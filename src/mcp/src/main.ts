#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { Client, PageIterator, PageCollection } from "@microsoft/microsoft-graph-client";
import fetch from 'isomorphic-fetch'; // Required polyfill for Graph client
import { logger } from "./logger.js";
import { AuthManager, AuthConfig, AuthMode } from "./auth.js";
import { LokkaClientId, LokkaDefaultTenantId, LokkaDefaultRedirectUri, getDefaultGraphApiVersion } from "./constants.js";
import { loadConfig, selectTenant, tenantConfigToAuthConfig, LokkaConfig } from "./config.js";

// Set up global fetch for the Microsoft Graph client
(global as any).fetch = fetch;

logger.info("Starting Lokka Multi-Microsoft API MCP Server (v0.3.0)");

// Initialize authentication and clients (module-level so tool closures can access them)
let authManager: AuthManager | null = null;
let graphClient: Client | null = null;
let lokkaAvailableTenants: Array<{ name: string }> = [];

// HTTP methods that require explicit user confirmation before execution
const METHODS_REQUIRING_CONFIRMATION = ["post", "put", "patch", "delete"] as const;

function requiresConfirmation(method: string): boolean {
  return (METHODS_REQUIRING_CONFIRMATION as readonly string[]).includes(method.toLowerCase());
}

// Start the server with stdio transport
async function main() {
  // Check USE_GRAPH_BETA environment variable
  const useGraphBeta = process.env.USE_GRAPH_BETA !== 'false';
  const defaultGraphApiVersion = getDefaultGraphApiVersion();

  logger.info(`Graph API default version: ${defaultGraphApiVersion} (USE_GRAPH_BETA=${process.env.USE_GRAPH_BETA || 'undefined'})`);

  // -------------------------------------------------------------------------
  // Step 1: Resolve authentication configuration
  // -------------------------------------------------------------------------
  let authConfig: AuthConfig;
  let lokkaConfig: LokkaConfig | null = null;

  const configPath = process.env.LOKKA_CONFIG;

  if (configPath) {
    // Load multi-tenant config file
    logger.info(`Loading multi-tenant config from: ${configPath}`);
    lokkaConfig = loadConfig(configPath);
    lokkaAvailableTenants = lokkaConfig.tenants.map((t) => ({ name: t.name }));

    const selectedTenantName = process.env.LOKKA_TENANT;
    const selectedTenant = selectTenant(lokkaConfig, selectedTenantName);
    logger.info(`Using tenant '${selectedTenant.name}' from config file`);
    authConfig = tenantConfigToAuthConfig(selectedTenant);
  } else {
    // Legacy single-tenant config via environment variables
    const useCertificate = process.env.USE_CERTIFICATE === 'true';
    const useInteractive = process.env.USE_INTERACTIVE === 'true';
    const useClientToken = process.env.USE_CLIENT_TOKEN === 'true';
    const initialAccessToken = process.env.ACCESS_TOKEN;

    let authMode: AuthMode;

    const enabledModes = [useClientToken, useInteractive, useCertificate].filter(Boolean);
    if (enabledModes.length > 1) {
      throw new Error(
        "Multiple authentication modes enabled. Please enable only one of USE_CLIENT_TOKEN, USE_INTERACTIVE, or USE_CERTIFICATE."
      );
    }

    if (useClientToken) {
      authMode = AuthMode.ClientProvidedToken;
      if (!initialAccessToken) {
        logger.info("Client token mode enabled but no initial token provided. Token must be set via set-access-token tool.");
      }
    } else if (useInteractive) {
      authMode = AuthMode.Interactive;
    } else if (useCertificate) {
      authMode = AuthMode.Certificate;
    } else {
      const hasClientCredentials = process.env.TENANT_ID && process.env.CLIENT_ID && process.env.CLIENT_SECRET;
      if (hasClientCredentials) {
        authMode = AuthMode.ClientCredentials;
      } else {
        authMode = AuthMode.Interactive;
        logger.info("No authentication mode specified and no client credentials found. Defaulting to interactive mode.");
      }
    }

    logger.info(`Starting with authentication mode: ${authMode}`);

    let tenantId: string | undefined;
    let clientId: string | undefined;

    if (authMode === AuthMode.Interactive) {
      tenantId = process.env.TENANT_ID || LokkaDefaultTenantId;
      clientId = process.env.CLIENT_ID || LokkaClientId;
      logger.info(`Interactive mode using tenant ID: ${tenantId}, client ID: ${clientId}`);
    } else {
      tenantId = process.env.TENANT_ID;
      clientId = process.env.CLIENT_ID;
    }

    const clientSecret = process.env.CLIENT_SECRET;
    const certificatePath = process.env.CERTIFICATE_PATH;
    const certificatePassword = process.env.CERTIFICATE_PASSWORD;

    if (authMode === AuthMode.ClientCredentials) {
      if (!tenantId || !clientId || !clientSecret) {
        throw new Error("Client credentials mode requires explicit TENANT_ID, CLIENT_ID, and CLIENT_SECRET environment variables");
      }
    } else if (authMode === AuthMode.Certificate) {
      if (!tenantId || !clientId || !certificatePath) {
        throw new Error("Certificate mode requires explicit TENANT_ID, CLIENT_ID, and CERTIFICATE_PATH environment variables");
      }
    }

    authConfig = {
      mode: authMode,
      tenantName: process.env.TENANT_NAME,
      tenantId,
      clientId,
      clientSecret,
      accessToken: initialAccessToken,
      redirectUri: process.env.REDIRECT_URI,
      certificatePath,
      certificatePassword,
    };
  }

  // -------------------------------------------------------------------------
  // Step 2: Initialise auth manager and Graph client
  // -------------------------------------------------------------------------
  authManager = new AuthManager(authConfig);
  await authManager.initialize();

  const authProvider = authManager.getGraphAuthProvider();
  graphClient = Client.initWithMiddleware({ authProvider });

  let tenantName = authManager.getTenantName();
  let tenantId = authConfig.tenantId;

  if (authConfig.mode === AuthMode.ClientProvidedToken && !authConfig.accessToken) {
    logger.info("Started in client token mode without initial token. Use set-access-token tool to provide authentication token.");
  } else {
    logger.info(`Authentication initialized successfully using ${authConfig.mode} mode`);
  }

  // -------------------------------------------------------------------------
  // Step 3: Log a prominent banner about which tenant is active
  // -------------------------------------------------------------------------
  let tenantDisplay = tenantName
    ? `${tenantName}${tenantId ? ` (${tenantId})` : ''}`
    : tenantId || "unknown";

  logger.info("=".repeat(60));
  logger.info(`ACTIVE TENANT: ${tenantDisplay}`);
  logger.info(`AUTH MODE    : ${authConfig.mode}`);
  if (lokkaAvailableTenants.length > 1) {
    logger.info(`ALL TENANTS  : ${lokkaAvailableTenants.map((t) => t.name).join(", ")}`);
  }
  logger.info("=".repeat(60));

  // -------------------------------------------------------------------------
  // Step 4: Create server with tenant name embedded in the server name
  // -------------------------------------------------------------------------
  const serverName = tenantName ? `Lokka-Microsoft [${tenantName}]` : "Lokka-Microsoft";
  const server = new McpServer({
    name: serverName,
    version: "0.3.1",
  });

  // Helper: build the tenant header line for tool responses
  function tenantHeader(): string {
    return `\u26a1 Tenant: ${tenantDisplay} | Auth: ${authConfig.mode}\n${"\u2500".repeat(60)}\n`;
  }

  // -------------------------------------------------------------------------
  // Tool: Lokka-Microsoft  (main API tool)
  // -------------------------------------------------------------------------
  server.tool(
    "Lokka-Microsoft",
    `A versatile tool to interact with Microsoft APIs including Microsoft Graph (Entra) and Azure Resource Management. ` +
    `Currently connected to tenant: ${tenantDisplay}. ` +
    `IMPORTANT: POST, PUT, PATCH and DELETE operations require explicit user confirmation (set confirm: true). ` +
    `For Graph API GET requests using advanced query parameters ($filter, $count, $search, $orderby), you are ADVISED to set 'consistencyLevel: "eventual"'.`,
    {
      apiType: z.enum(["graph", "azure"]).describe("Type of Microsoft API to query. Options: 'graph' for Microsoft Graph (Entra) or 'azure' for Azure Resource Management."),
      path: z.string().describe("The Azure or Graph API URL path to call (e.g. '/users', '/groups', '/subscriptions')"),
      method: z.enum(["get", "post", "put", "patch", "delete"]).describe("HTTP method to use. NOTE: POST, PUT, PATCH and DELETE require confirm: true. Only GET does not require confirmation."),
      apiVersion: z.string().optional().describe("Azure Resource Management API version (required for apiType Azure)"),
      subscriptionId: z.string().optional().describe("Azure Subscription ID (for Azure Resource Management)."),
      queryParams: z.record(z.string()).optional().describe("Query parameters for the request"),
      body: z.record(z.string(), z.any()).optional().describe("The request body (for POST, PUT, PATCH)"),
      graphApiVersion: z.enum(["v1.0", "beta"]).optional().default(defaultGraphApiVersion as "v1.0" | "beta").describe(`Microsoft Graph API version to use (default: ${defaultGraphApiVersion})`),
      fetchAll: z.boolean().optional().default(false).describe("Set to true to automatically fetch all pages for list results (e.g., users, groups). Default is false."),
      consistencyLevel: z.string().optional().describe("Graph API ConsistencyLevel header. ADVISED to be set to 'eventual' for Graph GET requests using advanced query parameters ($filter, $count, $search, $orderby)."),
      confirm: z.boolean().optional().describe("Required for POST, PUT, PATCH and DELETE operations (not needed for GET). Set to true only after the user has explicitly confirmed the operation."),
    },
    async ({
      apiType,
      path,
      method,
      apiVersion,
      subscriptionId,
      queryParams,
      body,
      graphApiVersion,
      fetchAll,
      consistencyLevel,
      confirm,
    }: {
      apiType: "graph" | "azure";
      path: string;
      method: "get" | "post" | "put" | "patch" | "delete";
      apiVersion?: string;
      subscriptionId?: string;
      queryParams?: Record<string, string>;
      body?: any;
      graphApiVersion: "v1.0" | "beta";
      fetchAll: boolean;
      consistencyLevel?: string;
      confirm?: boolean;
    }) => {
      // Override graphApiVersion if USE_GRAPH_BETA is explicitly set to false
      const effectiveGraphApiVersion = !useGraphBeta ? "v1.0" : graphApiVersion;

      // -----------------------------------------------------------------------
      // Confirmation gate for destructive / write operations
      // -----------------------------------------------------------------------
      if (requiresConfirmation(method) && confirm !== true) {
        const MAX_BODY_PREVIEW = 500;
        let bodyPreview = "";
        if (body) {
          const serialized = JSON.stringify(body, null, 2);
          const truncated = serialized.length > MAX_BODY_PREVIEW
            ? serialized.slice(0, MAX_BODY_PREVIEW) + "\n... (truncated)"
            : serialized;
          bodyPreview = `\nRequest body:\n${truncated}`;
        }
        const confirmationMessage =
          `\u26a0\ufe0f  CONFIRMATION REQUIRED\n` +
          `${"=".repeat(60)}\n` +
          `Tenant : ${tenantDisplay}\n` +
          `Method : ${method.toUpperCase()}\n` +
          `API    : ${apiType}\n` +
          `Path   : ${path}${bodyPreview}\n` +
          `${"=".repeat(60)}\n` +
          `This operation will MODIFY or DELETE data in the tenant above.\n` +
          `Please ask the user to explicitly confirm before proceeding.\n` +
          `Once confirmed, call this tool again with the same parameters and add: confirm: true`;
        return {
          content: [{ type: "text" as const, text: confirmationMessage }],
        };
      }

      logger.info(`Executing Lokka-Microsoft tool: apiType=${apiType}, path=${path}, method=${method}, graphApiVersion=${effectiveGraphApiVersion}, fetchAll=${fetchAll}, confirm=${confirm}`);
      let determinedUrl: string | undefined;

      try {
        let responseData: any;

        // --- Microsoft Graph Logic ---
        if (apiType === 'graph') {
          if (!graphClient) {
            throw new Error("Graph client not initialized");
          }
          determinedUrl = `https://graph.microsoft.com/${effectiveGraphApiVersion}`;

          let request = graphClient.api(path).version(effectiveGraphApiVersion);

          if (queryParams && Object.keys(queryParams).length > 0) {
            request = request.query(queryParams);
          }

          if (consistencyLevel) {
            request = request.header('ConsistencyLevel', consistencyLevel);
            logger.info(`Added ConsistencyLevel header: ${consistencyLevel}`);
          }

          switch (method.toLowerCase()) {
            case 'get':
              if (fetchAll) {
                logger.info(`Fetching all pages for Graph path: ${path}`);
                const firstPageResponse: PageCollection = await request.get();
                const odataContext = firstPageResponse['@odata.context'];

                if (!Array.isArray(firstPageResponse.value)) {
                  logger.info(`Response for ${path} is not a collection (no value array). Returning single response.`);
                  responseData = firstPageResponse;
                } else {
                  let allItems: any[] = [];
                  const callback = (item: any) => {
                    allItems.push(item);
                    return true;
                  };
                  const pageIterator = new PageIterator(graphClient!, firstPageResponse, callback);
                  await pageIterator.iterate();
                  responseData = {
                    '@odata.context': odataContext,
                    value: allItems
                  };
                  logger.info(`Finished fetching all Graph pages. Total items: ${allItems.length}`);
                }
              } else {
                logger.info(`Fetching single page for Graph path: ${path}`);
                responseData = await request.get();
              }
              break;
            case 'post':
              responseData = await request.post(body ?? {});
              break;
            case 'put':
              responseData = await request.put(body ?? {});
              break;
            case 'patch':
              responseData = await request.patch(body ?? {});
              break;
            case 'delete':
              responseData = await request.delete();
              if (responseData === undefined || responseData === null) {
                responseData = { status: "Success (No Content)" };
              }
              break;
            default:
              throw new Error(`Unsupported method: ${method}`);
          }
        }
        // --- Azure Resource Management Logic (using direct fetch) ---
        else {
          if (!authManager) {
            throw new Error("Auth manager not initialized");
          }
          determinedUrl = "https://management.azure.com";

          const azureCredential = authManager.getAzureCredential();
          const tokenResponse = await azureCredential.getToken("https://management.azure.com/.default");
          if (!tokenResponse || !tokenResponse.token) {
            throw new Error("Failed to acquire Azure access token");
          }

          let url = determinedUrl;
          if (subscriptionId) {
            url += `/subscriptions/${subscriptionId}`;
          }
          url += path;

          if (!apiVersion) {
            throw new Error("API version is required for Azure Resource Management queries");
          }
          const urlParams = new URLSearchParams({ 'api-version': apiVersion });
          if (queryParams) {
            for (const [key, value] of Object.entries(queryParams)) {
              urlParams.append(String(key), String(value));
            }
          }
          url += `?${urlParams.toString()}`;

          const headers: Record<string, string> = {
            'Authorization': `Bearer ${tokenResponse.token}`,
            'Content-Type': 'application/json'
          };
          const requestOptions: RequestInit = {
            method: method.toUpperCase(),
            headers: headers
          };
          if (["POST", "PUT", "PATCH"].includes(method.toUpperCase())) {
            requestOptions.body = body ? JSON.stringify(body) : JSON.stringify({});
          }

          if (fetchAll && method === 'get') {
            logger.info(`Fetching all pages for Azure RM starting from: ${url}`);
            let allValues: any[] = [];
            let currentUrl: string | null = url;

            while (currentUrl) {
              logger.info(`Fetching Azure RM page: ${currentUrl}`);
              const azureCredential = authManager.getAzureCredential();
              const currentPageTokenResponse = await azureCredential.getToken("https://management.azure.com/.default");
              if (!currentPageTokenResponse || !currentPageTokenResponse.token) {
                throw new Error("Failed to acquire Azure access token during pagination");
              }
              const currentPageHeaders = { ...headers, 'Authorization': `Bearer ${currentPageTokenResponse.token}` };
              const currentPageRequestOptions: RequestInit = { method: 'GET', headers: currentPageHeaders };

              const pageResponse = await fetch(currentUrl, currentPageRequestOptions);
              const pageText = await pageResponse.text();
              let pageData: any;
              try {
                pageData = pageText ? JSON.parse(pageText) : {};
              } catch (e) {
                logger.error(`Failed to parse JSON from Azure RM page: ${currentUrl}`, pageText);
                pageData = { rawResponse: pageText };
              }

              if (!pageResponse.ok) {
                logger.error(`API error on Azure RM page ${currentUrl}:`, pageData);
                throw new Error(`API error (${pageResponse.status}) during Azure RM pagination on ${currentUrl}: ${JSON.stringify(pageData)}`);
              }

              if (pageData.value && Array.isArray(pageData.value)) {
                allValues = allValues.concat(pageData.value);
              } else if (currentUrl === url && !pageData.nextLink) {
                allValues.push(pageData);
              } else if (currentUrl !== url) {
                logger.info(`[Warning] Azure RM response from ${currentUrl} did not contain a 'value' array.`);
              }
              currentUrl = pageData.nextLink || null;
            }
            responseData = { allValues: allValues };
            logger.info(`Finished fetching all Azure RM pages. Total items: ${allValues.length}`);
          } else {
            logger.info(`Fetching single page for Azure RM: ${url}`);
            const apiResponse = await fetch(url, requestOptions);
            const responseText = await apiResponse.text();
            try {
              responseData = responseText ? JSON.parse(responseText) : {};
            } catch (e) {
              logger.error(`Failed to parse JSON from single Azure RM page: ${url}`, responseText);
              responseData = { rawResponse: responseText };
            }
            if (!apiResponse.ok) {
              logger.error(`API error for Azure RM ${method} ${path}:`, responseData);
              throw new Error(`API error (${apiResponse.status}) for Azure RM: ${JSON.stringify(responseData)}`);
            }
          }
        }

        // --- Format and Return Result ---
        let resultText = tenantHeader();
        resultText += `Result for ${apiType} API (${apiType === 'graph' ? effectiveGraphApiVersion : apiVersion}) - ${method.toUpperCase()} ${path}:\n\n`;
        resultText += JSON.stringify(responseData, null, 2);

        if (!fetchAll && method === 'get') {
          const nextLinkKey = apiType === 'graph' ? '@odata.nextLink' : 'nextLink';
          if (responseData && responseData[nextLinkKey]) {
            resultText += `\n\nNote: More results are available. To retrieve all pages, add the parameter 'fetchAll: true' to your request.`;
          }
        }

        return {
          content: [{ type: "text" as const, text: resultText }],
        };

      } catch (error: any) {
        logger.error(`Error in Lokka-Microsoft tool (apiType: ${apiType}, path: ${path}, method: ${method}):`, error);
        if (!determinedUrl) {
          determinedUrl = apiType === 'graph'
            ? `https://graph.microsoft.com/${effectiveGraphApiVersion}`
            : "https://management.azure.com";
        }
        const errorBody = error.body ? (typeof error.body === 'string' ? error.body : JSON.stringify(error.body)) : 'N/A';
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              tenant: tenantDisplay,
              error: error instanceof Error ? error.message : String(error),
              statusCode: error.statusCode || 'N/A',
              errorBody: errorBody,
              attemptedBaseUrl: determinedUrl
            }),
          }],
          isError: true
        };
      }
    },
  );

  // -------------------------------------------------------------------------
  // Tool: set-access-token
  // -------------------------------------------------------------------------
  server.tool(
    "set-access-token",
    "Set or update the access token for Microsoft Graph authentication. Use this when the MCP Client has obtained a fresh token through interactive authentication.",
    {
      accessToken: z.string().describe("The access token obtained from Microsoft Graph authentication"),
      expiresOn: z.string().optional().describe("Token expiration time in ISO format (optional, defaults to 1 hour from now)")
    },
    async ({ accessToken, expiresOn }) => {
      try {
        const expirationDate = expiresOn ? new Date(expiresOn) : undefined;

        if (authManager?.getAuthMode() === AuthMode.ClientProvidedToken) {
          authManager.updateAccessToken(accessToken, expirationDate);

          const authProvider = authManager.getGraphAuthProvider();
          graphClient = Client.initWithMiddleware({ authProvider });

          return {
            content: [{
              type: "text" as const,
              text: "Access token updated successfully. You can now make Microsoft Graph requests on behalf of the authenticated user."
            }],
          };
        } else {
          return {
            content: [{
              type: "text" as const,
              text: "Error: MCP Server is not configured for client-provided token authentication. Set USE_CLIENT_TOKEN=true in environment variables."
            }],
            isError: true
          };
        }
      } catch (error: any) {
        logger.error("Error setting access token:", error);
        return {
          content: [{
            type: "text" as const,
            text: `Error setting access token: ${error.message}`
          }],
          isError: true
        };
      }
    }
  );

  // -------------------------------------------------------------------------
  // Tool: get-auth-status
  // -------------------------------------------------------------------------
  server.tool(
    "get-auth-status",
    "Check the current authentication status, active tenant information, and permission scopes of the MCP Server session.",
    {},
    async () => {
      try {
        const authMode = authManager?.getAuthMode() || "Not initialized";
        const isReady = authManager !== null;
        const tokenStatus = authManager ? await authManager.getTokenStatus() : { isExpired: false };

        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify({
              tenant: {
                name: tenantName || null,
                tenantId: authConfig.tenantId || null,
                display: tenantDisplay,
              },
              authMode,
              isReady,
              supportsTokenUpdates: authMode === AuthMode.ClientProvidedToken,
              tokenStatus,
              timestamp: new Date().toISOString()
            }, null, 2)
          }],
        };
      } catch (error: any) {
        return {
          content: [{
            type: "text" as const,
            text: `Error checking auth status: ${error.message}`
          }],
          isError: true
        };
      }
    }
  );

  // -------------------------------------------------------------------------
  // Tool: add-graph-permission
  // -------------------------------------------------------------------------
  server.tool(
    "add-graph-permission",
    "Request additional Microsoft Graph permission scopes by performing a fresh interactive sign-in. This tool only works in interactive authentication mode and should be used if any Graph API call returns permissions related errors.",
    {
      scopes: z.array(z.string()).describe("Array of Microsoft Graph permission scopes to request (e.g., ['User.Read', 'Mail.ReadWrite', 'Directory.Read.All'])")
    },
    async ({ scopes }) => {
      try {
        if (!authManager || authManager.getAuthMode() !== AuthMode.Interactive) {
          const currentMode = authManager?.getAuthMode() || "Not initialized";
          const clientId = authConfig.clientId || process.env.CLIENT_ID;

          let errorMessage = `Error: add-graph-permission tool is only available in interactive authentication mode. Current mode: ${currentMode}.\n\n`;

          if (currentMode === AuthMode.ClientCredentials) {
            errorMessage += `To add permissions in Client Credentials mode:\n`;
            errorMessage += `1. Open the Microsoft Entra admin center (https://entra.microsoft.com)\n`;
            errorMessage += `2. Navigate to Applications > App registrations\n`;
            errorMessage += `3. Find your application${clientId ? ` (Client ID: ${clientId})` : ''}\n`;
            errorMessage += `4. Go to API permissions\n`;
            errorMessage += `5. Click "Add a permission" and select Microsoft Graph\n`;
            errorMessage += `6. Choose "Application permissions" and add the required scopes:\n`;
            errorMessage += `   ${scopes.map(scope => `\u2022 ${scope}`).join('\n   ')}\n`;
            errorMessage += `7. Click "Grant admin consent" to approve the permissions\n`;
            errorMessage += `8. Restart the MCP server to use the new permissions`;
          } else if (currentMode === AuthMode.ClientProvidedToken) {
            errorMessage += `To add permissions in Client Provided Token mode:\n`;
            errorMessage += `1. Obtain a new access token that includes the required scopes:\n`;
            errorMessage += `   ${scopes.map(scope => `\u2022 ${scope}`).join('\n   ')}\n`;
            errorMessage += `2. When obtaining the token, ensure these scopes are included in the consent prompt\n`;
            errorMessage += `3. Use the set-access-token tool to update the server with the new token\n`;
            errorMessage += `4. The new token will include the additional permissions`;
          } else {
            errorMessage += `To use interactive permission requests, set USE_INTERACTIVE=true in environment variables and restart the server.`;
          }

          return {
            content: [{ type: "text" as const, text: errorMessage }],
            isError: true
          };
        }

        if (!scopes || scopes.length === 0) {
          return {
            content: [{ type: "text" as const, text: "Error: At least one permission scope must be specified." }],
            isError: true
          };
        }

        const invalidScopes = scopes.filter(scope => !scope.includes('.') || scope.trim() !== scope);
        if (invalidScopes.length > 0) {
          return {
            content: [{
              type: "text" as const,
              text: `Error: Invalid scope format detected: ${invalidScopes.join(', ')}. Scopes should be in format like 'User.Read' or 'Mail.ReadWrite'.`
            }],
            isError: true
          };
        }

        logger.info(`Requesting additional Graph permissions: ${scopes.join(', ')}`);

        const currentTenantId = authConfig.tenantId || LokkaDefaultTenantId;
        const currentClientId = authConfig.clientId || LokkaClientId;
        const redirectUri = authConfig.redirectUri || LokkaDefaultRedirectUri;

        logger.info(`Using tenant ID: ${currentTenantId}, client ID: ${currentClientId} for interactive authentication`);

        const { InteractiveBrowserCredential, DeviceCodeCredential } = await import("@azure/identity");

        authManager = null;
        graphClient = null;

        const scopeString = scopes.map(scope => `https://graph.microsoft.com/${scope}`).join(' ');
        logger.info(`Requesting fresh token with scopes: ${scopeString}`);

        console.log(`\n\ud83d\udd10 Requesting Additional Graph Permissions:`);
        console.log(`Scopes: ${scopes.join(', ')}`);
        console.log(`You will be prompted to sign in to grant these permissions.\n`);

        let newCredential;
        let tokenResponse;

        try {
          newCredential = new InteractiveBrowserCredential({
            tenantId: currentTenantId,
            clientId: currentClientId,
            redirectUri: redirectUri,
          });
          tokenResponse = await newCredential.getToken(scopeString);
        } catch (error) {
          logger.info("Interactive browser failed, falling back to device code flow");
          newCredential = new DeviceCodeCredential({
            tenantId: currentTenantId,
            clientId: currentClientId,
            userPromptCallback: (info) => {
              console.log(`\n\ud83d\udd10 Additional Permissions Required:`);
              console.log(`Please visit: ${info.verificationUri}`);
              console.log(`And enter code: ${info.userCode}`);
              console.log(`Requested scopes: ${scopes.join(', ')}\n`);
              return Promise.resolve();
            },
          });
          tokenResponse = await newCredential.getToken(scopeString);
        }

        if (!tokenResponse) {
          return {
            content: [{
              type: "text" as const,
              text: "Error: Failed to acquire access token with the requested scopes. Please check your permissions and try again."
            }],
            isError: true
          };
        }

        const newAuthConfig: AuthConfig = {
          mode: AuthMode.Interactive,
          tenantName: authConfig.tenantName,
          tenantId: currentTenantId,
          clientId: currentClientId,
          redirectUri
        };

        authManager = new AuthManager(newAuthConfig);
        (authManager as any).credential = newCredential;

        const newAuthProvider = authManager.getGraphAuthProvider();
        graphClient = Client.initWithMiddleware({ authProvider: newAuthProvider });

        const tokenStatus = await authManager.getTokenStatus();

        logger.info(`Successfully acquired fresh token with additional scopes: ${scopes.join(', ')}`);

        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify({
              message: "Successfully acquired additional Microsoft Graph permissions with fresh authentication",
              requestedScopes: scopes,
              tokenStatus,
              note: "A fresh sign-in was performed to ensure the new permissions are properly granted",
              timestamp: new Date().toISOString()
            }, null, 2)
          }],
        };

      } catch (error: any) {
        logger.error("Error requesting additional Graph permissions:", error);
        return {
          content: [{
            type: "text" as const,
            text: `Error requesting additional permissions: ${error.message}`
          }],
          isError: true
        };
      }
    }
  );

  // -------------------------------------------------------------------------
  // Tool: switch-tenant
  // -------------------------------------------------------------------------
  server.tool(
    "switch-tenant",
    "Switch the active Microsoft tenant at runtime. Only available when a multi-tenant config file is loaded (LOKKA_CONFIG). Use list-tenants to see available tenants.",
    {
      tenantName: z.string().describe("Name of the tenant to switch to (must match a tenant name in the config file)"),
    },
    async ({ tenantName: requestedTenant }) => {
      try {
        if (!lokkaConfig) {
          return {
            content: [{
              type: "text" as const,
              text: "Error: switch-tenant is only available when a multi-tenant config file is loaded. Set LOKKA_CONFIG to a JSON config file path and restart the server.",
            }],
            isError: true,
          };
        }

        const selectedTenant = selectTenant(lokkaConfig, requestedTenant);

        if (selectedTenant.name === tenantName) {
          return {
            content: [{
              type: "text" as const,
              text: `Already connected to tenant '${tenantName}'. No switch needed.`,
            }],
          };
        }

        logger.info(`Switching tenant from '${tenantName}' to '${selectedTenant.name}'`);

        const newAuthConfig = tenantConfigToAuthConfig(selectedTenant);
        const newAuthManager = new AuthManager(newAuthConfig);
        await newAuthManager.initialize();

        const newAuthProvider = newAuthManager.getGraphAuthProvider();
        const newGraphClient = Client.initWithMiddleware({ authProvider: newAuthProvider });

        // Update module-level and closure-captured variables
        authManager = newAuthManager;
        graphClient = newGraphClient;
        authConfig = newAuthConfig;
        tenantName = newAuthManager.getTenantName();
        tenantId = newAuthConfig.tenantId;
        tenantDisplay = tenantName
          ? `${tenantName}${tenantId ? ` (${tenantId})` : ''}`
          : tenantId || "unknown";

        logger.info(`Successfully switched to tenant: ${tenantDisplay}`);

        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify({
              message: `Successfully switched to tenant '${selectedTenant.name}'`,
              activeTenant: {
                name: tenantName,
                tenantId: tenantId || null,
                display: tenantDisplay,
                authMode: newAuthConfig.mode,
              },
              timestamp: new Date().toISOString(),
            }, null, 2),
          }],
        };
      } catch (error: any) {
        logger.error("Error switching tenant:", error);
        return {
          content: [{
            type: "text" as const,
            text: `Error switching tenant: ${error.message}`,
          }],
          isError: true,
        };
      }
    }
  );

  // -------------------------------------------------------------------------
  // Tool: list-tenants
  // -------------------------------------------------------------------------
  server.tool(
    "list-tenants",
    "List all configured tenants and show which one is currently active. To switch tenants, set the LOKKA_TENANT environment variable to a tenant name and restart the server.",
    {},
    async () => {
      const activeName = tenantName || tenantId || "unknown";
      const tenantList = lokkaAvailableTenants.length > 0
        ? lokkaAvailableTenants
        : [{ name: activeName }];

      const lines = tenantList.map((t) => {
        const isActive = t.name === activeName || (!tenantName && !tenantId);
        return `${isActive ? "\u25b6 " : "  "}${t.name}${isActive ? "  \u2190 ACTIVE" : ""}`;
      });

      return {
        content: [{
          type: "text" as const,
          text:
            `Configured tenants:\n\n${lines.join("\n")}\n\n` +
            `Active tenant: ${tenantDisplay}\n\n` +
            `To switch tenant, set LOKKA_TENANT=<name> and restart the MCP server.\n` +
            (configPath ? `Config file: ${configPath}` : "Using single-tenant environment variable configuration.\nTo use multiple tenants, create a JSON config file and set LOKKA_CONFIG=<path>."),
        }],
      };
    }
  );

  // -------------------------------------------------------------------------
  // Tool: restart
  // -------------------------------------------------------------------------
  server.tool(
    "restart",
    "Restart the Lokka MCP server process. The MCP client will automatically relaunch the server with the same configuration. Use this after changing environment variables or configuration files.",
    {},
    async () => {
      logger.info("Restart requested via restart tool. Exiting process for client-initiated relaunch.");
      // Send the response before exiting so the client receives the message
      setImmediate(() => process.exit(0));
      return {
        content: [{
          type: "text" as const,
          text: "Lokka MCP server is restarting. The MCP client will relaunch the server automatically.",
        }],
      };
    }
  );

  // -------------------------------------------------------------------------
  // Step 5: Connect transport
  // -------------------------------------------------------------------------
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  logger.error("Fatal error in main()", error);
  process.exit(1);
});
