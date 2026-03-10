// Multi-tenant configuration support for Lokka MCP Server
import { readFileSync } from "fs";
import { AuthMode } from "./auth.js";
export function loadConfig(configPath) {
    const content = readFileSync(configPath, "utf-8");
    const config = JSON.parse(content);
    if (!config.tenants || !Array.isArray(config.tenants)) {
        throw new Error("Invalid config: 'tenants' array is required");
    }
    if (config.tenants.length === 0) {
        throw new Error("Invalid config: 'tenants' array must not be empty");
    }
    for (const tenant of config.tenants) {
        if (!tenant.name || typeof tenant.name !== "string") {
            throw new Error("Invalid config: each tenant must have a 'name' string field");
        }
    }
    return config;
}
export function selectTenant(config, tenantName) {
    if (tenantName) {
        const found = config.tenants.find((t) => t.name.toLowerCase() === tenantName.toLowerCase());
        if (!found) {
            const available = config.tenants.map((t) => t.name).join(", ");
            throw new Error(`Tenant '${tenantName}' not found in config. Available tenants: ${available}`);
        }
        return found;
    }
    return config.tenants[0];
}
export function tenantConfigToAuthConfig(tenant) {
    let mode;
    switch ((tenant.authMode ?? "").toLowerCase()) {
        case "client_credentials":
            mode = AuthMode.ClientCredentials;
            break;
        case "certificate":
            mode = AuthMode.Certificate;
            break;
        case "client_provided_token":
            mode = AuthMode.ClientProvidedToken;
            break;
        case "interactive":
            mode = AuthMode.Interactive;
            break;
        default:
            // Auto-detect based on available fields
            if (tenant.tenantId && tenant.clientId && tenant.clientSecret) {
                mode = AuthMode.ClientCredentials;
            }
            else if (tenant.tenantId && tenant.clientId && tenant.certificatePath) {
                mode = AuthMode.Certificate;
            }
            else {
                mode = AuthMode.Interactive;
            }
    }
    return {
        mode,
        tenantName: tenant.name,
        tenantId: tenant.tenantId,
        clientId: tenant.clientId,
        clientSecret: tenant.clientSecret,
        certificatePath: tenant.certificatePath,
        certificatePassword: tenant.certificatePassword,
        redirectUri: tenant.redirectUri,
    };
}
