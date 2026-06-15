import { existsSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

const __dirname = dirname(fileURLToPath(import.meta.url));

/**
 * Reads `.envrc` in `baseDir`. Supports `KEY="value"` and `export KEY="value"`.
 * Does not replace keys already present in `process.env`.
 */
function loadEnvrc(baseDir) {
  const envrcPath = join(baseDir, ".envrc");
  if (!existsSync(envrcPath)) return;
  const text = readFileSync(envrcPath, "utf8");
  for (const rawLine of text.split("\n")) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const normalized = line.replace(/^export\s+/i, "");
    const eq = normalized.indexOf("=");
    if (eq <= 0) continue;
    const key = normalized.slice(0, eq).trim();
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) continue;
    let value = normalized.slice(eq + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    if (process.env[key] === undefined) {
      process.env[key] = value;
    }
  }
}

loadEnvrc(__dirname);
loadEnvrc(join(__dirname, "mynordic-oauth-test"));

const b2cAuthority =
  process.env.B2C_AUTHORITY ?? process.env.AZURE_AD_CIAM_URL ?? "";
const b2cClientId =
  process.env.B2C_CLIENT_ID ?? process.env.AZURE_NRF_CLOUD_CLIENT_ID ?? "";

export default defineConfig({
  define: {
    B2C_AUTHORITY: JSON.stringify(b2cAuthority),
    B2C_CLIENT_ID: JSON.stringify(b2cClientId),
  },
  server: {
    port: 8080,
  },
});
