import { readFileSync, existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

const __dirname = dirname(fileURLToPath(import.meta.url));

/**
 * Reads `.envrc` next to this config. Supports `KEY="value"` and `export KEY="value"`.
 * Does not replace keys already present in `process.env`.
 */
function loadEnvrc(baseDir = __dirname) {
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

loadEnvrc();

/** Serve /mynordic/from-portal (no trailing slash) — matches MSAL redirect URI exactly. */
function mynordicRouteMiddleware(req, _res, next) {
  const [pathname, search = ""] = (req.url ?? "").split("?");
  if (pathname === "/mynordic/from-portal") {
    req.url = `/mynordic/from-portal/index.html${search ? `?${search}` : ""}`;
  }
  next();
}

function mynordicRoutePlugin() {
  return {
    name: "mynordic-routes",
    configureServer(server) {
      server.middlewares.use(mynordicRouteMiddleware);
    },
    configurePreviewServer(server) {
      server.middlewares.use(mynordicRouteMiddleware);
    },
  };
}

export default defineConfig({
  plugins: [mynordicRoutePlugin()],
  define: {
    B2C_AUTHORITY: JSON.stringify(process.env.B2C_AUTHORITY ?? ""),
    B2C_CLIENT_ID: JSON.stringify(process.env.B2C_CLIENT_ID ?? ""),
  },
  build: {
    rollupOptions: {
      input: {
        main: join(__dirname, "index.html"),
        fromPortal: join(__dirname, "mynordic/from-portal/index.html"),
        sso: join(__dirname, "mynordic/sso.html"),
      },
    },
  },
  server: {
    port: 8080,
  },
});
