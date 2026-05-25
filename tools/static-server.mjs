import { createReadStream, existsSync, statSync } from "node:fs";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, "..");

const args = process.argv.slice(2);
const rootArg = args.find(arg => !arg.startsWith("--")) || ".";
const portArg = args.find(arg => arg.startsWith("--port="));
const port = Number(portArg ? portArg.split("=")[1] : process.env.PORT || 4173);
const rootDir = path.resolve(projectRoot, rootArg);

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".webp": "image/webp"
};

function send(res, status, body, contentType = "text/plain; charset=utf-8") {
  res.writeHead(status, {
    "Content-Type": contentType,
    "Cache-Control": "no-store"
  });
  res.end(body);
}

function resolveRequestPath(requestUrl) {
  const url = new URL(requestUrl, `http://localhost:${port}`);
  const decodedPath = decodeURIComponent(url.pathname);
  const normalizedPath = path.normalize(decodedPath).replace(/^(\.\.[/\\])+/, "");
  let filePath = path.join(rootDir, normalizedPath);

  const relative = path.relative(rootDir, filePath);
  if (relative.startsWith("..") || path.isAbsolute(relative)) return null;

  if (existsSync(filePath) && statSync(filePath).isDirectory()) {
    filePath = path.join(filePath, "index.html");
  }

  return filePath;
}

if (!Number.isFinite(port) || port <= 0) {
  console.error("[serve] Porta invalida.");
  process.exit(1);
}

if (!existsSync(rootDir) || !statSync(rootDir).isDirectory()) {
  console.error(`[serve] Diretorio nao encontrado: ${path.relative(projectRoot, rootDir) || rootDir}`);
  process.exit(1);
}

const server = http.createServer((req, res) => {
  const filePath = resolveRequestPath(req.url || "/");

  if (!filePath) {
    send(res, 403, "Acesso bloqueado.");
    return;
  }

  if (!existsSync(filePath) || !statSync(filePath).isFile()) {
    send(res, 404, "Arquivo nao encontrado.");
    return;
  }

  const ext = path.extname(filePath).toLowerCase();
  res.writeHead(200, {
    "Content-Type": mimeTypes[ext] || "application/octet-stream",
    "Cache-Control": "no-store"
  });
  createReadStream(filePath).pipe(res);
});

server.listen(port, () => {
  const servedRoot = path.relative(projectRoot, rootDir) || ".";
  console.log(`[serve] Servindo ${servedRoot} em http://localhost:${port}/`);
});
