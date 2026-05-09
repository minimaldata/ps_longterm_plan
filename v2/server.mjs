import { createReadStream, existsSync, statSync } from "node:fs";
import { createServer } from "node:http";
import { extname, join, normalize, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const PORT = Number(process.env.PORT || 8020);
const HOST = process.env.HOST || "127.0.0.1";
const ROOT = fileURLToPath(new URL(".", import.meta.url));
const REPO_ROOT = resolve(ROOT, "..");

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".svg": "image/svg+xml",
};

function resolveRequestPath(url) {
  const requestPath = decodeURIComponent(new URL(url, `http://${HOST}:${PORT}`).pathname);
  if (requestPath === "/napkinChart.js") return join(REPO_ROOT, "napkinChart.js");
  const normalizedPath = normalize(requestPath).replace(/^(\.\.[/\\])+/, "");
  const filePath = resolve(join(ROOT, normalizedPath === "/" ? "index.html" : normalizedPath));
  if (!filePath.startsWith(ROOT)) return null;
  if (existsSync(filePath) && statSync(filePath).isDirectory()) {
    return join(filePath, "index.html");
  }
  return filePath;
}

const server = createServer((request, response) => {
  const filePath = resolveRequestPath(request.url);
  if (!filePath || !existsSync(filePath)) {
    response.writeHead(404, { "content-type": "text/plain; charset=utf-8" });
    response.end("Not found");
    return;
  }

  const headers = {
    "cache-control": "no-store",
    "content-type": mimeTypes[extname(filePath)] || "application/octet-stream",
  };

  response.writeHead(200, headers);
  createReadStream(filePath).pipe(response);
});

server.listen(PORT, HOST, () => {
  console.log(`V2 server running at http://${HOST}:${PORT}/`);
});
