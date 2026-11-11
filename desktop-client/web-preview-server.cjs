const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");

const PORT = 3000;
const RENDERER_DIR = path.join(__dirname, "renderer");

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".png": "image/png",
  ".ico": "image/x-icon",
  ".json": "application/json",
};

const server = http.createServer((req, res) => {
  let reqUrl = req.url.split("?")[0];
  let relativePath = reqUrl === "/" ? "index.html" : reqUrl.replace(/^\//, "");
  let filePath = path.join(RENDERER_DIR, relativePath);

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain" });
      res.end("404 Not Found");
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, { "Content-Type": MIME_TYPES[ext] || "text/plain" });
    res.end(data);
  });
});

server.listen(PORT, "0.0.0.0", () => {
  console.log(`ZBOM Web Preview Server running at http://localhost:${PORT}`);
});
