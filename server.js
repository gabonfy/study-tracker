// Serves docs/ at https://localhost:3000 with a trusted local certificate (installed on first run).
const https = require("https");
const fs = require("fs");
const path = require("path");
const devCerts = require("office-addin-dev-certs");
const ROOT = path.join(__dirname, "docs");
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png", ".xml": "text/xml" };

devCerts.getHttpsServerOptions().then((opts) => {
  https.createServer(opts, (req, res) => {
    const urlPath = decodeURIComponent(req.url.split("?")[0]);
    const file = path.normalize(path.join(ROOT, urlPath === "/" ? "taskpane.html" : urlPath));
    if (!file.startsWith(ROOT)) { res.writeHead(403); return res.end(); }
    fs.readFile(file, (err, data) => {
      if (err) { res.writeHead(404); return res.end("Not found"); }
      res.writeHead(200, { "Content-Type": TYPES[path.extname(file)] || "application/octet-stream", "Cache-Control": "no-store" });
      res.end(data);
    });
  }).listen(3000, () => console.log("Study Tracker is being served at https://localhost:3000 – keep this window open while you use it."));
});
