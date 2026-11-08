// Usage: node set-url.js https://YOUR-NAME.github.io/study-tracker
// Writes manifest.xml pointing Word at the place where the docs/ folder is hosted.
const fs = require("fs");
const path = require("path");
const base = (process.argv[2] || "https://localhost:3000").replace(/\/+$/, "");
if (!/^https:\/\//.test(base)) { console.error("The address must start with https://"); process.exit(1); }
const tpl = fs.readFileSync(path.join(__dirname, "manifest.template.xml"), "utf8");
fs.writeFileSync(path.join(__dirname, "manifest.xml"), tpl.split("{{BASE}}").join(base));
console.log("manifest.xml now points to " + base);
