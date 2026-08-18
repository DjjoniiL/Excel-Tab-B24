"use strict";

const fs = require("node:fs");
const path = require("node:path");

const runtimeFiles = ["install.html", "install.js", "install.css", "index.html", "app.js", "style.css"];
const root = path.resolve(__dirname, "..");
const missing = runtimeFiles.filter((file) => !fs.existsSync(path.join(root, file)));

if (missing.length) {
  console.error(`Missing Marketplace runtime files: ${missing.join(", ")}`);
  process.exit(1);
}

console.log(`Marketplace runtime files present: ${runtimeFiles.join(", ")}`);
