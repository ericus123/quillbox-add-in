const fs = require("fs");
const path = require("path");
require("dotenv").config();

const BASE_URL = process.env.BASE_URL_DEV;

if (!BASE_URL) {
  console.error("Error: BASE_URL environment variable is required!");
  process.exit(1);
}

const input = path.resolve(__dirname, "manifest.template.xml");
const output = path.resolve(__dirname, "manifest.xml");

// Read template
let xml = fs.readFileSync(input, "utf8");

// Replace all instances of {{BASE_URL}}
xml = xml.replace(/{{BASE_URL}}/g, BASE_URL);

// Write to manifest.xml
fs.writeFileSync(output, xml);

console.log(`✅ Built manifest.xml with BASE_URL=${BASE_URL}`);
