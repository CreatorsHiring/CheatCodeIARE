const PizZip = require("pizzip");
const fs = require("fs");
const path = require("path");

const filePath = path.join(__dirname, "assets", "ReportTemplate.docx");
const content = fs.readFileSync(filePath, "binary");

const zip = new PizZip(content);
let docXml = zip.file("word/document.xml").asText();

// Replace {answer1} -> {@answer1} to enable raw XML injection
// We use a regex to be safe about potential xml tags inside the braces if any (though unlikely for simple text)
// But often Word splits text like {<w:t>answer</w:t>1}.
// This simple replace might fail if the placeholder is split.
// For now, we assume a clean template or we use a regex that handles split tags?
// Complex regex for split tags is hard.
// Let's try a simple replace first. If the user provided a clean template it might work.
// Actually, it's safer to just enable raw XML for ALL fields in the prompt logic, 
// BUT docxtemplater requires the @ prefix in the template itself.

// STRATEGY: 
// 1. Try simple replace.
// 2. If that fails, we might need to ask user to fix template.

// We will look for `{answer` and replace with `{@answer`.
const fixedXml = docXml.replace(/\{answer/g, "{@answer");

zip.file("word/document.xml", fixedXml);

const buffer = zip.generate({ type: "nodebuffer" });
fs.writeFileSync(filePath, buffer);

console.log("Template patched to use {@answer} for raw XML.");
