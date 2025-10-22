// scripts/docx-utils.js
const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

/**
 * Basic checks to ensure file exists and is a zip (.docx)
 */
function validateDocxFile(filePath) {
  if (!fs.existsSync(filePath)) throw new Error(`Template not found: ${filePath}`);
  const stats = fs.statSync(filePath);
  if (!stats.isFile() || stats.size < 100) {
    throw new Error(`Template appears invalid or too small (size=${stats.size}). Re-save & re-upload the .docx.`);
  }

  // read first 4 bytes to check PK signature
  const fd = fs.openSync(filePath, "r");
  const header = Buffer.alloc(4);
  fs.readSync(fd, header, 0, 4, 0);
  fs.closeSync(fd);
  if (header.readUInt32LE(0) !== 0x04034b50) { // little-endian 'PK\x03\x04'
    throw new Error("File does not appear to be a ZIP archive (missing PK header). Ensure this is a valid .docx file.");
  }
}

/**
 * Generate docx by replacing placeholders using docxtemplater.
 * Returns the output path on success or throws an informative error.
 */
function generateDocx(templatePath, outputDocxPath, data) {
  validateDocxFile(templatePath);

  const contentBuffer = fs.readFileSync(templatePath);
  let zip;
  try {
    zip = new PizZip(contentBuffer);
  } catch (err) {
    throw new Error(
      `Error while reading template as ZIP. The template may be corrupted or not a valid .docx.\n` +
      `Original error: ${err.message}\n` +
      `Suggestion: open template in Word -> Save As -> .docx, then re-upload.`
    );
  }

  let doc;
  try {
    doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  } catch (err) {
    throw new Error("Docxtemplater init error: " + err.message);
  }

  try {
    doc.render(data);
  } catch (err) {
    throw new Error(`Docxtemplater render error: ${err.message}\nCheck that your data matches the template placeholders.`);
  }

  const buf = doc.getZip().generate({ type: "nodebuffer" });
  fs.writeFileSync(outputDocxPath, buf);
  return outputDocxPath;
}

module.exports = { validateDocxFile, generateDocx };
