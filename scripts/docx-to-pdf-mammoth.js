// scripts/docx-to-pdf-mammoth.js
const fs = require("fs");
const mammoth = require("mammoth");
const puppeteer = require("puppeteer");

/**
 * Convert a filled DOCX -> PDF using mammoth + puppeteer
 * @param {string} docxPath
 * @param {string} outputPdfPath
 */
async function convertDocxToPdfUsingMammoth(docxPath, outputPdfPath) {
  if (!fs.existsSync(docxPath)) throw new Error("DOCX not found: " + docxPath);

  // 1) Convert DOCX to HTML using mammoth
  const result = await mammoth.convertToHtml({ path: docxPath });
  let html = result.value;
  const warnings = result.messages || [];
  if ((warnings || []).length) console.warn("Mammoth warnings:", warnings);

  // 2) Wrap HTML with minimal styles for printing
  const fullHtml = `
  <!doctype html>
  <html>
    <head>
      <meta charset="utf-8"/>
      <meta name="viewport" content="width=device-width, initial-scale=1"/>
      <style>
        body { font-family: "Times New Roman", serif; font-size: 12pt; margin: 28mm 20mm; color: #111; }
        p { margin: 0 0 8px; line-height: 1.45; }
        table { border-collapse: collapse; width: 100%; }
        table td, table th { border: 1px solid #ccc; padding: 6px; }
        .signature-line { margin-top: 30px; }
        /* preserve blank underlines */
        u { text-decoration: underline; }
      </style>
    </head>
    <body>${html}</body>
  </html>
  `;

  // 3) Use puppeteer to render HTML -> PDF
  // Note: launching puppeteer may take time on first run (Chromium download).
  const browser = await puppeteer.launch({
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
    headless: "new"
  });

  try {
    const page = await browser.newPage();
    await page.setContent(fullHtml, { waitUntil: "networkidle0", timeout: 0 });

    await page.pdf({
      path: outputPdfPath,
      format: "A4",
      printBackground: true,
      margin: { top: "18mm", bottom: "18mm", left: "15mm", right: "15mm" }
    });

    return outputPdfPath;
  } finally {
    await browser.close();
  }
}

module.exports = { convertDocxToPdfUsingMammoth };
