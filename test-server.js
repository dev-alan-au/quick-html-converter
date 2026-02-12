import express from "express";
import multer from "multer";
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import htmlToDocx from "@turbodocx/html-to-docx";
import puppeteer from "puppeteer-core";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.get("/", (req, res) => {
  res.send(`
    <h1>Simple HTML converter</h1>
    <h2>Upload HTML to convert to DOCX</h2>
    <form method="POST" action="/convert-to-docx" enctype="multipart/form-data">
      <div>
        <label for="fileName">File name</label>
        <input id="fileName" type="text" name="fileName" />
      </div>
      <div>
        <label for="htmlFile">File to convert</label>
        <input id="htmlFile" type="file" name="htmlFile" accept=".html" required />
      </div>
      <button type="submit">Convert</button>
    </form>
    <h2>Upload HTML to convert to PDF</h2>
    <form action="/convert-to-pdf" method="POST" enctype="multipart/form-data">
      <input type="file" name="htmlFile" accept=".html" required />
      <button type="submit">Convert to PDF</button>
    </form>
  `);
});

app.post("/convert-to-docx", upload.single("htmlFile"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send("No file uploaded.");
    }

    const fileName = req.body.fileName ?? 'output';

    const htmlContent = req.file.buffer.toString("utf-8");

    const docxBuffer = await htmlToDocx(htmlContent);

    res.set({
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": `attachment; filename=${fileName}.docx`,
    });

    res.send(docxBuffer);
  } catch (error) {
    console.error(error);
    res.status(500).send("Conversion failed.");
  }
});

app.post("/convert-to-pdf", upload.single("htmlFile"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send("No file uploaded.");
    }

    const htmlContent = req.file.buffer.toString("utf-8");

    const browser = await puppeteer.launch({  
          headless: true,
          defaultViewport: null,
          executablePath: "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
      });

    const page = await browser.newPage();

    await page.setContent(htmlContent, {
      waitUntil: "networkidle0",
    });

    const pdfBuffer = await page.pdf({
      format: "A4",
      printBackground: true,
      margin: {
        top: "20mm",
        right: "20mm",
        bottom: "20mm",
        left: "20mm",
      },
    });

    await browser.close();

    res.set({
      "Content-Type": "application/pdf",
      "Content-Disposition": "attachment; filename=output.pdf",
    });

    res.send(pdfBuffer);
  } catch (error) {
    console.error(error);
    res.status(500).send("PDF conversion failed.");
  }
});

const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
