import chromium from "@sparticuz/chromium";
import puppeteer from "puppeteer-core";
import multer from "multer";

export const config = {
  api: {
    bodyParser: false,
  },
};

const upload = multer({ storage: multer.memoryStorage() });

function runMiddleware(req, res, fn) {
  return new Promise((resolve, reject) => {
    fn(req, res, (result) => {
      if (result instanceof Error) return reject(result);
      resolve(result);
    });
  });
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).end();
  }

  await runMiddleware(req, res, upload.single("htmlFile"));

  const html = req.file.buffer.toString("utf-8");

  const browser = await puppeteer.launch({
    args: [
      ...chromium.args,
      "--hide-scrollbars",
      "--disable-web-security"
    ],
    defaultViewport: chromium.defaultViewport,
    executablePath: await chromium.executablePath(),
    headless: true,
  });

  const page = await browser.newPage();

  await page.setContent(html, { waitUntil: "networkidle0" });

  const pdfBuffer = await page.pdf({
    format: "A4",
    printBackground: true,
  });

  await browser.close();

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader(
    "Content-Disposition",
    'attachment; filename="output.pdf"'
  );

  res.status(200).end(pdfBuffer);
}
