import multer from "multer";
import htmlToDocx from "@turbodocx/html-to-docx";

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
      return resolve(result);
    });
  });
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).end();
  }

  await runMiddleware(req, res, upload.single("htmlFile"));

  const htmlContent = req.file.buffer.toString("utf-8");

  const docxBuffer = await htmlToDocx(htmlContent, null, {
    font: 'Arial',
    fontSize: 24,
    header: false,
    footer: false,
  })

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  );

  res.setHeader(
    "Content-Disposition",
    "attachment; filename=output.docx"
  );

  res.send(docxBuffer);
}