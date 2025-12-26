/* server.js - Updated for Cloud Deployment (Koyeb/Render) */
const express = require("express");
const multer = require("multer");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const { parse } = require("csv-parse/sync");
const archiver = require("archiver");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { exec } = require("child_process");
const { PassThrough } = require("stream");

const canvas = require("canvas");
const assert = require("assert");
const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.js");

class NodeCanvasFactory {
  create(width, height) {
    assert(width > 0 && height > 0, "Invalid canvas size");
    const canvasInstance = canvas.createCanvas(width, height);
    const context = canvasInstance.getContext("2d");
    return { canvas: canvasInstance, context: context };
  }
  reset(canvasAndContext, width, height) {
    assert(canvasAndContext.canvas, "Canvas is not specified");
    assert(width > 0 && height > 0, "Invalid canvas size");
    canvasAndContext.canvas.width = width;
    canvasAndContext.canvas.height = height;
  }
  destroy(canvasAndContext) {
    assert(canvasAndContext.canvas, "Canvas is not specified");
    canvasAndContext.canvas.width = 0;
    canvasAndContext.canvas.height = 0;
    canvasAndContext.canvas = null;
    canvasAndContext.context = null;
  }
}

async function convertPdfToImages(pdfBuffer, format = "png") {
  const uint8Array = new Uint8Array(pdfBuffer);
  const loadingTask = pdfjsLib.getDocument({
    data: uint8Array,
    cMapUrl: "./node_modules/pdfjs-dist/cmaps/",
    cMapPacked: true,
    canvasFactory: new NodeCanvasFactory(),
  });

  const pdfDocument = await loadingTask.promise;
  const pageCount = pdfDocument.numPages;
  const images = [];

  for (let i = 1; i <= pageCount; i++) {
    const page = await pdfDocument.getPage(i);
    const viewport = page.getViewport({ scale: 2.0 });
    const canvasFactory = new NodeCanvasFactory();
    const canvasAndContext = canvasFactory.create(viewport.width, viewport.height);
    const renderContext = {
      canvasContext: canvasAndContext.context,
      viewport: viewport,
      canvasFactory: canvasFactory,
    };
    await page.render(renderContext).promise;
    let imgBuffer;
    if (format === "jpg" || format === "jpeg") {
      imgBuffer = canvasAndContext.canvas.toBuffer("image/jpeg", { quality: 0.9 });
    } else {
      imgBuffer = canvasAndContext.canvas.toBuffer("image/png");
    }
    images.push(imgBuffer);
    page.cleanup();
  }
  return images;
}

const app = express();

// --- แก้ไขจุดที่ 1: ใช้ Port จาก Environment Variable ---
const port = process.env.PORT || 8000; 

app.use(express.static("public"));
app.use(express.urlencoded({ extended: true }));
const upload = multer({ storage: multer.memoryStorage() });

function findSoffice() {
  const candidates = [
    "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return `"${p}"`;
  }
  return null;
}
let sofficePath = process.env.SOFFICE_PATH && fs.existsSync(process.env.SOFFICE_PATH) 
  ? `"${process.env.SOFFICE_PATH}"` 
  : findSoffice();

if (!sofficePath) {
  console.error("✘ ไม่พบ LibreOffice");
  process.exit(1);
}
console.log("✔ ใช้ LibreOffice ที่:", sofficePath);

function convertToPdf(inputBuffer, ext) {
  return new Promise((resolve, reject) => {
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "lo-"));
    const inputName = "input." + ext;
    const inputPath = path.join(tempDir, inputName);
    const outputPath = path.join(tempDir, "input.pdf");
    fs.writeFileSync(inputPath, inputBuffer);
    const cmd = `${sofficePath} --headless --convert-to pdf --outdir "${tempDir}" "${inputPath}"`;
    exec(cmd, (error) => {
      if (error) return reject(new Error("LibreOffice error"));
      try {
        const pdfData = fs.readFileSync(outputPath);
        resolve(pdfData);
      } catch (e) {
        reject(new Error("PDF Output not found"));
      } finally {
        try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch (_) {}
      }
    });
  });
}

function createZipBuffer(files) {
  return new Promise((resolve, reject) => {
    const archive = archiver("zip", { zlib: { level: 9 } });
    const stream = new PassThrough();
    const chunks = [];
    stream.on("data", (chunk) => chunks.push(chunk));
    stream.on("end", () => resolve(Buffer.concat(chunks)));
    archive.on("error", (err) => reject(err));
    archive.pipe(stream);
    files.forEach((f) => archive.append(f.buffer, { name: f.filename }));
    archive.finalize();
  });
}

app.post("/merge", upload.fields([{ name: "template" }, { name: "datafile" }]), async (req, res) => {
  try {
    // Check Missing Files
    if (!req.files.template || !req.files.datafile) {
      return res.status(400).json({ success: false, message: "ERR_MISSING_FILES" });
    }

    const templateFile = req.files.template[0];
    const dataFile = req.files.datafile[0];
    const templateExt = templateFile.originalname.split(".").pop().toLowerCase();
    
    // Check Valid Template
    if (!["docx", "pptx"].includes(templateExt)) {
        return res.status(400).json({ success: false, message: "ERR_INVALID_TEMPLATE" });
    }

    let outputType = (req.body.outputType || "pdf").toLowerCase();
    const allowedTypes = ["pdf", "docx", "pptx", "png", "jpg"];
    if (!allowedTypes.includes(outputType)) outputType = "pdf";

    // Check Template Mismatch
    if ((outputType === "docx" && templateExt !== "docx") || 
        (outputType === "pptx" && templateExt !== "pptx")) {
      return res.status(400).json({ success: false, message: "ERR_TEMPLATE_MISMATCH" });
    }

    const records = parse(dataFile.buffer.toString("utf-8"), { columns: true, skip_empty_lines: true, trim: true });
    
    // Check Empty CSV
    if (records.length === 0) return res.status(400).json({ success: false, message: "ERR_CSV_EMPTY" });

    const outFiles = [];
    for (let i = 0; i < records.length; i++) {
      const row = records[i];
      const zip = new PizZip(templateFile.buffer);
      const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, delimiters: { start: "{{", end: "}}" } });
      try { doc.render(row); } catch (error) { continue; }
      
      const filledBuffer = doc.getZip().generate({ type: "nodebuffer" });
      const firstHeaderKey = Object.keys(row)[0]; 
      const firstColValue = row[firstHeaderKey] ? row[firstHeaderKey].toString().trim() : "NoName";
      const indexStr = (i + 1).toString().padStart(2, '0');
      const safeName = firstColValue.replace(/[^a-zA-Z0-9ก-๙\s-]/g, "");
      const baseName = `${indexStr}_${safeName}`;

      if (outputType === "pdf") {
        const pdfBuf = await convertToPdf(filledBuffer, templateExt);
        outFiles.push({ filename: `${baseName}.pdf`, buffer: pdfBuf, ext: "pdf" });
      } else if (outputType === "docx" || outputType === "pptx") {
        outFiles.push({ filename: `${baseName}.${outputType}`, buffer: filledBuffer, ext: outputType });
      } else if (outputType === "png" || outputType === "jpg") {
        const pdfBuf = await convertToPdf(filledBuffer, templateExt);
        try {
          const imageBuffers = await convertPdfToImages(pdfBuf, outputType);
          imageBuffers.forEach((imgBuf, idx) => {
             const suffix = imageBuffers.length > 1 ? `_${idx + 1}` : "";
             outFiles.push({ filename: `${baseName}${suffix}.${outputType}`, buffer: imgBuf, ext: outputType });
          });
        } catch (imgErr) {
          throw new Error("Convert Image Error");
        }
      }
    }

    const zipBuffer = await createZipBuffer(outFiles);
    const filesJson = outFiles.map((f) => ({
      filename: f.filename,
      mime: f.ext === "pdf" ? "application/pdf" : 
            f.ext === "png" ? "image/png" : 
            f.ext === "jpg" ? "image/jpeg" : "application/octet-stream",
      base64: f.buffer.toString("base64"),
    }));

    res.json({ success: true, label: outputType.toUpperCase(), files: filesJson, zipBase64: zipBuffer.toString("base64") });

  } catch (err) {
    console.error(err);
    if (!res.headersSent) res.status(500).json({ success: false, message: "Error: " + err.message });
  }
});

// --- แก้ไขจุดที่ 2: เพิ่ม '0.0.0.0' เพื่อให้เข้าถึงได้จากภายนอก ---
app.listen(port, '0.0.0.0', () => {
  console.log(`Server running at http://0.0.0.0:${port}`);
});