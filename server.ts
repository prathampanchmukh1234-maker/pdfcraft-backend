import express from "express";
import { createServer as createViteServer } from "vite";
import cors from "cors";
import multer from "multer";
import { createClient } from "@supabase/supabase-js";
import { PDFDocument, degrees, rgb, StandardFonts } from "pdf-lib";
import sharp from "sharp";
import archiver from "archiver";
import { v4 as uuidv4 } from "uuid";
import path from "path";
import fs from "fs";
import os from "os";
import { fileURLToPath } from "url";
import { execFile } from "child_process";
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";
import { Document, Packer, Paragraph, TextRun, AlignmentType } from "docx";
import ExcelJS from "exceljs";
import PptxGenJS from "pptxgenjs";
import mammoth from "mammoth";
import dotenv from "dotenv";

// Helper to extract structured data from PDF using pdfjs-dist
async function extractPdfData(buffer: Buffer) {
  const data = new Uint8Array(buffer);
  const loadingTask = pdfjs.getDocument({ 
    data,
    useSystemFonts: true,
    disableFontFace: true
  });
  const pdf = await loadingTask.promise;
  const numPages = pdf.numPages;
  const pagesData = [];

  for (let i = 1; i <= numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const viewport = page.getViewport({ scale: 1.0 });
    
    const items = textContent.items.map((item: any) => {
      const transform = item.transform;
      const fontSize = Math.sqrt(transform[0] * transform[0] + transform[1] * transform[1]);
      return {
        text: item.str,
        x: transform[4],
        y: viewport.height - transform[5], // Flip Y for standard coordinates
        width: item.width,
        height: item.height || fontSize,
        fontSize,
        fontName: item.fontName,
        hasEOL: item.hasEOL
      };
    });

    pagesData.push({
      pageNumber: i,
      width: viewport.width,
      height: viewport.height,
      items
    });
  }

  return pagesData;
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Load env reliably whether server is launched from repo root or backend directory.
const envCandidates = [
  path.resolve(process.cwd(), ".env"),
  path.resolve(__dirname, ".env"),
  path.resolve(__dirname, "..", ".env"),
];
for (const envPath of envCandidates) {
  if (fs.existsSync(envPath)) {
    dotenv.config({ path: envPath });
    break;
  }
}

// Supabase Setup
const supabaseUrl = process.env.VITE_SUPABASE_URL || "";
const supabaseServiceKey = process.env.SUPABASE_SERVICE_ROLE_KEY || "";
if (!supabaseUrl || !supabaseServiceKey) {
  console.error("Missing Supabase environment variables. Ensure VITE_SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are set.");
}
const supabase = createClient(supabaseUrl, supabaseServiceKey);

const app = express();
const PORT = process.env.PORT || 3000;

// Authentication Middleware
const authenticate = async (req: express.Request, res: express.Response, next: express.NextFunction) => {
  const authHeader = req.headers.authorization;
  if (!authHeader) {
    return res.status(401).json({ error: "No authorization header" });
  }

  const token = authHeader.split(" ")[1];
  const { data: { user }, error } = await supabase.auth.getUser(token);

  if (error || !user) {
    return res.status(401).json({ error: "Invalid or expired token" });
  }

  (req as any).user = user;
  next();
};

// CORS Setup
app.use(cors({
  origin: true,
  credentials: true
}));
app.use(express.json());

// Health check
app.get("/api/health", (req, res) => {
  res.json({ status: "ok" });
});

// Multer Setup
const upload = multer({ 
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024, // 50MB limit
  }
});

// Helper to ensure bucket exists and has correct configuration
async function ensureBucketExists() {
  const { data: buckets, error: listError } = await supabase.storage.listBuckets();
  if (listError) {
    console.error("Error listing buckets:", listError);
    throw new Error(`Failed to list storage buckets: ${listError.message}`);
  }

  const allowedMimeTypes = [
    "application/pdf", 
    "image/jpeg", 
    "image/png", 
    "application/zip",
    "application/x-zip-compressed",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  ];

  const bucketExists = buckets.some(b => b.name === "pdf-files");
  if (!bucketExists) {
    console.log("Creating 'pdf-files' bucket...");
    const { error: createError } = await supabase.storage.createBucket("pdf-files", {
      public: true,
      allowedMimeTypes,
      fileSizeLimit: 52428800 // 50MB
    });
    if (createError) {
      console.error("Error creating bucket:", createError);
      throw new Error(`Failed to create storage bucket: ${createError.message}`);
    } else {
      console.log("'pdf-files' bucket created successfully.");
    }
  } else {
    // Update existing bucket to ensure all MIME types are allowed
    console.log("Updating 'pdf-files' bucket configuration...");
    const { error: updateError } = await supabase.storage.updateBucket("pdf-files", {
      public: true,
      allowedMimeTypes,
      fileSizeLimit: 52428800
    });
    if (updateError) {
      console.error("Error updating bucket:", updateError);
      throw new Error(`Failed to update storage bucket: ${updateError.message}`);
    }
  }
}

// Helper to upload to Supabase Storage with user isolation
async function uploadToSupabase(buffer: Buffer, filename: string, mimeType: string, userId: string) {
  // Ensure bucket exists before upload
  await ensureBucketExists();

  const path = `users/${userId}/${uuidv4()}-${filename}`;
  const { data, error } = await supabase.storage
    .from("pdf-files")
    .upload(path, buffer, {
      contentType: mimeType,
      upsert: true,
    });

  if (error) {
    console.error("Supabase storage error:", error);
    throw new Error(`Supabase storage error: ${error.message}`);
  }
  
  const { data: { publicUrl } } = supabase.storage
    .from("pdf-files")
    .getPublicUrl(data.path);
    
  return publicUrl;
}

async function protectPdfWithQpdf(inputBuffer: Buffer, password: string): Promise<Buffer> {
  const tempId = uuidv4();
  const inputPath = path.join(os.tmpdir(), `${tempId}-input.pdf`);
  const outputPath = path.join(os.tmpdir(), `${tempId}-protected.pdf`);

  await fs.promises.writeFile(inputPath, inputBuffer);

  try {
    await new Promise<void>((resolve, reject) => {
      execFile(
        "qpdf",
        ["--encrypt", password, password, "256", "--", inputPath, outputPath],
        (error, _stdout, stderr) => {
          if (error) {
            const msg = stderr?.toString().trim() || error.message;
            reject(new Error(msg));
            return;
          }
          resolve();
        }
      );
    });

    return await fs.promises.readFile(outputPath);
  } finally {
    fs.promises.unlink(inputPath).catch(() => {});
    fs.promises.unlink(outputPath).catch(() => {});
  }
}

// Helper to log operation to history
async function logOperation(userId: string, type: string, originalName: string, url: string, size: number, status: string = 'completed', strict: boolean = false) {
  try {
    const processedName = `processed_${originalName}`;
    const { error } = await supabase.from("user_documents").insert({
      user_id: userId,
      service_type: type,
      original_file_name: originalName,
      processed_file_name: processedName,
      file_url: url,
      file_size: size,
      created_at: new Date().toISOString()
    });
    if (error) {
      throw new Error(error.message);
    }
  } catch (err) {
    console.error("Failed to log operation:", err);
    if (strict) {
      throw err;
    }
  }
}

// API Routes
app.post("/api/merge", authenticate, upload.array("files"), async (req, res) => {
  try {
    const files = req.files as Express.Multer.File[];
    const user = (req as any).user;

    if (!files || files.length < 2) {
      return res.status(400).json({ error: "At least two files are required for merging" });
    }

    const mergedPdf = await PDFDocument.create();
    for (const file of files) {
      const pdf = await PDFDocument.load(file.buffer);
      const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
      copiedPages.forEach((page) => mergedPdf.addPage(page));
    }

    const pdfBytes = await mergedPdf.save();
    const publicUrl = await uploadToSupabase(Buffer.from(pdfBytes), "merged.pdf", "application/pdf", user.id);

    await logOperation(user.id, "merge", files.map(f => f.originalname).join(", "), publicUrl, pdfBytes.length);

    res.json({ url: publicUrl });
  } catch (error: any) {
    console.error("Merge error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/split", authenticate, upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    const user = (req as any).user;
    const { pages } = req.body; // pages: "1-3"
    
    if (!file) return res.status(400).json({ error: "File is required" });

    const pdf = await PDFDocument.load(file.buffer);
    const pageCount = pdf.getPageCount();
    
    // Robust page selection logic
    const indices: number[] = [];
    const parts = pages.split(",");
    for (const part of parts) {
      if (part.includes("-")) {
        const [startStr, endStr] = part.split("-");
        const start = Math.max(0, (parseInt(startStr) || 1) - 1);
        const end = Math.min(pageCount - 1, (parseInt(endStr) || start + 1) - 1);
        for (let i = start; i <= end; i++) indices.push(i);
      } else {
        const pageNum = parseInt(part);
        if (!isNaN(pageNum)) {
          const idx = Math.max(0, Math.min(pageCount - 1, pageNum - 1));
          indices.push(idx);
        }
      }
    }
    
    if (indices.length === 0) indices.push(0);

    // If only one page, return PDF. If multiple, return ZIP.
    if (indices.length === 1) {
      const newPdf = await PDFDocument.create();
      const [copiedPage] = await newPdf.copyPages(pdf, [indices[0]]);
      newPdf.addPage(copiedPage);
      const pdfBytes = await newPdf.save();
      const publicUrl = await uploadToSupabase(Buffer.from(pdfBytes), `split-page-${indices[0] + 1}.pdf`, "application/pdf", user.id);
      await logOperation(user.id, "split", file.originalname, publicUrl, pdfBytes.length);
      return res.json({ url: publicUrl });
    } else {
      const zip = archiver('zip');
      const zipChunks: Buffer[] = [];
      
      // We need to collect the zip data into a buffer to upload to Supabase
      const zipBuffer = await new Promise<Buffer>((resolve, reject) => {
        const chunks: Buffer[] = [];
        zip.on('data', (chunk) => chunks.push(chunk));
        zip.on('end', () => resolve(Buffer.concat(chunks)));
        zip.on('error', reject);
        
        (async () => {
          for (const idx of indices) {
            const singlePagePdf = await PDFDocument.create();
            const [copiedPage] = await singlePagePdf.copyPages(pdf, [idx]);
            singlePagePdf.addPage(copiedPage);
            const bytes = await singlePagePdf.save();
            zip.append(Buffer.from(bytes), { name: `page-${idx + 1}.pdf` });
          }
          zip.finalize();
        })();
      });

      const publicUrl = await uploadToSupabase(zipBuffer, `split-pages.zip`, "application/zip", user.id);
      await logOperation(user.id, "split", file.originalname, publicUrl, zipBuffer.length);
      res.json({ url: publicUrl });
    }
  } catch (error: any) {
    console.error("Split error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/rotate", authenticate, upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    const user = (req as any).user;
    const { angle, pageIndex } = req.body; // angle: 90, 180, 270, pageIndex: optional
    
    if (!file) return res.status(400).json({ error: "File is required" });

    const pdf = await PDFDocument.load(file.buffer);
    const pages = pdf.getPages();
    
    if (pageIndex !== undefined && pageIndex !== null && pageIndex !== "") {
      const idx = parseInt(pageIndex);
      if (idx >= 0 && idx < pages.length) {
        const currentRotation = pages[idx].getRotation().angle;
        pages[idx].setRotation(degrees(currentRotation + Number(angle)));
      }
    } else {
      pages.forEach(page => {
        const currentRotation = page.getRotation().angle;
        page.setRotation(degrees(currentRotation + Number(angle)));
      });
    }

    const pdfBytes = await pdf.save();
    const publicUrl = await uploadToSupabase(Buffer.from(pdfBytes), "rotated.pdf", "application/pdf", user.id);

    await logOperation(user.id, "rotate", file.originalname, publicUrl, pdfBytes.length);

    res.json({ url: publicUrl });
  } catch (error: any) {
    console.error("Rotate error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/compress", authenticate, upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    const user = (req as any).user;
    
    if (!file) return res.status(400).json({ error: "File is required" });

    // Load with ignoreEncryption to handle more files
    const pdf = await PDFDocument.load(file.buffer, { ignoreEncryption: true });
    
    // Basic compression by re-saving with optimized settings
    // pdf-lib doesn't have advanced image recompression out of the box,
    // but useObjectStreams and other flags help.
    const pdfBytes = await pdf.save({ 
      useObjectStreams: true,
      addDefaultPage: false,
      updateFieldAppearances: false,
      objectsPerTick: 50
    });
    
    const publicUrl = await uploadToSupabase(Buffer.from(pdfBytes), "compressed.pdf", "application/pdf", user.id);

    await logOperation(user.id, "compress", file.originalname, publicUrl, pdfBytes.length);

    res.json({ url: publicUrl });
  } catch (error: any) {
    console.error("Compress error:", error);
    res.status(500).json({ error: error.message });
  }
});

async function processImageToPdf(files: Express.Multer.File[], userId: string) {
  const pdfDoc = await PDFDocument.create();
  for (const file of files) {
    let image;
    try {
      if (file.mimetype === 'image/png') {
        image = await pdfDoc.embedPng(file.buffer);
      } else {
        image = await pdfDoc.embedJpg(file.buffer);
      }
    } catch (e) {
      try {
        image = await pdfDoc.embedPng(file.buffer);
      } catch (e2) {
        image = await pdfDoc.embedJpg(file.buffer);
      }
    }
    const page = pdfDoc.addPage([image.width, image.height]);
    page.drawImage(image, { x: 0, y: 0, width: image.width, height: image.height });
  }
  const pdfBytes = await pdfDoc.save();
  return { pdfBytes, filename: "images-to-pdf.pdf" };
}

app.post("/api/image-to-pdf", authenticate, upload.array("files"), async (req, res) => {
  try {
    const files = req.files as Express.Multer.File[];
    const user = (req as any).user;
    if (!files || files.length === 0) return res.status(400).json({ error: "Images are required" });

    const { pdfBytes, filename } = await processImageToPdf(files, user.id);
    const publicUrl = await uploadToSupabase(Buffer.from(pdfBytes), filename, "application/pdf", user.id);
    await logOperation(user.id, "image-to-pdf", files.map(f => f.originalname).join(", "), publicUrl, pdfBytes.length);
    res.json({ url: publicUrl });
  } catch (error: any) {
    console.error("Image to PDF error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/jpg-to-pdf", authenticate, upload.array("files"), async (req, res) => {
  try {
    const files = req.files as Express.Multer.File[];
    const user = (req as any).user;
    if (!files || files.length === 0) return res.status(400).json({ error: "Images are required" });

    const { pdfBytes, filename } = await processImageToPdf(files, user.id);
    const publicUrl = await uploadToSupabase(Buffer.from(pdfBytes), filename, "application/pdf", user.id);
    await logOperation(user.id, "jpg-to-pdf", files.map(f => f.originalname).join(", "), publicUrl, pdfBytes.length);
    res.json({ url: publicUrl });
  } catch (error: any) {
    console.error("JPG to PDF error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/watermark", authenticate, upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    const user = (req as any).user;
    const { text } = req.body;
    
    if (!file) return res.status(400).json({ error: "File is required" });
    if (!text) return res.status(400).json({ error: "Watermark text is required" });

    const pdf = await PDFDocument.load(file.buffer);
    const pages = pdf.getPages();
    
    pages.forEach(page => {
      const { width, height } = page.getSize();
      page.drawText(text, {
        x: width / 4,
        y: height / 2,
        size: 50,
        opacity: 0.3,
        rotate: degrees(45),
        color: rgb(0.5, 0.5, 0.5)
      });
    });

    const pdfBytes = await pdf.save();
    const publicUrl = await uploadToSupabase(Buffer.from(pdfBytes), "watermarked.pdf", "application/pdf", user.id);

    await logOperation(user.id, "watermark", file.originalname, publicUrl, pdfBytes.length);

    res.json({ url: publicUrl });
  } catch (error: any) {
    console.error("Watermark error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/sign", authenticate, upload.fields([{ name: 'file', maxCount: 1 }, { name: 'signature', maxCount: 1 }]), async (req, res) => {
  try {
    const files = req.files as { [fieldname: string]: Express.Multer.File[] };
    const pdfFile = files['file']?.[0];
    const signatureFile = files['signature']?.[0];
    const user = (req as any).user;
    const { placedFields } = req.body;

    if (!pdfFile) return res.status(400).json({ error: "PDF file is required" });

    const pdfDoc = await PDFDocument.load(pdfFile.buffer);
    const pages = pdfDoc.getPages();

    if (placedFields) {
      const fields = JSON.parse(placedFields);
      for (const field of fields) {
        const fieldPageNum = field.pageNumber || 1;
        const targetPage = pages[Math.min(fieldPageNum - 1, pages.length - 1)];
        
        const pdfWidth = targetPage.getWidth();
        const pdfHeight = targetPage.getHeight();
        const screenWidth = 800; // Match frontend width
        const scaleFactor = pdfWidth / screenWidth;

        if (field.type === 'signature' && signatureFile) {
          try {
            const signatureImage = await pdfDoc.embedPng(signatureFile.buffer).catch(() => pdfDoc.embedJpg(signatureFile.buffer));
            const fieldScale = (field.scale || 1);
            const { width, height } = signatureImage.scale(fieldScale * scaleFactor * 0.4);
            
            const pdfX = field.x * scaleFactor;
            const pdfY = pdfHeight - (field.y * scaleFactor) - height;
            
            targetPage.drawImage(signatureImage, {
              x: pdfX,
              y: pdfY,
              width,
              height,
            });
          } catch (imgError) {
            console.error("Error embedding signature image:", imgError);
            targetPage.drawText(field.content || "Signed", {
              x: field.x * scaleFactor,
              y: pdfHeight - (field.y * scaleFactor) - 20,
              size: 20 * scaleFactor,
              color: rgb(0, 0, 0.5),
            });
          }
        } else {
          const fontSize = (field.type === 'initials' ? 30 : 16) * (field.scale || 1) * scaleFactor;
          targetPage.drawText(field.content || " ", {
            x: field.x * scaleFactor,
            y: pdfHeight - (field.y * scaleFactor) - fontSize,
            size: fontSize,
            color: rgb(0, 0, 0.5),
          });
        }
      }
    }

    const pdfBytes = await pdfDoc.save();
    const publicUrl = await uploadToSupabase(Buffer.from(pdfBytes), "signed.pdf", "application/pdf", user.id);

    await logOperation(user.id, "sign", pdfFile.originalname, publicUrl, pdfBytes.length);

    res.json({ url: publicUrl });
  } catch (error: any) {
    console.error("Sign error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/pdf-to-word", authenticate, upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    const user = (req as any).user;
    if (!file) return res.status(400).json({ error: "File is required" });

    console.log("Starting high-fidelity PDF to Word conversion for:", file.originalname);
    const pagesData = await extractPdfData(file.buffer);
    
    const sections = pagesData.map(page => {
      // Group items into lines based on Y coordinate with a tolerance
      const lines: any[][] = [];
      const sortedItems = [...page.items].sort((a, b) => a.y - b.y || a.x - b.x);
      
      if (sortedItems.length === 0) return { children: [new Paragraph({ children: [] })] };

      let currentLine: any[] = [];
      let lastY = sortedItems[0].y;
      
      sortedItems.forEach(item => {
        if (Math.abs(item.y - lastY) < 8) {
          currentLine.push(item);
        } else {
          lines.push(currentLine.sort((a, b) => a.x - b.x));
          currentLine = [item];
          lastY = item.y;
        }
      });
      if (currentLine.length > 0) lines.push(currentLine.sort((a, b) => a.x - b.x));

      // Convert lines to paragraphs
      const children = lines.map(line => {
        const text = line.map(item => item.text).join(" ").trim();
        if (!text) return null;

        // Basic alignment detection
        let alignment: any = AlignmentType.LEFT;
        const lineWidth = line[line.length - 1].x + line[line.length - 1].width - line[0].x;
        if (Math.abs(line[0].x + lineWidth / 2 - page.width / 2) < 50) {
          alignment = AlignmentType.CENTER;
        }

        return new Paragraph({
          children: [new TextRun({ 
            text, 
            size: Math.round(line[0].fontSize * 2),
            bold: line[0].fontName?.toLowerCase().includes('bold')
          })],
          alignment,
          spacing: { before: 120, after: 120 },
        });
      }).filter(p => p !== null) as Paragraph[];

      return {
        properties: {
          page: {
            size: {
              width: Math.round(page.width * 20),
              height: Math.round(page.height * 20),
            }
          }
        },
        children
      };
    });

    const doc = new Document({ sections });
    const buffer = await Packer.toBuffer(doc);
    const publicUrl = await uploadToSupabase(buffer, `${path.parse(file.originalname).name}.docx`, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", user.id);

    await logOperation(user.id, "pdf-to-word", file.originalname, publicUrl, buffer.length);
    res.json({ url: publicUrl });
  } catch (error: any) {
    console.error("PDF to Word error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/pdf-to-excel", authenticate, upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    const user = (req as any).user;
    if (!file) return res.status(400).json({ error: "File is required" });

    console.log("Starting high-fidelity PDF to Excel conversion for:", file.originalname);
    const pagesData = await extractPdfData(file.buffer);
    const workbook = new ExcelJS.Workbook();
    
    pagesData.forEach((page, pageIdx) => {
      const worksheet = workbook.addWorksheet(`Page ${pageIdx + 1}`);
      
      // Group items into rows with a tolerance
      const rows: any[][] = [];
      const sortedItems = [...page.items].sort((a, b) => a.y - b.y || a.x - b.x);
      
      if (sortedItems.length === 0) return;

      let currentRow: any[] = [];
      let lastY = sortedItems[0].y;
      
      sortedItems.forEach(item => {
        if (Math.abs(item.y - lastY) < 10) {
          currentRow.push(item);
        } else {
          rows.push(currentRow.sort((a, b) => a.x - b.x));
          currentRow = [item];
          lastY = item.y;
        }
      });
      if (currentRow.length > 0) rows.push(currentRow.sort((a, b) => a.x - b.x));

      // For each row, try to place items into columns based on X position
      rows.forEach(row => {
        const rowData: string[] = [];
        let lastX = 0;
        row.forEach(item => {
          // If there's a significant gap, skip some columns
          const gap = item.x - lastX;
          const skipCols = Math.max(0, Math.floor(gap / 50));
          for (let i = 0; i < skipCols; i++) rowData.push("");
          
          rowData.push(item.text);
          lastX = item.x + item.width;
        });
        worksheet.addRow(rowData);
      });
    });

    const buffer = await workbook.xlsx.writeBuffer() as Buffer;
    const publicUrl = await uploadToSupabase(buffer, `${path.parse(file.originalname).name}.xlsx`, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", user.id);

    await logOperation(user.id, "pdf-to-excel", file.originalname, publicUrl, buffer.length);
    res.json({ url: publicUrl });
  } catch (error: any) {
    console.error("PDF to Excel error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/pdf-to-powerpoint", authenticate, upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    const user = (req as any).user;
    if (!file) return res.status(400).json({ error: "File is required" });

    console.log("Starting high-fidelity PDF to PowerPoint conversion for:", file.originalname);
    const pagesData = await extractPdfData(file.buffer);
    const pres = new PptxGenJS();
    
    pagesData.forEach(page => {
      const slide = pres.addSlide();
      // Group items into lines to reduce number of text boxes
      const lines: any[][] = [];
      const sortedItems = [...page.items].sort((a, b) => a.y - b.y || a.x - b.x);
      
      let currentLine: any[] = [];
      let lastY = -1;
      
      sortedItems.forEach(item => {
        if (lastY === -1 || Math.abs(item.y - lastY) < 5) {
          currentLine.push(item);
        } else {
          lines.push(currentLine);
          currentLine = [item];
        }
        lastY = item.y;
      });
      if (currentLine.length > 0) lines.push(currentLine);

      lines.forEach(line => {
        const text = line.map(item => item.text).join(" ");
        const firstItem = line[0];
        slide.addText(text, {
          x: (firstItem.x / page.width) * 10, // Convert to inches (assuming 10 inch width)
          y: (firstItem.y / page.height) * 7.5, // Convert to inches (assuming 7.5 inch height)
          fontSize: firstItem.fontSize,
          color: "000000"
        });
      });
    });

    const buffer = await pres.write({ outputType: "nodebuffer" }) as Buffer;
    const publicUrl = await uploadToSupabase(buffer, `${path.parse(file.originalname).name}.pptx`, "application/vnd.openxmlformats-officedocument.presentationml.presentation", user.id);

    await logOperation(user.id, "pdf-to-powerpoint", file.originalname, publicUrl, buffer.length);
    res.json({ url: publicUrl });
  } catch (error: any) {
    console.error("PDF to PowerPoint error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/word-to-pdf", authenticate, upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    const user = (req as any).user;
    if (!file) return res.status(400).json({ error: "File is required" });

    const result = await mammoth.extractRawText({ buffer: file.buffer });
    const pdfDoc = await PDFDocument.create();
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const page = pdfDoc.addPage();
    const { height } = page.getSize();
    
    result.value.split('\n').forEach((line, index) => {
      if (index < 30) {
        page.drawText(line, { x: 50, y: height - 50 - index * 20, size: 12, font });
      }
    });

    const pdfBytes = await pdfDoc.save();
    const publicUrl = await uploadToSupabase(Buffer.from(pdfBytes), `${path.parse(file.originalname).name}.pdf`, "application/pdf", user.id);

    await logOperation(user.id, "word-to-pdf", file.originalname, publicUrl, pdfBytes.length);
    res.json({ url: publicUrl });
  } catch (error: any) {
    console.error("Word to PDF error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/excel-to-pdf", authenticate, upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    const user = (req as any).user;
    if (!file) return res.status(400).json({ error: "File is required" });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(file.buffer);
    const worksheet = workbook.getWorksheet(1);
    
    const pdfDoc = await PDFDocument.create();
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const page = pdfDoc.addPage();
    const { height } = page.getSize();
    
    let y = height - 50;
    worksheet?.eachRow((row, rowNumber) => {
      if (rowNumber < 30) {
        const text = row.values ? (row.values as any[]).join(" | ") : "";
        page.drawText(text, { x: 50, y, size: 10, font });
        y -= 20;
      }
    });

    const pdfBytes = await pdfDoc.save();
    const publicUrl = await uploadToSupabase(Buffer.from(pdfBytes), `${path.parse(file.originalname).name}.pdf`, "application/pdf", user.id);

    await logOperation(user.id, "excel-to-pdf", file.originalname, publicUrl, pdfBytes.length);
    res.json({ url: publicUrl });
  } catch (error: any) {
    console.error("Excel to PDF error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/pdf-to-jpg", authenticate, upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    const user = (req as any).user;
    if (!file) return res.status(400).json({ error: "File is required" });

    console.log("Starting high-fidelity PDF to JPG conversion for:", file.originalname);
    const pagesData = await extractPdfData(file.buffer);
    const zip = archiver('zip');
    
    const zipBuffer = await new Promise<Buffer>((resolve, reject) => {
      const chunks: Buffer[] = [];
      zip.on('data', (chunk) => chunks.push(chunk));
      zip.on('end', () => resolve(Buffer.concat(chunks)));
      zip.on('error', reject);
      
      (async () => {
        for (let i = 0; i < Math.min(pagesData.length, 10); i++) {
          const page = pagesData[i];
          
          // Create a SVG representation of the page
          const svg = `
            <svg width="${page.width}" height="${page.height}" xmlns="http://www.w3.org/2000/svg">
              <rect width="100%" height="100%" fill="white"/>
              ${page.items.map(item => `
                <text x="${item.x}" y="${item.y}" font-size="${item.fontSize}" font-family="Arial" fill="black">
                  ${item.text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}
                </text>
              `).join('')}
            </svg>
          `;
          
          const jpgBuffer = await sharp(Buffer.from(svg))
            .jpeg({ quality: 90 })
            .toBuffer();
          
          zip.append(jpgBuffer, { name: `page-${i + 1}.jpg` });
        }
        zip.finalize();
      })();
    });

    const publicUrl = await uploadToSupabase(zipBuffer, "extracted-images.zip", "application/zip", user.id);
    await logOperation(user.id, "pdf-to-jpg", file.originalname, publicUrl, zipBuffer.length);
    res.json({ url: publicUrl });

  } catch (error: any) {
    console.error("PDF to JPG error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/powerpoint-to-pdf", authenticate, upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    const user = (req as any).user;
    if (!file) return res.status(400).json({ error: "File is required" });

    // Basic text extraction from PPTX is hard without specialized libraries
    // We'll provide a placeholder or a basic implementation if possible
    const pdfDoc = await PDFDocument.create();
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const page = pdfDoc.addPage();
    page.drawText("PowerPoint to PDF conversion is currently limited to text extraction.", { x: 50, y: 700, size: 12, font });
    
    const pdfBytes = await pdfDoc.save();
    const publicUrl = await uploadToSupabase(Buffer.from(pdfBytes), `${path.parse(file.originalname).name}.pdf`, "application/pdf", user.id);

    await logOperation(user.id, "powerpoint-to-pdf", file.originalname, publicUrl, pdfBytes.length);
    res.json({ url: publicUrl });
  } catch (error: any) {
    console.error("PowerPoint to PDF error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/protect", authenticate, upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    const user = (req as any).user;
    const { password } = req.body;
    if (!file) return res.status(400).json({ error: "File is required" });
    if (!password) return res.status(400).json({ error: "Password is required" });

    const pdfBytes = await protectPdfWithQpdf(file.buffer, password);
    const publicUrl = await uploadToSupabase(Buffer.from(pdfBytes), `protected-${file.originalname}`, "application/pdf", user.id);

    await logOperation(user.id, "protect", file.originalname, publicUrl, pdfBytes.length);
    res.json({ url: publicUrl, message: "PDF protected successfully." });
  } catch (error: any) {
    console.error("Protect error:", error);
    const message = String(error?.message || "");
    if (message.toLowerCase().includes("not recognized") || message.toLowerCase().includes("enoent")) {
      return res.status(500).json({ error: "qpdf is not installed on the server. Install qpdf to enable PDF password protection." });
    }
    res.status(500).json({ error: message || "Failed to protect PDF" });
  }
});

app.post("/api/unlock", authenticate, upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    const user = (req as any).user;
    const { password } = req.body;
    if (!file) return res.status(400).json({ error: "File is required" });

    // pdf-lib does not support unlocking.
    // We'll return the original file with a message.
    const pdfDoc = await PDFDocument.load(file.buffer);
    const pdfBytes = await pdfDoc.save();
    const publicUrl = await uploadToSupabase(Buffer.from(pdfBytes), `unlocked-${file.originalname}`, "application/pdf", user.id);

    await logOperation(user.id, "unlock", file.originalname, publicUrl, pdfBytes.length);
    res.json({ url: publicUrl, message: "Unlock is currently simulated. The file is saved but not decrypted." });
  } catch (error: any) {
    console.error("Unlock error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/history", authenticate, async (req, res) => {
  try {
    const user = (req as any).user;
    const { data, error } = await supabase
      .from("user_documents")
      .select("*")
      .eq("user_id", user.id)
      .order("created_at", { ascending: false });

    if (error) {
      console.error("History fetch error:", JSON.stringify(error, null, 2));
      if (error.code === 'PGRST116' || error.message?.includes('relation "user_documents" does not exist')) {
        return res.status(500).json({ 
          error: "Database table 'user_documents' not found. Please ensure you have created the table in your Supabase dashboard as per the instructions." 
        });
      }
      return res.status(500).json({ 
        error: error.message,
        details: error.details,
        hint: error.hint,
        code: error.code
      });
    }
    res.json(data || []);
  } catch (error: any) {
    console.error("History route error:", error);
    res.status(500).json({ 
      error: error.message || "An unknown error occurred",
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
});

app.post("/api/upload-processed", authenticate, upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    const user = (req as any).user;
    const { operationType, originalName } = req.body;

    if (!file) return res.status(400).json({ error: "File is required" });

    const publicUrl = await uploadToSupabase(file.buffer, originalName || "processed.pdf", file.mimetype, user.id);

    await logOperation(user.id, operationType || "client-tool", originalName || "document.pdf", publicUrl, file.size, 'completed', true);

    res.json({ url: publicUrl });
  } catch (error: any) {
    console.error("Upload processed error:", error);
    res.status(500).json({ error: error.message });
  }
});

app.delete("/api/history/:id", authenticate, async (req, res) => {
  try {
    const user = (req as any).user;
    const { error } = await supabase
      .from("user_documents")
      .delete()
      .eq("id", req.params.id)
      .eq("user_id", user.id);

    if (error) throw error;
    res.json({ status: "ok" });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

// Vite middleware for development (serves frontend)
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
      root: path.join(__dirname, "..", "frontend"),
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.join(__dirname, "..", "dist")));
    app.get("*", (req, res) => {
      res.sendFile(path.join(__dirname, "..", "dist", "index.html"));
    });
  }

  app.listen(Number(PORT), "0.0.0.0", () => {
    console.log(`Backend server v1.0.1 running on port ${PORT}`);
  });
}

startServer();
