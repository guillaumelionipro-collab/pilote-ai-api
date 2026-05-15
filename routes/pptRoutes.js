import express from "express";
import multer from "multer";
import { importPptxTemplate } from "../services/pptImportEngine.js";
import { exportDeckToPptx } from "../services/pptExportEngine.js";

const router = express.Router();

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024,
  },
  fileFilter: (req, file, cb) => {
    if (!file.originalname.toLowerCase().endsWith(".pptx")) {
      return cb(new Error("Seuls les fichiers .pptx sont acceptés."));
    }
    cb(null, true);
  },
});

router.post("/import-template", upload.single("file"), async (req, res) => {
  try {
    if (!req.file?.buffer) {
      return res.status(400).json({ error: "Aucun fichier PowerPoint reçu." });
    }

    const result = await importPptxTemplate(req.file.buffer, req.file.originalname);
    res.json(result);
  } catch (error) {
    console.error("PPT import error:", error);
    res.status(500).json({
      error: "Erreur import PowerPoint.",
      details: error.message,
    });
  }
});

router.post("/export-pro", express.json({ limit: "50mb" }), async (req, res) => {
  try {
    const { fileName, buffer } = await exportDeckToPptx(req.body || {});

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
    res.send(buffer);
  } catch (error) {
    console.error("PPT export error:", error);
    res.status(500).json({
      error: "Erreur export PowerPoint.",
      details: error.message,
    });
  }
});

export default router;
