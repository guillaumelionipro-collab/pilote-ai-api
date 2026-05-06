import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import OpenAI from "openai";
import fs from "fs";
import ExcelJS from "exceljs";
import PDFDocument from "pdfkit";

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3001;
const MODEL = process.env.OPENAI_MODEL || "gpt-4o-mini";

app.use(cors());
app.use(express.json({ limit: "25mb" }));

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

const memoryFile = "memory-pro.json";

let memory = {
  profil: [],
  objectifs: [],
  projets: [],
  preferences: [],
  applications: [],
  faits_importants: [],
};

if (fs.existsSync(memoryFile)) {
  try {
    memory = JSON.parse(fs.readFileSync(memoryFile, "utf-8"));
  } catch {
    console.log("Mémoire invalide, reset mémoire.");
  }
}

function saveMemory() {
  fs.writeFileSync(memoryFile, JSON.stringify(memory, null, 2));
}

function formatMemory() {
  return `
PROFIL:
${memory.profil.join("\n")}

OBJECTIFS:
${memory.objectifs.join("\n")}

PROJETS:
${memory.projets.join("\n")}

PREFERENCES:
${memory.preferences.join("\n")}

APPLICATIONS:
${memory.applications.join("\n")}

FAITS IMPORTANTS:
${memory.faits_importants.join("\n")}
`;
}

function safeJson(text) {
  try {
    return JSON.parse(text);
  } catch {
    const match = text?.match(/\{[\s\S]*\}/);
    if (!match) return null;

    try {
      return JSON.parse(match[0]);
    } catch {
      return null;
    }
  }
}

async function analyseMemory(message) {
  try {
    const completion = await openai.chat.completions.create({
      model: MODEL,
      messages: [
        {
          role: "system",
          content: `
Analyse si le message contient une information durable à mémoriser.

Réponds uniquement en JSON valide :
{
  "shouldRemember": true/false,
  "category": "profil" | "objectifs" | "projets" | "preferences" | "applications" | "faits_importants",
  "memory": "phrase courte"
}
`,
        },
        { role: "user", content: message },
      ],
    });

    return safeJson(completion.choices[0].message.content) || {
      shouldRemember: false,
      category: "faits_importants",
      memory: "",
    };
  } catch {
    return {
      shouldRemember: false,
      category: "faits_importants",
      memory: "",
    };
  }
}

function addMemory(category, content) {
  if (!memory[category] || !content) return;

  const exists = memory[category].some(
    (item) => item.toLowerCase() === content.toLowerCase()
  );

  if (!exists) {
    memory[category].push(content);
    saveMemory();
  }
}

app.get("/", (req, res) => {
  res.json({
    status: "ok",
    message: "Pilot AI API active",
  });
});

app.get("/memory-pro", (req, res) => {
  res.json(memory);
});

app.delete("/memory-pro", (req, res) => {
  memory = {
    profil: [],
    objectifs: [],
    projets: [],
    preferences: [],
    applications: [],
    faits_importants: [],
  };

  saveMemory();

  res.json({ success: true });
});

app.post("/chat", async (req, res) => {
  try {
    const { message } = req.body;

    if (!message) {
      return res.json({
        reply: JSON.stringify({
          type: "chat",
          response: "Aucun message reçu.",
          data: {},
        }),
      });
    }

    const lower = message.toLowerCase().trim();

    if (lower === "/memory") {
      return res.json({
        reply: JSON.stringify({
          type: "chat",
          response: formatMemory(),
          data: {},
        }),
      });
    }

    if (lower === "/clear-memory") {
      memory = {
        profil: [],
        objectifs: [],
        projets: [],
        preferences: [],
        applications: [],
        faits_importants: [],
      };

      saveMemory();

      return res.json({
        reply: JSON.stringify({
          type: "chat",
          response: "Mémoire supprimée.",
          data: {},
        }),
      });
    }

    const memoryAnalysis = await analyseMemory(message);

    if (memoryAnalysis.shouldRemember && memoryAnalysis.memory) {
      addMemory(memoryAnalysis.category, memoryAnalysis.memory);
    }

    const completion = await openai.chat.completions.create({
      model: MODEL,
      messages: [
        {
          role: "system",
          content: `
Tu es Pilot AI, assistant personnel intelligent de Guillaume.

Tu dois répondre uniquement en JSON valide.

Format obligatoire :
{
  "type": "chat" | "excel" | "pdf" | "image",
  "response": "réponse claire",
  "data": {
    "title": "",
    "subtitle": "",
    "content": "",
    "prompt": "",
    "rows": []
  }
}

Règles :
- demande Excel/tableau/budget/suivi chiffré = type "excel"
- demande PDF/rapport/procédure/note/document = type "pdf"
- demande image/logo/visuel/mockup = type "image"
- sinon = type "chat"
- Aucun texte hors JSON.
- Réponds en français.

Mémoire :
${formatMemory()}
`,
        },
        { role: "user", content: message },
      ],
    });

    const raw = completion.choices[0].message.content;
    const parsed = safeJson(raw);

    res.json({
      reply: JSON.stringify(
        parsed || {
          type: "chat",
          response: raw || "Je suis prêt Guillaume.",
          data: {},
        }
      ),
    });
  } catch (error) {
    console.error("Erreur /chat :", error.message);

    res.json({
      reply: JSON.stringify({
        type: "chat",
        response: "Erreur serveur IA.",
        data: {},
      }),
    });
  }
});

app.post("/pilot", async (req, res) => {
  try {
    const { app: appName, mode, message, data } = req.body;

    const completion = await openai.chat.completions.create({
      model: MODEL,
      messages: [
        {
          role: "system",
          content: `
Tu es Pilot AI connecté aux applications de Guillaume.

Application : ${appName || "pilot-ai"}
Mode : ${mode || "general"}

Réponds uniquement en JSON valide :
{
  "type": "chat" | "excel" | "pdf" | "image",
  "response": "réponse claire",
  "data": {
    "title": "",
    "subtitle": "",
    "content": "",
    "prompt": "",
    "rows": []
  }
}

Si mode = metier-vitrage : réponds comme un référent technique vitrage automobile.
Si mode = runesis : réponds comme un coach performance running/trail.
`,
        },
        {
          role: "user",
          content: `
Message :
${message}

Données application :
${JSON.stringify(data || {}, null, 2)}
`,
        },
      ],
    });

    const raw = completion.choices[0].message.content;
    const parsed = safeJson(raw);

    res.json({
      reply: JSON.stringify(
        parsed || {
          type: "chat",
          response: raw || "Pilot est prêt.",
          data: {},
        }
      ),
    });
  } catch (error) {
    console.error("Erreur /pilot :", error.message);

    res.json({
      reply: JSON.stringify({
        type: "chat",
        response: "Erreur Pilot API.",
        data: {},
      }),
    });
  }
});

app.post("/create-excel", async (req, res) => {
  try {
    const { title, rows } = req.body;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Pilot AI");

    sheet.mergeCells("A1:D1");
    sheet.getCell("A1").value = title || "Tableau Pilot AI";
    sheet.getCell("A1").font = { size: 18, bold: true, color: { argb: "FFFFFFFF" } };
    sheet.getCell("A1").fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF0B1020" },
    };
    sheet.getCell("A1").alignment = { horizontal: "center" };

    sheet.addRow([]);

    const headerRow = sheet.addRow(["Catégorie", "Description", "Montant", "Statut"]);

    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF2563EB" },
      };
    });

    const dataRows =
      rows?.length > 0
        ? rows
        : [
            ["Revenus", "Salaire", 0, "À compléter"],
            ["Charges", "Loyer", 0, "À compléter"],
            ["Épargne", "Objectif", 0, "À compléter"],
            ["Reste à vivre", "Calcul", 0, "À compléter"],
          ];

    dataRows.forEach((row) => sheet.addRow(row));

    sheet.columns.forEach((col) => {
      col.width = 30;
    });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", "attachment; filename=pilot-ai.xlsx");

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error("Erreur Excel :", error.message);
    res.status(500).json({ error: "Erreur création Excel" });
  }
});

app.post("/create-pdf", async (req, res) => {
  try {
    const { title, subtitle, content } = req.body;

    const doc = new PDFDocument({ size: "A4", margin: 50 });

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", "attachment; filename=pilot-ai.pdf");

    doc.pipe(res);

    doc.rect(0, 0, 595, 120).fill("#0B1020");

    doc
      .fillColor("#FFFFFF")
      .fontSize(24)
      .font("Helvetica-Bold")
      .text(title || "Document Pilot AI", 50, 38);

    doc
      .fillColor("#67E8F9")
      .fontSize(11)
      .text(subtitle || "Document généré par Pilot AI", 50, 72);

    doc
      .fillColor("#111827")
      .fontSize(12)
      .font("Helvetica")
      .text(content || "Document généré automatiquement.", 50, 150, {
        width: 500,
        lineGap: 6,
      });

    doc.end();
  } catch (error) {
    console.error("Erreur PDF :", error.message);
    res.status(500).json({ error: "Erreur création PDF" });
  }
});

app.post("/create-image", async (req, res) => {
  try {
    const { prompt } = req.body;

    const image = await openai.images.generate({
      model: "gpt-image-1",
      prompt: prompt || "Image futuriste Pilot AI style Jarvis premium",
      size: "1024x1024",
    });

    res.json({
      imageUrl: `data:image/png;base64,${image.data[0].b64_json}`,
    });
  } catch (error) {
    console.error("Erreur image :", error.message);

    res.status(500).json({
      error: "Erreur génération image",
    });
  }
});

app.listen(PORT, () => {
  console.log(`PILOT AI API active on port ${PORT}`);
});