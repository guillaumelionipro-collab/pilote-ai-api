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

app.post("/api/meetings/transcribe", async (req, res) => {
  try {
    const { audioUrl } = req.body;

    if (!audioUrl) {
      return res.status(400).json({ error: "audioUrl manquante" });
    }

    const audioResponse = await fetch(audioUrl);

    if (!audioResponse.ok) {
      return res.status(400).json({
        error: "Impossible de récupérer le fichier audio",
      });
    }

    const arrayBuffer = await audioResponse.arrayBuffer();
    const audioBuffer = Buffer.from(arrayBuffer);

    const file = new File([audioBuffer], "reunion-centre-pilot.webm", {
      type: "audio/webm",
    });

    const transcription = await openai.audio.transcriptions.create({
      file,
      model: "gpt-4o-mini-transcribe",
      language: "fr",
      prompt: `
Contexte : réunion professionnelle du GIE Vitrage Auto Service.
Vocabulaire métier à reconnaître et préserver :
GIE VAS, Centre Pilot, Nico, Guillaume, MAEL, agenda MAEL, OR, ordre de réparation,
ADAS, calibration, pare-brise, vitrage automobile, SAV, Fiche de synthèse,
DSPC, MAIF, MACIF, ETAI, XGlass, Darva, Sidexa, facturation, pré-bilan,
rapport centre, responsable centre, procédure, audit, formation, qualité réseau.
La transcription doit être en français professionnel, ponctuée, lisible, avec les termes métier correctement écrits.
      `,
    });

    res.json({
      transcription: transcription.text,
    });
  } catch (error) {
    console.error("Erreur transcription premium réunion :", error);
    res.status(500).json({
      error: "Erreur transcription premium réunion",
      details: error.message,
    });
  }
});

app.post("/api/meetings/analyze", async (req, res) => {
  try {
    const { transcription } = req.body;

    if (!transcription) {
      return res.status(400).json({
        error: "Transcription manquante"
      });
    }

    const response = await openai.responses.create({
  model: MODEL,
  input: `
Tu es l'assistant exécutif premium du GIE Vitrage Auto Service.

MISSION :
Transformer une transcription brute de réunion en document professionnel clair, fluide, naturel et exploitable par une direction opérationnelle.

OBJECTIF PRIORITAIRE :
Produire une retranscription ultra naturelle, comme si une secrétaire de direction avait repris les notes :
- texte fluide ;
- ponctuation complète ;
- paragraphes propres ;
- suppression des hésitations ;
- correction des erreurs de reconnaissance vocale ;
- reformulation légère sans trahir le sens ;
- vocabulaire métier respecté ;
- aucun remplissage inutile.

CONTEXTE GIE :
Réseau : GIE Vitrage Auto Service / GIE VAS
Application : Centre Pilot
Assistant : Pilot AI
Référent : Guillaume
Direction : Nico / Monsieur Alibert

VOCABULAIRE MÉTIER À CONSERVER :
MAEL
agenda MAEL
banette
point de proximité
Sidexa
Darva
ETAI
XGlass
MACIF
MAIF
DSPC
OR
ordre de réparation
fiche de synthèse
prise en charge
pré-bilan
facturation
contrôle de gestion
ADAS
calibration
caméra
capteur
pare-brise
vitrage automobile
SAV
tour de véhicule
véhicule de courtoisie
audit
qualité
formation
non-conformité
traçabilité
procédure
pilotage
performance
CODIR
DG
direction

STYLE INTERDIT :
Ne pas écrire :
- "Lors de la réunion"
- "Plusieurs points ont été abordés"
- "Il a été évoqué"
- "Cette réunion a permis de"

STYLE ATTENDU :
Écrire de manière directe, humaine, professionnelle et opérationnelle.

Exemple attendu :
"Point de suivi sur le centre de Muret.

Le SAV reste en hausse depuis deux semaines. Une vérification des dossiers MAIF est nécessaire, notamment sur les photos, la DSPC et les OR signés.

Guillaume prendra l’analyse de Toulouse avant vendredi."

DÉTECTION DES ENGAGEMENTS :
Identifier les engagements verbaux :
- "je m’en occupe"
- "je prends le sujet"
- "je regarde Toulouse"
- "on fait ça jeudi"
- "Guillaume s’en charge"
- "Nico valide"
- "le responsable doit vérifier"

DÉTECTION DES ALERTES :
Identifier les alertes :
- retard
- urgence
- incident
- SAV
- erreur
- blocage
- non-conformité
- dossier incomplet
- risque assurance
- risque facturation
- problème qualité
- dérive centre

RÈGLES :
- Ne jamais inventer.
- Si une information manque, écrire "À définir".
- Les actions doivent être concrètes.
- Les alertes doivent être courtes et utiles.
- Les engagements doivent préciser personne + engagement + échéance si présente.

Retourne uniquement un JSON valide, sans texte autour.

STRUCTURE JSON EXACTE :
{
  "corrected_transcription": "",
  "executive_summary": "",
  "summary": "",
  "decisions": "",
  "risks": "",
  "next_steps": "",
  "urgency_level": "faible|moyen|élevé|critique",
  "confidence_score": 0,
  "alerts": [
    {
      "level": "faible|moyen|élevé|critique",
      "message": "",
      "topic": ""
    }
  ],
  "commitments": [
    {
      "person": "",
      "commitment": "",
      "due_date": "",
      "topic": ""
    }
  ],
  "actions": [
    {
      "action": "",
      "responsible": "",
      "due_date": "",
      "priority": "Basse|Moyenne|Haute|Critique",
      "centre": "",
      "topic": ""
    }
  ]
}

FORMAT DU executive_summary :
POINT DE SITUATION

Constat :
...

Décisions :
• ...
• ...

Alertes :
• ...

Engagements :
• ...

Prochaine étape :
...

TRANSCRIPTION BRUTE :
${transcription}
  `
});

    res.json({
      result: response.output_text
    });

  } catch (error) {
    console.error("Erreur analyse:", error);
    res.status(500).json({
      error: "Erreur analyse"
    });
  }
});

app.listen(PORT, () => {
  console.log(`PILOT AI API active on port ${PORT}`);
});