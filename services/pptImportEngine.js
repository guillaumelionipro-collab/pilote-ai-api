import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { v4 as uuidv4 } from "uuid";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  textNodeName: "text",
});

function asArray(value) {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function getTextFromShape(shape) {
  const paragraphs = asArray(shape?.["p:txBody"]?.["a:p"]);
  const lines = [];

  paragraphs.forEach((p) => {
    const runs = asArray(p?.["a:r"]);
    const text = runs
      .map((r) => r?.["a:t"])
      .filter(Boolean)
      .join("");
    if (text.trim()) lines.push(text.trim());
  });

  return lines.join("\n");
}

function emuToPx(value) {
  const n = Number(value || 0);
  return Math.round(n / 9525);
}

function extractPosition(shape) {
  const xfrm =
    shape?.["p:spPr"]?.["a:xfrm"] ||
    shape?.["p:pic"]?.["p:spPr"]?.["a:xfrm"] ||
    shape?.["p:grpSpPr"]?.["a:xfrm"];

  return {
    x: emuToPx(xfrm?.["a:off"]?.x),
    y: emuToPx(xfrm?.["a:off"]?.y),
    w: emuToPx(xfrm?.["a:ext"]?.cx) || 500,
    h: emuToPx(xfrm?.["a:ext"]?.cy) || 100,
  };
}

function extractShapeFill(shape) {
  const solidFill = shape?.["p:spPr"]?.["a:solidFill"];
  const srgb = solidFill?.["a:srgbClr"]?.val;
  if (srgb) return `#${srgb}`;
  return "transparent";
}

function extractTextColor(shape) {
  const paragraphs = asArray(shape?.["p:txBody"]?.["a:p"]);
  const firstRun = paragraphs?.[0]?.["a:r"];
  const solidFill = firstRun?.["a:rPr"]?.["a:solidFill"];
  const srgb = solidFill?.["a:srgbClr"]?.val;
  if (srgb) return `#${srgb}`;
  return "#10182f";
}

function extractFontSize(shape) {
  const paragraphs = asArray(shape?.["p:txBody"]?.["a:p"]);
  const firstRun = paragraphs?.[0]?.["a:r"];
  const sz = firstRun?.["a:rPr"]?.sz;
  if (!sz) return 28;
  return Math.round(Number(sz) / 100);
}

async function readXml(zip, path) {
  const file = zip.file(path);
  if (!file) return null;
  const xml = await file.async("string");
  return parser.parse(xml);
}

async function extractTheme(zip) {
  const themeXml = await readXml(zip, "ppt/theme/theme1.xml");
  const clrScheme = themeXml?.["a:theme"]?.["a:themeElements"]?.["a:clrScheme"];
  const colors = [];

  if (clrScheme) {
    Object.values(clrScheme).forEach((entry) => {
      const srgb = entry?.["a:srgbClr"]?.val;
      if (srgb) colors.push(`#${srgb}`);
    });
  }

  const fontScheme = themeXml?.["a:theme"]?.["a:themeElements"]?.["a:fontScheme"];
  const fonts = [
    fontScheme?.["a:majorFont"]?.["a:latin"]?.typeface,
    fontScheme?.["a:minorFont"]?.["a:latin"]?.typeface,
  ].filter(Boolean);

  return {
    colors: [...new Set(colors)],
    fonts: [...new Set(fonts)],
  };
}

async function extractRelationships(zip, slideNumber) {
  const relPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
  const relXml = await readXml(zip, relPath);
  const rels = asArray(relXml?.Relationships?.Relationship);
  const map = {};

  rels.forEach((rel) => {
    map[rel.Id] = rel.Target;
  });

  return map;
}

async function extractMedia(zip, target) {
  if (!target) return null;
  const cleanTarget = target.replace("../", "ppt/");
  const mediaPath = cleanTarget.startsWith("ppt/") ? cleanTarget : `ppt/${cleanTarget}`;
  const file = zip.file(mediaPath);
  if (!file) return null;

  const ext = mediaPath.split(".").pop()?.toLowerCase() || "png";
  const mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "svg" ? "image/svg+xml" : "image/png";
  const base64 = await file.async("base64");

  return `data:${mime};base64,${base64}`;
}

async function extractSlide(zip, slideNumber) {
  const slidePath = `ppt/slides/slide${slideNumber}.xml`;
  const slideXml = await readXml(zip, slidePath);
  if (!slideXml) return null;

  const rels = await extractRelationships(zip, slideNumber);
  const cSld = slideXml?.["p:sld"]?.["p:cSld"];
  const spTree = cSld?.["p:spTree"];

  const shapes = asArray(spTree?.["p:sp"]);
  const pics = asArray(spTree?.["p:pic"]);

  const elements = [];

  shapes.forEach((shape) => {
    const text = getTextFromShape(shape);
    const pos = extractPosition(shape);
    const fill = extractShapeFill(shape);

    if (text) {
      elements.push({
        id: uuidv4(),
        type: fill !== "transparent" ? "shape" : "text",
        x: pos.x,
        y: pos.y,
        w: pos.w,
        h: pos.h,
        text,
        fontSize: extractFontSize(shape),
        bold: true,
        color: extractTextColor(shape),
        bg: fill,
        radius: fill !== "transparent" ? 20 : 0,
        align: "left",
      });
    } else if (fill !== "transparent") {
      elements.push({
        id: uuidv4(),
        type: "shape",
        x: pos.x,
        y: pos.y,
        w: pos.w,
        h: pos.h,
        text: "",
        fontSize: 18,
        bold: false,
        color: "#10182f",
        bg: fill,
        radius: 20,
        align: "left",
      });
    }
  });

  for (const pic of pics) {
    const blip = pic?.["p:blipFill"]?.["a:blip"];
    const embed = blip?.embed;
    const target = rels[embed];
    const src = await extractMedia(zip, target);
    const pos = extractPosition({ "p:pic": pic });

    if (src) {
      elements.push({
        id: uuidv4(),
        type: "image",
        x: pos.x,
        y: pos.y,
        w: pos.w,
        h: pos.h,
        text: "",
        fontSize: 18,
        bold: false,
        color: "#10182f",
        bg: "transparent",
        radius: 0,
        align: "center",
        src,
      });
    }
  }

  return {
    id: uuidv4(),
    name: `Slide ${slideNumber}`,
    background: "#ffffff",
    elements,
    source: {
      path: slidePath,
      slideNumber,
    },
  };
}

export async function importPptxTemplate(buffer, fileName = "template.pptx") {
  const zip = await JSZip.loadAsync(buffer);
  const presentationXml = await readXml(zip, "ppt/presentation.xml");
  const slideIds = asArray(
    presentationXml?.["p:presentation"]?.["p:sldIdLst"]?.["p:sldId"]
  );

  const slideCount = slideIds.length || zip.file(/ppt\/slides\/slide\d+\.xml/).length;
  const theme = await extractTheme(zip);

  const slides = [];
  for (let i = 1; i <= slideCount; i++) {
    const slide = await extractSlide(zip, i);
    if (slide) slides.push(slide);
  }

  return {
    templateName: fileName,
    importedAt: new Date().toISOString(),
    theme,
    meta: {
      slideCount: slides.length,
      engine: "openxml-v1",
    },
    slides,
  };
}
