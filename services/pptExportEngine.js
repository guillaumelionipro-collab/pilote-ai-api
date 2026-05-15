import pptxgen from "pptxgenjs";
import { v4 as uuidv4 } from "uuid";

const SLIDE_W_PX = 1280;
const SLIDE_H_PX = 720;
const SLIDE_W_IN = 13.333;
const SLIDE_H_IN = 7.5;

function pxToIn(px, totalPx, totalIn) {
  return (Number(px || 0) / totalPx) * totalIn;
}

function cleanColor(value, fallback = "10182F") {
  if (!value || value === "transparent") return fallback;
  return String(value).replace("#", "").toUpperCase();
}

function safeText(value) {
  return String(value ?? "")
    .replace(/\u202F/g, " ")
    .replace(/\u00A0/g, " ")
    .replace(/[–—]/g, "-")
    .replace(/[“”]/g, '"')
    .replace(/[’]/g, "'");
}

export async function exportDeckToPptx(deckData = {}) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = deckData.author || "Centre Pilot";
  pptx.company = "GIE Auto Vitrage Service";
  pptx.subject = deckData.subject || "Formation";
  pptx.title = deckData.title || "formation-centre-pilot";

  const slides = Array.isArray(deckData.slides) ? deckData.slides : [];

  slides.forEach((s) => {
    const slide = pptx.addSlide();
    slide.background = { color: cleanColor(s.background, "FFFFFF") };

    const elements = Array.isArray(s.elements) ? s.elements : [];

    elements.forEach((el) => {
      const x = pxToIn(el.x, SLIDE_W_PX, SLIDE_W_IN);
      const y = pxToIn(el.y, SLIDE_H_PX, SLIDE_H_IN);
      const w = pxToIn(el.w, SLIDE_W_PX, SLIDE_W_IN);
      const h = pxToIn(el.h, SLIDE_H_PX, SLIDE_H_IN);

      if (el.type === "image" && el.src) {
        slide.addImage({
          data: el.src,
          x,
          y,
          w,
          h,
        });
        return;
      }

      if (el.type === "shape") {
        slide.addShape(pptx.ShapeType.roundRect, {
          x,
          y,
          w,
          h,
          fill: { color: cleanColor(el.bg, "FFFFFF") },
          line: { color: "DCE3F0", transparency: 35 },
          radius: 0.18,
        });
      }

      if (el.text) {
        slide.addText(safeText(el.text), {
          x,
          y,
          w,
          h,
          fontSize: Math.max(8, Math.round(Number(el.fontSize || 22) * 0.48)),
          bold: !!el.bold,
          color: cleanColor(el.color, "10182F"),
          align: el.align || "left",
          valign: "mid",
          fit: "shrink",
          margin: 0.08,
          breakLine: false,
        });
      }
    });
  });

  const fileName = `${deckData.title || "formation"}-${uuidv4()}.pptx`;
  const buffer = await pptx.write({ outputType: "nodebuffer" });

  return {
    fileName,
    buffer,
  };
}
