const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType,
  ShadingType, BorderStyle } = require("docx");
const FONT = "Arial", MAG = "E6007E", W = 9638;
const t = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 26, bold: o.bold, color: o.color });
const rich = (s, o = {}) => s.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map(x => x.startsWith("**") ? t(x.slice(2, -2), { ...o, bold: true }) : t(x, o));
const para = (s, o = {}) => new Paragraph({ spacing: { after: o.after ?? 80 }, alignment: o.align, children: rich(s, o) });
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

function box(title, colour, lines) {
  const kids = [new Paragraph({ spacing: { after: 100 }, children: [t(title, { size: 32, bold: true, color: colour })] }),
    ...lines.map((l, i) => new Paragraph({ spacing: { after: 70 }, children: [t(`${i + 1}.  `, { bold: true, color: colour }), ...rich(l)] }))];
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ children: [new TableCell({
    width: { size: W, type: WidthType.DXA },
    borders: { top: none, bottom: none, right: none, left: { style: BorderStyle.SINGLE, size: 48, color: colour } },
    shading: { fill: "F5F5F5", type: ShadingType.CLEAR, color: "auto" },
    margins: { top: 140, bottom: 120, left: 260, right: 200 }, children: kids })] })] });
}
const gap = (a = 180) => new Paragraph({ spacing: { after: a }, children: [] });

function colours() {
  const rows = [["C8A165", "Coloured square", "A pallet is there"], ["F3EBDD", "Stripes", "Free space"],
    ["FF0000", "Red square", "Wrong zone: move it"], ["BFBFBF", "Grey", "Not ours: don't use"]];
  const cw = [700, 3000, 5938];
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: cw, rows: rows.map(([c, a, b]) => new TableRow({ children: [
    new TableCell({ width: { size: cw[0], type: WidthType.DXA }, shading: { fill: c, type: ShadingType.CLEAR, color: "auto" },
      borders: { top: { style: BorderStyle.SINGLE, size: 12, color: "FFFFFF" }, bottom: { style: BorderStyle.SINGLE, size: 12, color: "FFFFFF" }, left: none, right: none }, children: [new Paragraph("")] }),
    new TableCell({ width: { size: cw[1], type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: none }, margins: { left: 200, top: 60, bottom: 60 }, children: [para(a, { bold: true, after: 0 })] }),
    new TableCell({ width: { size: cw[2], type: WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: none }, margins: { top: 60, bottom: 60 }, children: [para(b, { after: 0 })] }),
  ] })) });
}

const doc = new Document({ styles: { default: { document: { run: { font: FONT, size: 26 } } } }, sections: [{
  properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 900, left: 1134, right: 1134 } } },
  children: [
    new Paragraph({ spacing: { after: 40 }, children: [t("Floor Map: how to use", { size: 48, bold: true })] }),
    new Paragraph({ spacing: { after: 260 }, border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: MAG, space: 6 } },
      children: [t("Open Floor_Map_Simple.xlsx. No macros, nothing to set up. Everything is counted in pallets.", { size: 26, color: "595959" })] }),
    box("PUT A PALLET AWAY", "00843D", ["Click a square on the **Map**.", "Click the **little arrow** next to it.", "Pick the item. The square fills with colour."]),
    gap(),
    box("TAKE A PALLET OUT", "C00000", ["Click the square.", "Press **Delete**. The square goes back to stripes."]),
    gap(),
    box("MOVE A PALLET", "0070C0", ["Take it out of the old square (**Delete**).", "Put it in the new square (**arrow**, then pick)."]),
    gap(),
    box("NEW ITEM?", MAG, ["Go to the **Stock** sheet.", "Type it on the first empty row: **code**, **description**, **zone**.", "It now shows up in the arrow list on the Map."]),
    gap(260),
    new Paragraph({ spacing: { after: 120 }, children: [t("What the colours mean", { size: 32, bold: true })] }),
    colours(),
    gap(200),
    para("**LOW** on the Stock sheet means you're nearly out: time to order.", { size: 24 }),
  ] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log("ok"); });
