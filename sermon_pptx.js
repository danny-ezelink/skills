const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "The Root Problem";

// ─── COLOR PALETTE ───────────────────────────────────────────────────────────
const BG_DARK   = "0D0F1A";   // deep midnight
const BG_CARD   = "161929";   // card background
const GOLD      = "C9A84C";   // primary gold
const GOLD_LITE = "E8D5A3";   // light gold
const WHITE     = "FFFFFF";
const MUTED     = "8B92A5";
const RED_ACC   = "8B2635";   // deep crimson accent (blood)
const CREAM     = "F5EDD6";

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function makeShadow() {
  return { type: "outer", color: "000000", blur: 10, offset: 4, angle: 135, opacity: 0.5 };
}
function makeGoldBar(slide, x, y, w = 0.05, h = 0.7) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: GOLD }, line: { color: GOLD } });
}
function addBg(slide) {
  slide.background = { color: BG_DARK };
}
function addScriptureBug(slide, ref) {
  slide.addText(ref, {
    x: 0.35, y: 5.1, w: 9.3, h: 0.3,
    fontSize: 11, fontFace: "Calibri", color: GOLD, align: "right",
    italic: true, margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 1 — TITLE
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);

  // top accent bar full width
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: GOLD }, line: { color: GOLD } });
  // bottom accent bar
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.54, w: 10, h: 0.08, fill: { color: GOLD }, line: { color: GOLD } });

  // sermon series tag
  s.addText("SERMON SERIES  ·  GRACE, IDENTITY & THE GOSPEL", {
    x: 0.5, y: 0.28, w: 9, h: 0.3,
    fontSize: 10, fontFace: "Calibri", color: GOLD, align: "center",
    bold: true, charSpacing: 3, margin: 0,
  });

  // main title
  s.addText("The Root\nProblem", {
    x: 0.8, y: 0.85, w: 8.4, h: 2.6,
    fontSize: 72, fontFace: "Georgia", color: WHITE,
    bold: true, align: "center", margin: 0,
    shadow: makeShadow(),
  });

  // gold divider line
  s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 3.5, w: 3.0, h: 0.04, fill: { color: GOLD }, line: { color: GOLD } });

  // subtitle
  s.addText("What Nobody Told You About Sin", {
    x: 0.5, y: 3.7, w: 9, h: 0.55,
    fontSize: 22, fontFace: "Georgia", color: GOLD_LITE,
    italic: true, align: "center", margin: 0,
  });

  // scripture ref
  s.addText("Romans 5:19  ·  Romans 5:12  ·  2 Corinthians 5:17", {
    x: 0.5, y: 4.5, w: 9, h: 0.3,
    fontSize: 12, fontFace: "Calibri", color: MUTED, align: "center", margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 2 — KEY TEXT: Romans 5:19
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: BG_DARK }, line: { color: BG_DARK } });

  // large decorative quote mark
  s.addText("“", {
    x: 0.2, y: -0.1, w: 2, h: 1.8,
    fontSize: 160, fontFace: "Georgia", color: GOLD,
    bold: true, margin: 0, transparency: 70,
  });

  s.addText([
    { text: "For as by one man’s disobedience\nmany were ", options: { color: CREAM, fontSize: 26, fontFace: "Georgia", italic: true } },
    { text: "made sinners,", options: { color: GOLD, fontSize: 26, fontFace: "Georgia", italic: true, bold: true } },
    { text: "\nso by the obedience of one shall\nmany be ", options: { color: CREAM, fontSize: 26, fontFace: "Georgia", italic: true } },
    { text: "made righteous.", options: { color: GOLD, fontSize: 26, fontFace: "Georgia", italic: true, bold: true } },
  ], {
    x: 0.7, y: 0.9, w: 8.6, h: 3.5,
    align: "center", valign: "middle", margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 3.8, y: 4.6, w: 2.4, h: 0.04, fill: { color: GOLD }, line: { color: GOLD } });
  s.addText("Romans 5:19  (KJV)", {
    x: 0.5, y: 4.75, w: 9, h: 0.3,
    fontSize: 13, fontFace: "Calibri", color: GOLD, align: "center",
    italic: true, bold: true, margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 3 — THE HOOK
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("ASK THE ROOM:", {
    x: 0.5, y: 0.35, w: 9, h: 0.35,
    fontSize: 12, fontFace: "Calibri", color: GOLD, align: "center",
    bold: true, charSpacing: 4, margin: 0,
  });

  s.addText([
    { text: "“How many of you have ever tried to\n", options: {} },
    { text: "stop doing something wrong", options: { color: GOLD, bold: true } },
    { text: "— and ", options: {} },
    { text: "failed?", options: { color: GOLD, bold: true } },
    { text: "”", options: {} },
  ], {
    x: 0.5, y: 1.0, w: 9, h: 2.8,
    fontSize: 38, fontFace: "Georgia", color: WHITE,
    italic: true, align: "center", valign: "middle", margin: 0,
    shadow: makeShadow(),
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.54, w: 10, h: 0.06, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("You promised. You tried. You failed. You felt worse.", {
    x: 0.5, y: 4.5, w: 9, h: 0.45,
    fontSize: 16, fontFace: "Calibri", color: MUTED,
    italic: true, align: "center", margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 4 — BIG STATEMENT: Diagnosis Problem
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: "080A12" }, line: { color: "080A12" } });

  s.addText("You didn’t have a", {
    x: 0.5, y: 0.8, w: 9, h: 0.9,
    fontSize: 34, fontFace: "Georgia", color: MUTED,
    italic: true, align: "center", margin: 0,
  });
  s.addText("DISCIPLINE\nPROBLEM.", {
    x: 0.5, y: 1.6, w: 9, h: 1.8,
    fontSize: 72, fontFace: "Georgia", color: WHITE,
    bold: true, align: "center", margin: 0, shadow: makeShadow(),
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 3.55, w: 7, h: 0.04, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("You had a", {
    x: 0.5, y: 3.72, w: 9, h: 0.6,
    fontSize: 28, fontFace: "Georgia", color: MUTED,
    italic: true, align: "center", margin: 0,
  });
  s.addText("DIAGNOSIS PROBLEM.", {
    x: 0.5, y: 4.28, w: 9, h: 0.75,
    fontSize: 36, fontFace: "Georgia", color: GOLD,
    bold: true, align: "center", margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 5 — POINT 1: Sin Is a Condition
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("POINT 1", {
    x: 0.5, y: 0.25, w: 9, h: 0.3,
    fontSize: 11, fontFace: "Calibri", color: GOLD,
    bold: true, charSpacing: 5, align: "center", margin: 0,
  });

  s.addText("Sin Is Not What\nYou Think It Is", {
    x: 0.5, y: 0.7, w: 9, h: 1.9,
    fontSize: 52, fontFace: "Georgia", color: WHITE,
    bold: true, align: "center", margin: 0, shadow: makeShadow(),
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.72, w: 9, h: 0.04, fill: { color: GOLD }, line: { color: GOLD } });

  // two column truth statement
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.9, w: 4.1, h: 1.9, fill: { color: BG_CARD }, line: { color: "2A2E42" } });
  s.addText([
    { text: "RELIGION SAYS:", options: { color: GOLD, bold: true, breakLine: true } },
    { text: "“sin less”\n“do better”\n“try harder”", options: { color: MUTED, italic: true } },
  ], {
    x: 0.65, y: 3.0, w: 3.8, h: 1.6,
    fontSize: 16, fontFace: "Calibri", align: "center", valign: "middle", margin: 6,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 2.9, w: 4.1, h: 1.9, fill: { color: "1A0E0E" }, line: { color: RED_ACC } });
  s.addText([
    { text: "THE TRUTH:", options: { color: GOLD, bold: true, breakLine: true } },
    { text: "Sin is not an action.\nSin is a ", options: { color: CREAM } },
    { text: "condition.", options: { color: GOLD, bold: true } },
  ], {
    x: 5.55, y: 3.0, w: 3.8, h: 1.6,
    fontSize: 16, fontFace: "Calibri", align: "center", valign: "middle", margin: 6,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 6 — KEY TRUTH: Born a Sinner
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);

  s.addText("You don’t become a sinner", {
    x: 0.5, y: 0.55, w: 9, h: 0.7,
    fontSize: 28, fontFace: "Georgia", color: MUTED, italic: true,
    align: "center", margin: 0,
  });
  s.addText("because you sinned.", {
    x: 0.5, y: 1.2, w: 9, h: 0.85,
    fontSize: 36, fontFace: "Georgia", color: WHITE, bold: true,
    align: "center", margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 3.0, y: 2.15, w: 4, h: 0.05, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("You sin because you were", {
    x: 0.5, y: 2.35, w: 9, h: 0.7,
    fontSize: 28, fontFace: "Georgia", color: MUTED, italic: true,
    align: "center", margin: 0,
  });
  s.addText("born a sinner.", {
    x: 0.5, y: 3.0, w: 9, h: 0.9,
    fontSize: 50, fontFace: "Georgia", color: GOLD, bold: true,
    align: "center", margin: 0, shadow: makeShadow(),
  });

  s.addText("“Many were MADE sinners.”  — Romans 5:19", {
    x: 0.5, y: 4.4, w: 9, h: 0.45,
    fontSize: 14, fontFace: "Calibri", color: MUTED, italic: true,
    align: "center", margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 7 — ILLUSTRATION: Pulling Weeds
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("ILLUSTRATION", {
    x: 0.3, y: 0.25, w: 3, h: 0.3,
    fontSize: 10, fontFace: "Calibri", color: GOLD,
    bold: true, charSpacing: 4, margin: 0,
  });

  s.addText("The Weed Root", {
    x: 0.3, y: 0.65, w: 6, h: 0.8,
    fontSize: 42, fontFace: "Georgia", color: WHITE, bold: true,
    margin: 0, shadow: makeShadow(),
  });

  const lines = [
    "Treating sin by fixing behavior is like",
    "pulling weeds and leaving the root in the ground.",
    "",
    "You spend all afternoon yanking them out.",
    "The yard looks clean. You feel accomplished.",
    "",
    "But you didn’t kill it.",
    "You just separated it temporarily.",
    "",
    "The root is still there.",
    "And it will always grow back.",
  ];

  let yPos = 1.6;
  for (const line of lines) {
    if (line === "") { yPos += 0.15; continue; }
    const isKey = line.includes("root") && line.includes("still");
    const isKey2 = line.includes("always grow back");
    s.addText(line, {
      x: 0.3, y: yPos, w: 9.3, h: 0.34,
      fontSize: isKey || isKey2 ? 17 : 15,
      fontFace: isKey || isKey2 ? "Georgia" : "Calibri",
      color: isKey || isKey2 ? GOLD_LITE : CREAM,
      bold: isKey || isKey2,
      italic: !isKey && !isKey2,
      margin: 0,
    });
    yPos += isKey || isKey2 ? 0.38 : 0.32;
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 8 — POINT 2: Bloodline Problem + Romans 5:12
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("POINT 2", {
    x: 0.5, y: 0.22, w: 9, h: 0.3,
    fontSize: 11, fontFace: "Calibri", color: GOLD,
    bold: true, charSpacing: 5, align: "center", margin: 0,
  });
  s.addText("It’s a Bloodline Problem", {
    x: 0.5, y: 0.62, w: 9, h: 1.0,
    fontSize: 46, fontFace: "Georgia", color: WHITE, bold: true,
    align: "center", margin: 0, shadow: makeShadow(),
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 1.74, w: 8.4, h: 1.45, fill: { color: BG_CARD }, line: { color: "2A2E42" } });
  s.addText([
    { text: "“Wherefore, as by one man sin entered into the world,\nand death by sin; and so ", options: { italic: true } },
    { text: "death passed upon all men", options: { italic: true, bold: true, color: GOLD } },
    { text: ",\nfor that all have sinned.”", options: { italic: true } },
  ], {
    x: 1.0, y: 1.84, w: 8.0, h: 1.2,
    fontSize: 15.5, fontFace: "Georgia", color: CREAM,
    align: "center", valign: "middle", margin: 4,
  });
  addScriptureBug(s, "Romans 5:12  (KJV)");

  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 3.3, w: 8.4, h: 0.04, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText([
    { text: "Adam’s disobedience didn’t just set a bad example.\nIt ", options: {} },
    { text: "infected the entire bloodline.", options: { color: GOLD, bold: true } },
  ], {
    x: 0.5, y: 3.45, w: 9, h: 1.4,
    fontSize: 21, fontFace: "Georgia", color: CREAM,
    italic: true, align: "center", valign: "middle", margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 9 — ILLUSTRATION: Baby & Cookies
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("ILLUSTRATION", {
    x: 0.3, y: 0.25, w: 4, h: 0.3,
    fontSize: 10, fontFace: "Calibri", color: GOLD, bold: true, charSpacing: 4, margin: 0,
  });
  s.addText("Nobody Taught Them to Lie", {
    x: 0.3, y: 0.6, w: 9.3, h: 0.8,
    fontSize: 36, fontFace: "Georgia", color: WHITE, bold: true, margin: 0, shadow: makeShadow(),
  });

  s.addText("Why do babies never need to be taught how to lie?", {
    x: 0.3, y: 1.55, w: 9.3, h: 0.5,
    fontSize: 17, fontFace: "Calibri", color: CREAM, italic: true, margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 2.15, w: 9.3, h: 1.5, fill: { color: BG_CARD }, line: { color: "2A2E42" } });
  s.addText("A two-year-old — cookie crumbs on their face, chocolate on\ntheir hands — looks right at you and says:", {
    x: 0.5, y: 2.25, w: 9.0, h: 0.7,
    fontSize: 15, fontFace: "Calibri", color: CREAM, italic: true, align: "center", margin: 0,
  });
  s.addText("“I didn’t do it.”", {
    x: 0.5, y: 2.95, w: 9.0, h: 0.55,
    fontSize: 26, fontFace: "Georgia", color: GOLD, bold: true, italic: true, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.77, w: 9.3, h: 0.04, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText([
    { text: "That’s not learned behavior.", options: { color: MUTED } },
    { text: "  That’s ", options: { color: CREAM } },
    { text: "inherited nature.", options: { color: GOLD, bold: true } },
  ], {
    x: 0.3, y: 3.95, w: 9.3, h: 0.55,
    fontSize: 18, fontFace: "Georgia", italic: true, align: "center", margin: 0,
  });

  s.addText("The root was already there.", {
    x: 0.3, y: 4.6, w: 9.3, h: 0.4,
    fontSize: 15, fontFace: "Calibri", color: MUTED, italic: true, align: "center", margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 10 — POINT 3: Jesus Came to Replace
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("POINT 3", {
    x: 0.5, y: 0.22, w: 9, h: 0.3,
    fontSize: 11, fontFace: "Calibri", color: GOLD,
    bold: true, charSpacing: 5, align: "center", margin: 0,
  });

  s.addText("Jesus Didn’t Come\nto Improve You", {
    x: 0.5, y: 0.62, w: 9, h: 1.55,
    fontSize: 46, fontFace: "Georgia", color: WHITE, bold: true,
    align: "center", margin: 0, shadow: makeShadow(),
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 2.26, w: 7, h: 0.04, fill: { color: GOLD }, line: { color: GOLD } });

  // two columns
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.42, w: 4.1, h: 1.75, fill: { color: BG_CARD }, line: { color: "2A2E42" } });
  s.addText([
    { text: "Every Religion:", options: { color: MUTED, bold: true, breakLine: true } },
    { text: "“Here’s what you must\nDO to become acceptable.”", options: { color: CREAM, italic: true } },
  ], {
    x: 0.65, y: 2.52, w: 3.8, h: 1.5,
    fontSize: 15, fontFace: "Calibri", align: "center", valign: "middle", margin: 6,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 2.42, w: 4.1, h: 1.75, fill: { color: "0D1A0D" }, line: { color: "2A5C2A" } });
  s.addText([
    { text: "The Gospel:", options: { color: GOLD, bold: true, breakLine: true } },
    { text: "“Here’s what was\nDONE for you —\nbecause you couldn’t.”", options: { color: CREAM, italic: true } },
  ], {
    x: 5.55, y: 2.52, w: 3.8, h: 1.5,
    fontSize: 15, fontFace: "Calibri", align: "center", valign: "middle", margin: 6,
  });

  s.addText("He came to make ", {
    x: 0.5, y: 4.3, w: 5, h: 0.55,
    fontSize: 20, fontFace: "Georgia", color: MUTED, italic: true,
    align: "right", margin: 0,
  });
  s.addText("dead people alive.", {
    x: 5.3, y: 4.3, w: 4.2, h: 0.55,
    fontSize: 20, fontFace: "Georgia", color: GOLD, bold: true,
    align: "left", margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 11 — 2 Cor 5:17 — NEW CREATURE
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);

  s.addText("“", {
    x: 0.1, y: -0.2, w: 2, h: 1.8,
    fontSize: 160, fontFace: "Georgia", color: GOLD,
    bold: true, margin: 0, transparency: 70,
  });

  s.addText([
    { text: "Therefore if any man be in Christ,\nhe is a ", options: { color: CREAM, italic: true } },
    { text: "new creature:", options: { color: GOLD, bold: true, italic: true } },
    { text: "\nold things are passed away;\nbehold, ", options: { color: CREAM, italic: true } },
    { text: "all things are become new.", options: { color: GOLD, bold: true, italic: true } },
  ], {
    x: 0.6, y: 0.7, w: 8.8, h: 3.3,
    fontSize: 26, fontFace: "Georgia",
    align: "center", valign: "middle", margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 4.1, w: 3.0, h: 0.04, fill: { color: GOLD }, line: { color: GOLD } });
  s.addText("2 Corinthians 5:17  (KJV)", {
    x: 0.5, y: 4.25, w: 9, h: 0.3,
    fontSize: 13, fontFace: "Calibri", color: GOLD,
    italic: true, bold: true, align: "center", margin: 0,
  });

  // emphasis word
  s.addText("NEW.", {
    x: 0.5, y: 4.7, w: 9, h: 0.55,
    fontSize: 14, fontFace: "Calibri", color: MUTED,
    align: "center", margin: 0,
    italic: true,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 12 — NOT Improved. NEW.
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: "080A12" }, line: { color: "080A12" } });

  const items = [
    { label: "Not improved.", active: false },
    { label: "Not upgraded.", active: false },
    { label: "Not enhanced.", active: false },
    { label: "NEW.", active: true },
  ];

  s.addText("You are…", {
    x: 0.5, y: 0.25, w: 9, h: 0.55,
    fontSize: 22, fontFace: "Georgia", color: MUTED, italic: true,
    align: "center", margin: 0,
  });

  items.forEach((item, i) => {
    const yBase = 0.92 + i * 1.02;
    s.addShape(pres.shapes.RECTANGLE, {
      x: item.active ? 1.5 : 2.5,
      y: yBase,
      w: item.active ? 7 : 5,
      h: 0.82,
      fill: { color: item.active ? GOLD : BG_CARD },
      line: { color: item.active ? GOLD : "2A2E42" },
      shadow: item.active ? makeShadow() : undefined,
    });
    s.addText(item.label, {
      x: item.active ? 1.5 : 2.5,
      y: yBase + 0.04,
      w: item.active ? 7 : 5,
      h: 0.74,
      fontSize: item.active ? 40 : 28,
      fontFace: "Georgia",
      color: item.active ? "0D0F1A" : MUTED,
      bold: item.active,
      align: "center",
      valign: "middle",
      margin: 0,
    });
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 13 — ILLUSTRATION: Cat and Dog
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.08, h: 5.625, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("ILLUSTRATION", {
    x: 0.3, y: 0.22, w: 4, h: 0.3,
    fontSize: 10, fontFace: "Calibri", color: GOLD, bold: true, charSpacing: 4, margin: 0,
  });
  s.addText("The Cat & the Dog", {
    x: 0.3, y: 0.57, w: 9.3, h: 0.75,
    fontSize: 38, fontFace: "Georgia", color: WHITE, bold: true, margin: 0, shadow: makeShadow(),
  });

  s.addText("Can a cat become a dog by acting like a dog?", {
    x: 0.3, y: 1.42, w: 9.3, h: 0.42,
    fontSize: 17, fontFace: "Calibri", color: CREAM, italic: true, margin: 0,
  });

  const catLines = [
    "You can teach a cat to fetch.",
    "Put a collar on it. Call it a good boy.",
    "Train it to sit on command.",
    "But at midnight — it’s still licking itself",
    "and knocking things off the counter.",
    "Because training never changed the nature.",
  ];
  let cy = 1.95;
  catLines.forEach((line, i) => {
    const isLast = i === catLines.length - 1;
    s.addText(line, {
      x: 0.3, y: cy, w: 9.3, h: 0.32,
      fontSize: isLast ? 16 : 14,
      fontFace: isLast ? "Georgia" : "Calibri",
      color: isLast ? GOLD_LITE : MUTED,
      italic: !isLast, bold: isLast, margin: 0,
    });
    cy += isLast ? 0.38 : 0.3;
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.85, w: 9.3, h: 0.04, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText([
    { text: "Behavior modification cannot change ", options: { color: CREAM } },
    { text: "what you are.", options: { color: GOLD, bold: true } },
    { text: "\nOnly Jesus can do that.", options: { color: WHITE, bold: true } },
  ], {
    x: 0.3, y: 4.0, w: 9.3, h: 1.1,
    fontSize: 18, fontFace: "Georgia", italic: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 14 — POINT 4: Two Men — Romans 5:19 revisited
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("POINT 4  —  Romans 5:19", {
    x: 0.5, y: 0.22, w: 9, h: 0.3,
    fontSize: 11, fontFace: "Calibri", color: GOLD,
    bold: true, charSpacing: 4, align: "center", margin: 0,
  });
  s.addText("His Obedience\nBecame Your Identity", {
    x: 0.5, y: 0.62, w: 9, h: 1.55,
    fontSize: 40, fontFace: "Georgia", color: WHITE, bold: true,
    align: "center", margin: 0, shadow: makeShadow(),
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 2.28, w: 8, h: 0.04, fill: { color: GOLD }, line: { color: GOLD } });

  // Two man comparison
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.45, w: 4.0, h: 2.65, fill: { color: "1A0B0B" }, line: { color: RED_ACC } });
  s.addText("ADAM", {
    x: 0.6, y: 2.55, w: 3.8, h: 0.45,
    fontSize: 20, fontFace: "Georgia", color: RED_ACC, bold: true, align: "center", margin: 0,
  });
  s.addText("One man’s\nDISO­BEDIENCE", {
    x: 0.6, y: 3.05, w: 3.8, h: 0.8,
    fontSize: 20, fontFace: "Georgia", color: WHITE, bold: true, align: "center", margin: 0,
  });
  s.addText("made many sinners", {
    x: 0.6, y: 3.88, w: 3.8, h: 0.5,
    fontSize: 14, fontFace: "Calibri", color: MUTED, italic: true, align: "center", margin: 0,
  });
  s.addText("by birth", {
    x: 0.6, y: 4.38, w: 3.8, h: 0.45,
    fontSize: 13, fontFace: "Calibri", color: RED_ACC, bold: true, align: "center", margin: 0,
  });

  s.addText("VS", {
    x: 4.5, y: 3.3, w: 1.0, h: 0.7,
    fontSize: 22, fontFace: "Georgia", color: MUTED, bold: true, align: "center", valign: "middle", margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: 2.45, w: 4.0, h: 2.65, fill: { color: "0D1A0D" }, line: { color: "3A7A3A" } });
  s.addText("JESUS", {
    x: 5.6, y: 2.55, w: 3.8, h: 0.45,
    fontSize: 20, fontFace: "Georgia", color: GOLD, bold: true, align: "center", margin: 0,
  });
  s.addText("One Man’s\nOBEDIENCE", {
    x: 5.6, y: 3.05, w: 3.8, h: 0.8,
    fontSize: 20, fontFace: "Georgia", color: WHITE, bold: true, align: "center", margin: 0,
  });
  s.addText("made many righteous", {
    x: 5.6, y: 3.88, w: 3.8, h: 0.5,
    fontSize: 14, fontFace: "Calibri", color: CREAM, italic: true, align: "center", margin: 0,
  });
  s.addText("by new birth", {
    x: 5.6, y: 4.38, w: 3.8, h: 0.45,
    fontSize: 13, fontFace: "Calibri", color: GOLD, bold: true, align: "center", margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 15 — APPLICATION: 3 Shifts
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("APPLICATION", {
    x: 0.5, y: 0.22, w: 9, h: 0.3,
    fontSize: 11, fontFace: "Calibri", color: GOLD,
    bold: true, charSpacing: 5, align: "center", margin: 0,
  });
  s.addText("Three Shifts Starting Today", {
    x: 0.5, y: 0.6, w: 9, h: 0.85,
    fontSize: 38, fontFace: "Georgia", color: WHITE, bold: true,
    align: "center", margin: 0, shadow: makeShadow(),
  });

  const shifts = [
    { num: "1", head: "Stop out-behaving your nature.", body: "No effort changes your spiritual DNA. Only the blood of Jesus does that." },
    { num: "2", head: "Reframe how you see your sin.", body: "My flesh acted like flesh. But my identity is in Christ. His record is my record." },
    { num: "3", head: "Speak truth out loud.", body: "I am a new creature. Old things have passed away. I am righteous because of Jesus." },
  ];

  shifts.forEach((shift, i) => {
    const yBase = 1.6 + i * 1.24;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yBase, w: 9, h: 1.1, fill: { color: BG_CARD }, line: { color: "2A2E42" } });
    // number circle
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: yBase + 0.18, w: 0.75, h: 0.75, fill: { color: GOLD }, line: { color: GOLD } });
    s.addText(shift.num, {
      x: 0.6, y: yBase + 0.18, w: 0.75, h: 0.75,
      fontSize: 22, fontFace: "Georgia", color: "0D0F1A",
      bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText(shift.head, {
      x: 1.5, y: yBase + 0.1, w: 7.8, h: 0.42,
      fontSize: 16, fontFace: "Georgia", color: GOLD_LITE,
      bold: true, margin: 0,
    });
    s.addText(shift.body, {
      x: 1.5, y: yBase + 0.52, w: 7.8, h: 0.46,
      fontSize: 13, fontFace: "Calibri", color: MUTED,
      italic: true, margin: 0,
    });
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 16 — DECLARATION SLIDE (Shift 3 spoken)
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);

  s.addText("SAY THIS OUT LOUD:", {
    x: 0.5, y: 0.4, w: 9, h: 0.35,
    fontSize: 12, fontFace: "Calibri", color: GOLD,
    bold: true, charSpacing: 5, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 0.9, w: 8.4, h: 3.8, fill: { color: BG_CARD }, line: { color: GOLD } });
  // gold left accent bar inside box
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 0.9, w: 0.08, h: 3.8, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText([
    { text: "“I am a ", options: { color: CREAM } },
    { text: "new creature in Christ.", options: { color: GOLD, bold: true } },
    { text: "\n\nOld things have passed away.", options: { color: CREAM } },
    { text: "\n\nI don’t fight to become righteous.", options: { color: CREAM } },
    { text: "\n\nI already am — because of ", options: { color: CREAM } },
    { text: "Jesus.”", options: { color: GOLD, bold: true } },
  ], {
    x: 1.1, y: 1.05, w: 7.9, h: 3.45,
    fontSize: 22, fontFace: "Georgia",
    italic: true, align: "center", valign: "middle", margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 4.85, w: 3.0, h: 0.04, fill: { color: GOLD }, line: { color: GOLD } });
  s.addText("Based on 2 Corinthians 5:17  &  Romans 5:19", {
    x: 0.5, y: 4.98, w: 9, h: 0.28,
    fontSize: 11, fontFace: "Calibri", color: MUTED, italic: true, align: "center", margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 17 — CLOSING: Birth Certificate
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: "080A12" }, line: { color: "080A12" } });

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: GOLD }, line: { color: GOLD } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.56, w: 10, h: 0.06, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("Jesus didn’t hand you", {
    x: 0.5, y: 0.5, w: 9, h: 0.7,
    fontSize: 28, fontFace: "Georgia", color: MUTED, italic: true,
    align: "center", margin: 0,
  });
  s.addText("a checklist.", {
    x: 0.5, y: 1.15, w: 9, h: 0.75,
    fontSize: 44, fontFace: "Georgia", color: WHITE, bold: true,
    align: "center", margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 3.0, y: 2.05, w: 4.0, h: 0.05, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("He handed you", {
    x: 0.5, y: 2.25, w: 9, h: 0.65,
    fontSize: 28, fontFace: "Georgia", color: MUTED, italic: true,
    align: "center", margin: 0,
  });
  s.addText("a new birth certificate.", {
    x: 0.5, y: 2.85, w: 9, h: 0.9,
    fontSize: 48, fontFace: "Georgia", color: GOLD, bold: true,
    align: "center", margin: 0, shadow: makeShadow(),
  });

  s.addText("You were born a sinner by one man’s failure.", {
    x: 0.5, y: 4.0, w: 9, h: 0.4,
    fontSize: 15, fontFace: "Calibri", color: MUTED, italic: true,
    align: "center", margin: 0,
  });
  s.addText("You were reborn righteous by one Man’s faithfulness.", {
    x: 0.5, y: 4.42, w: 9, h: 0.4,
    fontSize: 15, fontFace: "Georgia", color: CREAM, italic: true, bold: true,
    align: "center", margin: 0,
  });
  s.addText("That is grace.", {
    x: 0.5, y: 4.92, w: 9, h: 0.35,
    fontSize: 16, fontFace: "Georgia", color: GOLD, bold: true,
    align: "center", margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 18 — PRAYER
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText("CLOSING PRAYER", {
    x: 0.5, y: 0.22, w: 9, h: 0.35,
    fontSize: 12, fontFace: "Calibri", color: GOLD,
    bold: true, charSpacing: 6, align: "center", margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 0.7, w: 8.4, h: 4.55, fill: { color: BG_CARD }, line: { color: "2A2E42" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 0.7, w: 0.06, h: 4.55, fill: { color: GOLD }, line: { color: GOLD } });

  s.addText([
    { text: "Father God,\n\n", options: { bold: true, color: GOLD } },
    { text: "I confess I’ve been fighting symptoms my whole life.\nTrying to be good enough. Trying to clean myself up.\nTrying to earn what You already gave me for free.\n\n", options: { color: CREAM } },
    { text: "Forgive me for trusting my effort more than\nYour Son’s sacrifice.\n\n", options: { color: CREAM } },
    { text: "Remind me today — I am not what I do.\nI am who You say I am.\n\n", options: { color: GOLD_LITE } },
    { text: "A new creature. Made righteous by the obedience of Jesus.\nI rest in that today. In Jesus’ name — Amen.", options: { color: WHITE, bold: true } },
  ], {
    x: 1.1, y: 0.85, w: 8.0, h: 4.25,
    fontSize: 14.5, fontFace: "Georgia",
    italic: true, valign: "top", margin: 6,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// WRITE FILE
// ══════════════════════════════════════════════════════════════════════════════
pres.writeFile({ fileName: "/home/user/skills/the-root-problem.pptx" })
  .then(() => console.log("Done: the-root-problem.pptx"))
  .catch(err => { console.error(err); process.exit(1); });
