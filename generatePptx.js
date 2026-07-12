// H2X Studio — Fee Proposal Generator
// Layout: LAYOUT_WIDE 13.333" x 7.5" (extracted from Taseco VI PPTX)
// Fonts: Toma Sans Light (body) + Toma Sans Bold (emphasis) + Playfair Display (headings)

const pptxgen = require('pptxgenjs');
const FONTS = require('./fontData');
const { H2X_LOGO } = require('./logoData');
const os = require('os');
const path = require('path');

// ─── BRAND ───────────────────────────────────────────────────────
const C = {
  black:    '141414',
  white:    'FFFFFF',
  offwhite: 'F9F9F7',
  cream:    'F2F0ED',
  terra:    'C7A07F',   // terracotta gold
  terradk:  'B8895F',
  gray:     '999999',
  graylt:   'CCCCCC',
  rule:     'DEDBD5',
  rowalt:   'F2F0ED',
  rowtot:   'E8E4DF',
  border:   'E0DDD8',
  dark:     '1A1814',
};

// Slide 13.333" x 7.5" (LAYOUT_WIDE)
const W = 13.333, H = 7.5;
const ML = 0.6667; // margin left (exact from Taseco)
const CW = 12.0;   // content width

// Font system — matches Taseco exactly
const F = {
  serif:     'Playfair Display',
  serifI:    'Playfair Display',
  light:     'Toma Sans Light',
  bold:      'Toma Sans Bold',
  semi:      'Toma Sans SemiBold',
  regular:   'Toma Sans Regular',
};

// Size system — extracted from XML
const SZ = {
  h1:    23,  // slide title
  h1lg:  32,  // cover project name
  h2:    22,  // section subtitle
  label:  9,  // all body labels (spaced caps)
  body:   9,  // all body text
  cover: 12,  // cover location line
  coverSub: 16, // cover "Interior Design Services"
  coverFee: 14, // cover "Fee Proposal"
  coverVer: 10, // version / date
  pagenum: 10,  // page number
  copy:    6,   // copyright
  pullquote: 14,
  teamName: 13,
};

// ─── SHARED SCOPE / STAGE DEFINITIONS (A / B / C) ─────────────────
// Single source of truth for stage names + weeks + ratio, so Scope of Services,
// Design Fee, and Document Issue Schedule always stay in sync.
function getStageDefs(scope) {
  if (scope === 'B') {
    // Full service, no lighting — Lighting offered separately as an add-on
    return [
      { stage: 'Appraisal and Briefing', weeks: 1, ratio: 15 },
      { stage: 'Feasibility & Preliminary Concept', weeks: 2, ratio: 20 },
      { stage: 'Final Concept', weeks: 4, ratio: 20 },
      { stage: 'Schematic Design', weeks: 4, ratio: 15 },
      { stage: 'Design Development', weeks: 4, ratio: 10 },
      { stage: 'Construction Documents', weeks: 6, ratio: 10 },
      { stage: 'BOQ (Bill of Quantity)', weeks: 3, ratio: 5 },
      { stage: 'Author Supervision', weeks: 4, ratio: 5 },
    ];
  }
  if (scope === 'C') {
    // Full service, Lighting bundled (optional) into Final Concept stage
    return [
      { stage: 'Appraisal and Briefing', weeks: 1, ratio: 15 },
      { stage: 'Feasibility & Preliminary Concept', weeks: 2, ratio: 20 },
      { stage: 'Final Concept + Lighting (Optional)', weeks: 4, ratio: 20 },
      { stage: 'Schematic Design', weeks: 4, ratio: 15 },
      { stage: 'Design Development', weeks: 4, ratio: 10 },
      { stage: 'Construction Documents', weeks: 6, ratio: 10 },
      { stage: 'BOQ (Bill of Quantity)', weeks: 3, ratio: 5 },
      { stage: 'Author Supervision', weeks: 4, ratio: 5 },
    ];
  }
  // 'A' (default) — Concept → Schematic only
  return [
    { stage: 'Appraisal and Briefing', weeks: 1, ratio: 15 },
    { stage: 'Feasibility & Preliminary Concept', weeks: 2, ratio: 20 },
    { stage: 'Final Concept + Lighting (Optional)', weeks: 4, ratio: 30 },
    { stage: 'Schematic Design', weeks: 4, ratio: 35 },
  ];
}

// Optional (not included in the base fee) items shown per scope on Scope of Services
function getOptionalItems(scope) {
  if (scope === 'A') {
    return ['Design Development', 'Construction Documents', 'BOQ (Bill of Quantity)', 'Author Supervision'];
  }
  return []; // B and C already include the full scope through Author Supervision
}


// Short description per canonical stage name — reused across Design Process & Workflow and other slides
const STAGE_DESC = {
  'Appraisal and Briefing': 'Kick-off, brief alignment, site assessment — the foundation for everything that follows.',
  'Feasibility & Preliminary Concept': 'Feasibility & operational assessment, adjusted floor plan, mood & feeling direction.',
  'Final Concept + Lighting (Optional)': 'Concept statement, final layout, moodboard, 3D key views — lighting concept optional.',
  'Final Concept': 'Concept statement, final layout, moodboard, 3D key views, material palette.',
  'Schematic Design': 'Annotated plans, RCP & floor-finish plans, key elevations, outline specification.',
  'Design Development': 'Detailed drawings, material schedules, FF&E specifications.',
  'Construction Documents': 'Full CD set, coordinated with consultants, ready for tender.',
  'BOQ (Bill of Quantity)': 'Tender-ready bill of quantities.',
  'Author Supervision': 'Bid review, RFI responses, site visits through to completion.',
};

function makePres() {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';

  // Embed all fonts
  try {
    pres.defineFont({ name: 'Playfair Display',       data: FONTS['PlayfairDisplay-Regular'],  type: 'otf' });
    pres.defineFont({ name: 'Playfair Display Italic', data: FONTS['PlayfairDisplay-Italic'],   type: 'ttf' });
    pres.defineFont({ name: 'Playfair Display SemiBold', data: FONTS['PlayfairDisplay-SemiBold'], type: 'ttf' });
    pres.defineFont({ name: 'Toma Sans Light',   data: FONTS['TomaSans-Light'],   type: 'otf' });
    pres.defineFont({ name: 'Toma Sans Regular', data: FONTS['TomaSans-Regular'], type: 'otf' });
    pres.defineFont({ name: 'Toma Sans Bold',    data: FONTS['TomaSans-Bold'],    type: 'otf' });
    pres.defineFont({ name: 'Toma Sans SemiBold', data: FONTS['TomaSans-SemiBold'], type: 'otf' });
  } catch(e) {
    // defineFont not available in this version — fonts will use system fallback
    console.log('Note: defineFont not available, using system fonts');
  }

  return pres;
}

// Defensive truncation for fields that must fit a fixed-size box (e.g. one-line bullets)
function truncateWords(text, maxWords = 22) {
  if (!text) return text;
  const words = String(text).trim().split(/\s+/);
  if (words.length <= maxWords) return text;
  return words.slice(0, maxWords).join(' ') + '…';
}

// Logo — actual H2X Studio logo image (client-provided)
// pos: 'left' (used on cover) or 'right' (used on all content slides)
// Falls back to text-drawn version for a hypothetical dark/light-on-dark background,
// since we only have a black-on-transparent version of the logo file.
function logo(s, light = false, pos = 'left') {
  const w = 0.75, h = 0.434; // aspect ratio ≈ 1.73:1 (107x62px source)
  const x = pos === 'right' ? (W - ML - w) : ML;
  if (light) {
    // No white/light version of the logo asset available — fall back to text mark.
    s.addText('H2X', { x, y: 0.289, w: 0.7, h: 0.32,
      fontSize: 18, fontFace: F.serifI, color: C.white, bold: false, margin: 0 });
    s.addText('STUDIO', { x, y: 0.59, w: 1.0, h: 0.2,
      fontSize: 7, fontFace: F.light, color: C.terra, charSpacing: 3, margin: 0 });
    return;
  }
  s.addImage({ data: H2X_LOGO, x, y: 0.31, w, h });
}

// Studio name top-right
function studioName(s, light = false) {
  s.addText('H2X Studio', {
    x: 8.6663, y: 0.2889, w: 4.0, h: 0.5,
    fontSize: SZ.h2, fontFace: F.serif, color: light ? C.white : C.black,
    align: 'right', margin: 0,
  });
}

// Page number — Playfair Display, terracotta, right-aligned
function pageNum(s, num) {
  s.addText(String(num).padStart(2, '0'), {
    x: 11.6663, y: 7.08, w: 1.0, h: 0.3,
    fontSize: SZ.pagenum, fontFace: F.serif, color: C.terra,
    align: 'right', margin: 0,
  });
}

// Copyright — bottom left
function copyright(s) {
  s.addText('©H2X.Studio 2026. All Rights Reserved', {
    x: ML, y: 7.08, w: 9.0, h: 0.3,
    fontSize: SZ.copy, fontFace: F.light, color: C.graylt, margin: 0,
  });
}

// Slide title + subtitle line
function slideTitle(s, title, subtitle = '') {
  s.background = { color: C.white };
  logo(s, false, 'right');

  s.addText(title, {
    x: ML, y: 0.5, w: CW, h: 0.6,
    fontSize: SZ.h1, fontFace: F.serif, color: C.black, margin: 0,
  });
  // Rule line — exact from Taseco
  s.addShape('rect', {
    x: ML, y: 1.34, w: CW, h: 0.008,
    fill: { color: C.rule }, line: { color: C.rule, width: 0 },
  });
  if (subtitle) {
    s.addText(subtitle, {
      x: ML, y: 1.0842, w: CW, h: 0.2308,
      fontSize: SZ.label, fontFace: F.semi, color: C.terra, margin: 0,
    });
  }
}

// Section label (spaced caps, terracotta)
function secLabel(s, text, x, y, w) {
  s.addText(text, {
    x, y, w, h: 0.3,
    fontSize: SZ.label, fontFace: F.light, color: C.terra,
    charSpacing: 1.5, margin: 0,
  });
}

// Body text
function bodyText(s, text, x, y, w, h, bold = false, color = null) {
  s.addText(text, {
    x, y, w, h,
    fontSize: SZ.body,
    fontFace: bold ? F.bold : F.light,
    color: color || C.black,
    margin: 0, valign: 'top', lineSpacingMultiple: 1.5,
  });
}

// Table header row — black bg, white Toma Sans Light spaced caps
function tblHeader(s, cells, y, rowH = 0.35, region = null) {
  const rx = region ? region.x : ML;
  const rw = region ? region.w : CW;
  s.addShape('rect', {
    x: rx, y, w: rw, h: rowH,
    fill: { color: C.black }, line: { color: C.black, width: 0 },
  });
  cells.forEach(cell => {
    s.addText(cell.text, {
      x: cell.x, y: y + 0.05, w: cell.w, h: rowH - 0.08,
      fontSize: cell.fontSize || SZ.label, fontFace: F.light, color: C.white,
      charSpacing: cell.fontSize ? 0.5 : 1.5, margin: 0, valign: 'middle',
      align: cell.align || 'left',
    });
  });
}

// Table data row
function tblRow(s, cells, y, rowH, bg) {
  s.addShape('rect', {
    x: ML, y, w: CW, h: rowH,
    fill: { color: bg }, line: { color: C.border, width: 0.5 },
  });
  cells.forEach(cell => {
    s.addText(cell.text || '', {
      x: cell.x, y: y + 0.06, w: cell.w, h: rowH - 0.1,
      fontSize: cell.fontSize || SZ.body,
      fontFace: cell.bold ? F.bold : F.light,
      color: cell.color || C.black,
      bold: !!cell.bold, italic: !!cell.italic,
      align: cell.align || 'left',
      margin: 0, valign: 'middle',
    });
  });
}

// ─── SLIDE 01 — COVER ────────────────────────────────────────────
function slide01(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  logo(s);
  studioName(s);

  // Top rule (x=1.5" in Taseco)
  s.addShape('rect', {
    x: 1.5, y: 0.3889, w: 7.1663, h: 0.008,
    fill: { color: C.rule }, line: { color: C.rule, width: 0 },
  });

  // Hero image box
  if (d.heroImage) {
    s.addImage({ data: d.heroImage, x: 2.0, y: 0.6, w: 9.333, h: 4.3, sizing: { type: 'cover', w: 9.333, h: 4.3 } });
  } else {
    s.addShape('rect', {
      x: 2.0, y: 0.6, w: 9.333, h: 4.3,
      fill: { color: C.cream }, line: { color: C.graylt, width: 0.5 },
    });
    s.addText('[ HERO IMAGE ]', {
      x: 2.0, y: 0.6, w: 9.333, h: 4.3,
      fontSize: 9, fontFace: F.light, color: C.gray,
      align: 'center', valign: 'middle', margin: 0,
    });
  }

  // Project name
  s.addText(d.projectName || 'Tên Dự Án', {
    x: ML, y: 4.98, w: CW, h: 0.5,
    fontSize: SZ.h1lg, fontFace: F.serif, color: C.black,
    align: 'center', margin: 0,
  });

  // Location
  s.addText((d.location || 'ĐỊA ĐIỂM').toUpperCase(), {
    x: ML, y: 5.54, w: CW, h: 0.3,
    fontSize: SZ.cover, fontFace: F.light, color: C.terra,
    align: 'center', charSpacing: 3, margin: 0,
  });

  // Service line
  s.addText('INTERIOR DESIGN SERVICES', {
    x: ML, y: 5.96, w: CW, h: 0.3,
    fontSize: SZ.coverSub, fontFace: F.serif, color: C.black,
    align: 'center', margin: 0,
  });

  // Fee Proposal
  s.addText('Fee Proposal', {
    x: ML, y: 6.3, w: CW, h: 0.3,
    fontSize: SZ.coverFee, fontFace: F.serif, color: C.black,
    align: 'center', bold: true, margin: 0,
  });

  // Version / Date
  s.addText(`PHIÊN BẢN  ${d.version || '01'}`, {
    x: ML, y: 5.2, w: 4.0, h: 0.3,
    fontSize: SZ.coverVer, fontFace: F.light, color: C.gray,
    charSpacing: 2, margin: 0,
  });
  s.addText(d.month || 'JUNE 2026', {
    x: 8.6663, y: 5.2, w: 4.0, h: 0.3,
    fontSize: SZ.coverVer, fontFace: F.light, color: C.gray,
    align: 'right', charSpacing: 2, margin: 0,
  });

  copyright(s);
}

// ─── SLIDE 02 — PROJECT UNDERSTANDING ───────────────────────────
function slide02(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Project Understanding',
    `${d.clientName || '[Client]'} · ${d.typology || '[Typology]'} · ${d.location || '[Location]'}`);

  // Left column x=0.6667, w=5.7198
  // Right column x=7.1444, w=5.7198
  const lx = ML, lw = 5.7198;
  const rx = 7.1444, rw = 5.7198;

  // THE OPPORTUNITY
  secLabel(s, 'CƠ HỘI', lx, 1.6, lw);
  s.addText(d.opportunity || '[Mô tả cơ hội dự án]', {
    x: lx, y: 1.9398, w: lw + 0.43, h: 2.8363,
    fontSize: SZ.body, fontFace: F.light, color: C.black,
    margin: 0, valign: 'top', lineSpacingMultiple: 1.5,
  });

  // PROJECT BRIEF
  secLabel(s, 'TÓM TẮT DỰ ÁN', lx, 4.6957, lw);
  const briefLines = [
    { text: '+ ', options: { fontFace: F.light, color: C.terra } },
    { text: `Chủ đầu tư · ${d.clientName || '[Client]'}`, options: { fontFace: F.light, color: C.black } },
    { text: '\n+ ', options: { fontFace: F.light, color: C.terra } },
    { text: `Vị trí · ${d.location || '[Location]'}`, options: { fontFace: F.light, color: C.black } },
    { text: '\n+ ', options: { fontFace: F.light, color: C.terra } },
    { text: `Loại hình · ${d.typology || '[Typology]'}`, options: { fontFace: F.light, color: C.black } },
    { text: '\n+ ', options: { fontFace: F.light, color: C.terra } },
    { text: `Diện tích · ~${d.area || '[X]'} m²`, options: { fontFace: F.light, color: C.black } },
    { text: '\n+ ', options: { fontFace: F.light, color: C.terra } },
    { text: `Khách hàng mục tiêu · ${d.targetGuest || '[Description]'}`, options: { fontFace: F.light, color: C.black } },
  ].map(l => ({ text: l.text, options: { ...l.options, fontSize: SZ.body } }));
  s.addText(briefLines, { x: lx, y: 4.983, w: 5.9439, h: 1.8536, margin: 0, valign: 'top', lineSpacingMultiple: 1.5 });

  // EXPERIENCE INTENT
  secLabel(s, 'ĐỊNH HƯỚNG TRẢI NGHIỆM', rx, 1.6933, rw);
  s.addText(d.experienceIntent || '[Định hướng trải nghiệm]', {
    x: rx, y: 1.9182, w: rw, h: 2.7377,
    fontSize: SZ.body, fontFace: F.light, color: C.black,
    margin: 0, valign: 'top', lineSpacingMultiple: 1.5,
  });

  // WHY H2X
  secLabel(s, 'VÌ SAO H2X', rx, 4.7074, rw);
  s.addText(d.whyH2x || 'H2X (Human + Hospitality × eXperiences) thiết kế dựa trên vận hành và hành trình trải nghiệm của khách hàng.', {
    x: rx, y: 4.983, w: 5.2931, h: 1.3215,
    fontSize: SZ.body, fontFace: F.light, color: C.black,
    margin: 0, valign: 'top', lineSpacingMultiple: 1.5,
  });

  pageNum(s, 2);
  copyright(s);
}

// ─── SLIDE 03 — SCOPE OF SERVICES ───────────────────────────────
function slide03(pres, d) {
  const s = pres.addSlide();
  const scope = d.scope || 'A';
  const subtitles = {
    A: 'This engagement covers Concept & Schematic Design — Lighting offered as an optional add-on within Concept',
    B: 'Full-service engagement — Concept through Construction Documents & Site Supervision · Lighting offered separately',
    C: 'Full-service engagement — Concept through Construction Documents & Site Supervision · Lighting bundled as optional add-on',
  };
  slideTitle(s, 'Scope of Services', subtitles[scope] || subtitles.A);

  const stagesA = [
    { stage: 'Appraisal and Briefing', status: '✓ Included', desc: 'Kick-off, brief alignment, A&B material review, site assessment.' },
    { stage: 'Feasibility & Preliminary Concept', status: '✓ Included', desc: 'Feasibility & operational assessment, adjusted concept floor plan, mood & feeling recommendations.' },
    { stage: 'Final Concept + Lighting (Optional)', status: '✓ Included', desc: 'Concept statement, final layout, moodboard, 3D key views, material palette. Lighting concept offered as optional add-on.' },
    { stage: 'Schematic Design', status: '✓ Included', desc: 'Annotated plans, RCP & floor-finish plans, key elevations, outline specification schedules.' },
    { stage: 'Design Development · Construction Documents · BOQ', status: 'Optional', desc: 'Available as a follow-on package — fees proposed by addendum at the agreed rate once Schematic is approved.' },
    { stage: 'Author Supervision', status: 'Optional', desc: 'Offered separately on request once the project moves into construction.' },
  ];
  const stagesB = [
    { stage: 'Appraisal and Briefing', status: '✓ Included', desc: 'Kick-off, brief alignment, A&B material review, site assessment.' },
    { stage: 'Feasibility & Preliminary Concept', status: '✓ Included', desc: 'Feasibility & operational assessment, adjusted concept floor plan, mood & feeling recommendations.' },
    { stage: 'Final Concept', status: '✓ Included', desc: 'Concept statement, final layout, moodboard, 3D key views, material palette.' },
    { stage: 'Schematic Design', status: '✓ Included', desc: 'Annotated plans, RCP & floor-finish plans, key elevations, outline spec schedules.' },
    { stage: 'Design Development', status: '✓ Included', desc: 'Detailed drawings, material schedules, FF&E specifications.' },
    { stage: 'Construction Documents', status: '✓ Included', desc: 'Full CD set, coordinated with consultants.' },
    { stage: 'BOQ (Bill of Quantity)', status: '✓ Included', desc: 'Tender-ready bill of quantities.' },
    { stage: 'Author Supervision', status: '✓ Included', desc: 'Bid review, RFI responses, site visits through to completion.' },
  ];
  const stagesC = [
    { stage: 'Appraisal and Briefing', status: '✓ Included', desc: 'Kick-off, brief alignment, A&B material review, site assessment.' },
    { stage: 'Feasibility & Preliminary Concept', status: '✓ Included', desc: 'Feasibility & operational assessment, adjusted concept floor plan, mood & feeling recommendations.' },
    { stage: 'Final Concept + Lighting (Optional)', status: '✓ Included', desc: 'Concept statement, final layout, moodboard, 3D key views, material palette. Lighting concept offered as optional add-on.' },
    { stage: 'Schematic Design', status: '✓ Included', desc: 'Annotated plans, RCP & floor-finish plans, key elevations, outline spec schedules.' },
    { stage: 'Design Development', status: '✓ Included', desc: 'Detailed drawings, material schedules, FF&E specifications.' },
    { stage: 'Construction Documents', status: '✓ Included', desc: 'Full CD set, coordinated with consultants.' },
    { stage: 'BOQ (Bill of Quantity)', status: '✓ Included', desc: 'Tender-ready bill of quantities.' },
    { stage: 'Author Supervision', status: '✓ Included', desc: 'Bid review, RFI responses, site visits through to completion.' },
  ];
  const stages = scope === 'B' ? stagesB : scope === 'C' ? stagesC : stagesA;

  // Native table (avoids a rendering bug seen with many manually-drawn
  // overlapping shapes when row count is high — see Hourly Rates slide fix)
  const rowH = stages.length > 8 ? 0.44 : stages.length > 6 ? 0.52 : 0.72;
  const descFont = stages.length > 8 ? 7.5 : stages.length > 6 ? 8 : SZ.body;
  const scopeRows = [
    [
      { text: 'STAGE IN SCOPE', options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8 } },
      { text: '', options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8 } },
      { text: 'WHAT IT DELIVERS', options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8 } },
    ],
    ...stages.map((st, i) => {
      const isOpt = st.status === 'Optional';
      const bg = i % 2 === 0 ? C.rowalt : C.white;
      const fg = isOpt ? C.gray : C.black;
      return [
        { text: st.stage, options: { fill: { color: bg }, color: fg, fontFace: F.light, bold: !isOpt, italic: isOpt, fontSize: descFont, valign: 'top' } },
        { text: isOpt ? 'Optional' : '\u2713 Included', options: { fill: { color: bg }, color: isOpt ? C.gray : C.terradk, fontFace: F.light, bold: !isOpt, fontSize: 7.5, align: 'center', valign: 'top' } },
        { text: st.desc, options: { fill: { color: bg }, color: fg, fontFace: F.light, fontSize: descFont, valign: 'top' } },
      ];
    }),
  ];
  s.addTable(scopeRows, {
    x: ML, y: 1.55, w: CW,
    colW: [3.5, 1.4, 7.1],
    rowH: [0.35, ...stages.map(() => rowH)],
    border: { type: 'solid', color: C.border, pt: 0.5 },
    margin: [0.04, 0.08, 0.04, 0.08],
    autoPage: false,
  });

  if (scope === 'B') {
    const noteY = 1.55 + 0.35 + stages.length * rowH + 0.12;
    s.addText('Lighting Design — offered separately as an add-on at any stage, not bundled into the fee above.', {
      x: ML, y: noteY, w: CW, h: 0.25,
      fontSize: 8, fontFace: F.light, color: C.gray, italic: true, margin: 0,
    });
  }

  pageNum(s, 3);
  copyright(s);
}

// ─── SLIDE 04 — ZONING PROGRAMME ────────────────────────────────
function slide04(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Indicative Zoning Programme',
    `H2X đề xuất ~${d.area || '[X]'} m² · ~${d.peakGuests || '[n]'} khách peak · ~${d.sqmPerGuest || '7'} m²/khách · to be confirmed in Concept`);

  const zones = d.zones || [];

  // Header — from Taseco slide 4
  tblHeader(s, [
    { text: 'KHU VỰC', x: ML + 0.06, w: 3.7 },
    { text: 'DIỆN TÍCH', x: ML + 3.86, w: 1.2 },
    { text: 'CHỖ', x: ML + 5.06, w: 0.9 },
    { text: 'LÝ DO VẬN HÀNH', x: ML + 6.06, w: 6.0 },
  ], 1.3883, 0.35);

  const rowH = 0.56;
  zones.forEach((z, i) => {
    const isTotal = i === zones.length - 1 && z.zone && z.zone.toLowerCase().includes('total');
    const y = 1.3883 + 0.35 + i * rowH;
    const bg = isTotal ? C.rowtot : i % 2 === 0 ? C.rowalt : C.white;
    tblRow(s, [
      { text: z.zone || '', x: ML + 0.1, w: 3.6, bold: isTotal, color: isTotal ? C.terradk : C.black },
      { text: z.area ? `${z.area} m²` : '—', x: ML + 3.86, w: 1.1, align: 'center', bold: isTotal, color: isTotal ? C.terradk : C.black },
      { text: z.seats || '—', x: ML + 5.06, w: 0.85, align: 'center' },
      { text: z.rationale || '', x: ML + 6.06, w: 6.0, color: isTotal ? C.gray : C.black, italic: isTotal },
    ], y, rowH, bg);
  });

  // Total row if not already included
  if (zones.length > 0 && !zones[zones.length-1].zone?.toLowerCase().includes('total')) {
    const totalArea = zones.reduce((sum, z) => sum + (parseFloat(z.area)||0), 0);
    const y = 1.3883 + 0.35 + zones.length * rowH;
    tblRow(s, [
      { text: 'Total (net usable)', x: ML + 0.1, w: 3.6, bold: true, color: C.terradk },
      { text: `~${d.area || totalArea} m²`, x: ML + 3.86, w: 1.1, align: 'center', bold: true, color: C.terradk },
      { text: '', x: ML + 5.06, w: 0.85 },
      { text: 'Circulation (~15%) distributed within each zone.', x: ML + 6.06, w: 6.0, color: C.gray, italic: true },
    ], y, rowH, C.rowtot);
  }

  pageNum(s, 4);
  copyright(s);
}

// ─── SLIDE 05 — DESIGN LANGUAGE ─────────────────────────────────
function slide05(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  logo(s, false, 'right');

  s.addText('Design Language', {
    x: ML, y: 0.5, w: CW, h: 0.55,
    fontSize: SZ.h1, fontFace: F.serif, color: C.black, margin: 0,
  });
  s.addText(d.designTagline || 'Quiet luxury with a Vietnamese soul — warm materials, light from shadow, restraint over display', {
    x: ML, y: 1.0842, w: CW, h: 0.2308,
    fontSize: SZ.label, fontFace: F.semi, color: C.terra, margin: 0,
  });
  s.addShape('rect', {
    x: ML, y: 1.34, w: CW, h: 0.008,
    fill: { color: C.rule }, line: { color: C.rule, width: 0 },
  });

  // 3 image boxes
  const pillars = d.designPillars || [
    { label: 'WELCOME', sub: 'ARRIVAL RITUAL · THE HUMAN TOUCH' },
    { label: 'F&B', sub: 'REFINED · SENSORY · UNHURRIED' },
    { label: 'SOCIAL', sub: 'WARM · ATMOSPHERIC · CONNECTED' },
  ];
  const imgW = 3.84, imgH = 4.8, imgY = 1.45;
  const moodImages = d.moodImages || [];
  pillars.forEach((p, i) => {
    const x = ML + i * (imgW + 0.18);
    const img = moodImages[i];
    if (img) {
      s.addImage({ data: img, x, y: imgY, w: imgW, h: imgH, sizing: { type: 'cover', w: imgW, h: imgH } });
    } else {
      s.addShape('rect', {
        x, y: imgY, w: imgW, h: imgH,
        fill: { color: C.cream }, line: { color: C.graylt, width: 0.5 },
      });
      s.addText('[ MOOD IMAGE ]', {
        x, y: imgY, w: imgW, h: imgH,
        fontSize: 8, fontFace: F.light, color: C.gray,
        align: 'center', valign: 'middle', margin: 0,
      });
    }
    s.addText(p.label || '', {
      x, y: imgY + imgH + 0.08, w: imgW, h: 0.25,
      fontSize: SZ.label, fontFace: F.bold, color: C.black,
      charSpacing: 2, margin: 0, align: 'center',
    });
    s.addText(p.sub || '', {
      x, y: imgY + imgH + 0.35, w: imgW, h: 0.22,
      fontSize: 8, fontFace: F.bold, color: C.gray,
      charSpacing: 1, margin: 0, align: 'center',
    });
  });

  pageNum(s, 5);
  copyright(s);
}

// ─── SLIDE 06 — LIGHTING DESIGN ──────────────────────────────────
function slide06(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Lighting Design — Built In, From Concept',
    'Included in the design fee · H2X\'s in-house strength · the single biggest lever on atmosphere');

  // Pull quote
  s.addText('"Light is not a final layer — it is the atmosphere. So we design it from the first Concept sketch, not after the walls are built."', {
    x: ML, y: 1.45, w: CW, h: 0.6,
    fontSize: SZ.pullquote, fontFace: F.serifI, color: C.black,
    italic: true, margin: 0,
  });

  // Two columns
  const cols = [
    {
      x: ML, w: 5.72, label: 'WHY INTEGRATE FROM CONCEPT',
      items: [
        'Light shapes mood, brand identity & guest experience from day one — splitting it out invites vision conflict and costly late redesign',
        'The lighting concept evolves in parallel with space, materials & FF&E → photorealistic renders early, fewer iteration loops',
        'No interface risk between an interior designer and an external lighting consultant',
        'Optimises the lighting budget — a meaningful share of FF&E — instead of value-engineering it away at the end',
      ],
    },
    {
      x: 7.28, w: 5.72, label: 'WHAT H2X DELIVERS (IN THE FEE)',
      items: [
        d.lightingStrategy ? truncateWords(d.lightingStrategy, 22) : 'Zoned atmosphere strategy — dramatic for F&B, functional for business, circadian/calming for Quiet & Wellness',
        'Daylight integration with the view; warm indirect & cove lighting',
        'Fixture & control concept; scene-setting by time of day & guest wave',
        'Vietnam energy-code compliant; premium, buildable fixtures via our long-standing supplier network',
      ],
    },
  ];

  cols.forEach(col => {
    secLabel(s, col.label, col.x, 2.2, col.w);
    col.items.forEach((item, i) => {
      s.addText([
        { text: '+ ', options: { fontFace: F.light, color: C.terra, fontSize: SZ.body } },
        { text: item, options: { fontFace: F.light, color: C.black, fontSize: SZ.body } },
      ], { x: col.x, y: 2.6 + i * 1.05, w: col.w, h: 0.9, margin: 0, valign: 'top', lineSpacingMultiple: 1.5 });
    });
  });

  pageNum(s, 6);
  copyright(s);
}

// ─── SLIDE 07 — PROJECT TEAM ──────────────────────────────────────
function slide07(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Project Team', 'A senior, multidisciplinary team — with in-house Lighting Design');

  const team = d.team || [
    { role: 'CHIEF CREATIVE OFFICER', name: 'Gia Huy (Michael)', title: 'M.Architect · 24+ yrs · sets vision & design direction', dark: true },
    { role: 'CHIEF OPERATING OFFICER', name: 'Nguyễn Văn Toàn (Karo)', title: 'Architect · 20+ yrs · sets vision & design direction', dark: true },
    { role: 'CONCEPT LEAD', name: 'Nguyễn Mạnh Hùng (Henry)', title: 'Spatial concept, narrative & guest journey' },
    { role: 'HEAD OF FF&E', name: 'Nguyễn Thị Thúy Vy (Vivian)', title: 'Furniture, fixtures & material curation' },
    { role: 'TECHNICAL LEAD', name: 'Nguyễn Vương Linh (Lucas)', title: 'Schematic documentation & consultant coordination' },
    { role: '3D VISUALISATION & PM', name: 'Đỗ Danh Sơn', title: 'Photorealistic renders · schedule, reviews & delivery' },
    { role: 'LIGHTING DESIGNER', name: '[Lighting Lead]', title: 'Atmosphere, fixture & control strategy' },
    { role: 'AUDIO VISUAL DESIGNER', name: '[AV Lead]', title: 'Soundscape, AV & display concept' },
    { role: 'BRAND LEAD', name: '[Brand Lead]', title: 'Naming, identity & brand guidelines' },
  ];

  const cw = 3.84, ch = 1.62;
  team.forEach((m, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = ML + col * (cw + 0.18);
    const y = 1.5 + row * (ch + 0.15);

    s.addShape('rect', {
      x, y, w: cw, h: ch,
      fill: { color: m.dark ? C.dark : C.rowalt },
      line: { color: C.border, width: 0.5 },
    });
    s.addText(m.role, {
      x: x + 0.18, y: y + 0.15, w: cw - 0.3, h: 0.25,
      fontSize: 7.5, fontFace: F.light, color: C.terra,
      charSpacing: 1.5, margin: 0,
    });
    s.addText(m.name, {
      x: x + 0.18, y: y + 0.44, w: cw - 0.3, h: 0.4,
      fontSize: SZ.teamName, fontFace: F.serif,
      color: m.dark ? C.white : C.black, margin: 0,
    });
    s.addText(m.title, {
      x: x + 0.18, y: y + 0.88, w: cw - 0.3, h: 0.65,
      fontSize: SZ.body, fontFace: F.light,
      color: m.dark ? C.graylt : C.gray,
      margin: 0, valign: 'top', lineSpacingMultiple: 1.5,
    });
  });

  pageNum(s, 7);
  copyright(s);
}

// ─── SLIDE 08 — DESIGN PROCESS ───────────────────────────────────
function slide08(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Design Process & Workflow', 'Approval-gated at every stage \u00b7 First Draft \u2192 Review \u2192 Final Approval \u2014 auto-adjusts to K\u1ecbch b\u1ea3n A / B / C');

  const scope = d.scope || 'A';
  const defs = getStageDefs(scope);
  const SHORT_LABEL = {
    'Appraisal and Briefing': 'Appraisal &\nBriefing',
    'Feasibility & Preliminary Concept': 'Feasibility &\nPrelim. Concept',
    'Final Concept + Lighting (Optional)': 'Final Concept\n+ Lighting*',
    'Final Concept': 'Final\nConcept',
    'Schematic Design': 'Schematic\nDesign',
    'Design Development': 'Design\nDevelopment',
    'Construction Documents': 'Construction\nDocuments',
    'BOQ (Bill of Quantity)': 'BOQ',
    'Author Supervision': 'Author\nSupervision',
  };
  const stages = defs.map(st => SHORT_LABEL[st.stage] || st.stage);

  // Flow diagram — box count/labels always match the selected Scope A/B/C
  const count = stages.length;
  const gap = count > 5 ? 0.12 : 0.2;
  const bw = (CW - (count - 1) * gap) / count;
  const boxFont = count > 6 ? 7.5 : count > 4 ? 8.5 : 10;
  stages.forEach((st, i) => {
    const x = ML + i * (bw + gap);
    const dark = i === stages.length - 1;
    s.addShape('rect', {
      x, y: 1.55, w: bw, h: 0.85,
      fill: { color: dark ? C.black : C.white },
      line: { color: dark ? C.black : C.graylt, width: 1 },
    });
    s.addText(st, {
      x: x + 0.04, y: 1.6, w: bw - 0.08, h: 0.7,
      fontSize: boxFont, fontFace: F.serif, color: dark ? C.white : C.black,
      align: 'center', valign: 'middle', margin: 0, lineSpacingMultiple: 1.05,
    });
    if (i < stages.length - 1) {
      s.addText('\u2192', {
        x: x + bw + 0.01, y: 1.78, w: gap - 0.02, h: 0.3,
        fontSize: count > 5 ? 9 : 12, fontFace: F.light, color: C.gray, align: 'center', margin: 0,
      });
    }
  });

  // OUR METHODOLOGY
  secLabel(s, 'OUR METHODOLOGY', ML, 2.65, CW);
  s.addText('We provide a step-by-step method for how we approach a project to achieve the best results at each stage. Every stage runs through a First Draft, one consolidated Review round, then Final Approval before moving forward \u2014 giving the Client full visibility and control at each gate.', {
    x: ML, y: 2.95, w: CW, h: 0.65,
    fontSize: SZ.body, fontFace: F.light, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.4,
  });

  // OUR PROCESS — one card per stage, always matching the selected Scope
  secLabel(s, 'OUR PROCESS', ML, 3.75, CW);
  const cardY = 4.05;
  const cGap = 0.14;
  const cw = (CW - (count - 1) * cGap) / count;
  const cardFont = count > 6 ? 6.5 : count > 4 ? 7 : 7.5;
  defs.forEach((st, i) => {
    const x = ML + i * (cw + cGap);
    s.addShape('rect', { x, y: cardY, w: cw, h: 2.15, fill: { color: C.rowalt }, line: { color: C.border, width: 0.5 } });
    s.addText(String(i + 1), { x: x + 0.08, y: cardY + 0.08, w: cw - 0.16, h: 0.24, fontSize: 8.5, fontFace: F.bold, color: C.terradk, margin: 0 });
    s.addText(st.stage, {
      x: x + 0.08, y: cardY + 0.32, w: cw - 0.16, h: 0.55,
      fontSize: cardFont + 1, fontFace: F.bold, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.15,
    });
    s.addText(STAGE_DESC[st.stage] || '', {
      x: x + 0.08, y: cardY + 0.9, w: cw - 0.16, h: 1.15,
      fontSize: cardFont, fontFace: F.light, color: C.gray, margin: 0, valign: 'top', lineSpacingMultiple: 1.25,
    });
  });

  s.addText(`Methodology: ${defs.map(st => st.stage.split(' + ')[0].split(' (')[0]).join(' \u2192 ')}${scope === 'A' ? ' \u2014 full service available as a follow-on addendum.' : '.'}${scope !== 'B' ? ' *Lighting concept offered as an optional add-on.' : ''}`, {
    x: ML, y: H - 0.6, w: CW, h: 0.22,
    fontSize: 7.5, fontFace: F.light, color: C.gray, margin: 0,
  });

  pageNum(s, 8);
  copyright(s);
}

function slideWorkStageDeliverable(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Work Stage Deliverable', 'How each stage moves from draft to signed approval');

  // Flow diagram
  const flow = ['Kick-off\nMeeting', 'First\nDraft', 'First\nReview', 'Final\nApproval', 'Signed Stage\nApproved'];
  const fw = 2.2, fgap = 0.28;
  flow.forEach((f, i) => {
    const x = ML + i * (fw + fgap);
    const dark = i === flow.length - 1;
    s.addShape('rect', {
      x, y: 1.55, w: fw, h: 0.85,
      fill: { color: dark ? C.black : C.white },
      line: { color: dark ? C.black : C.graylt, width: 1 },
    });
    s.addText(f, {
      x: x + 0.08, y: 1.6, w: fw - 0.16, h: 0.72, fontSize: 10, fontFace: F.serif,
      color: dark ? C.white : C.black, align: 'center', valign: 'middle', margin: 0,
    });
    if (i < flow.length - 1) {
      s.addText('→', { x: x + fw + 0.01, y: 1.78, w: fgap - 0.02, h: 0.35, fontSize: 12, fontFace: F.light, color: C.gray, align: 'center', margin: 0 });
    }
  });

  // Explanation blocks
  const blocks = [
    { h: 'First Draft', b: 'The Designer prepares a First Draft based on information received and issues it to the Project Manager.' },
    { h: 'First Review', b: 'The PM issues the First Draft to all project stakeholders and forwards their consolidated comments and instructions to the Designer.' },
    { h: 'Final Approval', b: 'After the final presentation, the Designer collates the received comments and instructions. A stage-completion letter is then issued, capturing any minor items to carry into the next stage.' },
  ];
  blocks.forEach((bl, i) => {
    const x = ML + i * (3.85 + 0.28);
    s.addShape('rect', { x, y: 2.75, w: 3.85, h: 2.15, fill: { color: C.rowalt }, line: { color: C.border, width: 0.5 } });
    s.addText(bl.h, { x: x + 0.18, y: 2.88, w: 3.55, h: 0.32, fontSize: 11, fontFace: F.serif, color: C.black, margin: 0 });
    s.addText(bl.b, {
      x: x + 0.18, y: 3.24, w: 3.55, h: 1.55, fontSize: SZ.body, fontFace: F.light, color: C.black,
      margin: 0, valign: 'top', lineSpacingMultiple: 1.4,
    });
  });

  s.addText('Additional rounds: any round requested beyond those stated above will be subject to an extension of time and additional fees as per the Agreement.', {
    x: ML, y: 5.2, w: CW, h: 0.4, fontSize: 8, fontFace: F.light, color: C.terradk, italic: true, margin: 0, lineSpacingMultiple: 1.3,
  });
  s.addText('An Appraisal & Briefing stage is added ahead of Feasibility to align on project approach and planning. Stage durations may be adjusted based on H2X\u2019s evaluation of scope, deliverables, and required Client review/approval time at each stage.', {
    x: ML, y: 5.65, w: CW, h: 0.55, fontSize: 8, fontFace: F.light, color: C.gray, margin: 0, lineSpacingMultiple: 1.3,
  });

  pageNum(s, 9);
  copyright(s);
}

// ─── SLIDE 10 — PROJECT ORGANISATION CHART ───────────────────────
function slideOrgChart(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Project Organisation Chart', 'A senior, multidisciplinary team structured around your project');

  const team = d.team || [
    { role: 'CHIEF CREATIVE OFFICER', name: 'Gia Huy (Michael)' },
    { role: 'CHIEF OPERATING OFFICER', name: 'Nguy\u1ec5n V\u0103n To\u00e0n (Karo)' },
    { role: 'CONCEPT LEAD', name: 'Nguy\u1ec5n M\u1ea1nh H\u00f9ng (Henry)' },
    { role: 'HEAD OF FF&E', name: 'Nguy\u1ec5n Th\u1ecb Th\u00fay Vy (Vivian)' },
    { role: 'TECHNICAL LEAD', name: 'Nguy\u1ec5n V\u01b0\u01a1ng Linh (Lucas)' },
    { role: '3D VISUALISATION & PM', name: '\u0110\u1ed7 Danh S\u01a1n' },
    { role: 'LIGHTING DESIGNER', name: '[Lighting Lead]' },
    { role: 'AUDIO VISUAL DESIGNER', name: '[AV Lead]' },
    { role: 'BRAND LEAD', name: '[Brand Lead]' },
  ];
  const leaders = team.slice(0, 2);
  const heads = team.slice(2);

  const topW = 3.2, topX = ML + (CW - topW) / 2;
  s.addShape('rect', { x: topX, y: 1.5, w: topW, h: 0.5, fill: { color: C.black }, line: { color: C.black, width: 0 } });
  s.addText('H2X STUDIO', { x: topX, y: 1.5, w: topW, h: 0.5, fontSize: 11, fontFace: F.serif, color: C.white, align: 'center', valign: 'middle', margin: 0, charSpacing: 1 });

  const lw = 3.6, lgap = 0.3;
  const lTotalW = leaders.length * lw + (leaders.length - 1) * lgap;
  const lStartX = ML + (CW - lTotalW) / 2;
  const t2Y = 2.25;
  s.addShape('line', { x: ML + CW / 2, y: 2.0, w: 0, h: 0.25, line: { color: C.graylt, width: 1 } });
  leaders.forEach((m, i) => {
    const x = lStartX + i * (lw + lgap);
    s.addShape('rect', { x, y: t2Y, w: lw, h: 0.58, fill: { color: C.terradk }, line: { color: C.terradk, width: 0 } });
    s.addText(m.role, { x, y: t2Y + 0.05, w: lw, h: 0.22, fontSize: 7, fontFace: F.light, color: C.white, align: 'center', margin: 0, charSpacing: 1 });
    s.addText(m.name, { x, y: t2Y + 0.26, w: lw, h: 0.3, fontSize: 10, fontFace: F.serif, color: C.white, align: 'center', margin: 0 });
  });

  const n = heads.length;
  const hgapFixed = 0.14;
  const hw = (CW - (n - 1) * hgapFixed) / n;
  const t3Y = 3.35;
  s.addShape('line', { x: ML + CW / 2, y: t2Y + 0.58, w: 0, h: t3Y - (t2Y + 0.58), line: { color: C.graylt, width: 1 } });
  heads.forEach((m, i) => {
    const x = ML + i * (hw + hgapFixed);
    s.addShape('line', { x: x + hw / 2, y: t3Y - 0.15, w: 0, h: 0.15, line: { color: C.graylt, width: 1 } });
    s.addShape('rect', { x, y: t3Y, w: hw, h: 0.95, fill: { color: C.rowalt }, line: { color: C.border, width: 0.5 } });
    s.addText(m.role, { x: x + 0.06, y: t3Y + 0.08, w: hw - 0.12, h: 0.42, fontSize: 6, fontFace: F.light, color: C.terra, align: 'center', margin: 0, valign: 'top', lineSpacingMultiple: 1.15 });
    s.addText(m.name, { x: x + 0.06, y: t3Y + 0.52, w: hw - 0.12, h: 0.38, fontSize: 7.5, fontFace: F.serif, color: C.black, align: 'center', margin: 0, valign: 'top' });
  });

  const t4Y = 4.75;
  const support = ['Architecture &\nConcept Team', 'Interior &\nFF&E Team', 'Lighting &\nTechnical Team', '3D Visualisation\nTeam', 'Brand & AV\nTeam'];
  const sw = (CW - (support.length - 1) * 0.14) / support.length;
  s.addText('Supported by', { x: ML, y: t4Y - 0.32, w: 3, h: 0.25, fontSize: 8, fontFace: F.light, color: C.gray, italic: true, margin: 0 });
  support.forEach((sup, i) => {
    const x = ML + i * (sw + 0.14);
    s.addShape('rect', { x, y: t4Y, w: sw, h: 0.6, fill: { color: C.white }, line: { color: C.graylt, width: 0.75 } });
    s.addText(sup, { x: x + 0.06, y: t4Y, w: sw - 0.12, h: 0.6, fontSize: 7, fontFace: F.light, color: C.black, align: 'center', valign: 'middle', margin: 0, lineSpacingMultiple: 1.15 });
  });

  s.addText('Every project is led directly by the CCO/COO and a dedicated Lighting Designer works in-house from Concept, not bolted on later.', {
    x: ML, y: 5.65, w: CW, h: 0.3, fontSize: 8, fontFace: F.light, color: C.gray, italic: true, margin: 0,
  });

  pageNum(s, 10);
  copyright(s);
}

// ─── SLIDE 11 — DESIGN WORK PHASE ────────────────────────────────
function slideDesignWorkPhase(pres, d) {
  const s = pres.addSlide();
  const scope = d.scope || 'A';
  slideTitle(s, 'Design Work Phase', `What H2X delivers at each stage \u2014 K\u1ecbch b\u1ea3n ${scope} \u2014 deliverable-level detail`);

  const WORK_PHASE_DELIVERABLE = {
    'Appraisal and Briefing': 'Kick-off meeting; scope alignment; site visit; review of existing materials (site plans, structural/M&E drawings).',
    'Feasibility & Preliminary Concept': 'Feasibility & operational assessment; preliminary zoning layout; material board; design narrative (DWG/PDF).',
    'Final Concept + Lighting (Optional)': 'Look & Feel / Moodboard & story, coloured plan, 3D colour renders, digital material board \u2014 all in PDF.',
    'Final Concept': 'Look & Feel / Moodboard & story, coloured plan, 3D colour renders, digital material board \u2014 all in PDF.',
    'Schematic Design': 'Annotated plans, RCP & floor-finish plan, key elevations, outline specification schedules (DWG/PDF).',
    'Design Development': 'Full technical drawings: fit-out, built-in, MEP coordination, loose furniture list, specification schedules (DWG/PDF).',
    'Construction Documents': 'Complete construction drawing set for site execution, coordinated across all consultants (DWG/PDF).',
    'BOQ (Bill of Quantity)': 'Tender-ready bill of quantities.',
    'Author Supervision': 'Bid review, RFI responses, and periodic site visits through to completion.',
  };
  const defs = getStageDefs(scope);
  const phases = defs.map(st => ({
    name: st.stage,
    lead: st.stage.includes('Lighting') ? 'H2X (Design + Lighting)' : 'H2X',
    deliverable: WORK_PHASE_DELIVERABLE[st.stage] || '',
  }));

  tblHeader(s, [
    { text: 'PHASE', x: ML + 0.06, w: 3.4 },
    { text: 'LEAD', x: ML + 3.55, w: 2.1 },
    { text: 'KEY DELIVERABLES', x: ML + 5.75, w: 6.25 },
  ], 1.55, 0.32);

  const rowH = phases.length > 6 ? 0.55 : phases.length > 4 ? 0.7 : 1.05;
  phases.forEach((p, i) => {
    const y = 1.55 + 0.32 + i * rowH;
    const bg = i % 2 === 0 ? C.rowalt : C.white;
    s.addShape('rect', { x: ML, y, w: CW, h: rowH - 0.06, fill: { color: bg }, line: { color: C.border, width: 0.5 } });
    s.addText(p.name, { x: ML + 0.16, y: y + 0.06, w: 3.25, h: rowH - 0.14, fontSize: SZ.body, fontFace: F.bold, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.2 });
    s.addText(p.lead, { x: ML + 3.55, y: y + 0.06, w: 2.0, h: rowH - 0.14, fontSize: 7.5, fontFace: F.light, color: C.terradk, margin: 0, valign: 'top' });
    s.addText(p.deliverable, { x: ML + 5.75, y: y + 0.06, w: 6.15, h: rowH - 0.14, fontSize: 7.5, fontFace: F.light, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.2 });
  });

  s.addText('Detailed purpose, scope-of-work and full deliverable breakdown for each phase follow on the next pages.', {
    x: ML, y: 1.55 + 0.32 + phases.length * rowH + 0.15, w: CW, h: 0.25,
    fontSize: 7.5, fontFace: F.light, color: C.gray, italic: true, margin: 0,
  });

  pageNum(s, 11);
  copyright(s);
}

// ─── PHASE EXPLAINER (generic renderer, used for the 6 detail slides below) ──
function slidePhaseExplainer(pres, d, pageN, data) {
  const s = pres.addSlide();
  slideTitle(s, data.title, data.leadLine);

  secLabel(s, 'PURPOSE OF THE STAGE', ML, 1.55, CW);
  s.addText(data.purpose, {
    x: ML, y: 1.83, w: CW, h: 0.55,
    fontSize: SZ.body, fontFace: F.light, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.35,
  });

  secLabel(s, 'COVERED WITHIN THE STAGE', ML, 2.5, 7.0);
  let cy = 2.78;
  data.covered.forEach(line => {
    s.addText('+', { x: ML, y: cy, w: 0.2, h: 0.3, fontSize: SZ.body, fontFace: F.bold, color: C.terra, margin: 0 });
    s.addText(line, {
      x: ML + 0.22, y: cy, w: 6.78, h: 0.5,
      fontSize: 8, fontFace: F.light, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.25,
    });
    cy += 0.02 + Math.ceil(line.length / 78) * 0.22;
  });

  secLabel(s, 'DELIVERABLES', ML + 7.4, 2.5, 4.9);
  let dy = 2.78;
  data.deliverables.forEach(line => {
    s.addText('+', { x: ML + 7.4, y: dy, w: 0.2, h: 0.3, fontSize: SZ.body, fontFace: F.bold, color: C.terra, margin: 0 });
    s.addText(line, {
      x: ML + 7.62, y: dy, w: 4.7, h: 0.5,
      fontSize: 8, fontFace: F.light, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.25,
    });
    dy += 0.02 + Math.ceil(line.length / 55) * 0.22;
  });

  pageNum(s, pageN);
  copyright(s);
}

function slideBriefing(pres, d) {
  slidePhaseExplainer(pres, d, 12, {
    title: 'Briefing and Mobilisation',
    leadLine: 'LEAD \u2014 H2X & Client \u00b7 SUPPORT \u2014 Consultants',
    purpose: 'Ensure the project kicks off with the full design and consultant teams on board, and briefs agreed.',
    covered: [
      'H2X is introduced to the rest of the Design Team; overall scope and master-planning are reviewed to identify any scope gaps.',
      'H2X completes Appraisal & Briefing with all relevant material to kick off the Concept phase: site plans, elevations, existing structural and M&E drawings (DWG files), site photos, renderings, landscaping.',
      'Site visit (where possible) with the Client and Lead Consultant to review location, surroundings, site restrictions/limitations.',
      'Review of the brief with the Client to check for updates since contract agreement, guidelines, program requirements and scope of works.',
    ],
    deliverables: ['Kick-off meeting'],
  });
}

function slideFeasibility(pres, d) {
  slidePhaseExplainer(pres, d, 13, {
    title: 'Feasibility (Pre-Concept Design)',
    leadLine: 'LEAD \u2014 H2X \u00b7 SUPPORT \u2014 Consultants',
    purpose: 'Generate a design narrative.',
    covered: [
      'Research the unique attributes of the location and site, including a deeper understanding of the brief and brand aspirations.',
      'Prepare a selection of images to determine the design concept direction by means of photographic references.',
      'Prepare preliminary interior zoning layout showing the functional programme.',
      'Prepare material board (photographic references) and design narrative.',
      '1st presentation with Client for interior concept direction.',
    ],
    deliverables: ['Overview Concept Page (DWG/PDF)'],
  });
}

function slideConceptDesignExplainer(pres, d) {
  slidePhaseExplainer(pres, d, 14, {
    title: 'Concept Design',
    leadLine: 'LEAD \u2014 H2X \u00b7 SUPPORT \u2014 Consultants \u00b7 Runs concurrently with Feasibility',
    purpose: 'Establish an agreed Concept intent and direction, reflecting the character of the project areas covered in the brief.',
    covered: [
      'Narrative pages: creating a unique point to base the design on.',
      'Mood & Feel presentation: inspiration images and initial sketches explaining overall look and feel.',
      'Furniture and Fixtures: initial selection of furniture, lighting, fixtures and fittings for feedback, incl. initial custom sketches.',
      'Material boards: initial presentation of architectural and upholstery finishes in key areas.',
    ],
    deliverables: ['Look & Feel, Moodboard and story (PDF)', 'Coloured Plan', '3D Colour Renders', 'Digital material board (PDF)'],
  });
}

function slideSchematicExplainer(pres, d) {
  slidePhaseExplainer(pres, d, 15, {
    title: 'Schematic Design',
    leadLine: 'LEAD \u2014 H2X \u00b7 SUPPORT \u2014 Consultants \u00b7 Kicks off once Concept Design is approved',
    purpose: 'Provide DWG drawings and basic schedules to the design team for development into the Detail Design package.',
    covered: [
      'Plans clearly annotated with key finishes, including RCP and floor-finishes plan.',
      'Key elevations clearly annotated with key finishes, agreed at the end of Concept Design.',
      'Outline specification schedules for architectural finishes, presented as index pages for further development.',
    ],
    deliverables: ['Plans package (DWG/PDF)', 'Elevations (DWG/PDF)', 'Table of Materials (Office files/PDF)', 'General layout, floor finish, electrical & RCP plans', 'Key sections showing furniture elevations'],
  });
}

function slideDevelopmentExplainer(pres, d) {
  slidePhaseExplainer(pres, d, 16, {
    title: 'Development Design',
    leadLine: 'LEAD \u2014 H2X \u00b7 SUPPORT \u2014 Consultants \u00b7 Develops from the approved Schematic Design',
    purpose: 'Develop the entire technical drawing set to ensure the design follows the direction set in Schematic Design.',
    covered: [
      'Full development of technical drawings for the entire interior, including walls and separate items.',
      'Complete technical solutions for the entire project, ensuring accurate estimation, bidding invitation and design feasibility.',
      'Coordination of MEP design across the full technical document set.',
    ],
    deliverables: ['Technical docs: Fit-out, Built-in (DWG/PDF)', 'Technical docs: full MEP design (DWG/PDF)', 'Specifications (Digital list)', 'Loose Furniture list', 'RCP, wall/floor finish & electrical plans', 'Large-scale detail drawings'],
  });
}

function slideConstructionDocsExplainer(pres, d) {
  slidePhaseExplainer(pres, d, 17, {
    title: 'Construction Documents',
    leadLine: 'LEAD \u2014 H2X \u00b7 SUPPORT \u2014 Consultants \u00b7 Final phase, developed from approved Development Design',
    purpose: 'Produce a complete set of detailed technical construction drawings for site execution \u2014 the official documentation for construction on site.',
    covered: [
      'Develop the full set of technical construction drawings based on the approved design: fit-out and built-in drawings.',
      'Detailed drawings of walls, ceilings, flooring, loose furniture, and built components.',
      'Precise technical drawings, specifications and schedules ensuring accuracy, feasibility and coordination across all construction activities.',
    ],
    deliverables: ['Construction Drawings Package (DWG/PDF)', 'Schedules & technical specs for hard finishes, sanitary fittings, equipment', 'FF&E selection'],
  });
}

// ─── SLIDE 09 — DETAILED DESIGN STAGE
// ─── SLIDE 09 — DETAILED DESIGN STAGE ────────────────────────────
function slide09(pres, d) {
  const s = pres.addSlide();
  const scope = d.scope || 'A';
  slideTitle(s, 'Detailed Design Stage', `Design stage breakdown · Kịch bản ${scope} · weeks & ratio synced with Document Issue Schedule`);

  const defs = d.docSchedule && d.docSchedule.length ? d.docSchedule : getStageDefs(scope);
  const totalWeeks = defs.reduce((sum, st) => sum + (Number(st.weeks) || 0), 0);

  tblHeader(s, [
    { text: '#', x: ML + 0.06, w: 0.6, align: 'center' },
    { text: 'PHASE', x: ML + 0.75, w: 7.5 },
    { text: 'WEEKS', x: ML + 8.35, w: 1.2, align: 'center' },
    { text: 'RATIO', x: ML + 9.65, w: 2.35, align: 'center' },
  ], 1.55, 0.35);

  const rowH = defs.length > 5 ? 0.5 : 0.65;
  defs.forEach((st, i) => {
    const y = 1.55 + 0.35 + i * rowH;
    tblRow(s, [
      { text: String(i + 1), x: ML + 0.1, w: 0.55, align: 'center' },
      { text: st.stage, x: ML + 0.75, w: 7.45 },
      { text: String(st.weeks), x: ML + 8.35, w: 1.15, align: 'center' },
      { text: st.ratio + '%', x: ML + 9.65, w: 2.35, align: 'center' },
    ], y, rowH - 0.05, i % 2 === 0 ? C.white : C.rowalt);
  });

  const totY = 1.55 + 0.35 + defs.length * rowH;
  tblRow(s, [
    { text: 'TOTAL', x: ML + 0.75, w: 7.7, bold: true, color: C.terradk },
    { text: String(totalWeeks), x: ML + 8.35, w: 1.15, align: 'center', bold: true, color: C.terradk },
    { text: '100%', x: ML + 9.65, w: 2.35, align: 'center', bold: true, color: C.terradk },
  ], totY, 0.4, C.rowtot);

  s.addText('Client Approval required at the end of each stage before proceeding.' + (scope === 'A' ? ' Design Development / Construction Documents / BOQ / Author Supervision available as a follow-on addendum.' : ' Bid/Tender/Procurement and Handover priced separately by others.'), {
    x: ML, y: totY + 0.5, w: CW, h: 0.35,
    fontSize: 7.5, fontFace: F.light, color: C.gray, italic: true, margin: 0, lineSpacingMultiple: 1.3,
  });

  pageNum(s, 20);
  copyright(s);
}

// ─── SLIDE 10 — DETAILED SOW LIST ────────────────────────────────
function slide10(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Detailed SOW List', 'Scope of Works breakdown by deliverable, mapped to K\u1ecbch b\u1ea3n A / B / C');

  const scope = d.scope || 'A';
  const groups = d.sowGroups || [
    { num: '0', name: 'Appraisal and Briefing', opt: [1, 1, 1] },
    { num: '1', name: 'Feasibility & Preliminary Concept Design', items: [
      { num: '1.1', name: 'Feasibility and Operational Assessment', opt: [1, 1, 1] },
      { num: '1.2', name: 'Adjusted Concept Design Floor Plan', opt: [1, 1, 1] },
      { num: '1.3', name: 'Mood & Feeling Recommendations', opt: [1, 1, 1] },
    ]},
    { num: '2', name: 'Final Concept', items: [
      { num: '2.1', name: 'Concept Statement, Final Layout & Moodboard', opt: [1, 1, 1] },
      { num: '2.2', name: 'Picture & Video Render', opt: [1, 1, 1] },
    ]},
    { num: '3', name: 'Schematic Design', items: [
      { num: '3.1', name: 'Schematic Drawings', opt: [1, 1, 1] },
      { num: '3.2', name: 'Lighting Design Consulting', opt: [1, 0, 1], note: 'Optional / offered as add-on in B' },
      { num: '3.3', name: 'Signage & AV Consulting', opt: [0, 0, 0], note: 'By specialist consultant, add-on' },
    ]},
    { num: '4', name: 'Design Development', items: [
      { num: '4.1', name: 'Interior Design Development', opt: [0, 1, 1] },
      { num: '4.2', name: 'MEP Design Development', opt: [0, 1, 1] },
    ]},
    { num: '5', name: 'Construction Phase', items: [
      { num: '5.1', name: 'Construction Documents', opt: [0, 1, 1] },
      { num: '5.2', name: 'BOQ (Bill of Quantity)', opt: [0, 1, 1], note: 'For tender quantities' },
      { num: '5.3', name: 'Bid / Tender / Procurement', opt: [0, 0, 0], note: 'Client\u2019s responsibility' },
      { num: '5.4', name: 'VE and Budget Control', opt: [0, 0, 0] },
    ]},
    { num: '6', name: 'Author Supervision', items: [
      { num: '6.1', name: 'Bid Review & RFI Response', opt: [0, 1, 1] },
      { num: '6.2', name: 'Periodic Site Visits', opt: [0, 1, 1] },
    ]},
    { num: '7', name: 'Handover', opt: [0, 1, 1] },
  ];
  const optLabels = ['K\u1ecaCH B\u1ea2N A', 'K\u1ecaCH B\u1ea2N B', 'K\u1ecaCH B\u1ea2N C'];
  const activeCol = scope === 'A' ? 0 : scope === 'B' ? 1 : 2;
  const optW = 1.0, optX = [ML + 5.45, ML + 6.55, ML + 7.65];

  tblHeader(s, [
    { text: '#', x: ML + 0.06, w: 0.5, align: 'center' },
    { text: 'SCOPE OF WORKS', x: ML + 0.65, w: 4.7 },
    { text: optLabels[0], x: optX[0], w: optW, align: 'center', fontSize: 6.5 },
    { text: optLabels[1], x: optX[1], w: optW, align: 'center', fontSize: 6.5 },
    { text: optLabels[2], x: optX[2], w: optW, align: 'center', fontSize: 6.5 },
    { text: 'NOTE', x: ML + 8.75, w: 3.25 },
  ], 1.55, 0.3);

  const GH = 0.2, RH = 0.165;
  let y = 1.55 + 0.3;
  const drawCheckRow = (num, name, opt, note, y, h, bg) => {
    tblRow(s, [
      { text: num, x: ML + 0.1, w: 0.45, align: 'center', fontSize: 7 },
      { text: name, x: ML + 0.65, w: 4.65, fontSize: 7 },
      { text: '', x: ML + 8.75, w: 0 },
    ], y, h, bg);
    [0, 1, 2].forEach(c => {
      s.addShape('rect', { x: optX[c], y, w: optW, h, fill: { color: bg }, line: { color: C.border, width: 0.5 } });
      if (opt[c]) {
        s.addText('✓', { x: optX[c], y, w: optW, h, fontSize: 8, fontFace: F.bold,
          color: c === activeCol ? C.terradk : C.black, align: 'center', valign: 'middle', margin: 0 });
      }
    });
    if (note) {
      s.addText(note, { x: ML + 8.75, y, w: 3.2, h, fontSize: 6.5, fontFace: F.light, color: C.gray, italic: true, valign: 'middle', margin: 0 });
    }
  };

  groups.forEach((g, gi) => {
    const bgBase = gi % 2 === 0 ? C.white : C.rowalt;
    if (g.items) {
      s.addShape('rect', { x: ML, y, w: CW, h: GH, fill: { color: C.cream }, line: { color: C.border, width: 0.5 } });
      s.addText(`${g.num}   ${g.name}`, { x: ML + 0.1, y, w: 8.5, h: GH, fontSize: 8, fontFace: F.bold, color: C.terradk, margin: 0, valign: 'middle' });
      y += GH;
      g.items.forEach((it, i) => {
        drawCheckRow(it.num, it.name, it.opt, it.note, y, RH, i % 2 === 0 ? C.white : C.rowalt);
        y += RH;
      });
    } else {
      drawCheckRow(g.num, g.name, g.opt, g.note, y, RH, bgBase);
      y += RH;
    }
  });

  const noteY = y + 0.1;
  s.addText(`Option ${activeCol + 1} (highlighted ✓) is the scope proposed in this fee proposal.`, {
    x: ML, y: noteY, w: CW, h: 0.2, fontSize: 7.5, fontFace: F.light, color: C.terradk, italic: true, margin: 0,
  });

  pageNum(s, 21);
  copyright(s);
}

// ─── SLIDE 11 — DOCUMENT ISSUE SCHEDULE ──────────────────────────
function slide11(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Document Issue Schedule', 'Preliminary schedule · Interior Design + Lighting Design · lump sum by phase');

  const scope = d.scope || 'A';
  const defs = d.docSchedule && d.docSchedule.length ? d.docSchedule : getStageDefs(scope);
  const sched = defs.map((st, i) => ({ num: String(i + 1), stage: st.stage, weeks: Number(st.weeks) || 0, ratio: st.ratio }));
  const totalWeeks = sched.reduce((sum, st) => sum + st.weeks, 0);

  tblHeader(s, [
    { text: 'STAGE', x: ML + 0.06, w: 0.6 },
    { text: 'PHASE', x: ML + 0.75, w: 7.5 },
    { text: 'WEEKS', x: ML + 8.35, w: 1.2, align: 'center' },
    { text: 'RATIO', x: ML + 9.65, w: 2.35, align: 'center' },
  ], 1.55, 0.35);

  const rowH = sched.length > 4 ? 0.5 : 0.65;
  sched.forEach((st, i) => {
    const y = 1.55 + 0.35 + i * rowH;
    tblRow(s, [
      { text: st.num, x: ML + 0.1, w: 0.55, align: 'center' },
      { text: st.stage, x: ML + 0.75, w: 7.45 },
      { text: String(st.weeks), x: ML + 8.35, w: 1.15, align: 'center' },
      { text: st.ratio + '%', x: ML + 9.65, w: 2.35, align: 'center' },
    ], y, rowH - 0.05, i % 2 === 0 ? C.rowalt : C.white);
  });

  const totY = 1.55 + 0.35 + sched.length * rowH;
  tblRow(s, [
    { text: 'TOTAL', x: ML + 0.75, w: 7.7, bold: true, color: C.terradk },
    { text: String(totalWeeks), x: ML + 8.35, w: 1.15, align: 'center', bold: true, color: C.terradk },
    { text: '100%', x: ML + 9.65, w: 2.35, align: 'center', bold: true, color: C.terradk },
  ], totY, 0.4, C.rowtot);

  // Proportional timeline bar
  secLabel(s, 'TIMELINE OVERVIEW (BY WEEK)', ML, totY + 0.5, CW);
  const barY = totY + 0.8, barH = 0.45;
  const shades = [C.black, C.terradk, C.terra, C.gray, C.dark, C.rowtot];
  let bx = ML;
  sched.forEach((st, i) => {
    const segW = totalWeeks > 0 ? (st.weeks / totalWeeks) * CW : 0;
    s.addShape('rect', { x: bx, y: barY, w: segW, h: barH, fill: { color: shades[i % shades.length] }, line: { color: C.white, width: 1 } });
    s.addText(`${st.weeks}w`, {
      x: bx, y: barY, w: segW, h: barH, fontSize: 8, fontFace: F.bold,
      color: C.white, align: 'center', valign: 'middle', margin: 0,
    });
    bx += segW;
  });
  s.addText(`Estimated total duration: ~${totalWeeks} weeks (excludes Client approval turnaround time).`, {
    x: ML, y: barY + barH + 0.1, w: CW, h: 0.22, fontSize: 7.5, fontFace: F.light, color: C.gray, italic: true, margin: 0,
  });

  pageNum(s, 18);
  copyright(s);
}

// ─── SLIDE 12 — DESIGN FEE ────────────────────────────────────────
function slide12(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Design Fee', 'Interior Design + Lighting Design · lump sum by phase');

  const scope = d.scope || 'A';
  const fee = parseFloat(d.feeTotal) || 0;
  const fmt = n => n > 0 ? n.toLocaleString('en-US') : '[—]';

  const defs = getStageDefs(scope);
  const stages = defs.map((st, i) => ({
    num: String(i + 1), stage: st.stage, ratio: st.ratio + '%', amt: Math.round(fee * st.ratio / 100),
  }));

  tblHeader(s, [
    { text: 'STAGE', x: ML + 0.06, w: 0.6 },
    { text: 'PHASE', x: ML + 0.75, w: 7.5 },
    { text: 'RATIO', x: ML + 8.35, w: 1.2 },
    { text: 'FEE (USD)', x: ML + 9.65, w: 2.35, align: 'center' },
  ], 1.55, 0.35);

  const rowH = stages.length > 4 ? 0.5 : 0.7;
  stages.forEach((st, i) => {
    const y = 1.55 + 0.35 + i * rowH;
    tblRow(s, [
      { text: st.num, x: ML + 0.1, w: 0.55, align: 'center' },
      { text: st.stage, x: ML + 0.75, w: 7.45 },
      { text: st.ratio, x: ML + 8.35, w: 1.15, align: 'center' },
      { text: fmt(st.amt), x: ML + 9.65, w: 2.35, align: 'center' },
    ], y, rowH - 0.05, i % 2 === 0 ? C.rowalt : C.white);
  });

  // Subtotal
  const totY = 1.55 + 0.35 + stages.length * rowH;
  tblRow(s, [
    { text: 'INTERIOR + LIGHTING — SUBTOTAL', x: ML + 0.75, w: 8.8, bold: true, color: C.terradk },
    { text: '100%', x: ML + 8.35, w: 1.15, align: 'center', bold: true, color: C.terradk },
    { text: fmt(fee), x: ML + 9.65, w: 2.35, align: 'center', bold: true, color: C.terradk },
  ], totY, 0.45, C.rowtot);

  // Benchmark + notes (consolidated into one tight block to fit varying row counts)
  const area = parseFloat(d.area) || 0;
  const perSqm = area > 0 ? Math.round(fee / area) : 0;
  const scopeNote = scope === 'A'
    ? 'Design Development / Construction Documents / BOQ / Author Supervision available as a follow-on addendum.'
    : 'Lighting Design available as an optional add-on — see Add-on Workstreams.';
  const noteLines = [
    `≈ ${perSqm > 0 ? perSqm + ' USD / m²' : '[USD/m²]'} · ${d.area || '[X]'} m² · indicative benchmark for this typology.`,
    scopeNote,
    'Fees above exclude taxes and other charges.',
    'Using additional services within our ecosystem entitles you to a discount on the design fee.',
  ].join('\n');
  s.addText(noteLines, {
    x: ML, y: totY + 0.5, w: CW, h: 0.62,
    fontSize: 7.5, fontFace: F.light, color: C.gray, italic: true, margin: 0, lineSpacingMultiple: 1.25,
  });

  pageNum(s, 19);
  copyright(s);
}

// ─── SLIDE 13 — ADD-ON VALUE ─────────────────────────────────────
function slide13(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Branding · Signage · AV — Scope & Value', null);
  s.addText('Three workstreams that turn a beautiful space into a coherent, memorable destination — offered at concept stage', {
    x: ML, y: 1.0842, w: CW, h: 0.2308,
    fontSize: SZ.label, fontFace: F.semi, color: C.terra, margin: 0,
  });

  const addons = [
    { title: 'Brand Identity & Naming', partner: 'Ziva',
      scope: ['Lounge name & story', 'Logo & core visual identity', 'Brand guidelines (essentials)'],
      why: ['A named space earns loyalty, PR & partnership value', 'One coherent guest-facing story across every touchpoint'] },
    { title: 'Signage & Wayfinding', partner: 'Ziva',
      scope: ['Bilingual wayfinding strategy', 'Signage concept & key types', 'Specification direction'],
      why: ['Effortless navigation lowers anxiety & staff load', 'Signage is the brand at eye level — venue-grade legibility'] },
    { title: 'Audio-Visual Design', partner: 'H2X',
      scope: ['Signature soundscape direction', 'AV & display integration concept', 'Zoning of sound by area'],
      why: ['Sound completes the atmosphere lighting begins', 'AV as ambience, not noise — reliable, calm information'] },
  ];

  const cw = 3.84;
  addons.forEach((a, i) => {
    const x = ML + i * (cw + 0.18);
    s.addShape('rect', { x, y: 1.5, w: cw, h: 5.5, fill: { color: C.rowalt }, line: { color: C.border, width: 0.5 } });
    s.addText(a.title, { x: x + 0.18, y: 1.65, w: cw - 0.3, h: 0.45, fontSize: 13, fontFace: F.serif, color: C.black, margin: 0 });
    s.addText(`Partner: ${a.partner}`, { x: x + 0.18, y: 2.14, w: cw - 0.3, h: 0.22, fontSize: 8, fontFace: F.bold, color: C.terra, margin: 0 });
    secLabel(s, 'SCOPE (1 CONCEPT)', x + 0.18, 2.42, cw - 0.3);
    s.addText(a.scope.map(l => `+ ${l}`).join('\n'), { x: x + 0.18, y: 2.72, w: cw - 0.3, h: 0.95, fontSize: SZ.body, fontFace: F.light, color: C.black, margin: 0 });
    secLabel(s, 'WHY IT MATTERS', x + 0.18, 3.75, cw - 0.3);
    s.addText(a.why.map(l => `+ ${l}`).join('\n'), { x: x + 0.18, y: 4.05, w: cw - 0.3, h: 1.8, fontSize: SZ.body, fontFace: F.light, color: C.black, margin: 0 });
  });

  pageNum(s, 22);
  copyright(s);
}

// ─── SLIDE 14 — ADD-ON FEE ───────────────────────────────────────
function slide14(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Add-on Workstreams', 'Concept stage · itemised so each can be selected independently · design fees only (fabrication/production excluded)');

  const fmt = n => n > 0 ? n.toLocaleString('en-US') : '[—]';
  const feeMain = parseFloat(d.feeTotal) || 0;
  const feeBrand = d.addBrand ? parseFloat(d.feeBrand) || 0 : 0;
  const feeSign = d.addSignage ? parseFloat(d.feeSignage) || 0 : 0;
  const feeAV = d.addAV ? parseFloat(d.feeAV) || 0 : 0;
  const addons = feeBrand + feeSign + feeAV;
  const grand = feeMain + addons;

  const rows = [
    { ws: 'Brand Identity & Naming', partner: 'Ziva', scope: 'Naming, logo, core identity & essential guidelines.', fee: feeBrand, show: d.addBrand },
    { ws: 'Signage & Wayfinding', partner: 'Ziva', scope: 'Bilingual wayfinding strategy & signage concept.', fee: feeSign, show: d.addSignage },
    { ws: 'Audio-Visual Design', partner: 'H2X', scope: 'Soundscape, AV & display concept.', fee: feeAV, show: d.addAV },
  ];

  tblHeader(s, [
    { text: 'WORKSTREAM', x: ML + 0.06, w: 3.2 },
    { text: 'PARTNER', x: ML + 3.35, w: 1.1 },
    { text: 'SCOPE (CONCEPT, 1 DIRECTION)', x: ML + 4.55, w: 5.5 },
    { text: 'FEE (USD)', x: ML + 10.15, w: 1.85, align: 'center' },
  ], 1.55, 0.35);

  rows.forEach((r, i) => {
    const y = 1.55 + 0.35 + i * 0.65;
    tblRow(s, [
      { text: r.ws, x: ML + 0.1, w: 3.1 },
      { text: r.partner, x: ML + 3.35, w: 1.05, align: 'center' },
      { text: r.scope, x: ML + 4.55, w: 5.45 },
      { text: r.show ? fmt(r.fee) : 'Available on request', x: ML + 10.15, w: 1.85, align: 'center', italic: !r.show, color: r.show ? null : C.gray, fontSize: r.show ? SZ.body : 7.5 },
    ], y, 0.6, i % 2 === 0 ? C.rowalt : C.white);
  });

  const subY = 1.55 + 0.35 + rows.length * 0.65;
  tblRow(s, [
    { text: 'Add-ons subtotal', x: ML + 0.1, w: 9.9, bold: true, color: C.terradk },
    { text: fmt(addons), x: ML + 10.15, w: 1.85, align: 'center', bold: true, color: C.terradk },
  ], subY, 0.55, C.rowtot);

  // Grand total
  const gtY = subY + 0.65;
  s.addShape('rect', { x: ML, y: gtY, w: CW, h: 0.7, fill: { color: C.black }, line: { color: C.black, width: 0 } });
  s.addText('GRAND TOTAL', { x: ML + 0.15, y: gtY + 0.08, w: 6, h: 0.32, fontSize: 10, fontFace: F.light, color: C.white, charSpacing: 2, margin: 0 });
  s.addText(`Interior + Lighting (${fmt(feeMain)}) + Add-ons (${fmt(addons)}). Excludes taxes.`, {
    x: ML + 0.15, y: gtY + 0.42, w: 8, h: 0.22, fontSize: 7.5, fontFace: F.light, color: C.gray, margin: 0,
  });
  s.addText(`${fmt(grand)} USD`, {
    x: ML + 10.15, y: gtY, w: 1.85, h: 0.7, fontSize: 12, fontFace: F.bold, color: C.terra, align: 'center', valign: 'middle', margin: 0,
  });

  // Discount
  s.addShape('rect', { x: ML, y: gtY + 0.78, w: CW, h: 0.35, fill: { color: 'EDE8E2' }, line: { color: 'D6C8B8', width: 0.5 } });
  s.addText(`Early-commitment partnership discount −10% if signed within 14 days  →  ${fmt(Math.round(grand * 0.9))} USD`, {
    x: ML + 0.15, y: gtY + 0.82, w: CW - 0.3, h: 0.25, fontSize: SZ.body, fontFace: F.light, color: C.terradk, italic: true, margin: 0,
  });

  pageNum(s, 23);
  copyright(s);
}

// ─── SLIDE 15 — PAYMENT SCHEDULE ────────────────────────────────
function slide15(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Payment Schedule', 'Milestone-based, aligned to stage approvals');

  const scope = d.scope || 'A';
  const fee = parseFloat(d.feeTotal) || 0;
  const fmt = n => n > 0 ? n.toLocaleString('en-US') : '[—]';

  const defs = getStageDefs(scope);
  const pmts = defs.map((st, i) => ({
    num: String(i + 1),
    m: `${st.stage} approval`,
    pct: st.ratio + '%',
    amt: Math.round(fee * st.ratio / 100),
    note: i === 0 ? 'Upon signing' : (i === defs.length - 1 ? 'Final milestone' : ''),
  }));

  tblHeader(s, [
    { text: '#', x: ML + 0.06, w: 0.5, align: 'center' },
    { text: 'MILESTONE', x: ML + 0.65, w: 6.55 },
    { text: '%', x: ML + 7.3, w: 0.75, align: 'center' },
    { text: 'FEE DESIGN (USD)', x: ML + 8.15, w: 1.85, align: 'center' },
    { text: 'NOTE', x: ML + 10.1, w: 1.9 },
  ], 1.55, 0.35);

  const rowH = pmts.length > 6 ? 0.44 : pmts.length > 4 ? 0.52 : 0.72;
  pmts.forEach((p, i) => {
    const y = 1.55 + 0.35 + i * rowH;
    tblRow(s, [
      { text: p.num, x: ML + 0.1, w: 0.45, align: 'center' },
      { text: p.m, x: ML + 0.65, w: 6.5 },
      { text: p.pct, x: ML + 7.3, w: 0.7, align: 'center' },
      { text: fmt(p.amt), x: ML + 8.15, w: 1.8, align: 'center' },
      { text: p.note, x: ML + 10.1, w: 1.85, fontSize: 7.5, color: C.gray, italic: true },
    ], y, rowH - 0.05, i % 2 === 0 ? C.rowalt : C.white);
  });

  const totY = 1.55 + 0.35 + pmts.length * rowH;
  tblRow(s, [
    { text: 'Subtotal — Design Fee', x: ML + 0.65, w: 6.5, bold: true, color: C.terradk },
    { text: '100%', x: ML + 7.3, w: 0.7, align: 'center', bold: true, color: C.terradk },
    { text: fmt(fee), x: ML + 8.15, w: 1.8, align: 'center', bold: true, color: C.terradk },
    { text: '', x: ML + 10.1, w: 1.85 },
  ], totY, 0.38, C.rowtot);

  // Add-ons (if any selected) — compact single line + grand total, auto-summed
  const addonRows = [
    d.addBrand ? { label: 'Brand Identity & Naming', amt: parseFloat(d.feeBrand) || 0 } : null,
    d.addSignage ? { label: 'Signage & Wayfinding', amt: parseFloat(d.feeSignage) || 0 } : null,
    d.addAV ? { label: 'Audio-Visual Design', amt: parseFloat(d.feeAV) || 0 } : null,
  ].filter(Boolean);
  const addonsTotal = addonRows.reduce((sum, r) => sum + r.amt, 0);

  let y2 = totY + 0.38;
  if (addonRows.length) {
    tblRow(s, [
      { text: `Add-ons — ${addonRows.map(r => r.label).join(' · ')}`, x: ML + 0.65, w: 7.3, italic: true, color: C.gray, fontSize: 8 },
      { text: fmt(addonsTotal), x: ML + 8.15, w: 1.8, align: 'center', italic: true, color: C.gray },
      { text: 'Billed on delivery', x: ML + 10.1, w: 1.85, fontSize: 7.5, color: C.gray, italic: true },
    ], y2, 0.32, C.white);
    y2 += 0.32;
    tblRow(s, [
      { text: 'GRAND TOTAL (Design Fee + Add-ons)', x: ML + 0.65, w: 7.3, bold: true, color: C.white },
      { text: fmt(fee + addonsTotal), x: ML + 8.15, w: 1.8, align: 'center', bold: true, color: C.white },
      { text: '', x: ML + 10.1, w: 1.85 },
    ], y2, 0.38, C.black);
    y2 += 0.38;
  }

  s.addText('Add-on workstreams billed 50% on kick-off and 50% on delivery unless folded into the schedule above. Invoices in USD; payment within 14 days of invoice.', {
    x: ML, y: y2 + 0.08, w: CW, h: 0.3,
    fontSize: 7, fontFace: F.light, color: C.gray, italic: true, margin: 0, lineSpacingMultiple: 1.15,
  });

  pageNum(s, 24);
  copyright(s);
}

// ─── SLIDE 16 — INCLUSIONS & EXCLUSIONS ──────────────────────────
function slide16(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Inclusions & Exclusions', null);

  const scope = d.scope || 'A';
  const incA = [
    'Interior design — Appraisal through Schematic Design',
    'Lighting Design — concept for Final Concept stage (optional, integrated in the design fee)',
    'Feasibility assessment, zoning & space planning, design narrative, moodboards',
    '3D key-view renders · material & FF&E palette',
    'Schematic plans, RCP, elevations & outline specifications',
    'Brand, Signage & AV — concept stage (1 direction each) if selected',
  ];
  const incBC = [
    'Interior design — Appraisal through Author Supervision (full service)',
    scope === 'C' ? 'Lighting Design — concept for Final Concept stage (optional, integrated in the design fee)' : 'Lighting Design — offered separately as an add-on (not bundled in this fee)',
    'Feasibility assessment, zoning & space planning, design narrative, moodboards',
    '3D key-view renders · material & FF&E palette',
    'Full technical drawing set: Schematic, Design Development, Construction Documents',
    'BOQ (Bill of Quantity) — tender-ready',
    'Author Supervision — bid review, RFI responses, periodic site visits',
    'Brand, Signage & AV — concept stage (1 direction each) if selected',
  ];
  const inc = scope === 'A' ? incA : incBC;

  const excA = [
    'Design Development, Construction Documents, BOQ & Author Supervision (follow-on addendum)',
    'Development & production of branding, signage & AV beyond concept',
    'Base-building architecture; MEP, structural & façade engineering',
    'Kitchen consultant, acoustics & fire engineering',
    'Permits & approvals · taxes/VAT',
    'Reimbursables: travel, printing, mock-ups (at cost with prior approval)',
    'Construction & site works',
  ];
  const excBC = [
    scope === 'B' ? 'Lighting Design (offered separately as an add-on)' : null,
    'Development & production of branding, signage & AV beyond concept',
    'Base-building architecture; MEP, structural & façade engineering',
    'Kitchen consultant, acoustics & fire engineering',
    'Permits & approvals · taxes/VAT',
    'Reimbursables: travel, printing, mock-ups (at cost with prior approval)',
    'Construction & site works (physical build)',
    'Tender / bid management beyond bid review (procurement is Client\u2019s responsibility)',
  ].filter(Boolean);
  const exc = scope === 'A' ? excA : excBC;

  [
    { label: 'INCLUSIONS', items: inc, x: ML, bg: C.rowalt, color: C.black },
    { label: 'EXCLUSIONS', items: exc, x: ML + 6.18, bg: C.white, color: C.gray },
  ].forEach(col => {
    s.addShape('rect', { x: col.x, y: 1.55, w: 5.9, h: 5.6, fill: { color: col.bg }, line: { color: C.border, width: 0.5 } });
    secLabel(s, col.label, col.x + 0.2, 1.68, 5.6);
    s.addText(col.items.map(l => `+ ${l}`).join('\n'), {
      x: col.x + 0.2, y: 1.98, w: 5.65, h: 4.98,
      fontSize: SZ.body, fontFace: F.light, color: col.color,
      margin: 0, valign: 'top', lineSpacingMultiple: 1.5,
    });
  });

  pageNum(s, 25);
  copyright(s);
}

// ─── SLIDE 17 — TERMS & CONDITIONS ───────────────────────────────
function slide17(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Terms & Conditions', null);

  const terms = [
    { h: 'Revisions', b: 'Each stage includes a First Draft + one consolidated review round; further rounds by addendum (time + fee).' },
    { h: 'Site visits', b: 'Up to three (3) included in the design fee; additional visits charged at cost.' },
    { h: 'Reimbursables', b: 'Travel, printing & physical mock-ups billed at cost with prior approval.' },
    { h: 'Intellectual property', b: 'Design IP transfers to the Client upon full settlement of fees.' },
    { h: 'Validity', b: 'This proposal is valid for 30 days from the date of issue.' },
    { h: 'Currency & tax', b: 'All fees in USD, exclusive of VAT and other statutory charges.' },
    { h: 'Schedule', b: 'Timeline assumes timely Client feedback at each approval gate.' },
  ];

  terms.forEach((t, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = ML + col * 6.18;
    const y = 1.55 + row * 1.12;
    s.addShape('rect', { x, y, w: 5.9, h: 1.02, fill: { color: col === 0 ? C.rowalt : C.white }, line: { color: C.border, width: 0.5 } });
    s.addText(t.h, { x: x + 0.18, y: y + 0.1, w: 5.6, h: 0.26, fontSize: SZ.body, fontFace: F.bold, color: C.terradk, margin: 0 });
    s.addText(t.b, { x: x + 0.18, y: y + 0.38, w: 5.6, h: 0.55, fontSize: SZ.body, fontFace: F.light, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.5 });
  });

  pageNum(s, 26);
  copyright(s);
}

// ─── SLIDE 18 — THANK YOU ────────────────────────────────────────
function slide18(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  studioName(s);

  // Left image box (from Taseco: x=0, y=0, w≈4.6, h=7.5)
  if (d.closingImage) {
    s.addImage({ data: d.closingImage, x: 0, y: 0, w: 4.5, h: H, sizing: { type: 'cover', w: 4.5, h: H } });
  } else {
    s.addShape('rect', { x: 0, y: 0, w: 4.5, h: H, fill: { color: C.cream }, line: { color: C.graylt, width: 0.5 } });
    s.addText('[ CLOSING IMAGE ]', { x: 0, y: 0, w: 4.5, h: H, fontSize: 8, fontFace: F.light, color: C.gray, align: 'center', valign: 'middle', margin: 0 });
  }

  // Thank You
  s.addText('Thank You', {
    x: 5.2, y: 1.6, w: 7.7, h: 1.0,
    fontSize: 42, fontFace: F.serif, color: C.terra, margin: 0,
  });

  // CTA
  s.addText(d.ctaText || 'We would welcome a working session to walk through the zoning & layout options, align on scope & budget, and confirm the schedule — so we can move into Concept without delay.', {
    x: 5.2, y: 2.78, w: 7.7, h: 0.85,
    fontSize: SZ.body, fontFace: F.light, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.5,
  });

  // Contacts
  const contacts = [
    d.phone || '+84 9 6652 6662',
    `${d.website || 'www.h2xstudio.com.vn'} · ${d.email || 'info@h2xstudio.com'}`,
    d.addressHN || 'HQ — 26 Trần Hưng Đạo, Hà Nội',
    d.addressHCM || 'Branch — 199D Nguyễn Văn Hưởng, Thảo Điền, TP HCM',
  ];
  contacts.forEach((c, i) => {
    const y = 3.74 + i * 0.5;
    if (i > 0) s.addShape('rect', { x: 5.2, y: y - 0.06, w: 7.7, h: 0.008, fill: { color: C.rule }, line: { color: C.rule, width: 0 } });
    s.addText(c, {
      x: 5.2, y, w: 7.7, h: 0.38,
      fontSize: i === 0 ? 11 : SZ.body, fontFace: i === 0 ? F.bold : F.light,
      color: C.black, bold: i === 0, margin: 0,
    });
  });

  copyright(s);
}

// ─── MAIN EXPORT ─────────────────────────────────────────────────
// ─── SLIDE — HOURLY RATES & LITHOGRAPHY CHARGES ──────────────────
function slideHourlyRates(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Hourly Rates & Lithography Charges', 'Appendix C & D \u2014 additional services time charges & printing rates (USD)');

  const rates = d.hourlyRates || [
    { role: 'Principal / Creative Director', hr: '$29', day: '$230' },
    { role: 'Chief Operating Officer', hr: '$23', day: '$185' },
    { role: 'Senior Technical Expert', hr: '$25', day: '$200' },
    { role: 'Associate Director / Team Lead', hr: '$13', day: '$104' },
    { role: 'Deputy Team Lead', hr: '$11', day: '$88' },
    { role: 'Senior Designer', hr: '$10', day: '$81' },
    { role: 'Senior 3D Visualiser', hr: '$8', day: '$65' },
    { role: 'Designer / Technical Designer', hr: '$6', day: '$50' },
    { role: 'FF&E Specialist', hr: '$5', day: '$38' },
    { role: 'Accounting / Admin', hr: '$3.50', day: '$27' },
  ];

  // Left — Hourly / Daily Rates (native table — avoids a rendering bug seen with
  // many manually-drawn overlapping shapes on this text-dense slide)
  secLabel(s, 'APPENDIX C \u2014 HOURLY / DAILY RATES', ML, 1.5, 6.9);
  const rateRows = [
    [
      { text: 'ROLE', options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8, bold: false } },
      { text: 'PER HOUR', options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8, align: 'center' } },
      { text: 'PER DAY', options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8, align: 'center' } },
    ],
    ...rates.map((r, i) => {
      const bg = i % 2 === 0 ? C.rowalt : C.white;
      return [
        { text: r.role, options: { fill: { color: bg }, color: C.black, fontFace: F.light, fontSize: 7.5 } },
        { text: r.hr, options: { fill: { color: bg }, color: C.black, fontFace: F.light, fontSize: 7.5, align: 'center' } },
        { text: r.day, options: { fill: { color: bg }, color: C.black, fontFace: F.light, fontSize: 7.5, align: 'center' } },
      ];
    }),
  ];
  s.addTable(rateRows, {
    x: ML, y: 1.8, w: 6.9,
    colW: [4.6, 1.15, 1.15],
    rowH: [0.32, ...rates.map(() => 0.32)],
    border: { type: 'solid', color: C.border, pt: 0.5 },
    margin: [0.03, 0.06, 0.03, 0.06],
    valign: 'middle',
    autoPage: false,
  });

  // Right — Lithography Charges
  const rx = ML + 7.3, rw = CW - 7.3;
  secLabel(s, 'APPENDIX D \u2014 LITHOGRAPHY & PRINTING (USD)', rx, 1.5, rw);
  const litho = [
    { size: 'A4', bw: '$0.05', col: '$0.20', mnt: '$2.50' },
    { size: 'A3', bw: '$0.10', col: '$0.40', mnt: '$4.00' },
    { size: 'A2', bw: '$0.15', col: '$1.20', mnt: '$7.00' },
    { size: 'A1', bw: '$0.20', col: '$2.50', mnt: '$12' },
    { size: 'A0', bw: '$0.35', col: '$4.50', mnt: '$18' },
  ];
  const lithoRows = [
    [
      { text: 'SIZE', options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8, align: 'center' } },
      { text: 'B&W', options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8, align: 'center' } },
      { text: 'COLOUR', options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8, align: 'center' } },
      { text: 'MOUNT', options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8, align: 'center' } },
    ],
    ...litho.map((r, i) => {
      const bg = i % 2 === 0 ? C.rowalt : C.white;
      return [
        { text: r.size, options: { fill: { color: bg }, color: C.black, fontFace: F.light, fontSize: 7.5, align: 'center' } },
        { text: r.bw, options: { fill: { color: bg }, color: C.black, fontFace: F.light, fontSize: 7.5, align: 'center' } },
        { text: r.col, options: { fill: { color: bg }, color: C.black, fontFace: F.light, fontSize: 7.5, align: 'center' } },
        { text: r.mnt, options: { fill: { color: bg }, color: C.black, fontFace: F.light, fontSize: 7.5, align: 'center' } },
      ];
    }),
  ];
  s.addTable(lithoRows, {
    x: rx, y: 1.8, w: rw,
    colW: [0.95, 1.25, 1.25, 1.25],
    rowH: [0.32, ...litho.map(() => 0.32)],
    border: { type: 'solid', color: C.border, pt: 0.5 },
    margin: [0.03, 0.03, 0.03, 0.03],
    valign: 'middle',
    autoPage: false,
  });

  const litY = 1.8 + 0.32 + litho.length * 0.34 + 0.15;
  s.addText('Other: Soft copy transfer $12\u2013$19/time \u00b7 Binding $2\u2013$4/set \u00b7 Print resize $1/copy', {
    x: rx, y: litY, w: rw, h: 0.4,
    fontSize: 7.5, fontFace: F.light, color: C.gray, italic: true, margin: 0, lineSpacingMultiple: 1.25,
  });

  const noteY = 1.8 + 0.32 + rates.length * 0.36 + 0.15;
  s.addText('Rates exclude VAT and other applicable taxes \u00b7 Reference conversion 1 USD \u2248 26,000 VND \u00b7 Rates apply for the current year and may be revised periodically.', {
    x: ML, y: Math.max(noteY, litY + 0.5), w: CW, h: 0.4,
    fontSize: 7.5, fontFace: F.light, color: C.gray, italic: true, margin: 0, lineSpacingMultiple: 1.25,
  });

  pageNum(s, 27);
  copyright(s);
}

// ─── SLIDE — BIM FEE SCHEDULE ─────────────────────────────────────
function slideBIMFee(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'BIM Fee Schedule', 'Optional premium for delivering the project using Building Information Modelling');

  const bimRoles = d.bimRates || [
    { role: 'Partner / Design Director', hr: '$360', day: '$3,240', wk: '$16,200', mo: '$64,800' },
    { role: 'Associate Designer', hr: '$190', day: '$1,710', wk: '$8,550', mo: '$34,200' },
    { role: 'Senior Interior Designer / Senior FF&E', hr: '$130', day: '$1,170', wk: '$5,850', mo: '$23,400' },
    { role: 'Interior Designer / FF&E', hr: '$100', day: '$900', wk: '$4,500', mo: '$18,000' },
    { role: 'Senior CAD/BIM Technician', hr: '$70', day: '$630', wk: '$3,150', mo: '$12,600' },
  ];

  secLabel(s, 'BIM RATES \u2014 MONTHLY & HOURLY (USD)', ML, 1.5, CW);
  tblHeader(s, [
    { text: 'DESIGNATION', x: ML + 0.06, w: 4.0 },
    { text: 'HOURLY', x: ML + 4.1, w: 1.9, align: 'center' },
    { text: 'DAILY', x: ML + 6.05, w: 1.9, align: 'center' },
    { text: 'WEEKLY', x: ML + 8.0, w: 1.9, align: 'center' },
    { text: 'MONTHLY', x: ML + 9.95, w: 2.0, align: 'center' },
  ], 1.8, 0.32);
  bimRoles.forEach((r, i) => {
    const y = 1.8 + 0.32 + i * 0.4;
    tblRow(s, [
      { text: r.role, x: ML + 0.1, w: 3.95, fontSize: 8 },
      { text: r.hr, x: ML + 4.1, w: 1.85, align: 'center', fontSize: 8 },
      { text: r.day, x: ML + 6.05, w: 1.85, align: 'center', fontSize: 8 },
      { text: r.wk, x: ML + 8.0, w: 1.85, align: 'center', fontSize: 8 },
      { text: r.mo, x: ML + 9.95, w: 1.95, align: 'center', fontSize: 8 },
    ], y, 0.35, i % 2 === 0 ? C.rowalt : C.white);
  });

  const lodY = 1.8 + 0.32 + bimRoles.length * 0.4 + 0.25;
  secLabel(s, 'LOD STANDARDS AT H2X', ML, lodY, CW);
  const lods = [
    { l: 'LOD 100 (Concept Design)', d: 'Massing, general orientation and basic space allocations for broad visualisation.' },
    { l: 'LOD 200 (Schematic Design)', d: 'Preliminary dimensions, specific locations and relationships between model elements.' },
    { l: 'LOD 300/350 (Detailed Design)', d: 'Detailed geometry, accurate dimensions and clash detection across all consultants.' },
  ];
  const lw = (CW - 2 * 0.2) / 3;
  lods.forEach((item, i) => {
    const x = ML + i * (lw + 0.2);
    const y = lodY + 0.3;
    s.addShape('rect', { x, y, w: lw, h: 1.15, fill: { color: C.rowalt }, line: { color: C.border, width: 0.5 } });
    s.addText(item.l, { x: x + 0.12, y: y + 0.1, w: lw - 0.24, h: 0.4, fontSize: 8.5, fontFace: F.bold, color: C.terradk, margin: 0, valign: 'top', lineSpacingMultiple: 1.2 });
    s.addText(item.d, { x: x + 0.12, y: y + 0.5, w: lw - 0.24, h: 0.6, fontSize: 7.5, fontFace: F.light, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.25 });
  });

  s.addText('BIM authored in Autodesk Revit \u00b7 coordination via Navisworks \u00b7 aligned to ISO 19650-1/2 \u00b7 BIM delivery is an optional premium, quoted separately from the base Design Fee.', {
    x: ML, y: lodY + 1.6, w: CW, h: 0.35,
    fontSize: 7.5, fontFace: F.light, color: C.gray, italic: true, margin: 0, lineSpacingMultiple: 1.25,
  });

  pageNum(s, 28);
  copyright(s);
}

async function generateProposal(data) {
  const pres = makePres();
  pres.title = `H2X Studio — ${data.projectName || 'Fee Proposal'}`;
  pres.author = 'H2X Studio';

  slide01(pres, data);
  slide02(pres, data);
  slide03(pres, data);
  slide04(pres, data);
  slide05(pres, data);
  slide06(pres, data);
  slide07(pres, data);
  slide08(pres, data);
  slideWorkStageDeliverable(pres, data);  // Work Stage Deliverable (page 9)
  slideOrgChart(pres, data);              // Project Organisation Chart (page 10)
  slideDesignWorkPhase(pres, data);       // Design Work Phase (page 11)
  slideBriefing(pres, data);              // Briefing and Mobilisation (page 12)
  slideFeasibility(pres, data);           // Feasibility / Pre-Concept Design (page 13)
  slideConceptDesignExplainer(pres, data);// Concept Design (page 14)
  slideSchematicExplainer(pres, data);    // Schematic Design (page 15)
  slideDevelopmentExplainer(pres, data);  // Development Design (page 16)
  slideConstructionDocsExplainer(pres, data); // Construction Documents (page 17)

  slide11(pres, data);  // Document Issue Schedule (page 18)
  slide12(pres, data);  // Design Fee (page 19)
  slide09(pres, data);  // Detailed Design Stage (page 20)
  slide10(pres, data);  // Detailed SOW List (page 21)

  slide13(pres, data);  // Branding · Signage · AV — Scope & Value (page 22)
  slide14(pres, data);  // Add-on Workstreams — fee table (page 23)

  slide15(pres, data);  // Payment Schedule (page 24)
  slide16(pres, data);  // Inclusions & Exclusions (page 25)
  slide17(pres, data);  // Terms & Conditions (page 26)
  slideHourlyRates(pres, data); // Hourly Rates & Lithography Charges (page 27)
  slideBIMFee(pres, data);      // BIM Fee Schedule (page 28)
  slide18(pres, data);  // Thank You

  const fileName = path.join(os.tmpdir(), `H2X_FeeProposal_${(data.projectName || 'Project').replace(/\s+/g,'_')}_v${data.version || '01'}.pptx`);
  await pres.writeFile({ fileName });
  return fileName;
}

module.exports = { generateProposal };
