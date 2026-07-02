// H2X Studio — Fee Proposal Generator
// Layout: LAYOUT_WIDE 13.333" x 7.5"  |  1 inch = 914400 EMU
// Fonts: Playfair Display (headings) + SVN-Gilroy (body)

const pptxgen = require('pptxgenjs');
const FONTS = require('./fontData');

// ─── BRAND CONSTANTS ─────────────────────────────────────────────
const C = {
  black:      '141414',
  offwhite:   'F9F9F7',
  cream:      'FBFBF9',
  terracotta: 'C7A07F',
  terracottaDark: 'B8895F',
  gold:       'D6BBA6',
  gray:       '999999',
  grayLight:  'CCCCCC',
  grayMid:    '888888',
  white:      'FFFFFF',
  darkBg:     '1A1814',
  rowAlt:     'F2F0ED',
  rowTotal:   'E8E4DF',
  border:     'E0DDD8',
  rule:       'DEDBD5',
};

// Slide 13.333" x 7.5" (LAYOUT_WIDE)
const W = 13.333, H = 7.5;
const ML = 0.667; // margin left
const MR = 0.667; // margin right
const CW = W - ML - MR; // content width = 12.0"

const F = {
  serif:   'Playfair Display',
  sans:    'SVN-Gilroy',
};

// Font size helpers (matches Taseco exactly)
const SZ = {
  h1: 23,    // slide title
  h1lg: 32,  // cover title
  h2: 16,    // subtitle
  label: 9,  // section labels (spaced caps)
  body: 9,   // body text
  small: 6,  // copyright
  pagenum: 10,
  pullquote: 14,
};

// ─── HELPERS ─────────────────────────────────────────────────────
function makePres() {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
  return pres;
}

function logo(s, light = false) {
  const col = light ? C.white : C.black;
  const sub = light ? C.gold : C.terracotta;
  s.addText([
    { text: 'H2X', options: { bold: true, fontSize: 16, color: col, fontFace: F.serif } },
    { text: '\nSTUDIO', options: { fontSize: 7.5, color: sub, fontFace: F.sans, charSpacing: 4 } },
  ], { x: 0.5, y: 0.28, w: 1.4, h: 0.6, margin: 0 });
}

function pageNum(s, num, light = false) {
  const col = light ? C.gold : C.terracotta;
  s.addText(String(num).padStart(2, '0'), {
    x: W - 1.334, y: H - 0.42, w: 1.0, h: 0.3,
    fontSize: SZ.pagenum, fontFace: F.serif, color: col, align: 'right', margin: 0,
  });
}

function copyright(s, light = false) {
  s.addText('©H2X.Studio 2026. All Rights Reserved', {
    x: ML, y: H - 0.42, w: 9, h: 0.3,
    fontSize: SZ.small, fontFace: F.sans, color: light ? '555555' : C.grayLight, margin: 0,
  });
}

function slideTitle(s, title, subtitle, dark = false) {
  const bg = dark ? C.darkBg : C.offwhite;
  const tc = dark ? C.white : C.black;
  const sc = dark ? C.gold : C.terracotta;
  s.background = { color: bg };
  logo(s, dark);

  s.addText(title, {
    x: ML, y: 0.5, w: CW, h: 0.6,
    fontSize: SZ.h1, fontFace: F.serif, color: tc, margin: 0,
  });
  // Rule line
  s.addShape('rect', {
    x: ML, y: 1.04, w: CW, h: 0.008,
    fill: { color: dark ? '3A3530' : C.rule }, line: { color: dark ? '3A3530' : C.rule, width: 0 },
  });
  if (subtitle) {
    s.addText(subtitle, {
      x: ML, y: 1.084, w: CW, h: 0.231,
      fontSize: SZ.label, fontFace: F.sans, color: sc, margin: 0,
    });
  }
}

function sectionLabel(s, text, x, y, w, light = false) {
  s.addText(text, {
    x, y, w, h: 0.3,
    fontSize: SZ.label, fontFace: F.sans,
    color: light ? C.gold : C.terracotta,
    bold: false, charSpacing: 2, margin: 0,
  });
}

function bodyText(s, lines, x, y, w, h, color = null) {
  // lines: array of {text, bold?, italic?} or plain strings
  const items = lines.map((l, i) => {
    const isLast = i === lines.length - 1;
    if (typeof l === 'string') {
      return { text: l, options: { breakLine: !isLast, fontSize: SZ.body, fontFace: F.sans, color: color || C.black } };
    }
    return {
      text: l.text,
      options: {
        breakLine: !isLast,
        fontSize: SZ.body,
        fontFace: l.bold ? 'SVN-Gilroy Bold' : l.italic ? F.sans : F.sans,
        bold: !!l.bold,
        italic: !!l.italic,
        color: color || C.black,
      },
    };
  });
  s.addText(items, { x, y, w, h, margin: 0, valign: 'top' });
}

function tableHeaderRow(s, cells, y, rowH = 0.3) {
  // cells: [{text, x, w}]
  s.addShape('rect', {
    x: ML, y, w: CW, h: rowH,
    fill: { color: C.black }, line: { color: C.black, width: 0 },
  });
  cells.forEach(cell => {
    s.addText(cell.text, {
      x: cell.x, y: y + 0.04, w: cell.w, h: rowH - 0.06,
      fontSize: SZ.label, fontFace: F.sans, color: C.white,
      charSpacing: 2, bold: false, margin: 0, valign: 'middle',
      align: cell.align || 'left',
    });
  });
}

function tableRow(s, cells, y, rowH, bg) {
  s.addShape('rect', {
    x: ML, y, w: CW, h: rowH,
    fill: { color: bg }, line: { color: C.border, width: 0.5 },
  });
  cells.forEach(cell => {
    s.addText(cell.text, {
      x: cell.x, y: y + 0.06, w: cell.w, h: rowH - 0.1,
      fontSize: cell.fontSize || SZ.body,
      fontFace: cell.bold ? 'SVN-Gilroy SemiBold' : F.sans,
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
  s.background = { color: C.offwhite };
  logo(s, false);

  s.addText('H2X Studio', {
    x: W - 5.0, y: 0.289, w: 4.666, h: 0.5,
    fontSize: 22, fontFace: F.serif, color: C.black, align: 'right', margin: 0,
  });

  // Top rule
  s.addShape('rect', { x: ML, y: 0.85, w: CW, h: 0.008, fill: { color: C.rule }, line: { color: C.rule, width: 0 } });

  // Hero image placeholder (user will replace in PPT)
  s.addShape('rect', {
    x: 2.0, y: 0.95, w: 9.333, h: 4.0,
    fill: { color: '2A2520' }, line: { color: '3A3530', width: 1 },
  });
  s.addText('[ HERO IMAGE — thay bằng ảnh dự án ]', {
    x: 2.0, y: 0.95, w: 9.333, h: 4.0,
    fontSize: 9, fontFace: F.sans, color: '444038',
    align: 'center', valign: 'middle', margin: 0,
  });

  // Project name
  s.addText(d.projectName || 'Tên Dự Án', {
    x: ML, y: 5.04, w: CW, h: 0.5,
    fontSize: SZ.h1lg, fontFace: F.serif, color: C.white, align: 'center', margin: 0,
  });
  s.addText((d.location || 'ĐỊA ĐIỂM').toUpperCase(), {
    x: ML, y: 5.6, w: CW, h: 0.3,
    fontSize: 12, fontFace: F.sans, color: C.terracotta,
    align: 'center', charSpacing: 3, margin: 0,
  });
  s.addText('INTERIOR DESIGN SERVICES', {
    x: ML, y: 6.0, w: CW, h: 0.3,
    fontSize: 14, fontFace: F.serif, color: C.grayLight, align: 'center', margin: 0,
  });
  s.addText('Fee Proposal', {
    x: ML, y: 6.34, w: CW, h: 0.3,
    fontSize: 14, fontFace: F.serif, color: C.white, align: 'center', bold: true, margin: 0,
  });

  s.addText(`VERSION ${d.version || '01'}`, {
    x: ML, y: H - 0.45, w: 4, h: 0.3,
    fontSize: 10, fontFace: F.sans, color: C.gray, charSpacing: 2, margin: 0,
  });
  s.addText(d.month || 'JUNE 2026', {
    x: W - 5.0, y: H - 0.45, w: 4.333, h: 0.3,
    fontSize: 10, fontFace: F.sans, color: C.gray, align: 'right', charSpacing: 2, margin: 0,
  });
  copyright(s, true);
}

// ─── SLIDE 02 — PROJECT UNDERSTANDING ───────────────────────────
function slide02(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Project Understanding', `${d.clientName} · ${d.typology} · ${d.location}`);

  // 2-column layout: left 6.156" | right 5.293" (matches Taseco)
  const lw = 6.156, rw = 5.72, rx = 7.144;
  const contentY = 1.6;

  // Left column
  sectionLabel(s, 'THE OPPORTUNITY', ML, contentY, lw);
  s.addText(d.opportunity || '[Mô tả cơ hội dự án — tầm quan trọng và tiềm năng]', {
    x: ML, y: contentY + 0.34, w: lw, h: 2.8,
    fontSize: SZ.body, fontFace: F.sans, color: C.black, margin: 0, valign: 'top',
  });

  sectionLabel(s, 'PROJECT BRIEF', ML, contentY + 2.8 + 0.5, lw);
  const briefLines = [
    { text: '+ ', bold: true },
    { text: `Client · ${d.clientName || '[Client]'}` },
    { text: `\n+ `, bold: true },
    { text: `Location · ${d.location || '[Location]'}` },
    { text: `\n+ `, bold: true },
    { text: `Typology · ${d.typology || '[Typology]'}` },
    { text: `\n+ `, bold: true },
    { text: `Area · ~${d.area || '[X]'} m²` },
    { text: `\n+ `, bold: true },
    { text: `Target guest · ${d.targetGuest || '[Description]'}` },
  ];
  // Render brief as rich text
  const briefItems = briefLines.map((l, i) => ({
    text: l.text,
    options: {
      fontSize: SZ.body,
      fontFace: l.bold ? 'SVN-Gilroy Bold' : F.sans,
      color: C.black,
    },
  }));
  s.addText(briefItems, {
    x: ML, y: contentY + 3.5, w: lw, h: 2.2,
    margin: 0, valign: 'top',
  });

  // Right column
  sectionLabel(s, 'EXPERIENCE INTENT', rx, contentY, rw);
  s.addText(d.experienceIntent || '[Định hướng trải nghiệm — cảm xúc và hành trình khách]', {
    x: rx, y: contentY + 0.34, w: rw, h: 2.8,
    fontSize: SZ.body, fontFace: F.sans, color: C.black, margin: 0, valign: 'top',
  });

  sectionLabel(s, 'WHY H2X', rx, contentY + 2.8 + 0.5, rw);
  s.addText(d.whyH2x || 'H2X brings deep hospitality experience, integrated Lighting Design, and a proven track record across Vietnam\'s premium hospitality typologies — ready to begin.', {
    x: rx, y: contentY + 3.5, w: rw, h: 1.8,
    fontSize: SZ.body, fontFace: F.sans, color: C.black, margin: 0, valign: 'top',
  });

  pageNum(s, 2);
  copyright(s);
}

// ─── SLIDE 03 — SCOPE OF SERVICES ───────────────────────────────
function slide03(pres, d) {
  const s = pres.addSlide();
  const scope = d.scope || 'A'; // A = Concept→Schematic, B = Full Services
  slideTitle(s, 'Scope of Services', null);

  // Subtitle after rule
  s.addText(
    scope === 'A'
      ? 'This engagement covers Concept & Schematic Design, with Lighting Design integrated throughout'
      : 'Full-service engagement — Concept through Construction Documents & Site Supervision',
    { x: ML, y: 1.15, w: CW, h: 0.25, fontSize: SZ.label, fontFace: F.sans, color: C.terracotta, margin: 0 }
  );

  const stagesA = [
    { stage: 'Briefing & Mobilization', status: '✓ Included', desc: 'Kick-off, brief alignment, A&B material review, site assessment.' },
    { stage: 'Concept Design + Lighting Concept', status: '✓ Included', desc: 'Zoning & space planning, design narrative, moodboards, 3D key views, material palette, preliminary lighting concept.' },
    { stage: 'Schematic Design + Lighting Schematic', status: '✓ Included', desc: 'Annotated plans, RCP & floor-finish plans, key elevations, outline specification schedules, schematic lighting.' },
    { stage: 'Design Development · Construction Documents · BOQ', status: 'Optional', desc: 'Available as a follow-on package — fees proposed by addendum at the agreed rate once Schematic is approved.' },
    { stage: 'Tender Support · Site Supervision', status: 'Optional', desc: 'Offered separately on request.' },
  ];

  const stagesB = [
    { stage: 'Briefing & Mobilization', status: '✓ Included', desc: 'Kick-off, brief alignment, A&B material review, site assessment.' },
    { stage: 'Concept Design + Lighting Concept', status: '✓ Included', desc: 'Zoning & space planning, design narrative, moodboards, 3D key views, material palette, preliminary lighting concept.' },
    { stage: 'Schematic Design + Lighting Schematic', status: '✓ Included', desc: 'Annotated plans, RCP & floor-finish plans, key elevations, outline specs, schematic lighting.' },
    { stage: 'Design Development', status: '✓ Included', desc: 'Detailed drawings, material schedules, FF&E specifications, consultant coordination.' },
    { stage: 'Construction Documents + BOQ', status: '✓ Included', desc: 'Full CD set, tender-ready bill of quantities.' },
    { stage: 'Tender Support · Site Supervision', status: '✓ Included', desc: 'Bid review, RFI responses, site visits through to completion.' },
  ];

  const stages = scope === 'B' ? stagesB : stagesA;

  // Header row
  const hdrY = 1.48;
  tableHeaderRow(s, [
    { text: 'STAGE IN SCOPE', x: ML + 0.06, w: 4.5 },
    { text: 'WHAT IT DELIVERS', x: ML + 5.0, w: 7.0 },
  ], hdrY, 0.35);

  stages.forEach((st, i) => {
    const y = hdrY + 0.35 + i * 0.82;
    const bg = i % 2 === 0 ? C.rowAlt : C.offwhite;
    const isOptional = st.status === 'Optional';

    s.addShape('rect', {
      x: ML, y, w: CW, h: 0.78,
      fill: { color: bg }, line: { color: C.border, width: 0.5 },
    });

    // Stage name
    s.addText(st.stage, {
      x: ML + 0.1, y: y + 0.1, w: 4.4, h: 0.56,
      fontSize: SZ.body, fontFace: isOptional ? F.sans : F.sans,
      color: isOptional ? C.grayMid : C.black,
      italic: isOptional, bold: !isOptional,
      margin: 0, valign: 'top',
    });

    // Status badge
    s.addText(st.status, {
      x: ML + 4.6, y: y + 0.2, w: 0.8, h: 0.3,
      fontSize: 8, fontFace: F.sans,
      color: isOptional ? C.gray : C.terracottaDark,
      bold: !isOptional, margin: 0, valign: 'middle', align: 'center',
    });

    // Description
    s.addText(st.desc, {
      x: ML + 5.0, y: y + 0.1, w: 7.1, h: 0.56,
      fontSize: SZ.body, fontFace: F.sans,
      color: isOptional ? C.grayMid : C.black,
      margin: 0, valign: 'top',
    });
  });

  pageNum(s, 3);
  copyright(s);
}

// ─── SLIDE 04 — ZONING PROGRAMME ────────────────────────────────
function slide04(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Indicative Zoning Programme', null);

  const zones = d.zones || [
    { zone: 'Welcome & Concierge', area: '', seats: '', rationale: 'Single access-control point, app/face check-in, welcome ritual; staff sightline over arrival.' },
    { zone: 'Main Social & Relaxation', area: '', seats: '', rationale: 'Varied seating; apron/runway view on premium row; the social heart.' },
    { zone: 'Premium F&B', area: '', seats: '', rationale: 'Buffet + show kitchen + bar, adjacent to BOH for the shortest service run.' },
    { zone: 'Quiet / Wellness / Recovery', area: '', seats: '', rationale: 'Nap & shower suites, relaxation; acoustically separated.' },
    { zone: 'Business & Productivity', area: '', seats: '', rationale: 'Workstations + meeting room, buffering social and quiet.' },
    { zone: 'Amenities (WC / Family / Prayer)', area: '', seats: '', rationale: 'Distributed near both social and quiet zones; fully accessible.' },
    { zone: 'Back-of-House', area: '', seats: '', rationale: 'Kitchen, prep, store, staff, waste — service corridor never crosses guest flow.' },
    { zone: 'Pre-boarding / Exit', area: '', seats: '', rationale: 'Smooth exit to gate; boarding displays; proactive flight reminders.' },
    { zone: `Total (net usable)  ~${d.area || '[X]'} m²`, area: `~${d.area || '[X]'}`, seats: '', rationale: 'Circulation (~15%) distributed within each zone.' },
  ];

  const hdrY = 1.48;
  tableHeaderRow(s, [
    { text: 'ZONE', x: ML + 0.06, w: 3.5 },
    { text: 'AREA', x: ML + 3.7, w: 1.0 },
    { text: 'SEATS', x: ML + 4.8, w: 1.0 },
    { text: 'OPERATIONAL RATIONALE', x: ML + 5.9, w: 6.2 },
  ], hdrY, 0.35);

  zones.forEach((z, i) => {
    const isTotal = i === zones.length - 1;
    const y = hdrY + 0.35 + i * 0.56;
    const bg = isTotal ? C.rowTotal : i % 2 === 0 ? C.rowAlt : C.offwhite;

    s.addShape('rect', {
      x: ML, y, w: CW, h: 0.52,
      fill: { color: bg }, line: { color: C.border, width: 0.5 },
    });

    [
      { text: isTotal ? `Total (net usable)` : z.zone, x: ML + 0.1, w: 3.4, bold: isTotal, color: isTotal ? C.terracottaDark : C.black },
      { text: z.area || '—', x: ML + 3.7, w: 0.95, align: 'center', bold: isTotal },
      { text: z.seats || '—', x: ML + 4.8, w: 0.95, align: 'center' },
      { text: isTotal ? `~${d.area || '[X]'} m² · Circulation (~15%) distributed within each zone.` : z.rationale, x: ML + 5.9, w: 6.15, color: isTotal ? C.grayMid : C.black, italic: isTotal },
    ].forEach(cell => {
      s.addText(cell.text, {
        x: cell.x, y: y + 0.1, w: cell.w, h: 0.32,
        fontSize: SZ.body, fontFace: cell.bold ? 'SVN-Gilroy Bold' : F.sans,
        color: cell.color || C.black, bold: !!cell.bold, italic: !!cell.italic,
        align: cell.align || 'left', margin: 0, valign: 'middle',
      });
    });
  });

  s.addText(`H2X proposal for ~${d.area || '[X]'} m² · ~${d.peakGuests || '[n]'} guests at peak · ~${d.sqmPerGuest || '7'} m²/guest (premium benchmark) · to be confirmed in Concept`, {
    x: ML, y: H - 0.6, w: CW, h: 0.22,
    fontSize: 7.5, fontFace: F.sans, color: C.gray, italic: true, margin: 0,
  });

  pageNum(s, 4);
  copyright(s);
}

// ─── SLIDE 05 — DESIGN LANGUAGE ─────────────────────────────────
function slide05(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.offwhite };
  logo(s);

  s.addText('Design Language', {
    x: ML, y: 0.5, w: CW, h: 0.55,
    fontSize: SZ.h1, fontFace: F.serif, color: C.black, margin: 0,
  });
  s.addText(d.designTagline || 'Quiet luxury with a Vietnamese soul — warm materials, light from shadow, restraint over display', {
    x: ML, y: 1.06, w: CW, h: 0.25,
    fontSize: SZ.label, fontFace: F.sans, color: C.terracotta, margin: 0,
  });
  s.addShape('rect', { x: ML, y: 1.35, w: CW, h: 0.008, fill: { color: C.rule }, line: { color: C.rule, width: 0 } });

  // 3 mood image boxes
  const imgW = 3.9, imgH = 4.5;
  const imgPillars = d.designPillars || [
    { label: 'WELCOME', sub: 'ARRIVAL RITUAL · THE HUMAN TOUCH' },
    { label: 'F&B', sub: 'REFINED · SENSORY · UNHURRIED' },
    { label: 'SOCIAL', sub: 'WARM · ATMOSPHERIC · CONNECTED' },
  ];

  imgPillars.forEach((p, i) => {
    const x = ML + i * (imgW + 0.2);
    s.addShape('rect', {
      x, y: 1.45, w: imgW, h: imgH,
      fill: { color: '2A2520' }, line: { color: '3A3530', width: 1 },
    });
    s.addText('[ MOOD IMAGE ]', {
      x, y: 1.45, w: imgW, h: imgH,
      fontSize: 8, fontFace: F.sans, color: '444038',
      align: 'center', valign: 'middle', margin: 0,
    });
    s.addText(p.label, {
      x, y: 6.05, w: imgW, h: 0.25,
      fontSize: 8.5, fontFace: 'SVN-Gilroy Bold', color: C.black,
      bold: true, charSpacing: 2, margin: 0,
    });
    s.addText(p.sub, {
      x, y: 6.3, w: imgW, h: 0.22,
      fontSize: 7.5, fontFace: F.sans, color: C.gray, charSpacing: 1, margin: 0,
    });
  });

  pageNum(s, 5);
  copyright(s);
}

// ─── SLIDE 06 — LIGHTING DESIGN ──────────────────────────────────
function slide06(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.offwhite };
  logo(s, false);

  s.addText('Lighting Design — Built In, From Concept', {
    x: ML, y: 0.5, w: CW, h: 0.6,
    fontSize: SZ.h1, fontFace: F.serif, color: C.black, margin: 0,
  });
  s.addShape('rect', { x: ML, y: 1.04, w: CW, h: 0.008, fill: { color: C.rule }, line: { color: C.rule, width: 0 } });

  s.addText('"Light is not a final layer — it is the atmosphere. So we design it from the first Concept sketch, not after the walls are built."', {
    x: ML, y: 1.1, w: CW, h: 0.6,
    fontSize: SZ.pullquote, fontFace: F.serif, color: C.terracotta, italic: true, align: 'center', margin: 0,
  });

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
      x: 6.947, w: 5.72, label: 'WHAT H2X DELIVERS (IN THE FEE)',
      items: [
        d.lightingStrategy || 'Zoned atmosphere strategy — dramatic for F&B, functional for business, circadian/calming for Quiet & Wellness',
        'Daylight integration; warm indirect & cove lighting',
        'Fixture & control concept; scene-setting by time of day & flight wave',
        'Vietnam energy-code compliant; premium, buildable fixtures via our long-standing supplier network',
      ],
    },
  ];

  cols.forEach(col => {
    sectionLabel(s, col.label, col.x, 2.0, col.w, false);
    col.items.forEach((item, i) => {
      s.addText([
        { text: '+ ', options: { color: C.terracotta, bold: true, fontFace: 'SVN-Gilroy Bold', fontSize: SZ.body } },
        { text: item, options: { color: C.white, fontFace: F.sans, fontSize: SZ.body } },
      ], {
        x: col.x, y: 2.38 + i * 0.86, w: col.w, h: 0.78,
        margin: 0, valign: 'top',
      });
    });
  });

  s.addText('Included in the design fee · H2X\'s in-house strength · the single biggest lever on atmosphere', {
    x: ML, y: H - 0.6, w: CW, h: 0.22,
    fontSize: 7.5, fontFace: F.sans, color: C.gray, italic: true, margin: 0,
  });

  pageNum(s, 6, false);
  copyright(s, false);
}

// ─── SLIDE 07 — PROJECT TEAM ──────────────────────────────────────
function slide07(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Project Team', 'A senior, multidisciplinary team — with in-house Lighting Design');

  const team = d.team || [
    { role: 'CREATIVE DIRECTOR', name: 'Gia Huy (Michael)', title: 'M.Architect · 23+ yrs · sets vision & design direction', dark: true },
    { role: 'LEAD DESIGNER — CONCEPT', name: 'Nguyễn Mạnh Hùng (Henry)', title: 'Leads spatial concept, narrative & guest journey' },
    { role: 'LIGHTING DESIGNER (IN-HOUSE)', name: '[Lighting Lead]', title: 'Atmosphere, fixture & control strategy from Concept' },
    { role: 'HEAD OF FF&E', name: 'Nguyễn Thị Thúy Vy (Vivian)', title: 'Furniture, fixtures & material curation' },
    { role: 'TECHNICAL LEAD', name: 'Nguyễn Vương Linh (Lucas)', title: 'Schematic documentation & consultant coordination' },
    { role: '3D VISUALISATION & PM', name: 'Đỗ Danh Sơn', title: 'Photorealistic renders · schedule, reviews & delivery' },
  ];

  const cardW = 3.8, cardH = 1.55;
  team.forEach((m, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = ML + col * (cardW + 0.2);
    const y = 1.45 + row * (cardH + 0.18);

    s.addShape('rect', {
      x, y, w: cardW, h: cardH,
      fill: { color: m.dark ? C.darkBg : C.rowAlt },
      line: { color: C.border, width: 0.5 },
      shadow: { type: 'outer', color: '000000', blur: 8, offset: 2, angle: 45, opacity: 0.08 },
    });
    s.addText(m.role, {
      x: x + 0.18, y: y + 0.15, w: cardW - 0.3, h: 0.25,
      fontSize: 7.5, fontFace: F.sans, color: m.dark ? C.terracotta : C.terracottaDark,
      charSpacing: 1.5, margin: 0,
    });
    s.addText(m.name, {
      x: x + 0.18, y: y + 0.44, w: cardW - 0.3, h: 0.38,
      fontSize: 13, fontFace: F.serif, color: m.dark ? C.white : C.black, margin: 0,
    });
    s.addText(m.title, {
      x: x + 0.18, y: y + 0.86, w: cardW - 0.3, h: 0.56,
      fontSize: SZ.body, fontFace: F.sans, color: m.dark ? C.grayLight : C.gray,
      margin: 0, valign: 'top',
    });
  });

  pageNum(s, 7);
  copyright(s);
}

// ─── SLIDE 08 — DESIGN PROCESS ───────────────────────────────────
function slide08(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Design Process & Workflow', 'Approval-gated at every stage · this engagement runs to Schematic');

  const scope = d.scope || 'A';
  const stages = scope === 'B'
    ? [
        { label: 'Briefing &\nMobilization', sub: '', dark: false },
        { label: 'Concept Design', sub: '+ Lighting', dark: false },
        { label: 'Schematic Design', sub: '+ Lighting', dark: false },
        { label: 'Design Development', sub: '', dark: false },
        { label: 'Construction\nDocuments', sub: '', dark: true },
      ]
    : [
        { label: 'Briefing &\nMobilization', sub: '', dark: false },
        { label: 'Concept Design', sub: '+ Lighting', dark: false },
        { label: 'Schematic Design', sub: '+ Lighting', dark: true },
      ];

  const boxW = scope === 'B' ? 2.28 : 3.85;
  const gap = scope === 'B' ? 0.18 : 0.3;

  stages.forEach((st, i) => {
    const x = ML + i * (boxW + gap);
    s.addShape('rect', {
      x, y: 1.45, w: boxW, h: 1.0,
      fill: { color: st.dark ? C.black : C.offwhite },
      line: { color: st.dark ? C.black : C.grayLight, width: 1 },
    });
    s.addText(st.label, {
      x: x + 0.1, y: 1.52, w: boxW - 0.2, h: 0.56,
      fontSize: 12, fontFace: F.serif, color: st.dark ? C.white : C.black,
      align: 'center', margin: 0,
    });
    if (st.sub) {
      s.addText(st.sub, {
        x: x + 0.1, y: 2.1, w: boxW - 0.2, h: 0.28,
        fontSize: 10, fontFace: F.sans, color: C.terracotta, align: 'center', margin: 0,
      });
    }
    if (i < stages.length - 1) {
      s.addText('→', {
        x: x + boxW + 0.01, y: 1.78, w: gap, h: 0.35,
        fontSize: 14, fontFace: F.sans, color: C.grayMid, align: 'center', margin: 0,
      });
    }
  });

  // 3 info cards
  const infoW = 3.85;
  const infos = [
    { h: 'Cadence', body: 'Each stage: a First Draft, one consolidated Review round, then Final Approval & a stage-completion sign-off.\n\nAdditional rounds beyond this are by addendum (time + fee).' },
    { h: 'Tools', body: 'Revit + Enscape; photorealistic 3D key views from Concept so decisions are made on what you will actually see.' },
    { h: 'Indicative Timeline', body: `Mobilization ~1 wk · Concept ~4 wks · Schematic ~4 wks${scope === 'B' ? ' · DD ~6 wks · CD/BOQ ~8 wks' : ''}, plus client-approval gates — about ${scope === 'B' ? '24–28' : '9–11'} weeks.` },
  ];
  infos.forEach((inf, i) => {
    const x = ML + i * (infoW + 0.3);
    s.addShape('rect', {
      x, y: 2.8, w: infoW, h: 2.3,
      fill: { color: C.rowAlt }, line: { color: C.border, width: 0.5 },
    });
    s.addText(inf.h, {
      x: x + 0.18, y: 2.95, w: infoW - 0.3, h: 0.32,
      fontSize: 12, fontFace: F.serif, color: C.black, margin: 0,
    });
    s.addText(inf.body, {
      x: x + 0.18, y: 3.3, w: infoW - 0.3, h: 1.65,
      fontSize: SZ.body, fontFace: F.sans, color: C.black, margin: 0, valign: 'top',
    });
  });

  s.addText(`Methodology: Brief → Concept (approval) → Schematic (approval).${scope === 'B' ? ' Design Development → Construction Documents → BOQ → Supervision follow the same gated rhythm.' : ' Should the project continue, Design Development → Construction Documents → BOQ follow the same gated rhythm.'}`, {
    x: ML, y: H - 0.55, w: CW, h: 0.22,
    fontSize: 7.5, fontFace: F.sans, color: C.gray, margin: 0,
  });

  pageNum(s, 8);
  copyright(s);
}

// ─── SLIDE 09 — DESIGN FEE ───────────────────────────────────────
function slide09(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Design Fee', 'Interior Design + Lighting Design · lump sum by phase');

  const scope = d.scope || 'A';
  const fee = parseFloat(d.feeTotal) || 0;

  const stagesA = [
    { num: '1', stage: 'Briefing & Mobilization', ratio: '15%', fee: Math.round(fee * 0.15) },
    { num: '2', stage: 'Concept Design (incl. Lighting Concept)', ratio: '50%', fee: Math.round(fee * 0.50) },
    { num: '3', stage: 'Schematic Design (incl. Lighting Schematic)', ratio: '35%', fee: Math.round(fee * 0.35) },
  ];

  const stagesB = [
    { num: '1', stage: 'Briefing & Mobilization', ratio: '5%', fee: Math.round(fee * 0.05) },
    { num: '2', stage: 'Concept Design (incl. Lighting Concept)', ratio: '20%', fee: Math.round(fee * 0.20) },
    { num: '3', stage: 'Schematic Design (incl. Lighting Schematic)', ratio: '20%', fee: Math.round(fee * 0.20) },
    { num: '4', stage: 'Design Development', ratio: '25%', fee: Math.round(fee * 0.25) },
    { num: '5', stage: 'Construction Documents + BOQ', ratio: '20%', fee: Math.round(fee * 0.20) },
    { num: '6', stage: 'Tender Support + Site Supervision', ratio: '10%', fee: Math.round(fee * 0.10) },
  ];

  const stages = scope === 'B' ? stagesB : stagesA;
  const fmt = n => n > 0 ? n.toLocaleString('en-US') : '[—]';

  // Header
  const hdrY = 1.48;
  tableHeaderRow(s, [
    { text: 'STAGE', x: ML + 0.06, w: 0.5 },
    { text: 'PHASE', x: ML + 0.7, w: 7.0 },
    { text: 'RATIO', x: ML + 8.1, w: 1.5 },
    { text: 'FEE (USD)', x: ML + 9.7, w: 2.5, align: 'right' },
  ], hdrY, 0.35);

  stages.forEach((st, i) => {
    const y = hdrY + 0.35 + i * 0.56;
    const bg = i % 2 === 0 ? C.rowAlt : C.offwhite;
    tableRow(s, [
      { text: st.num, x: ML + 0.1, w: 0.5, align: 'center' },
      { text: st.stage, x: ML + 0.7, w: 7.0 },
      { text: st.ratio, x: ML + 8.1, w: 1.5, align: 'center' },
      { text: `${fmt(st.fee)}`, x: ML + 9.7, w: 2.4, align: 'right' },
    ], y, 0.52, bg);
  });

  // Subtotal row
  const totalY = hdrY + 0.35 + stages.length * 0.56;
  s.addShape('rect', { x: ML, y: totalY, w: CW, h: 0.52, fill: { color: C.rowTotal }, line: { color: C.border, width: 0.5 } });
  s.addText('INTERIOR + LIGHTING — SUBTOTAL', {
    x: ML + 0.7, y: totalY + 0.1, w: 7.5, h: 0.32,
    fontSize: SZ.body, fontFace: 'SVN-Gilroy Bold', color: C.terracottaDark, bold: true, margin: 0,
  });
  s.addText('100%', { x: ML + 8.1, y: totalY + 0.1, w: 1.5, h: 0.32, fontSize: SZ.body, fontFace: 'SVN-Gilroy Bold', color: C.terracottaDark, bold: true, align: 'center', margin: 0 });
  s.addText(fmt(fee), { x: ML + 9.7, y: totalY + 0.1, w: 2.4, h: 0.32, fontSize: SZ.body, fontFace: 'SVN-Gilroy Bold', color: C.terracottaDark, bold: true, align: 'right', margin: 0 });

  // Benchmark note
  const area = parseFloat(d.area) || 0;
  const perSqm = area > 0 ? Math.round(fee / area) : 0;
  s.addText(`≈ ${perSqm > 0 ? perSqm + ' USD / m²' : '[USD/m²]'} · ${d.area || '[X]'} m²`, {
    x: W - 5.5, y: totalY + 0.62, w: 5.0, h: 0.3,
    fontSize: SZ.body, fontFace: 'SVN-Gilroy Bold', color: C.black, align: 'right', margin: 0,
  });
  s.addText('Indicative benchmark for this typology (Lighting included). Excludes taxes & other fees.\nDesign Development / Construction Documents / BOQ available as a follow-on addendum.', {
    x: ML, y: totalY + 0.62, w: 8.5, h: 0.5,
    fontSize: 8, fontFace: F.sans, color: C.gray, margin: 0,
  });

  pageNum(s, 9);
  copyright(s);
}

// ─── SLIDE 10 — ADD-ON VALUE ─────────────────────────────────────
function slide10(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Branding · Signage · AV — Scope & Value', null);
  s.addText('Three workstreams that turn a beautiful space into a coherent, memorable destination — offered at concept stage', {
    x: ML, y: 1.15, w: CW, h: 0.25, fontSize: SZ.label, fontFace: F.sans, color: C.terracotta, margin: 0,
  });

  const addons = [
    {
      title: 'Brand Identity & Naming', partner: 'Ziva',
      scope: ['Lounge name & story', 'Logo & core visual identity', 'Brand guidelines (essentials)'],
      why: ['A named space earns loyalty, PR & partnership value', 'One coherent guest-facing story across every touchpoint'],
    },
    {
      title: 'Signage & Wayfinding', partner: 'Ziva',
      scope: ['Bilingual wayfinding strategy', 'Signage concept & key types', 'Specification direction'],
      why: ['Effortless navigation lowers anxiety & staff load', 'Signage is the brand at eye level — venue-grade legibility'],
    },
    {
      title: 'Audio-Visual Design', partner: 'H2X',
      scope: ['Signature soundscape direction', 'AV & display integration concept', 'Zoning of sound by area'],
      why: ['Sound completes the atmosphere lighting begins', 'AV as ambience, not noise — reliable, calm information'],
    },
  ];

  const cw = 3.85;
  addons.forEach((a, i) => {
    const x = ML + i * (cw + 0.2);
    s.addShape('rect', {
      x, y: 1.48, w: cw, h: 5.4,
      fill: { color: C.rowAlt }, line: { color: C.border, width: 0.5 },
    });
    s.addText(a.title, { x: x + 0.18, y: 1.62, w: cw - 0.3, h: 0.5, fontSize: 13, fontFace: F.serif, color: C.black, margin: 0 });
    s.addText(`Partner: ${a.partner}`, { x: x + 0.18, y: 2.15, w: cw - 0.3, h: 0.22, fontSize: 8, fontFace: 'SVN-Gilroy Bold', color: C.terracotta, bold: true, margin: 0 });
    sectionLabel(s, 'SCOPE (1 CONCEPT)', x + 0.18, 2.44, cw - 0.3);
    const scopeItems = a.scope.map((l, idx) => ({
      text: l, options: { bullet: true, breakLine: idx < a.scope.length - 1, fontSize: SZ.body, fontFace: F.sans, color: C.black },
    }));
    s.addText(scopeItems, { x: x + 0.18, y: 2.72, w: cw - 0.3, h: 1.0, margin: 0, valign: 'top' });
    sectionLabel(s, 'WHY IT MATTERS', x + 0.18, 3.82, cw - 0.3);
    const whyItems = a.why.map((l, idx) => ({
      text: l, options: { bullet: true, breakLine: idx < a.why.length - 1, fontSize: SZ.body, fontFace: F.sans, color: C.black },
    }));
    s.addText(whyItems, { x: x + 0.18, y: 4.1, w: cw - 0.3, h: 1.6, margin: 0, valign: 'top' });
  });

  pageNum(s, 10);
  copyright(s);
}

// ─── SLIDE 11 — ADD-ON FEE TABLE ─────────────────────────────────
function slide11(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Add-on Workstreams', 'Concept stage · itemised so each can be selected independently · design fees only (fabrication/production excluded)');

  const feeMain = parseFloat(d.feeTotal) || 0;
  const feeBrand = parseFloat(d.feeBrand) || 0;
  const feeSign = parseFloat(d.feeSignage) || 0;
  const feeAV = parseFloat(d.feeAV) || 0;
  const feeAddons = feeBrand + feeSign + feeAV;
  const feeGrand = feeMain + feeAddons;
  const feeDiscount = Math.round(feeGrand * 0.9);
  const fmt = n => n > 0 ? n.toLocaleString('en-US') : '[—]';

  const rows = [
    { ws: 'Brand Identity & Naming', partner: 'Ziva', scope: 'Naming, logo, core identity & essential guidelines.', fee: feeBrand },
    { ws: 'Signage & Wayfinding', partner: 'Ziva', scope: 'Bilingual wayfinding strategy & signage concept.', fee: feeSign },
    { ws: 'Audio-Visual Design', partner: 'H2X', scope: 'Soundscape, AV & display concept.', fee: feeAV },
  ];

  const hdrY = 1.48;
  tableHeaderRow(s, [
    { text: 'WORKSTREAM', x: ML + 0.06, w: 3.0 },
    { text: 'PARTNER', x: ML + 3.15, w: 1.1 },
    { text: 'SCOPE (CONCEPT, 1 DIRECTION)', x: ML + 4.35, w: 5.3 },
    { text: 'FEE (USD)', x: ML + 9.75, w: 2.35, align: 'right' },
  ], hdrY, 0.35);

  rows.forEach((r, i) => {
    const y = hdrY + 0.35 + i * 0.56;
    tableRow(s, [
      { text: r.ws, x: ML + 0.1, w: 3.0 },
      { text: r.partner, x: ML + 3.15, w: 1.05, align: 'center' },
      { text: r.scope, x: ML + 4.35, w: 5.25 },
      { text: fmt(r.fee), x: ML + 9.75, w: 2.3, align: 'right' },
    ], y, 0.52, i % 2 === 0 ? C.rowAlt : C.offwhite);
  });

  // Subtotal row
  const subtotalY = hdrY + 0.35 + rows.length * 0.56;
  s.addShape('rect', { x: ML, y: subtotalY, w: CW, h: 0.52, fill: { color: C.rowTotal }, line: { color: C.border, width: 0.5 } });
  s.addText('Add-ons subtotal', { x: ML + 0.1, y: subtotalY + 0.1, w: 9.5, h: 0.32, fontSize: SZ.body, fontFace: 'SVN-Gilroy Bold', color: C.terracottaDark, bold: true, margin: 0 });
  s.addText(fmt(feeAddons), { x: ML + 9.75, y: subtotalY + 0.1, w: 2.3, h: 0.32, fontSize: SZ.body, fontFace: 'SVN-Gilroy Bold', color: C.terracottaDark, bold: true, align: 'right', margin: 0 });

  // Grand total dark bar
  const grandY = subtotalY + 0.62;
  s.addShape('rect', { x: ML, y: grandY, w: CW, h: 0.65, fill: { color: C.black }, line: { color: C.black, width: 0 } });
  s.addText('GRAND TOTAL', { x: ML + 0.15, y: grandY + 0.08, w: 5, h: 0.28, fontSize: 10, fontFace: F.sans, color: C.white, bold: true, charSpacing: 2, margin: 0 });
  s.addText(`Interior + Lighting (${fmt(feeMain)}) + Add-ons concept stage (${fmt(feeAddons)}). Excludes taxes & other fees.`, {
    x: ML + 0.15, y: grandY + 0.36, w: 8, h: 0.22, fontSize: 7.5, fontFace: F.sans, color: C.gray, margin: 0,
  });
  s.addText(fmt(feeGrand) + ' USD', {
    x: ML + 8.2, y: grandY + 0.08, w: 3.9, h: 0.42, fontSize: 16, fontFace: F.serif, color: C.terracotta, align: 'right', margin: 0,
  });

  // Discount strip
  s.addShape('rect', { x: ML, y: grandY + 0.72, w: CW, h: 0.38, fill: { color: 'EDE8E2' }, line: { color: 'D6C8B8', width: 0.5 } });
  s.addText(`Early-commitment partnership discount −10% if signed within 14 days  →  ${fmt(feeDiscount)} USD`, {
    x: ML + 0.15, y: grandY + 0.76, w: CW - 0.3, h: 0.26, fontSize: SZ.body, fontFace: F.sans, color: C.terracottaDark, italic: true, margin: 0,
  });

  pageNum(s, 11);
  copyright(s);
}

// ─── SLIDE 12 — PAYMENT SCHEDULE ────────────────────────────────
function slide12(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Payment Schedule', 'Milestone-based, aligned to stage approvals');

  const scope = d.scope || 'A';
  const fee = parseFloat(d.feeTotal) || 0;
  const fmt = n => n > 0 ? n.toLocaleString('en-US') : '[—]';

  const pmtA = [
    { num: '1', milestone: 'Upon signing — mobilization', pct: '20%', amt: Math.round(fee * 0.20) },
    { num: '2', milestone: 'Concept Design approval', pct: '45%', amt: Math.round(fee * 0.45) },
    { num: '3', milestone: 'Schematic Design approval', pct: '35%', amt: Math.round(fee * 0.35) },
  ];

  const pmtB = [
    { num: '1', milestone: 'Upon signing — mobilization', pct: '10%', amt: Math.round(fee * 0.10) },
    { num: '2', milestone: 'Concept Design approval', pct: '20%', amt: Math.round(fee * 0.20) },
    { num: '3', milestone: 'Schematic Design approval', pct: '20%', amt: Math.round(fee * 0.20) },
    { num: '4', milestone: 'Design Development approval', pct: '20%', amt: Math.round(fee * 0.20) },
    { num: '5', milestone: 'Construction Documents delivery', pct: '20%', amt: Math.round(fee * 0.20) },
    { num: '6', milestone: 'Project completion / Final supervision', pct: '10%', amt: Math.round(fee * 0.10) },
  ];

  const pmts = scope === 'B' ? pmtB : pmtA;

  const hdrY = 1.48;
  tableHeaderRow(s, [
    { text: '#', x: ML + 0.06, w: 0.7, align: 'center' },
    { text: 'MILESTONE', x: ML + 0.85, w: 7.8 },
    { text: '%', x: ML + 8.75, w: 1.0, align: 'center' },
    { text: 'INTERIOR + LIGHTING (USD)', x: ML + 9.85, w: 2.25, align: 'right' },
  ], hdrY, 0.35);

  pmts.forEach((p, i) => {
    const y = hdrY + 0.35 + i * 0.62;
    tableRow(s, [
      { text: p.num, x: ML + 0.1, w: 0.6, align: 'center' },
      { text: p.milestone, x: ML + 0.85, w: 7.75 },
      { text: p.pct, x: ML + 8.75, w: 0.95, align: 'center' },
      { text: fmt(p.amt), x: ML + 9.85, w: 2.2, align: 'right' },
    ], y, 0.58, i % 2 === 0 ? C.rowAlt : C.offwhite);
  });

  // Total row
  const totY = hdrY + 0.35 + pmts.length * 0.62;
  s.addShape('rect', { x: ML, y: totY, w: CW, h: 0.52, fill: { color: C.rowTotal }, line: { color: C.border, width: 0.5 } });
  s.addText('Total', { x: ML + 0.85, y: totY + 0.1, w: 7.75, h: 0.32, fontSize: SZ.body, fontFace: 'SVN-Gilroy Bold', color: C.terracottaDark, bold: true, margin: 0 });
  s.addText('100%', { x: ML + 8.75, y: totY + 0.1, w: 0.95, h: 0.32, fontSize: SZ.body, fontFace: 'SVN-Gilroy Bold', color: C.terracottaDark, bold: true, align: 'center', margin: 0 });
  s.addText(fmt(fee), { x: ML + 9.85, y: totY + 0.1, w: 2.2, h: 0.32, fontSize: SZ.body, fontFace: 'SVN-Gilroy Bold', color: C.terracottaDark, bold: true, align: 'right', margin: 0 });

  s.addText('Add-on workstreams (Brand, Signage, AV — concept) are billed 50% on workstream kick-off and 50% on delivery, or may be folded into the schedule above. Invoices in USD; payment within 14 days of invoice.', {
    x: ML, y: totY + 0.62, w: CW, h: 0.38,
    fontSize: 8, fontFace: F.sans, color: C.gray, margin: 0,
  });

  pageNum(s, 12);
  copyright(s);
}

// ─── SLIDE 13 — INCLUSIONS & EXCLUSIONS ──────────────────────────
function slide13(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Inclusions & Exclusions', null);

  const inc = [
    'Interior design — Concept & Schematic stages',
    'Lighting Design — concept & schematic (integrated, in the design fee)',
    'Zoning & space planning, design narrative, moodboards',
    '3D key-view renders · material & FF&E palette',
    'Schematic plans, RCP, elevations & outline specifications',
    'Brand, Signage & AV — concept stage (1 direction each) if selected',
  ];
  const exc = [
    'Design Development, Construction Documents & BOQ (follow-on addendum)',
    'Development & production of branding, signage & AV beyond concept',
    'Base-building architecture; MEP, structural & façade engineering',
    'Kitchen consultant, acoustics & fire engineering',
    'Permits & approvals · taxes/VAT',
    'Reimbursables: travel, printing, mock-ups (at cost with prior approval)',
    'Construction & supervision (unless contracted under Full Services)',
  ];

  [
    { label: 'INCLUSIONS', items: inc, x: ML, bg: C.rowAlt, color: C.black },
    { label: 'EXCLUSIONS', items: exc, x: ML + 6.28, bg: C.offwhite, color: C.grayMid },
  ].forEach(col => {
    s.addShape('rect', {
      x: col.x, y: 1.48, w: 5.9, h: 5.6,
      fill: { color: col.bg }, line: { color: C.border, width: 0.5 },
    });
    sectionLabel(s, col.label, col.x + 0.2, 1.62, 5.7, false);
    const items = col.items.map((l, idx) => ({
      text: l,
      options: { bullet: true, breakLine: idx < col.items.length - 1, fontSize: SZ.body, fontFace: F.sans, color: col.color },
    }));
    s.addText(items, { x: col.x + 0.2, y: 1.98, w: 5.65, h: 4.8, margin: 0, valign: 'top' });
  });

  pageNum(s, 13);
  copyright(s);
}

// ─── SLIDE 14 — TERMS & CONDITIONS ───────────────────────────────
function slide14(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Terms & Conditions', null);

  const terms = [
    { h: 'Revisions', b: 'Each stage includes a First Draft + one consolidated review round; further rounds by addendum (time + fee).' },
    { h: 'Site visits', b: 'Up to three (3) included in the design fee; additional visits charged at cost.' },
    { h: 'Reimbursables', b: 'Travel, printing & physical mock-ups billed at cost with prior approval.' },
    { h: 'Intellectual property', b: 'Design IP transfers to the Client upon full settlement of fees.' },
    { h: 'Validity', b: `This proposal is valid for 30 days from the date of issue.` },
    { h: 'Currency & tax', b: 'All fees in USD, exclusive of VAT and other statutory charges.' },
    { h: 'Schedule', b: 'Timeline assumes timely Client feedback at each approval gate.' },
  ];

  terms.forEach((t, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = ML + col * 6.28;
    const y = 1.55 + row * 1.05;
    s.addShape('rect', {
      x, y, w: 6.0, h: 0.95,
      fill: { color: col === 0 ? C.rowAlt : C.offwhite },
      line: { color: C.border, width: 0.5 },
    });
    s.addText(t.h, { x: x + 0.18, y: y + 0.1, w: 5.7, h: 0.26, fontSize: SZ.body, fontFace: 'SVN-Gilroy Bold', color: C.terracottaDark, bold: true, margin: 0 });
    s.addText(t.b, { x: x + 0.18, y: y + 0.38, w: 5.7, h: 0.5, fontSize: SZ.body, fontFace: F.sans, color: C.black, margin: 0, valign: 'top' });
  });

  pageNum(s, 14);
  copyright(s);
}

// ─── SLIDE 15 — THANK YOU ────────────────────────────────────────
function slide15(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  logo(s, false);
  s.addText('H2X Studio', {
    x: W - 5.0, y: 0.289, w: 4.666, h: 0.5,
    fontSize: 22, fontFace: F.serif, color: C.black, align: 'right', margin: 0,
  });

  // Left image
  s.addShape('rect', { x: 0, y: 0, w: 4.6, h: H, fill: { color: C.cream }, line: { color: C.grayLight, width: 1 } });
  s.addText('[ CLOSING IMAGE ]', { x: 0, y: 0, w: 4.6, h: H, fontSize: 9, fontFace: F.sans, color: C.gray, align: 'center', valign: 'middle', margin: 0 });

  s.addText('Thank You', {
    x: 5.2, y: 1.8, w: 7.5, h: 1.1,
    fontSize: 42, fontFace: F.serif, color: C.terracotta, margin: 0,
  });

  s.addText(d.ctaText || 'We would welcome a working session to walk through the zoning & layout options, align on scope & budget, and confirm the schedule — so we can move into Concept without delay.', {
    x: 5.2, y: 3.0, w: 7.6, h: 0.9,
    fontSize: SZ.body, fontFace: F.sans, color: C.black, margin: 0, valign: 'top',
  });

  const contacts = [
    d.phone || '+84 9 6652 6662',
    `${d.website || 'www.h2xstudio.com.vn'} · ${d.email || 'info@h2xstudio.com'}`,
    d.addressHN || 'HQ — 26 Trần Hưng Đạo, Hà Nội',
    d.addressHCM || 'Branch — 199D Nguyễn Văn Hưởng, Thảo Điền, TP HCM',
  ];

  contacts.forEach((c, i) => {
    const y = 4.1 + i * 0.48;
    if (i > 0) s.addShape('rect', { x: 5.2, y: y - 0.06, w: 7.6, h: 0.008, fill: { color: C.rule }, line: { color: C.rule, width: 0 } });
    s.addText(c, {
      x: 5.2, y, w: 7.6, h: 0.36,
      fontSize: i === 0 ? 11 : SZ.body, fontFace: i === 0 ? 'SVN-Gilroy Bold' : F.sans,
      color: i === 0 ? C.black : C.gray, bold: i === 0, margin: 0,
    });
  });

  copyright(s, false);
}

// ─── MAIN EXPORT ─────────────────────────────────────────────────
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
  slide09(pres, data);

  // Only add add-on slides if any add-on is selected
  if (data.addBrand || data.addSignage || data.addAV) {
    slide10(pres, data);
    slide11(pres, data);
  }

  slide12(pres, data);
  slide13(pres, data);
  slide14(pres, data);
  slide15(pres, data);

  const fileName = `/tmp/H2X_FeeProposal_${(data.projectName || 'Project').replace(/\s+/g,'_')}_v${data.version || '01'}.pptx`;
  await pres.writeFile({ fileName });
  return fileName;
}

module.exports = { generateProposal };
