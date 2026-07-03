// H2X Studio — Fee Proposal Generator
// Layout: LAYOUT_WIDE 13.333" x 7.5" (extracted from Taseco VI PPTX)
// Fonts: Toma Sans Light (body) + Toma Sans Bold (emphasis) + Playfair Display (headings)

const pptxgen = require('pptxgenjs');
const FONTS = require('./fontData');

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

// ─── HELPERS ─────────────────────────────────────────────────────
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

// Logo — text-based (H2X STUDIO)
function logo(s, light = false) {
  const col = light ? C.white : C.black;
  const sub = light ? C.terra : C.terra;
  s.addText('H2X', { x: ML, y: 0.289, w: 0.7, h: 0.32,
    fontSize: 18, fontFace: F.serifI, color: col, bold: false, margin: 0 });
  s.addText('STUDIO', { x: ML, y: 0.59, w: 1.0, h: 0.2,
    fontSize: 7, fontFace: F.light, color: sub, charSpacing: 3, margin: 0 });
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
  logo(s);
  studioName(s);

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
      fontSize: SZ.label, fontFace: F.light, color: C.terra, margin: 0,
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
    margin: 0, valign: 'top',
  });
}

// Table header row — black bg, white Toma Sans Light spaced caps
function tblHeader(s, cells, y, rowH = 0.35) {
  s.addShape('rect', {
    x: ML, y, w: CW, h: rowH,
    fill: { color: C.black }, line: { color: C.black, width: 0 },
  });
  cells.forEach(cell => {
    s.addText(cell.text, {
      x: cell.x, y: y + 0.05, w: cell.w, h: rowH - 0.08,
      fontSize: SZ.label, fontFace: F.light, color: C.white,
      charSpacing: 1.5, margin: 0, valign: 'middle',
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
  s.addShape('rect', {
    x: 2.0, y: 0.6, w: 9.333, h: 4.3,
    fill: { color: C.cream }, line: { color: C.graylt, width: 0.5 },
  });
  s.addText('[ HERO IMAGE ]', {
    x: 2.0, y: 0.6, w: 9.333, h: 4.3,
    fontSize: 9, fontFace: F.light, color: C.gray,
    align: 'center', valign: 'middle', margin: 0,
  });

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
    margin: 0, valign: 'top',
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
  s.addText(briefLines, { x: lx, y: 4.983, w: 5.9439, h: 1.8536, margin: 0, valign: 'top' });

  // EXPERIENCE INTENT
  secLabel(s, 'ĐỊNH HƯỚNG TRẢI NGHIỆM', rx, 1.6933, rw);
  s.addText(d.experienceIntent || '[Định hướng trải nghiệm]', {
    x: rx, y: 1.9182, w: rw, h: 2.7377,
    fontSize: SZ.body, fontFace: F.light, color: C.black,
    margin: 0, valign: 'top',
  });

  // WHY H2X
  secLabel(s, 'VÌ SAO H2X', rx, 4.7074, rw);
  s.addText(d.whyH2x || 'H2X (Human + Hospitality × eXperiences) thiết kế dựa trên vận hành và hành trình trải nghiệm của khách hàng.', {
    x: rx, y: 4.983, w: 5.2931, h: 1.3215,
    fontSize: SZ.body, fontFace: F.light, color: C.black,
    margin: 0, valign: 'top',
  });

  pageNum(s, 2);
  copyright(s);
}

// ─── SLIDE 03 — SCOPE OF SERVICES ───────────────────────────────
function slide03(pres, d) {
  const s = pres.addSlide();
  const scope = d.scope || 'A';
  slideTitle(s, 'Scope of Services',
    scope === 'A'
      ? 'This engagement covers Concept & Schematic Design, with Lighting Design integrated throughout'
      : 'Full-service engagement — Concept through Construction Documents & Site Supervision'
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
    { stage: 'Schematic Design + Lighting Schematic', status: '✓ Included', desc: 'Annotated plans, RCP & floor-finish plans, key elevations, outline spec schedules, schematic lighting.' },
    { stage: 'Design Development', status: '✓ Included', desc: 'Detailed drawings, material schedules, FF&E specifications.' },
    { stage: 'Construction Documents + BOQ', status: '✓ Included', desc: 'Full CD set, tender-ready bill of quantities.' },
    { stage: 'Tender Support · Site Supervision', status: '✓ Included', desc: 'Bid review, RFI responses, site visits through to completion.' },
  ];
  const stages = scope === 'B' ? stagesB : stagesA;

  // Header — exact positions from Taseco slide 3
  tblHeader(s, [
    { text: 'STAGE IN SCOPE', x: ML + 0.06, w: 3.5 },
    { text: 'WHAT IT DELIVERS', x: ML + 5.0, w: 7.2 },
  ], 1.55, 0.35);

  stages.forEach((st, i) => {
    const isOpt = st.status === 'Optional';
    const y = 1.55 + 0.35 + i * 0.95;
    const bg = i % 2 === 0 ? C.rowalt : C.white;
    tblRow(s, [
      { text: st.stage, x: ML + 0.1, w: 4.0, bold: !isOpt, italic: isOpt, color: isOpt ? C.gray : C.black },
      { text: isOpt ? 'Optional' : '✓ Included', x: ML + 4.2, w: 0.9, color: isOpt ? C.gray : C.terradk, bold: !isOpt, align: 'center', fontSize: 8 },
      { text: st.desc, x: ML + 5.2, w: 6.9, color: isOpt ? C.gray : C.black },
    ], y, 0.88, bg);
  });

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
  logo(s);
  studioName(s);

  s.addText('Design Language', {
    x: ML, y: 0.5, w: CW, h: 0.55,
    fontSize: SZ.h1, fontFace: F.serif, color: C.black, margin: 0,
  });
  s.addText(d.designTagline || 'Quiet luxury with a Vietnamese soul — warm materials, light from shadow, restraint over display', {
    x: ML, y: 1.0842, w: CW, h: 0.2308,
    fontSize: SZ.label, fontFace: F.light, color: C.terra, margin: 0,
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
  pillars.forEach((p, i) => {
    const x = ML + i * (imgW + 0.18);
    s.addShape('rect', {
      x, y: imgY, w: imgW, h: imgH,
      fill: { color: C.cream }, line: { color: C.graylt, width: 0.5 },
    });
    s.addText('[ MOOD IMAGE ]', {
      x, y: imgY, w: imgW, h: imgH,
      fontSize: 8, fontFace: F.light, color: C.gray,
      align: 'center', valign: 'middle', margin: 0,
    });
    s.addText(p.label || '', {
      x, y: imgY + imgH + 0.08, w: imgW, h: 0.25,
      fontSize: SZ.label, fontFace: F.bold, color: C.black,
      charSpacing: 2, margin: 0,
    });
    s.addText(p.sub || '', {
      x, y: imgY + imgH + 0.35, w: imgW, h: 0.22,
      fontSize: 8, fontFace: F.light, color: C.gray,
      charSpacing: 1, margin: 0,
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
        d.lightingStrategy || 'Zoned atmosphere strategy — dramatic for F&B, functional for business, circadian/calming for Quiet & Wellness',
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
      ], { x: col.x, y: 2.6 + i * 1.05, w: col.w, h: 0.9, margin: 0, valign: 'top' });
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
    { role: 'CREATIVE DIRECTOR', name: 'Gia Huy (Michael)', title: 'M.Architect · 23+ yrs · sets vision & design direction', dark: true },
    { role: 'LEAD DESIGNER — CONCEPT', name: 'Nguyễn Mạnh Hùng (Henry)', title: 'Leads spatial concept, narrative & guest journey' },
    { role: 'LIGHTING DESIGNER (IN-HOUSE)', name: '[Lighting Lead]', title: 'Atmosphere, fixture & control strategy from Concept' },
    { role: 'HEAD OF FF&E', name: 'Nguyễn Thị Thúy Vy (Vivian)', title: 'Furniture, fixtures & material curation' },
    { role: 'TECHNICAL LEAD', name: 'Nguyễn Vương Linh (Lucas)', title: 'Schematic documentation & consultant coordination' },
    { role: '3D VISUALISATION & PM', name: 'Đỗ Danh Sơn', title: 'Photorealistic renders · schedule, reviews & delivery' },
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
    ? ['Briefing &\nMobilization', 'Concept Design\n+ Lighting', 'Schematic Design\n+ Lighting', 'Design\nDevelopment', 'Construction\nDocuments']
    : ['Briefing &\nMobilization', 'Concept Design\n+ Lighting', 'Schematic Design\n+ Lighting'];

  const bw = scope === 'B' ? 2.28 : 3.85, gap = scope === 'B' ? 0.18 : 0.28;
  stages.forEach((st, i) => {
    const x = ML + i * (bw + gap);
    const dark = i === stages.length - 1;
    s.addShape('rect', {
      x, y: 1.55, w: bw, h: 0.95,
      fill: { color: dark ? C.black : C.white },
      line: { color: dark ? C.black : C.graylt, width: 1 },
    });
    s.addText(st, {
      x: x + 0.1, y: 1.62, w: bw - 0.2, h: 0.78,
      fontSize: 11, fontFace: F.serif, color: dark ? C.white : C.black,
      align: 'center', margin: 0,
    });
    if (i < stages.length - 1) {
      s.addText('→', {
        x: x + bw + 0.02, y: 1.85, w: gap - 0.04, h: 0.35,
        fontSize: 13, fontFace: F.light, color: C.gray, align: 'center', margin: 0,
      });
    }
  });

  const infos = [
    { h: 'Cadence', b: 'Each stage: a First Draft, one consolidated Review round, then Final Approval & a stage-completion sign-off.\n\nAdditional rounds by addendum (time + fee).' },
    { h: 'Tools', b: 'Revit + Enscape; photorealistic 3D key views from Concept so decisions are made on what you will actually see.' },
    { h: 'Indicative Timeline', b: scope === 'B'
        ? 'Mobilization ~1 wk · Concept ~4 wks · Schematic ~4 wks · DD ~6 wks · CD/BOQ ~8 wks + approval gates.'
        : 'Mobilization ~1 wk · Concept ~4 wks · Schematic ~4 wks + client-approval gates — about 9–11 weeks.' },
  ];
  infos.forEach((inf, i) => {
    const x = ML + i * (3.85 + 0.28);
    s.addShape('rect', {
      x, y: 2.65, w: 3.85, h: 2.5,
      fill: { color: C.rowalt }, line: { color: C.border, width: 0.5 },
    });
    s.addText(inf.h, {
      x: x + 0.18, y: 2.78, w: 3.55, h: 0.35,
      fontSize: 11, fontFace: F.serif, color: C.black, margin: 0,
    });
    s.addText(inf.b, {
      x: x + 0.18, y: 3.16, w: 3.55, h: 1.85,
      fontSize: SZ.body, fontFace: F.light, color: C.black,
      margin: 0, valign: 'top',
    });
  });

  s.addText(`Methodology: Brief → Concept (approval) → Schematic (approval). ${scope === 'B' ? 'Design Development → Construction Documents → BOQ → Supervision follow the same gated rhythm.' : 'Should the project continue, Design Development → Construction Documents → BOQ follow the same gated rhythm.'}`, {
    x: ML, y: H - 0.6, w: CW, h: 0.22,
    fontSize: 7.5, fontFace: F.light, color: C.gray, margin: 0,
  });

  pageNum(s, 8);
  copyright(s);
}

// ─── SLIDE 09 — DESIGN FEE ───────────────────────────────────────
function slide09(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Design Fee', 'Interior Design + Lighting Design · Concept → Schematic · lump sum by phase');

  const scope = d.scope || 'A';
  const fee = parseFloat(d.feeTotal) || 0;
  const fmt = n => n > 0 ? n.toLocaleString('en-US') : '[—]';

  const stagesA = [
    { num: '1', stage: 'Briefing & Mobilization', ratio: '15%', amt: Math.round(fee * 0.15) },
    { num: '2', stage: 'Concept Design (incl. Lighting Concept)', ratio: '50%', amt: Math.round(fee * 0.50) },
    { num: '3', stage: 'Schematic Design (incl. Lighting Schematic)', ratio: '35%', amt: Math.round(fee * 0.35) },
  ];
  const stagesB = [
    { num: '1', stage: 'Briefing & Mobilization', ratio: '5%', amt: Math.round(fee * 0.05) },
    { num: '2', stage: 'Concept Design + Lighting', ratio: '20%', amt: Math.round(fee * 0.20) },
    { num: '3', stage: 'Schematic Design + Lighting', ratio: '20%', amt: Math.round(fee * 0.20) },
    { num: '4', stage: 'Design Development', ratio: '25%', amt: Math.round(fee * 0.25) },
    { num: '5', stage: 'Construction Documents + BOQ', ratio: '20%', amt: Math.round(fee * 0.20) },
    { num: '6', stage: 'Tender Support + Site Supervision', ratio: '10%', amt: Math.round(fee * 0.10) },
  ];
  const stages = scope === 'B' ? stagesB : stagesA;

  tblHeader(s, [
    { text: 'STAGE', x: ML + 0.06, w: 0.6 },
    { text: 'PHASE', x: ML + 0.75, w: 7.5 },
    { text: 'RATIO', x: ML + 8.35, w: 1.2 },
    { text: 'FEE (USD)', x: ML + 9.65, w: 2.45, align: 'right' },
  ], 1.55, 0.35);

  stages.forEach((st, i) => {
    const y = 1.55 + 0.35 + i * 0.7;
    tblRow(s, [
      { text: st.num, x: ML + 0.1, w: 0.55, align: 'center' },
      { text: st.stage, x: ML + 0.75, w: 7.45 },
      { text: st.ratio, x: ML + 8.35, w: 1.15, align: 'center' },
      { text: fmt(st.amt), x: ML + 9.65, w: 2.4, align: 'right' },
    ], y, 0.65, i % 2 === 0 ? C.rowalt : C.white);
  });

  // Subtotal
  const totY = 1.55 + 0.35 + stages.length * 0.7;
  tblRow(s, [
    { text: 'INTERIOR + LIGHTING — SUBTOTAL', x: ML + 0.75, w: 8.8, bold: true, color: C.terradk },
    { text: '100%', x: ML + 8.35, w: 1.15, align: 'center', bold: true, color: C.terradk },
    { text: fmt(fee), x: ML + 9.65, w: 2.4, align: 'right', bold: true, color: C.terradk },
  ], totY, 0.55, C.rowtot);

  // Benchmark note
  const area = parseFloat(d.area) || 0;
  const perSqm = area > 0 ? Math.round(fee / area) : 0;
  s.addText(`≈ ${perSqm > 0 ? perSqm + ' USD / m²' : '[USD/m²]'} · ${d.area || '[X]'} m² · indicative benchmark for this typology (Lighting included). Excludes taxes & other fees.`, {
    x: ML, y: totY + 0.65, w: CW, h: 0.3,
    fontSize: 7.5, fontFace: F.light, color: C.gray, italic: true, margin: 0,
  });
  s.addText('Design Development / Construction Documents / BOQ available as a follow-on addendum.', {
    x: ML, y: totY + 0.95, w: CW, h: 0.25,
    fontSize: 7.5, fontFace: F.light, color: C.gray, margin: 0,
  });

  pageNum(s, 9);
  copyright(s);
}

// ─── SLIDE 10 — ADD-ON VALUE ─────────────────────────────────────
function slide10(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, 'Branding · Signage · AV — Scope & Value', null);
  s.addText('Three workstreams that turn a beautiful space into a coherent, memorable destination — offered at concept stage', {
    x: ML, y: 1.0842, w: CW, h: 0.2308,
    fontSize: SZ.label, fontFace: F.light, color: C.terra, margin: 0,
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

  pageNum(s, 10);
  copyright(s);
}

// ─── SLIDE 11 — ADD-ON FEE ───────────────────────────────────────
function slide11(pres, d) {
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
  ].filter(r => r.show);

  tblHeader(s, [
    { text: 'WORKSTREAM', x: ML + 0.06, w: 3.2 },
    { text: 'PARTNER', x: ML + 3.35, w: 1.1 },
    { text: 'SCOPE (CONCEPT, 1 DIRECTION)', x: ML + 4.55, w: 5.5 },
    { text: 'FEE (USD)', x: ML + 10.15, w: 1.95, align: 'right' },
  ], 1.55, 0.35);

  rows.forEach((r, i) => {
    const y = 1.55 + 0.35 + i * 0.65;
    tblRow(s, [
      { text: r.ws, x: ML + 0.1, w: 3.1 },
      { text: r.partner, x: ML + 3.35, w: 1.05, align: 'center' },
      { text: r.scope, x: ML + 4.55, w: 5.45 },
      { text: fmt(r.fee), x: ML + 10.15, w: 1.9, align: 'right' },
    ], y, 0.6, i % 2 === 0 ? C.rowalt : C.white);
  });

  const subY = 1.55 + 0.35 + rows.length * 0.65;
  tblRow(s, [
    { text: 'Add-ons subtotal', x: ML + 0.1, w: 9.9, bold: true, color: C.terradk },
    { text: fmt(addons), x: ML + 10.15, w: 1.9, align: 'right', bold: true, color: C.terradk },
  ], subY, 0.55, C.rowtot);

  // Grand total
  const gtY = subY + 0.65;
  s.addShape('rect', { x: ML, y: gtY, w: CW, h: 0.7, fill: { color: C.black }, line: { color: C.black, width: 0 } });
  s.addText('GRAND TOTAL', { x: ML + 0.15, y: gtY + 0.08, w: 6, h: 0.32, fontSize: 10, fontFace: F.light, color: C.white, charSpacing: 2, margin: 0 });
  s.addText(`Interior + Lighting (${fmt(feeMain)}) + Add-ons (${fmt(addons)}). Excludes taxes.`, {
    x: ML + 0.15, y: gtY + 0.42, w: 8, h: 0.22, fontSize: 7.5, fontFace: F.light, color: C.gray, margin: 0,
  });
  s.addText(fmt(grand) + ' USD', {
    x: ML + 8.2, y: gtY + 0.08, w: 3.9, h: 0.45, fontSize: 16, fontFace: F.serif, color: C.terra, align: 'right', margin: 0,
  });

  // Discount
  s.addShape('rect', { x: ML, y: gtY + 0.78, w: CW, h: 0.35, fill: { color: 'EDE8E2' }, line: { color: 'D6C8B8', width: 0.5 } });
  s.addText(`Early-commitment partnership discount −10% if signed within 14 days  →  ${fmt(Math.round(grand * 0.9))} USD`, {
    x: ML + 0.15, y: gtY + 0.82, w: CW - 0.3, h: 0.25, fontSize: SZ.body, fontFace: F.light, color: C.terradk, italic: true, margin: 0,
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
    { num: '1', m: 'Upon signing — mobilization', pct: '20%', amt: Math.round(fee * 0.20) },
    { num: '2', m: 'Concept Design approval', pct: '45%', amt: Math.round(fee * 0.45) },
    { num: '3', m: 'Schematic Design approval', pct: '35%', amt: Math.round(fee * 0.35) },
  ];
  const pmtB = [
    { num: '1', m: 'Upon signing — mobilization', pct: '10%', amt: Math.round(fee * 0.10) },
    { num: '2', m: 'Concept Design approval', pct: '20%', amt: Math.round(fee * 0.20) },
    { num: '3', m: 'Schematic Design approval', pct: '20%', amt: Math.round(fee * 0.20) },
    { num: '4', m: 'Design Development approval', pct: '20%', amt: Math.round(fee * 0.20) },
    { num: '5', m: 'Construction Documents delivery', pct: '20%', amt: Math.round(fee * 0.20) },
    { num: '6', m: 'Project completion / Final supervision', pct: '10%', amt: Math.round(fee * 0.10) },
  ];
  const pmts = scope === 'B' ? pmtB : pmtA;

  tblHeader(s, [
    { text: '#', x: ML + 0.06, w: 0.6, align: 'center' },
    { text: 'MILESTONE', x: ML + 0.75, w: 8.2 },
    { text: '%', x: ML + 9.05, w: 1.0, align: 'center' },
    { text: 'INTERIOR + LIGHTING (USD)', x: ML + 10.15, w: 1.95, align: 'right' },
  ], 1.55, 0.35);

  pmts.forEach((p, i) => {
    const y = 1.55 + 0.35 + i * 0.72;
    tblRow(s, [
      { text: p.num, x: ML + 0.1, w: 0.55, align: 'center' },
      { text: p.m, x: ML + 0.75, w: 8.15 },
      { text: p.pct, x: ML + 9.05, w: 0.95, align: 'center' },
      { text: fmt(p.amt), x: ML + 10.15, w: 1.9, align: 'right' },
    ], y, 0.67, i % 2 === 0 ? C.rowalt : C.white);
  });

  const totY = 1.55 + 0.35 + pmts.length * 0.72;
  tblRow(s, [
    { text: 'Total', x: ML + 0.75, w: 8.15, bold: true, color: C.terradk },
    { text: '100%', x: ML + 9.05, w: 0.95, align: 'center', bold: true, color: C.terradk },
    { text: fmt(fee), x: ML + 10.15, w: 1.9, align: 'right', bold: true, color: C.terradk },
  ], totY, 0.55, C.rowtot);

  s.addText('Add-on workstreams (Brand, Signage, AV) billed 50% on kick-off and 50% on delivery, or may be folded into the schedule above. Invoices in USD; payment within 14 days of invoice.', {
    x: ML, y: totY + 0.65, w: CW, h: 0.38,
    fontSize: 8, fontFace: F.light, color: C.gray, margin: 0,
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
    { label: 'INCLUSIONS', items: inc, x: ML, bg: C.rowalt, color: C.black },
    { label: 'EXCLUSIONS', items: exc, x: ML + 6.18, bg: C.white, color: C.gray },
  ].forEach(col => {
    s.addShape('rect', { x: col.x, y: 1.55, w: 5.9, h: 5.6, fill: { color: col.bg }, line: { color: C.border, width: 0.5 } });
    secLabel(s, col.label, col.x + 0.2, 1.68, 5.6);
    s.addText(col.items.map(l => `+ ${l}`).join('\n'), {
      x: col.x + 0.2, y: 1.98, w: 5.65, h: 4.98,
      fontSize: SZ.body, fontFace: F.light, color: col.color,
      margin: 0, valign: 'top',
    });
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
    s.addText(t.b, { x: x + 0.18, y: y + 0.38, w: 5.6, h: 0.55, fontSize: SZ.body, fontFace: F.light, color: C.black, margin: 0, valign: 'top' });
  });

  pageNum(s, 14);
  copyright(s);
}

// ─── SLIDE 15 — THANK YOU ────────────────────────────────────────
function slide15(pres, d) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  studioName(s);

  // Left image box (from Taseco: x=0, y=0, w≈4.6, h=7.5)
  s.addShape('rect', { x: 0, y: 0, w: 4.5, h: H, fill: { color: C.cream }, line: { color: C.graylt, width: 0.5 } });
  s.addText('[ CLOSING IMAGE ]', { x: 0, y: 0, w: 4.5, h: H, fontSize: 8, fontFace: F.light, color: C.gray, align: 'center', valign: 'middle', margin: 0 });

  // Thank You
  s.addText('Thank You', {
    x: 5.2, y: 1.6, w: 7.7, h: 1.0,
    fontSize: 42, fontFace: F.serif, color: C.terra, margin: 0,
  });

  // CTA
  s.addText(d.ctaText || 'We would welcome a working session to walk through the zoning & layout options, align on scope & budget, and confirm the schedule — so we can move into Concept without delay.', {
    x: 5.2, y: 2.78, w: 7.7, h: 0.85,
    fontSize: SZ.body, fontFace: F.light, color: C.black, margin: 0, valign: 'top',
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
