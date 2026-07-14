// H2X Studio — Fee Proposal Generator
// Layout: LAYOUT_WIDE 13.333" x 7.5" (extracted from Taseco VI PPTX)
// Fonts: Toma Sans Light (body) + Toma Sans Bold (emphasis) + Playfair Display (headings)

const pptxgen = require('pptxgenjs');
const FONTS = require('./fontData');
const { H2X_LOGO } = require('./logoData');
const os = require('os');
const path = require('path');

// ─── NGÔN NGỮ (EN mặc định / VI khi data.language==='vi') ─────────
let CURRENT_LANG = 'en';
function setLang(lang) { CURRENT_LANG = (lang === 'vi') ? 'vi' : 'en'; }
function L(en, vi) { return CURRENT_LANG === 'vi' ? vi : en; }

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

// Tên các giai đoạn thiết kế — bản dịch chính thức của H2X. Khóa (key) LUÔN giữ tiếng Anh
// vì được dùng làm định danh khớp với front-end (data-phase) và các bảng dữ liệu khác.
const STAGE_VI = {
  'Appraisal and Briefing': 'Đánh giá và tóm tắt',
  'Feasibility & Preliminary Concept': 'Tính khả thi và ý tưởng sơ bộ',
  'Final Concept': 'Thiết kế Ý tưởng',
  'Final Concept + Lighting (Optional)': 'Thiết kế Ý tưởng + Chiếu sáng (Tùy chọn)',
  'Schematic Design': 'Thiết kế Cơ sở',
  'Design Development': 'Thiết kế kĩ thuật',
  'Construction Documents': 'Thiết kế bản vẽ thi công',
  'BOQ (Bill of Quantity)': 'Bảng khối lượng',
  'Author Supervision': 'Giám sát tác giả',
};
function stageLabel(stage) { return CURRENT_LANG === 'vi' ? (STAGE_VI[stage] || stage) : stage; }

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
    return CURRENT_LANG === 'vi'
      ? [stageLabel('Design Development'), stageLabel('Construction Documents'), stageLabel('BOQ (Bill of Quantity)'), stageLabel('Author Supervision')]
      : ['Design Development', 'Construction Documents', 'BOQ (Bill of Quantity)', 'Author Supervision'];
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
const STAGE_DESC_VI = {
  'Appraisal and Briefing': 'Khởi động dự án, thống nhất yêu cầu, khảo sát hiện trạng — nền tảng cho mọi bước tiếp theo.',
  'Feasibility & Preliminary Concept': 'Đánh giá khả thi & vận hành, mặt bằng điều chỉnh, định hướng phong cách & cảm xúc thiết kế.',
  'Final Concept + Lighting (Optional)': 'Tuyên ngôn ý tưởng, mặt bằng hoàn chỉnh, moodboard, phối cảnh 3D chính — ý tưởng chiếu sáng tùy chọn.',
  'Final Concept': 'Tuyên ngôn ý tưởng, mặt bằng hoàn chỉnh, moodboard, phối cảnh 3D chính, bảng vật liệu.',
  'Schematic Design': 'Bản vẽ mặt bằng có chú thích, mặt bằng trần & vật liệu sàn, mặt đứng chính, thuyết minh kỹ thuật sơ bộ.',
  'Design Development': 'Bản vẽ chi tiết, bảng vật liệu, thông số kỹ thuật FF&E.',
  'Construction Documents': 'Bộ hồ sơ thi công đầy đủ, phối hợp với các đơn vị tư vấn, sẵn sàng đấu thầu.',
  'BOQ (Bill of Quantity)': 'Bảng khối lượng sẵn sàng cho đấu thầu.',
  'Author Supervision': 'Rà soát hồ sơ dự thầu, phản hồi RFI, giám sát công trình tới khi hoàn thành.',
};
function stageDesc(stage) { return CURRENT_LANG === 'vi' ? (STAGE_DESC_VI[stage] || STAGE_DESC[stage] || '') : (STAGE_DESC[stage] || ''); }

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
  s.addText(L('©H2X.Studio 2026. All Rights Reserved', '©H2X.Studio 2026. Bảo lưu mọi quyền'), {
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
    s.addText(L('[ HERO IMAGE ]', '[ ẢNH BÌA ]'), {
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
  s.addText(L('INTERIOR DESIGN SERVICES', 'DỊCH VỤ THIẾT KẾ NỘI THẤT'), {
    x: ML, y: 5.96, w: CW, h: 0.3,
    fontSize: SZ.coverSub, fontFace: F.serif, color: C.black,
    align: 'center', margin: 0,
  });

  // Fee Proposal
  s.addText(L('Fee Proposal', 'Báo giá Thiết kế'), {
    x: ML, y: 6.3, w: CW, h: 0.3,
    fontSize: SZ.coverFee, fontFace: F.serif, color: C.black,
    align: 'center', bold: true, margin: 0,
  });

  // Version / Date
  s.addText(`${L('VERSION', 'PHIÊN BẢN')}  ${d.version || '01'}`, {
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
  slideTitle(s, L('Project Understanding', 'Hiểu biết Dự án'),
    `${d.clientName || '[Client]'} · ${d.typology || '[Typology]'} · ${d.location || '[Location]'}`);

  // Left column x=0.6667, w=5.7198
  // Right column x=7.1444, w=5.7198
  const lx = ML, lw = 5.7198;
  const rx = 7.1444, rw = 5.7198;

  // THE OPPORTUNITY
  secLabel(s, L('THE OPPORTUNITY', 'CƠ HỘI'), lx, 1.6, lw);
  s.addText(d.opportunity || L('[Project opportunity description]', '[Mô tả cơ hội dự án]'), {
    x: lx, y: 1.9398, w: lw + 0.43, h: 2.8363,
    fontSize: SZ.body, fontFace: F.light, color: C.black,
    margin: 0, valign: 'top', lineSpacingMultiple: 1.5,
  });

  // PROJECT BRIEF
  secLabel(s, L('PROJECT BRIEF', 'TÓM TẮT DỰ ÁN'), lx, 4.6957, lw);
  const briefLines = [
    { text: '+ ', options: { fontFace: F.light, color: C.terra } },
    { text: `${L('Client', 'Chủ đầu tư')} · ${d.clientName || '[Client]'}`, options: { fontFace: F.light, color: C.black } },
    { text: '\n+ ', options: { fontFace: F.light, color: C.terra } },
    { text: `${L('Location', 'Vị trí')} · ${d.location || '[Location]'}`, options: { fontFace: F.light, color: C.black } },
    { text: '\n+ ', options: { fontFace: F.light, color: C.terra } },
    { text: `${L('Typology', 'Loại hình')} · ${d.typology || '[Typology]'}`, options: { fontFace: F.light, color: C.black } },
    { text: '\n+ ', options: { fontFace: F.light, color: C.terra } },
    { text: `${L('Area', 'Diện tích')} · ~${d.area || '[X]'} m²`, options: { fontFace: F.light, color: C.black } },
    { text: '\n+ ', options: { fontFace: F.light, color: C.terra } },
    { text: `${L('Target guest', 'Khách hàng mục tiêu')} · ${d.targetGuest || '[Description]'}`, options: { fontFace: F.light, color: C.black } },
  ].map(l => ({ text: l.text, options: { ...l.options, fontSize: SZ.body } }));
  s.addText(briefLines, { x: lx, y: 4.983, w: 5.9439, h: 1.8536, margin: 0, valign: 'top', lineSpacingMultiple: 1.5 });

  // EXPERIENCE INTENT
  secLabel(s, L('EXPERIENCE INTENT', 'ĐỊNH HƯỚNG TRẢI NGHIỆM'), rx, 1.6933, rw);
  s.addText(d.experienceIntent || L('[Experience intent]', '[Định hướng trải nghiệm]'), {
    x: rx, y: 1.9182, w: rw, h: 2.7377,
    fontSize: SZ.body, fontFace: F.light, color: C.black,
    margin: 0, valign: 'top', lineSpacingMultiple: 1.5,
  });

  // WHY H2X
  secLabel(s, L('WHY H2X', 'VÌ SAO H2X'), rx, 4.7074, rw);
  s.addText(d.whyH2x || L('H2X (Human + Hospitality × eXperiences) designs based on operation and guest journey.', 'H2X (Human + Hospitality × eXperiences) thiết kế dựa trên vận hành và hành trình trải nghiệm của khách hàng.'), {
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
  const subtitlesEN = {
    A: 'This engagement covers Concept & Schematic Design — Lighting offered as an optional add-on within Concept',
    B: 'Full-service engagement — Concept through Construction Documents & Site Supervision · Lighting offered separately',
    C: 'Full-service engagement — Concept through Construction Documents & Site Supervision · Lighting bundled as optional add-on',
  };
  const subtitlesVI = {
    A: 'Hợp đồng bao gồm Thiết kế Ý tưởng & Thiết kế Cơ sở — Chiếu sáng là hạng mục bổ sung tùy chọn trong giai đoạn Ý tưởng',
    B: 'Hợp đồng trọn gói — Từ Ý tưởng đến Hồ sơ Thi công & Giám sát công trình · Chiếu sáng chào giá riêng',
    C: 'Hợp đồng trọn gói — Từ Ý tưởng đến Hồ sơ Thi công & Giám sát công trình · Chiếu sáng là hạng mục bổ sung tùy chọn',
  };
  const subtitles = CURRENT_LANG === 'vi' ? subtitlesVI : subtitlesEN;
  slideTitle(s, L('Scope of Services', 'Phạm vi Dịch vụ'), subtitles[scope] || subtitles.A);

  const descEN = {
    kickoff: 'Kick-off, brief alignment, A&B material review, site assessment.',
    feasibility: 'Feasibility & operational assessment, adjusted concept floor plan, mood & feeling recommendations.',
    finalConceptOpt: 'Concept statement, final layout, moodboard, 3D key views, material palette. Lighting concept offered as optional add-on.',
    schematic: 'Annotated plans, RCP & floor-finish plans, key elevations, outline specification schedules.',
    schematic2: 'Annotated plans, RCP & floor-finish plans, key elevations, outline spec schedules.',
    followOn: 'Available as a follow-on package — fees proposed by addendum at the agreed rate once Schematic is approved.',
    supervisionOpt: 'Offered separately on request once the project moves into construction.',
    finalConcept: 'Concept statement, final layout, moodboard, 3D key views, material palette.',
    devDesign: 'Detailed drawings, material schedules, FF&E specifications.',
    cd: 'Full CD set, coordinated with consultants.',
    boq: 'Tender-ready bill of quantities.',
    supervision: 'Bid review, RFI responses, site visits through to completion.',
  };
  const descVI = {
    kickoff: 'Khởi động, thống nhất yêu cầu, rà soát tài liệu Đánh giá & Tóm tắt, khảo sát hiện trạng.',
    feasibility: 'Đánh giá khả thi & vận hành, mặt bằng ý tưởng điều chỉnh, đề xuất phong cách & cảm xúc thiết kế.',
    finalConceptOpt: 'Tuyên ngôn ý tưởng, mặt bằng hoàn chỉnh, moodboard, phối cảnh 3D chính, bảng vật liệu. Ý tưởng chiếu sáng là hạng mục bổ sung tùy chọn.',
    schematic: 'Mặt bằng có chú thích, mặt bằng trần & vật liệu sàn, mặt đứng chính, bảng thuyết minh kỹ thuật sơ bộ.',
    schematic2: 'Mặt bằng có chú thích, mặt bằng trần & vật liệu sàn, mặt đứng chính, bảng thuyết minh kỹ thuật sơ bộ.',
    followOn: 'Là gói công việc tiếp theo — phí đề xuất qua phụ lục hợp đồng theo đơn giá đã thống nhất sau khi Thiết kế Cơ sở được duyệt.',
    supervisionOpt: 'Chào giá riêng theo yêu cầu khi dự án bước vào giai đoạn thi công.',
    finalConcept: 'Tuyên ngôn ý tưởng, mặt bằng hoàn chỉnh, moodboard, phối cảnh 3D chính, bảng vật liệu.',
    devDesign: 'Bản vẽ chi tiết, bảng vật liệu, thông số kỹ thuật FF&E.',
    cd: 'Bộ hồ sơ thi công đầy đủ, phối hợp với đơn vị tư vấn.',
    boq: 'Bảng khối lượng sẵn sàng đấu thầu.',
    supervision: 'Rà soát hồ sơ dự thầu, phản hồi RFI, giám sát công trình tới khi hoàn thành.',
  };
  const DS = CURRENT_LANG === 'vi' ? descVI : descEN;
  const statusIncluded = L('\u2713 Included', '\u2713 Bao gồm');
  const statusOptional = L('Optional', 'Tùy chọn');
  const devCdBoqLabel = L('Design Development · Construction Documents · BOQ', `${stageLabel('Design Development')} · ${stageLabel('Construction Documents')} · ${stageLabel('BOQ (Bill of Quantity)')}`);

  const stagesA = [
    { stage: stageLabel('Appraisal and Briefing'), status: statusIncluded, desc: DS.kickoff },
    { stage: stageLabel('Feasibility & Preliminary Concept'), status: statusIncluded, desc: DS.feasibility },
    { stage: stageLabel('Final Concept + Lighting (Optional)'), status: statusIncluded, desc: DS.finalConceptOpt },
    { stage: stageLabel('Schematic Design'), status: statusIncluded, desc: DS.schematic },
    { stage: devCdBoqLabel, status: statusOptional, desc: DS.followOn },
    { stage: stageLabel('Author Supervision'), status: statusOptional, desc: DS.supervisionOpt },
  ];
  const stagesB = [
    { stage: stageLabel('Appraisal and Briefing'), status: statusIncluded, desc: DS.kickoff },
    { stage: stageLabel('Feasibility & Preliminary Concept'), status: statusIncluded, desc: DS.feasibility },
    { stage: stageLabel('Final Concept'), status: statusIncluded, desc: DS.finalConcept },
    { stage: stageLabel('Schematic Design'), status: statusIncluded, desc: DS.schematic2 },
    { stage: stageLabel('Design Development'), status: statusIncluded, desc: DS.devDesign },
    { stage: stageLabel('Construction Documents'), status: statusIncluded, desc: DS.cd },
    { stage: stageLabel('BOQ (Bill of Quantity)'), status: statusIncluded, desc: DS.boq },
    { stage: stageLabel('Author Supervision'), status: statusIncluded, desc: DS.supervision },
  ];
  const stagesC = [
    { stage: stageLabel('Appraisal and Briefing'), status: statusIncluded, desc: DS.kickoff },
    { stage: stageLabel('Feasibility & Preliminary Concept'), status: statusIncluded, desc: DS.feasibility },
    { stage: stageLabel('Final Concept + Lighting (Optional)'), status: statusIncluded, desc: DS.finalConceptOpt },
    { stage: stageLabel('Schematic Design'), status: statusIncluded, desc: DS.schematic2 },
    { stage: stageLabel('Design Development'), status: statusIncluded, desc: DS.devDesign },
    { stage: stageLabel('Construction Documents'), status: statusIncluded, desc: DS.cd },
    { stage: stageLabel('BOQ (Bill of Quantity)'), status: statusIncluded, desc: DS.boq },
    { stage: stageLabel('Author Supervision'), status: statusIncluded, desc: DS.supervision },
  ];
  const stages = scope === 'B' ? stagesB : scope === 'C' ? stagesC : stagesA;

  // Native table (avoids a rendering bug seen with many manually-drawn
  // overlapping shapes when row count is high — see Hourly Rates slide fix)
  const rowH = stages.length > 8 ? 0.44 : stages.length > 6 ? 0.52 : 0.72;
  const descFont = stages.length > 8 ? 7.5 : stages.length > 6 ? 8 : SZ.body;
  const scopeRows = [
    [
      { text: L('STAGE IN SCOPE', 'GIAI ĐOẠN TRONG PHẠM VI'), options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8 } },
      { text: '', options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8 } },
      { text: L('WHAT IT DELIVERS', 'NỘI DUNG BÀN GIAO'), options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8 } },
    ],
    ...stages.map((st, i) => {
      const isOpt = st.status === statusOptional;
      const bg = i % 2 === 0 ? C.rowalt : C.white;
      const fg = isOpt ? C.gray : C.black;
      return [
        { text: st.stage, options: { fill: { color: bg }, color: fg, fontFace: F.light, bold: !isOpt, italic: isOpt, fontSize: descFont, valign: 'top' } },
        { text: isOpt ? statusOptional : statusIncluded, options: { fill: { color: bg }, color: isOpt ? C.gray : C.terradk, fontFace: F.light, bold: !isOpt, fontSize: 7.5, align: 'center', valign: 'top' } },
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
    s.addText(L('Lighting Design — offered separately as an add-on at any stage, not bundled into the fee above.', 'Thiết kế Chiếu sáng — chào giá riêng như một hạng mục bổ sung ở bất kỳ giai đoạn nào, không bao gồm trong phí nêu trên.'), {
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
  slideTitle(s, L('Indicative Zoning Programme', 'Chương trình Phân khu (Zoning)'),
    `${L('H2X proposes', 'H2X đề xuất')} ~${d.area || '[X]'} m² · ~${d.peakGuests || '[n]'} ${L('peak guests', 'khách peak')} · ~${d.sqmPerGuest || '7'} ${L('m²/guest', 'm²/khách')} · ${L('to be confirmed in Concept', 'sẽ chốt ở giai đoạn Ý tưởng')}`);

  const zones = d.zones || [];

  // Header — from Taseco slide 4
  tblHeader(s, [
    { text: L('ZONE', 'KHU VỰC'), x: ML + 0.06, w: 3.7 },
    { text: L('AREA', 'DIỆN TÍCH'), x: ML + 3.86, w: 1.2 },
    { text: L('SEATS', 'CHỖ'), x: ML + 5.06, w: 0.9 },
    { text: L('OPERATIONAL RATIONALE', 'LÝ DO VẬN HÀNH'), x: ML + 6.06, w: 6.0 },
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
      { text: L('Total (net usable)', 'Tổng (diện tích sử dụng thuần)'), x: ML + 0.1, w: 3.6, bold: true, color: C.terradk },
      { text: `~${d.area || totalArea} m²`, x: ML + 3.86, w: 1.1, align: 'center', bold: true, color: C.terradk },
      { text: '', x: ML + 5.06, w: 0.85 },
      { text: L('Circulation (~15%) distributed within each zone.', 'Diện tích lưu thông (~15%) phân bổ trong từng khu vực.'), x: ML + 6.06, w: 6.0, color: C.gray, italic: true },
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

  s.addText(L('Design Language', 'Ngôn ngữ Thiết kế'), {
    x: ML, y: 0.5, w: CW, h: 0.55,
    fontSize: SZ.h1, fontFace: F.serif, color: C.black, margin: 0,
  });
  s.addText(d.designTagline || L('Quiet luxury with a Vietnamese soul — warm materials, light from shadow, restraint over display', 'Sang trọng tĩnh lặng mang hồn Việt — vật liệu ấm áp, ánh sáng từ bóng đổ, tiết chế thay vì phô trương'), {
    x: ML, y: 1.0842, w: CW, h: 0.2308,
    fontSize: SZ.label, fontFace: F.semi, color: C.terra, margin: 0,
  });
  s.addShape('rect', {
    x: ML, y: 1.34, w: CW, h: 0.008,
    fill: { color: C.rule }, line: { color: C.rule, width: 0 },
  });

  // 3 image boxes
  const pillars = d.designPillars || (CURRENT_LANG === 'vi' ? [
    { label: 'ĐÓN TIẾP', sub: 'NGHI THỨC ĐÓN KHÁCH · SỰ CHẠM ẤM ÁP' },
    { label: 'F&B', sub: 'TINH TẾ · GIÀU CẢM XÚC · THONG THẢ' },
    { label: 'GIAO LƯU', sub: 'ẤM ÁP · GIÀU KHÔNG KHÍ · KẾT NỐI' },
  ] : [
    { label: 'WELCOME', sub: 'ARRIVAL RITUAL · THE HUMAN TOUCH' },
    { label: 'F&B', sub: 'REFINED · SENSORY · UNHURRIED' },
    { label: 'SOCIAL', sub: 'WARM · ATMOSPHERIC · CONNECTED' },
  ]);
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
      s.addText(L('[ MOOD IMAGE ]', '[ ẢNH MOOD ]'), {
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
  slideTitle(s, L('Lighting Design — Built In, From Concept', 'Thiết kế Chiếu sáng — Tích hợp từ Ý tưởng'),
    L('Included in the design fee · H2X\'s in-house strength · the single biggest lever on atmosphere',
      'Đã bao gồm trong phí thiết kế · Thế mạnh nội bộ của H2X · yếu tố quyết định lớn nhất đến không khí không gian'));

  // Pull quote
  s.addText(L('"Light is not a final layer — it is the atmosphere. So we design it from the first Concept sketch, not after the walls are built."',
    '"Ánh sáng không phải là lớp hoàn thiện cuối cùng — nó chính là không khí của không gian. Vì vậy chúng tôi thiết kế ánh sáng ngay từ phác thảo Ý tưởng đầu tiên, không phải sau khi tường đã xây xong."'), {
    x: ML, y: 1.45, w: CW, h: 0.6,
    fontSize: SZ.pullquote, fontFace: F.serifI, color: C.black,
    italic: true, margin: 0,
  });

  // Two columns
  const cols = CURRENT_LANG === 'vi' ? [
    {
      x: ML, w: 5.72, label: 'VÌ SAO TÍCH HỢP TỪ Ý TƯỞNG',
      items: [
        'Ánh sáng định hình không khí, nhận diện thương hiệu & trải nghiệm khách ngay từ đầu — tách riêng dễ gây xung đột tầm nhìn và phải thiết kế lại tốn kém về sau',
        'Ý tưởng chiếu sáng phát triển song song với không gian, vật liệu & FF&E → có phối cảnh chân thực sớm, ít vòng lặp chỉnh sửa hơn',
        'Không có rủi ro xung đột giữa kiến trúc sư nội thất và đơn vị tư vấn chiếu sáng bên ngoài',
        'Tối ưu ngân sách chiếu sáng — một phần đáng kể trong FF&E — thay vì bị cắt giảm vào phút chót',
      ],
    },
    {
      x: 7.28, w: 5.72, label: 'H2X BÀN GIAO GÌ (TRONG PHÍ)',
      items: [
        d.lightingStrategy ? truncateWords(d.lightingStrategy, 22) : 'Chiến lược không khí theo từng khu vực — ấn tượng cho F&B, chức năng cho khu làm việc, nhịp sinh học/thư giãn cho khu Yên tĩnh & Chăm sóc sức khỏe',
        'Tích hợp ánh sáng tự nhiên theo tầm nhìn; ánh sáng gián tiếm ấm & đèn hắt trần',
        'Ý tưởng thiết bị & điều khiển; kịch bản ánh sáng theo thời điểm trong ngày & luồng khách',
        'Tuân thủ quy chuẩn năng lượng Việt Nam; thiết bị cao cấp, khả thi thi công qua mạng lưới nhà cung cấp lâu năm của H2X',
      ],
    },
  ] : [
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
  slideTitle(s, L('Project Team', 'Đội ngũ Dự án'), L('A senior, multidisciplinary team — with in-house Lighting Design', 'Đội ngũ cấp cao, đa lĩnh vực — có sẵn năng lực Thiết kế Chiếu sáng nội bộ'));

  const team = d.team || (CURRENT_LANG === 'vi' ? [
    { role: 'GIÁM ĐỐC SÁNG TẠO (CCO)', name: 'Gia Huy (Michael)', title: 'Th.S Kiến trúc · 24+ năm · định hướng tầm nhìn & thiết kế', dark: true },
    { role: 'GIÁM ĐỐC VẬN HÀNH (COO)', name: 'Nguyễn Văn Toàn (Karo)', title: 'Kiến trúc sư · 20+ năm · định hướng tầm nhìn & thiết kế', dark: true },
    { role: 'TRƯỞNG NHÓM Ý TƯỞNG', name: 'Nguyễn Mạnh Hùng (Henry)', title: 'Ý tưởng không gian, câu chuyện thiết kế & hành trình khách' },
    { role: 'TRƯỞNG PHÒNG FF&E', name: 'Nguyễn Thị Thúy Vy (Vivian)', title: 'Nội thất, thiết bị & tuyển chọn vật liệu' },
    { role: 'TRƯỞNG PHÒNG KỸ THUẬT', name: 'Nguyễn Vương Linh (Lucas)', title: 'Hồ sơ kỹ thuật & phối hợp đơn vị tư vấn' },
    { role: '3D DIỄN HỌA & QUẢN LÝ DỰ ÁN', name: 'Đỗ Danh Sơn', title: 'Phối cảnh chân thực · tiến độ, rà soát & bàn giao' },
    { role: 'THIẾT KẾ CHIẾU SÁNG', name: '[Lighting Lead]', title: 'Không khí không gian, thiết bị & chiến lược điều khiển' },
    { role: 'THIẾT KẾ NGHE NHÌN', name: '[AV Lead]', title: 'Âm thanh, tích hợp AV & màn hình' },
    { role: 'PHỤ TRÁCH THƯƠNG HIỆU', name: '[Brand Lead]', title: 'Đặt tên, nhận diện & hướng dẫn thương hiệu' },
  ] : [
    { role: 'CHIEF CREATIVE OFFICER', name: 'Gia Huy (Michael)', title: 'M.Architect · 24+ yrs · sets vision & design direction', dark: true },
    { role: 'CHIEF OPERATING OFFICER', name: 'Nguyễn Văn Toàn (Karo)', title: 'Architect · 20+ yrs · sets vision & design direction', dark: true },
    { role: 'CONCEPT LEAD', name: 'Nguyễn Mạnh Hùng (Henry)', title: 'Spatial concept, narrative & guest journey' },
    { role: 'HEAD OF FF&E', name: 'Nguyễn Thị Thúy Vy (Vivian)', title: 'Furniture, fixtures & material curation' },
    { role: 'TECHNICAL LEAD', name: 'Nguyễn Vương Linh (Lucas)', title: 'Schematic documentation & consultant coordination' },
    { role: '3D VISUALISATION & PM', name: 'Đỗ Danh Sơn', title: 'Photorealistic renders · schedule, reviews & delivery' },
    { role: 'LIGHTING DESIGNER', name: '[Lighting Lead]', title: 'Atmosphere, fixture & control strategy' },
    { role: 'AUDIO VISUAL DESIGNER', name: '[AV Lead]', title: 'Soundscape, AV & display concept' },
    { role: 'BRAND LEAD', name: '[Brand Lead]', title: 'Naming, identity & brand guidelines' },
  ]);

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
  slideTitle(s, L('Design Process & Workflow', 'Quy trình & Cách thức Làm việc'), L('Approval-gated at every stage \u00b7 First Draft \u2192 Review \u2192 Final Approval \u2014 auto-adjusts to K\u1ecbch b\u1ea3n A / B / C', 'Có bước duyệt ở mọi giai đoạn · Bản nháp đầu → Rà soát → Duyệt cuối cùng — tự động khớp theo Kịch bản A / B / C'));

  const scope = d.scope || 'A';
  const defs = getStageDefs(scope);
  const SHORT_LABEL_EN = {
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
  const SHORT_LABEL_VI = {
    'Appraisal and Briefing': 'Đánh giá &\nTóm tắt',
    'Feasibility & Preliminary Concept': 'Khả thi &\nÝ tưởng sơ bộ',
    'Final Concept + Lighting (Optional)': 'Ý tưởng\n+ Chiếu sáng*',
    'Final Concept': 'Thiết kế\nÝ tưởng',
    'Schematic Design': 'Thiết kế\nCơ sở',
    'Design Development': 'Thiết kế\nKỹ thuật',
    'Construction Documents': 'Bản vẽ\nThi công',
    'BOQ (Bill of Quantity)': 'Bảng\nKhối lượng',
    'Author Supervision': 'Giám sát\nTác giả',
  };
  const SHORT_LABEL = CURRENT_LANG === 'vi' ? SHORT_LABEL_VI : SHORT_LABEL_EN;
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
  secLabel(s, L('OUR METHODOLOGY', 'PHƯƠNG PHÁP LÀM VIỆC'), ML, 2.65, CW);
  s.addText(L('We provide a step-by-step method for how we approach a project to achieve the best results at each stage. Every stage runs through a First Draft, one consolidated Review round, then Final Approval before moving forward \u2014 giving the Client full visibility and control at each gate.',
    'Chúng tôi triển khai theo phương pháp từng bước rõ ràng để đạt kết quả tốt nhất ở mỗi giai đoạn. Mỗi giai đoạn đều đi qua Bản nháp đầu, một vòng Rà soát tổng hợp, rồi đến Duyệt cuối cùng trước khi tiếp tục — giúp Khách hàng luôn nắm rõ và kiểm soát tại mỗi mốc duyệt.'), {
    x: ML, y: 2.95, w: CW, h: 0.65,
    fontSize: SZ.body, fontFace: F.light, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.4,
  });

  // OUR PROCESS — one card per stage, always matching the selected Scope
  secLabel(s, L('OUR PROCESS', 'QUY TRÌNH CHI TIẾT'), ML, 3.75, CW);
  const cardY = 4.05;
  const cGap = 0.14;
  const cw = (CW - (count - 1) * cGap) / count;
  const cardFont = count > 6 ? 6.5 : count > 4 ? 7 : 7.5;
  defs.forEach((st, i) => {
    const x = ML + i * (cw + cGap);
    s.addShape('rect', { x, y: cardY, w: cw, h: 2.15, fill: { color: C.rowalt }, line: { color: C.border, width: 0.5 } });
    s.addText(String(i + 1), { x: x + 0.08, y: cardY + 0.08, w: cw - 0.16, h: 0.24, fontSize: 8.5, fontFace: F.bold, color: C.terradk, margin: 0 });
    s.addText(stageLabel(st.stage), {
      x: x + 0.08, y: cardY + 0.32, w: cw - 0.16, h: 0.55,
      fontSize: cardFont + 1, fontFace: F.bold, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.15,
    });
    s.addText(stageDesc(st.stage), {
      x: x + 0.08, y: cardY + 0.9, w: cw - 0.16, h: 1.15,
      fontSize: cardFont, fontFace: F.light, color: C.gray, margin: 0, valign: 'top', lineSpacingMultiple: 1.25,
    });
  });

  const methodologyPrefix = L('Methodology: ', 'Quy trình: ');
  const followOnSuffix = L(' \u2014 full service available as a follow-on addendum.', ' — dịch vụ trọn gói có thể bổ sung qua phụ lục hợp đồng.');
  const lightingSuffix = L(' *Lighting concept offered as an optional add-on.', ' *Ý tưởng chiếu sáng là hạng mục bổ sung tùy chọn.');
  s.addText(`${methodologyPrefix}${defs.map(st => stageLabel(st.stage).split(' + ')[0].split(' (')[0]).join(' \u2192 ')}${scope === 'A' ? followOnSuffix : '.'}${scope !== 'B' ? lightingSuffix : ''}`, {
    x: ML, y: H - 0.6, w: CW, h: 0.22,
    fontSize: 7.5, fontFace: F.light, color: C.gray, margin: 0,
  });

  pageNum(s, 8);
  copyright(s);
}

function slideWorkStageDeliverable(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, L('Work Stage Deliverable', 'Sản phẩm theo Giai đoạn'), L('How each stage moves from draft to signed approval', 'Cách mỗi giai đoạn đi từ bản nháp đến khi được duyệt chính thức'));

  // Flow diagram
  const flow = L(['Kick-off\nMeeting', 'First\nDraft', 'First\nReview', 'Final\nApproval', 'Signed Stage\nApproved'],
    ['Họp\nKhởi động', 'Bản nháp\nĐầu tiên', 'Rà soát\nLần đầu', 'Duyệt\nCuối cùng', 'Giai đoạn\nĐã duyệt']);
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
  const blocks = L([
    { h: 'First Draft', b: 'The Designer prepares a First Draft based on information received and issues it to the Project Manager.' },
    { h: 'First Review', b: 'The PM issues the First Draft to all project stakeholders and forwards their consolidated comments and instructions to the Designer.' },
    { h: 'Final Approval', b: 'After the final presentation, the Designer collates the received comments and instructions. A stage-completion letter is then issued, capturing any minor items to carry into the next stage.' },
  ], [
    { h: 'Bản nháp Đầu tiên', b: 'Nhà thiết kế chuẩn bị Bản nháp Đầu tiên dựa trên thông tin đã nhận và gửi cho Quản lý Dự án.' },
    { h: 'Rà soát Lần đầu', b: 'PM gửi Bản nháp Đầu tiên tới tất cả các bên liên quan của dự án và tổng hợp ý kiến, chỉ đạo gửi lại cho Nhà thiết kế.' },
    { h: 'Duyệt Cuối cùng', b: 'Sau buổi thuyết trình cuối cùng, Nhà thiết kế tổng hợp các ý kiến và chỉ đạo đã nhận. Sau đó, thư hoàn thành giai đoạn sẽ được phát hành, ghi nhận các hạng mục nhỏ cần chuyển sang giai đoạn tiếp theo.' },
  ]);
  blocks.forEach((bl, i) => {
    const x = ML + i * (3.85 + 0.28);
    s.addShape('rect', { x, y: 2.75, w: 3.85, h: 2.15, fill: { color: C.rowalt }, line: { color: C.border, width: 0.5 } });
    s.addText(bl.h, { x: x + 0.18, y: 2.88, w: 3.55, h: 0.32, fontSize: 11, fontFace: F.serif, color: C.black, margin: 0 });
    s.addText(bl.b, {
      x: x + 0.18, y: 3.24, w: 3.55, h: 1.55, fontSize: SZ.body, fontFace: F.light, color: C.black,
      margin: 0, valign: 'top', lineSpacingMultiple: 1.4,
    });
  });

  s.addText(L('Additional rounds: any round requested beyond those stated above will be subject to an extension of time and additional fees as per the Agreement.',
    'Vòng bổ sung: bất kỳ vòng rà soát nào yêu cầu thêm ngoài các vòng nêu trên sẽ phát sinh gia hạn thời gian và phí bổ sung theo Hợp đồng.'), {
    x: ML, y: 5.2, w: CW, h: 0.4, fontSize: 8, fontFace: F.light, color: C.terradk, italic: true, margin: 0, lineSpacingMultiple: 1.3,
  });
  s.addText(L('An Appraisal & Briefing stage is added ahead of Feasibility to align on project approach and planning. Stage durations may be adjusted based on H2X\u2019s evaluation of scope, deliverables, and required Client review/approval time at each stage.',
    'Giai đoạn Đánh giá & Tóm tắt được bổ sung trước giai đoạn Khả thi để thống nhất cách tiếp cận và kế hoạch dự án. Thời lượng từng giai đoạn có thể điều chỉnh dựa trên đánh giá của H2X về phạm vi, sản phẩm bàn giao và thời gian rà soát/duyệt cần thiết của Khách hàng ở mỗi giai đoạn.'), {
    x: ML, y: 5.65, w: CW, h: 0.55, fontSize: 8, fontFace: F.light, color: C.gray, margin: 0, lineSpacingMultiple: 1.3,
  });

  pageNum(s, 9);
  copyright(s);
}

// ─── SLIDE 10 — PROJECT ORGANISATION CHART ───────────────────────
function slideOrgChart(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, L('Project Organisation Chart', 'Sơ đồ Tổ chức Dự án'), L('A senior, multidisciplinary team structured around your project', 'Đội ngũ cấp cao, đa lĩnh vực, được tổ chức xoay quanh dự án của bạn'));

  const team = d.team || (CURRENT_LANG === 'vi' ? [
    { role: 'GIÁM ĐỐC SÁNG TẠO (CCO)', name: 'Gia Huy (Michael)' },
    { role: 'GIÁM ĐỐC VẬN HÀNH (COO)', name: 'Nguy\u1ec5n V\u0103n To\u00e0n (Karo)' },
    { role: 'TRƯỞNG NHÓM Ý TƯỞNG', name: 'Nguy\u1ec5n M\u1ea1nh H\u00f9ng (Henry)' },
    { role: 'TRƯỞNG PHÒNG FF&E', name: 'Nguy\u1ec5n Th\u1ecb Th\u00fay Vy (Vivian)' },
    { role: 'TRƯỞNG PHÒNG KỸ THUẬT', name: 'Nguy\u1ec5n V\u01b0\u01a1ng Linh (Lucas)' },
    { role: '3D DIỄN HỌA & QUẢN LÝ DỰ ÁN', name: '\u0110\u1ed7 Danh S\u01a1n' },
    { role: 'THIẾT KẾ CHIẾU SÁNG', name: '[Lighting Lead]' },
    { role: 'THIẾT KẾ NGHE NHÌN', name: '[AV Lead]' },
    { role: 'PHỤ TRÁCH THƯƠNG HIỆU', name: '[Brand Lead]' },
  ] : [
    { role: 'CHIEF CREATIVE OFFICER', name: 'Gia Huy (Michael)' },
    { role: 'CHIEF OPERATING OFFICER', name: 'Nguy\u1ec5n V\u0103n To\u00e0n (Karo)' },
    { role: 'CONCEPT LEAD', name: 'Nguy\u1ec5n M\u1ea1nh H\u00f9ng (Henry)' },
    { role: 'HEAD OF FF&E', name: 'Nguy\u1ec5n Th\u1ecb Th\u00fay Vy (Vivian)' },
    { role: 'TECHNICAL LEAD', name: 'Nguy\u1ec5n V\u01b0\u01a1ng Linh (Lucas)' },
    { role: '3D VISUALISATION & PM', name: '\u0110\u1ed7 Danh S\u01a1n' },
    { role: 'LIGHTING DESIGNER', name: '[Lighting Lead]' },
    { role: 'AUDIO VISUAL DESIGNER', name: '[AV Lead]' },
    { role: 'BRAND LEAD', name: '[Brand Lead]' },
  ]);
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
  const support = L(['Architecture &\nConcept Team', 'Interior &\nFF&E Team', 'Lighting &\nTechnical Team', '3D Visualisation\nTeam', 'Brand & AV\nTeam'],
    ['Kiến trúc &\nÝ tưởng', 'Nội thất &\nFF&E', 'Chiếu sáng &\nKỹ thuật', 'Diễn họa\n3D', 'Thương hiệu\n& AV']);
  const sw = (CW - (support.length - 1) * 0.14) / support.length;
  s.addText(L('Supported by', 'Được hỗ trợ bởi'), { x: ML, y: t4Y - 0.32, w: 3, h: 0.25, fontSize: 8, fontFace: F.light, color: C.gray, italic: true, margin: 0 });
  support.forEach((sup, i) => {
    const x = ML + i * (sw + 0.14);
    s.addShape('rect', { x, y: t4Y, w: sw, h: 0.6, fill: { color: C.white }, line: { color: C.graylt, width: 0.75 } });
    s.addText(sup, { x: x + 0.06, y: t4Y, w: sw - 0.12, h: 0.6, fontSize: 7, fontFace: F.light, color: C.black, align: 'center', valign: 'middle', margin: 0, lineSpacingMultiple: 1.15 });
  });

  s.addText(L('Every project is led directly by the CCO/COO and a dedicated Lighting Designer works in-house from Concept, not bolted on later.',
    'Mọi dự án đều do CCO/COO trực tiếp phụ trách, cùng một Chuyên viên Thiết kế Chiếu sáng làm việc nội bộ ngay từ giai đoạn Ý tưởng, không bổ sung sau.'), {
    x: ML, y: 5.65, w: CW, h: 0.3, fontSize: 8, fontFace: F.light, color: C.gray, italic: true, margin: 0,
  });

  pageNum(s, 10);
  copyright(s);
}

// ─── SLIDE 11 — DESIGN WORK PHASE ────────────────────────────────
function slideDesignWorkPhase(pres, d) {
  const s = pres.addSlide();
  const scope = d.scope || 'A';
  slideTitle(s, L('Design Work Phase', 'Các Giai đoạn Thiết kế'), `${L('What H2X delivers at each stage', 'H2X bàn giao gì ở từng giai đoạn')} \u2014 K\u1ecbch b\u1ea3n ${scope} \u2014 ${L('deliverable-level detail', 'chi tiết theo sản phẩm bàn giao')}`);

  const WORK_PHASE_DELIVERABLE_EN = {
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
  const WORK_PHASE_DELIVERABLE_VI = {
    'Appraisal and Briefing': 'Họp khởi động; thống nhất phạm vi; khảo sát hiện trạng; rà soát tài liệu sẵn có (mặt bằng hiện trạng, bản vẽ kết cấu/M&E).',
    'Feasibility & Preliminary Concept': 'Đánh giá khả thi & vận hành; mặt bằng phân khu sơ bộ; bảng vật liệu; thuyết minh ý tưởng thiết kế (DWG/PDF).',
    'Final Concept + Lighting (Optional)': 'Look & Feel / Moodboard & câu chuyện thiết kế, mặt bằng màu, phối cảnh 3D màu, bảng vật liệu số \u2014 tất cả dạng PDF.',
    'Final Concept': 'Look & Feel / Moodboard & câu chuyện thiết kế, mặt bằng màu, phối cảnh 3D màu, bảng vật liệu số \u2014 tất cả dạng PDF.',
    'Schematic Design': 'Mặt bằng có chú thích, mặt bằng trần & vật liệu sàn, mặt đứng chính, bảng thuyết minh kỹ thuật sơ bộ (DWG/PDF).',
    'Design Development': 'Bản vẽ kỹ thuật đầy đủ: thi công nội thất, đồ nội thất cố định, phối hợp M&E, danh mục đồ rời, bảng thông số kỹ thuật (DWG/PDF).',
    'Construction Documents': 'Bộ hồ sơ bản vẽ thi công đầy đủ, phối hợp với toàn bộ đơn vị tư vấn (DWG/PDF).',
    'BOQ (Bill of Quantity)': 'Bảng khối lượng sẵn sàng đấu thầu.',
    'Author Supervision': 'Rà soát hồ sơ dự thầu, phản hồi RFI, và giám sát công trình định kỳ tới khi hoàn thành.',
  };
  const WORK_PHASE_DELIVERABLE = CURRENT_LANG === 'vi' ? WORK_PHASE_DELIVERABLE_VI : WORK_PHASE_DELIVERABLE_EN;
  const defs = getStageDefs(scope);
  const phases = defs.map(st => ({
    name: stageLabel(st.stage),
    lead: st.stage.includes('Lighting') ? L('H2X (Design + Lighting)', 'H2X (Thiết kế + Chiếu sáng)') : 'H2X',
    deliverable: WORK_PHASE_DELIVERABLE[st.stage] || '',
  }));

  tblHeader(s, [
    { text: L('PHASE', 'GIAI ĐOẠN'), x: ML + 0.06, w: 3.4 },
    { text: L('LEAD', 'PHỤ TRÁCH'), x: ML + 3.55, w: 2.1 },
    { text: L('KEY DELIVERABLES', 'SẢN PHẨM BÀN GIAO CHÍNH'), x: ML + 5.75, w: 6.25 },
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

  s.addText(L('Detailed purpose, scope-of-work and full deliverable breakdown for each phase follow on the next pages.', 'Mục đích chi tiết, phạm vi công việc và sản phẩm bàn giao đầy đủ cho từng giai đoạn được trình bày ở các trang tiếp theo.'), {
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

  secLabel(s, L('PURPOSE OF THE STAGE', 'MỤC ĐÍCH GIAI ĐOẠN'), ML, 1.55, CW);
  s.addText(data.purpose, {
    x: ML, y: 1.83, w: CW, h: 0.55,
    fontSize: SZ.body, fontFace: F.light, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.35,
  });

  secLabel(s, L('COVERED WITHIN THE STAGE', 'NỘI DUNG TRONG GIAI ĐOẠN'), ML, 2.5, 7.0);
  let cy = 2.78;
  data.covered.forEach(line => {
    s.addText('+', { x: ML, y: cy, w: 0.2, h: 0.3, fontSize: SZ.body, fontFace: F.bold, color: C.terra, margin: 0 });
    s.addText(line, {
      x: ML + 0.22, y: cy, w: 6.78, h: 0.5,
      fontSize: 8, fontFace: F.light, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.25,
    });
    cy += 0.02 + Math.ceil(line.length / 78) * 0.22;
  });

  secLabel(s, L('DELIVERABLES', 'SẢN PHẨM BÀN GIAO'), ML + 7.4, 2.5, 4.9);
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
  const en = {
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
  };
  const vi = {
    title: 'Tiếp nhận Yêu cầu & Khởi động',
    leadLine: 'PHỤ TRÁCH — H2X & Khách hàng · HỖ TRỢ — Đơn vị tư vấn',
    purpose: 'Đảm bảo dự án khởi động với đầy đủ đội ngũ thiết kế và tư vấn, cùng yêu cầu đã được thống nhất.',
    covered: [
      'H2X được giới thiệu với các thành viên còn lại trong Đội ngũ Thiết kế; rà soát phạm vi tổng thể và quy hoạch chung để xác định các khoảng trống (nếu có).',
      'H2X hoàn thành Đánh giá & Tóm tắt với đầy đủ tài liệu cần thiết để khởi động giai đoạn Ý tưởng: mặt bằng công trình, mặt đứng, bản vẽ kết cấu & M&E hiện trạng (file DWG), ảnh hiện trạng, hình ảnh phối cảnh, cảnh quan.',
      'Khảo sát công trình (nếu có thể) cùng Khách hàng và Tư vấn trưởng để đánh giá vị trí, khu vực xung quanh, các hạn chế/ràng buộc của công trình.',
      'Rà soát yêu cầu cùng Khách hàng để cập nhật các thay đổi kể từ khi ký hợp đồng, hướng dẫn, yêu cầu chương trình và phạm vi công việc.',
    ],
    deliverables: ['Họp khởi động dự án'],
  };
  slidePhaseExplainer(pres, d, 12, CURRENT_LANG === 'vi' ? vi : en);
}

function slideFeasibility(pres, d) {
  const en = {
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
  };
  const vi = {
    title: 'Khả thi (Tiền Ý tưởng)',
    leadLine: 'PHỤ TRÁCH — H2X · HỖ TRỢ — Đơn vị tư vấn',
    purpose: 'Xây dựng câu chuyện thiết kế (design narrative).',
    covered: [
      'Nghiên cứu các đặc điểm riêng của vị trí và công trình, đồng thời hiểu sâu hơn về yêu cầu và định hướng thương hiệu.',
      'Chuẩn bị bộ sưu tập hình ảnh tham chiếu để xác định định hướng ý tưởng thiết kế.',
      'Chuẩn bị mặt bằng phân khu nội thất sơ bộ thể hiện chương trình công năng.',
      'Chuẩn bị bảng vật liệu (hình ảnh tham chiếu) và câu chuyện thiết kế.',
      'Thuyết trình lần 1 với Khách hàng về định hướng ý tưởng nội thất.',
    ],
    deliverables: ['Trang Tổng quan Ý tưởng (DWG/PDF)'],
  };
  slidePhaseExplainer(pres, d, 13, CURRENT_LANG === 'vi' ? vi : en);
}

function slideConceptDesignExplainer(pres, d) {
  const en = {
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
  };
  const vi = {
    title: 'Thiết kế Ý tưởng',
    leadLine: 'PHỤ TRÁCH — H2X · HỖ TRỢ — Đơn vị tư vấn · Thực hiện song song với giai đoạn Khả thi',
    purpose: 'Xác lập định hướng và ý tưởng thiết kế đã thống nhất, phản ánh đúng tính chất của các khu vực dự án theo yêu cầu.',
    covered: [
      'Trang câu chuyện thiết kế: xây dựng điểm nhấn riêng làm nền tảng cho thiết kế.',
      'Thuyết trình Mood & Feel: hình ảnh cảm hứng và phác thảo ban đầu diễn giải tổng thể phong cách và cảm xúc không gian.',
      'Nội thất & Thiết bị: lựa chọn ban đầu về nội thất, chiếu sáng, thiết bị và phụ kiện để lấy ý kiến phản hồi, kèm phác thảo tùy chỉnh ban đầu.',
      'Bảng vật liệu: trình bày ban đầu về hoàn thiện kiến trúc và bọc nệm tại các khu vực chính.',
    ],
    deliverables: ['Look & Feel, Moodboard và câu chuyện thiết kế (PDF)', 'Mặt bằng màu', 'Phối cảnh 3D màu', 'Bảng vật liệu kỹ thuật số (PDF)'],
  };
  slidePhaseExplainer(pres, d, 14, CURRENT_LANG === 'vi' ? vi : en);
}

function slideSchematicExplainer(pres, d) {
  const en = {
    title: 'Schematic Design',
    leadLine: 'LEAD \u2014 H2X \u00b7 SUPPORT \u2014 Consultants \u00b7 Kicks off once Concept Design is approved',
    purpose: 'Provide DWG drawings and basic schedules to the design team for development into the Detail Design package.',
    covered: [
      'Plans clearly annotated with key finishes, including RCP and floor-finishes plan.',
      'Key elevations clearly annotated with key finishes, agreed at the end of Concept Design.',
      'Outline specification schedules for architectural finishes, presented as index pages for further development.',
    ],
    deliverables: ['Plans package (DWG/PDF)', 'Elevations (DWG/PDF)', 'Table of Materials (Office files/PDF)', 'General layout, floor finish, electrical & RCP plans', 'Key sections showing furniture elevations'],
  };
  const vi = {
    title: 'Thiết kế Cơ sở',
    leadLine: 'PHỤ TRÁCH — H2X · HỖ TRỢ — Đơn vị tư vấn · Bắt đầu sau khi Thiết kế Ý tưởng được duyệt',
    purpose: 'Cung cấp bản vẽ DWG và các bảng biểu cơ bản cho đội ngũ thiết kế để phát triển thành hồ sơ Thiết kế Chi tiết.',
    covered: [
      'Mặt bằng có chú thích rõ ràng về vật liệu hoàn thiện chính, bao gồm mặt bằng trần và vật liệu sàn.',
      'Mặt đứng chính có chú thích rõ ràng về vật liệu hoàn thiện, đã thống nhất vào cuối giai đoạn Thiết kế Ý tưởng.',
      'Bảng thuyết minh kỹ thuật sơ bộ cho vật liệu hoàn thiện kiến trúc, trình bày dưới dạng trang mục lục để phát triển tiếp.',
    ],
    deliverables: ['Bộ hồ sơ mặt bằng (DWG/PDF)', 'Mặt đứng (DWG/PDF)', 'Bảng vật liệu (File Office/PDF)', 'Mặt bằng tổng thể, vật liệu sàn, điện & trần', 'Mặt cắt chính thể hiện mặt đứng nội thất'],
  };
  slidePhaseExplainer(pres, d, 15, CURRENT_LANG === 'vi' ? vi : en);
}

function slideDevelopmentExplainer(pres, d) {
  const en = {
    title: 'Development Design',
    leadLine: 'LEAD \u2014 H2X \u00b7 SUPPORT \u2014 Consultants \u00b7 Develops from the approved Schematic Design',
    purpose: 'Develop the entire technical drawing set to ensure the design follows the direction set in Schematic Design.',
    covered: [
      'Full development of technical drawings for the entire interior, including walls and separate items.',
      'Complete technical solutions for the entire project, ensuring accurate estimation, bidding invitation and design feasibility.',
      'Coordination of MEP design across the full technical document set.',
    ],
    deliverables: ['Technical docs: Fit-out, Built-in (DWG/PDF)', 'Technical docs: full MEP design (DWG/PDF)', 'Specifications (Digital list)', 'Loose Furniture list', 'RCP, wall/floor finish & electrical plans', 'Large-scale detail drawings'],
  };
  const vi = {
    title: 'Thiết kế Phát triển',
    leadLine: 'PHỤ TRÁCH — H2X · HỖ TRỢ — Đơn vị tư vấn · Phát triển từ Thiết kế Cơ sở đã duyệt',
    purpose: 'Phát triển toàn bộ bộ bản vẽ kỹ thuật để đảm bảo thiết kế bám sát định hướng đã xác lập ở giai đoạn Thiết kế Cơ sở.',
    covered: [
      'Phát triển đầy đủ bản vẽ kỹ thuật cho toàn bộ nội thất, bao gồm tường và các hạng mục riêng lẻ.',
      'Giải pháp kỹ thuật hoàn chỉnh cho toàn bộ dự án, đảm bảo tính chính xác trong dự toán, mời thầu và khả thi thiết kế.',
      'Phối hợp thiết kế MEP xuyên suốt toàn bộ hồ sơ kỹ thuật.',
    ],
    deliverables: ['Hồ sơ kỹ thuật: Fit-out, Built-in (DWG/PDF)', 'Hồ sơ kỹ thuật: Thiết kế MEP đầy đủ (DWG/PDF)', 'Thông số kỹ thuật (Danh mục điện tử)', 'Danh mục nội thất rời', 'Mặt bằng trần, vật liệu tường/sàn & điện', 'Bản vẽ chi tiết tỷ lệ lớn'],
  };
  slidePhaseExplainer(pres, d, 16, CURRENT_LANG === 'vi' ? vi : en);
}

function slideConstructionDocsExplainer(pres, d) {
  const en = {
    title: 'Construction Documents',
    leadLine: 'LEAD \u2014 H2X \u00b7 SUPPORT \u2014 Consultants \u00b7 Final phase, developed from approved Development Design',
    purpose: 'Produce a complete set of detailed technical construction drawings for site execution \u2014 the official documentation for construction on site.',
    covered: [
      'Develop the full set of technical construction drawings based on the approved design: fit-out and built-in drawings.',
      'Detailed drawings of walls, ceilings, flooring, loose furniture, and built components.',
      'Precise technical drawings, specifications and schedules ensuring accuracy, feasibility and coordination across all construction activities.',
    ],
    deliverables: ['Construction Drawings Package (DWG/PDF)', 'Schedules & technical specs for hard finishes, sanitary fittings, equipment', 'FF&E selection'],
  };
  const vi = {
    title: 'Thiết kế bản vẽ thi công',
    leadLine: 'PHỤ TRÁCH — H2X · HỖ TRỢ — Đơn vị tư vấn · Giai đoạn cuối, phát triển từ Thiết kế Phát triển đã duyệt',
    purpose: 'Xuất bộ hồ sơ bản vẽ kỹ thuật thi công chi tiết đầy đủ để triển khai tại công trường — tài liệu chính thức phục vụ thi công.',
    covered: [
      'Phát triển đầy đủ bộ bản vẽ kỹ thuật thi công dựa trên thiết kế đã duyệt: bản vẽ fit-out và built-in.',
      'Bản vẽ chi tiết tường, trần, sàn, nội thất rời và các hạng mục xây dựng cố định.',
      'Bản vẽ kỹ thuật, thông số và bảng biểu chính xác đảm bảo độ chuẩn xác, tính khả thi và phối hợp xuyên suốt các hoạt động thi công.',
    ],
    deliverables: ['Bộ hồ sơ Bản vẽ Thi công (DWG/PDF)', 'Bảng biểu & thông số kỹ thuật vật liệu hoàn thiện cứng, thiết bị vệ sinh, thiết bị', 'Lựa chọn FF&E'],
  };
  slidePhaseExplainer(pres, d, 17, CURRENT_LANG === 'vi' ? vi : en);
}

// ─── SLIDE 09 — DETAILED DESIGN STAGE
// ─── SLIDE 09 — DETAILED DESIGN STAGE ────────────────────────────
function slide09(pres, d) {
  const s = pres.addSlide();
  const scope = d.scope || 'A';
  slideTitle(s, L('Detailed Design Stage', 'Chi tiết các Giai đoạn Thiết kế'), `${L('Design stage breakdown', 'Phân bổ theo giai đoạn thiết kế')} · Kịch bản ${scope} · ${L('weeks & ratio synced with Document Issue Schedule', 'số tuần & tỷ lệ đồng bộ với Tiến độ Phát hành Hồ sơ')}`);

  const defs = d.docSchedule && d.docSchedule.length ? d.docSchedule : getStageDefs(scope);
  const totalWeeks = defs.reduce((sum, st) => sum + (Number(st.weeks) || 0), 0);

  tblHeader(s, [
    { text: '#', x: ML + 0.06, w: 0.6, align: 'center' },
    { text: L('PHASE', 'GIAI ĐOẠN'), x: ML + 0.75, w: 7.5 },
    { text: L('WEEKS', 'SỐ TUẦN'), x: ML + 8.35, w: 1.2, align: 'center' },
    { text: L('RATIO', 'TỶ LỆ'), x: ML + 9.65, w: 2.35, align: 'center' },
  ], 1.55, 0.35);

  const rowH = defs.length > 5 ? 0.5 : 0.65;
  defs.forEach((st, i) => {
    const y = 1.55 + 0.35 + i * rowH;
    tblRow(s, [
      { text: String(i + 1), x: ML + 0.1, w: 0.55, align: 'center' },
      { text: stageLabel(st.stage), x: ML + 0.75, w: 7.45 },
      { text: String(st.weeks), x: ML + 8.35, w: 1.15, align: 'center' },
      { text: st.ratio + '%', x: ML + 9.65, w: 2.35, align: 'center' },
    ], y, rowH - 0.05, i % 2 === 0 ? C.white : C.rowalt);
  });

  const totY = 1.55 + 0.35 + defs.length * rowH;
  tblRow(s, [
    { text: L('TOTAL', 'TỔNG'), x: ML + 0.75, w: 7.7, bold: true, color: C.terradk },
    { text: String(totalWeeks), x: ML + 8.35, w: 1.15, align: 'center', bold: true, color: C.terradk },
    { text: '100%', x: ML + 9.65, w: 2.35, align: 'center', bold: true, color: C.terradk },
  ], totY, 0.4, C.rowtot);

  s.addText(L('Client Approval required at the end of each stage before proceeding.', 'Cần Khách hàng duyệt vào cuối mỗi giai đoạn trước khi tiếp tục.') + (scope === 'A' ? L(' Design Development / Construction Documents / BOQ / Author Supervision available as a follow-on addendum.', ' Thiết kế kĩ thuật / Bản vẽ thi công / Bảng khối lượng / Giám sát tác giả có thể bổ sung qua phụ lục hợp đồng.') : L(' Bid/Tender/Procurement and Handover priced separately by others.', ' Đấu thầu/Mua sắm và Bàn giao được báo giá riêng bởi đơn vị khác.')), {
    x: ML, y: totY + 0.5, w: CW, h: 0.35,
    fontSize: 7.5, fontFace: F.light, color: C.gray, italic: true, margin: 0, lineSpacingMultiple: 1.3,
  });

  pageNum(s, 20);
  copyright(s);
}

// ─── SLIDE 10 — DETAILED SOW LIST ────────────────────────────────
function slide10(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, L('Detailed SOW List', 'Danh mục Công việc Chi tiết'), L('Scope of Works breakdown by deliverable, mapped to K\u1ecbch b\u1ea3n A / B / C', 'Phân bổ Phạm vi Công việc theo sản phẩm bàn giao, khớp với Kịch bản A / B / C'));

  const scope = d.scope || 'A';
  const groupsEN = [
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
  const groupsVI = [
    { num: '0', name: 'Đánh giá và tóm tắt', opt: [1, 1, 1] },
    { num: '1', name: 'Tính khả thi và ý tưởng sơ bộ', items: [
      { num: '1.1', name: 'Đánh giá khả thi và vận hành', opt: [1, 1, 1] },
      { num: '1.2', name: 'Mặt bằng ý tưởng điều chỉnh', opt: [1, 1, 1] },
      { num: '1.3', name: 'Đề xuất phong cách & cảm xúc thiết kế', opt: [1, 1, 1] },
    ]},
    { num: '2', name: 'Thiết kế Ý tưởng', items: [
      { num: '2.1', name: 'Tuyên ngôn ý tưởng, mặt bằng hoàn chỉnh & Moodboard', opt: [1, 1, 1] },
      { num: '2.2', name: 'Hình ảnh & Video Render', opt: [1, 1, 1] },
    ]},
    { num: '3', name: 'Thiết kế Cơ sở', items: [
      { num: '3.1', name: 'Bản vẽ Thiết kế Cơ sở', opt: [1, 1, 1] },
      { num: '3.2', name: 'Tư vấn Thiết kế Chiếu sáng', opt: [1, 0, 1], note: 'Tùy chọn / chào giá riêng ở Kịch bản B' },
      { num: '3.3', name: 'Tư vấn Biển báo & Nghe nhìn (AV)', opt: [0, 0, 0], note: 'Do đơn vị tư vấn chuyên môn, hạng mục bổ sung' },
    ]},
    { num: '4', name: 'Thiết kế kĩ thuật', items: [
      { num: '4.1', name: 'Thiết kế kĩ thuật Nội thất', opt: [0, 1, 1] },
      { num: '4.2', name: 'Thiết kế kĩ thuật MEP', opt: [0, 1, 1] },
    ]},
    { num: '5', name: 'Giai đoạn Thi công', items: [
      { num: '5.1', name: 'Thiết kế bản vẽ thi công', opt: [0, 1, 1] },
      { num: '5.2', name: 'Bảng khối lượng (BOQ)', opt: [0, 1, 1], note: 'Phục vụ khối lượng đấu thầu' },
      { num: '5.3', name: 'Đấu thầu / Mua sắm', opt: [0, 0, 0], note: 'Trách nhiệm của Khách hàng' },
      { num: '5.4', name: 'Kiểm soát Giá trị & Ngân sách', opt: [0, 0, 0] },
    ]},
    { num: '6', name: 'Giám sát tác giả', items: [
      { num: '6.1', name: 'Rà soát Hồ sơ Dự thầu & Phản hồi RFI', opt: [0, 1, 1] },
      { num: '6.2', name: 'Giám sát Công trình Định kỳ', opt: [0, 1, 1] },
    ]},
    { num: '7', name: 'Bàn giao', opt: [0, 1, 1] },
  ];
  const groups = d.sowGroups && d.sowGroups.length ? d.sowGroups : (CURRENT_LANG === 'vi' ? groupsVI : groupsEN);
  const optLabels = [L('SCENARIO A', 'KỊCH BẢN A'), L('SCENARIO B', 'KỊCH BẢN B'), L('SCENARIO C', 'KỊCH BẢN C')];
  const activeCol = scope === 'A' ? 0 : scope === 'B' ? 1 : 2;
  const optW = 1.0, optX = [ML + 5.45, ML + 6.55, ML + 7.65];

  tblHeader(s, [
    { text: '#', x: ML + 0.06, w: 0.5, align: 'center' },
    { text: L('SCOPE OF WORKS', 'PHẠM VI CÔNG VIỆC'), x: ML + 0.65, w: 4.7 },
    { text: optLabels[0], x: optX[0], w: optW, align: 'center', fontSize: 6.5 },
    { text: optLabels[1], x: optX[1], w: optW, align: 'center', fontSize: 6.5 },
    { text: optLabels[2], x: optX[2], w: optW, align: 'center', fontSize: 6.5 },
    { text: L('NOTE', 'GHI CHÚ'), x: ML + 8.75, w: 3.25 },
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
  s.addText(L(`Option ${activeCol + 1} (highlighted \u2713) is the scope proposed in this fee proposal.`, `Cột ${activeCol + 1} (đánh dấu ✓ nổi bật) là phạm vi được đề xuất trong báo giá này.`), {
    x: ML, y: noteY, w: CW, h: 0.2, fontSize: 7.5, fontFace: F.light, color: C.terradk, italic: true, margin: 0,
  });

  pageNum(s, 21);
  copyright(s);
}

// ─── SLIDE 11 — DOCUMENT ISSUE SCHEDULE ──────────────────────────
function slide11(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, L('Document Issue Schedule', 'Tiến độ Phát hành Hồ sơ'), L('Preliminary schedule · Interior Design + Lighting Design · lump sum by phase', 'Tiến độ sơ bộ · Thiết kế Nội thất + Thiết kế Chiếu sáng · trọn gói theo giai đoạn'));

  const scope = d.scope || 'A';
  const defs = d.docSchedule && d.docSchedule.length ? d.docSchedule : getStageDefs(scope);
  const sched = defs.map((st, i) => ({ num: String(i + 1), stage: st.stage, weeks: Number(st.weeks) || 0, ratio: st.ratio }));
  const totalWeeks = sched.reduce((sum, st) => sum + st.weeks, 0);

  tblHeader(s, [
    { text: L('STAGE', 'GIAI ĐOẠN'), x: ML + 0.06, w: 0.6 },
    { text: L('PHASE', 'GIAI ĐOẠN'), x: ML + 0.75, w: 7.5 },
    { text: L('WEEKS', 'SỐ TUẦN'), x: ML + 8.35, w: 1.2, align: 'center' },
    { text: L('RATIO', 'TỶ LỆ'), x: ML + 9.65, w: 2.35, align: 'center' },
  ], 1.55, 0.35);

  const rowH = sched.length > 4 ? 0.5 : 0.65;
  sched.forEach((st, i) => {
    const y = 1.55 + 0.35 + i * rowH;
    tblRow(s, [
      { text: st.num, x: ML + 0.1, w: 0.55, align: 'center' },
      { text: stageLabel(st.stage), x: ML + 0.75, w: 7.45 },
      { text: String(st.weeks), x: ML + 8.35, w: 1.15, align: 'center' },
      { text: st.ratio + '%', x: ML + 9.65, w: 2.35, align: 'center' },
    ], y, rowH - 0.05, i % 2 === 0 ? C.rowalt : C.white);
  });

  const totY = 1.55 + 0.35 + sched.length * rowH;
  tblRow(s, [
    { text: L('TOTAL', 'TỔNG'), x: ML + 0.75, w: 7.7, bold: true, color: C.terradk },
    { text: String(totalWeeks), x: ML + 8.35, w: 1.15, align: 'center', bold: true, color: C.terradk },
    { text: '100%', x: ML + 9.65, w: 2.35, align: 'center', bold: true, color: C.terradk },
  ], totY, 0.4, C.rowtot);

  // Proportional timeline bar
  secLabel(s, L('TIMELINE OVERVIEW (BY WEEK)', 'TỔNG QUAN TIẾN ĐỘ (THEO TUẦN)'), ML, totY + 0.5, CW);
  const barY = totY + 0.8, barH = 0.45;
  const shades = [C.black, C.terradk, C.terra, C.gray, C.dark, C.rowtot];
  let bx = ML;
  sched.forEach((st, i) => {
    const segW = totalWeeks > 0 ? (st.weeks / totalWeeks) * CW : 0;
    s.addShape('rect', { x: bx, y: barY, w: segW, h: barH, fill: { color: shades[i % shades.length] }, line: { color: C.white, width: 1 } });
    s.addText(`${st.weeks}${L('w', 't')}`, {
      x: bx, y: barY, w: segW, h: barH, fontSize: 8, fontFace: F.bold,
      color: C.white, align: 'center', valign: 'middle', margin: 0,
    });
    bx += segW;
  });
  s.addText(L(`Estimated total duration: ~${totalWeeks} weeks (excludes Client approval turnaround time).`, `Tổng thời lượng dự kiến: ~${totalWeeks} tuần (chưa gồm thời gian Khách hàng rà soát & duyệt).`), {
    x: ML, y: barY + barH + 0.1, w: CW, h: 0.22, fontSize: 7.5, fontFace: F.light, color: C.gray, italic: true, margin: 0,
  });

  pageNum(s, 18);
  copyright(s);
}

// ─── SLIDE 12 — DESIGN FEE ────────────────────────────────────────
function slide12(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, L('Design Fee', 'Phí Thiết kế'), L('Interior Design + Lighting Design · lump sum by phase', 'Thiết kế Nội thất + Thiết kế Chiếu sáng · trọn gói theo giai đoạn'));

  const scope = d.scope || 'A';
  const fee = parseFloat(d.feeTotal) || 0;
  const fmt = n => n > 0 ? n.toLocaleString('en-US') : '[—]';

  const defs = getStageDefs(scope);
  const stages = defs.map((st, i) => ({
    num: String(i + 1), stage: st.stage, ratio: st.ratio + '%', amt: Math.round(fee * st.ratio / 100),
  }));

  tblHeader(s, [
    { text: L('STAGE', 'GIAI ĐOẠN'), x: ML + 0.06, w: 0.6 },
    { text: L('PHASE', 'GIAI ĐOẠN'), x: ML + 0.75, w: 7.5 },
    { text: L('RATIO', 'TỶ LỆ'), x: ML + 8.35, w: 1.2 },
    { text: L('FEE (USD)', 'PHÍ (USD)'), x: ML + 9.65, w: 2.35, align: 'center' },
  ], 1.55, 0.35);

  const rowH = stages.length > 4 ? 0.5 : 0.7;
  stages.forEach((st, i) => {
    const y = 1.55 + 0.35 + i * rowH;
    tblRow(s, [
      { text: st.num, x: ML + 0.1, w: 0.55, align: 'center' },
      { text: stageLabel(st.stage), x: ML + 0.75, w: 7.45 },
      { text: st.ratio, x: ML + 8.35, w: 1.15, align: 'center' },
      { text: fmt(st.amt), x: ML + 9.65, w: 2.35, align: 'center' },
    ], y, rowH - 0.05, i % 2 === 0 ? C.rowalt : C.white);
  });

  // Subtotal
  const totY = 1.55 + 0.35 + stages.length * rowH;
  tblRow(s, [
    { text: L('INTERIOR + LIGHTING — SUBTOTAL', 'NỘI THẤT + CHIẾU SÁNG — TẠM TÍNH'), x: ML + 0.75, w: 8.8, bold: true, color: C.terradk },
    { text: '100%', x: ML + 8.35, w: 1.15, align: 'center', bold: true, color: C.terradk },
    { text: fmt(fee), x: ML + 9.65, w: 2.35, align: 'center', bold: true, color: C.terradk },
  ], totY, 0.45, C.rowtot);

  // Benchmark + notes (consolidated into one tight block to fit varying row counts)
  const area = parseFloat(d.area) || 0;
  const perSqm = area > 0 ? Math.round(fee / area) : 0;
  const scopeNote = scope === 'A'
    ? L('Design Development / Construction Documents / BOQ / Author Supervision available as a follow-on addendum.', 'Thiết kế kĩ thuật / Bản vẽ thi công / Bảng khối lượng / Giám sát tác giả có thể bổ sung qua phụ lục hợp đồng.')
    : L('Lighting Design available as an optional add-on — see Add-on Workstreams.', 'Thiết kế Chiếu sáng có thể chào giá riêng như hạng mục bổ sung — xem Hạng mục Bổ sung.');
  const noteLines = [
    L(`≈ ${perSqm > 0 ? perSqm + ' USD / m²' : '[USD/m²]'} · ${d.area || '[X]'} m² · indicative benchmark for this typology.`,
      `≈ ${perSqm > 0 ? perSqm + ' USD / m²' : '[USD/m²]'} · ${d.area || '[X]'} m² · định mức tham khảo cho loại hình này.`),
    scopeNote,
    L('Fees above exclude taxes and other charges.', 'Phí nêu trên chưa bao gồm thuế và các chi phí khác.'),
    L('Using additional services within our ecosystem entitles you to a discount on the design fee.', 'Sử dụng thêm các dịch vụ trong hệ sinh thái của chúng tôi sẽ được ưu đãi trên phí thiết kế.'),
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
  slideTitle(s, L('Branding · Signage · AV — Scope & Value', 'Thương hiệu · Biển báo · Nghe nhìn — Phạm vi & Giá trị'), null);
  s.addText(L('Three workstreams that turn a beautiful space into a coherent, memorable destination — offered at concept stage', 'Ba hạng mục biến một không gian đẹp thành một điểm đến nhất quán, đáng nhớ — chào giá ở giai đoạn ý tưởng'), {
    x: ML, y: 1.0842, w: CW, h: 0.2308,
    fontSize: SZ.label, fontFace: F.semi, color: C.terra, margin: 0,
  });

  const addonsEN = [
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
  const addonsVI = [
    { title: 'Nhận diện thương hiệu & Đặt tên', partner: 'Ziva',
      scope: ['Tên & câu chuyện của không gian', 'Logo & nhận diện thị giác cốt lõi', 'Hướng dẫn thương hiệu (cơ bản)'],
      why: ['Một không gian có tên riêng tạo lòng trung thành, giá trị PR & hợp tác', 'Một câu chuyện nhất quán xuyên suốt mọi điểm chạm với khách'] },
    { title: 'Biển báo & Chỉ dẫn', partner: 'Ziva',
      scope: ['Chiến lược chỉ dẫn song ngữ', 'Ý tưởng biển báo & các loại chính', 'Định hướng thông số kỹ thuật'],
      why: ['Điều hướng dễ dàng giảm lo lắng & giảm tải cho nhân viên', 'Biển báo chính là thương hiệu ở tầm mắt — độ rõ đạt chuẩn công trình'] },
    { title: 'Thiết kế Nghe nhìn', partner: 'H2X',
      scope: ['Định hướng âm thanh đặc trưng', 'Ý tưởng tích hợp AV & màn hình', 'Phân vùng âm thanh theo khu vực'],
      why: ['Âm thanh hoàn thiện không khí mà ánh sáng khởi tạo', 'AV tạo không khí, không phải tiếng ồn — thông tin đáng tin cậy, nhẹ nhàng'] },
  ];
  const addons = CURRENT_LANG === 'vi' ? addonsVI : addonsEN;

  const cw = 3.84;
  addons.forEach((a, i) => {
    const x = ML + i * (cw + 0.18);
    s.addShape('rect', { x, y: 1.5, w: cw, h: 5.5, fill: { color: C.rowalt }, line: { color: C.border, width: 0.5 } });
    s.addText(a.title, { x: x + 0.18, y: 1.65, w: cw - 0.3, h: 0.45, fontSize: 13, fontFace: F.serif, color: C.black, margin: 0 });
    s.addText(`${L('Partner', 'Đối tác')}: ${a.partner}`, { x: x + 0.18, y: 2.14, w: cw - 0.3, h: 0.22, fontSize: 8, fontFace: F.bold, color: C.terra, margin: 0 });
    secLabel(s, L('SCOPE (1 CONCEPT)', 'PHẠM VI (1 Ý TƯỞNG)'), x + 0.18, 2.42, cw - 0.3);
    s.addText(a.scope.map(l => `+ ${l}`).join('\n'), { x: x + 0.18, y: 2.72, w: cw - 0.3, h: 0.95, fontSize: SZ.body, fontFace: F.light, color: C.black, margin: 0 });
    secLabel(s, L('WHY IT MATTERS', 'VÌ SAO QUAN TRỌNG'), x + 0.18, 3.75, cw - 0.3);
    s.addText(a.why.map(l => `+ ${l}`).join('\n'), { x: x + 0.18, y: 4.05, w: cw - 0.3, h: 1.8, fontSize: SZ.body, fontFace: F.light, color: C.black, margin: 0 });
  });

  pageNum(s, 22);
  copyright(s);
}

// ─── SLIDE 14 — ADD-ON FEE ───────────────────────────────────────
function slide14(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, L('Add-on Workstreams', 'Hạng mục Bổ sung'), L('Concept stage · itemised so each can be selected independently · design fees only (fabrication/production excluded)', 'Giai đoạn ý tưởng · tách riêng từng hạng mục để chọn độc lập · chỉ gồm phí thiết kế (chưa gồm gia công/sản xuất)'));

  const fmt = n => n > 0 ? n.toLocaleString('en-US') : '[—]';
  const feeMain = parseFloat(d.feeTotal) || 0;
  const feeBrand = d.addBrand ? parseFloat(d.feeBrand) || 0 : 0;
  const feeSign = d.addSignage ? parseFloat(d.feeSignage) || 0 : 0;
  const feeAV = d.addAV ? parseFloat(d.feeAV) || 0 : 0;
  const addons = feeBrand + feeSign + feeAV;
  const grand = feeMain + addons;

  const rows = [
    { ws: L('Brand Identity & Naming', 'Nhận diện thương hiệu & Đặt tên'), partner: 'Ziva', scope: L('Naming, logo, core identity & essential guidelines.', 'Đặt tên, logo, nhận diện cốt lõi & hướng dẫn cơ bản.'), fee: feeBrand, show: d.addBrand },
    { ws: L('Signage & Wayfinding', 'Biển báo & Chỉ dẫn'), partner: 'Ziva', scope: L('Bilingual wayfinding strategy & signage concept.', 'Chiến lược chỉ dẫn song ngữ & ý tưởng biển báo.'), fee: feeSign, show: d.addSignage },
    { ws: L('Audio-Visual Design', 'Thiết kế Nghe nhìn'), partner: 'H2X', scope: L('Soundscape, AV & display concept.', 'Ý tưởng âm thanh, AV & màn hình.'), fee: feeAV, show: d.addAV },
  ];

  tblHeader(s, [
    { text: L('WORKSTREAM', 'HẠNG MỤC'), x: ML + 0.06, w: 3.2 },
    { text: L('PARTNER', 'ĐỐI TÁC'), x: ML + 3.35, w: 1.1 },
    { text: L('SCOPE (CONCEPT, 1 DIRECTION)', 'PHẠM VI (Ý TƯỞNG, 1 ĐỊNH HƯỚNG)'), x: ML + 4.55, w: 5.5 },
    { text: L('FEE (USD)', 'PHÍ (USD)'), x: ML + 10.15, w: 1.85, align: 'center' },
  ], 1.55, 0.35);

  rows.forEach((r, i) => {
    const y = 1.55 + 0.35 + i * 0.65;
    tblRow(s, [
      { text: r.ws, x: ML + 0.1, w: 3.1 },
      { text: r.partner, x: ML + 3.35, w: 1.05, align: 'center' },
      { text: r.scope, x: ML + 4.55, w: 5.45 },
      { text: r.show ? fmt(r.fee) : L('Available on request', 'Chào giá theo yêu cầu'), x: ML + 10.15, w: 1.85, align: 'center', italic: !r.show, color: r.show ? null : C.gray, fontSize: r.show ? SZ.body : 7.5 },
    ], y, 0.6, i % 2 === 0 ? C.rowalt : C.white);
  });

  const subY = 1.55 + 0.35 + rows.length * 0.65;
  tblRow(s, [
    { text: L('Add-ons subtotal', 'Tạm tính hạng mục bổ sung'), x: ML + 0.1, w: 9.9, bold: true, color: C.terradk },
    { text: fmt(addons), x: ML + 10.15, w: 1.85, align: 'center', bold: true, color: C.terradk },
  ], subY, 0.55, C.rowtot);

  // Grand total
  const gtY = subY + 0.65;
  s.addShape('rect', { x: ML, y: gtY, w: CW, h: 0.7, fill: { color: C.black }, line: { color: C.black, width: 0 } });
  s.addText(L('GRAND TOTAL', 'TỔNG CỘNG'), { x: ML + 0.15, y: gtY + 0.08, w: 6, h: 0.32, fontSize: 10, fontFace: F.light, color: C.white, charSpacing: 2, margin: 0 });
  s.addText(L(`Interior + Lighting (${fmt(feeMain)}) + Add-ons (${fmt(addons)}). Excludes taxes.`, `Nội thất + Chiếu sáng (${fmt(feeMain)}) + Hạng mục bổ sung (${fmt(addons)}). Chưa gồm thuế.`), {
    x: ML + 0.15, y: gtY + 0.42, w: 8, h: 0.22, fontSize: 7.5, fontFace: F.light, color: C.gray, margin: 0,
  });
  s.addText(`${fmt(grand)} USD`, {
    x: ML + 10.15, y: gtY, w: 1.85, h: 0.7, fontSize: 12, fontFace: F.bold, color: C.terra, align: 'center', valign: 'middle', margin: 0,
  });

  // Discount
  s.addShape('rect', { x: ML, y: gtY + 0.78, w: CW, h: 0.35, fill: { color: 'EDE8E2' }, line: { color: 'D6C8B8', width: 0.5 } });
  s.addText(L(`Early-commitment partnership discount −10% if signed within 14 days  →  ${fmt(Math.round(grand * 0.9))} USD`, `Ưu đãi cam kết sớm −10% nếu ký hợp đồng trong vòng 14 ngày  →  ${fmt(Math.round(grand * 0.9))} USD`), {
    x: ML + 0.15, y: gtY + 0.82, w: CW - 0.3, h: 0.25, fontSize: SZ.body, fontFace: F.light, color: C.terradk, italic: true, margin: 0,
  });

  pageNum(s, 23);
  copyright(s);
}

// ─── SLIDE 15 — PAYMENT SCHEDULE ────────────────────────────────
function slide15(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, L('Payment Schedule', 'Tiến độ Thanh toán'), L('Milestone-based, aligned to stage approvals', 'Theo mốc thanh toán, khớp với các mốc duyệt giai đoạn'));

  const scope = d.scope || 'A';
  const fee = parseFloat(d.feeTotal) || 0;
  const fmt = n => n > 0 ? n.toLocaleString('en-US') : '[—]';

  const defs = getStageDefs(scope);
  const pmts = defs.map((st, i) => ({
    num: String(i + 1),
    m: `${stageLabel(st.stage)} ${L('approval', 'được duyệt')}`,
    pct: st.ratio + '%',
    amt: Math.round(fee * st.ratio / 100),
    note: i === 0 ? L('Upon signing', 'Khi ký hợp đồng') : (i === defs.length - 1 ? L('Final milestone', 'Mốc cuối cùng') : ''),
  }));

  tblHeader(s, [
    { text: '#', x: ML + 0.06, w: 0.5, align: 'center' },
    { text: L('MILESTONE', 'MỐC THANH TOÁN'), x: ML + 0.65, w: 6.55 },
    { text: '%', x: ML + 7.3, w: 0.75, align: 'center' },
    { text: L('FEE DESIGN (USD)', 'PHÍ THIẾT KẾ (USD)'), x: ML + 8.15, w: 1.85, align: 'center' },
    { text: L('NOTE', 'GHI CHÚ'), x: ML + 10.1, w: 1.9 },
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
    { text: L('Subtotal — Design Fee', 'Tạm tính — Phí Thiết kế'), x: ML + 0.65, w: 6.5, bold: true, color: C.terradk },
    { text: '100%', x: ML + 7.3, w: 0.7, align: 'center', bold: true, color: C.terradk },
    { text: fmt(fee), x: ML + 8.15, w: 1.8, align: 'center', bold: true, color: C.terradk },
    { text: '', x: ML + 10.1, w: 1.85 },
  ], totY, 0.38, C.rowtot);

  // Add-ons (if any selected) — compact single line + grand total, auto-summed
  const addonRows = [
    d.addBrand ? { label: L('Brand Identity & Naming', 'Nhận diện thương hiệu & Đặt tên'), amt: parseFloat(d.feeBrand) || 0 } : null,
    d.addSignage ? { label: L('Signage & Wayfinding', 'Biển báo & Chỉ dẫn'), amt: parseFloat(d.feeSignage) || 0 } : null,
    d.addAV ? { label: L('Audio-Visual Design', 'Thiết kế Nghe nhìn'), amt: parseFloat(d.feeAV) || 0 } : null,
  ].filter(Boolean);
  const addonsTotal = addonRows.reduce((sum, r) => sum + r.amt, 0);

  let y2 = totY + 0.38;
  if (addonRows.length) {
    tblRow(s, [
      { text: `${L('Add-ons', 'Hạng mục bổ sung')} — ${addonRows.map(r => r.label).join(' · ')}`, x: ML + 0.65, w: 7.3, italic: true, color: C.gray, fontSize: 8 },
      { text: fmt(addonsTotal), x: ML + 8.15, w: 1.8, align: 'center', italic: true, color: C.gray },
      { text: L('Billed on delivery', 'Thanh toán khi bàn giao'), x: ML + 10.1, w: 1.85, fontSize: 7.5, color: C.gray, italic: true },
    ], y2, 0.32, C.white);
    y2 += 0.32;
    tblRow(s, [
      { text: L('GRAND TOTAL (Design Fee + Add-ons)', 'TỔNG CỘNG (Phí Thiết kế + Hạng mục Bổ sung)'), x: ML + 0.65, w: 7.3, bold: true, color: C.white },
      { text: fmt(fee + addonsTotal), x: ML + 8.15, w: 1.8, align: 'center', bold: true, color: C.white },
      { text: '', x: ML + 10.1, w: 1.85 },
    ], y2, 0.38, C.black);
    y2 += 0.38;
  }

  s.addText(L('Add-on workstreams billed 50% on kick-off and 50% on delivery unless folded into the schedule above. Invoices in USD; payment within 14 days of invoice.', 'Hạng mục bổ sung thanh toán 50% khi khởi động và 50% khi bàn giao, trừ khi đã gộp vào tiến độ nêu trên. Hóa đơn bằng USD; thanh toán trong vòng 14 ngày kể từ ngày xuất hóa đơn.'), {
    x: ML, y: y2 + 0.08, w: CW, h: 0.3,
    fontSize: 7, fontFace: F.light, color: C.gray, italic: true, margin: 0, lineSpacingMultiple: 1.15,
  });

  pageNum(s, 24);
  copyright(s);
}

// ─── SLIDE 16 — INCLUSIONS & EXCLUSIONS ──────────────────────────
function slide16(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, L('Inclusions & Exclusions', 'Hạng mục Bao gồm & Không bao gồm'), null);

  const scope = d.scope || 'A';
  const isVI = CURRENT_LANG === 'vi';
  const incA = isVI ? [
    'Thiết kế nội thất — Đánh giá đến hết Thiết kế Cơ sở',
    'Thiết kế Chiếu sáng — ý tưởng cho giai đoạn Thiết kế Ý tưởng (tùy chọn, gộp trong phí thiết kế)',
    'Đánh giá khả thi, phân khu & quy hoạch không gian, thuyết minh ý tưởng, moodboard',
    'Phối cảnh 3D các góc chính · bảng vật liệu & FF&E',
    'Mặt bằng cơ sở, mặt bằng trần, mặt đứng & thuyết minh kỹ thuật sơ bộ',
    'Thương hiệu, Biển báo & Nghe nhìn — giai đoạn ý tưởng (1 định hướng mỗi hạng mục) nếu được chọn',
  ] : [
    'Interior design — Appraisal through Schematic Design',
    'Lighting Design — concept for Final Concept stage (optional, integrated in the design fee)',
    'Feasibility assessment, zoning & space planning, design narrative, moodboards',
    '3D key-view renders · material & FF&E palette',
    'Schematic plans, RCP, elevations & outline specifications',
    'Brand, Signage & AV — concept stage (1 direction each) if selected',
  ];
  const incBC = isVI ? [
    'Thiết kế nội thất — Đánh giá đến hết Giám sát tác giả (trọn gói)',
    scope === 'C' ? 'Thiết kế Chiếu sáng — ý tưởng cho giai đoạn Thiết kế Ý tưởng (tùy chọn, gộp trong phí thiết kế)' : 'Thiết kế Chiếu sáng — chào giá riêng như hạng mục bổ sung (không gộp trong phí này)',
    'Đánh giá khả thi, phân khu & quy hoạch không gian, thuyết minh ý tưởng, moodboard',
    'Phối cảnh 3D các góc chính · bảng vật liệu & FF&E',
    'Bộ bản vẽ kỹ thuật đầy đủ: Thiết kế Cơ sở, Thiết kế kĩ thuật, Bản vẽ thi công',
    'Bảng khối lượng (BOQ) — sẵn sàng đấu thầu',
    'Giám sát tác giả — rà soát hồ sơ dự thầu, phản hồi RFI, giám sát công trình định kỳ',
    'Thương hiệu, Biển báo & Nghe nhìn — giai đoạn ý tưởng (1 định hướng mỗi hạng mục) nếu được chọn',
  ] : [
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

  const excA = isVI ? [
    'Thiết kế kĩ thuật, Bản vẽ thi công, Bảng khối lượng & Giám sát tác giả (bổ sung qua phụ lục hợp đồng)',
    'Phát triển & sản xuất thương hiệu, biển báo & nghe nhìn ngoài giai đoạn ý tưởng',
    'Kiến trúc phần vỏ công trình; kỹ thuật MEP, kết cấu & mặt dựng',
    'Tư vấn bếp, âm học & phòng cháy chữa cháy',
    'Giấy phép & phê duyệt · thuế/VAT',
    'Chi phí hoàn trả: đi lại, in ấn, mô hình mẫu (theo thực tế, cần duyệt trước)',
    'Thi công & công tác tại công trường',
  ] : [
    'Design Development, Construction Documents, BOQ & Author Supervision (follow-on addendum)',
    'Development & production of branding, signage & AV beyond concept',
    'Base-building architecture; MEP, structural & façade engineering',
    'Kitchen consultant, acoustics & fire engineering',
    'Permits & approvals · taxes/VAT',
    'Reimbursables: travel, printing, mock-ups (at cost with prior approval)',
    'Construction & site works',
  ];
  const excBC = isVI ? [
    scope === 'B' ? 'Thiết kế Chiếu sáng (chào giá riêng như hạng mục bổ sung)' : null,
    'Phát triển & sản xuất thương hiệu, biển báo & nghe nhìn ngoài giai đoạn ý tưởng',
    'Kiến trúc phần vỏ công trình; kỹ thuật MEP, kết cấu & mặt dựng',
    'Tư vấn bếp, âm học & phòng cháy chữa cháy',
    'Giấy phép & phê duyệt · thuế/VAT',
    'Chi phí hoàn trả: đi lại, in ấn, mô hình mẫu (theo thực tế, cần duyệt trước)',
    'Thi công & công tác tại công trường (xây dựng thực tế)',
    'Quản lý đấu thầu ngoài phạm vi rà soát hồ sơ dự thầu (mua sắm là trách nhiệm của Khách hàng)',
  ].filter(Boolean) : [
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
    { label: L('INCLUSIONS', 'BAO GỒM'), items: inc, x: ML, bg: C.rowalt, color: C.black },
    { label: L('EXCLUSIONS', 'KHÔNG BAO GỒM'), items: exc, x: ML + 6.18, bg: C.white, color: C.gray },
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
  slideTitle(s, L('Terms & Conditions', 'Điều khoản & Điều kiện'), null);

  const termsEN = [
    { h: 'Revisions', b: 'Each stage includes a First Draft + one consolidated review round; further rounds by addendum (time + fee).' },
    { h: 'Site visits', b: 'Up to three (3) included in the design fee; additional visits charged at cost.' },
    { h: 'Reimbursables', b: 'Travel, printing & physical mock-ups billed at cost with prior approval.' },
    { h: 'Intellectual property', b: 'Design IP transfers to the Client upon full settlement of fees.' },
    { h: 'Validity', b: 'This proposal is valid for 30 days from the date of issue.' },
    { h: 'Currency & tax', b: 'All fees in USD, exclusive of VAT and other statutory charges.' },
    { h: 'Schedule', b: 'Timeline assumes timely Client feedback at each approval gate.' },
  ];
  const termsVI = [
    { h: 'Chỉnh sửa', b: 'Mỗi giai đoạn bao gồm 1 Bản nháp đầu + 1 vòng rà soát tổng hợp; các vòng tiếp theo tính phụ lục bổ sung (thời gian + phí).' },
    { h: 'Khảo sát công trình', b: 'Bao gồm tối đa ba (3) lần trong phí thiết kế; thêm lần sẽ tính phí theo thực tế.' },
    { h: 'Chi phí hoàn trả', b: 'Đi lại, in ấn & mô hình mẫu vật lý tính theo thực tế, cần duyệt trước.' },
    { h: 'Sở hữu trí tuệ', b: 'Bản quyền thiết kế chuyển giao cho Khách hàng sau khi thanh toán đầy đủ.' },
    { h: 'Hiệu lực', b: 'Báo giá này có hiệu lực 30 ngày kể từ ngày phát hành.' },
    { h: 'Đơn vị tiền tệ & thuế', b: 'Toàn bộ phí tính bằng USD, chưa bao gồm VAT và các khoản thuế/phí khác theo quy định.' },
    { h: 'Tiến độ', b: 'Timeline giả định Khách hàng phản hồi đúng hạn ở mỗi mốc duyệt.' },
  ];
  const terms = CURRENT_LANG === 'vi' ? termsVI : termsEN;

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
    s.addText(L('[ CLOSING IMAGE ]', '[ ẢNH KẾT ]'), { x: 0, y: 0, w: 4.5, h: H, fontSize: 8, fontFace: F.light, color: C.gray, align: 'center', valign: 'middle', margin: 0 });
  }

  // Thank You
  s.addText(L('Thank You', 'Trân trọng cảm ơn'), {
    x: 5.2, y: 1.6, w: 7.7, h: 1.0,
    fontSize: 42, fontFace: F.serif, color: C.terra, margin: 0,
  });

  // CTA
  s.addText(d.ctaText || L('We would welcome a working session to walk through the zoning & layout options, align on scope & budget, and confirm the schedule — so we can move into Concept without delay.', 'Chúng tôi rất mong có một buổi làm việc để cùng trao đổi các phương án phân khu & mặt bằng, thống nhất phạm vi & ngân sách, và chốt tiến độ — để có thể bước vào giai đoạn Ý tưởng ngay.'), {
    x: 5.2, y: 2.78, w: 7.7, h: 0.85,
    fontSize: SZ.body, fontFace: F.light, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.5,
  });

  // Contacts
  const contacts = [
    d.phone || '+84 9 6652 6662',
    `${d.website || 'www.h2xstudio.com.vn'} · ${d.email || 'info@h2xstudio.com'}`,
    d.addressHN || `${L('HQ', 'Trụ sở')} — 26 Trần Hưng Đạo, Hà Nội`,
    d.addressHCM || `${L('Branch', 'Chi nhánh')} — 199D Nguyễn Văn Hưởng, Thảo Điền, TP HCM`,
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
  slideTitle(s, L('Hourly Rates & Lithography Charges', 'Đơn giá Giờ công & Phí In ấn'), L('Appendix C & D \u2014 additional services time charges & printing rates (USD)', 'Phụ lục C & D \u2014 đơn giá giờ công dịch vụ bổ sung & phí in ấn (USD)'));

  const ratesEN = [
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
  const ratesVI = [
    { role: 'Giám đốc Sáng tạo / Chủ trì', hr: '$29', day: '$230' },
    { role: 'Giám đốc Vận hành', hr: '$23', day: '$185' },
    { role: 'Chuyên gia Kỹ thuật Cao cấp', hr: '$25', day: '$200' },
    { role: 'Phó Giám đốc / Trưởng nhóm', hr: '$13', day: '$104' },
    { role: 'Phó Trưởng nhóm', hr: '$11', day: '$88' },
    { role: 'Thiết kế viên Cao cấp', hr: '$10', day: '$81' },
    { role: 'Chuyên viên 3D Diễn họa Cao cấp', hr: '$8', day: '$65' },
    { role: 'Thiết kế viên / Thiết kế viên Kỹ thuật', hr: '$6', day: '$50' },
    { role: 'Chuyên viên FF&E', hr: '$5', day: '$38' },
    { role: 'Kế toán / Hành chính', hr: '$3.50', day: '$27' },
  ];
  const rates = d.hourlyRates || (CURRENT_LANG === 'vi' ? ratesVI : ratesEN);

  // Left — Hourly / Daily Rates (native table — avoids a rendering bug seen with
  // many manually-drawn overlapping shapes on this text-dense slide)
  secLabel(s, L('APPENDIX C \u2014 HOURLY / DAILY RATES', 'PHỤ LỤC C \u2014 ĐƠN GIÁ GIỜ / NGÀY CÔNG'), ML, 1.5, 6.9);
  const rateRows = [
    [
      { text: L('ROLE', 'VAI TRÒ'), options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8, bold: false } },
      { text: L('PER HOUR', 'THEO GIỜ'), options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8, align: 'center' } },
      { text: L('PER DAY', 'THEO NGÀY'), options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8, align: 'center' } },
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
  secLabel(s, L('APPENDIX D \u2014 LITHOGRAPHY & PRINTING (USD)', 'PHỤ LỤC D \u2014 IN ẤN & LITHOGRAPHY (USD)'), rx, 1.5, rw);
  const litho = [
    { size: 'A4', bw: '$0.05', col: '$0.20', mnt: '$2.50' },
    { size: 'A3', bw: '$0.10', col: '$0.40', mnt: '$4.00' },
    { size: 'A2', bw: '$0.15', col: '$1.20', mnt: '$7.00' },
    { size: 'A1', bw: '$0.20', col: '$2.50', mnt: '$12' },
    { size: 'A0', bw: '$0.35', col: '$4.50', mnt: '$18' },
  ];
  const lithoRows = [
    [
      { text: L('SIZE', 'KÍCH THƯỚC'), options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8, align: 'center' } },
      { text: L('B&W', 'ĐEN TRẮNG'), options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8, align: 'center' } },
      { text: L('COLOUR', 'MÀU'), options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8, align: 'center' } },
      { text: L('MOUNT', 'BỒI BẢNG'), options: { fill: { color: C.black }, color: 'FFFFFF', fontFace: F.light, fontSize: 8, align: 'center' } },
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
  s.addText(L('Other: Soft copy transfer $12\u2013$19/time \u00b7 Binding $2\u2013$4/set \u00b7 Print resize $1/copy', 'Khác: Chuyển file mềm $12\u2013$19/lần \u00b7 Đóng gáy $2\u2013$4/bộ \u00b7 Chỉnh kích thước in $1/bản'), {
    x: rx, y: litY, w: rw, h: 0.4,
    fontSize: 7.5, fontFace: F.light, color: C.gray, italic: true, margin: 0, lineSpacingMultiple: 1.25,
  });

  const noteY = 1.8 + 0.32 + rates.length * 0.36 + 0.15;
  s.addText(L('Rates exclude VAT and other applicable taxes \u00b7 Reference conversion 1 USD \u2248 26,000 VND \u00b7 Rates apply for the current year and may be revised periodically.', 'Đơn giá chưa bao gồm VAT và các loại thuế liên quan khác \u00b7 Tỷ giá tham khảo 1 USD \u2248 26.000 VND \u00b7 Đơn giá áp dụng cho năm hiện tại và có thể được điều chỉnh định kỳ.'), {
    x: ML, y: Math.max(noteY, litY + 0.5), w: CW, h: 0.4,
    fontSize: 7.5, fontFace: F.light, color: C.gray, italic: true, margin: 0, lineSpacingMultiple: 1.25,
  });

  pageNum(s, 27);
  copyright(s);
}

// ─── SLIDE — BIM FEE SCHEDULE ─────────────────────────────────────
function slideBIMFee(pres, d) {
  const s = pres.addSlide();
  slideTitle(s, L('BIM Fee Schedule', 'Bảng Phí BIM'), L('Optional premium for delivering the project using Building Information Modelling', 'Phụ phí tùy chọn khi triển khai dự án bằng công nghệ Building Information Modelling (BIM)'));

  const bimRolesEN = [
    { role: 'Partner / Design Director', hr: '$360', day: '$3,240', wk: '$16,200', mo: '$64,800' },
    { role: 'Associate Designer', hr: '$190', day: '$1,710', wk: '$8,550', mo: '$34,200' },
    { role: 'Senior Interior Designer / Senior FF&E', hr: '$130', day: '$1,170', wk: '$5,850', mo: '$23,400' },
    { role: 'Interior Designer / FF&E', hr: '$100', day: '$900', wk: '$4,500', mo: '$18,000' },
    { role: 'Senior CAD/BIM Technician', hr: '$70', day: '$630', wk: '$3,150', mo: '$12,600' },
  ];
  const bimRolesVI = [
    { role: 'Đối tác / Giám đốc Thiết kế', hr: '$360', day: '$3,240', wk: '$16,200', mo: '$64,800' },
    { role: 'Thiết kế viên Chính', hr: '$190', day: '$1,710', wk: '$8,550', mo: '$34,200' },
    { role: 'Thiết kế viên Nội thất Cao cấp / FF&E Cao cấp', hr: '$130', day: '$1,170', wk: '$5,850', mo: '$23,400' },
    { role: 'Thiết kế viên Nội thất / FF&E', hr: '$100', day: '$900', wk: '$4,500', mo: '$18,000' },
    { role: 'Kỹ thuật viên CAD/BIM Cao cấp', hr: '$70', day: '$630', wk: '$3,150', mo: '$12,600' },
  ];
  const bimRoles = d.bimRates || (CURRENT_LANG === 'vi' ? bimRolesVI : bimRolesEN);

  secLabel(s, L('BIM RATES \u2014 MONTHLY & HOURLY (USD)', 'ĐƠN GIÁ BIM \u2014 THEO THÁNG & THEO GIỜ (USD)'), ML, 1.5, CW);
  tblHeader(s, [
    { text: L('DESIGNATION', 'CHỨC DANH'), x: ML + 0.06, w: 4.0 },
    { text: L('HOURLY', 'THEO GIỜ'), x: ML + 4.1, w: 1.9, align: 'center' },
    { text: L('DAILY', 'THEO NGÀY'), x: ML + 6.05, w: 1.9, align: 'center' },
    { text: L('WEEKLY', 'THEO TUẦN'), x: ML + 8.0, w: 1.9, align: 'center' },
    { text: L('MONTHLY', 'THEO THÁNG'), x: ML + 9.95, w: 2.0, align: 'center' },
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
  secLabel(s, L('LOD STANDARDS AT H2X', 'TIÊU CHUẨN LOD TẠI H2X'), ML, lodY, CW);
  const lodsEN = [
    { l: 'LOD 100 (Concept Design)', d: 'Massing, general orientation and basic space allocations for broad visualisation.' },
    { l: 'LOD 200 (Schematic Design)', d: 'Preliminary dimensions, specific locations and relationships between model elements.' },
    { l: 'LOD 300/350 (Detailed Design)', d: 'Detailed geometry, accurate dimensions and clash detection across all consultants.' },
  ];
  const lodsVI = [
    { l: 'LOD 100 (Thiết kế Ý tưởng)', d: 'Khối hình tổng thể, định hướng chung và phân bổ không gian cơ bản để hình dung tổng quan.' },
    { l: 'LOD 200 (Thiết kế Cơ sở)', d: 'Kích thước sơ bộ, vị trí cụ thể và mối quan hệ giữa các thành phần mô hình.' },
    { l: 'LOD 300/350 (Thiết kế Chi tiết)', d: 'Hình học chi tiết, kích thước chính xác và kiểm tra xung đột giữa các đơn vị tư vấn.' },
  ];
  const lods = CURRENT_LANG === 'vi' ? lodsVI : lodsEN;
  const lw = (CW - 2 * 0.2) / 3;
  lods.forEach((item, i) => {
    const x = ML + i * (lw + 0.2);
    const y = lodY + 0.3;
    s.addShape('rect', { x, y, w: lw, h: 1.15, fill: { color: C.rowalt }, line: { color: C.border, width: 0.5 } });
    s.addText(item.l, { x: x + 0.12, y: y + 0.1, w: lw - 0.24, h: 0.4, fontSize: 8.5, fontFace: F.bold, color: C.terradk, margin: 0, valign: 'top', lineSpacingMultiple: 1.2 });
    s.addText(item.d, { x: x + 0.12, y: y + 0.5, w: lw - 0.24, h: 0.6, fontSize: 7.5, fontFace: F.light, color: C.black, margin: 0, valign: 'top', lineSpacingMultiple: 1.25 });
  });

  s.addText(L('BIM authored in Autodesk Revit \u00b7 coordination via Navisworks \u00b7 aligned to ISO 19650-1/2 \u00b7 BIM delivery is an optional premium, quoted separately from the base Design Fee.', 'BIM được thực hiện trên Autodesk Revit \u00b7 phối hợp qua Navisworks \u00b7 tuân theo tiêu chuẩn ISO 19650-1/2 \u00b7 Phí BIM là khoản phụ phí tùy chọn, báo giá riêng ngoài Phí Thiết kế cơ bản.'), {
    x: ML, y: lodY + 1.6, w: CW, h: 0.35,
    fontSize: 7.5, fontFace: F.light, color: C.gray, italic: true, margin: 0, lineSpacingMultiple: 1.25,
  });

  pageNum(s, 28);
  copyright(s);
}

async function generateProposal(data) {
  setLang(data.language);
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
