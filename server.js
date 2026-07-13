const express = require('express');
const path = require('path');
const fs = require('fs');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');
const { generateProposal } = require('./generatePptx');

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ─── FIREBASE AUTH VERIFICATION (no service account needed — verifies ─
// ─── the ID token's signature against Google's public keys) ──────────
const FIREBASE_PROJECT_ID = 'h2x-tools-auth';
const client = jwksClient({
  jwksUri: 'https://www.googleapis.com/service_accounts/v1/metadata/x509/securetoken@system.gserviceaccount.com',
});

function getSigningKey(header, callback) {
  client.getSigningKey(header.kid, (err, key) => {
    if (err) return callback(err);
    callback(null, key.publicKey || key.rsaPublicKey);
  });
}

// Verifies the request is from a genuinely signed-in Firebase user of this
// project. Does NOT check per-tool Firestore permission (that's enforced
// client-side by Firestore Security Rules) — this middleware only blocks
// fully anonymous/direct API abuse from outside the app.
function requireAuth(req, res, next) {
  const authHeader = req.headers.authorization || '';
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;
  if (!token) {
    return res.status(401).json({ error: 'Chưa đăng nhập. Vui lòng tải lại trang và đăng nhập bằng Google.' });
  }
  jwt.verify(token, getSigningKey, {
    algorithms: ['RS256'],
    audience: FIREBASE_PROJECT_ID,
    issuer: `https://securetoken.google.com/${FIREBASE_PROJECT_ID}`,
  }, (err, decoded) => {
    if (err) {
      return res.status(401).json({ error: 'Phiên đăng nhập không hợp lệ hoặc đã hết hạn. Vui lòng tải lại trang.' });
    }
    req.userEmail = decoded.email;
    next();
  });
}

// ─── TYPOLOGY ZONE TEMPLATES ─────────────────────────────────────
const ZONE_TEMPLATES = {
  'Hotel & Resort': [
    'Lobby & Arrival',
    'Lounge & All-day Dining',
    'Signature Restaurant / Bar',
    'Pool & Outdoor Leisure',
    'Spa & Fitness',
    'Meeting & Event Space',
    'Back-of-House',
    'Guest Circulation & Lift Lobby'
  ],
  'F&B Restaurant': [
    'Entry & Reception / Host Stand',
    'Main Dining Room',
    'Private Dining Room (VIP)',
    'Bar & Pre-dining Lounge',
    'Wine Cellar / Display',
    'Amenities (WC / Cloakroom)',
    'Back-of-House / Kitchen',
    'Staff & Service Corridor'
  ],
  'Lounge & Club': [
    'Entry & Welcome',
    'Main Lounge & Social',
    'F&B / Bar Counter',
    'VIP / Members Area',
    'Live Entertainment / DJ Zone',
    'Amenities (WC)',
    'Back-of-House',
    'Smoking / Outdoor Terrace'
  ],
  'Airport & Transit Lounge': [
    'Welcome & Concierge',
    'Main Lounge & Relaxation',
    'Premium F&B / Bar',
    'Quiet & Wellness Zone',
    'Business & Productivity',
    'Amenities (WC / Shower / Prayer)',
    'Back-of-House',
    'Pre-boarding / Exit'
  ],
  'Private Banking & VIP Lounge': [
    'Reception & Concierge',
    'Private Client Lounge',
    'Meeting / Advisory Rooms',
    'F&B & Refreshment Bar',
    'Executive Board Room',
    'Amenities (WC / Cloakroom)',
    'Back-of-House / Vault Support',
    'Discreet Client Entry / Exit'
  ],
  'Wellness & Spa': [
    'Reception & Welcome Ritual',
    'Relaxation Lounge',
    'Treatment Rooms',
    'Wet Area (Sauna / Steam / Pool)',
    'Yoga & Movement Studio',
    'Amenities (Changing / Lockers)',
    'Back-of-House / Staff',
    'Product Retail & Exit'
  ],
  'Luxury Retail & Showroom': [
    'Storefront & Entry Display',
    'Main Sales Floor',
    'VIP / Private Shopping Suite',
    'Product Feature / Hero Display',
    'Fitting / Consultation Rooms',
    'Amenities (WC / Lounge)',
    'Back-of-House / Stockroom',
    'Staff & Service Entry'
  ],
  'Gallery & Event Space': [
    'Entry & Welcome / Reception',
    'Main Gallery / Exhibition Floor',
    'Featured / Highlight Space',
    'Collector / VIP Lounge',
    'Multi-purpose Event Hall',
    'Amenities (WC)',
    'Back-of-House / Storage',
    'Loading & Handling'
  ],
  'Workplace & Experience Center': [
    'Reception & Brand Welcome',
    'Open Workspace / Collaboration',
    'Experience / Showcase Zone',
    'Meeting & Conference Rooms',
    'Breakout & Café Lounge',
    'Amenities (WC / Wellness Room)',
    'Back-of-House / IT & Storage',
    'Staff Circulation'
  ],
  'Villa & Mansion': [
    'Entry Foyer & Arrival',
    'Living & Family Room',
    'Dining & Kitchen',
    'Master Suite',
    'Guest Bedrooms',
    'Outdoor / Garden & Pool',
    'Back-of-House / Staff Quarters',
    'Garage & Service Entry'
  ],
  'Penthouse & Duplex': [
    'Entry Foyer & Arrival',
    'Living & Social Level',
    'Dining & Open Kitchen',
    'Master Suite & Walk-in',
    'Private Terrace / Sky Lounge',
    'Guest Suite(s)',
    'Home Office / Media Room',
    'Back-of-House / Utility'
  ],
  'Mixed-use & Lifestyle Development': [
    'Main Entry & Arrival Plaza',
    'Retail / F&B Frontage',
    'Lobby & Concierge',
    'Shared Amenity / Clubhouse',
    'Co-working / Community Space',
    'Amenities (WC / Wellness)',
    'Back-of-House / Loading',
    'Vertical Circulation & Lift Core'
  ]
};

const TYPOLOGY_CONTEXT = {
  'Hotel & Resort': 'luxury hotel resort public area interior design zoning programme best practices',
  'F&B Restaurant': 'fine dining restaurant interior design zoning space planning best practices Michelin',
  'Lounge & Club': 'nightlife lounge club interior design zoning space planning best practices',
  'Airport & Transit Lounge': 'premium airport transit lounge design zoning programme best practices',
  'Private Banking & VIP Lounge': 'private banking VIP client lounge interior design zoning best practices',
  'Wellness & Spa': 'luxury spa wellness center interior design zoning space planning',
  'Luxury Retail & Showroom': 'luxury retail flagship showroom interior design zoning space planning',
  'Gallery & Event Space': 'art gallery exhibition event space interior design zoning programme',
  'Workplace & Experience Center': 'corporate workplace brand experience center interior design zoning',
  'Villa & Mansion': 'luxury villa mansion residential interior design space planning',
  'Penthouse & Duplex': 'luxury penthouse duplex residential interior design space planning',
  'Mixed-use & Lifestyle Development': 'mixed-use lifestyle development interior design zoning master planning'
};

// ─── AI RESEARCH ENDPOINT ────────────────────────────────────────
app.post('/api/research', requireAuth, async (req, res) => {
  const { area, typology, mood, location, projectName, description, userApiKey } = req.body;

  const apiKey = (userApiKey && userApiKey.trim()) || process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(400).json({
      error: 'Chưa có API key. Nhập API key của bạn ở góc trên (biểu tượng 🔑) để dùng tính năng AI, hoặc liên hệ quản trị viên để cấu hình key mặc định.',
      needsApiKey: true,
    });
  }

  // Get correct zones for this typology
  const zones = ZONE_TEMPLATES[typology] || ZONE_TEMPLATES['Hotel & Resort'];
  const zoneList = zones.join(', ');
  const searchContext = TYPOLOGY_CONTEXT[typology] || typology;

  try {
    const prompt = `You are an expert hospitality interior designer and space planner at H2X Studio Vietnam, specializing in ${typology} projects.

PROJECT BRIEF:
- Project Name: ${projectName || 'Hospitality Project'}
- Typology: ${typology}
- Location: ${location || 'Vietnam'}
- Total Area: ~${area || '1400'} m²
- Mood/Direction: ${mood || 'Contemporary luxury'}
${description ? `- Additional Description (from client): ${description}\n  → Use this description to refine the zoning and design direction below. If it describes a space that differs from the standard ${typology} template (e.g. a hybrid or unusual programme), let it override generic assumptions — the 8 mandatory zone NAMES below still apply structurally, but re-interpret their sizing, rationale, and design direction around what the client actually described.` : ''}

TASK: Based on best practices for ${typology} interior design (referencing international standards, Michelin-starred restaurants, award-winning hospitality projects), propose a zoning programme and design direction.

MANDATORY ZONES FOR ${typology.toUpperCase()} (use EXACTLY these 8 zones in this order):
${zones.map((z, i) => `${i + 1}. ${z}`).join('\n')}

RULES:
1. All 8 zones must sum to exactly ${area || 1400} m²
2. Area distribution must be realistic for ${typology} (e.g., kitchen = 25-30% for fine dining, main dining = 40-50%)
3. Seats/capacity must be realistic for each zone type (fine dining: 1.8-2.5 m²/cover, lounge: 4-6 m²/person)
4. Rationale must reflect operational best practices specific to ${typology}
5. Design language must suit ${typology} — NOT generic hospitality
6. designPillar1/2/3 must be the THREE most important experiential or operational moments for THIS SPECIFIC typology — chosen freely, not a fixed template. Do NOT default to a generic "arrival → main activity → departure" journey unless that framing genuinely fits the typology (e.g. it fits an airport lounge, but a fine-dining restaurant has no "departure" moment worth a pillar — better choices there might be arrival/hospitality, the dining experience itself, and something else operationally central like the bar, service ritual, or kitchen theatre). Pick whatever three pillars best represent what makes THIS typology's design distinctive.
7. lightingStrategy MUST be a single short sentence (max ~20 words). It renders in a small fixed-size box next to three other one-line bullets — a long or multi-sentence answer will visually overflow and overlap other text on the slide. Keep it as brief as the other bullets, not a technical spec.

Return ONLY this JSON (no markdown, no backticks, no text before/after):

{"zoningProposal":[{"zone":"${zones[0]}","area":150,"seats":20,"rationale":"Specific operational rationale for this zone in ${typology}"},{"zone":"${zones[1]}","area":400,"seats":80,"rationale":"..."},{"zone":"${zones[2]}","area":200,"seats":40,"rationale":"..."},{"zone":"${zones[3]}","area":120,"seats":20,"rationale":"..."},{"zone":"${zones[4]}","area":80,"seats":10,"rationale":"..."},{"zone":"${zones[5]}","area":80,"seats":0,"rationale":"..."},{"zone":"${zones[6]}","area":250,"seats":0,"rationale":"..."},{"zone":"${zones[7]}","area":120,"seats":0,"rationale":"..."}],"designTagline":"Compelling one-line aesthetic direction for this ${typology}","designPillar1":{"label":"LABEL SPECIFIC TO THIS TYPOLOGY'S MOST IMPORTANT MOMENT","sub":"3 EVOCATIVE WORDS"},"designPillar2":{"label":"LABEL FOR THE SECOND MOST IMPORTANT MOMENT","sub":"3 EVOCATIVE WORDS"},"designPillar3":{"label":"LABEL FOR THE THIRD MOST IMPORTANT MOMENT","sub":"3 EVOCATIVE WORDS"},"lightingStrategy":"ONE short sentence (max ~20 words) — a specific lighting technique for ${typology}, same length as a single bullet point, NOT a multi-sentence explanation","opportunity":"2-3 sentences about why this ${typology} project matters in this market","experienceIntent":"2-3 sentences describing the guest journey specific to ${typology}","whyH2X":"2 sentences about H2X Studio expertise in ${typology} projects","feeRangeLow":${Math.round(parseFloat(area||1400) * 80)},"feeRangeHigh":${Math.round(parseFloat(area||1400) * 120)}}`;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 8192,
        tools: [{
          type: 'web_search_20250305',
          name: 'web_search'
        }],
        messages: [{
          role: 'user',
          content: `First search for "${searchContext}" to get current best practices, then use that research to answer:\n\n${prompt}`
        }]
      })
    });

    const data = await response.json();
    console.log('Response type:', data.type, '| Stop reason:', data.stop_reason);

    if (data.error) {
      // Fallback: try without web search
      console.log('Web search failed, trying without:', data.error.message);
      const response2 = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify({
          model: 'claude-sonnet-4-6',
          max_tokens: 8192,
          messages: [{ role: 'user', content: prompt }]
        })
      });
      const data2 = await response2.json();
      if (data2.error) return res.status(500).json({ error: data2.error.message });
      const text2 = data2.content?.[0]?.text || '{}';
      const clean2 = text2.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      const parsed2 = JSON.parse(clean2);
      parsed2._zoneTemplate = zones;
      parsed2._typology = typology;
      return res.json(parsed2);
    }

    // Extract text from response (may include tool_use blocks)
    let text = '';
    for (const block of (data.content || [])) {
      if (block.type === 'text') text += block.text;
    }

    if (!text) text = '{}';
    console.log('Text preview:', text.substring(0, 150));

    if (data.stop_reason === 'max_tokens') {
      console.warn('WARNING: Response truncated by max_tokens limit — JSON is likely incomplete. Consider raising max_tokens further.');
    }

    const clean = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    // Extract JSON from text if mixed with prose
    const jsonMatch = clean.match(/\{[\s\S]*\}/);
    const jsonStr = jsonMatch ? jsonMatch[0] : clean;

    let parsed;
    try {
      parsed = JSON.parse(jsonStr);
    } catch (e) {
      const truncated = data.stop_reason === 'max_tokens';
      console.error('Parse error:', e.message, truncated ? '(caused by max_tokens truncation)' : '');
      return res.status(500).json({
        error: truncated
          ? 'AI response was cut off (max_tokens too low). Increase max_tokens in server.js.'
          : 'Parse failed',
        raw: text.substring(0, 300)
      });
    }

    // Ensure zones match the template
    parsed._zoneTemplate = zones;
    parsed._typology = typology;

    res.json(parsed);

  } catch (err) {
    console.error('Research error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── GENERATE PPTX ENDPOINT ──────────────────────────────────────
app.post('/api/generate', requireAuth, async (req, res) => {
  try {
    const data = req.body;
    console.log('Generating proposal for:', data.projectName);
    const filePath = await generateProposal(data);
    const fileName = path.basename(filePath);
    res.download(filePath, fileName, (err) => {
      if (err) console.error('Download error:', err);
      try { fs.unlinkSync(filePath); } catch (e) {}
    });
  } catch (err) {
    console.error('Generation error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── ZONE TEMPLATE ENDPOINT ──────────────────────────────────────
app.get('/api/zones/:typology', (req, res) => {
  const typology = decodeURIComponent(req.params.typology);
  const zones = ZONE_TEMPLATES[typology] || ZONE_TEMPLATES['Hotel & Resort'];
  res.json({ zones });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`H2X Proposal Generator running on port ${PORT}`);
});
