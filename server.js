const express = require('express');
const path = require('path');
const fs = require('fs');
const { generateProposal } = require('./generatePptx');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ─── TYPOLOGY ZONE TEMPLATES ─────────────────────────────────────
const ZONE_TEMPLATES = {
  'Premium / Business Airport Lounge': [
    'Welcome & Concierge',
    'Main Lounge & Relaxation',
    'Premium F&B / Bar',
    'Quiet & Wellness Zone',
    'Business & Productivity',
    'Amenities (WC / Shower / Prayer)',
    'Back-of-House',
    'Pre-boarding / Exit'
  ],
  'F&B Fine Dining': [
    'Entry & Reception / Host Stand',
    'Main Dining Room',
    'Private Dining Room (VIP)',
    'Bar & Pre-dining Lounge',
    'Wine Cellar / Display',
    'Amenities (WC / Cloakroom)',
    'Back-of-House / Kitchen',
    'Staff & Service Corridor'
  ],
  'F&B Casual-Premium': [
    'Entry & Waiting Area',
    'Main Dining Floor',
    'Bar & Social Counter',
    'Outdoor / Terrace Dining',
    'Private / Semi-private Section',
    'Amenities (WC)',
    'Back-of-House / Kitchen',
    'Storage & Staff Area'
  ],
  'Gallery / Exhibition Space': [
    'Entry & Welcome / Reception',
    'Main Gallery Floor',
    'Featured / Highlight Gallery',
    'Collector Lounge & VIP Area',
    'Education & Workshop Space',
    'Amenities (WC)',
    'Back-of-House / Storage',
    'Loading & Art Handling'
  ],
  'Boutique Hotel Public Area': [
    'Lobby & Check-in',
    'Lounge & Social Area',
    'F&B / All-day Dining',
    'Bar & Evening Lounge',
    'Fitness & Wellness',
    'Amenities (WC / Concierge)',
    'Back-of-House',
    'Arrival / Departure Zone'
  ],
  'Wellness / Spa': [
    'Reception & Welcome Ritual',
    'Relaxation Lounge',
    'Treatment Rooms',
    'Wet Area (Sauna / Steam / Pool)',
    'Yoga & Movement Studio',
    'Amenities (Changing / Lockers)',
    'Back-of-House / Staff',
    'Product Retail & Exit'
  ]
};

const TYPOLOGY_CONTEXT = {
  'Premium / Business Airport Lounge': 'premium airport business lounge design zoning programme best practices 2024',
  'F&B Fine Dining': 'fine dining restaurant interior design zoning space planning best practices Michelin',
  'F&B Casual-Premium': 'casual premium restaurant interior design space planning zones layout',
  'Gallery / Exhibition Space': 'art gallery exhibition space interior design zoning programme',
  'Boutique Hotel Public Area': 'boutique hotel lobby lounge public area interior design zoning',
  'Wellness / Spa': 'luxury spa wellness center interior design zoning space planning'
};

// ─── AI RESEARCH ENDPOINT ────────────────────────────────────────
app.post('/api/research', async (req, res) => {
  const { area, typology, mood, location, projectName } = req.body;

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY not configured' });
  }

  // Get correct zones for this typology
  const zones = ZONE_TEMPLATES[typology] || ZONE_TEMPLATES['Premium / Business Airport Lounge'];
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

TASK: Based on best practices for ${typology} interior design (referencing international standards, Michelin-starred restaurants, award-winning hospitality projects), propose a zoning programme and design direction.

MANDATORY ZONES FOR ${typology.toUpperCase()} (use EXACTLY these 8 zones in this order):
${zones.map((z, i) => `${i + 1}. ${z}`).join('\n')}

RULES:
1. All 8 zones must sum to exactly ${area || 1400} m²
2. Area distribution must be realistic for ${typology} (e.g., kitchen = 25-30% for fine dining, main dining = 40-50%)
3. Seats/capacity must be realistic for each zone type (fine dining: 1.8-2.5 m²/cover, lounge: 4-6 m²/person)
4. Rationale must reflect operational best practices specific to ${typology}
5. Design language must suit ${typology} — NOT generic hospitality

Return ONLY this JSON (no markdown, no backticks, no text before/after):

{"zoningProposal":[{"zone":"${zones[0]}","area":150,"seats":20,"rationale":"Specific operational rationale for this zone in ${typology}"},{"zone":"${zones[1]}","area":400,"seats":80,"rationale":"..."},{"zone":"${zones[2]}","area":200,"seats":40,"rationale":"..."},{"zone":"${zones[3]}","area":120,"seats":20,"rationale":"..."},{"zone":"${zones[4]}","area":80,"seats":10,"rationale":"..."},{"zone":"${zones[5]}","area":80,"seats":0,"rationale":"..."},{"zone":"${zones[6]}","area":250,"seats":0,"rationale":"..."},{"zone":"${zones[7]}","area":120,"seats":0,"rationale":"..."}],"designTagline":"Compelling one-line aesthetic direction for this ${typology}","designPillar1":{"label":"ARRIVAL","sub":"3 EVOCATIVE WORDS"},"designPillar2":{"label":"DINING","sub":"3 EVOCATIVE WORDS"},"designPillar3":{"label":"DEPARTURE","sub":"3 EVOCATIVE WORDS"},"lightingStrategy":"Specific lighting strategy for ${typology} — reference actual techniques","opportunity":"2-3 sentences about why this ${typology} project matters in this market","experienceIntent":"2-3 sentences describing the guest journey specific to ${typology}","whyH2X":"2 sentences about H2X Studio expertise in ${typology} projects","feeRangeLow":${Math.round(parseFloat(area||1400) * 80)},"feeRangeHigh":${Math.round(parseFloat(area||1400) * 120)}}`;

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
      return res.json(JSON.parse(clean2));
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
app.post('/api/generate', async (req, res) => {
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
  const zones = ZONE_TEMPLATES[typology] || ZONE_TEMPLATES['Premium / Business Airport Lounge'];
  res.json({ zones });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`H2X Proposal Generator running on port ${PORT}`);
});
