const express = require('express');
const path = require('path');
const fs = require('fs');
const { generateProposal } = require('./generatePptx');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ─── AI RESEARCH ENDPOINT ────────────────────────────────────────
app.post('/api/research', async (req, res) => {
  const { area, typology, mood, location, projectName } = req.body;
  try {
    const prompt = `You are an expert hospitality interior designer helping H2X Studio in Vietnam prepare a fee proposal for a client.

PROJECT:
- Name: ${projectName || 'Premium Hospitality Project'}
- Typology: ${typology || 'Premium Lounge'}
- Location: ${location || 'Vietnam'}
- Total Area: ~${area || '1400'} m²
- Mood/Direction: ${mood || 'Luxury Vietnamese contemporary'}

Based on this, provide structured recommendations in JSON format ONLY (no markdown, no backticks):

{
  "zoningProposal": [
    {"zone": "Zone name", "area": 200, "areaPercent": 14, "seats": 30, "rationale": "Operational rationale in English"}
  ],
  "designTagline": "One-line aesthetic tagline in English",
  "designPillar1": {"label": "ARRIVAL", "sub": "3 DESCRIPTIVE WORDS"},
  "designPillar2": {"label": "F&B", "sub": "3 DESCRIPTIVE WORDS"},
  "designPillar3": {"label": "SOCIAL", "sub": "3 DESCRIPTIVE WORDS"},
  "lightingStrategy": "One sentence describing zoned lighting approach for this typology",
  "opportunity": "2-3 sentences in English about market opportunity and project significance",
  "experienceIntent": "2-3 sentences in English describing the guest journey and emotional experience",
  "whyH2X": "2 sentences in English about why H2X is uniquely suited for this project",
  "benchmarkNote": "Fee benchmark note e.g. ≈ 107 USD/m² · 1,400 m²",
  "feeRangeLow": 120000,
  "feeRangeHigh": 160000
}

Zones must sum to approximately ${area || 1400} m². Include 8 zones: Welcome/Entry, Main Social/Lounge, F&B/Bar, Quiet/Wellness, Business/Productivity, Amenities, Back-of-House, Transition/Exit. The last zone should have remaining area. Return ONLY valid JSON.`;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 1500,
        messages: [{ role: 'user', content: prompt }]
      })
    });

    const data = await response.json();
    const text = data.content?.[0]?.text || '{}';
    
    let parsed;
    try {
      parsed = JSON.parse(text.replace(/```json|```/g, '').trim());
    } catch {
      parsed = { error: 'Parse failed', raw: text.substring(0, 200) };
    }
    res.json(parsed);
  } catch (err) {
    console.error('Research error:', err);
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
      try { fs.unlinkSync(filePath); } catch(e) {}
    });
  } catch (err) {
    console.error('Generation error:', err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`H2X Proposal Generator → http://localhost:${PORT}`);
});
