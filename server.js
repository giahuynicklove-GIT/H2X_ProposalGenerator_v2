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
  
  const apiKey = process.env.ANTHROPIC_API_KEY;
  console.log('API Key exists:', !!apiKey);
  console.log('API Key prefix:', apiKey ? apiKey.substring(0, 15) + '...' : 'MISSING');
  
  try {
    const prompt = `You are an expert hospitality interior designer helping H2X Studio in Vietnam prepare a fee proposal for a client.

PROJECT:
- Name: ${projectName || 'Premium Hospitality Project'}
- Typology: ${typology || 'Premium Lounge'}
- Location: ${location || 'Vietnam'}
- Total Area: ~${area || '1400'} m²
- Mood/Direction: ${mood || 'Luxury Vietnamese contemporary'}

Based on this, provide structured recommendations in JSON format ONLY (no markdown, no backticks, no explanation):

{"zoningProposal":[{"zone":"Zone name","area":200,"seats":30,"rationale":"Operational rationale in English"}],"designTagline":"One-line aesthetic tagline in English","designPillar1":{"label":"ARRIVAL","sub":"3 DESCRIPTIVE WORDS"},"designPillar2":{"label":"F&B","sub":"3 DESCRIPTIVE WORDS"},"designPillar3":{"label":"SOCIAL","sub":"3 DESCRIPTIVE WORDS"},"lightingStrategy":"One sentence describing zoned lighting approach","opportunity":"2-3 sentences about market opportunity","experienceIntent":"2-3 sentences about guest journey","whyH2X":"2 sentences about why H2X suits this project","feeRangeLow":120000,"feeRangeHigh":160000}

Include exactly 8 zones summing to approximately ${area || 1400} m²: Welcome/Entry, Main Social/Lounge, F&B/Bar, Quiet/Wellness, Business/Productivity, Amenities, Back-of-House, Transition/Exit. Return ONLY valid JSON, nothing else.`;

    console.log('Calling Anthropic API...');
    
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 2000,
        messages: [{ role: 'user', content: prompt }]
      })
    });

    console.log('Anthropic status:', response.status);
    
    const data = await response.json();
    console.log('Anthropic response type:', data.type);
    console.log('Anthropic response preview:', JSON.stringify(data).substring(0, 300));
    
    if (data.error) {
      console.error('Anthropic error:', data.error);
      return res.status(500).json({ error: data.error.message || 'Anthropic API error' });
    }
    
    const text = data.content?.[0]?.text || '{}';
    console.log('Text preview:', text.substring(0, 200));
    
    let parsed;
    try {
      const clean = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      parsed = JSON.parse(clean);
    } catch (parseErr) {
      console.error('JSON parse error:', parseErr.message);
      console.error('Raw text:', text.substring(0, 500));
      return res.status(500).json({ error: 'Failed to parse AI response', raw: text.substring(0, 200) });
    }
    
    console.log('Success! Returning parsed data');
    res.json(parsed);
    
  } catch (err) {
    console.error('Research error:', err.message);
    console.error('Stack:', err.stack);
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
