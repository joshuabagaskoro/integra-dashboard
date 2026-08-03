// api/parse-loading-report.js
// Serverless function to parse loading reports via Claude API

const Anthropic = require('@anthropic-ai/sdk');

const client = new Anthropic();

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { chatText } = req.body;

    if (!chatText) {
      return res.status(400).json({ error: 'chatText required' });
    }

    // Call Claude API to parse the loading report
    const message = await client.messages.create({
      model: 'claude-opus-4-6',
      max_tokens: 1024,
      messages: [
        {
          role: 'user',
          content: `You are a loading report parser. Extract the following from this shipment loading report and return ONLY valid JSON (no markdown, no extra text):

{
  "shipmentNumber": <integer>,
  "qtyOnBoard": <number in WMT>,
  "progressPercent": <number 0-100>,
  "balanceDue": <number in WMT>,
  "reportDate": "YYYY-MM-DD"
}

If any field is missing, set to null. Return ONLY the JSON object.

Loading Report:
${chatText}`
        }
      ]
    });

    // Extract the response text
    const responseText = message.content[0].type === 'text' 
      ? message.content[0].text 
      : '';

    // Parse JSON response
    let parsed;
    try {
      parsed = JSON.parse(responseText);
    } catch (e) {
      return res.status(400).json({ 
        error: 'Failed to parse Claude response', 
        details: responseText 
      });
    }

    res.status(200).json(parsed);
  } catch (error) {
    console.error('Claude parse error:', error);
    res.status(500).json({ 
      error: 'Failed to parse loading report', 
      details: error.message 
    });
  }
}
