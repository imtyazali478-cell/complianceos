export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-api-key, anthropic-version');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).end();

  const key = process.env.ANTHROPIC_KEY;
  if (!key) return res.status(500).json({ 
    content: [{type:'text', text:'API key missing'}] 
  });

  try {const finalBody = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
finalBody.model = 'claude-haiku-4-5-20251001';
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': key,
        'anthropic-version': '2023-06-01'
      },
     body: JSON.stringify(finalBody)
    });
    
    if (!response.ok) {
      const err = await response.text();
      return res.status(200).json({ 
        content: [{type:'text', text:'API Error: ' + err}] 
      });
    }
    
    const data = await response.json();
    return res.status(200).json(data);
  } catch (e) {
    return res.status(200).json({ 
      content: [{type:'text', text:'Error: ' + e.message}] 
    });
  }
}
