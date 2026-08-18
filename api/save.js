// Vercel Serverless Function: Data Save Endpoint
export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method === 'POST') {
    return res.status(200).json({
      status: 'success',
      timestamp: new Date().toISOString()
    });
  }

  res.status(405).json({ error: 'Method not allowed' });
}
