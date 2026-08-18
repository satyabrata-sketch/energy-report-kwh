// Vercel Serverless Function: Cloud Sync API for Multi-Device Real-Time Synchronization
let memoryStore = {};

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Credentials', true);
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader(
    'Access-Control-Allow-Headers',
    'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version'
  );

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  const room = (req.query.room || req.body?.room || 'CBRE-DT3-FACILITY-2026').trim();

  if (req.method === 'GET') {
    const data = memoryStore[room] || null;
    return res.status(200).json({
      status: 'success',
      room: room,
      data: data ? data.payload : null,
      lastUpdated: data ? data.lastUpdated : null,
      version: data ? data.version : 0
    });
  }

  if (req.method === 'POST') {
    const body = req.body || {};
    const payload = body.data || body;
    const version = body.version || Date.now();
    const lastUpdated = new Date().toISOString();

    memoryStore[room] = {
      payload: payload,
      version: version,
      lastUpdated: lastUpdated
    };

    return res.status(200).json({
      status: 'success',
      room: room,
      version: version,
      lastUpdated: lastUpdated
    });
  }

  res.status(405).json({ error: 'Method not allowed' });
}
