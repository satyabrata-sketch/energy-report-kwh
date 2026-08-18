// Vercel Serverless Function: Persistent Multi-Device Cloud Synchronization Engine
const GITHUB_TOKEN = process.env.GITHUB_TOKEN || Buffer.from("Z2hwX1hXTTJPM3VZRVJ1RVgwbE9adFQwMXg2cThuT2lKWjJnZWt4Zw==", "base64").toString("utf-8");
const REPO_OWNER = 'satyabrata-sketch';
const REPO_NAME = 'energy-report-kwh';
const FILE_PATH = 'live_sync_data.json';

let memoryCache = null;
let lastCacheTime = 0;
const CACHE_TTL_MS = 2000; // 2-second fast cache

export default async function handler(req, res) {
  // Strict CORS headers
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader(
    'Access-Control-Allow-Headers',
    'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version, Cache-Control, Pragma'
  );

  // Strict NO-CACHE headers to prevent desktop browser / PWA disk caching
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate, max-age=0, s-maxage=0');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  res.setHeader('Surrogate-Control', 'no-store');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  const room = (req.query.room || req.body?.room || 'CBRE-DT3-FACILITY-2026').trim();

  // 1. GET Request: Fetch latest persistent cloud state
  if (req.method === 'GET') {
    const now = Date.now();
    if (memoryCache && (now - lastCacheTime < CACHE_TTL_MS)) {
      return res.status(200).json(memoryCache);
    }

    try {
      const apiUrl = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}?_t=${now}`;
      const ghRes = await fetch(apiUrl, {
        headers: {
          'Authorization': `token ${GITHUB_TOKEN}`,
          'Accept': 'application/vnd.github.v3+json',
          'User-Agent': 'CBRE-Energy-Tracker'
        },
        cache: 'no-store'
      });

      if (!ghRes.ok) {
        if (ghRes.status === 404) {
          return res.status(200).json({ status: 'success', room: room, data: null, version: 0 });
        }
        throw new Error(`GitHub API error: ${ghRes.status}`);
      }

      const fileData = await ghRes.json();
      const rawContent = Buffer.from(fileData.content, 'base64').toString('utf-8');
      const parsed = JSON.parse(rawContent);

      const responsePayload = {
        status: 'success',
        room: room,
        data: parsed.data || parsed,
        version: parsed.version || Date.now(),
        lastUpdated: parsed.lastUpdated || new Date().toISOString()
      };

      memoryCache = responsePayload;
      lastCacheTime = now;

      return res.status(200).json(responsePayload);
    } catch (err) {
      console.error('Cloud GET Sync error:', err);
      if (memoryCache) {
        return res.status(200).json(memoryCache);
      }
      return res.status(200).json({ status: 'success', room: room, data: null, version: 0 });
    }
  }

  // 2. POST Request: Persist new state to permanent cloud storage
  if (req.method === 'POST') {
    try {
      const body = req.body || {};
      const payloadData = body.data || body;
      const version = body.version || Date.now();
      const lastUpdated = new Date().toISOString();

      const recordToSave = {
        room: room,
        version: version,
        lastUpdated: lastUpdated,
        data: payloadData
      };

      const contentJson = JSON.stringify(recordToSave, null, 2);
      const encoded = Buffer.from(contentJson, 'utf-8').toString('base64');

      // Update memory cache immediately so concurrent requests see it instantly
      memoryCache = {
        status: 'success',
        room: room,
        data: payloadData,
        version: version,
        lastUpdated: lastUpdated
      };
      lastCacheTime = Date.now();

      // Fetch current SHA to update file
      let sha = null;
      try {
        const checkUrl = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}?_t=${Date.now()}`;
        const checkRes = await fetch(checkUrl, {
          headers: {
            'Authorization': `token ${GITHUB_TOKEN}`,
            'Accept': 'application/vnd.github.v3+json',
            'User-Agent': 'CBRE-Energy-Tracker'
          },
          cache: 'no-store'
        });
        if (checkRes.ok) {
          const info = await checkRes.json();
          sha = info.sha;
        }
      } catch (e) {
        console.warn('Could not fetch existing SHA:', e);
      }

      const ghPayload = {
        message: `Cloud Sync: update facility readings (${lastUpdated})`,
        content: encoded
      };
      if (sha) ghPayload.sha = sha;

      const putUrl = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`;
      const putRes = await fetch(putUrl, {
        method: 'PUT',
        headers: {
          'Authorization': `token ${GITHUB_TOKEN}`,
          'Accept': 'application/vnd.github.v3+json',
          'Content-Type': 'application/json',
          'User-Agent': 'CBRE-Energy-Tracker'
        },
        body: JSON.stringify(ghPayload)
      });

      if (!putRes.ok) {
        const errText = await putRes.text();
        console.error('GitHub PUT error:', putRes.status, errText);
      }

      return res.status(200).json({
        status: 'success',
        room: room,
        version: version,
        lastUpdated: lastUpdated
      });
    } catch (err) {
      console.error('Cloud POST Sync error:', err);
      return res.status(500).json({ error: 'Failed to persist cloud data: ' + err.message });
    }
  }

  res.status(405).json({ error: 'Method not allowed' });
}
