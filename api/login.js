export default function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { password } = req.body || {};
  const appPassword = process.env.APP_PASSWORD;

  if (!appPassword) {
    return res.status(500).json({ error: 'Server misconfigured' });
  }

  if (password === appPassword) {
    return res.status(200).json({ success: true });
  }

  return res.status(401).json({ success: false, error: 'Incorrect password' });
}
