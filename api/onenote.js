const axios = require("axios");

module.exports = async (req, res) => {
  // ✅ CORS headers to allow frontend to reach this backend
  res.setHeader("Access-Control-Allow-Origin", "https://kids-tech-frontend.vercel.app");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  const userAccessToken = req.headers.authorization?.split(" ")[1];

  if (!userAccessToken) {
    return res.status(401).json({ error: "Missing Authorization header" });
  }

  // Get env vars
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const backendUserId = process.env.BACKEND_USER_ID;
  const notebookId = process.env.ONENOTE_NOTEBOOK_ID;

  try {
    // ✅ Step 1: Exchange token using OBO flow
    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const tokenResponse = await axios.post(
      tokenUrl,
      new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
        requested_token_use: "on_behalf_of",
        scope: "https://graph.microsoft.com/.default",
        assertion: userAccessToken,
      }),
      {
        headers: {
          "Content-Type": "

