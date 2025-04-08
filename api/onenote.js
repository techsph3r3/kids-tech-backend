const axios = require("axios");

module.exports = async (req, res) => {
  // ✅ Allow frontend calls from Vercel
  res.setHeader("Access-Control-Allow-Origin", "https://kids-tech-frontend.vercel.app");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  // ✅ Handle preflight request
  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  const userAccessToken = req.headers.authorization?.split(" ")[1];

  if (!userAccessToken) {
    return res.status(401).json({ error: "Missing Authorization header" });
  }

  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const backendUserId = process.env.BACKEND_USER_ID;
  const notebookId = process.env.ONENOTE_NOTEBOOK_ID;

  try {
    // Step 1: Exchange token using OBO flow
    const tokenResponse = await axios.post(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
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
          "Content-Type": "application/x-www-form-urlencoded",
        },
      }
    );

    const oboAccessToken = tokenResponse.data.access_token;

    // Step 2: Use token to call Graph on behalf of fixed backend account
    const graphResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${backendUserId}/onenote/notebooks/${notebookId}/sections`,
      {
        headers: {
          Authorization: `Bearer ${oboAccessToken}`,
        },
      }
    );

    return res.status(200).json(graphResponse.data);
  } catch (error) {
    const details
