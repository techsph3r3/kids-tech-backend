const axios = require("axios");

module.exports = async (req, res) => {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const notebookId = process.env.ONENOTE_NOTEBOOK_ID;

  try {
    // STEP 1: Get token using client_credentials
    const tokenResponse = await axios.post(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
      }),
      {
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
      }
    );

    const accessToken = tokenResponse.data.access_token;

    // STEP 2: Call OneNote API with token
    const graphResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${process.env.MS_USER_ID}/onenote/notebooks/${notebookId}/sections`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );

    res.status(200).json(graphResponse.data);
  } catch (error) {
    console.error("Graph Error", error?.response?.data || error.message);
    res.status(500).json({ error: "Failed to fetch OneNote data" });
  }
};
