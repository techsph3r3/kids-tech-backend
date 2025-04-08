const axios = require("axios");

module.exports = async (req, res) => {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const notebookId = process.env.ONENOTE_NOTEBOOK_ID;
  const userId = process.env.MS_USER_ID;

  if (!tenantId || !clientId || !clientSecret || !notebookId || !userId) {
    return res.status(400).json({
      error: "Missing required environment variables",
      variables: {
        TENANT_ID: !!tenantId,
        CLIENT_ID: !!clientId,
        CLIENT_SECRET: !!clientSecret,
        ONENOTE_NOTEBOOK_ID: !!notebookId,
        MS_USER_ID: !!userId,
      },
    });
  }

  try {
    // Step 1: Get access token from Microsoft Identity Platform
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

    // Step 2: Use token to call Microsoft Graph API
    const graphResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${userId}/onenote/notebooks/${notebookId}/sections`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );

    // Step 3: Return OneNote section data
    res.status(200).json(graphResponse.data);
  } catch (error) {
    console.error("Graph API Error:", error?.response?.data || error.message);
    res.status(500).json({
      error: "Failed to fetch OneNote data",
      details: error?.response?.data || error.message,
    });
  }
};
