const axios = require("axios");

module.exports = async (req, res) => {
  const authHeader = req.headers.authorization;

  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    return res.status(401).json({ error: "Missing or invalid auth token" });
  }

  const accessToken = authHeader.split(" ")[1];

  try {
    const graphResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/me/onenote/notebooks`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );

    res.status(200).json(graphResponse.data);
  } catch (error) {
    console.error("Graph API Error:", error?.response?.data || error.message);
    res.status(500).json({
      error: "Failed to fetch OneNote data",
      details: error?.response?.data || error.message,
    });
  }
};

