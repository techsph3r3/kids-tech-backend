const axios = require("axios");

module.exports = async (req, res) => {
  // ✅ Handle CORS
  res.setHeader("Access-Control-Allow-Origin", "https://kids-tech-frontend.vercel.app");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    return res.status(200).end(); // ✅ Preflight handled
  }

  const authHeader = req.headers.authorization;

  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    return res.status(401).json({ error: "Missing or invalid auth token" });
  }

  const accessToken = authHeader.split(" ")[1];

  try {
    const graphResponse = await axios.get(
      "https://graph.microsoft.com/v1.0/me/onenote/notebooks",
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );

    res.status(200).json(graphResponse.data);
} catch (error) {
  const errorData = error?.response?.data || error.message || "Unknown error";
  console.error("Graph API Error:", errorData);

  res.status(500).json({
    error: "Failed to fetch OneNote data",
    details: errorData,
  });
}
};
