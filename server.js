import express from 'express';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
dotenv.config();

const app = express();
const port = process.env.PORT || 3000;

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const NOTEBOOK_NAME = "Tech2025_2026";
const SECTION_NAME = "Quarter 1";
const USER_ID = "6674f220-79cb-429b-986f-e88f53d48a91";

app.get('/', (req, res) => {
  res.send('✅ Kids Tech Backend is running!');
});

app.get('/api/pages', async (req, res) => {
  try {
    console.log("🔐 Fetching access token...");
    const tokenRes = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: CLIENT_ID,
        scope: 'https://graph.microsoft.com/.default',
        client_secret: CLIENT_SECRET,
        grant_type: 'client_credentials',
      }),
    });

    const tokenData = await tokenRes.json();
    if (!tokenData.access_token) {
      console.error("❌ Failed to get access token:", tokenData);
      return res.status(500).json({ error: 'Failed to get token', details: tokenData });
    }

    const accessToken = tokenData.access_token;
    console.log("✅ Got token, fetching notebooks...");

    const notebooksRes = await fetch(`https://graph.microsoft.com/v1.0/users/${USER_ID}/onenote/notebooks`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const notebooksData = await notebooksRes.json();

    if (!notebooksData.value) {
      console.error("❌ Notebooks fetch failed:", notebooksData);
      return res.status(500).json({ error: 'Notebook fetch failed', details: notebooksData });
    }

    const targetNotebook = notebooksData.value.find(nb => nb.displayName === NOTEBOOK_NAME);
    if (!targetNotebook) {
      console.warn("❗ Notebook not found");
      return res.status(404).json({ error: 'Notebook not found' });
    }

    console.log("✅ Notebook found:", targetNotebook.id);

    const sectionsRes = await fetch(`https://graph.microsoft.com/v1.0/users/${USER_ID}/onenote/notebooks/${targetNotebook.id}/sections`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const sectionsData = await sectionsRes.json();

    if (!sectionsData.value) {
      console.error("❌ Sections fetch failed:", sectionsData);
      return res.status(500).json({ error: 'Sections fetch failed', details: sectionsData });
    }

    const targetSection = sectionsData.value.find(sec => sec.displayName === SECTION_NAME);
    if (!targetSection) {
      console.warn("❗ Section not found");
      return res.status(404).json({ error: 'Section not found' });
    }

    console.log("✅ Section found:", targetSection.id);

    const pagesRes = await fetch(`https://graph.microsoft.com/v1.0/users/${USER_ID}/onenote/sections/${targetSection.id}/pages`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const pagesData = await pagesRes.json();

    if (!pagesData.value) {
      console.error("❌ Pages fetch failed:", pagesData);
      return res.status(500).json({ error: 'Pages fetch failed', details: pagesData });
    }

    console.log("✅ Pages fetched:", pagesData.value.length);
    const pages = pagesData.value.map(page => ({
      id: page.id,
      title: page.title,
      url: page.links?.oneNoteWebUrl?.href || null,
    }));

    res.json(pages);
  } catch (err) {
    console.error("❌ UNEXPECTED ERROR:", err);
    res.status(500).json({ error: 'Internal Server Error', message: err.message });
  }
});

app.listen(port, () => {
  console.log(`✅ Server running on port ${port}`);
});
