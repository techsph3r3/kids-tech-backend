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

app.get('/api/pages', async (req, res) => {
  try {
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
    const accessToken = tokenData.access_token;

    const notebooksRes = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks', {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const notebooksData = await notebooksRes.json();
    const targetNotebook = notebooksData.value.find(nb => nb.displayName === NOTEBOOK_NAME);

    if (!targetNotebook) return res.status(404).json({ error: 'Notebook not found' });

    const sectionsRes = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/notebooks/${targetNotebook.id}/sections`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const sectionsData = await sectionsRes.json();
    const targetSection = sectionsData.value.find(sec => sec.displayName === SECTION_NAME);

    if (!targetSection) return res.status(404).json({ error: 'Section not found' });

    const pagesRes = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/sections/${targetSection.id}/pages`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const pagesData = await pagesRes.json();

    const pages = (pagesData.value || []).map(page => ({
      id: page.id,
      title: page.title,
      url: page.links?.oneNoteWebUrl?.href || null,
    }));

    res.json(pages);
  } catch (err) {
    console.error('Error:', err);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.listen(port, () => {
  console.log(`✅ Server running on port ${port}`);
});
