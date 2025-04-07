import express from 'express';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
dotenv.config();

const app = express();
const port = process.env.PORT || 3000;

// Allow frontend to call backend
app.use(express.json());
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*'); // For testing; restrict later
  res.setHeader('Access-Control-Allow-Headers', 'Authorization, Content-Type');
  next();
});

app.get('/', (req, res) => {
  res.send('✅ Backend is running with delegated auth!');
});

app.get('/api/pages', async (req, res) => {
  const authHeader = req.headers.authorization;

  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return res.status(401).json({ error: 'Missing or invalid Authorization header' });
  }

  const accessToken = authHeader.split(' ')[1];

  try {
    // Get notebooks
    const notebooksRes = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks', {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });

    const notebooks = await notebooksRes.json();
    const targetNotebook = notebooks.value.find(nb => nb.displayName === "Tech2025_2026");

    if (!targetNotebook) {
      return res.status(404).json({ error: 'Notebook not found' });
    }

    // Get sections
    const sectionsRes = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/notebooks/${targetNotebook.id}/sections`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const sections = await sectionsRes.json();
    const targetSection = sections.value.find(sec => sec.displayName === "Quarter 1");

    if (!targetSection) {
      return res.status(404).json({ error: 'Section not found' });
    }

    // Get pages
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
  } catch (error) {
    console.error("❌ Error fetching data:", error);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.listen(port, () => {
  console.log(`✅ Server running on port ${port}`);
});
