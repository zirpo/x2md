const express = require('express');
const multer = require('multer');
const path = require('path');
const converter = require('./services/converter');

const app = express();
const port = 3001;

// Configure multer for memory storage
const upload = multer({ storage: multer.memoryStorage() });

// Serve static files from 'public' directory
app.use(express.static('public'));

// API endpoint for file conversion
app.post('/api/convert', upload.single('file'), async (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'No file uploaded' });
    }

    try {
        const markdown = await converter.convert(
            req.file.buffer,
            req.file.mimetype,
            req.file.originalname
        );
        res.json({ markdown });
    } catch (error) {
        console.error('Conversion error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.listen(port, () => {
    console.log(`Server listening at http://localhost:${port}`);
});
