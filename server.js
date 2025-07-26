const express = require('express');
const multer = require('multer');
const path = require('path');
const { exec } = require('child_process');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Serve static files (HTML)
app.use(express.static('public'));

// Multer config for file uploads
const upload = multer({ dest: 'uploads/' });

// Handle file upload
app.post('/upload', upload.single('file'), (req, res) => {
    const uploadedFile = req.file;
    const tempPath = uploadedFile.path;
    const originalExt = path.extname(uploadedFile.originalname).toLowerCase();
    const newPath = tempPath + originalExt;

    // Rename with original extension
    fs.rename(tempPath, newPath, (err) => {
        if (err) return res.status(500).send('File rename failed.');

        // Call Python script
        const outputPath = path.join(__dirname, 'output.xlsx');
        const command = `python scripts/convert.py "${newPath}" "${outputPath}"`;


        exec(command, (error, stdout, stderr) => {
            if (error) {
                console.error('Python Error:', stderr);
                return res.status(500).send('Python script failed.');
            }

            // Send the resulting Excel file
            const outputPath = path.join(__dirname, 'output.xlsx');
            res.download(outputPath, 'placemarks_output.xlsx', (err) => {
                if (err) {
                    console.error('Download error:', err);
                }
                // Optionally clean up
                fs.unlink(newPath, () => {});
                fs.unlink(outputPath, () => {});
            });
        });
    });
});

app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
});
