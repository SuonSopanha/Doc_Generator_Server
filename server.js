const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const multer = require('multer');
const mammoth = require('mammoth'); // Mammoth to extract text from DOCX
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { Document, Packer, Paragraph, TextRun } = require('docx'); // docx.js

// Initialize the Express app
const app = express();

// Define a port
const PORT = process.env.PORT || 4000;

// CORS configuration
app.use(cors());

// Body-parser middleware for handling JSON and URL-encoded data
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// Multer configuration for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/'); // Files will be stored in an "uploads" folder
  },
  filename: (req, file, cb) => {
    cb(null, `${Date.now()}-${file.originalname}`); // Create a unique file name
  },
});

const upload = multer({ storage: storage });

// Basic route
app.get('/', (req, res) => {
  res.send('Hello, World!');
});

// Helper function to read and parse Excel data
const readExcelData = (filePath) => {
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 }); // Get data as array of arrays
  return data;
};

// Helper function to create a Word document using docx.js with the provided content
const generateDocx = async (content, outputFilePath) => {
  // Create a new Document
  const doc = new Document({
    sections: [{
      properties: {},
      children: [
        new Paragraph({
          children: [
            new TextRun(content),
          ],
        }),
      ],
    }],
    creator: 'Your Name', // Specify the creator of the document
    title: 'Generated Document', // Optionally specify a title
  });

  // Write the document to a buffer
  const buffer = await Packer.toBuffer(doc);

  // Write the buffer to a file
  fs.writeFileSync(outputFilePath, buffer);
};

// Route to handle multiple files and document merge
app.post('/upload', upload.fields([
  { name: 'docFile', maxCount: 1 },
  { name: 'excelFile', maxCount: 1 },
]), async (req, res) => {
  if (!req.files || !req.files['docFile'] || !req.files['excelFile']) {
    return res.status(400).send('Both DOC and Excel files are required.');
  }

  const docFile = req.files['docFile'][0];
  const excelFile = req.files['excelFile'][0];

  const docxFilePath = path.join(__dirname, docFile.path);

  // Read the DOCX file using Mammoth to extract text
  let result;
  try {
    result = await mammoth.extractRawText({ path: docxFilePath });
  } catch (err) {
    console.error('Error reading DOCX file:', err);
    return res.status(500).send('Error reading DOCX file.');
  }

  const docContent = result.value; // Raw text extracted from the DOCX file
  console.log('docContent:', docContent);

  // Load the Excel data (first row should contain headers)
  const excelData = readExcelData(path.join(__dirname, excelFile.path));

  // Excel headers (first row) will serve as placeholders mapping
  const headers = excelData[0]; // e.g., ["name", "title"]

  // Array to store generated file paths
  const generatedFiles = [];

  // Loop through all rows (starting from the second row)
  for (let i = 1; i < excelData.length; i++) {
    const rowData = excelData[i]; // e.g., ["John Doe", "Developer"]

    let updatedContent = docContent;

    // Replace placeholders in the docContent with corresponding Excel data
    headers.forEach((header, index) => {
      const placeholder = `\\[\\[${header.trim()}\\]\\]`; // Escape special characters for regex
      const value = rowData[index] || '';

      // Replace placeholders (case insensitive and global)
      updatedContent = updatedContent.replace(new RegExp(placeholder, 'gi'), value);
    });

    // Generate a Word document with the updated content (create new doc for each)
    const outputFilePath = `uploads/output-${Date.now()}-${i}.docx`;
    await generateDocx(updatedContent, path.join(__dirname, outputFilePath));

    // Add the generated file path to the list
    generatedFiles.push(outputFilePath);
  }

  // Send response with the list of generated files
  res.json({
    message: 'Documents generated successfully',
    files: generatedFiles,
  });
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
