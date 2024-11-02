const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const path = require('path');

async function readDocx(filePath) {
    try {
        // Read the file content as binary
        const content = fs.readFileSync(filePath, "binary");
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, lineBreaks: true });

        // Extract text
        const text = doc.getFullText();
        console.log('Extracted Text:', text);

        // Generate new DOCX file with the extracted content
        generateDocx(text);
    } catch (error) {
        console.error('Error reading DOCX file:', error);
    }
}

function generateDocx(content) {
    const zip = new PizZip();
    const doc = new Docxtemplater(zip, { paragraphLoop: true, lineBreaks: true });

    // Set data in the template (replace placeholders if any)
    doc.setData({ content });

    try {
        // Render the document
        doc.render();
    } catch (error) {
        console.error('Error during rendering:', error);
        return;
    }

    // Generate the buffer for the new document
    const buf = doc.getZip().generate({ type: 'nodebuffer' });

    // Write the buffer to a new DOCX file
    const outputFilePath = path.join(__dirname, 'Generated-Document.docx');
    fs.writeFileSync(outputFilePath, buf);
    console.log('New DOCX file generated:', outputFilePath);
}

// Use an absolute path to the DOCX file
const inputFilePath = path.join(__dirname, 'English-Homework.docx');
readDocx(inputFilePath);
