const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const multer = require("multer");
const mammoth = require("mammoth"); // Optional if you still want to extract text from DOCX
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const archiver = require("archiver");

const app = express();
const PORT = process.env.PORT || 4000;

app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// Multer configuration for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, "uploads/"); // Files will be stored in an "uploads" folder
  },
  filename: (req, file, cb) => {
    cb(null, `${Date.now()}-${file.originalname}`); // Create a unique file name
  },
});

const upload = multer({ storage: storage });

// Basic route
app.get("/", (req, res) => {
  res.send("Hello, World!");
});

// Helper function to read and parse Excel data
const readExcelData = (filePath) => {
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  return data;
};

// Route to handle multiple files and document merge
app.post(
  "/upload",
  upload.fields([
    { name: "docFile", maxCount: 1 },
    { name: "excelFile", maxCount: 1 },
  ]),
  async (req, res) => {
    if (!req.files || !req.files["docFile"] || !req.files["excelFile"]) {
      return res.status(400).send("Both DOC and Excel files are required.");
    }

    const docFile = req.files["docFile"][0];
    const excelFile = req.files["excelFile"][0];

    const docxFilePath = path.join(__dirname, docFile.path);

    // Read the DOCX file using PizZip and Docxtemplater
    let zip, doc;
    try {
      const content = fs.readFileSync(docxFilePath, "binary");
      zip = new PizZip(content);
      doc = new Docxtemplater(zip, { paragraphLoop: true, lineBreaks: true });
    } catch (err) {
      console.error("Error reading DOCX file:", err);
      return res.status(500).send("Error reading DOCX file.");
    }

    // Read Excel data
    const excelData = readExcelData(path.join(__dirname, excelFile.path));
    const headers = excelData[0]; // e.g., ["name", "title"]
    const generatedFiles = [];

    // Loop through all rows (starting from the second row)
    // Loop through all rows (starting from the second row)
    for (let i = 1; i < excelData.length; i++) {
      const rowData = excelData[i];

      // Create a data object for the current row
      const data = {};
      headers.forEach((header, index) => {
        data[header.trim()] = rowData[index] || ""; // Map headers to values
      });

      // Create a new Docxtemplater instance for each document
      const newDoc = new Docxtemplater(
        new PizZip(fs.readFileSync(docxFilePath, "binary")),
        { paragraphLoop: true, lineBreaks: true }
      );

      // Set data in the document
      newDoc.setData(data);

      // Render the document
      try {
        newDoc.render();
      } catch (error) {
        console.error("Error during rendering:", error);
        return res.status(500).send("Error rendering document.");
      }

      // Generate the buffer for the new document
      const buf = newDoc.getZip().generate({ type: "nodebuffer" });

      // Write the buffer to a new DOCX file
      const outputFilePath = path.join(
        __dirname,
        `uploads/output-${Date.now()}-${i}.docx`
      );
      fs.writeFileSync(outputFilePath, buf);
      generatedFiles.push(outputFilePath); // Add to generated files
    }

    // Create a ZIP file to send the generated documents
    const zipFilePath = path.join(
      __dirname,
      `uploads/generated-documents-${Date.now()}.zip`
    );
    const zipStream = fs.createWriteStream(zipFilePath);
    const archive = archiver("zip", {
      zlib: { level: 9 },
    });

    zipStream.on("close", () => {
      console.log(`ZIP file created: ${zipFilePath}`);
      res.download(zipFilePath, (err) => {
        if (err) {
          console.error("Error sending ZIP file:", err);
          res.status(500).send("Error sending ZIP file.");
        } else {
          // Clean up generated files after sending
          generatedFiles.forEach((file) => fs.unlinkSync(file)); // Delete generated documents
          fs.unlinkSync(zipFilePath); // Delete the ZIP file
        }
      });
    });

    // Pipe the ZIP stream to the archive
    archive.pipe(zipStream);

    // Append the generated files to the ZIP archive
    generatedFiles.forEach((file) => {
      archive.file(file, { name: path.basename(file) });
    });

    // Finalize the archive
    archive.finalize();
  }
);

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
