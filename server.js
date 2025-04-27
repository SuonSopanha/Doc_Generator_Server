const express = require('express');
const fs = require('fs').promises;
const path = require('path');

// Delete all files in uploads folder on server start
(async () => {
  const uploadsDir = path.join(__dirname, 'uploads');
  try {
    const files = await fs.readdir(uploadsDir);
    for (const file of files) {
      const filePath = path.join(uploadsDir, file);
      const stat = await fs.lstat(filePath);
      if (stat.isFile()) {
        await fs.unlink(filePath);
      }
    }
    console.log('Uploads folder cleared.');
  } catch (err) {
    if (err.code === 'ENOENT') {
      // uploads folder does not exist, ignore
    } else {
      console.error('Error clearing uploads folder:', err);
    }
  }
})();

const cors = require("cors");
const bodyParser = require("body-parser");
const multer = require("multer");
const xlsx = require("xlsx");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const DocxMerger = require("docx-merger");
const archiver = require("archiver");
archiver.registerFormat("zip-encrypted", require("archiver-zip-encrypted"));
const libre = require("libreoffice-convert");

const app = express();
const PORT = process.env.PORT || 4000;

// Configure CORS with specific options
app.use(cors({
  origin: true, // Allow all origins
  methods: ['GET', 'POST', 'OPTIONS'], // Allow these methods
  allowedHeaders: ['Content-Type', 'Authorization'], // Allow these headers
  credentials: true, // Allow credentials
  preflightContinue: false,
  optionsSuccessStatus: 204
}));

// Handle preflight requests
app.options('*', cors());

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// Multer configuration for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, "uploads/");
  },
  filename: (req, file, cb) => {
    cb(null, `${Date.now()}-${file.originalname}`);
  },
});

const upload = multer({ storage: storage });

// Helper function to read and parse Excel data
const readExcelData = (filePath) => {
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  return data;
};

const readCsvData = (filePath) => {
  const fileContent = fs.readFileSync(filePath, "utf8");
  const rows = fileContent
    .trim()
    .split("\n")
    .map((row) => row.split(","));
  return rows;
};

const readJsonData = (filePath) => {
  console.log(filePath);
  const fileContent = fs.readFileSync(filePath, "utf8");
  const jsonData = JSON.parse(fileContent);
  console.log(jsonData);

  // Ensure jsonData is an array
  if (!Array.isArray(jsonData)) {
    throw new Error('JSON data must be an array of objects');
  }

  // Get headers from the first object
  const headers = Object.keys(jsonData[0]);

  // Convert each object to an array of values in the same order as headers
  const rows = jsonData.map(obj => headers.map(header => obj[header] || ""));

  // Return in same format as Excel/CSV: [headers, ...rows]
  return [headers, ...rows];
};

const filterData = (data, startRow, endRow, option) => {
  return data.slice(startRow, endRow + 1).filter((_, index) => {
    if (option === "odd") return (startRow + index) % 2 !== 0;
    if (option === "even") return (startRow + index) % 2 === 0;
    return true; // Return all if no valid option is given
  });
};

const conditionalData = (data, column, operator, value) => {
  if (!column || !operator || value === undefined) return data;
  
  const headers = data[0];
  const columnIndex = headers.indexOf(column);
  
  if (columnIndex === -1) return data;
  
  return data.filter((row, index) => {
    if (index === 0) return true; // Keep headers
    const cellValue = row[columnIndex];
    
    switch (operator) {
      case '=':
      case '==':
        return cellValue == value;
      case '!=':
      case '<>':
        return cellValue != value;
      case '>':
        return cellValue > value;
      case '<':
        return cellValue < value;
      case '>=':
        return cellValue >= value;
      case '<=':
        return cellValue <= value;
      case 'contains':
        return String(cellValue).includes(String(value));
      case 'startsWith':
        return String(cellValue).startsWith(String(value));
      case 'endsWith':
        return String(cellValue).endsWith(String(value));
      default:
        return true;
    }
  });
};

// Helper function to generate individual documents
const generateDocuments = async (
  docxFilePath,
  excelData,
  filterType,
  range = {},
  mergingCondition = null
) => {
  const headers = excelData[0];
  const generatedFiles = [];

  // First apply the merging condition if it exists
  let filteredData = mergingCondition 
    ? conditionalData(excelData, mergingCondition.column, mergingCondition.operator, mergingCondition.value)
    : excelData;

  let startRow = 1;
  let endRow = filteredData.length;
  let increment = 1;

  if(filterType === 'odd') {
    startRow = 1;
    increment = 2;
  } else if(filterType === 'even') {
    startRow = 2;
    increment = 2;
  } else if(filterType === 'custom') {
    startRow = Math.max(1, parseInt(range.from) || 1);
    endRow = Math.min(filteredData.length, (parseInt(range.to) + 1) || filteredData.length);
    increment = 1;
  }

  console.log('Filtered data:', filteredData);

  for (let i = startRow; i < endRow; i += increment) {
    const rowData = filteredData[i];
    
    // Skip if row is undefined
    if (!rowData) continue;
    
    const data = {};

    headers.forEach((header, index) => {
      // Ensure header is trimmed and handle undefined values
      const headerKey = header.trim();
      data[headerKey] = rowData[index] !== undefined ? rowData[index] : "";
    });

    const content = fs.readFileSync(docxFilePath, "binary");
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      lineBreaks: true,
    });

    doc.setData(data);
    doc.render();

    const buf = doc.getZip().generate({ type: "nodebuffer" });
    const outputFilePath = path.join(
      __dirname,
      `uploads/output-${Date.now()}-${i}.docx`
    );
    fs.writeFileSync(outputFilePath, buf);
    generatedFiles.push(outputFilePath);
  }

  if (generatedFiles.length === 0) {
    throw new Error("No documents were generated. The filter conditions may have excluded all rows.");
  }

  return generatedFiles;
};

const convertDocxToPdf = (docxFilePath) => {
  return new Promise((resolve, reject) => {
    const pdfFilePath = docxFilePath.replace(/\.docx$/, ".pdf");
    const fileBuffer = fs.readFileSync(docxFilePath);

    libre.convert(fileBuffer, ".pdf", undefined, (err, done) => {
      if (err) {
        console.error("Error converting DOCX to PDF:", err);
        reject(err);
        return;
      }
      fs.writeFileSync(pdfFilePath, done);
      resolve(pdfFilePath);
    });
  });
};

// Helper function to handle single file output
// Helper function to handle single file output
const handleSingleFileOutput = async (
  generatedFiles,
  res,
  outputExtension,
  encryptedPassword
) => {
  try {
    // Read file buffers from generated files
    const fileBuffers = generatedFiles.map((filePath) =>
      fs.readFileSync(filePath)
    );

    // Merge DOCX files
    const merger = new DocxMerger(
      { removeTrailingLineBreaks: true },
      fileBuffers
    );
    const finalDocxPath = path.join(
      __dirname,
      `uploads/combined-output-${Date.now()}.docx`
    );

    merger.save("nodebuffer", async function (mergedBuffer) {
      fs.writeFileSync(finalDocxPath, mergedBuffer);

      let finalOutputPath = finalDocxPath;

      // Convert to PDF if required
      if (outputExtension === "pdf") {
        try {
          finalOutputPath = await convertDocxToPdf(finalDocxPath);
          fs.unlinkSync(finalDocxPath); // Remove intermediate DOCX file
        } catch (error) {
          console.error("Error during DOCX to PDF conversion:", error);
          res.status(500).send("Error converting to PDF.");
          return;
        }
      }

      // Prepare the ZIP archive
      const zipFilePath = path.join(
        __dirname,
        `uploads/output-archive-${Date.now()}.zip`
      );

      const output = fs.createWriteStream(zipFilePath);
      const archive = encryptedPassword
        ? archiver("zip-encrypted", {
            zlib: { level: 9 },
            encryptionMethod: "aes256",
            password: encryptedPassword,
          })
        : archiver("zip", { zlib: { level: 9 } });

      output.on("close", () => {
        console.log(`ZIP file created (${archive.pointer()} total bytes)`);

        // Send ZIP file to the client
        res.download(zipFilePath, path.basename(zipFilePath), (err) => {
          if (err) {
            console.error("Error sending ZIP file:", err);
            res.status(500).send("Error sending ZIP file.");
          }

          // Cleanup temporary files
          try {
            fs.unlinkSync(zipFilePath);
            fs.unlinkSync(finalOutputPath);
          } catch (error) {
            console.error("Error deleting ZIP or final output file:", error);
          }

          generatedFiles.forEach((file) => {
            try {
              fs.unlinkSync(file);
            } catch (error) {
              console.error(`Error deleting file ${file}:`, error);
            }
          });
        });
      });

      archive.on("error", (err) => {
        console.error("Error creating ZIP archive:", err);
        res.status(500).send("Error creating ZIP archive.");
      });

      // Pipe archive data to the output file
      archive.pipe(output);

      // Append the final output file to the archive
      archive.append(fs.createReadStream(finalOutputPath), {
        name: path.basename(finalOutputPath),
      });

      // Finalize the archive
      archive.finalize();
    });
  } catch (error) {
    console.error("Error combining files or creating ZIP:", error);

    // Cleanup temporary files
    generatedFiles.forEach((file) => {
      try {
        fs.unlinkSync(file);
      } catch (error) {
        console.error(`Error deleting file ${file}:`, error);
      }
    });

    res.status(500).send("Error combining files or creating ZIP.");
  }
};

// Helper function to handle multiple file output (ZIP)
const handleMultipleFileOutput = async (
  generatedFiles,
  res,
  outputExtension,
  encryptedPassword
) => {
  const zipFilePath = path.join(
    __dirname,
    `uploads/documents-${Date.now()}.zip`
  );
  const output = fs.createWriteStream(zipFilePath);
  const archive = archiver("zip", {
    zlib: { level: 9 },
  });

  output.on("close", () => {
    res.download(zipFilePath, "documents.zip", (err) => {
      if (err) {
        console.error("Error sending ZIP file:", err);
        res.status(500).send("Error sending ZIP file.");
      }

      // Clean up all files after sending
      generatedFiles.forEach((file) => {
        try {
          fs.unlinkSync(file);
        } catch (error) {
          console.error(`Error deleting file ${file}:`, error);
        }
      });

      try {
        fs.unlinkSync(zipFilePath);
      } catch (error) {
        console.error("Error deleting ZIP file:", error);
      }
    });
  });

  archive.on("error", (err) => {
    console.error("Error creating ZIP:", err);
    res.status(500).send("Error creating ZIP file.");

    // Clean up generated files on error
    generatedFiles.forEach((file) => {
      try {
        fs.unlinkSync(file);
      } catch (error) {
        console.error(`Error deleting file ${file}:`, error);
      }
    });
  });

  archive.pipe(output);

  // Add each file to the archive with a proper name
  generatedFiles.forEach((file, index) => {
    archive.file(file, { name: `document-${index + 1}.docx` });
  });

  await archive.finalize();
};

// Main route to handle document processing
app.post(
  "/upload",
  upload.fields([
    { name: "docFile", maxCount: 1 },
    { name: "excelFile", maxCount: 1 },
  ]),
  async (req, res) => {
    if (!req.files || !req.files["docFile"] || !req.files["excelFile"]) {
      return res.status(400).send("Both template document and data file (Excel/CSV/JSON) are required.");
    }

    const docFile = req.files["docFile"][0];
    const dataFile = req.files["excelFile"][0];
    const docxFilePath = path.join(__dirname, docFile.path);
    const dataFilePath = path.join(__dirname, dataFile.path);

    try {
      // Determine file type and read data accordingly
      let data;
      const fileExtension = path.extname(dataFile.originalname).toLowerCase();
      
      switch (fileExtension) {
        case '.xlsx':
        case '.xls':
          data = readExcelData(dataFilePath);
          break;
        case '.csv':
          data = readCsvData(dataFilePath);
          break;
        case '.json':
          data = readJsonData(dataFilePath);
          break;
        default:
          throw new Error('Unsupported file format. Please use Excel (.xlsx/.xls), CSV (.csv), or JSON (.json) files.');
      }
      
      // Parse mergingCondition from the request body
      let mergingCondition = null;
      try {
        if (req.body.mergingCondition) {
          // If mergingCondition is sent as a string (which happens with multipart/form-data), parse it
          const parsedCondition = typeof req.body.mergingCondition === 'string' 
            ? JSON.parse(req.body.mergingCondition)
            : req.body.mergingCondition;
            
          mergingCondition = {
            column: parsedCondition.column,
            operator: parsedCondition.operator,
            value: parsedCondition.value
          };
        }
      } catch (error) {
        console.error('Error parsing mergingCondition:', error);
        mergingCondition = null;
      }

      console.log('Parsed mergingCondition:', mergingCondition);

      const generatedFiles = await generateDocuments(
        docxFilePath,
        data,
        req.body.filterType,
        { from: req.body.customFrom, to: req.body.customTo },
        mergingCondition
      );
      const outputExtension = req.body.outputExtension || "docx";
      const encryptedPassword = req.body.password;

      // Handle output based on format
      if (req.body.outputFormat === "single") {
        await handleSingleFileOutput(
          generatedFiles,
          res,
          outputExtension,
          encryptedPassword
        );
      } else {
        await handleMultipleFileOutput(
          generatedFiles,
          res,
          outputExtension,
          encryptedPassword
        );
      }

      // Clean up uploaded files
      try {
        fs.unlinkSync(docxFilePath);
        fs.unlinkSync(dataFilePath);
      } catch (error) {
        console.error("Error cleaning up uploaded files:", error);
      }
    } catch (error) {
      console.error("Error processing documents:", error);
      res.status(500).send("Error processing documents. error: " + error.message);
    }
  }
);

// Basic route
app.get("/", (req, res) => {
  res.send("Hello, World!");
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
