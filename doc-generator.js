const fs = require('fs');
const fsPromises = require('fs').promises;
const path = require('path');
const xlsx = require('xlsx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const DocxMerger = require('docx-merger');
const archiver = require('archiver');
archiver.registerFormat('zip-encrypted', require('archiver-zip-encrypted'));
const libre = require('libreoffice-convert');

const UPLOADS_DIR = path.join(__dirname, 'uploads'); // Centralized uploads directory path

// Ensure uploads directory exists on module load
(async () => {
  try {
    await fsPromises.mkdir(UPLOADS_DIR, { recursive: true });
  } catch (error) {
    console.error('Failed to create uploads directory on init:', error);
  }
})();

const readExcelData = (filePath) => {
  console.log('Reading Excel:', filePath);
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  return data;
};

const readCsvData = (filePath) => {
  console.log('Reading CSV:', filePath);
  const fileContent = fs.readFileSync(filePath, 'utf8');
  const rows = fileContent
    .trim()
    .split('\n')
    .map((row) => row.split(','));
  return rows;
};

const readJsonData = (filePath) => {
  console.log('Reading JSON:', filePath);
  const fileContent = fs.readFileSync(filePath, 'utf8');
  const jsonData = JSON.parse(fileContent);

  if (!Array.isArray(jsonData) || jsonData.length === 0) {
    throw new Error('JSON data must be a non-empty array of objects');
  }

  const headers = Object.keys(jsonData[0]);
  const rows = jsonData.map(obj => headers.map(header => obj[header] || ''));
  return [headers, ...rows];
};

const filterData = (data, startRow, endRow, option) => {
  return data.slice(startRow, endRow + 1).filter((_, index) => {
    if (option === 'odd') return (startRow + index) % 2 !== 0;
    if (option === 'even') return (startRow + index) % 2 === 0;
    return true; 
  });
};

const conditionalData = (data, column, operator, value) => {
  if (!column || !operator || value === undefined) return data;
  const headers = data[0];
  const columnIndex = headers.indexOf(column);
  if (columnIndex === -1) return data;
  return data.filter((row, index) => {
    if (index === 0) return true; 
    const cellValue = row[columnIndex];
    switch (operator) {
      case '=':
      case '==': return cellValue == value;
      case '!=':
      case '<>': return cellValue != value;
      case '>': return cellValue > value;
      case '<': return cellValue < value;
      case '>=': return cellValue >= value;
      case '<=': return cellValue <= value;
      case 'contains': return String(cellValue).includes(String(value));
      case 'startsWith': return String(cellValue).startsWith(String(value));
      case 'endsWith': return String(cellValue).endsWith(String(value));
      default: return true;
    }
  });
};

const generateDocuments = async (docxFilePath, data, filterType, range = {}, mergingCondition = null) => {
  const headers = data[0];
  const generatedFiles = [];
  let filteredData = mergingCondition
    ? conditionalData(data, mergingCondition.column, mergingCondition.operator, mergingCondition.value)
    : data;

  let startRow = 1;
  let endRow = filteredData.length;
  let increment = 1;

  if (filterType === 'odd') {
    startRow = 1; 
    increment = 2;
  } else if (filterType === 'even') {
    startRow = 2;
    increment = 2;
  } else if (filterType === 'custom') {
    startRow = Math.max(1, parseInt(range.from) || 1);
    endRow = Math.min(filteredData.length, (parseInt(range.to) + 1) || filteredData.length);
  }

  for (let i = startRow; i < endRow; i += increment) {
    const rowData = filteredData[i];
    if (!rowData) continue;
    const dataObj = {};
    headers.forEach((header, index) => {
      const headerKey = String(header).trim();
      dataObj[headerKey] = rowData[index] !== undefined ? rowData[index] : '';
    });

    const content = fs.readFileSync(docxFilePath, 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, lineBreaks: true });
    doc.setData(dataObj);
    doc.render();
    const buf = doc.getZip().generate({ type: 'nodebuffer' });
    const outputFilePath = path.join(__dirname, `uploads/output-${Date.now()}-${i}.docx`);
    fs.writeFileSync(outputFilePath, buf);
    generatedFiles.push(outputFilePath);
  }
  if (generatedFiles.length === 0) {
    throw new Error('No documents were generated. Filter conditions/range may have excluded all rows.');
  }
  return generatedFiles;
};

const convertDocxToPdf = (docxFilePath) => {
  return new Promise((resolve, reject) => {
    const pdfFilePath = docxFilePath.replace(/\.docx$/, '.pdf');
    const fileBuffer = fs.readFileSync(docxFilePath);
    libre.convert(fileBuffer, '.pdf', undefined, (err, done) => {
      if (err) {
        console.error('Error converting DOCX to PDF:', err);
        return reject(err);
      }
      fs.writeFileSync(pdfFilePath, done);
      resolve(pdfFilePath);
    });
  });
};

// Returns path to the final ZIP file
const handleSingleFileOutput = async (generatedIndividualFiles, outputExtension, encryptedPassword) => {
  return new Promise(async (resolve, reject) => {
    let finalDocxPathForMerging;
    try {
      const fileBuffers = generatedIndividualFiles.map(fp => fs.readFileSync(fp));
      const merger = new DocxMerger({ removeTrailingLineBreaks: true }, fileBuffers);
      finalDocxPathForMerging = path.join(__dirname, `uploads/combined-output-${Date.now()}.docx`);
      
      merger.save('nodebuffer', async (mergedBuffer) => {
        try {
          fs.writeFileSync(finalDocxPathForMerging, mergedBuffer);
          let finalFileToZip = finalDocxPathForMerging;
          let tempPdfPath = null;

          if (outputExtension === 'pdf') {
            tempPdfPath = await convertDocxToPdf(finalDocxPathForMerging);
            finalFileToZip = tempPdfPath;
            await fsPromises.unlink(finalDocxPathForMerging); // remove merged .docx after .pdf conversion
          }

          const zipFilePath = path.join(__dirname, `uploads/package-single-${Date.now()}.zip`);
          const output = fs.createWriteStream(zipFilePath);
          const archive = encryptedPassword
            ? archiver('zip-encrypted', { zlib: { level: 9 }, encryptionMethod: 'aes256', password: encryptedPassword })
            : archiver('zip', { zlib: { level: 9 } });

          output.on('close', async () => {
            console.log(`Single file ZIP created: ${zipFilePath}`);
            // Cleanup intermediate individual .docx files
            await Promise.all(generatedIndividualFiles.map(f => fsPromises.unlink(f).catch(e => console.error(`Cleanup error for ${f}:`, e))));
            // If PDF was created, the merged DOCX was already deleted. The PDF (finalFileToZip) is in the zip.
            // If DOCX was zipped, finalFileToZip is finalDocxPathForMerging, which is in the zip.
            // The file that was put into the zip (finalFileToZip) should NOT be deleted here, it's the content of the zip.
            resolve(zipFilePath);
          });
          archive.on('error', reject);
          archive.pipe(output);
          archive.file(finalFileToZip, { name: path.basename(finalFileToZip) });
          await archive.finalize();

        } catch (processingError) {
          reject(processingError);
        }
      });
    } catch (err) {
      // Cleanup on error
      if (finalDocxPathForMerging) await fsPromises.unlink(finalDocxPathForMerging).catch(e => console.error(`Cleanup error for ${finalDocxPathForMerging}:`, e));
      await Promise.all(generatedIndividualFiles.map(f => fsPromises.unlink(f).catch(e => console.error(`Cleanup error for ${f}:`, e))));
      reject(err);
    }
  });
};

// Returns path to the final ZIP file
const handleMultipleFileOutput = async (generatedIndividualFiles, outputExtension, encryptedPassword) => {
  return new Promise(async (resolve, reject) => {
    let filesToArchivePaths = [...generatedIndividualFiles];
    let tempPdfPaths = []; // To keep track of PDFs if conversion happens, for later cleanup if needed

    try {
      if (outputExtension === 'pdf') {
        const pdfConversionPromises = generatedIndividualFiles.map(async (docxFile) => {
          const pdfPath = await convertDocxToPdf(docxFile);
          tempPdfPaths.push(pdfPath); // Add to temp list for potential cleanup on error
          return pdfPath;
        });
        filesToArchivePaths = await Promise.all(pdfConversionPromises);
        // Original docx files (generatedIndividualFiles) can be cleaned up after zipping the PDFs
      }

      const zipFilePath = path.join(__dirname, `uploads/package-multi-${Date.now()}.zip`);
      const output = fs.createWriteStream(zipFilePath);
      const archive = encryptedPassword ? archiver('zip-encrypted', { zlib: { level: 9 }, encryptionMethod: 'aes256', password: encryptedPassword }) : archiver('zip', { zlib: { level: 9 } }); // Add encryption if needed for multi-file zip

      output.on('close', async () => {
        console.log(`Multiple file ZIP created: ${zipFilePath}`);
        // Cleanup original .docx files that were generated
        await Promise.all(generatedIndividualFiles.map(f => fsPromises.unlink(f).catch(e => console.error(`Cleanup error for original docx ${f}:`, e))));
        // If PDFs were created and zipped, they are the content of the zip. They shouldn't be deleted here.
        // If original DOCX were zipped, they are already cleaned up by the above line.
        resolve(zipFilePath);
      });
      archive.on('error', reject);
      archive.pipe(output);
      filesToArchivePaths.forEach((filePath, index) => {
        archive.file(filePath, { name: `document-${index + 1}${path.extname(filePath)}` });
      });
      await archive.finalize();

    } catch (err) {
      // Cleanup on error
      await Promise.all(generatedIndividualFiles.map(f => fsPromises.unlink(f).catch(e => console.error(`Cleanup error for ${f}:`, e))));
      await Promise.all(tempPdfPaths.map(f => fsPromises.unlink(f).catch(e => console.error(`Cleanup error for temp PDF ${f}:`, e))));
      reject(err);
    }
  });
};

// This function is for direct HTTP handling, will be replaced by queue logic for background processing
const processAndRespond = async (req, res) => {
  if (!req.files || !req.files['docFile'] || !req.files['excelFile']) {
    return res.status(400).send('Both template document and data file are required.');
  }
  const docFile = req.files['docFile'][0];
  const dataFile = req.files['excelFile'][0];
  const { filterType, customFrom, customTo, mergingCondition: mergingConditionString, outputFormat, outputExtension, password: encryptedPassword } = req.body;

  let finalZipPath;
  try {
    const jobData = {
      docxFilePath: docFile.path,
      dataFilePath: dataFile.path,
      originalDataFileName: dataFile.originalname,
      filterType,
      customFrom,
      customTo,
      mergingConditionString,
      outputFormat,
      outputExtension,
      encryptedPassword
    };
    finalZipPath = await generateAndPackageDocuments(jobData);
    res.download(finalZipPath, path.basename(finalZipPath), async (downloadErr) => {
      if (downloadErr) console.error('Error sending file:', downloadErr);
      await fsPromises.unlink(finalZipPath).catch(e => console.error('Cleanup error for final zip:', e));
      // Original uploaded files (docFile.path, dataFile.path) are cleaned up by the worker/caller of generateAndPackageDocuments
      // Here, for direct HTTP, we should clean them.
      await fsPromises.unlink(docFile.path).catch(e => console.error('Cleanup error for uploaded docx:', e));
      await fsPromises.unlink(dataFile.path).catch(e => console.error('Cleanup error for uploaded data file:', e));
    });
  } catch (error) {
    console.error('Error in processAndRespond:', error);
    if (finalZipPath) await fsPromises.unlink(finalZipPath).catch(e => console.error('Cleanup error for final zip on error:', e));
    await fsPromises.unlink(docFile.path).catch(e => console.error('Cleanup error for uploaded docx on error:', e));
    await fsPromises.unlink(dataFile.path).catch(e => console.error('Cleanup error for uploaded data file on error:', e));
    if (!res.headersSent) {
      res.status(500).send(`Error processing documents: ${error.message}`);
    }
  }
};

// New main function for worker
const generateAndPackageDocuments = async (jobData) => {
  const { 
    docxFilePath, 
    dataFilePath, 
    originalDataFileName, 
    filterType, 
    customFrom, 
    customTo, 
    mergingConditionString, 
    outputFormat, 
    outputExtension, 
    encryptedPassword 
  } = jobData;

  try {
    let data;
    const fileExtension = path.extname(originalDataFileName).toLowerCase();
    switch (fileExtension) {
      case '.xlsx': case '.xls': data = readExcelData(dataFilePath); break;
      case '.csv': data = readCsvData(dataFilePath); break;
      case '.json': data = readJsonData(dataFilePath); break;
      default: throw new Error('Unsupported data file format.');
    }

    let mergingConditionObj = null;
    if (mergingConditionString) {
        try { mergingConditionObj = JSON.parse(mergingConditionString); } 
        catch (e) { throw new Error('Invalid merging condition format.'); }
    }

    const generatedIndividualFiles = await generateDocuments(
      docxFilePath,
      data,
      filterType,
      { from: customFrom, to: customTo },
      mergingConditionObj
    );

    let finalPackagePath;
    if (outputFormat === 'single') {
      finalPackagePath = await handleSingleFileOutput(generatedIndividualFiles, outputExtension, encryptedPassword);
    } else { // multiple
      finalPackagePath = await handleMultipleFileOutput(generatedIndividualFiles, outputExtension, encryptedPassword);
    }
    // Intermediate individual files (generatedIndividualFiles) are cleaned up by handleSingle/MultipleFileOutput.
    // PDF versions of individual files (if created by handleMultipleFileOutput) are also cleaned up by it after zipping.
    // The finalPackagePath is the path to the ZIP file. This should be returned.
    // The original docxFilePath and dataFilePath (input to this function) should be cleaned by the worker after this function successfully returns.
    return finalPackagePath; 

  } catch (error) {
    console.error('Error in generateAndPackageDocuments:', error.message);
    // Do not clean up docxFilePath/dataFilePath here, worker should do it based on job success/failure.
    // Intermediate files should have been attempted to be cleaned by sub-functions.
    throw error; 
  }
};

module.exports = {
  processAndRespond, 
  generateAndPackageDocuments, 
  readExcelData, 
  readCsvData, 
  readJsonData,
  filterData,
  conditionalData,
  generateDocuments,
  convertDocxToPdf,
  // handleSingleFileOutput and handleMultipleFileOutput are primarily internal to generateAndPackageDocuments now
  // but can be exported if needed for testing or other specific uses.
  // handleSingleFileOutput, 
  // handleMultipleFileOutput
};
