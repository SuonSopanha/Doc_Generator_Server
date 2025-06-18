const express = require('express');
const fsPromises = require('fs').promises;
const path = require('path');
const cors = require('cors');
const bodyParser = require('body-parser');
const multer = require('multer');
const { Queue } = require('bullmq');
const IORedis = require('ioredis');
const { v4: uuidv4 } = require('uuid'); // For generating unique job IDs

const UPLOADS_DIR = path.join(__dirname, 'uploads');
const QUEUE_NAME = 'documentGeneration';
const FILE_CLEANUP_INTERVAL = parseInt(process.env.FILE_CLEANUP_INTERVAL || 3) * 60 * 1000; // Default to 3 minutes if not set

// Ensure you have REDIS_HOST and REDIS_PORT environment variables set up if not using defaults
const redisConnectionOptions = {
  host: process.env.REDIS_HOST || '127.0.0.1',
  port: parseInt(process.env.REDIS_PORT, 10) || 6379,
  maxRetriesPerRequest: null,
  enableReadyCheck: false
};

console.log(`[Server] Connecting to Redis for BullMQ at ${redisConnectionOptions.host}:${redisConnectionOptions.port}`);
const documentQueue = new Queue(QUEUE_NAME, { connection: new IORedis(redisConnectionOptions) });

// Optional: Clean uploads folder on server start (be careful with this in production if files are meant to persist)
(async () => {
  try {
    await fsPromises.mkdir(UPLOADS_DIR, { recursive: true });
    // const files = await fsPromises.readdir(UPLOADS_DIR);
    // for (const file of files) {
    //   await fsPromises.unlink(path.join(UPLOADS_DIR, file));
    // }
    // console.log('[Server] Uploads folder cleared (initialization).');
  } catch (err) {
    console.error('[Server] Error initializing uploads folder:', err);
  }
})();

const app = express();
const PORT = process.env.PORT || 4000;

// Function to clean up old files
async function cleanupOldFiles() {
  try {
    const files = await fsPromises.readdir(UPLOADS_DIR);
    const now = Date.now();
    
    for (const file of files) {
      const filePath = path.join(UPLOADS_DIR, file);
      const stats = await fsPromises.stat(filePath);
      const age = now - stats.mtimeMs;
      
      if (age > FILE_CLEANUP_INTERVAL) {
        try {
          await fsPromises.unlink(filePath);
          console.log(`[Cleanup] Deleted old file: ${file}`);
        } catch (error) {
          console.error(`[Cleanup] Error deleting file ${file}:`, error);
        }
      }
    }
  } catch (error) {
    console.error('[Cleanup] Error during file cleanup:', error);
  }
}

// Start cleanup interval
setInterval(cleanupOldFiles, FILE_CLEANUP_INTERVAL);

// Initial cleanup on server start
cleanupOldFiles();

app.use(cors({
  origin: true, 
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: true,
  preflightContinue: false,
  optionsSuccessStatus: 204
}));
app.options('*', cors());

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    fsPromises.mkdir(UPLOADS_DIR, { recursive: true })
      .then(() => cb(null, UPLOADS_DIR))
      .catch(err => cb(err));
  },
  filename: (req, file, cb) => {
    // Using a unique ID in filename to avoid potential collisions if original names are identical
    cb(null, `${Date.now()}-${uuidv4().slice(0,8)}-${file.originalname}`);
  },
});
const upload = multer({ storage: storage });

// File download endpoint
app.get("/api/download/:fileId", async (req, res) => {
  try {
    const filePath = path.join(__dirname, 'uploads', req.params.fileId);
    if (await fsPromises.access(filePath).then(() => true).catch(() => false)) {
      res.download(filePath, path.basename(filePath), (err) => {
        if (err) {
          console.error('Error downloading file:', err);
        }
        // Delete file after download
      });
    } else {
      res.status(404).json({ error: 'File not found' });
    }
  } catch (error) {
    console.error('Error handling download:', error);
    res.status(500).json({ error: 'Error handling download' });
  }
});

app.post('/api/upload', upload.fields([{ name: 'docFile', maxCount: 1 }, { name: 'excelFile', maxCount: 1 }]), async (req, res) => {
  if (!req.files || !req.files['docFile'] || !req.files['excelFile']) {
    return res.status(400).send('Both template document and data file are required.');
  }

  const docFile = req.files['docFile'][0];
  const dataFile = req.files['excelFile'][0];

  const jobId = uuidv4(); // Generate a unique job ID

  const jobData = {
    jobId: jobId,
    docxFilePath: docFile.path, // Path to the uploaded template file
    dataFilePath: dataFile.path, // Path to the uploaded data file
    originalDataFileName: dataFile.originalname, // Keep original name for extension detection
    filterType: req.body.filterType,
    customFrom: req.body.customFrom,
    customTo: req.body.customTo,
    mergingConditionString: req.body.mergingCondition, // Expecting stringified JSON or object
    outputFormat: req.body.outputFormat || 'single',
    outputExtension: req.body.outputExtension || 'docx',
    encryptedPassword: req.body.password
  };

  try {
    await documentQueue.add('generateDocument', jobData, { jobId: jobId }); // Use custom job ID
    console.log(`[Server] Job ${jobId} added to queue with data:`, JSON.stringify(jobData, null, 2));
    res.status(200).json({ 
      message: 'Document generation request accepted. Processing in background.', 
      jobId: jobId 
      // You might want to add a URL here for the client to poll job status
      // e.g., statusUrl: `/api/job-status/${jobId}`
    });
  } catch (error) {
    console.error('[Server] Error adding job to queue:', error);
    // If adding to queue fails, try to clean up uploaded files
    await fsPromises.unlink(docFile.path).catch(e => console.error(`Cleanup error for ${docFile.path}: ${e}`));
    await fsPromises.unlink(dataFile.path).catch(e => console.error(`Cleanup error for ${dataFile.path}: ${e}`));
    res.status(500).send('Error submitting document generation request.');
  }
});

// Optional: Endpoint to check job status (basic example)
app.get('/api/job-status/:jobId', async (req, res) => {
  const { jobId } = req.params;
  try {
    const job = await documentQueue.getJob(jobId);
    if (!job) {
      return res.status(404).json({ jobId, status: 'not_found', message: 'Job not found.' });
    }
    const state = await job.getState();
    const progress = job.progress;
    const returnValue = job.returnvalue;
    const failedReason = job.failedReason;

    res.json({
      jobId,
      status: state,
      progress,
      timestamp: job.timestamp ? new Date(job.timestamp) : null,
      processedOn: job.processedOn ? new Date(job.processedOn) : null,
      finishedOn: job.finishedOn ? new Date(job.finishedOn) : null,
      returnValue, // This would contain { jobId, finalPackagePath } upon successful completion
      failedReason
    });
  } catch (error) {
    console.error(`[Server] Error fetching status for job ${jobId}:`, error);
    res.status(500).send('Error fetching job status.');
  }
});

app.get('/api/hello', (req, res) => {
  res.send('Hello, World!');
});

app.listen(PORT, () => {
  console.log(`[Server] Server is running on http://localhost:${PORT}`);
  console.log(`[Server] Document generation queue '${QUEUE_NAME}' is active.`);
});
