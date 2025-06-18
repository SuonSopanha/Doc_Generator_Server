const { Worker } = require('bullmq');
const IORedis = require('ioredis');
const fsPromises = require('fs').promises;
const path = require('path');
const { generateAndPackageDocuments } = require('./doc-generator'); // Correct function name

const QUEUE_NAME = 'documentGeneration';
const UPLOADS_DIR = path.join(__dirname, 'uploads');

// Ensure you have REDIS_HOST and REDIS_PORT environment variables set up if not using defaults
const redisConnectionOptions = {
  host: process.env.REDIS_HOST || '127.0.0.1',
  port: parseInt(process.env.REDIS_PORT, 10) || 6379,
  maxRetriesPerRequest: null, 
  enableReadyCheck: false 
};

console.log(`[Worker] Connecting to Redis at ${redisConnectionOptions.host}:${redisConnectionOptions.port}`);

const worker = new Worker(QUEUE_NAME, async job => {
  console.log(`[Worker] Picked up job ${job.id} with name ${job.name}`);
  console.log('[Worker] Job data:', JSON.stringify(job.data, null, 2));

  const { docxFilePath, dataFilePath, originalDataFileName } = job.data;
  let finalPackagePath;

  try {
    // generateAndPackageDocuments is expected to return the path to the final ZIP file
    finalPackagePath = await generateAndPackageDocuments(job.data);
    console.log(`[Worker] Job ${job.id} processed successfully. Final package at: ${finalPackagePath}`);
    

  } catch (error) {
    console.error(`[Worker] Error processing job ${job.id}:`, error.message, error.stack);
    throw error; // This will mark the job as failed in BullMQ
  } finally {
    // Cleanup the original uploaded files (template and data file) that were processed by this job.
    console.log(`[Worker] Cleaning up original uploaded files for job ${job.id}: ${docxFilePath}, ${dataFilePath}`);
    if (docxFilePath) {
      await fsPromises.unlink(docxFilePath).catch(e => console.error(`[Worker] Error deleting original docx ${docxFilePath} for job ${job.id}: ${e.message}`));
    }
    if (dataFilePath) {
      await fsPromises.unlink(dataFilePath).catch(e => console.error(`[Worker] Error deleting original data file ${dataFilePath} for job ${job.id}: ${e.message}`));
    }
  }
  // Convert absolute path to relative path within uploads directory
  const relativePath = path.relative(UPLOADS_DIR, finalPackagePath);
  return { jobId: job.id, finalPackagePath: relativePath }; // Return relative path for the job completion

}, { connection: new IORedis(redisConnectionOptions), concurrency: process.env.WORKER_CONCURRENCY || 3 }); // Adjust concurrency as needed

worker.on('completed', (job, result) => {
  console.log(`[Worker] Job ${job.id} has completed. Result: ${JSON.stringify(result)}`);
});

worker.on('failed', (job, err) => {
  console.error(`[Worker] Job ${job.id} has failed with error: ${err.message}`);
  console.error('[Worker] Failed job stack:', err.stack);
  // Add more detailed logging or error reporting (e.g., to an error tracking service)
});

worker.on('error', err => {
  // This is for errors in the worker instance itself, not job failures
  console.error('[Worker] BullMQ worker instance error:', err);
});

console.log(`[Worker] Worker started for queue: ${QUEUE_NAME}. Waiting for jobs...`);

// Graceful shutdown
async function shutdown() {
  console.log('[Worker] Shutting down worker...');
  await worker.close();
  console.log('[Worker] Worker closed.');
  process.exit(0);
}

process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);
