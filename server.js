const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const multer = require('multer');

// Initialize the Express app
const app = express();

// Define a port
const PORT = process.env.PORT || 3000;

// CORS configuration
app.use(cors());

// Body-parser middleware for handling JSON and URL-encoded data
app.use(bodyParser.json()); // Parse JSON bodies
app.use(bodyParser.urlencoded({ extended: true })); // Parse URL-encoded bodies

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

// Route to handle JSON data
app.post('/json-data', (req, res) => {
  console.log(req.body); // Logs the received JSON data
  res.json({ message: 'JSON data received', data: req.body });
});

// Route to handle file upload
app.post('/upload', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }
  res.json({ message: 'File uploaded successfully', file: req.file });
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
