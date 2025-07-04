// routes/sheetRoutes.js
const express = require('express');
const router = express.Router();
const Sheet = require('../models/Sheet'); // Import your Mongoose model

// --- API ENDPOINTS ---

// 1. Save/Create a new spreadsheet or Update an existing one
// POST /api/sheets
router.post('/', async (req, res) => {
  try {
    const { name, worksheets } = req.body; // Expecting name and the full worksheets object
  console.log("Incoming data:", JSON.stringify(req.body, null, 2));
    if (!name || !worksheets) {
      return res.status(400).json({ message: 'Sheet name and data are required.' });
    }

    // Try to find if a sheet with this name already exists
    let sheet = await Sheet.findOne({ name });

    if (sheet) {
      // If exists, update it
      sheet.worksheets = worksheets;
      await sheet.save();
      return res.status(200).json({ message: 'Spreadsheet updated successfully', sheet });
    } else {
      // If not, create a new one
      sheet = new Sheet({ name, worksheets });
      await sheet.save();
      return res.status(201).json({ message: 'Spreadsheet saved successfully', sheet });
    }
  } catch (err) {
    if (err.code === 11000) { // MongoDB duplicate key error (for 'name' unique index)
      return res.status(409).json({ message: 'A sheet with this name already exists.' });
    }
    console.error('Error saving/updating sheet:', err);
    res.status(500).json({ message: 'Server error during save/update.' });
  }
});

// 2. Load a specific spreadsheet by ID or Name (prefer ID if possible, but Name is user-friendly)
// GET /api/sheets/:name
router.get('/:name', async (req, res) => {
  try {
    const sheet = await Sheet.findOne({ name: req.params.name });
    if (!sheet) {
      return res.status(404).json({ message: 'Spreadsheet not found.' });
    }
    res.status(200).json(sheet); // Send the full sheet object
  } catch (err) {
    console.error('Error loading sheet:', err);
    res.status(500).json({ message: 'Server error during load.' });
  }
});

// 3. (Optional) Get a list of all saved spreadsheet names (for "Open" dialog)
// GET /api/sheets
router.get('/', async (req, res) => {
  try {
    const sheetsList = await Sheet.find({}, 'name createdAt updatedAt'); // Only retrieve name and timestamps
    res.status(200).json(sheetsList);
  } catch (err) {
    console.error('Error fetching sheets list:', err);
    res.status(500).json({ message: 'Server error fetching list.' });
  }
});

// 4. (Optional) Delete a spreadsheet
// DELETE /api/sheets/:name
router.delete('/:name', async (req, res) => {
  try {
    const result = await Sheet.deleteOne({ name: req.params.name });
    if (result.deletedCount === 0) {
      return res.status(404).json({ message: 'Spreadsheet not found.' });
    }
    res.status(200).json({ message: 'Spreadsheet deleted successfully.' });
  } catch (err) {
    console.error('Error deleting sheet:', err);
    res.status(500).json({ message: 'Server error during delete.' });
  }
});

module.exports = router;