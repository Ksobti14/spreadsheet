const mongoose = require('mongoose');

const sheetschema = new mongoose.Schema({
  name: {
    type: String,
    required: true,
    unique: true,
    trim: true
  },
  worksheets: {
    type: mongoose.Schema.Types.Mixed, // Allows flexible/nested data
    required: true
  },
  createdAt: {
    type: Date,
    default: Date.now
  },
  updatedAt: {
    type: Date,
    default: Date.now
  }
});

sheetschema.pre('save', function (next) {
  this.updatedAt = Date.now();
  next();
});

module.exports = mongoose.model('Sheet', sheetschema);
