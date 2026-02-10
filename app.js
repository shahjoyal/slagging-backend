// app.js - unified server for both sites (Express + MongoDB)
// Install dependencies:
// npm i express body-parser mongoose multer xlsx exceljs dotenv cors express-session bcryptjs

const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');
const xlsx = require('xlsx');
const multer = require('multer');
const mongoose = require('mongoose');
const ExcelJS = require('exceljs');
const cors = require('cors');
const session = require('express-session');
const bcrypt = require('bcryptjs');
const fs = require("fs");
const { RandomForestRegression } = require("ml-random-forest");

require('dotenv').config();



const app = express();

const modelJSON = JSON.parse(fs.readFileSync("aft_model.json"));
const model = RandomForestRegression.load(modelJSON);

// load preprocessing cutoffs saved during training (optional)
let preprocessing = null;
try {
  preprocessing = JSON.parse(fs.readFileSync('preprocessing.json'));
  console.log('Loaded preprocessing.json');
  // sanity-check shape
  if (!Array.isArray(preprocessing.lower) || !Array.isArray(preprocessing.upper)) {
    console.warn('preprocessing.json missing expected lower/upper arrays — ignoring');
    preprocessing = null;
  }
} catch (err) {
  console.warn('preprocessing.json not found or invalid — will skip clipping:', String(err));
  preprocessing = null;
}

// ---------- CONFIG ----------
const MONGODB_URI = process.env.MONGODB_URI || 'YOUR_MONGODB_URI_HERE';
const PORT = process.env.PORT || 5000; // default 5000 to match your fetch
console.log('MONGODB_URI config status:', (MONGODB_URI && !MONGODB_URI.includes('YOUR_MONGODB_URI_HERE')) ? 'using env' : 'MONGODB_URI not set - edit .env or app.js');

// ---------- MIDDLEWARE ----------
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static('js'));

// If your frontend is served from a different origin and you use credentials (sessions),
// configure CORS accordingly. For simple same-origin setups, default cors() is okay.
app.use(cors()); // dev-friendly; tighten origin & credentials in production

// If behind a reverse proxy (nginx, load balancer), enable trust proxy so req.ip and x-forwarded-for behave correctly.
// If unsure and you are deploying behind a proxy, set true.
app.set('trust proxy', true);

// ---------- SESSION ----------
app.use(session({
  secret: process.env.SESSION_SECRET || 'please_change_this_to_a_strong_secret',
  resave: false,
  saveUninitialized: false,
  cookie: {
    httpOnly: true,
    // maxAge = 7 days (adjust as needed)
    maxAge: 7 * 24 * 3600 * 1000
    // In production use secure: true if using HTTPS
  }
}));

// ---------- MONGOOSE CONNECT ----------
mongoose.connect(MONGODB_URI, {
  useNewUrlParser: true,
  useUnifiedTopology: true,
  serverSelectionTimeoutMS: 30000
})
.then(() => console.log('MongoDB connected successfully'))
.catch(err => {
  console.error('MongoDB connection error:', err);
});

// ---------- SCHEMA & MODELS ----------
// flexible generic schema for existing collections
const flexibleSchema = new mongoose.Schema({}, { strict: false });
// const SlaggingData = mongoose.models.SlaggingData || mongoose.model('SlaggingData', flexibleSchema);

// Force Coal model to use 'coals' collection (as in your DB)
let Coal;
try {
  Coal = mongoose.model('Coal');
} catch (e) {
  Coal = mongoose.model('Coal', flexibleSchema, 'coals'); // explicit collection name
}
console.log('Coal collection name:', Coal.collection && Coal.collection.name);

// ---------- USER MODEL (for trials / subscription) ----------
const userSchema = new mongoose.Schema({
  email: { type: String, required: true, unique: true, index: true },
  passwordHash: { type: String, required: true },
  trialsLeft: { type: Number, default: 5 },            // number of remaining trials
  lockedUntil: { type: Date, default: null },          // when lock expires
  lastIP: { type: String, default: null },             // last known IP
  ipHistory: [{ ip: String, when: Date }],             // optional history
  createdAt: { type: Date, default: Date.now }
}, { timestamps: true });

const User = mongoose.models.User || mongoose.model('User', userSchema);

// ---------- HELPERS ----------
function excelSerialToDate(serial) {
  const excelEpoch = new Date(Date.UTC(1900, 0, 1));
  const daysOffset = serial - 1;
  const date = new Date(excelEpoch.getTime() + daysOffset * 24 * 60 * 60 * 1000);
  return date.toISOString().split('T')[0];
}

/**
 * normalizeCoalDoc: converts a DB document (various field-naming variants)
 * into a canonical object expected by frontends:
 */
function normalizeCoalDoc(raw) {
  if (!raw) return null;
  const o = (raw.toObject ? raw.toObject() : Object.assign({}, raw));
  const id = String(o._id || o.id || '');

  const coalName = o.coal || o.name || o['Coal source name'] || o['Coal Source Name'] || '';
  const transportId = o['Transport ID'] || o.transportId || o.transport_id || null;

  const canonicalKeys = ['SiO2','Al2O3','Fe2O3','CaO','MgO','Na2O','K2O','TiO2','SO3','P2O5','Mn3O4','Sulphur (S)','GCV'];

  const aliasMap = {
    'SiO2': 'SiO2', 'SiO₂': 'SiO2',
    'Al2O3': 'Al2O3', 'Al₂O₃': 'Al2O3',
    'Fe2O3': 'Fe2O3', 'Fe₂O₃': 'Fe2O3',
    'CaO': 'CaO',
    'MgO': 'MgO',
    'Na2O': 'Na2O',
    'K2O': 'K2O',
    'TiO2': 'TiO2', 'TiO₂': 'TiO2',
    'SO3': 'SO3', 'SO₃': 'SO3',
    'P2O5': 'P2O5', 'P₂O₅': 'P2O5',
    'Mn3O4': 'Mn3O4', 'Mn₃O₄': 'Mn3O4',
    'Sulphur (S)': 'Sulphur (S)',
    'SulphurS': 'Sulphur (S)', 'Sulphur': 'Sulphur (S)', 'S': 'Sulphur (S)',
    'GCV': 'GCV', 'Gcv': 'GCV', 'gcv': 'GCV'
  };

  const properties = {};
  canonicalKeys.forEach(k => properties[k] = null);

  function collectFrom(obj) {
    if (!obj) return;
    Object.keys(obj).forEach(k => {
      const trimmed = String(k).trim();
      let mapped = aliasMap[trimmed] || null;
      if (!mapped) {
        const normalizedKey = trimmed.replace(/₂/g,'2').replace(/₃/g,'3').replace(/₄/g,'4');
        mapped = aliasMap[normalizedKey] || null;
      }
      if (mapped) {
        const val = obj[k];
        properties[mapped] = (val === '' || val === null || val === undefined) ? null : (isNaN(Number(val)) ? val : Number(val));
      }
    });
  }

  collectFrom(o);
  if (o.properties && typeof o.properties === 'object') collectFrom(o.properties);

  if ((properties['GCV'] === null || properties['GCV'] === undefined) && (o.gcv || o.GCV || o.Gcv)) {
    properties['GCV'] = o.gcv || o.GCV || o.Gcv;
  }

  Object.keys(properties).forEach(k => {
    const v = properties[k];
    if (v !== null && v !== undefined && !isNaN(Number(v))) properties[k] = Number(v);
  });

  const gcvVal = properties['GCV'];

  return {
    _id: o._id,
    id,
    coal: coalName,
    coalType: coalName,
    transportId,
    gcv: gcvVal,
    properties
  };
}

// IP helper (works with X-Forwarded-For when trust proxy=true)
function getClientIp(req) {
  const xff = req.headers['x-forwarded-for'];
  if (xff) return xff.split(',')[0].trim();
  return req.ip || (req.connection && req.connection.remoteAddress) || null;
}

// Auth middleware
async function requireAuth(req, res, next) {
  try {
    const uid = req.session && req.session.userId;
    if (!uid) return res.status(401).json({ error: 'Not authenticated' });
    const user = await User.findById(uid);
    if (!user) {
      req.session.destroy?.(()=>{});
      return res.status(401).json({ error: 'User not found' });
    }
    // check lock
    if (user.lockedUntil && user.lockedUntil > new Date()) {
      return res.status(403).json({ error: 'Account locked until ' + user.lockedUntil.toISOString() });
    }
    req.currentUser = user;
    next();
  } catch (err) {
    console.error('auth error', err);
    res.status(500).json({ error: 'Authentication error' });
  }
}

// // ---------- ROUTES ----------
// // root route - keep existing login page
// app.get('/', (req, res) => {
//   res.sendFile(path.join(__dirname, 'public', 'login.html'));
// });

// // download template
// // download template (coal-oriented)
// // download template (coal-oriented) — optional data export with ?includeData=true
// app.get("/download-template", async (req, res) => {
//   try {
//     const includeData = String(req.query.includeData || '').toLowerCase() === 'true';

//     const workbook = new ExcelJS.Workbook();
//     const worksheet = workbook.addWorksheet("Coal Upload Template");

//     const instructionText = `Instruction for filling the sheet:
// 1. Row 1 contains instructions (delete this row when uploading).
// 2. Row 2 must be headers. Required header: "Coal" (name).
// 3. Other helpful headers: SiO2, Al2O3, Fe2O3, CaO, MgO, Na2O, K2O, TiO2, SO3, P2O5, Mn3O4, Sulphur, GCV, Cost, Transport ID, Shipment date.
// 4. Leave empty any missing values. Save as .xlsx and upload using the 'file' field.`;

//     worksheet.mergeCells('A1:Q1');
//     const instructionCell = worksheet.getCell('A1');
//     instructionCell.value = instructionText;
//     instructionCell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
//     instructionCell.font = { bold: true };
//     instructionCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFEFEF' } };
//     worksheet.getRow(1).height = 90;

//     // Header row (row index 2)
//     const headers = [
//       "Coal", "SiO2", "Al2O3", "Fe2O3", "CaO", "MgO", "Na2O", "K2O",
//       "TiO2", "SO3", "P2O5", "Mn3O4", "Sulphur", "GCV", "Cost", "Transport ID", "Shipment date"
//     ];
//     worksheet.addRow(headers);
//     const headerRow = worksheet.getRow(2);
//     headerRow.font = { bold: true };
//     headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
//     headerRow.eachCell((cell) => {
//       cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB0E0E6' } };
//       cell.border = {
//         top: { style: 'thin' }, left: { style: 'thin' },
//         bottom: { style: 'thin' }, right: { style: 'thin' }
//       };
//     });

//     // Set reasonable column widths
//     headers.forEach((_, index) => worksheet.getColumn(index + 1).width = 18);

//     if (includeData) {
//       // Fetch docs from coals collection
//       const docs = await Coal.find({}, { __v: 0 }).lean().exec();

//       // Helper to choose existing field variants
//       function pick(o, ...keys) {
//         for (const k of keys) {
//           if (o && o[k] !== undefined && o[k] !== null) return o[k];
//         }
//         return '';
//       }

//       // Append each DB row into the sheet in the same header order
//       for (const d of docs) {
//         const rowValues = [
//           // Coal name
//           pick(d, 'coal', 'Coal', 'name', 'Coal source name'),
//           // oxides & properties — try common canonical and alternate keys
//           pick(d, 'SiO2', 'SiO₂', 'SiO 2'),
//           pick(d, 'Al2O3', 'Al₂O₃'),
//           pick(d, 'Fe2O3', 'Fe₂O₃'),
//           pick(d, 'CaO'),
//           pick(d, 'MgO'),
//           pick(d, 'Na2O'),
//           pick(d, 'K2O'),
//           pick(d, 'TiO2', 'TiO₂'),
//           pick(d, 'SO3', 'SO₃'),
//           pick(d, 'P2O5', 'P₂O₅'),
//           pick(d, 'Mn3O4', 'Mn₃O₄', 'MN3O4'),
//           // Sulphur field may be SulphurS / Sulphur / S / Sulphur (S)
//           pick(d, 'SulphurS', 'Sulphur (S)', 'Sulphur', 'S'),
//           // GCV/gcv
//           pick(d, 'GCV', 'gcv', 'Gcv'),
//           // cost
//           pick(d, 'cost', 'Cost'),
//           // transport id and shipment date
//           pick(d, 'Transport ID', 'transportId', 'transport_id'),
//           // for shipment date prefer ISO string if Date object or string
//           (() => {
//             const sd = pick(d, 'shipmentDate', 'Shipment date', 'shipment_date');
//             if (!sd) return '';
//             if (sd instanceof Date) return sd.toISOString().split('T')[0];
//             // if mongo stores as object like {"$date": "..."} or as string, coerce to string
//             return String(sd);
//           })()
//         ];
//         worksheet.addRow(rowValues);
//       }
//     }

//     res.setHeader("Content-Disposition", `attachment; filename=${includeData ? 'Coal_Data_Export.xlsx' : 'Coal_Upload_Template.xlsx'}`);
//     res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//     await workbook.xlsx.write(res);
//     res.end();
//   } catch (err) {
//     console.error('/download-template error:', err);
//     res.status(500).send('Template generation failed');
//   }
// });



// // multer memory storage for uploads
// const storage = multer.memoryStorage();
// const upload = multer({ storage });

// // upload excel -> SlaggingData collection
// // upload excel -> insert into 'coals' collection
// // upload excel -> insert into 'coals' collection (robust header detection & normalization)
// app.post("/upload-excel", upload.single("file"), async (req, res) => {
//   if (!req.file) return res.status(400).json({ error: "No file uploaded" });
//   try {
//     const workbook = xlsx.read(req.file.buffer, { type: "buffer" });
//     if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
//       return res.status(400).json({ error: "No sheets found in workbook" });
//     }

//     const sheetName = workbook.SheetNames[0];
//     const sheet = workbook.Sheets[sheetName];

//     // Read as rows array so we can detect header row robustly (handles instruction/merged header rows)
//     const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: null });

//     // find header row index by looking for typical header tokens
//     const headerRowIndex = rows.findIndex(r => Array.isArray(r) && r.some(cell => {
//       if (!cell) return false;
//       const s = String(cell).toLowerCase();
//       return /coal|sio2|sio₂|al2o3|gcv|sulphur|si o2|al₂o₃|fe2o3/.test(s);
//     }));

//     if (headerRowIndex === -1) {
//       return res.status(400).json({ error: "Could not find header row in sheet. Ensure headers like 'Coal' or 'SiO2' exist." });
//     }

//     const rawHeaders = rows[headerRowIndex].map(h => (h === null || h === undefined) ? '' : String(h).trim());
//     const dataRows = rows.slice(headerRowIndex + 1);

//     // header map (variants -> canonical keys)
//     const headerMap = {
//       'coal': 'coal', 'coal source name': 'coal', 'coal source': 'coal', 'name': 'coal',

//       'sio2': 'SiO2', 'sio₂': 'SiO2', 'si o2': 'SiO2',
//       'al2o3': 'Al2O3', 'al₂o₃': 'Al2O3',
//       'fe2o3': 'Fe2O3', 'fe₂o₃': 'Fe2O3',
//       'cao': 'CaO', 'mgo': 'MgO',
//       'na2o': 'Na2O', 'k2o': 'K2O',
//       'tio2': 'TiO2', 'tio₂': 'TiO2',
//       'so3': 'SO3', 'so₃': 'SO3',
//       'p2o5': 'P2O5', 'p₂o₅': 'P2O5',
//       'mn3o4': 'Mn3O4', 'mn₃o₄': 'Mn3O4',

//       'sulphur (s)': 'SulphurS', 'sulphur': 'SulphurS', 'sulphurs': 'SulphurS', 's': 'SulphurS',
//       'gcv': 'gcv', 'gcv.': 'gcv', 'g c v': 'gcv',
//       'cost': 'cost', 'price': 'cost',
//       'transport id': 'Transport ID', 'data uploaded by tps': 'uploadedBy', 'shipment date': 'shipmentDate', 'type of transport': 'transportType'
//     };

//     // helper to canonicalize header string
//     function canonicalHeader(h) {
//       if (h === null || h === undefined) return '';
//       const s = String(h).trim();
//       const simple = s.replace(/[\s_\-\.]/g, '').replace(/₂/g,'2').replace(/₃/g,'3').replace(/₄/g,'4').toLowerCase();
//       // try exact headerMap keys first
//       const direct = Object.keys(headerMap).find(k => k.toLowerCase() === s.toLowerCase());
//       if (direct) return headerMap[direct];
//       const found = Object.keys(headerMap).find(k => k.replace(/[\s_\-\.]/g,'').replace(/₂/g,'2').replace(/₃/g,'3').replace(/₄/g,'4').toLowerCase() === simple);
//       return found ? headerMap[found] : s; // fallback to original header text if not found
//     }

//     // build canonical headers array
//     const canonicalHeaders = rawHeaders.map(h => canonicalHeader(h));

//     // map rows to objects
//     const parsed = dataRows.map((row, rowIndex) => {
//       // skip completely empty rows
//       if (!Array.isArray(row) || row.every(c => c === null || (typeof c === 'string' && c.trim() === ''))) return null;

//       const out = {};
//       for (let i = 0; i < canonicalHeaders.length; i++) {
//         const key = canonicalHeaders[i];
//         // skip empty header slots
//         if (!key) continue;
//         let val = row[i] === undefined ? null : row[i];

//         // convert excel date serials if header indicates date (optional)
//         if (key.toLowerCase().includes('date') && typeof val === 'number' && val > 0 && val < 2958465) {
//           val = excelSerialToDate(val);
//         } else if (val === '') {
//           val = null;
//         }

//         // convert numeric-like strings to Number
//         if (val !== null && typeof val !== 'number') {
//           const maybeNum = Number(String(val).replace(/,/g, '').trim());
//           if (!Number.isNaN(maybeNum)) val = Math.round(maybeNum * 100) / 100;
//         }

//         out[key] = val;
//       }

//       // require coal name
//       if (!out.coal || String(out.coal).trim() === '') return null;
//       return out;
//     }).filter(Boolean);

//     if (!parsed.length) {
//       return res.status(400).json({ error: "No valid data rows found after header (check your Excel file)" });
//     }

//     // Insert into coals collection; unordered so one bad row won't stop others
//     const inserted = await Coal.insertMany(parsed, { ordered: false });
//     res.json({ message: "Data uploaded successfully", rowsParsed: parsed.length, rowsInserted: inserted.length, sample: inserted.slice(0,5) });

//   } catch (error) {
//     console.error("Error processing file (upload-excel):", error);
//     res.status(500).json({ error: "Failed to process file", details: String(error) });
//   }
// });



// // fetch raw SlaggingData
// // fetch all coals
// app.get("/fetch-data", async (req, res) => {
//   try {
//     const data = await Coal.find({}, { __v: 0 }).lean();
//     res.json(data);
//   } catch (error) {
//     console.error("Error fetching data:", error);
//     res.status(500).json({ error: "Failed to fetch data" });
//   }
// });

// // delete route -> remove docs from 'coals'
// app.delete("/delete-data", async (req, res) => {
//   try {
//     const { ids } = req.body;
//     if (!Array.isArray(ids) || ids.length === 0) return res.status(400).json({ error: "No IDs provided" });

//     const result = await Coal.deleteMany({ _id: { $in: ids } });
//     if (result.deletedCount === 0) return res.status(404).json({ error: "No data found" });
//     res.json({ message: `${result.deletedCount} data deleted successfully` });
//   } catch (error) {
//     console.error("Error deleting data:", error);
//     res.status(500).json({ error: "Failed to delete data" });
//   }
// });


function calculateAFT(values) {
  // values: array of 9 oxide numbers in order [SiO2, Al2O3, Fe2O3, CaO, MgO, Na2O, K2O, SO3, TiO2]
  try {
    // Build and preprocess features the SAME way as training
    const features = buildFeaturesForModel(values);
    const prepped = applyPreprocessing(features);

    // predict expects a 2D array (rows)
    const pred = model.predict([prepped])[0];

    // safety clamp to reasonable physical bounds (adjust if needed)
    const clamped = Math.max(1000, Math.min(1600, Math.round(pred)));

    return clamped;
  } catch (err) {
    console.warn('ML prediction failed — falling back to formulaAFT. Error:', String(err));
    // fallback to pure formula if ML throws
    return Math.round(formulaAFT(values));
  }
}

function formulaAFT(values) {
  const [SiO2, Al2O3, Fe2O3, CaO, MgO, Na2O, K2O, SO3, TiO2] = values;
  const sumSiAl = SiO2 + Al2O3;

  if (sumSiAl < 55) {
    return 1245 +
      1.1 * SiO2 +
      0.95 * Al2O3 -
      2.5 * Fe2O3 -
      2.98 * CaO -
      4.5 * MgO -
      7.89 * (Na2O + K2O) -
      1.7 * SO3 -
      0.63 * TiO2;
  } else if (sumSiAl < 75) {
    return 1323 +
      1.45 * SiO2 +
      0.683 * Al2O3 -
      2.39 * Fe2O3 -
      3.1 * CaO -
      4.5 * MgO -
      7.49 * (Na2O + K2O) -
      2.1 * SO3 -
      0.63 * TiO2;
  } else {
    return 1395 +
      1.2 * SiO2 +
      0.9 * Al2O3 -
      2.5 * Fe2O3 -
      3.1 * CaO -
      4.5 * MgO -
      7.2 * (Na2O + K2O) -
      1.7 * SO3 -
      0.63 * TiO2;
  }
}

function buildFeatures(values) {
  const [
    SiO2,
    Al2O3,
    Fe2O3,
    CaO,
    MgO,
    Na2O,
    K2O,
    SO3,
    TiO2,
  ] = values;

  const SiAl = SiO2 + 0.8 * Al2O3;   
  const Flux = CaO + MgO + Fe2O3;    
  const Alk  = Na2O + K2O;            

  return [
    SiAl,
    Flux,
    Alk,
    SO3,
    TiO2,
  ];
}

// Build the SAME feature vector used in training:
// [SiO2, Al2O3, Fe2O3, CaO, MgO, Na2O, K2O, SO3, TiO2, baseAFT]
function buildFeaturesForModel(values) {
  const [SiO2, Al2O3, Fe2O3, CaO, MgO, Na2O, K2O, SO3, TiO2] = values;
  const base = formulaAFT(values);
  return [
    SiO2, Al2O3, Fe2O3, CaO, MgO,
    Na2O, K2O, SO3, TiO2, base
  ];
}

// Apply 1%/99% winsorization cutoffs saved during training.
// If preprocessing.json not present, return features unchanged.
function applyPreprocessing(features) {
  if (!preprocessing) return features;
  const out = features.slice();
  for (let i = 0; i < out.length; i++) {
    const lo = preprocessing.lower[i];
    const hi = preprocessing.upper[i];
    if (typeof lo === 'number' && out[i] < lo) out[i] = lo;
    if (typeof hi === 'number' && out[i] > hi) out[i] = hi;
  }
  return out;
}


app.post("/calculate-aft", async (req, res) => {
  try {
    const values = req.body && req.body.values;
    if (!Array.isArray(values) || values.length < 9) {
      return res.status(400).json({ error: "Please send values: array of 9 oxides [SiO2,Al2O3,Fe2O3,CaO,MgO,Na2O,K2O,SO3,TiO2]" });
    }

    // build features (includes base formula as 10th feature)
    const features = buildFeaturesForModel(values);
    const prepped = applyPreprocessing(features);

    console.log("Base AFT (formula):", features[9]);
    console.log("Features for ML model:", prepped);

    // Predict using the model (2D input)
    let predicted = null;
    try {
      predicted = model.predict([prepped])[0];
    } catch (err) {
      console.warn('Model.predict failed:', String(err));
      predicted = features[9]; // fallback: use base formula if model fails
    }

    // clamp and round
    predicted = Math.round(Math.max(1000, Math.min(1600, predicted)));

    console.log("Final AFT Prediction:", predicted);
    return res.status(200).json({ prediction: predicted, method: 'ml+base' });
  } catch (err) {
    console.error('/calculate-aft error:', err);
    return res.status(500).json({ error: 'Internal error calculating AFT', details: String(err) });
  }
});


// // ---------- AUTH ROUTES ----------

// // POST /auth/login  { email, password }
// app.post('/auth/login', async (req, res) => {
//   try {
//     const { email, password } = req.body;
//     if (!email || !password) return res.status(400).json({ error: 'Email & password required' });

//     const user = await User.findOne({ email: email.toLowerCase().trim() });
//     if (!user) return res.status(401).json({ error: 'Invalid credentials' });

//     // check locked
//     if (user.lockedUntil && user.lockedUntil > new Date()) {
//       return res.status(403).json({ error: 'Account locked until ' + user.lockedUntil.toISOString() });
//     }

//     const ok = await bcrypt.compare(password, user.passwordHash);
//     if (!ok) return res.status(401).json({ error: 'Invalid credentials' });

//     // update last IP
//     const ip = getClientIp(req);
//     user.lastIP = ip;
//     user.ipHistory = user.ipHistory || [];
//     user.ipHistory.push({ ip, when: new Date() });
//     await user.save();

//     // create session
//     req.session.userId = user._id.toString();

//     res.json({ message: 'Logged in', trialsLeft: user.trialsLeft });
//   } catch (err) {
//     console.error('/auth/login error', err);
//     res.status(500).json({ error: 'Login failed' });
//   }
// });

// // POST /auth/logout
// app.post('/auth/logout', (req, res) => {
//   req.session.destroy(err => {
//     if (err) console.error('session destroy err', err);
//     res.json({ message: 'Logged out' });
//   });
// });

// // GET /auth/status - returns basic auth info
// app.get('/auth/status', async (req, res) => {
//   try {
//     if (!req.session || !req.session.userId) return res.json({ authenticated: false });
//     const user = await User.findById(req.session.userId, 'email trialsLeft lockedUntil lastIP');
//     if (!user) return res.json({ authenticated: false });
//     return res.json({
//       authenticated: true,
//       email: user.email,
//       trialsLeft: user.trialsLeft,
//       lockedUntil: user.lockedUntil,
//       lastIP: user.lastIP
//     });
//   } catch (err) {
//     console.error('/auth/status error', err);
//     res.status(500).json({ error: 'Status check failed' });
//   }
// });


// ---------- ROUTES TO REPLACE ----------

// root route
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// DOWNLOAD TEMPLATE (optionally include data via ?includeData=true)
app.get("/download-template", async (req, res) => {
  try {
    const includeData = String(req.query.includeData || '').toLowerCase() === 'true';
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Coal Upload Template");

    const instructionText = `Instructions:
1) Delete row 1 (instructions) before uploading.
2) Row 2 is headers. Required header: "Coal".
3) Optional header: "Coal ID" (or "Coal Source ID") to supply your own IDs.
4) Other example headers: SiO2, Al2O3, Fe2O3, CaO, MgO, Na2O, K2O, TiO2, SO3, P2O5, Mn3O4, Sulphur, GCV, Cost, Transport ID, Shipment date.`;
    worksheet.mergeCells('A1:R1');
    const instructionCell = worksheet.getCell('A1');
    instructionCell.value = instructionText;
    instructionCell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getRow(1).height = 80;

    // header row - include Coal ID as first header
    const headers = [
      "Coal ID", "Coal", "SiO2", "Al2O3", "Fe2O3", "CaO", "MgO", "Na2O", "K2O",
      "TiO2", "SO3", "P2O5", "Mn3O4", "SulphurS", "gcv", "Cost"
    ];
    worksheet.addRow(headers);
    const headerRow = worksheet.getRow(2);
    headerRow.font = { bold: true };
    headerRow.alignment = { horizontal: 'center' };

    headers.forEach((_, i) => worksheet.getColumn(i + 1).width = 18);

    if (includeData) {
      const docs = await Coal.find({}, { __v: 0 }).lean().exec();
      function pick(o, ...keys) {
        for (const k of keys) if (o && o[k] !== undefined && o[k] !== null) return o[k];
        return '';
      }
      for (const d of docs) {
        const rowValues = [
          pick(d, 'coalId', 'Coal ID'),
          pick(d, 'coal', 'Coal', 'name'),
          pick(d, 'SiO2'), pick(d, 'Al2O3'), pick(d, 'Fe2O3'), pick(d, 'CaO'), pick(d, 'MgO'),
          pick(d, 'Na2O'), pick(d, 'K2O'), pick(d, 'TiO2'), pick(d, 'SO3'), pick(d, 'P2O5'),
          pick(d, 'Mn3O4'), pick(d, 'Sulphur', 'S'), pick(d, 'GCV'), pick(d, 'cost'),
          pick(d, 'Transport ID','transportId'),
          (d.shipmentDate instanceof Date) ? d.shipmentDate.toISOString().split('T')[0] : (d.shipmentDate || '')
        ];
        worksheet.addRow(rowValues);
      }
    }

    res.setHeader("Content-Disposition", `attachment; filename=${includeData ? 'Coal_Data_Export.xlsx' : 'Coal_Upload_Template.xlsx'}`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('/download-template error:', err);
    res.status(500).send('Template generation failed');
  }
});

// ---------- MULTER (file upload) SETUP ----------
const storage = multer.memoryStorage();          // keep file in memory buffer
const upload = multer({
  storage,
  limits: { fileSize: 50 * 1024 * 1024 }        // optional: limit to 50MB (adjust if needed)
});


// UPLOAD EXCEL
// app.post("/upload-excel", upload.single("file"), async (req, res) => {
//   if (!req.file) return res.status(400).json({ error: "No file uploaded" });
//   try {
//     const workbook = xlsx.read(req.file.buffer, { type: "buffer" });
//     if (!workbook.SheetNames || workbook.SheetNames.length === 0) return res.status(400).json({ error: "No sheets found" });
//     const sheet = workbook.Sheets[workbook.SheetNames[0]];
//     const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: null });

//     // find header row index by common tokens
//     const headerRowIndex = rows.findIndex(r => Array.isArray(r) && r.some(c => c && /coal|sio2|gcv|sulphur/i.test(String(c))));
//     if (headerRowIndex === -1) return res.status(400).json({ error: "Could not find header row. Ensure headers like 'Coal' exist." });

//     const rawHeaders = rows[headerRowIndex].map(h => (h === null || h === undefined) ? '' : String(h).trim());
//     const dataRows = rows.slice(headerRowIndex + 1);

//     // header mapping including coalId variants
//     const headerMap = {
//       'coal': 'coal', 'coal source name': 'coal', 'name': 'coal',
//       'coalid': 'coalId', 'coal id': 'coalId', 'coal_id': 'coalId', 'coalsourceid': 'coalId', 'coal source id': 'coalId',
//       'sio2': 'SiO2','al2o3':'Al2O3','fe2o3':'Fe2O3','cao':'CaO','mgo':'MgO',
//       'na2o':'Na2O','k2o':'K2O','tio2':'TiO2','so3':'SO3','p2o5':'P2O5','mn3o4':'Mn3O4',
//       'sulphur':'Sulphur', 's': 'Sulphur', 'gcv':'GCV', 'cost':'cost', 'transport id':'Transport ID', 'shipment date':'shipmentDate'
//     };

//     function canonicalHeader(s) {
//       if (!s) return '';
//       const simple = String(s).replace(/[\s_\-\.]/g,'').replace(/[₂₃₄]/g, m => ({'₂':'2','₃':'3','₄':'4'}[m])).toLowerCase();
//       const direct = Object.keys(headerMap).find(k => k.toLowerCase() === String(s).toLowerCase());
//       if (direct) return headerMap[direct];
//       const found = Object.keys(headerMap).find(k => k.replace(/[\s_\-\.]/g,'').toLowerCase() === simple);
//       return found ? headerMap[found] : s;
//     }

//     const canonicalHeaders = rawHeaders.map(h => canonicalHeader(h));
//     const parsed = dataRows.map(row => {
//       if (!Array.isArray(row) || row.every(c => c === null || (typeof c === 'string' && c.trim() === ''))) return null;
//       const out = {};
//       for (let i = 0; i < canonicalHeaders.length; i++) {
//         const key = canonicalHeaders[i];
//         if (!key) continue;
//         let val = row[i] === undefined ? null : row[i];
//         if (key.toLowerCase().includes('date') && typeof val === 'number') val = excelSerialToDate(val);
//         if (val === '') val = null;
//         if (val !== null && typeof val !== 'number') {
//           const maybeNum = Number(String(val).replace(/,/g, '').trim());
//           if (!Number.isNaN(maybeNum)) val = Math.round(maybeNum * 100) / 100;
//         }
//         out[key] = val;
//       }
//       if (!out.coal || String(out.coal).trim() === '') return null;
//       return out;
//     }).filter(Boolean);

//     if (!parsed.length) return res.status(400).json({ error: "No valid data rows found after header" });

//     // assign sequential coalId when missing (continuing from max existing numeric coalId)
//     if (parsed.length) {
//       const existingMaxDoc = await Coal.find({ coalId: { $exists: true } }).sort({ coalId: -1 }).limit(1).lean().exec();
//       let nextCoalId = 1;
//       if (existingMaxDoc && existingMaxDoc.length) {
//         const candidate = Number(existingMaxDoc[0].coalId);
//         if (Number.isFinite(candidate)) nextCoalId = Math.max(1, Math.floor(candidate) + 1);
//       }
//       parsed.forEach(r => {
//         if (r.coalId === undefined || r.coalId === null || String(r.coalId).trim() === '') {
//           r.coalId = String(nextCoalId);
//           nextCoalId++;
//         } else {
//           r.coalId = String(r.coalId).trim();
//         }
//       });
//     }

//     const inserted = await Coal.insertMany(parsed, { ordered: false });
//     res.json({ message: "Data uploaded successfully", rowsParsed: parsed.length, rowsInserted: inserted.length, sample: inserted.slice(0,5) });
//   } catch (error) {
//     console.error("Error processing file (upload-excel):", error);
//     res.status(500).json({ error: "Failed to process file", details: String(error) });
//   }
// });
// REPLACE the existing /upload-excel route with this block
// REPLACE your existing /upload-excel route with this block
app.post("/upload-excel", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file uploaded" });

  try {
    const workbook = xlsx.read(req.file.buffer, { type: "buffer" });
    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
      return res.status(400).json({ error: "No sheets found in workbook" });
    }

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: null });

    // find header row index (robust: looks for 'coal' token)
    const headerRowIndex = rows.findIndex(r => Array.isArray(r) && r.some(c => c && /coal/i.test(String(c))));
    if (headerRowIndex === -1) {
      return res.status(400).json({ error: "Could not find header row. Ensure 'Coal' header exists." });
    }

    const rawHeaders = rows[headerRowIndex].map(h => (h === null || h === undefined) ? '' : String(h).trim());
    const dataRows = rows.slice(headerRowIndex + 1);

    // Canonical DB keys we will allow (strict: unknown headers are ignored)
    const allowedCanonicalKeys = [
      'coalId','coal',
      'SiO2','Al2O3','Fe2O3','CaO','MgO','Na2O','K2O','TiO2',
      'SO3','P2O5','Mn3O4',
      'Sulphur','GCV','cost'
    ];

    // Header alias map -> canonical DB key
    const headerMap = {
      'coal': 'coal',
      'coalid': 'coalId', 'coal id': 'coalId', 'coal_id': 'coalId',
      'sio2': 'SiO2', 'si o2': 'SiO2', 'si02': 'SiO2', 'si o₂': 'SiO2',
      'al2o3': 'Al2O3', 'al₂o₃': 'Al2O3',
      'fe2o3': 'Fe2O3', 'fe₂o₃': 'Fe2O3',
      'cao': 'CaO', 'mgo': 'MgO',
      'na2o': 'Na2O', 'k2o': 'K2O',
      'tio2': 'TiO2', 'tio₂': 'TiO2',
      'so3': 'SO3', 'so₃': 'SO3',
      'p2o5': 'P2O5', 'p₂o₅': 'P2O5',
      'mn3o4': 'Mn3O4', 'mn₃o₄': 'Mn3O4',
      'sulphur': 'Sulphur', 'sulphurs': 'Sulphur', 's': 'Sulphur',
      'gcv': 'GCV', 'g c v': 'GCV', 'g.c.v.': 'GCV',
      'cost': 'cost', 'price': 'cost', 'rate': 'cost'
    };

    function normalizeHeaderString(s) {
      if (s === null || s === undefined) return '';
      let t = String(s).trim();
      t = t.replace(/₂/g, '2').replace(/₃/g, '3').replace(/₄/g, '4');
      t = t.toLowerCase().replace(/[\(\)\[\]\.,\/\\\-\_]/g, ' ').replace(/\s+/g, ' ').trim();
      return t;
    }

    function canonicalHeader(raw) {
      if (!raw && raw !== 0) return '';
      const norm = normalizeHeaderString(raw);

      if (headerMap.hasOwnProperty(norm)) {
        const mapped = headerMap[norm];
        if (allowedCanonicalKeys.includes(mapped)) return mapped;
        return '';
      }

      const compact = norm.replace(/\s+/g, '');
      const foundKey = Object.keys(headerMap).find(k => k.replace(/\s+/g,'') === compact);
      if (foundKey) {
        const mapped = headerMap[foundKey];
        if (allowedCanonicalKeys.includes(mapped)) return mapped;
      }

      const directAllowed = allowedCanonicalKeys.find(k => k.toLowerCase() === String(raw).trim().toLowerCase());
      if (directAllowed) return directAllowed;
      return '';
    }

    const canonicalHeaders = rawHeaders.map(h => canonicalHeader(h));

    if (!canonicalHeaders.includes('coal')) {
      return res.status(400).json({ error: "Required header 'Coal' not found. Ensure file has a 'Coal' column." });
    }

    // Parse rows into canonical-keyed objects (unknown columns skipped)
    const parsed = dataRows.map((row) => {
      if (!Array.isArray(row) || row.every(c => c === null || (typeof c === 'string' && c.trim() === ''))) return null;
      const out = {};
      for (let i = 0; i < canonicalHeaders.length; i++) {
        const key = canonicalHeaders[i];
        if (!key) continue;
        let val = row[i] === undefined ? null : row[i];

        // convert numeric-ish strings to Number (limit to 2 decimals)
        if (val !== null && typeof val !== 'number') {
          const maybeNum = Number(String(val).replace(/,/g, '').trim());
          if (!Number.isNaN(maybeNum)) val = Math.round(maybeNum * 100) / 100;
        }

        if (val === '') val = null;
        out[key] = val;
      }
      if (!out.coal || String(out.coal).trim() === '') return null;
      return out;
    }).filter(Boolean);

    if (!parsed.length) {
      return res.status(400).json({ error: "No valid data rows found after header (check your Excel rows)" });
    }

    // assign sequential coalId when missing (keep existing logic)
    const docs = parsed.slice();
    if (docs.length) {
      const existingMaxDoc = await Coal.find({ coalId: { $exists: true } }).sort({ coalId: -1 }).limit(1).lean().exec();
      let nextCoalId = 1;
      if (existingMaxDoc && existingMaxDoc.length) {
        const candidate = Number(existingMaxDoc[0].coalId);
        if (Number.isFinite(candidate)) nextCoalId = Math.max(1, Math.floor(candidate) + 1);
      }
      docs.forEach(r => {
        if (r.coalId === undefined || r.coalId === null || String(r.coalId).trim() === '') {
          r.coalId = String(nextCoalId);
          nextCoalId++;
        } else {
          r.coalId = String(r.coalId).trim();
        }
      });
    }

    // Final sanitize: keep only allowed canonical keys
    const sanitized = docs.map(r => {
      const o = {};
      Object.keys(r).forEach(k => {
        if (allowedCanonicalKeys.includes(k)) o[k] = r[k];
      });
      return o;
    });

    // --- NEW: map canonical keys to existing DB field names (so we don't create duplicates)
    // Build mapping canonicalKey -> actual DB field name (if present)
    const dbFieldMap = {}; // e.g. { 'Sulphur': 'SulphurS', 'GCV': 'gcv', ... }

    const sampleDoc = await Coal.findOne().lean(); // inspect one existing doc (if any)
    if (sampleDoc) {
      Object.keys(sampleDoc).forEach(fieldName => {
        // determine if this DB field corresponds to one of our canonical keys
        const can = canonicalHeader(fieldName);
        if (can && allowedCanonicalKeys.includes(can)) {
          dbFieldMap[can] = fieldName; // use the exact DB field name found
        }
      });
    }

    // ensure defaults requested by you: map Sulphur->SulphurS and GCV->gcv when not inferred
    if (!dbFieldMap['Sulphur']) dbFieldMap['Sulphur'] = 'SulphurS';
    if (!dbFieldMap['GCV']) dbFieldMap['GCV'] = 'gcv';

    // For all other canonical keys not mapped, use the canonical name as the DB field name
    allowedCanonicalKeys.forEach(k => {
      if (!dbFieldMap[k]) dbFieldMap[k] = k;
    });

    // transform sanitized docs to use DB field names (so we won't create 'Sulphur' or 'GCV' if mapping points to 'SulphurS'/'gcv')
    const finalDocs = sanitized.map(rowCanonical => {
      const obj = {};
      Object.keys(rowCanonical).forEach(canKey => {
        const targetField = dbFieldMap[canKey] || canKey;
        // If both targetField and the canonical name would collide, we prefer writing to targetField.
        obj[targetField] = rowCanonical[canKey];
      });
      return obj;
    });

    // Insert into DB
    const inserted = await Coal.insertMany(finalDocs, { ordered: false });

    return res.json({
      message: "Data uploaded successfully",
      rowsParsed: finalDocs.length,
      rowsInserted: inserted.length,
      fieldMappingPreview: dbFieldMap,
      sampleInserted: inserted.slice(0,5)
    });

  } catch (err) {
    console.error("Error in upload-excel (strict header handler + DB-mapping):", err);
    return res.status(500).json({ error: "Failed to process file", details: String(err) });
  }
});

// FETCH ALL DATA
app.get("/fetch-data", async (req, res) => {
  try {
    const data = await Coal.find({}, { __v: 0 }).lean();
    res.json(data);
  } catch (error) {
    console.error("Error fetching data:", error);
    res.status(500).json({ error: "Failed to fetch data" });
  }
});

// DELETE selected rows — supports Mongo _id (24-hex) and coalId
app.delete("/delete-data", async (req, res) => {
  try {
    const { ids } = req.body;
    if (!Array.isArray(ids) || ids.length === 0) return res.status(400).json({ error: "No IDs provided" });

    const objectIds = ids.filter(id => /^[0-9a-fA-F]{24}$/.test(String(id)));
    const coalIdValues = ids.filter(id => !/^[0-9a-fA-F]{24}$/.test(String(id)));

    let totalDeleted = 0;
    if (objectIds.length) {
      const r1 = await Coal.deleteMany({ _id: { $in: objectIds } });
      totalDeleted += (r1 && r1.deletedCount) ? r1.deletedCount : 0;
    }
    if (coalIdValues.length) {
      const r2 = await Coal.deleteMany({ coalId: { $in: coalIdValues } });
      totalDeleted += (r2 && r2.deletedCount) ? r2.deletedCount : 0;
    }

    if (totalDeleted === 0) return res.status(404).json({ error: "No data found" });
    res.json({ message: `${totalDeleted} data deleted successfully` });
  } catch (error) {
    console.error("Error deleting data:", error);
    res.status(500).json({ error: "Failed to delete data" });
  }
});

// AUTH ROUTES
app.post('/auth/login', async (req, res) => {
  try {
    const { email, password } = req.body;
    if (!email || !password) return res.status(400).json({ error: 'Email & password required' });
    const user = await User.findOne({ email: email.toLowerCase().trim() });
    if (!user) return res.status(401).json({ error: 'Invalid credentials' });

    if (user.lockedUntil && user.lockedUntil > new Date()) {
      return res.status(403).json({ error: 'Account locked until ' + user.lockedUntil.toISOString() });
    }
    const ok = await bcrypt.compare(password, user.passwordHash);
    if (!ok) return res.status(401).json({ error: 'Invalid credentials' });

    const ip = getClientIp(req);
    user.lastIP = ip;
    user.ipHistory = user.ipHistory || [];
    user.ipHistory.push({ ip, when: new Date() });
    await user.save();

    req.session.userId = user._id.toString();
    res.json({ message: 'Logged in' });
  } catch (err) {
    console.error('/auth/login error', err);
    res.status(500).json({ error: 'Login failed' });
  }
});

app.post('/auth/logout', (req, res) => {
  req.session.destroy(err => {
    if (err) console.error('session destroy err', err);
    res.json({ message: 'Logged out' });
  });
});

app.get('/auth/status', async (req, res) => {
  try {
    if (!req.session || !req.session.userId) return res.json({ authenticated: false });
    const user = await User.findById(req.session.userId, 'email trialsLeft lockedUntil lastIP');
    if (!user) return res.json({ authenticated: false });
    return res.json({
      authenticated: true,
      email: user.email,
      trialsLeft: user.trialsLeft,
      lockedUntil: user.lockedUntil,
      lastIP: user.lastIP
    });
  } catch (err) {
    console.error('/auth/status error', err);
    res.status(500).json({ error: 'Status check failed' });
  }
});



// ---------- OPTIMIZE ROUTE (modified to enforce trials & log IP) ----------
// app.post("/optimize", requireAuth, async (req, res) => {
//   try {
//     const user = req.currentUser;
//     // ensure not locked (requireAuth already checks but double-check)
//     if (user.lockedUntil && user.lockedUntil > new Date()) {
//       return res.status(403).json({ error: 'Account locked until ' + user.lockedUntil.toISOString() });
//     }

//     // check trials
//     if ((user.trialsLeft || 0) <= 0) {
//       // enforce lock for 24 hours starting now
//       user.lockedUntil = new Date(Date.now() + 24 * 3600 * 1000);
//       await user.save();
//       req.session.destroy(()=>{});
//       return res.status(403).json({ error: 'Trials exhausted. Account locked for 24 hours.' });
//     }

//     // record IP for this calculation call
//     const ip = getClientIp(req);
//     user.lastIP = ip;
//     user.ipHistory = user.ipHistory || [];
//     user.ipHistory.push({ ip, when: new Date() });

//     // ---- YOUR EXISTING OPTIMIZATION LOGIC STARTS HERE ----
//     const { blends } = req.body;
//     if (!blends || !Array.isArray(blends) || blends.length === 0) {
//         return res.status(400).json({ error: "Invalid blend data" });
//     }
//     const oxideCols = ['SiO2', 'Al2O3', 'Fe2O3', 'CaO', 'MgO', 'Na2O', 'K2O', 'SO3', 'TiO2'];
//     const coalNames = blends.map(b => b.coal);
//     const oxideValues = blends.map(b => oxideCols.map(col => b.properties[col] || 0));
//     const minMaxBounds = blends.map(b => [b.min, b.max]);
//     const costsPerTon = blends.map(b => b.cost);
//     const gcvValue = blends.map(b => b.properties.Gcv);
//     const individualCoalAFTs = oxideValues.map((vals, i) => ({
//       coal: coalNames[i],
//       predicted_aft: calculateAFT(vals)
//     }));
//     function* generateCombinations(bounds, step) {
//       function* helper(index, combo) {
//         if (index === bounds.length) {
//           const sum = combo.reduce((a, b) => a + b, 0);
//           if (sum === 100) yield combo;
//           return;
//         }
//         const [min, max] = bounds[index];
//         for (let i = min; i <= max; i += step) yield* helper(index + 1, [...combo, i]);
//       }
//       yield* helper(0, []);
//     }
//     const step = 1;
//     const validBlends = [];
//     for (const blend of generateCombinations(minMaxBounds, step)) {
//       const weights = blend.map(x => x / 100);
//       const blendedOxides = oxideCols.map((_, i) =>
//         oxideValues.reduce((sum, val, idx) => sum + val[i] * weights[idx], 0)
//       );
//       const predictedAFT = calculateAFT(blendedOxides);
//       const totalgcv = blend.reduce((sum, pct, i ) => sum + pct*gcvValue[i], 0) / 100;
//       const totalCost = blend.reduce((sum, pct, i) => sum + pct * costsPerTon[i], 0) / 100;
//       validBlends.push({ blend, predicted_aft: predictedAFT, cost: totalCost, gcv: totalgcv, blended_oxides: blendedOxides });
//     }
//     if (validBlends.length === 0) return res.status(404).json({ message: "No valid blends found" });
//     const aftVals = validBlends.map(b => b.predicted_aft);
//     const costVals = validBlends.map(b => b.cost);
//     const aftMin = Math.min(...aftVals);
//     const aftMax = Math.max(...aftVals);
//     const costMin = Math.min(...costVals);
//     const costMax = Math.max(...costVals);
//     const blendScores = validBlends.map((b, i) => {
//       const aftNorm = (b.predicted_aft - aftMin) / (aftMax - aftMin);
//       const costNorm = (costMax - b.cost) / (costMax - costMin);
//       return aftNorm + costNorm;
//     });
//     const bestAftBlend = validBlends[aftVals.indexOf(Math.max(...aftVals))];
//     const cheapestBlend = validBlends[costVals.indexOf(Math.min(...costVals))];
//     const balancedBlend = validBlends[blendScores.indexOf(Math.max(...blendScores))];
//     const currentWeights = blends.map(b => b.current / 100);
//     const currentBlendedOxides = oxideCols.map((_, i) =>
//       oxideValues.reduce((sum, val, idx) => sum + val[i] * currentWeights[idx], 0)
//     );
//     const currentAFT = calculateAFT(currentBlendedOxides);
//     const currentGCV = blends.reduce((sum, b, i) => sum + (b.current * gcvValue[i]), 0) / 100;
//     const currentCost = blends.reduce((sum, b, i) => sum + (b.current * costsPerTon[i]), 0) / 100;
//     const currentBlend = { blend: blends.map(b => b.current), predicted_aft: currentAFT, gcv: currentGCV, cost: currentCost };
//     // ---- YOUR EXISTING OPTIMIZATION LOGIC ENDS HERE ----

//     // decrement trials and save user
//     user.trialsLeft = (user.trialsLeft || 1) - 1;
//     // if hits 0 we lock and destroy session
//     if (user.trialsLeft <= 0) {
//       user.lockedUntil = new Date(Date.now() + 24 * 3600 * 1000);
//       await user.save();
//       req.session.destroy(()=>{});
//       return res.status(200).json({
//          message: 'Calculation ran and this was your final trial. Account locked for 24 hours.',
//          best_aft_blend: bestAftBlend,
//          cheapest_blend: cheapestBlend,
//          balanced_blend: balancedBlend,
//          current_blend: currentBlend,
//          individual_coal_afts: individualCoalAFTs,
//          trialsLeft: 0,
//          lockedUntil: user.lockedUntil
//       });
//     } else {
//       await user.save();
//       return res.json({
//         best_aft_blend: bestAftBlend,
//         cheapest_blend: cheapestBlend,
//         balanced_blend: balancedBlend,
//         current_blend: currentBlend,
//         individual_coal_afts: individualCoalAFTs,
//         trialsLeft: user.trialsLeft
//       });
//     }
//   } catch (err) {
//     console.error("Optimization error:", err);
//     res.status(500).json({ error: "Internal server error" });
//   }
// });
// ---------- OPTIMIZE ROUTE (replace the existing handler with this) ----------
app.post("/optimize", requireAuth, async (req, res) => {
  try {
    const user = req.currentUser;
    // double-check lock
    if (user.lockedUntil && user.lockedUntil > new Date()) {
      return res.status(403).json({ error: 'Account locked until ' + user.lockedUntil.toISOString() });
    }

    // check trials
    if ((user.trialsLeft || 0) <= 0) {
      user.lockedUntil = new Date(Date.now() + 24 * 3600 * 1000);
      await user.save();
      req.session.destroy(()=>{});
      return res.status(403).json({ error: 'Trials exhausted. Account locked for 24 hours.' });
    }

    // record IP for this calculation call
    const ip = getClientIp(req);
    user.lastIP = ip;
    user.ipHistory = user.ipHistory || [];
    user.ipHistory.push({ ip, when: new Date() });

    // ---- OPTIMIZATION LOGIC ----
    const { blends } = req.body;
    if (!blends || !Array.isArray(blends) || blends.length === 0) {
      return res.status(400).json({ error: "Invalid blend data" });
    }

    // oxide columns order used by calculateAFT
    const oxideCols = ['SiO2', 'Al2O3', 'Fe2O3', 'CaO', 'MgO', 'Na2O', 'K2O', 'SO3', 'TiO2'];

    // prepare arrays
    const coalNames = blends.map(b => b.coal || '');
    const oxideValues = blends.map(b => oxideCols.map(col => {
      const v = (b.properties && (b.properties[col] ?? b.properties[col.toUpperCase()] ?? b.properties[col.toLowerCase()])) ?? 0;
      const n = Number(v);
      return Number.isFinite(n) ? n : 0;
    }));

    // sanitize min/max/current
    const minMaxBounds = blends.map(b => {
      const min = Number.isFinite(Number(b.min)) ? Number(b.min) : 0;
      const max = Number.isFinite(Number(b.max)) ? Number(b.max) : 100;
      // ensure sensible order and clamp
      const mm = Math.max(0, Math.min(100, min));
      const mx = Math.max(mm, Math.min(100, max));
      return [Math.round(mm), Math.round(mx)];
    });

    // sanitize costs: if blank/invalid -> 0
    const costsPerTon = blends.map(b => {
      if (b.cost === null || b.cost === undefined || b.cost === '') return 0;
      const n = Number(String(b.cost).replace(/,/g, '').trim());
      return Number.isFinite(n) ? n : 0;
    });

    // sanitize gcv values (if missing -> 0)
    const gcvValue = blends.map(b => {
      const gRaw = b.properties && (b.properties.GCV ?? b.properties.Gcv ?? b.properties.gcv ?? b.properties.gcv);
      const n = Number(gRaw);
      return Number.isFinite(n) ? n : 0;
    });

    // step granularity (1% step)
    const step = 1;

    // generator for combinations with bounds and integer steps that sum to 100
    function* generateCombinations(bounds, idx = 0, acc = []) {
      if (idx === bounds.length - 1) {
        // last one is determined so that sum is 100 - sum(acc)
        const sumSoFar = acc.reduce((s, v) => s + v, 0);
        const lastVal = 100 - sumSoFar;
        const [minLast, maxLast] = bounds[idx];
        if (lastVal >= minLast && lastVal <= maxLast) {
          yield [...acc, lastVal];
        }
        return;
      }
      const [min, max] = bounds[idx];
      for (let v = min; v <= max; v += step) {
        const sumSoFar = acc.reduce((s, vv) => s + vv, 0) + v;
        // quick pruning: if sumSoFar > 100, break this loop (since further v increases sum)
        if (sumSoFar > 100) break;
        // minimal possible remaining (all remaining mins)
        const remainingMin = bounds.slice(idx + 1).reduce((s, b) => s + b[0], 0);
        const remainingMax = bounds.slice(idx + 1).reduce((s, b) => s + b[1], 0);
        // if even with all max we can't reach 100, skip this v
        if (sumSoFar + remainingMax < 100) continue;
        // if even with all min we already exceed 100, skip
        if (sumSoFar + remainingMin > 100) continue;
        yield* generateCombinations(bounds, idx + 1, [...acc, v]);
      }
    }

    // Evaluate valid blends
    const validBlends = [];
    for (const blend of generateCombinations(minMaxBounds)) {
      // weights as fractions for oxide blending
      const weights = blend.map(x => x / 100);
      // blended oxides
      const blendedOxides = oxideCols.map((_, oi) =>
        oxideValues.reduce((sum, val, idx) => sum + val[oi] * weights[idx], 0)
      );

      // predicted AFT using your calculateAFT (assumed defined above)
      const predictedAFT = calculateAFT(blendedOxides);
      console.log('Evaluated blend:', blend, 'Predicted AFT:', predictedAFT);

      // total GCV and cost: blend array is percentage integers, divide by 100
      const totalGcv = blend.reduce((sum, pct, i) => sum + pct * (gcvValue[i] || 0), 0) / 100;
      const totalCost = blend.reduce((sum, pct, i) => sum + pct * (costsPerTon[i] || 0), 0) / 100;

      validBlends.push({
        blend, // array of integer percentages summing to 100
        predicted_aft: predictedAFT,
        cost: totalCost,
        gcv: totalGcv,
        blended_oxides: blendedOxides
      });
    }

    if (validBlends.length === 0) {
      return res.status(404).json({ message: "No valid blends found" });
    }

    // helper arrays
    const aftVals = validBlends.map(b => b.predicted_aft);
    const costVals = validBlends.map(b => b.cost);

    // compute mins/max safely
    const aftMin = Math.min(...aftVals);
    const aftMax = Math.max(...aftVals);
    const costMin = Math.min(...costVals);
    const costMax = Math.max(...costVals);

    // scoring to find a "balanced" blend (maximize aft normalized + cost normalized)
    const blendScores = validBlends.map(b => {
      const aftNorm = (aftMax === aftMin) ? 0.5 : (b.predicted_aft - aftMin) / (aftMax - aftMin);
      // costNorm higher when cost is lower
      const costNorm = (costMax === costMin) ? 0.5 : (costMax - b.cost) / (costMax - costMin);
      return aftNorm + costNorm;
    });

    // pick best variants
    const indexOfBestAft = aftVals.indexOf(Math.max(...aftVals));
    const indexOfCheapest = costVals.indexOf(Math.min(...costVals));
    const indexOfBalanced = blendScores.indexOf(Math.max(...blendScores));

    const bestAftBlend = validBlends[indexOfBestAft];
    const cheapestBlend = validBlends[indexOfCheapest];
    const balancedBlend = validBlends[indexOfBalanced];

    // compute current blend (from user-supplied 'current' values)
    const currentWeights = blends.map(b => {
      const n = Number(b.current);
      return Number.isFinite(n) ? n / 100 : 0;
    });

    const currentBlendedOxides = oxideCols.map((_, oi) =>
      oxideValues.reduce((sum, val, idx) => sum + val[oi] * (currentWeights[idx] || 0), 0)
    );
    const currentAFT = calculateAFT(currentBlendedOxides);
    console.log('Current blend oxides:', currentBlendedOxides, 'Predicted AFT:', currentAFT);
    const currentGCV = blends.reduce((sum, b, i) => sum + (Number(b.current) || 0) * (gcvValue[i] || 0), 0) / 100;
    const currentCost = blends.reduce((sum, b, i) => sum + (Number(b.current) || 0) * (costsPerTon[i] || 0), 0) / 100;
    const currentBlend = { blend: blends.map(b => Number(b.current) || 0), predicted_aft: currentAFT, gcv: currentGCV, cost: currentCost };

    // individual coal AFTs (based on each coal's oxide vector)
    const individualCoalAFTs = oxideValues.map((vals, i) => ({
      coal: coalNames[i] || `coal-${i}`,
      predicted_aft: calculateAFT(vals)
    }));

    console.log('Optimization complete. Best AFT blend:', individualCoalAFTs);

    // Decrement trial for user and save (if you want to count this run)
    user.trialsLeft = Math.max(0, (user.trialsLeft || 0) - 1);
    if (user.trialsLeft <= 0) {
      user.lockedUntil = new Date(Date.now() + 24 * 3600 * 1000);
    }
    await user.save();

    console.log(`best AFT: ${bestAftBlend.blend}, cheapest cost: ${cheapestBlend.blend}, balanced score: ${balancedBlend.blend}, trials left: ${user.trialsLeft}`)

    // return response shaped to client expectations
    return res.json({
      best_aft_blend: bestAftBlend,
      cheapest_blend: cheapestBlend,
      balanced_blend: balancedBlend,
      current_blend: currentBlend,
      individual_coal_afts: individualCoalAFTs
    });

  } catch (err) {
    console.error('/optimize error', err);
    return res.status(500).json({ error: 'Optimization failed', details: String(err) });
  }
});

// ---------- COMPATIBILITY API (for second website / input.html) ----------

// Return array of normalized coal docs for dropdowns
app.get(['/api/coal','/api/coals','/api/coal/list','/api/coal/all'], async (req, res) => {
  try {
    const docs = await Coal.find({}).lean().exec();
    const normalized = docs.map(d => normalizeCoalDoc(d));
    return res.json(normalized);
  } catch (err) {
    console.error('GET /api/coals error:', err);
    return res.status(500).json({ error: err.message || 'Server error' });
  }
});

// Minimal payload for names-only requests
app.get('/api/coalnames', async (req, res) => {
  try {
    const docs = await Coal.find({}, { coal: 1 }).lean().exec();
    const minimal = docs.map(d => ({ _id: d._id, coal: d.coal || d['Coal source name'] || d.name }));
    return res.json(minimal);
  } catch (err) {
    console.error('GET /api/coalnames error:', err);
    return res.status(500).json({ error: err.message || 'Server error' });
  }
});

// Return shape expected by model.html (coal_data: [...])
app.get('/get_coal_types', async (req, res) => {
  try {
    const docs = await Coal.find({}).lean().exec();
    const requiredProps = [
      "SiO2", "Al2O3", "Fe2O3", "CaO", "MgO", "Na2O", "K2O", "TiO2",
      "SO3", "P2O5", "Mn3O4", "Sulphur (S)", "GCV"
    ];
    const coalData = docs.map(row => {
      const id = String(row._id || row.id || '');
      const coalType = row.coal || row.name || row['Coal source name'] || '';
      const transportId = row['Transport ID'] || row.transportId || null;
      const properties = {};
      requiredProps.forEach(prop => {
        properties[prop] = row[prop] !== undefined ? row[prop] : (row[prop.replace('2','₂')] !== undefined ? row[prop.replace('2','₂')] : null);
      });
      if ((properties['GCV'] === null || properties['GCV'] === undefined) && (row.gcv || row.GCV || row.Gcv)) {
        properties['GCV'] = row.gcv || row.GCV || row.Gcv;
      }
      if ((properties['Sulphur (S)'] === null || properties['Sulphur (S)'] === undefined)) {
        properties['Sulphur (S)'] = row['Sulphur (S)'] || row['SulphurS'] || row['Sulphur'] || row.S || null;
      }
      return { id, coalType, transportId, properties };
    });
    return res.json({ coal_data: coalData });
  } catch (error) {
    console.error('/get_coal_types error:', error);
    return res.status(500).json({ error: 'Failed to fetch coal types' });
  }
});
// POST /consume-trial
// Decrements trialsLeft by 1 for the logged-in user, locks account for 24h if it reaches 0,
// logs the user out by destroying the session, and returns current trials and lockedUntil.
app.post('/consume-trial', requireAuth, async (req, res) => {
  try {
    const user = req.currentUser; // set by requireAuth

    // If already locked (should be handled by requireAuth) return locked info
    if (user.lockedUntil && user.lockedUntil > new Date()) {
      return res.status(403).json({ error: 'Account locked', lockedUntil: user.lockedUntil, trialsLeft: user.trialsLeft });
    }

    // Decrement only if > 0
    user.trialsLeft = (user.trialsLeft || 0) - 1;

    // If trials go to 0 or negative, lock for 24 hours and destroy session
    if (user.trialsLeft <= 0) {
      user.trialsLeft = 0;
      user.lockedUntil = new Date(Date.now() + 24 * 3600 * 1000);
      await user.save();

      // destroy session (log out)
      req.session.destroy(err => {
        if (err) console.error('session destroy error during consume-trial', err);
        // respond to client (session is gone)
        return res.json({ message: 'Trials exhausted. Account locked for 24 hours.', trialsLeft: user.trialsLeft, lockedUntil: user.lockedUntil });
      });
      return;
    }

    // Otherwise, save and return trialsLeft
    await user.save();
    return res.json({ message: 'Trial consumed', trialsLeft: user.trialsLeft, lockedUntil: user.lockedUntil || null });
  } catch (err) {
    console.error('/consume-trial error', err);
    return res.status(500).json({ error: 'Internal server error' });
  }
});

// ---------- START SERVER ----------
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});

module.exports = app;
