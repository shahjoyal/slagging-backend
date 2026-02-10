// trainAFTModel_direct.js
// npm i csv-parser ml-random-forest
// node trainAFTModel_direct.js

const fs = require('fs');
const csv = require('csv-parser');
const { RandomForestRegression } = require('ml-random-forest');

const CSV_FILE = 'aft_training_data1.csv'; // <-- put your CSV here
const rows = [];

// -------------------- domain helpers --------------------
function formulaAFT(values) {
  const [SiO2, Al2O3, Fe2O3, CaO, MgO, Na2O, K2O, SO3, TiO2] = values;
  const sumSiAl = SiO2 + Al2O3;
  if (sumSiAl < 55) {
    return 1245 + 1.1*SiO2 + 0.95*Al2O3 - 2.5*Fe2O3 - 2.98*CaO - 4.5*MgO - 7.89*(Na2O+K2O) - 1.7*SO3 - 0.63*TiO2;
  } else if (sumSiAl < 75) {
    return 1323 + 1.45*SiO2 + 0.683*Al2O3 - 2.39*Fe2O3 - 3.1*CaO - 4.5*MgO - 7.49*(Na2O+K2O) - 2.1*SO3 - 0.63*TiO2;
  } else {
    return 1395 + 1.2*SiO2 + 0.9*Al2O3 - 2.5*Fe2O3 - 3.1*CaO - 4.5*MgO - 7.2*(Na2O+K2O) - 1.7*SO3 - 0.63*TiO2;
  }
}

// Build features: use the raw 9 oxide values + the base formula output as an extra feature
function buildFeaturesDirect(values) {
  const base = formulaAFT(values);
  // feature vector: [SiO2, Al2O3, Fe2O3, CaO, MgO, Na2O, K2O, SO3, TiO2, baseAFT]
  return [...values.slice(0,9), base];
}

// -------------------- tiny util functions --------------------
function asFloat(v) {
  if (v === null || v === undefined) return null;
  const n = parseFloat(String(v).toString().replace(/,/g, ''));
  return Number.isFinite(n) ? n : null;
}

function quantile(sortedArr, q) {
  if (!sortedArr.length) return 0;
  if (q <= 0) return sortedArr[0];
  if (q >= 1) return sortedArr[sortedArr.length-1];
  const pos = (sortedArr.length - 1) * q;
  const base = Math.floor(pos);
  const rest = pos - base;
  if (sortedArr[base+1] !== undefined) {
    return sortedArr[base] + rest * (sortedArr[base+1] - sortedArr[base]);
  } else {
    return sortedArr[base];
  }
}

function meanAbs(a, b) {
  let s = 0;
  for (let i=0;i<a.length;i++) s += Math.abs(a[i] - b[i]);
  return s / a.length;
}
function rmseCalc(a, b) {
  let s = 0;
  for (let i=0;i<a.length;i++) s += Math.pow(a[i]-b[i],2);
  return Math.sqrt(s/a.length);
}
function percentWithin(a, b, tol) {
  let c = 0;
  for (let i=0;i<a.length;i++) if (Math.abs(a[i]-b[i]) <= tol) c++;
  return (c/a.length)*100;
}

// Fisher-Yates shuffle
function shuffleArray(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

// -------------------- reading CSV --------------------
if (!fs.existsSync(CSV_FILE)) {
  console.error("CSV file not found:", CSV_FILE);
  process.exit(1);
}
fs.createReadStream(CSV_FILE)
  .pipe(csv())
  .on('data', (r) => rows.push(r))
  .on('end', main)
  .on('error', (err) => { console.error('CSV read error', err); process.exit(1); });

// -------------------- main --------------------
function main() {
  // parse rows, drop incomplete
  const data = [];
  rows.forEach((row, i) => {
    const values = [
      asFloat(row.SiO2), asFloat(row.Al2O3), asFloat(row.Fe2O3),
      asFloat(row.CaO), asFloat(row.MgO), asFloat(row.Na2O),
      asFloat(row.K2O), asFloat(row.SO3), asFloat(row.TiO2)
    ];
    const aft = asFloat(row.AFT);
    if (values.some(v => v === null) || aft === null) return; // skip incomplete rows
    data.push({ index: i, values, aft });
  });

  if (!data.length) {
    console.error("No valid rows after parsing CSV.");
    process.exit(1);
  }

  console.log(`Rows after removing missing: ${data.length}`);

  // log AFT IQR outliers (do not drop automatically)
  const afts = data.map(d => d.aft).slice().sort((a,b)=>a-b);
  const q1 = quantile(afts, 0.25), q3 = quantile(afts, 0.75), iqr = q3 - q1;
  const lower = q1 - 1.5*iqr, upper = q3 + 1.5*iqr;
  const aftOutliers = data.filter(d => d.aft < lower || d.aft > upper).map(d=>d.index);
  console.log('AFT IQR bounds:', lower, upper, 'AFT outlier row indices (CSV row index):', aftOutliers);

  // Build full feature matrix (direct AFT target)
  const Xraw = data.map(d => buildFeaturesDirect(d.values)); // 10 features: 9 oxides + baseAFT
  const Y = data.map(d => d.aft);

  // Winsorize features at 1% / 99% and save cutoffs (so runtime API can apply same clipping)
  const Xt = Xraw.map(r => r.slice()); // deep-ish copy
  const cutoffs = { lower: [], upper: [] };
  for (let c = 0; c < Xt[0].length; c++) {
    const col = Xt.map(r => r[c]).slice().sort((a,b)=>a-b);
    const lo = quantile(col, 0.01), hi = quantile(col, 0.99);
    cutoffs.lower.push(lo);
    cutoffs.upper.push(hi);
    for (let r = 0; r < Xt.length; r++) {
      if (Xt[r][c] < lo) Xt[r][c] = lo;
      else if (Xt[r][c] > hi) Xt[r][c] = hi;
    }
  }
  fs.writeFileSync('preprocessing.json', JSON.stringify(cutoffs, null, 2));
  console.log('Saved preprocessing cutoffs to preprocessing.json');

  // shuffle & split (80/20)
  const ids = shuffleArray(Array.from({length: Xt.length}, (_,i) => i));
  const split = Math.floor(0.8 * ids.length);
  const trainIdx = ids.slice(0, split), testIdx = ids.slice(split);

  const X_train = trainIdx.map(i => Xt[i]);
  const Y_train = trainIdx.map(i => Y[i]);
  const X_test = testIdx.map(i => Xt[i]);
  const Y_test = testIdx.map(i => Y[i]);
  const original_test_rows = testIdx.map(i => data[i].index);

  console.log('Train size:', X_train.length, 'Test size:', X_test.length);

  // conservative RF grid for direct AFT prediction
  const grid = [
    { nEstimators: 100, maxDepth: 10, minNumSamples: 4, replacement: false },
    { nEstimators: 150, maxDepth: 12, minNumSamples: 3, replacement: false },
    { nEstimators: 200, maxDepth: 14, minNumSamples: 2, replacement: false }
  ];

  let best = { mae: Infinity, params: null, modelJSON: null, stats: null };

  for (const cfg of grid) {
    console.log('Training RF with', cfg);
    const rf = new RandomForestRegression({
      nEstimators: cfg.nEstimators,
      maxFeatures: Math.max(1, Math.floor(Math.sqrt(X_train[0].length))),
      maxDepth: cfg.maxDepth,
      minNumSamples: cfg.minNumSamples,
      replacement: !!cfg.replacement,
      seed: 42
    });

    rf.train(X_train, Y_train);

    const preds = rf.predict(X_test);
    const mae = meanAbs(Y_test, preds);
    const rmse = rmseCalc(Y_test, preds);
    const within20 = percentWithin(Y_test, preds, 20);

    console.log(` -> MAE: ${mae.toFixed(2)}, RMSE: ${rmse.toFixed(2)}, % within ±20: ${within20.toFixed(1)}%`);

    if (mae < best.mae) {
      best.mae = mae;
      best.params = cfg;
      best.modelJSON = rf.toJSON();
      best.stats = { mae, rmse, within20, preds };
    }
  }

  // Save best model and metadata
  if (best.modelJSON) {
    fs.writeFileSync('aft_model.json', JSON.stringify(best.modelJSON));
    console.log('Saved best model to aft_model.json with params', best.params, 'stats', best.stats);
    const meta = {
      params: best.params,
      date: new Date().toISOString(),
      train_size: X_train.length,
      test_size: X_test.length
    };
    fs.writeFileSync('model_meta.json', JSON.stringify(meta, null, 2));
    console.log('Saved model_meta.json');
  } else {
    console.error('No model was trained.');
    process.exit(1);
  }

  // Evaluate on test set and list rows with abs error > 20 (compare PRED to actual)
  const bestPreds = best.stats.preds;
  const badRows = [];
  for (let i = 0; i < bestPreds.length; i++) {
    const idx = testIdx[i];
    const csvRow = data[idx].index;
    const actualAFT = data[idx].aft;
    const predictedAFT = Math.round(bestPreds[i]);
    const absError = Math.round(Math.abs(actualAFT - predictedAFT));

    if (absError > 20) {
      badRows.push({ csvRowIndex: csvRow, actualAFT, predictedAFT, absError });
    }
  }

  console.log('Number of test rows with abs error > 20:', badRows.length);
  console.log('Example rows with abs error > 20 (CSV row index, actual, predicted, absError):');
  console.log(badRows.slice(0, 200));
  console.log('Done.');
}
