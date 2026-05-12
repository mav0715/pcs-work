// ============================================================
// PCS MIS — MACHINE INTELLIGENCE SYSTEM
// AppsScript_MIS.js | Champo Carpets, Panipat | Final Version
// ============================================================
// SETUP:
// 1. Paste into Apps Script editor bound to MIS sheet
// 2. Update MIS_CONFIG below with real values
// 3. Run installTriggers() ONCE
// 4. Deploy as Web App → Execute as Me → Anyone can access
// 5. Copy deployment URL → mis-output-form.html SCRIPT_URL
// 6. Run recalcTargets() ONCE to seed OUTPUT_TARGETS tab
//
// MANUAL EDIT SAFETY:
// Script reads live sheet values on every call — never caches.
// Editing raw input cols (1-46) in DAILY_OUTPUT is safe.
// NEVER edit calculated cols (47-94) — script overwrites them.
// NEVER delete header rows or rename tabs.
//
// TO ADD WEAVING LATER (3 changes only):
// 1. Set WEAVING count in opCounts inside recalcTargets()
// 2. Add "WEAVING" to depts array in complianceSummary()
// 3. Add "WEAVING" to depts array in sendWeeklyEmail()
// ============================================================

const MIS_CONFIG = {
  MIS_SHEET_ID:        "16hMfvllFpw2WmIPvD_KBPbopUjQmLGq7y1y0AS_91qM",
  PCS_SHEET_ID:        "1S2ikmWM_GlWbYVuOlPQBbGEebtnO31qNcc8m7-CTaRU",
  PASSWORD:            "mis2025champo",
  ADMIN_EMAIL:         "saiyam@champocarpets.com",
  WASENDER_URL:        "REPLACE_WITH_WASENDER_API_URL",
  WASENDER_TOKEN:      "REPLACE_WITH_WASENDER_TOKEN",
  WORKING_DAYS:        26,
  REMINDER_DAYS:       [1, 6],   // Mon=1, Sat=6
  REMINDER_HOURS:      [10, 16], // 10:30am, 4:30pm
  REJECTION_FLAG_MULT: 1.5,
  REJECTION_FLAG_MIN:  2,
  REJECTION_AVG_DAYS:  30,
};

// ─── TAB NAMES ───────────────────────────────────────────────
const TABS = {
  ARTICLES:  "ARTICLES",
  BUYERS:    "BUYERS",
  RATES:     "RATES",
  MACHINES:  "MACHINES",
  DAILY:     "DAILY_OUTPUT",
  SAMPLING:  "SAMPLING_LOG",
  COST_SUM:  "COST_SUMMARY",
  ART_COST:  "ARTICLE_COST_SUMMARY",
  TARGETS:   "OUTPUT_TARGETS",
  HEADS:     "DEPT_HEADS",
  DIFF:      "DIFFICULTY_CONFIG",
  FX:        "CURRENCY_LOG",
};

// ─── COLUMN MAPS (1-based) ───────────────────────────────────
// ARTICLES — 26 columns
const CA = {
  CODE:1,NAME:2,CAT:3,SIZE:4,
  LEN:5,WID:6,UNIT:7,
  P1:8,P2:9,P3:10,P4:11,
  TUFT:12,ACTIVE:13,STATUS:14,
  REED:15,PICKS:16,COLORS:17,
  SHRINK_V:18,SHRINK_U:19,
  MARGIN_WV:20,MARGIN_WU:21,
  MARGIN_LV:22,MARGIN_LU:23,
  DEF_UNIT:24,
  SPEC_FROM:25,SPEC_TO:26,
};

// DAILY_OUTPUT — 94 columns
const CD = {
  // RAW INPUTS (1-46)
  ROW_ID:1,DATE:2,PERIOD:3,DEPT:4,MACHINE_ID:5,SHIFT:6,FACILITY:7,
  BUYER:8,ART_CODE:9,ART_NAME:10,SIZE_LABEL:11,
  LEN_FIN:12,WID_FIN:13,UNIT:14,CATEGORY:15,
  TUFT_TYPE:16,COLORS:17,COMPLEXITY:18,SPM:19,
  PANEL_TYPE:20,OUTPUT_UNIT:21,RAW_OUTPUT:22,
  PCS_ACROSS:23,GAP_V:24,GAP_U:25,
  SHRINK_V:26,SHRINK_U:27,
  MARGIN_WV:28,MARGIN_WU:29,
  MARGIN_LV:30,MARGIN_LU:31,
  REED:32,PICKS:33,
  SPEC_VER:34,
  PCS_PROD:35,REJ_QTY:36,REJ_UNIT:37,STITCH_PC:38,
  PAPERS_PR:39,PAPERS_REJ:40,PAPER_SRC:41,
  OPS_DAY:42,OPS_NIGHT:43,RUN_TYPE:44,
  BY:45,AT:46,
  // CALCULATED (47-94) — NEVER MANUALLY EDIT
  MACH_W_CM:47,FAB_L_CM:48,GROSS_SQM:49,GROSS_SQFT:50,
  SHRINK_CM:51,SHRINK_SQM:52,SHRINK_SQFT:53,USABLE_W:54,
  MARGIN_W_CM:55,MARGIN_L_CM:56,CUT_W_CM:57,CUT_L_CM:58,
  PCS_IN_L:59,TOTAL_PCS:60,NET_CONS_SQM:61,GAP_SQM:62,
  WASTE_SQM:63,WASTE_SQFT:64,WASTE_PCT:65,WASTE_COST:66,GAP_SQFT:67,
  SQM_PC:68,SQFT_PC:69,TOT_SQM:70,TOT_SQFT:71,
  REJ_SQM:72,REJ_SQFT:73,NET_SQM:74,NET_SQFT:75,REJ_RATE:76,
  TOT_ST:77,ST_WASTE:78,NET_ST:79,
  TOT_OPS:80,FORMULA:81,OP_RATE:82,FIXED_RATE:83,
  MANPWR:84,FIXED_COST:85,DAY_COST:86,
  SHARE_PCT:87,ATTR_COST:88,
  CPT_SQM:89,CPT_SQFT:90,CPT_1000:91,COST_LOSS:92,
  DIFF_SCORE:93,NORM_OUT:94,
};

// SAMPLING_LOG
const CS = {ID:1,DATE:2,PERIOD:3,DEPT:4,N_DES:5,CREATED:6,SENT:7,CONV:8,COST:9,CPD:10,BY:11,AT:12};

// RATES
const CR = {DEPT:1,FORMULA:2,OP_RATE:3,FIXED:4,GREEN:5,RED:6,UNIT:7,FROM:8,TO:9};

// MACHINES
const CM = {ID:1,NAME:2,DEPT:3,TYPE:4,FAC:5,SPEC1:6,SPEC2:7,PIN:8,STATUS:9,MODE:10,NOTES:11};

// DEPT_HEADS
const CH = {PIN:1,NAME:2,DEPTS:3,FAC:4,ACTIVE:5,PHONE:6};

// OUTPUT_TARGETS
const CT = {DEPT:1,DAY_COST:2,WD:3,G_SQFT:4,Y_SQFT:5,R_SQFT:6,G_SQM:7,Y_SQM:8,R_SQM:9,G_1000:10,R_1000:11,G_MON:12,R_MON:13,UPDATED:14};

// DIFFICULTY_CONFIG
const CDIFF = {DEPT:1,FACTOR:2,BAND:3,MIN:4,MAX:5,MULT:6};

// ─── CORE HELPERS ────────────────────────────────────────────
function ss()        { return SpreadsheetApp.openById(MIS_CONFIG.MIS_SHEET_ID); }
function tab(name)   { const s=ss().getSheetByName(name); if(!s) throw new Error("Tab missing: "+name); return s; }
function rows(name)  { const s=tab(name),l=s.getLastRow(); return l<2?[]:s.getRange(2,1,l-1,s.getLastColumn()).getValues(); }
function r2(n)       { return Math.round((parseFloat(n)||0)*100)/100; }
function r4(n)       { return Math.round((parseFloat(n)||0)*10000)/10000; }
function fl(n)       { return Math.floor(parseFloat(n)||0); }
function fmtDate(d)  { return Utilities.formatDate(d instanceof Date?d:new Date(d),Session.getScriptTimeZone(),"yyyy-MM-dd"); }
function respond(o)  { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function getPeriod(d) {
  const day=(d instanceof Date?d:new Date(d)).getDate();
  return day<=11?"P1":day<=22?"P2":"P3";
}

function nextId(tabName,prefix,col,pad) {
  const data=rows(tabName);
  if(!data.length) return `${prefix}-${"1".padStart(pad,"0")}`;
  const nums=data.map(r=>parseInt(String(r[col-1]).replace(prefix+"-",""))||0).filter(n=>n>0);
  return `${prefix}-${String((nums.length?Math.max(...nums):0)+1).padStart(pad,"0")}`;
}

function toNum(v)    { return parseFloat(v)||0; }
function toCm(v,u)   { return u==="in"?(toNum(v)*2.54):toNum(v); }
function toPct(v,u,base) { return u==="%"?(base*(toNum(v)/100)):toCm(v,u); }

// ─── doGet + doPost ───────────────────────────────────────────
function doGet()  { return HtmlService.createHtmlOutput("<h2>MIS API running ✓</h2>"); }

function doPost(e) {
  try {
    const {action,password,data} = JSON.parse(e.postData.contents);
    if(password!==MIS_CONFIG.PASSWORD) return respond({status:"error",message:"Invalid password"});
    const map = {
      verifyPin,getMachines,getArticles,getBuyers,
      submitOutput,submitSampling,addArticle,updateArticleSpec,
      getDashboardData,getTargets,
      getCurrencyRates,getCurrencyHistory,
    };
    if(!map[action]) return respond({status:"error",message:"Unknown action: "+action});
    return respond(map[action](data));
  } catch(err) {
    return respond({status:"error",message:err.toString()});
  }
}

// ─── VERIFY PIN ──────────────────────────────────────────────
function verifyPin(d) {
  for(const r of rows(TABS.HEADS)) {
    if(String(r[CH.PIN-1])===String(d.pin)&&r[CH.ACTIVE-1]==="Y") {
      return {
        status:"ok",
        name:r[CH.NAME-1],
        depts:String(r[CH.DEPTS-1]).split(",").map(x=>x.trim()),
        facility:r[CH.FAC-1]||"PNP",
        phone:r[CH.PHONE-1]||"",
      };
    }
  }
  return {status:"error",message:"PIN not found or inactive"};
}

// ─── GET MACHINES ─────────────────────────────────────────────
function getMachines(d) {
  const {depts,facility}=d;
  return {status:"ok",machines:rows(TABS.MACHINES)
    .filter(r=>r[CM.STATUS-1]==="Active"
      &&(!facility||r[CM.FAC-1]===facility)
      &&(!depts||depts.includes(r[CM.DEPT-1])))
    .map(r=>({id:r[CM.ID-1],name:r[CM.NAME-1],dept:r[CM.DEPT-1],
      type:r[CM.TYPE-1],spec1:r[CM.SPEC1-1],spec2:r[CM.SPEC2-1],mode:r[CM.MODE-1]}))};
}

// ─── GET ARTICLES ─────────────────────────────────────────────
// Returns only current active spec (Spec_Effective_To = blank)
function getArticles(d) {
  const dept=d.dept||"";
  return {status:"ok",articles:rows(TABS.ARTICLES)
    .filter(r=>r[CA.ACTIVE-1]!=="N" && !r[CA.SPEC_TO-1]) // active + current spec
    .filter(r=>{
      if(dept==="TUFTING_CUT")  return r[CA.TUFT-1]==="Cut";
      if(dept==="TUFTING_LOOP") return r[CA.TUFT-1]==="HKD";
      return true;
    })
    .map(r=>({
      code:r[CA.CODE-1],name:r[CA.NAME-1],category:r[CA.CAT-1],
      sizeLabel:r[CA.SIZE-1],length:r[CA.LEN-1],width:r[CA.WID-1],unit:r[CA.UNIT-1],
      processes:[r[CA.P1-1],r[CA.P2-1],r[CA.P3-1],r[CA.P4-1]].filter(Boolean),
      tuftType:r[CA.TUFT-1],status:r[CA.STATUS-1],
      reed:r[CA.REED-1],picks:r[CA.PICKS-1],colors:r[CA.COLORS-1],
      shrinkV:r[CA.SHRINK_V-1],shrinkU:r[CA.SHRINK_U-1],
      marginWV:r[CA.MARGIN_WV-1],marginWU:r[CA.MARGIN_WU-1],
      marginLV:r[CA.MARGIN_LV-1],marginLU:r[CA.MARGIN_LU-1],
      defUnit:r[CA.DEF_UNIT-1],
      specFrom:r[CA.SPEC_FROM-1],
    }))};
}

// ─── GET BUYERS ───────────────────────────────────────────────
function getBuyers() {
  return {status:"ok",buyers:rows(TABS.BUYERS).filter(r=>r[2]!=="N").map(r=>({code:r[0],name:r[1]}))};
}

// ─── ADD ARTICLE (incomplete) ─────────────────────────────────
function addArticle(d) {
  const row=new Array(26).fill("");
  row[CA.CODE-1]=d.code||d.name; row[CA.NAME-1]=d.name;
  row[CA.CAT-1]=d.category||"";  row[CA.ACTIVE-1]="Y";
  row[CA.STATUS-1]="Incomplete";  row[CA.DEF_UNIT-1]=d.defUnit||"meters";
  row[CA.SPEC_FROM-1]=new Date();
  tab(TABS.ARTICLES).appendRow(row);
  return {status:"ok",message:"Saved as Incomplete — complete in ARTICLES tab"};
}

// ─── UPDATE ARTICLE SPEC (spec versioning) ───────────────────
// Called when owner adjusts reed/picks/margins and chooses to save as new spec
function updateArticleSpec(d) {
  const sh   = tab(TABS.ARTICLES);
  const data = sh.getDataRange().getValues();
  const today = new Date();

  // Find current active spec row for this article
  for(let i=1;i<data.length;i++) {
    const r=data[i];
    if(r[CA.CODE-1]!==d.articleCode) continue;
    if(r[CA.SPEC_TO-1]) continue; // already closed
    // Close this spec
    sh.getRange(i+1,CA.SPEC_TO,1,1).setValue(today);
    break;
  }

  // Append new spec row — copy all fields, update changed ones
  const lastActive = data.slice(1).findLast(r=>r[CA.CODE-1]===d.articleCode);
  if(!lastActive) return {status:"error",message:"Article not found: "+d.articleCode};

  const newRow = [...lastActive];
  newRow[CA.REED-1]      = d.reed      || lastActive[CA.REED-1];
  newRow[CA.PICKS-1]     = d.picks     || lastActive[CA.PICKS-1];
  newRow[CA.COLORS-1]    = d.colors    || lastActive[CA.COLORS-1];
  newRow[CA.SHRINK_V-1]  = d.shrinkV   || lastActive[CA.SHRINK_V-1];
  newRow[CA.SHRINK_U-1]  = d.shrinkU   || lastActive[CA.SHRINK_U-1];
  newRow[CA.MARGIN_WV-1] = d.marginWV  || lastActive[CA.MARGIN_WV-1];
  newRow[CA.MARGIN_WU-1] = d.marginWU  || lastActive[CA.MARGIN_WU-1];
  newRow[CA.MARGIN_LV-1] = d.marginLV  || lastActive[CA.MARGIN_LV-1];
  newRow[CA.MARGIN_LU-1] = d.marginLU  || lastActive[CA.MARGIN_LU-1];
  newRow[CA.SPEC_FROM-1] = today;
  newRow[CA.SPEC_TO-1]   = "";
  sh.appendRow(newRow);
  return {status:"ok",message:"New spec version saved"};
}

// ─── RATES LOOKUP ─────────────────────────────────────────────
function getRate(dept,date) {
  const d=date instanceof Date?date:new Date(date);
  for(const r of rows(TABS.RATES)) {
    if(r[CR.DEPT-1]!==dept) continue;
    const from=r[CR.FROM-1],to=r[CR.TO-1];
    if((!(from instanceof Date)||d>=from)&&(!(to instanceof Date)||d<=to)) {
      return {formula:r[CR.FORMULA-1],opRate:toNum(r[CR.OP_RATE-1]),
              fixedRate:toNum(r[CR.FIXED-1]),green:toNum(r[CR.GREEN-1]),
              red:toNum(r[CR.RED-1]),unit:r[CR.UNIT-1]||"/sqft"};
    }
  }
  return null;
}

// ─── DIFFICULTY SCORE ────────────────────────────────────────
function getDiffScore(dept,colors,reed,picks,spm) {
  const diffRows=rows(TABS.DIFF);
  function getM(factorName,val) {
    const n=toNum(val);
    const bands=diffRows.filter(r=>(r[CDIFF.DEPT-1]===dept||r[CDIFF.DEPT-1]==="ALL")&&r[CDIFF.FACTOR-1]===factorName);
    for(const b of bands) {
      const mn=toNum(b[CDIFF.MIN-1]),mx=toNum(b[CDIFF.MAX-1]);
      if(mn===0&&mx===0) return toNum(b[CDIFF.MULT-1]); // technique multiplier
      if(n>=mn&&n<=mx)   return toNum(b[CDIFF.MULT-1]);
    }
    return 1.0;
  }
  let score=1.0;
  score *= getM("Color_Band",colors);
  if(dept==="WEAVING") { score *= getM("Reed_Band",reed); score *= getM("Picks_Band",picks); }
  if(dept==="EMBROIDERY") score *= getM("SPM_Band",spm);
  if(dept==="TUFTING_CUT"||dept==="TUFTING_LOOP"||dept==="TUFTING_ROBO") score *= getM("Technique",0);
  return r4(score);
}

// ─── AREA + YIELD CALCULATIONS ────────────────────────────────
function calcWeaving(d, machWidthIn) {
  const machWcm   = machWidthIn * 2.54;
  const lenFincm  = toCm(d.lenFin, d.unit);
  const widFincm  = toCm(d.widFin, d.unit);

  // Shrinkage
  const shrinkCm  = d.shrinkU==="%"
    ? machWcm*(toNum(d.shrinkV)/100)
    : toCm(d.shrinkV,"in");
  const usableWcm = machWcm - shrinkCm;

  // Stitch margins
  const marginWcm = d.marginWU==="%"
    ? widFincm*(toNum(d.marginWV)/100)
    : toCm(d.marginWV,"in");
  const marginLcm = d.marginLU==="%"
    ? lenFincm*(toNum(d.marginLV)/100)
    : toCm(d.marginLV,"in");

  // Fabric cut per piece
  const cutWcm = widFincm + marginWcm;
  const cutLcm = lenFincm + marginLcm;

  // Fabric length from raw output
  let fabLcm = 0;
  const rawOut = toNum(d.rawOutput);
  if(d.outputUnit==="meters") fabLcm = rawOut * 100;
  else if(d.outputUnit==="sqm") fabLcm = rawOut / (machWcm/100) * 100; // back-calc
  else fabLcm = toNum(d.pcsAcross) * cutWcm; // pcs mode — estimate

  // Gap between pieces
  const gapCm = toCm(d.gapV, d.gapU||"cm");

  // Yield
  const pcsInWidth  = d.pcsAcross ? fl(toNum(d.pcsAcross)) : fl(usableWcm / cutWcm);
  const pcsInLength = fabLcm > 0 ? fl(fabLcm / (cutLcm + gapCm)) : 0;
  const totalPcs    = pcsInWidth * pcsInLength;

  // Areas
  const grossSqm    = r4((machWcm * fabLcm) / 10000);
  const grossSqft   = r4(grossSqm * 10.7639);
  const shrinkSqm   = r4((shrinkCm * fabLcm) / 10000);
  const shrinkSqft  = r4(shrinkSqm * 10.7639);
  const netConsSqm  = r4(totalPcs * (cutWcm * cutLcm) / 10000);
  const gapSqm      = r4((gapCm * pcsInLength * machWcm) / 10000);
  const gapSqft     = r4(gapSqm * 10.7639);
  const sqmPerPc    = r4((lenFincm * widFincm) / 10000); // finished size
  const sqftPerPc   = r4(sqmPerPc * 10.7639);
  const totSqm      = r4(totalPcs * sqmPerPc);
  const totSqft     = r4(totSqm * 10.7639);
  const wasteSqm    = r4(grossSqm - netConsSqm - shrinkSqm);
  const wasteSqft   = r4(wasteSqm * 10.7639);
  const wastePct    = grossSqm > 0 ? r2((wasteSqm/grossSqm)*100) : 0;

  return {
    machWcm,fabLcm,grossSqm,grossSqft,
    shrinkCm,shrinkSqm,shrinkSqft,usableWcm,
    marginWcm,marginLcm,cutWcm,cutLcm,
    pcsInLength,totalPcs,netConsSqm,gapSqm,gapSqft,
    wasteSqm,wasteSqft,wastePct,
    sqmPerPc,sqftPerPc,totSqm,totSqft,
  };
}

function calcStdArea(lenFin,widFin,unit,pcs) {
  const l = toCm(lenFin,unit), w = toCm(widFin,unit);
  const sqmPc = r4((l*w)/10000);
  const sqftPc= r4(sqmPc*10.7639);
  return {sqmPc,sqftPc,totSqm:r2(pcs*sqmPc),totSqft:r2(pcs*sqftPc)};
}

// ─── NORMALISE REJECTION ──────────────────────────────────────
function normaliseRej(rejQty,rejUnit,sqmPc,sqftPc,totSqm,totSqft) {
  const qty=toNum(rejQty);
  let rejSqm=0,rejSqft=0;
  if(rejUnit==="sqm")  { rejSqm=r2(qty); rejSqft=r2(qty*10.7639); }
  else if(rejUnit==="sqft") { rejSqft=r2(qty); rejSqm=r2(qty/10.7639); }
  else if(rejUnit==="meters") {
    // weaving: rejection in meters → multiply by machine width (use totSqm/rawOutput ratio)
    // Approximation: rejSqm = qty * (totSqm / rawOutputMeters) — caller handles
    rejSqm=r2(qty*sqmPc); rejSqft=r2(qty*sqftPc);
  } else {
    // pcs
    rejSqm=r2(qty*sqmPc); rejSqft=r2(qty*sqftPc);
  }
  const netSqm  = r2(Math.max(0,totSqm-rejSqm));
  const netSqft = r2(Math.max(0,totSqft-rejSqft));
  const rejRate = totSqft>0 ? r2((rejSqft/totSqft)*100) : 0;
  return {rejSqm,rejSqft,netSqm,netSqft,rejRate};
}

// ─── SUBMIT OUTPUT ────────────────────────────────────────────
function submitOutput(d) {
  const date   = new Date(d.date);
  const rate   = getRate(d.dept,date);
  if(!rate) throw new Error("No rate for "+d.dept+" on "+d.date);

  const rowId  = nextId(TABS.DAILY,"DO",CD.ROW_ID,5);

  // Machine width
  const machRows = rows(TABS.MACHINES);
  const mach     = machRows.find(r=>r[CM.ID-1]===d.machineId);
  const machWidIn= mach ? toNum(mach[CM.SPEC1-1]) : 0;

  // Raw inputs
  const rawOut  = toNum(d.rawOutput);
  const pcsProd = toNum(d.pcsProd)||rawOut;
  const rejQty  = toNum(d.rejectedQty);
  const rejUnit = d.rejectedUnit||"pcs";
  const stPc    = toNum(d.stitchesPerPc);
  const ppPr    = toNum(d.papersPrinted);
  const ppRej   = toNum(d.papersRejected);
  const opsD    = toNum(d.operatorsDay);
  const opsN    = toNum(d.operatorsNight);
  const totOps  = opsD+opsN;

  // Difficulty score
  const diffScore = getDiffScore(
    d.dept, d.noOfColors||d.colors,
    d.reedCount, d.picksPerInch, d.spm
  );

  // Area calculations — dept-specific
  let grossSqm=0,grossSqft=0,shrinkCm=0,shrinkSqm=0,shrinkSqft=0,usableW=0;
  let marginWcm=0,marginLcm=0,cutWcm=0,cutLcm=0;
  let pcsInL=0,totalPcs=0,netConsSqm=0,gapSqm=0,gapSqft=0;
  let wasteSqm=0,wasteSqft=0,wastePct=0,wasteCost=0;
  let sqmPc=0,sqftPc=0,totSqm=0,totSqft=0;
  let fabLcm=0,machWcm=0;

  if(d.dept==="WEAVING") {
    const w = calcWeaving({
      lenFin:d.lenFin||d.length, widFin:d.widFin||d.width, unit:d.unit,
      outputUnit:d.outputUnit, rawOutput:rawOut,
      pcsAcross:d.pcsAcross, gapV:d.gapV, gapU:d.gapU,
      shrinkV:d.shrinkV, shrinkU:d.shrinkU,
      marginWV:d.marginWV, marginWU:d.marginWU,
      marginLV:d.marginLV, marginLU:d.marginLU,
    }, machWidIn);
    ({grossSqm,grossSqft,shrinkCm,shrinkSqm,shrinkSqft,usableWcm:usableW,
      marginWcm,marginLcm,cutWcm,cutLcm,pcsInLength:pcsInL,totalPcs,
      netConsSqm,gapSqm,gapSqft,wasteSqm,wasteSqft,wastePct,
      sqmPerPc:sqmPc,sqftPerPc:sqftPc,totSqm,totSqft,
      machWcm,fabLcm} = w);
  } else {
    // Knitting, Tufting, Embroidery, Printing
    const sa = calcStdArea(d.length,d.width,d.unit,pcsProd);
    sqmPc=sa.sqmPc; sqftPc=sa.sqftPc;
    totSqm=sa.totSqm; totSqft=sa.totSqft;
    machWcm  = machWidIn*2.54;
    grossSqm = totSqm; grossSqft = totSqft;
  }

  // Rejection normalisation
  // For weaving: rejection in meters — sqmPc = sqm per meter of fabric
  const sqmPerMeterWeaving = machWcm>0&&fabLcm>0 ? r4((machWcm*fabLcm)/10000/(rawOut||1)) : sqmPc;
  const rejSqmPc  = d.dept==="WEAVING" && rejUnit==="meters" ? sqmPerMeterWeaving : sqmPc;
  const rejSqftPc = rejSqmPc*10.7639;
  const {rejSqm,rejSqft,netSqm,netSqft,rejRate} = normaliseRej(rejQty,rejUnit,rejSqmPc,rejSqftPc,totSqm,totSqft);

  // Stitches (embroidery)
  const totSt     = d.dept==="EMBROIDERY" ? r2(pcsProd*stPc) : 0;
  const stWaste   = d.dept==="EMBROIDERY" ? r2(rejQty*stPc)  : 0;
  const netSt     = r2(totSt-stWaste);

  // Cost (Type_A or Type_B)
  let manpwr=0,fixedCost=rate.fixedRate,dayCost=0;
  if(rate.formula==="Type_B") {
    manpwr   = r2(totOps*rate.opRate);
    dayCost  = r2(manpwr+fixedCost);
  } else {
    dayCost  = fixedCost;
  }

  // Placeholder share — recalcShare() corrects proportionally after append
  const attrCost = dayCost;
  const cptSqft  = totSqft>0 ? r2(attrCost/totSqft) : 0;
  const cptSqm   = totSqm>0  ? r2(attrCost/totSqm)  : 0;
  const cpt1000  = totSt>0   ? r2(attrCost/totSt*1000) : 0;
  const costLoss = totSqft>0 ? r2((rejSqft/totSqft)*attrCost) : 0;
  wasteCost      = grossSqm>0 ? r2((wasteSqm/grossSqm)*attrCost) : 0;
  const normOut  = diffScore>0 ? r4(totSqm/diffScore) : totSqm;

  // Build row (94 columns)
  const row = new Array(94).fill("");
  // Raw inputs
  row[CD.ROW_ID-1]=rowId;         row[CD.DATE-1]=date;
  row[CD.PERIOD-1]=getPeriod(date);row[CD.DEPT-1]=d.dept;
  row[CD.MACHINE_ID-1]=d.machineId;row[CD.SHIFT-1]=d.shift;
  row[CD.FACILITY-1]=d.facility||"PNP";
  row[CD.BUYER-1]=d.buyer||"";    row[CD.ART_CODE-1]=d.articleCode||"";
  row[CD.ART_NAME-1]=d.articleName||"";row[CD.SIZE_LABEL-1]=d.sizeLabel||"";
  row[CD.LEN_FIN-1]=toNum(d.length)||"";row[CD.WID_FIN-1]=toNum(d.width)||"";
  row[CD.UNIT-1]=d.unit||"cm";    row[CD.CATEGORY-1]=d.category||"";
  row[CD.TUFT_TYPE-1]=d.tuftingType||"";
  row[CD.COLORS-1]=toNum(d.noOfColors)||"";
  row[CD.COMPLEXITY-1]=d.complexity||"";row[CD.SPM-1]=toNum(d.spm)||"";
  row[CD.PANEL_TYPE-1]=d.panelType||"NA";
  row[CD.OUTPUT_UNIT-1]=d.outputUnit||"pcs";row[CD.RAW_OUTPUT-1]=rawOut;
  row[CD.PCS_ACROSS-1]=toNum(d.pcsAcross)||"";
  row[CD.GAP_V-1]=toNum(d.gapV)||"";row[CD.GAP_U-1]=d.gapU||"";
  row[CD.SHRINK_V-1]=toNum(d.shrinkV)||"";row[CD.SHRINK_U-1]=d.shrinkU||"";
  row[CD.MARGIN_WV-1]=toNum(d.marginWV)||"";row[CD.MARGIN_WU-1]=d.marginWU||"";
  row[CD.MARGIN_LV-1]=toNum(d.marginLV)||"";row[CD.MARGIN_LU-1]=d.marginLU||"";
  row[CD.REED-1]=toNum(d.reedCount)||"";row[CD.PICKS-1]=toNum(d.picksPerInch)||"";
  row[CD.SPEC_VER-1]=d.specFrom||"";
  row[CD.PCS_PROD-1]=pcsProd;     row[CD.REJ_QTY-1]=rejQty;
  row[CD.REJ_UNIT-1]=rejUnit;     row[CD.STITCH_PC-1]=stPc||"";
  row[CD.PAPERS_PR-1]=ppPr||"";   row[CD.PAPERS_REJ-1]=ppRej||"";
  row[CD.PAPER_SRC-1]=d.paperSource||"";
  row[CD.OPS_DAY-1]=opsD;         row[CD.OPS_NIGHT-1]=opsN;
  row[CD.RUN_TYPE-1]=d.runType||"Original";
  row[CD.BY-1]=d.submittedBy||""; row[CD.AT-1]=new Date();
  // Calculated
  row[CD.MACH_W_CM-1]=r2(machWcm);row[CD.FAB_L_CM-1]=r2(fabLcm);
  row[CD.GROSS_SQM-1]=grossSqm;   row[CD.GROSS_SQFT-1]=grossSqft;
  row[CD.SHRINK_CM-1]=r2(shrinkCm);row[CD.SHRINK_SQM-1]=shrinkSqm;
  row[CD.SHRINK_SQFT-1]=shrinkSqft;row[CD.USABLE_W-1]=r2(usableW);
  row[CD.MARGIN_W_CM-1]=r2(marginWcm);row[CD.MARGIN_L_CM-1]=r2(marginLcm);
  row[CD.CUT_W_CM-1]=r2(cutWcm);  row[CD.CUT_L_CM-1]=r2(cutLcm);
  row[CD.PCS_IN_L-1]=pcsInL;      row[CD.TOTAL_PCS-1]=totalPcs;
  row[CD.NET_CONS_SQM-1]=netConsSqm;row[CD.GAP_SQM-1]=gapSqm;
  row[CD.WASTE_SQM-1]=wasteSqm;   row[CD.WASTE_SQFT-1]=wasteSqft;
  row[CD.WASTE_PCT-1]=wastePct;   row[CD.WASTE_COST-1]=wasteCost;
  row[CD.GAP_SQFT-1]=gapSqft;
  row[CD.SQM_PC-1]=sqmPc;         row[CD.SQFT_PC-1]=sqftPc;
  row[CD.TOT_SQM-1]=totSqm;       row[CD.TOT_SQFT-1]=totSqft;
  row[CD.REJ_SQM-1]=rejSqm;       row[CD.REJ_SQFT-1]=rejSqft;
  row[CD.NET_SQM-1]=netSqm;       row[CD.NET_SQFT-1]=netSqft;
  row[CD.REJ_RATE-1]=rejRate;
  row[CD.TOT_ST-1]=totSt||"";     row[CD.ST_WASTE-1]=stWaste||"";
  row[CD.NET_ST-1]=netSt||"";
  row[CD.TOT_OPS-1]=totOps;       row[CD.FORMULA-1]=rate.formula;
  row[CD.OP_RATE-1]=rate.opRate;  row[CD.FIXED_RATE-1]=fixedCost;
  row[CD.MANPWR-1]=manpwr;        row[CD.FIXED_COST-1]=fixedCost;
  row[CD.DAY_COST-1]=dayCost;
  row[CD.SHARE_PCT-1]=100;        row[CD.ATTR_COST-1]=attrCost;
  row[CD.CPT_SQM-1]=cptSqm;      row[CD.CPT_SQFT-1]=cptSqft;
  row[CD.CPT_1000-1]=cpt1000;     row[CD.COST_LOSS-1]=costLoss;
  row[CD.DIFF_SCORE-1]=diffScore; row[CD.NORM_OUT-1]=normOut;

  tab(TABS.DAILY).appendRow(row);

  // ── NIGHT SHIFT: SAME AS DAY ──────────────────────────────
  // If owner selected "Same as Day", create an identical second row
  // with Shift = "Night". All calculations are identical — same output,
  // same article, same machine, same operators. Row ID incremented.
  if(d.nightShift === "Same") {
    const nightRow = [...row];
    nightRow[CD.ROW_ID-1] = nextId(TABS.DAILY,"DO",CD.ROW_ID,5);
    nightRow[CD.SHIFT-1]  = "Night";
    nightRow[CD.AT-1]     = new Date();
    tab(TABS.DAILY).appendRow(nightRow);
  }

  // Recalc proportional cost share across all entries same dept+date
  recalcShare(d.dept, date);
  recalcTargets();
  trySendConfirmation(d,{rowId,totSqft,totSqm,totSt,cptSqft,cpt1000,dayCost,rejRate,wastePct,rate});

  return {status:"ok",rowId,totSqft,totSqm,totSt:totSt||0,
    cptSqft,cpt1000,dayCost,rejRate,wastePct,diffScore,formula:rate.formula,
    nightDuplicated: d.nightShift==="Same"};
}

// ─── RECALC PROPORTIONAL SHARE ────────────────────────────────
function recalcShare(dept, date) {
  const sh=tab(TABS.DAILY), last=sh.getLastRow();
  if(last<2) return;
  const dateStr=fmtDate(date);
  const allData=sh.getRange(2,1,last-1,94).getValues();
  const matches=[];
  for(let i=0;i<allData.length;i++) {
    const r=allData[i];
    if(r[CD.DEPT-1]!==dept) continue;
    const rd=r[CD.DATE-1]; if(!rd) continue;
    if(fmtDate(rd)!==dateStr) continue;
    if(r[CD.RUN_TYPE-1]==="Sampling") continue;
    matches.push({row:i+2,
      sqft:toNum(r[CD.TOT_SQFT-1]),st:toNum(r[CD.TOT_ST-1]),
      dayCost:toNum(r[CD.DAY_COST-1]),
      rejSqft:toNum(r[CD.REJ_SQFT-1]),grossSqm:toNum(r[CD.GROSS_SQM-1]),
      wasteSqm:toNum(r[CD.WASTE_SQM-1])});
  }
  if(!matches.length) return;
  const isEmb=dept==="EMBROIDERY";
  const totalOut=matches.reduce((s,r)=>s+(isEmb?r.st:r.sqft),0);
  const dayCost=matches[0].dayCost;
  for(const m of matches) {
    const out=isEmb?m.st:m.sqft;
    const share=totalOut>0?r2((out/totalOut)*100):0;
    const attr=r2(dayCost*share/100);
    const cSqft=m.sqft>0?r2(attr/m.sqft):0;
    const cSqm=m.sqft>0?r2(attr/(m.sqft/10.7639)):0;
    const c1000=m.st>0?r2(attr/m.st*1000):0;
    const cLoss=m.sqft>0?r2((m.rejSqft/m.sqft)*attr):0;
    const wCost=m.grossSqm>0?r2((m.wasteSqm/m.grossSqm)*attr):0;
    sh.getRange(m.row,CD.SHARE_PCT,1,1).setValue(share);
    sh.getRange(m.row,CD.ATTR_COST,1,1).setValue(attr);
    sh.getRange(m.row,CD.CPT_SQFT, 1,1).setValue(cSqft);
    sh.getRange(m.row,CD.CPT_SQM,  1,1).setValue(cSqm);
    sh.getRange(m.row,CD.CPT_1000, 1,1).setValue(c1000);
    sh.getRange(m.row,CD.COST_LOSS,1,1).setValue(cLoss);
    sh.getRange(m.row,CD.WASTE_COST,1,1).setValue(wCost);
  }
}

// ─── SUBMIT SAMPLING ──────────────────────────────────────────
function submitSampling(d) {
  const date=new Date(d.date),rowId=nextId(TABS.SAMPLING,"SL",CS.ID,4);
  const created=parseInt(d.designsCreated)||0,sent=parseInt(d.designsSent)||0;
  const nDes=parseInt(d.noDesigners)||1,conv=created>0?r2(sent/created*100):0;
  const dateStr=fmtDate(date);
  const sampCost=rows(TABS.DAILY)
    .filter(r=>{const rd=r[CD.DATE-1];return rd&&fmtDate(rd)===dateStr&&r[CD.DEPT-1]===d.dept&&r[CD.RUN_TYPE-1]==="Sampling";})
    .reduce((s,r)=>s+toNum(r[CD.ATTR_COST-1]),0);
  const cpd=sent>0?r2(sampCost/sent):0;
  const row=new Array(12).fill("");
  row[CS.ID-1]=rowId;row[CS.DATE-1]=date;row[CS.PERIOD-1]=getPeriod(date);
  row[CS.DEPT-1]=d.dept;row[CS.N_DES-1]=nDes;row[CS.CREATED-1]=created;
  row[CS.SENT-1]=sent;row[CS.CONV-1]=conv;row[CS.COST-1]=r2(sampCost);
  row[CS.CPD-1]=cpd;row[CS.BY-1]=d.submittedBy||"";row[CS.AT-1]=new Date();
  tab(TABS.SAMPLING).appendRow(row);
  return {status:"ok",rowId,sampCost:r2(sampCost),cpd,conv};
}

// ─── RECALC OUTPUT TARGETS ────────────────────────────────────
function recalcTargets() {
  const rateRows=rows(TABS.RATES),sh=tab(TABS.TARGETS),wd=MIS_CONFIG.WORKING_DAYS;
  // ← UPDATE WEAVING OPERATOR COUNT HERE when confirmed with Somdutt
  const opCounts={KNITTING:6,EMBROIDERY:10,WEAVING:0,
    TUFTING_CUT:0,TUFTING_LOOP:0,TUFTING_ROBO:0,PRINTING:0};
  const latest={};
  for(const r of rateRows){if(r[CR.DEPT-1])latest[r[CR.DEPT-1]]=r;}
  const out=[];
  for(const[dept,r]of Object.entries(latest)){
    const formula=r[CR.FORMULA-1],opRate=toNum(r[CR.OP_RATE-1]);
    const fixed=toNum(r[CR.FIXED-1]),green=toNum(r[CR.GREEN-1]);
    const red=toNum(r[CR.RED-1]),unit=r[CR.UNIT-1]||"/sqft";
    const ops=opCounts[dept]||0;
    const dayCost=formula==="Type_A"?fixed:r2(ops*opRate+fixed);
    const yellow=green>0&&red>0?r2((green+red)/2):green;
    const gSqft=green>0?r2(dayCost/green):0;
    const ySqft=yellow>0?r2(dayCost/yellow):0;
    const rSqft=red>0?r2(dayCost/red):0;
    const isEm=unit.includes("1000");
    out.push([dept,dayCost,wd,gSqft,ySqft,rSqft,
      r2(gSqft/10.7639),r2(ySqft/10.7639),r2(rSqft/10.7639),
      isEm&&green>0?Math.round(dayCost/green*1000):0,
      isEm&&red>0?Math.round(dayCost/red*1000):0,
      r2(gSqft*wd),r2(rSqft*wd),new Date()]);
  }
  const last=sh.getLastRow();
  if(last>1)sh.getRange(2,1,last-1,sh.getLastColumn()).clearContent();
  if(out.length)sh.getRange(2,1,out.length,out[0].length).setValues(out);
}

// ─── GET TARGETS ─────────────────────────────────────────────
function getTargets() {
  const result={};
  for(const r of rows(TABS.TARGETS)){
    if(!r[CT.DEPT-1])continue;
    result[r[CT.DEPT-1]]={dayCost:r[CT.DAY_COST-1],
      gSqft:r[CT.G_SQFT-1],ySqft:r[CT.Y_SQFT-1],rSqft:r[CT.R_SQFT-1],
      gSqm:r[CT.G_SQM-1],gMonSqft:r[CT.G_MON-1],rMonSqft:r[CT.R_MON-1],
      g1000:r[CT.G_1000-1],r1000:r[CT.R_1000-1]};
  }
  return {status:"ok",targets:result};
}

// ─── DASHBOARD DATA ───────────────────────────────────────────
function getDashboardData(d) {
  const fac=d.facility||"PNP",now=new Date();
  return {status:"ok",asOf:now.toISOString(),
    today:todaySummary(fac,now),mtd:mtdSummary(fac,now),
    targets:getTargets().targets,flags:rejectionFlags(fac,now),
    incomplete:incompleteArticles(),compliance:complianceSummary(fac,now),
    sampling:samplingSummary(fac,now),
    cushionStatus:getCushionCompleteness(fac),
    processStatus:getProcessCompleteness(fac)};
}

function todaySummary(fac,now) {
  const ds=fmtDate(now);
  const data=rows(TABS.DAILY).filter(r=>{
    const rd=r[CD.DATE-1];
    return rd&&fmtDate(rd)===ds&&r[CD.FACILITY-1]===fac&&r[CD.RUN_TYPE-1]!=="Sampling";
  });
  const byDept={},costSeen={};
  for(const r of data){
    const dep=r[CD.DEPT-1];
    if(!byDept[dep])byDept[dep]={sqft:0,sqm:0,st:0,cost:0,rejSqft:0,wasteSqm:0,grossSqm:0};
    byDept[dep].sqft    +=toNum(r[CD.TOT_SQFT-1]);
    byDept[dep].sqm     +=toNum(r[CD.TOT_SQM-1]);
    byDept[dep].st      +=toNum(r[CD.TOT_ST-1]);
    byDept[dep].rejSqft +=toNum(r[CD.REJ_SQFT-1]);
    byDept[dep].wasteSqm+=toNum(r[CD.WASTE_SQM-1]);
    byDept[dep].grossSqm+=toNum(r[CD.GROSS_SQM-1]);
    if(!costSeen[dep]){costSeen[dep]=true;byDept[dep].cost=toNum(r[CD.DAY_COST-1]);}
  }
  for(const dep of Object.keys(byDept)){
    const b=byDept[dep];
    b.sqft=r2(b.sqft);b.sqm=r2(b.sqm);
    b.rejRate=b.sqft>0?r2(b.rejSqft/b.sqft*100):0;
    b.wastePct=b.grossSqm>0?r2(b.wasteSqm/b.grossSqm*100):0;
    b.cptSqft=b.sqft>0?r2(b.cost/b.sqft):0;
    b.cpt1000=b.st>0?r2(b.cost/b.st*1000):0;
  }
  return byDept;
}

function mtdSummary(fac,now) {
  const y=now.getFullYear(),m=now.getMonth(),start=new Date(y,m,1);
  const data=rows(TABS.DAILY).filter(r=>{
    const rd=r[CD.DATE-1];if(!rd)return false;
    const dt=rd instanceof Date?rd:new Date(rd);
    return dt>=start&&dt<=now&&r[CD.FACILITY-1]===fac&&r[CD.RUN_TYPE-1]!=="Sampling";
  });
  const byDept={},costSeen={};
  for(const r of data){
    const dep=r[CD.DEPT-1],key=dep+"|"+fmtDate(r[CD.DATE-1]);
    if(!byDept[dep])byDept[dep]={sqft:0,sqm:0,st:0,cost:0,wasteSqm:0,grossSqm:0};
    byDept[dep].sqft    +=toNum(r[CD.TOT_SQFT-1]);
    byDept[dep].sqm     +=toNum(r[CD.TOT_SQM-1]);
    byDept[dep].st      +=toNum(r[CD.TOT_ST-1]);
    byDept[dep].wasteSqm+=toNum(r[CD.WASTE_SQM-1]);
    byDept[dep].grossSqm+=toNum(r[CD.GROSS_SQM-1]);
    if(!costSeen[key]){costSeen[key]=true;byDept[dep].cost+=toNum(r[CD.DAY_COST-1]);}
  }
  const daysInMon=new Date(y,m+1,0).getDate(),dayOf=now.getDate(),daysLeft=daysInMon-dayOf;
  for(const dep of Object.keys(byDept)){
    const b=byDept[dep];
    b.sqft=r2(b.sqft);b.sqm=r2(b.sqm);
    b.avgCptSqft=b.sqft>0?r2(b.cost/b.sqft):0;
    b.avgCpt1000=b.st>0?r2(b.cost/b.st*1000):0;
    b.wastePct=b.grossSqm>0?r2(b.wasteSqm/b.grossSqm*100):0;
    const dSqft=dayOf>0?b.sqft/dayOf:0,dCost=dayOf>0?b.cost/dayOf:0;
    const pSqft=r2(b.sqft+dSqft*daysLeft),pCost=r2(b.cost+dCost*daysLeft);
    b.projCptSqft=pSqft>0?r2(pCost/pSqft):0;
    const pSt=b.st>0&&dayOf>0?b.st+b.st/dayOf*daysLeft:0;
    b.projCpt1000=pSt>0?r2(pCost/pSt*1000):0;
  }
  return byDept;
}

function rejectionFlags(fac,now) {
  const cutoff=new Date(now);cutoff.setDate(cutoff.getDate()-MIS_CONFIG.REJECTION_AVG_DAYS);
  const data=rows(TABS.DAILY).filter(r=>r[CD.FACILITY-1]===fac&&r[CD.RUN_TYPE-1]!=="Sampling");
  const todayStr=fmtDate(now),avgMap={};
  for(const r of data){
    const rd=r[CD.DATE-1];if(!rd)continue;
    const dt=rd instanceof Date?rd:new Date(rd);if(dt<cutoff)continue;
    const dep=r[CD.DEPT-1];if(!avgMap[dep])avgMap[dep]=[];
    avgMap[dep].push(toNum(r[CD.REJ_RATE-1]));
  }
  const flags=[];
  for(const r of data){
    const rd=r[CD.DATE-1];if(!rd)continue;
    if(fmtDate(rd)!==todayStr)continue;
    const dep=r[CD.DEPT-1],rate=toNum(r[CD.REJ_RATE-1]);
    const arr=avgMap[dep]||[],avg=arr.length?r2(arr.reduce((s,v)=>s+v,0)/arr.length):0;
    if(rate>avg*MIS_CONFIG.REJECTION_FLAG_MULT&&rate>MIS_CONFIG.REJECTION_FLAG_MIN)
      flags.push({dept:dep,machine:r[CD.MACHINE_ID-1],article:r[CD.ART_CODE-1],
        todayRate:rate,avgRate:avg,rejQty:r[CD.REJ_QTY-1],rejUnit:r[CD.REJ_UNIT-1]});
  }
  return flags;
}

function incompleteArticles() {
  return rows(TABS.ARTICLES).filter(r=>r[CA.STATUS-1]==="Incomplete")
    .map(r=>({code:r[CA.CODE-1],name:r[CA.NAME-1]}));
}

function complianceSummary(fac,now) {
  const cutoff=new Date(now);cutoff.setDate(cutoff.getDate()-7);
  const data=rows(TABS.DAILY).filter(r=>r[CD.FACILITY-1]===fac);
  const seen=new Set();
  for(const r of data){
    const rd=r[CD.DATE-1];if(!rd)continue;
    const dt=rd instanceof Date?rd:new Date(rd);
    if(dt>=cutoff)seen.add(r[CD.DEPT-1]+"|"+fmtDate(dt));
  }
  // ← ADD "WEAVING" HERE when weaving form is built
  const depts=["KNITTING","TUFTING_CUT","TUFTING_ROBO","TUFTING_LOOP","EMBROIDERY","PRINTING"];
  const out={};
  for(const dep of depts){
    let n=0;
    for(let i=0;i<7;i++){const d=new Date(now);d.setDate(d.getDate()-i);if(seen.has(dep+"|"+fmtDate(d)))n++;}
    out[dep]={submitted:n,expected:7,pct:r2(n/7*100)};
  }
  return out;
}

function samplingSummary(fac,now) {
  const y=now.getFullYear(),m=now.getMonth(),start=new Date(y,m,1);
  const data=rows(TABS.SAMPLING).filter(r=>{
    const rd=r[CS.DATE-1];if(!rd)return false;
    const dt=rd instanceof Date?rd:new Date(rd);return dt>=start&&dt<=now;
  });
  const s={created:0,sent:0,cost:0};
  for(const r of data){s.created+=parseInt(r[CS.CREATED-1])||0;s.sent+=parseInt(r[CS.SENT-1])||0;s.cost+=toNum(r[CS.COST-1]);}
  return {...s,conv:s.created>0?r2(s.sent/s.created*100):0,cpd:s.sent>0?r2(s.cost/s.sent):0};
}

// ─── CUSHION COMPLETENESS ─────────────────────────────────────
// Tracks Front/Back panel submission status per article per buyer
function getCushionCompleteness(fac) {
  const data=rows(TABS.DAILY).filter(r=>
    r[CD.FACILITY-1]===fac&&
    r[CD.DEPT-1]==="WEAVING"&&
    r[CD.CATEGORY-1]==="Cushion"
  );
  const map={};
  for(const r of data){
    const key=r[CD.ART_CODE-1]+"|"+r[CD.BUYER-1];
    if(!map[key])map[key]={article:r[CD.ART_CODE-1],buyer:r[CD.BUYER-1],front:null,back:null};
    const pt=r[CD.PANEL_TYPE-1];
    const dt=fmtDate(r[CD.DATE-1]);
    const sqm=toNum(r[CD.TOT_SQM-1]);
    if(pt==="Front"||pt==="Both")map[key].front={date:dt,sqm};
    if(pt==="Back"||pt==="Both") map[key].back ={date:dt,sqm};
  }
  return Object.values(map).map(v=>({
    ...v,
    complete:!!(v.front&&v.back),
    status:v.front&&v.back?"✓ Complete":v.front?"⟳ Back Pending":v.back?"⟳ Front Pending":"⟳ Both Pending"
  }));
}

// ─── PROCESS COMPLETENESS ────────────────────────────────────
// For each article+buyer, checks which processes have submissions
function getProcessCompleteness(fac) {
  const artRows=rows(TABS.ARTICLES).filter(r=>!r[CA.SPEC_TO-1]);
  const dailyRows=rows(TABS.DAILY).filter(r=>r[CD.FACILITY-1]===fac&&r[CD.RUN_TYPE-1]!=="Sampling");
  const submitted=new Set();
  for(const r of dailyRows) submitted.add(r[CD.ART_CODE-1]+"|"+r[CD.BUYER-1]+"|"+r[CD.DEPT-1]);
  const result=[];
  const seen=new Set();
  for(const r of dailyRows){
    const artCode=r[CD.ART_CODE-1],buyer=r[CD.BUYER-1];
    const key=artCode+"|"+buyer;
    if(seen.has(key))continue;seen.add(key);
    const artMaster=artRows.find(a=>a[CA.CODE-1]===artCode);
    if(!artMaster)continue;
    const processes=[artMaster[CA.P1-1],artMaster[CA.P2-1],artMaster[CA.P3-1],artMaster[CA.P4-1]].filter(Boolean);
    const status=processes.map(p=>({
      process:p,
      submitted:submitted.has(artCode+"|"+buyer+"|"+p.toUpperCase()),
    }));
    const complete=status.every(s=>s.submitted);
    result.push({article:artCode,buyer,processes:status,complete});
  }
  return result;
}

// ─── WASENDER ────────────────────────────────────────────────
function sendWasender(phone,msg) {
  if(!MIS_CONFIG.WASENDER_URL||MIS_CONFIG.WASENDER_URL.includes("REPLACE"))return;
  try{UrlFetchApp.fetch(MIS_CONFIG.WASENDER_URL,{method:"post",contentType:"application/json",
    headers:{"Authorization":"Bearer "+MIS_CONFIG.WASENDER_TOKEN},
    payload:JSON.stringify({phone,message:msg}),muteHttpExceptions:true});}
  catch(e){Logger.log("Wasender: "+e);}
}

const DEPT_HI={KNITTING:"निटिंग",WEAVING:"बुनाई",EMBROIDERY:"कढ़ाई",
  TUFTING_CUT:"टफ्टिंग कट",TUFTING_LOOP:"टफ्टिंग लूप",TUFTING_ROBO:"टफ्टिंग रोबो",PRINTING:"प्रिंटिंग"};

function trySendConfirmation(d,calc) {
  try{
    const head=rows(TABS.HEADS).find(r=>String(r[CH.DEPTS-1]).includes(d.dept)&&r[CH.ACTIVE-1]==="Y");
    if(!head||!head[CH.PHONE-1])return;
    const ds=Utilities.formatDate(new Date(d.date),Session.getScriptTimeZone(),"dd/MM/yyyy");
    const isEmb=calc.rate.unit&&calc.rate.unit.includes("1000");
    const outLine=isEmb?`${(calc.totSt||0).toLocaleString()} stitches`:`${calc.totSqft} sqft / ${calc.totSqm} sqm`;
    const cLine=isEmb?`₹${calc.cpt1000}/1000 stitches`:`₹${calc.cptSqft}/sqft`;
    const wasteStr=calc.wastePct>0?`\nWastage: ${calc.wastePct}%`:"";
    const msg=
`✓ MIS Updated — ${ds} (${d.shift})
Dept: ${d.dept} | Machine: ${d.machineId}
Output: ${outLine}
Cost: ${cLine} | Day Total: ₹${calc.dayCost}
Rejection: ${calc.rejRate}%${wasteStr}

✓ MIS अपडेट हुआ — ${ds} (${d.shift==="Day"?"दिन":"रात"})
विभाग: ${DEPT_HI[d.dept]||d.dept} | मशीन: ${d.machineId}
उत्पादन: ${outLine}
लागत: ${cLine} | कुल: ₹${calc.dayCost}
अस्वीकृति: ${calc.rejRate}%${wasteStr}`;
    sendWasender(String(head[CH.PHONE-1]),msg);
  }catch(e){Logger.log("Confirmation failed: "+e);}
}

// ─── REMINDERS (Mon+Sat 10:30 + 16:30) ───────────────────────
function sendReminders() {
  const now=new Date(),day=now.getDay(),hour=now.getHours();
  if(!MIS_CONFIG.REMINDER_DAYS.includes(day))return;
  if(!MIS_CONFIG.REMINDER_HOURS.includes(hour))return;
  const url="https://mav0715.github.io/pcs-Champo/mis-output-form.html";
  for(const r of rows(TABS.HEADS)){
    if(r[CH.ACTIVE-1]!=="Y"||!r[CH.PHONE-1])continue;
    const name=r[CH.NAME-1],pin=r[CH.PIN-1];
    sendWasender(String(r[CH.PHONE-1]),
`Hi ${name}, please update today's production data on the MIS portal.
Your PIN: ${pin}
Link: ${url}

नमस्ते ${name}, कृपया आज का उत्पादन डेटा MIS पोर्टल पर अपडेट करें।
आपका PIN: ${pin}
Link: ${url}`);
  }
}

// ─── WEEKLY EMAIL (Sunday 8am) ────────────────────────────────
function sendWeeklyEmail() {
  const now=new Date(),cutoff=new Date(now);cutoff.setDate(cutoff.getDate()-7);
  const data=rows(TABS.DAILY);
  // ← ADD "WEAVING" HERE when weaving form is built
  const depts=["KNITTING","TUFTING_CUT","TUFTING_ROBO","TUFTING_LOOP","EMBROIDERY","PRINTING"];
  const subMap={};
  for(const r of data){
    const rd=r[CD.DATE-1];if(!rd)continue;
    const dt=rd instanceof Date?rd:new Date(rd);if(dt<cutoff)continue;
    const dep=r[CD.DEPT-1];subMap[dep]=(subMap[dep]||0)+1;
  }
  let html=`<h2 style="font-family:sans-serif">MIS Weekly Report — ${Utilities.formatDate(now,Session.getScriptTimeZone(),"dd MMM yyyy")}</h2>`;
  html+=`<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-family:sans-serif;font-size:14px">`;
  html+=`<tr style="background:#0d1117;color:#fff"><th>Dept Head</th><th>Department</th><th>Submissions</th><th>Status</th></tr>`;
  for(const h of rows(TABS.HEADS)){
    if(h[CH.ACTIVE-1]!=="Y")continue;
    const name=h[CH.NAME-1];
    for(const dep of String(h[CH.DEPTS-1]).split(",").map(x=>x.trim())){
      if(!depts.includes(dep))continue;
      const n=subMap[dep]||0;
      const status=n>=5?"✓ Active":n>=2?"⚠️ Low":"✗ Not Responding";
      const color=n>=5?"#22c55e":n>=2?"#f59e0b":"#ef4444";
      html+=`<tr><td>${name}</td><td>${dep}</td><td style="text-align:center">${n}</td><td style="color:${color};font-weight:bold">${status}</td></tr>`;
    }
  }
  html+=`</table>`;
  const inc=incompleteArticles();
  if(inc.length){
    html+=`<h3 style="font-family:sans-serif;color:#f59e0b">⚠️ ${inc.length} Articles Need Completion</h3>`;
    html+=`<ul style="font-family:sans-serif">${inc.map(a=>`<li>${a.code} — ${a.name}</li>`).join("")}</ul>`;
  }
  html+=`<p style="font-family:sans-serif;color:#636e7b;font-size:12px">Reminders: Mon+Sat 10:30am & 4:30pm</p>`;
  GmailApp.sendEmail(MIS_CONFIG.ADMIN_EMAIL,
    `MIS Weekly Report — ${Utilities.formatDate(now,Session.getScriptTimeZone(),"dd MMM yyyy")}`,
    "View in HTML mode.",{htmlBody:html});
}

// ─── CURRENCY CONVERSION ──────────────────────────────────────
// Real-time INR → USD / AUD / CAD / EUR / GBP
// Monthly avg rates saved as permanent record in CURRENCY_LOG tab.
// Quarterly + yearly averages derived from monthly records.
//
// CURRENCY_LOG tab columns (auto-created if missing):
// Month_Key | Fetched_At | INR_USD | INR_AUD | INR_CAD | INR_EUR | INR_GBP |
// Avg_Cost_INR | Avg_Cost_USD | Avg_Cost_AUD | Avg_Cost_CAD | Avg_Cost_EUR | Avg_Cost_GBP |
// Dept | Source

const FX_API = "https://api.exchangerate-api.com/v4/latest/INR";
const FX_CURRENCIES = ["USD","AUD","CAD","EUR","GBP"];

// Called from form confirmation screen + dashboard
// Returns live rates + costs converted to all currencies
function getCurrencyRates(data) {
  const avgCostInr = toNum(data.avgCostInr); // cost per sqft or per 1000st in INR
  const dept       = data.dept || "";

  // Fetch live rates
  let rates = {};
  try {
    const resp = UrlFetchApp.fetch(FX_API, {muteHttpExceptions:true});
    const json = JSON.parse(resp.getContentText());
    FX_CURRENCIES.forEach(c => { rates[c] = r4(json.rates[c] || 0); });
  } catch(e) {
    Logger.log("FX fetch failed: " + e);
    // Fall back to last saved rates for this month
    rates = getLastSavedRates();
  }

  // Convert avg cost
  const converted = {};
  FX_CURRENCIES.forEach(c => {
    converted[c] = rates[c] > 0 ? r4(avgCostInr * rates[c]) : 0;
  });

  return {
    status: "ok",
    rates,
    avgCostInr,
    converted,
    timestamp: new Date().toISOString(),
  };
}

// Get last saved rates from CURRENCY_LOG (fallback when API fails)
function getLastSavedRates() {
  try {
    const data = rows(TABS.FX);
    if (!data.length) return {};
    const last = data[data.length - 1];
    return {
      USD: toNum(last[2]), AUD: toNum(last[3]),
      CAD: toNum(last[4]), EUR: toNum(last[5]), GBP: toNum(last[6]),
    };
  } catch(e) { return {}; }
}

// Save monthly avg FX rate + avg cost in all currencies
// Called by monthly trigger on 1st of each month
function saveMonthlyFXRecord() {
  const now    = new Date();
  const y      = now.getFullYear();
  const m      = now.getMonth(); // 0-based
  // Build month key for PREVIOUS month (we're saving last month's avg)
  const prevM  = m === 0 ? 11 : m - 1;
  const prevY  = m === 0 ? y - 1 : y;
  const monthKey = `${prevY}-${String(prevM + 1).padStart(2,"0")}`;

  // Check if already saved for this month
  try {
    const existing = rows(TABS.FX);
    if (existing.some(r => String(r[0]) === monthKey)) {
      Logger.log("FX record already exists for " + monthKey);
      return;
    }
  } catch(e) {}

  // Fetch current rates (snapshot at month-end)
  let rates = {};
  try {
    const resp = UrlFetchApp.fetch(FX_API, {muteHttpExceptions:true});
    const json = JSON.parse(resp.getContentText());
    FX_CURRENCIES.forEach(c => { rates[c] = r4(json.rates[c] || 0); });
  } catch(e) {
    Logger.log("FX monthly save failed — API error: " + e);
    return;
  }

  // Calculate avg cost per dept for the previous month from DAILY_OUTPUT
  const depts = ["KNITTING","TUFTING_CUT","TUFTING_ROBO","TUFTING_LOOP","EMBROIDERY","PRINTING","WEAVING"];
  const start = new Date(prevY, prevM, 1);
  const end   = new Date(prevY, prevM + 1, 0); // last day of prev month
  const dailyData = rows(TABS.DAILY).filter(r => {
    const rd = r[CD.DATE-1]; if(!rd) return false;
    const dt = rd instanceof Date ? rd : new Date(rd);
    return dt >= start && dt <= end && r[CD.RUN_TYPE-1] !== "Sampling";
  });

  const sh = tab(TABS.FX);

  for (const dept of depts) {
    const deptRows = dailyData.filter(r => r[CD.DEPT-1] === dept);
    if (!deptRows.length) continue;

    const isEmb = dept === "EMBROIDERY";
    const totalCost = (() => {
      // Sum cost once per day (not per row — fixed cost counted once/day)
      const seen = new Set();
      let total = 0;
      deptRows.forEach(r => {
        const key = fmtDate(r[CD.DATE-1]);
        if(!seen.has(key)) { seen.add(key); total += toNum(r[CD.DAY_COST-1]); }
      });
      return total;
    })();

    const totalOut = deptRows.reduce((s,r) =>
      s + toNum(isEmb ? r[CD.TOT_ST-1] : r[CD.TOT_SQFT-1]), 0);

    const avgCostInr = isEmb
      ? (totalOut > 0 ? r4(totalCost / totalOut * 1000) : 0)   // per 1000 stitches
      : (totalOut > 0 ? r4(totalCost / totalOut) : 0);          // per sqft

    const row = [
      monthKey,
      new Date().toISOString(),
      rates.USD, rates.AUD, rates.CAD, rates.EUR, rates.GBP,
      avgCostInr,
      rates.USD > 0 ? r4(avgCostInr * rates.USD) : 0,
      rates.AUD > 0 ? r4(avgCostInr * rates.AUD) : 0,
      rates.CAD > 0 ? r4(avgCostInr * rates.CAD) : 0,
      rates.EUR > 0 ? r4(avgCostInr * rates.EUR) : 0,
      rates.GBP > 0 ? r4(avgCostInr * rates.GBP) : 0,
      dept,
      "exchangerate-api.com",
    ];
    sh.appendRow(row);
  }

  Logger.log("✅ Monthly FX record saved for " + monthKey);
}

// Get historical currency data for dashboard
// Returns monthly records + derived quarterly + yearly averages
function getCurrencyHistory(data) {
  const dept = data.dept || "";
  const allRows_ = rows(TABS.FX);
  const filtered = dept
    ? allRows_.filter(r => r[13] === dept)
    : allRows_;

  if (!filtered.length) return {status:"ok", monthly:[], quarterly:[], yearly:[]};

  // Monthly records
  const monthly = filtered.map(r => ({
    monthKey:    r[0],
    avgCostInr:  toNum(r[7]),
    avgCostUsd:  toNum(r[8]),
    avgCostAud:  toNum(r[9]),
    avgCostCad:  toNum(r[10]),
    avgCostEur:  toNum(r[11]),
    avgCostGbp:  toNum(r[12]),
    dept:        r[13],
    rateUsd:     toNum(r[2]),
    rateAud:     toNum(r[3]),
    rateCad:     toNum(r[4]),
    rateEur:     toNum(r[5]),
    rateGbp:     toNum(r[6]),
  }));

  // Quarterly averages — group by year+quarter
  const qMap = {};
  monthly.forEach(m => {
    const [y, mo] = m.monthKey.split("-").map(Number);
    const q = Math.ceil(mo / 3);
    const qKey = `${y}-Q${q}|${m.dept}`;
    if (!qMap[qKey]) qMap[qKey] = {key:`${y}-Q${q}`,dept:m.dept,records:[],
      inr:0,usd:0,aud:0,cad:0,eur:0,gbp:0};
    qMap[qKey].records.push(m);
    qMap[qKey].inr += m.avgCostInr; qMap[qKey].usd += m.avgCostUsd;
    qMap[qKey].aud += m.avgCostAud; qMap[qKey].cad += m.avgCostCad;
    qMap[qKey].eur += m.avgCostEur; qMap[qKey].gbp += m.avgCostGbp;
  });
  const quarterly = Object.values(qMap).map(q => {
    const n = q.records.length;
    return {quarterKey:q.key,dept:q.dept,months:n,
      avgCostInr:r4(q.inr/n),avgCostUsd:r4(q.usd/n),
      avgCostAud:r4(q.aud/n),avgCostCad:r4(q.cad/n),
      avgCostEur:r4(q.eur/n),avgCostGbp:r4(q.gbp/n)};
  });

  // Yearly averages — group by year
  const yMap = {};
  monthly.forEach(m => {
    const y = m.monthKey.split("-")[0];
    const yKey = `${y}|${m.dept}`;
    if (!yMap[yKey]) yMap[yKey] = {year:y,dept:m.dept,records:[],
      inr:0,usd:0,aud:0,cad:0,eur:0,gbp:0};
    yMap[yKey].records.push(m);
    yMap[yKey].inr += m.avgCostInr; yMap[yKey].usd += m.avgCostUsd;
    yMap[yKey].aud += m.avgCostAud; yMap[yKey].cad += m.avgCostCad;
    yMap[yKey].eur += m.avgCostEur; yMap[yKey].gbp += m.avgCostGbp;
  });
  const yearly = Object.values(yMap).map(y => {
    const n = y.records.length;
    return {year:y.year,dept:y.dept,months:n,
      avgCostInr:r4(y.inr/n),avgCostUsd:r4(y.usd/n),
      avgCostAud:r4(y.aud/n),avgCostCad:r4(y.cad/n),
      avgCostEur:r4(y.eur/n),avgCostGbp:r4(y.gbp/n)};
  });

  return {status:"ok", monthly, quarterly, yearly};
}

// ─── INSTALL TRIGGERS — RUN ONCE ─────────────────────────────
function installTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t=>["sendReminders","sendWeeklyEmail","recalcTargets","saveMonthlyFXRecord"].includes(t.getHandlerFunction()))
    .forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("sendReminders").timeBased().everyHours(1).create();
  ScriptApp.newTrigger("sendWeeklyEmail").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(8).create();
  ScriptApp.newTrigger("recalcTargets").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  // Monthly FX record — runs on 1st of every month at 6am
  ScriptApp.newTrigger("saveMonthlyFXRecord").timeBased().onMonthDay(1).atHour(6).create();
  Logger.log("✅ Triggers installed. Run recalcTargets() now to seed OUTPUT_TARGETS.");
}
