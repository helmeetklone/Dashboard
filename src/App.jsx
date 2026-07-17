import { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  PieChart, Pie, Cell, ResponsiveContainer, LineChart, Line, LabelList } from "recharts";

// ── THEMES ───────────────────────────────────────────────────────────────────
const DARK  = { bg:"#060d1a",card:"#0c1a2e",cardAlt:"#081422",border:"#162840",text:"#ddeeff",muted:"#5a7fa8",inputBg:"#0c1a2e",rowAlt:"rgba(255,255,255,0.02)",rowHover:"rgba(37,99,235,0.08)" };
const LIGHT = { bg:"#f0f4fa",card:"#ffffff",cardAlt:"#f5f8ff",border:"#d0dff0",text:"#0f2040",muted:"#6b8aad",inputBg:"#ffffff",rowAlt:"rgba(0,0,0,0.02)",rowHover:"rgba(37,99,235,0.05)" };


// ── DEFAULT VALIDATION PARAMETERS ────────────────────────────────────────────
const DEFAULT_PARAMS = {
  dur_short: 2,      // menit — di bawah ini = SHORT
  dur_long:  30,     // menit — di atas ini = LONG
  dis_near:  50,     // meter — di bawah ini = NEAR
  dis_far:   200,    // meter — di atas ini = FAR
};

const P = {
  a1:"#22c55e",a2:"#f59e0b",a3:"#3b82f6",
  valid:"#22c55e",observe:"#f59e0b",investigate:"#ef4444",incomplete:"#3b82f6",
  normal:"#22c55e",short:"#f97316",long:"#a855f7",
  near:"#22c55e",mid:"#f59e0b",far:"#ef4444",match:"#22c55e",notmatch:"#ef4444",
  accent:"#f0a63d",
  regions:["#6366f1","#ec4899","#14b8a6","#f97316","#8b5cf6","#06b6d4","#84cc16","#f43f5e","#a78bfa","#34d399"],
};

const ACT = [
  {key:"A1 - NORMAL",    label:"A1 - Normal",    short:"A1",color:P.a1},
  {key:"A2 - ANOMALY",   label:"A2 - Anomaly",   short:"A2",color:P.a2},
  {key:"A3 - INCOMPLETE",label:"A3 - Incomplete", short:"A3",color:P.a3},
];
const VIS = [
  {key:"VALID",label:"Valid",color:P.valid},{key:"OBSERVE",label:"Observe",color:P.observe},
  {key:"INVESTIGATE",label:"Investigate",color:P.investigate},{key:"INCOMPLETE",label:"Incomplete",color:P.incomplete},
];

const pct  = (n,d) => d ? +((n/d)*100).toFixed(1) : 0;
const pctS = (n,d) => pct(n,d).toFixed(1)+"%";
const fmtK = n => n>=1000?(n/1000).toFixed(1)+"K":String(n||0);


// ── COMPUTE VALIDATION COLUMNS from raw data ─────────────────────────────────
function haversineM(lat1,lon1,lat2,lon2){
  const R=6371000,toR=Math.PI/180;
  const dLat=(lat2-lat1)*toR, dLon=(lon2-lon1)*toR;
  const sd=Math.sin(dLat/2),sl=Math.sin(dLon/2);
  const a=sd*sd+Math.cos(lat1*toR)*Math.cos(lat2*toR)*sl*sl;
  return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}

function computeValidation(row, params=DEFAULT_PARAMS){
  const p = {...DEFAULT_PARAMS, ...params};
  // Already has computed columns? Use them
  if(row["Visit Status"]&&["VALID","OBSERVE","INVESTIGATE","INCOMPLETE"].includes(String(row["Visit Status"]).toUpperCase()))
    return row;

  const r = {...row};

  // 1. Visit Duration
  const tIn  = row["Actual Visit Time"];
  const tOut = row["Actual Check-Out Time"];
  const hasIn  = tIn  != null && tIn  !== "";
  const hasOut = tOut != null && tOut !== "";
  let durMin = null;
  if(hasIn && hasOut){
    const d1=new Date(tIn), d2=new Date(tOut);
    if(!isNaN(d1)&&!isNaN(d2)) durMin=(d2-d1)/60000;
  }
  if(durMin===null && row["Visit Duration (Menit)"]!=null)
    durMin = parseFloat(row["Visit Duration (Menit)"]);
  r["Visit Duration (Menit)"] = durMin;

  // 2. Distance Check In/Out (from GPS coords if missing)
  let distIn  = parseFloat(row["Distance Check In (Meter)"])||null;
  let distOut = parseFloat(row["Distance Check Out (Meter)"])||null;
  const roLat = parseFloat(row["RO Latitude"]);
  const roLon = parseFloat(row["RO Longitude"]);
  if((distIn===null||isNaN(distIn)) && !isNaN(roLat)&&!isNaN(roLon)){
    const ciLat=parseFloat(row["Check-In Latitude"]), ciLon=parseFloat(row["Check-In Longitude"]);
    if(!isNaN(ciLat)&&!isNaN(ciLon)) distIn=haversineM(ciLat,ciLon,roLat,roLon);
  }
  if((distOut===null||isNaN(distOut)) && !isNaN(roLat)&&!isNaN(roLon)){
    const coLat=parseFloat(row["Check-Out Latitude"]), coLon=parseFloat(row["Check-Out Longitude"]);
    if(!isNaN(coLat)&&!isNaN(coLon)) distOut=haversineM(coLat,coLon,roLat,roLon);
  }
  r["Distance Check In (Meter)"]  = distIn  ?? row["Distance Check In (Meter)"];
  r["Distance Check Out (Meter)"] = distOut ?? row["Distance Check Out (Meter)"];

  // 3. Duration Status
  const durSt = !hasIn||!hasOut?"INCOMPLETE"
    : durMin<p.dur_short?"SHORT"
    : durMin>p.dur_long?"LONG"
    : "NORMAL";
  r["Duration Status"] = durSt;
  r["_DUR"] = durSt;

  // 4. Distance Status & Location Status
  const maxDist = Math.max(distIn||0, distOut||0);
  const disSt = !hasIn||!hasOut?"INCOMPLETE"
    : maxDist<=p.dis_near?"NEAR"
    : maxDist<=p.dis_far?"MID"
    : "FAR";
  const locSt = !hasIn||!hasOut?"INCOMPLETE"
    : disSt==="NEAR"?"MATCH":"NOT MATCH";
  r["Distance Status"]  = disSt;
  r["Location Status"]  = locSt;
  r["_DIS"] = disSt;
  r["_LOC"] = locSt;

  // 5. Visit Status
  const vs = !hasIn||!hasOut?"INCOMPLETE"
    : durSt==="SHORT"&&(disSt==="MID"||disSt==="FAR")&&locSt==="NOT MATCH"?"INVESTIGATE"
    : durSt==="NORMAL"&&disSt==="NEAR"&&locSt==="MATCH"?"VALID"
    : "OBSERVE";
  r["Visit Status"] = vs;
  r["_VS"] = vs;

  // 6. Activity Status
  const as1 = vs==="VALID"?"A1 - NORMAL"
    : vs==="INCOMPLETE"?"A3 - INCOMPLETE"
    : "A2 - ANOMALY";
  r["Activity Status"] = as1;

  return r;
}

function extractDate(val) {
  if(!val) return null;
  if(typeof val==="number"){ const d=new Date(Math.round((val-25569)*86400*1000)); return d.toISOString().slice(0,10); }
  const d=new Date(val); return isNaN(d)?null:d.toISOString().slice(0,10);
}
function getRegionCode(cl) {
  if(!cl) return "??";
  const m=cl.match(/^([A-Z]{2,3})[- _]/); return m?m[1]:cl.slice(0,3).toUpperCase();
}

// Mapping kode region ke kepanjangannya
const REGION_NAMES = {
  "BN":"Bali Nusra","CJ":"Central Java","EJ":"East Java","JB":"Jabotabek",
  "KM":"Kalimantan","NS":"North Sumatera","SS":"South Sumatera",
  "SW":"Sulawesi","WJ":"West Java",
};
function fmtPeriod(dr){
  if(!dr||!dr.min) return null;
  const fmt=d=>new Date(d).toLocaleDateString("id-ID",{day:"2-digit",month:"short",year:"numeric"});
  if(dr.min===dr.max) return fmt(dr.min);
  const days=Math.round((new Date(dr.max)-new Date(dr.min))/(1000*60*60*24))+1;
  return `${fmt(dr.min)} – ${fmt(dr.max)} (${days} hari)`;
}

// ── GROUP TREND DATA ──────────────────────────────────────────────────────────
function groupTrend(trend, period) {
  if(!trend||!trend.length) return [];
  if(period==="daily") return trend.map(d=>({...d,name:d.date.slice(5),_date:d.date}));
  const map={};
  trend.forEach(d=>{
    const dt=new Date(d.date);
    let key="";
    if(period==="weekly"){
      // ISO week: Monday-based
      const tmp=new Date(dt); tmp.setHours(0,0,0,0);
      tmp.setDate(tmp.getDate()-((tmp.getDay()+6)%7));
      key=tmp.toISOString().slice(0,10);
    } else if(period==="monthly"){
      key=d.date.slice(0,7); // YYYY-MM
    } else if(period==="quarterly"){
      const q=Math.ceil((dt.getMonth()+1)/3);
      key=`${dt.getFullYear()}-Q${q}`;
    } else if(period==="half"){
      const h=dt.getMonth()<6?1:2;
      key=`${dt.getFullYear()}-H${h}`;
    } else if(period==="yearly"){
      key=`${dt.getFullYear()}`;
    }
    if(!map[key]) map[key]={key,total:0,A1:0,A2:0,A3:0,_date:key};
    map[key].total+=d.total; map[key].A1+=d.A1;
    map[key].A2+=d.A2; map[key].A3+=d.A3;
  });
  return Object.values(map).sort((a,b)=>a.key.localeCompare(b.key)).map(d=>({
    ...d,
    name:d.key.length===10?d.key.slice(5):d.key // show MM-DD for daily, full for others
  }));
}

function regionFullName(code){
  return REGION_NAMES[code]||code;
}

// ── READ FILE — reads critical cols by direct cell reference (100% reliable) ─────
// Parser CSV ringan (tanpa dependency tambahan) — support quoted field & koma di dalam quote
function parseCSVText(text) {
  text=text.replace(/^\uFEFF/,"").replace(/\r\n/g,"\n").replace(/\r/g,"\n");
  const rows=[];
  let row=[],field="",inQuotes=false;
  for(let i=0;i<text.length;i++){
    const c=text[i];
    if(inQuotes){
      if(c==='"'){
        if(text[i+1]==='"'){field+='"';i++;}
        else inQuotes=false;
      } else field+=c;
    } else {
      if(c==='"') inQuotes=true;
      else if(c===','){row.push(field);field="";}
      else if(c==='\n'){row.push(field);rows.push(row);row=[];field="";}
      else field+=c;
    }
  }
  if(field.length||row.length){row.push(field);rows.push(row);}
  if(!rows.length) return [];
  const headers=rows[0].map(h=>h.trim());
  return rows.slice(1).filter(r=>r.length>1||r[0]!=="").map(r=>{
    const obj={};
    headers.forEach((h,i)=>{obj[h]=r[i]!==undefined&&r[i]!==""?r[i]:null;});
    return obj;
  });
}

function readFileRows(wbOrBuf, sheetName=null) {
  // Terima workbook yang SUDAH di-parse (hindari re-parse buffer berkali-kali — берat utk file besar di mobile)
  const wb = (wbOrBuf && wbOrBuf.SheetNames) ? wbOrBuf : XLSX.read(wbOrBuf, {type:"array", cellDates:true, cellHTML:false});
  const targetSheet = sheetName||wb.SheetNames[0];
  const ws = wb.Sheets[targetSheet];
  if(!ws||!ws["!ref"]) throw new Error("Sheet kosong: "+targetSheet);

  const range = XLSX.utils.decode_range(ws["!ref"]);

  // Scan header row → build map: columnName → [list of col indices]
  const hdrMap = {};
  for(let c = range.s.c; c <= range.e.c; c++){
    const cell = ws[XLSX.utils.encode_cell({r: range.s.r, c})];
    if(cell && cell.v != null){
      const h = String(cell.v).trim();
      if(!hdrMap[h]) hdrMap[h] = [];
      hdrMap[h].push(c);
    }
  }

  // Identify critical column positions (case-insensitive)
  const vsCol   = findCol("Visit Status");
  const asLastC = hdrMap["Activity Status"]  ? hdrMap["Activity Status"].slice(-1)[0] : findCol("Activity Status");
  const durCol  = findCol("Duration Status");
  const disCol  = findCol("Distance Status");
  const locCol  = findCol("Location Status");
  console.log("[XL] cols → VS:"+vsCol+" AS:"+asLastC+" DUR:"+durCol+" DIS:"+disCol+" LOC:"+locCol);

  // Case-insensitive column finder — defined as function so it's hoisted
  function findCol(name){
    const lo = name.toLowerCase();
    if(hdrMap[name]) return hdrMap[name][0];
    for(const [k,v] of Object.entries(hdrMap)){
      if(k.toLowerCase()===lo) return v[0];
    }
    return -1;
  }

  // Get cell value: tries .v (raw value) then .w (formatted string)
  function getCellVal(row, col){
    if(col < 0) return null;
    const cell = ws[XLSX.utils.encode_cell({r: row, c: col})];
    if(!cell) return null;
    if(cell.v != null) return cell.v;
    if(cell.w != null) return cell.w;
    return null;
  }

  // Read all rows via sheet_to_json (handles dates, numbers etc.)
  const rawRows = XLSX.utils.sheet_to_json(ws, {defval: null});

  // Inject directly-read values for critical columns
  const rows = rawRows.map((row, idx) => {
    const rowNum = range.s.r + 1 + idx; // 0-indexed row in sheet (skip header)
    const vs  = getCellVal(rowNum, vsCol);
    const as1 = getCellVal(rowNum, asLastC);
    const dur = getCellVal(rowNum, durCol);
    const dis = getCellVal(rowNum, disCol);
    const loc = getCellVal(rowNum, locCol);
    return {
      ...row,
      "_VS":  vs  != null ? String(vs)  : null,
      "_AS1": as1 != null ? String(as1) : null,
      "_DUR": dur != null ? String(dur) : null,
      "_DIS": dis != null ? String(dis) : null,
      "_LOC": loc != null ? String(loc) : null,
    };
  });

  if(!rows.length) throw new Error("File kosong");
  // Apply validation computation for raw data (without green columns)
  const processedRows = rows.map(r => computeValidation(r));
  return {rows: processedRows};
}


// ── PROCESS ROWS ──────────────────────────────────────────────────────────────
function processRows(rows) {
  // total dihitung tanpa Consignment Visit (konsisten dengan KPI A1/A2/A3)
  const total=rows.filter(r=>String(r["Activity Type"]||"").trim()!=="Consignment Visit").length;
  const actC={"A1 - NORMAL":0,"A2 - ANOMALY":0,"A3 - INCOMPLETE":0};
  const visC={VALID:0,OBSERVE:0,INVESTIGATE:0,INCOMPLETE:0};
  const durC={NORMAL:0,SHORT:0,LONG:0};
  const disC={NEAR:0,MID:0,FAR:0,INCOMPLETE:0};
  const locC={MATCH:0,"NOT MATCH":0,INCOMPLETE:0};
  const inRangeC={YES:0,NO:0};
  const outMap={},canvMap={},dateMap={},visitMap={},vtMap={};
  let minDate=null,maxDate=null;
  // Reason aggregation for Investigate & Observe
  const reasonMap={investigate:{},observe:{}};
  const addReason=(bucket,label)=>{ reasonMap[bucket][label]=(reasonMap[bucket][label]||0)+1; };

  rows.forEach(r=>{
    // Use directly-read cell values (_VS, _AS1 etc.) — bypasses SheetJS duplicate key issues
    // _VS = Visit Status cell value (direct cell read, always correct)
    // _AS1 = last "Activity Status" col = A1/A2/A3 (fallback)
    let vs = String(r["_VS"]||r["Visit Status"]||"").toUpperCase();
    if(!["VALID","OBSERVE","INVESTIGATE","INCOMPLETE"].includes(vs)){
      // Fallback: use Activity Status last col (A1-NORMAL / A2-ANOMALY / A3-INCOMPLETE)
      const rawAS = String(r["_AS1"]||r["Activity Status"]||"");
      if(rawAS.startsWith("A1"))     vs="VALID";
      else if(rawAS.startsWith("A2"))vs="OBSERVE";
      else if(rawAS.startsWith("A3"))vs="INCOMPLETE";
      else {
        // Last resort: distance-based classification
        const hasIn  = r["Check-In Latitude"]  != null && r["Check-In Longitude"]  != null;
        const hasOut = r["Check-Out Latitude"] != null && r["Check-Out Longitude"] != null;
        if(!hasIn||!hasOut) vs="INCOMPLETE";
        else {
          const mx=Math.max(parseFloat(r["Distance Check In (Meter)"])||0, parseFloat(r["Distance Check Out (Meter)"])||0);
          vs = mx>5000?"INVESTIGATE":mx>500?"OBSERVE":"VALID";
        }
      }
    }
    const as1 = vs==="VALID"?"A1 - NORMAL":vs==="OBSERVE"||vs==="INVESTIGATE"?"A2 - ANOMALY":vs==="INCOMPLETE"?"A3 - INCOMPLETE":"";
    r["_CAS1"]=as1; r["_CVS"]=vs;
    const vt=String(r["Activity Type"]||"Unknown").trim();
    if(!vtMap[vt])vtMap[vt]={type:vt,total:0,A1:0,A2:0,A3:0};
    vtMap[vt].total++;
    if(as1==="A1 - NORMAL")vtMap[vt].A1++;
    else if(as1==="A2 - ANOMALY")vtMap[vt].A2++;
    else if(as1==="A3 - INCOMPLETE")vtMap[vt].A3++; // store for getCanvasserRows & outlet drill
    // Track anomaly reasons for Investigate & Observe
    if(vs==="INVESTIGATE"||vs==="OBSERVE"){
      const bucket=vs==="INVESTIGATE"?"investigate":"observe";
      const durSt2=String(r["_DUR"]||r["Duration Status"]||"").toUpperCase();
      const disSt2=String(r["_DIS"]||r["Distance Status"]||"").toUpperCase();
      const locSt2=String(r["_LOC"]||r["Location Status"]||"").toUpperCase();
      const inR2=String(r["In Range"]||"").toLowerCase();
      const dur2=parseFloat(r["Visit Duration (Menit)"]);
      const dIn2=parseFloat(r["Distance Check In (Meter)"]);
      const dOt2=parseFloat(r["Distance Check Out (Meter)"]);
      if(durSt2==="SHORT"||(dur2>0&&dur2<2)) addReason(bucket,"⏱ Durasi singkat");
      if(durSt2==="LONG"||(dur2>30)) addReason(bucket,"⏱ Durasi panjang");
      if(dIn2>5000||dOt2>5000) addReason(bucket,"🚨 Jarak sangat jauh");
      else if(dIn2>200||dOt2>200) addReason(bucket,"📍 Jarak jauh");
      if(locSt2==="NOT MATCH") addReason(bucket,"📌 Lokasi tidak match");
      if(inR2==="no"||inR2==="n") addReason(bucket,"🎯 Out of range");
    }
    // Consignment Visit dikecualikan dari KPI utama (A1/A2/A3 overview, pie chart, key insights),
    // tapi tetap dihitung di vtMap (Visit Type breakdown) di atas.
    if(vt==="Consignment Visit") return;
    // Duration Status — read from cell first, then compute from timestamps or raw duration
    let dur = String(r["_DUR"]!=null?r["_DUR"]:r["Duration Status"]!=null?r["Duration Status"]:"").trim().toUpperCase();
    if(!["NORMAL","SHORT","LONG"].includes(dur)){
      // Try Visit Duration (Menit) column
      let dm = parseFloat(r["Visit Duration (Menit)"]);
      // If null/NaN, compute from actual check-in/out timestamps (always stored as real values)
      if(isNaN(dm)){
        const tIn  = r["Actual Visit Time"]     ? new Date(r["Actual Visit Time"])     : null;
        const tOut = r["Actual Check-Out Time"] ? new Date(r["Actual Check-Out Time"]) : null;
        if(tIn && tOut && !isNaN(tIn) && !isNaN(tOut) && tOut > tIn){
          dm = (tOut - tIn) / 60000; // milliseconds → minutes
        }
      }
      dur = isNaN(dm) ? "" : dm < 2 ? "SHORT" : dm > 60 ? "LONG" : "NORMAL";
    }

    // Distance Status — read from cell, fallback compute from Distance Check In (Meter)
    let dis = String(r["_DIS"]!=null?r["_DIS"]:r["Distance Status"]!=null?r["Distance Status"]:"").trim().toUpperCase();
    if(!["NEAR","MID","FAR","INCOMPLETE"].includes(dis)){
      const hasIn=r["Check-In Latitude"]!=null&&r["Check-In Longitude"]!=null;
      if(!hasIn) dis="INCOMPLETE";
      else{
        const dm=parseFloat(r["Distance Check In (Meter)"])||0;
        dis=dm<=100?"NEAR":dm<=1000?"MID":"FAR";
      }
    }

    // Location Status — read from cell, fallback compute from coordinates vs RO
    let loc = String(r["_LOC"]!=null?r["_LOC"]:r["Location Status"]!=null?r["Location Status"]:"").trim();
    if(!["MATCH","NOT MATCH","INCOMPLETE"].includes(loc.toUpperCase())){
      const hasIn=r["Check-In Latitude"]!=null&&r["Check-In Longitude"]!=null;
      const hasRO=r["RO Latitude"]!=null&&r["RO Longitude"]!=null;
      if(!hasIn) loc="INCOMPLETE";
      else if(!hasRO) loc="MATCH"; // can't verify, assume match
      else{
        const dlat=Math.abs(parseFloat(r["Check-In Latitude"])-parseFloat(r["RO Latitude"]));
        const dlng=Math.abs(parseFloat(r["Check-In Longitude"])-parseFloat(r["RO Longitude"]));
        // ~0.005 deg ≈ 500m
        loc=(dlat<0.005&&dlng<0.005)?"MATCH":"NOT MATCH";
      }
    }
    loc=loc.toUpperCase();
    const ot  = r["Outlet Type"]||"Unknown";
    const nm  = r["Canvasser"]||"Unknown";
    const cid = String(r["Canvasser ID"]||nm).trim();
    const cl  = r["Cluster"]||"Unknown";
    const rgn = getRegionCode(cl);
    const dt  = extractDate(r["Planned Visit Date"]);
    const durM= parseFloat(r["Visit Duration (Menit)"]);
    const disM= parseFloat(r["Distance Check In (Meter)"]);

    if(actC[as1]!==undefined) actC[as1]++;
    if(visC[vs] !==undefined) visC[vs]++;
    if(durC[dur]!==undefined) durC[dur]++;
    if(disC[dis]!==undefined) disC[dis]++;
    if(locC[loc]!==undefined) locC[loc]++;
    // In Range
    const inR=String(r["In Range"]||"").trim().toUpperCase();
    const inRKey=inR==="YES"||inR==="Y"||inR==="1"||inR==="TRUE"?"YES":"NO";
    if(r["In Range"]!=null) inRangeC[inRKey]++;

    if(!outMap[ot]) outMap[ot]={type:ot,total:0,A1:0,A2:0,A3:0,VALID:0,OBSERVE:0,INVESTIGATE:0,INCOMPLETE:0};
    outMap[ot].total++;
    if(as1==="A1 - NORMAL")    outMap[ot].A1++;
    if(as1==="A2 - ANOMALY")   outMap[ot].A2++;
    if(as1==="A3 - INCOMPLETE")outMap[ot].A3++;
    if(visC[vs]!==undefined)   outMap[ot][vs]=(outMap[ot][vs]||0)+1;
    // Census tracking
    const censusVal = String(r["RO Census"]||"").trim().toUpperCase();
    const isCensus = censusVal==="Y"||censusVal==="YES"||censusVal==="1"||censusVal==="TRUE";
    const ck = isCensus?"Census":"Non-Census";
    if(!outMap["__"+ck]) outMap["__"+ck]={type:ck,_isCensus:true,total:0,A1:0,A2:0,A3:0};
    outMap["__"+ck].total++;
    if(as1==="A1 - NORMAL")    outMap["__"+ck].A1++;
    if(as1==="A2 - ANOMALY")   outMap["__"+ck].A2++;
    if(as1==="A3 - INCOMPLETE")outMap["__"+ck].A3++;

    if(!canvMap[cid]) canvMap[cid]={id:cid,name:nm,cluster:cl,region:rgn,total:0,A1:0,A2:0,A3:0,VALID:0,OBSERVE:0,INVESTIGATE:0,INCOMPLETE:0,durSum:0,durCnt:0,disSum:0,disCnt:0,
      DUR_NORMAL:0,DUR_SHORT:0,DUR_LONG:0,DIS_NEAR:0,DIS_MID:0,DIS_FAR:0,DIS_INC:0,LOC_MATCH:0,LOC_NOTMATCH:0,LOC_INC:0,IR_YES:0,IR_NO:0};
    canvMap[cid].total++;
    if(as1==="A1 - NORMAL")    canvMap[cid].A1++;
    if(as1==="A2 - ANOMALY")   canvMap[cid].A2++;
    if(as1==="A3 - INCOMPLETE")canvMap[cid].A3++;
    if(visC[vs]!==undefined)   canvMap[cid][vs]=(canvMap[cid][vs]||0)+1;
    if(!isNaN(durM)){canvMap[cid].durSum+=durM;canvMap[cid].durCnt++;}
    if(!isNaN(disM)){canvMap[cid].disSum+=disM;canvMap[cid].disCnt++;}
    // In Range per canvasser
    if(r["In Range"]!=null){if(inRKey==="YES")canvMap[cid].IR_YES++;else canvMap[cid].IR_NO++;}
    // Duration/Distance/Location per canvasser
    if(dur==="NORMAL")    canvMap[cid].DUR_NORMAL++;
    else if(dur==="SHORT")canvMap[cid].DUR_SHORT++;
    else if(dur==="LONG") canvMap[cid].DUR_LONG++;
    if(dis==="NEAR")      canvMap[cid].DIS_NEAR++;
    else if(dis==="MID")  canvMap[cid].DIS_MID++;
    else if(dis==="FAR")  canvMap[cid].DIS_FAR++;
    else if(dis==="INCOMPLETE") canvMap[cid].DIS_INC++;
    if(loc==="MATCH")         canvMap[cid].LOC_MATCH++;
    else if(loc==="NOT MATCH")canvMap[cid].LOC_NOTMATCH++;
    else if(loc==="INCOMPLETE")canvMap[cid].LOC_INC++;

    if(dt){
      if(!minDate||dt<minDate) minDate=dt;
      if(!maxDate||dt>maxDate) maxDate=dt;
      if(!dateMap[dt]) dateMap[dt]={date:dt,total:0,A1:0,A2:0,A3:0};
      dateMap[dt].total++;
      if(as1==="A1 - NORMAL")    dateMap[dt].A1++;
      if(as1==="A2 - ANOMALY")   dateMap[dt].A2++;
      if(as1==="A3 - INCOMPLETE")dateMap[dt].A3++;
    }

    const outId=r["Outlet ID"]!=null?String(r["Outlet ID"]):null;
    if(outId&&dt){
      const key=`${outId}|${dt}`;
      if(!visitMap[key]) visitMap[key]={outlet:String(r["Outlet"]||outId),outletId:outId,cluster:String(cl),date:dt,visits:[],statuses:[]};
      visitMap[key].visits.push(String(nm));
      visitMap[key].statuses.push(String(as1||"–"));
    }
  });

  const canvassers=Object.values(canvMap).map(c=>({...c,
    avgDur:c.durCnt?+(c.durSum/c.durCnt).toFixed(1):null,
    avgDis:c.disCnt?+(c.disSum/c.disCnt).toFixed(1):null,
    a1p:pct(c.A1,c.total),a2p:pct(c.A2,c.total),a3p:pct(c.A3,c.total),invP:pct(c.INVESTIGATE,c.total),
  }));

  const censusData = Object.values(outMap).filter(d=>d._isCensus).sort((a,b)=>b.total-a.total);
  return {
    total,actC,visC,durC,disC,locC,inRangeC,
    visitTypeData:Object.values(vtMap).sort((a,b)=>b.total-a.total),
    outletData:Object.values(outMap).filter(d=>!d._isCensus).sort((a,b)=>b.total-a.total),
    censusData,
    canvassers,
    dateRange:{min:minDate,max:maxDate},
    reasonMap,
    trend:Object.values(dateMap).sort((a,b)=>a.date.localeCompare(b.date)),
    duplicates:Object.values(visitMap).filter(v=>Array.isArray(v.visits)&&v.visits.length>1).sort((a,b)=>b.visits.length-a.visits.length),
  };
}

// ── AGGREGATE helper ──────────────────────────────────────────────────────────
function aggregateList(dataList) {
  const sumC=(key)=>{const m={};dataList.forEach(r=>Object.entries(r[key]||{}).forEach(([k,v])=>{m[k]=(m[k]||0)+v;}));return m;};
  const mergeArr=(key,gk)=>{
    const m={};
    dataList.forEach(r=>(r[key]||[]).forEach(item=>{
      const k=item[gk];
      if(!m[k])m[k]={...item,total:0,A1:0,A2:0,A3:0,VALID:0,OBSERVE:0,INVESTIGATE:0,INCOMPLETE:0};
      ["total","A1","A2","A3","VALID","OBSERVE","INVESTIGATE","INCOMPLETE"].forEach(f=>{m[k][f]=(m[k][f]||0)+(item[f]||0);});
    }));
    return Object.values(m).sort((a,b)=>b.total-a.total);
  };
  const tMap={};
  dataList.forEach(r=>(r.trend||[]).forEach(d=>{
    if(!tMap[d.date])tMap[d.date]={date:d.date,total:0,A1:0,A2:0,A3:0};
    ["total","A1","A2","A3"].forEach(f=>{tMap[d.date][f]=(tMap[d.date][f]||0)+(d[f]||0);});
  }));
  const CANV_KEYS=["total","A1","A2","A3","VALID","OBSERVE","INVESTIGATE","INCOMPLETE","durSum","durCnt","disSum","disCnt","DUR_NORMAL","DUR_SHORT","DUR_LONG","DIS_NEAR","DIS_MID","DIS_FAR","DIS_INC","LOC_MATCH","LOC_NOTMATCH","LOC_INC","IR_YES","IR_NO"];
  const cMap={};
  dataList.forEach(r=>(r.canvassers||[]).forEach(c=>{
    const key=c.id||c.name;
    if(!cMap[key])cMap[key]={...c};
    else CANV_KEYS.forEach(k=>{cMap[key][k]=(cMap[key][k]||0)+(c[k]||0);});
  }));
  const canvassers=Object.values(cMap).map(c=>({...c,
    avgDur:c.durCnt?+(c.durSum/c.durCnt).toFixed(1):null,
    avgDis:c.disCnt?+(c.disSum/c.disCnt).toFixed(1):null,
    a1p:pct(c.A1,c.total),a2p:pct(c.A2,c.total),a3p:pct(c.A3,c.total),invP:pct(c.INVESTIGATE,c.total),
  }));
  return {
    total:dataList.reduce((s,r)=>s+(r.total||0),0),
    actC:sumC("actC"),visC:sumC("visC"),durC:sumC("durC"),disC:sumC("disC"),locC:sumC("locC"),inRangeC:sumC("inRangeC"),
    visitTypeData:mergeArr("visitTypeData","type"),
    outletData:mergeArr("outletData","type"),
    censusData:mergeArr("censusData","type"),
    canvassers,
    reasonMap:(()=>{
      const merged={investigate:{},observe:{}};
      dataList.forEach(d=>{
        ["investigate","observe"].forEach(b=>{
          Object.entries(d.reasonMap?.[b]||{}).forEach(([k,v])=>{ merged[b][k]=(merged[b][k]||0)+v; });
        });
      });
      return merged;
    })(),
    dateRange:(()=>{
      const mins=dataList.map(r=>r.dateRange?.min).filter(Boolean).sort();
      const maxs=dataList.map(r=>r.dateRange?.max).filter(Boolean).sort();
      return mins.length?{min:mins[0],max:maxs[maxs.length-1]}:null;
    })(),
    trend:Object.values(tMap).sort((a,b)=>a.date.localeCompare(b.date)),
    duplicates:dataList.flatMap(r=>(r.duplicates||[])).sort((a,b)=>b.visits.length-a.visits.length),
  };
}

// ── TOOLTIP ───────────────────────────────────────────────────────────────────
function Tip({active,payload,label,t}){
  if(!active||!payload?.length)return null;
  return(
    <div style={{background:t.card,border:`1px solid ${t.border}`,borderRadius:10,padding:"10px 14px",fontSize:12,color:t.text,boxShadow:"0 8px 32px rgba(0,0,0,0.4)"}}>
      <p style={{fontWeight:700,marginBottom:6,color:P.accent}}>{label}</p>
      {payload.map((p,i)=><p key={i} style={{color:p.color,margin:"2px 0"}}>{p.name}: <b>{typeof p.value==="number"?p.value.toLocaleString():p.value}</b></p>)}
    </div>
  );
}
const Bar3=({A1,A2,A3,total})=>(
  <div style={{display:"flex",height:6,borderRadius:4,overflow:"hidden",width:"100%"}}>
    {[{v:A1,c:P.a1},{v:A2,c:P.a2},{v:A3,c:P.a3}].map((s,i)=>s.v>0&&<div key={i} style={{width:pct(s.v,total)+"%",background:s.c}}/>)}
  </div>
);

// ── PAGINATION ───────────────────────────────────────────────────────────────
function Pagination({page,setPage,total,pageSize,t}){
  const tp=Math.ceil(total/pageSize);
  if(tp<=1||!total) return null;
  const s=page*pageSize+1,e=Math.min((page+1)*pageSize,total);
  const ps=Array.from({length:tp},(_,i)=>i).filter(i=>Math.abs(i-page)<=2);
  const btn=(a,d)=>({background:a?P.accent:d?"transparent":t.cardAlt,color:a?"#fff":d?t.muted:t.text,border:`1px solid ${a?"transparent":t.border}`,borderRadius:6,padding:"5px 10px",fontSize:11,fontWeight:700,cursor:d?"default":"pointer",opacity:d?0.4:1});
  return(
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 14px",borderTop:`1px solid ${t.border}`,background:t.cardAlt,flexWrap:"wrap",gap:6,flexShrink:0}}>
      <span style={{fontSize:11,color:t.muted}}>{s}–{e} dari {total}</span>
      <div style={{display:"flex",gap:4}}>
        <button onClick={()=>setPage(0)} disabled={page===0} style={btn(false,page===0)}>«</button>
        <button onClick={()=>setPage(p=>Math.max(0,p-1))} disabled={page===0} style={btn(false,page===0)}>‹</button>
        {ps.map(i=><button key={i} onClick={()=>setPage(i)} style={btn(i===page,false)}>{i+1}</button>)}
        <button onClick={()=>setPage(p=>Math.min(tp-1,p+1))} disabled={page>=tp-1} style={btn(false,page>=tp-1)}>›</button>
        <button onClick={()=>setPage(tp-1)} disabled={page>=tp-1} style={btn(false,page>=tp-1)}>»</button>
      </div>
    </div>
  );
}

// ── VISIT TYPE DRILL PANEL ───────────────────────────────────────────────────
function VisitTypeDrillPanel({drill,onClose,t,onCanvasserClick}){
  const [pg,setPg]=useState(0);
  const [sBy,setSBy]=useState("total");
  const [sDir,setSDir]=useState("desc");
  const [srch,setSrch]=useState("");
  const PG=10;
  const COLOR="#06b6d4";
  useEffect(()=>{setPg(0);setSrch("");},[drill?.visitType,drill?.statusFilter]);
  if(!drill) return null;
  const {visitType,label,rows}=drill;
  const SC=label==="A1"?P.a1:label==="A2"?P.a2:label==="A3"?P.a3:COLOR;
  const filt=[...rows]
    .filter(r=>srch?r.name.toLowerCase().includes(srch.toLowerCase())||(r.cluster||"").toLowerCase().includes(srch.toLowerCase()):true)
    .sort((a,b)=>sDir==="desc"?b[sBy]-a[sBy]:a[sBy]-b[sBy]);
  const list=filt.slice(pg*PG,(pg+1)*PG);
  const sBtn=(lbl,sk)=>(
    <button onClick={()=>{if(sBy===sk)setSDir(d=>d==="desc"?"asc":"desc");else{setSBy(sk);setSDir("desc");setPg(0);}}}
      style={{background:sBy===sk?SC:t.cardAlt,color:sBy===sk?"#fff":t.muted,border:"1px solid "+t.border,borderRadius:6,padding:"3px 10px",fontSize:10,fontWeight:700,cursor:"pointer"}}>
      {lbl}{sBy===sk?(sDir==="desc"?" ↓":" ↑"):""}
    </button>
  );
  return(
    <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:1050,display:"flex",alignItems:"flex-end",background:"rgba(0,0,0,0.65)",backdropFilter:"blur(4px)"}} onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{width:"100%",maxHeight:"85vh",background:t.card,borderRadius:"20px 20px 0 0",border:`1px solid ${t.border}`,overflow:"hidden",display:"flex",flexDirection:"column",boxShadow:"0 -8px 40px rgba(0,0,0,0.5)",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
        <div style={{padding:"14px 18px 10px",borderBottom:`1px solid ${t.border}`,display:"flex",alignItems:"center",gap:12,flexShrink:0}}>
          <div style={{flex:1}}>
            <div style={{fontWeight:800,fontSize:15,color:t.text}}>{visitType==="Regular Visit"?"🚗":visitType==="Ad-Hoc Visit"?"⚡":"📦"} {visitType}</div>
            <div style={{fontSize:11,color:t.muted,marginTop:2,display:"flex",gap:8,alignItems:"center"}}>
              {label!=="Semua"&&<span style={{background:SC+"20",color:SC,padding:"1px 8px",borderRadius:999,fontSize:10,fontWeight:700}}>{label==="A1"?"A1 - Normal":label==="A2"?"A2 - Anomaly":"A3 - Incomplete"}</span>}
              <span>{rows.length} canvasser · {rows.reduce((s,r)=>s+r.total,0).toLocaleString()} aktivitas</span>
            </div>
          </div>
          <button onClick={onClose} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:13,fontWeight:700}}>✕</button>
        </div>
        <div style={{padding:"8px 18px",borderBottom:`1px solid ${t.border}`,flexShrink:0}}>
          <input placeholder="🔍 Cari canvasser / cluster..." value={srch} onChange={e=>{setSrch(e.target.value);setPg(0);}}
            style={{width:"100%",background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"7px 12px",fontSize:12,outline:"none",boxSizing:"border-box"}}/>
          <div style={{display:"flex",gap:6,marginTop:6,flexWrap:"wrap",alignItems:"center"}}>
            <span style={{fontSize:10,color:t.muted,fontWeight:600}}>Sort:</span>
            {sBtn("Total","total")}{sBtn("A1","A1")}{sBtn("A2","A2")}{sBtn("A3","A3")}{sBtn("Outlet","outletCount")}
          </div>
        </div>
        <div style={{overflowY:"auto",flex:1,scrollbarWidth:"none",msOverflowStyle:"none"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
            <thead style={{position:"sticky",top:0,background:t.card,zIndex:1}}>
              <tr style={{background:t.cardAlt}}>
                {["#","Canvasser","Region","Cluster","Total","A1","A2","A3","Outlet (jml)","Outlet"].map(h=>(
                  <th key={h} style={{padding:"9px 10px",textAlign:"left",fontSize:11,fontWeight:700,color:t.muted,whiteSpace:"nowrap",borderBottom:`1px solid ${t.border}`}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {list.map((r,i)=>(
                <tr key={r.id||i} style={{borderBottom:`1px solid ${t.border}`,background:i%2===0?"transparent":t.rowAlt}}>
                  <td style={{padding:"7px 10px",color:t.muted,fontSize:10}}>{pg*PG+i+1}</td>
                  <td style={{padding:"7px 10px",fontWeight:600,color:t.text,whiteSpace:"nowrap"}}>{r.name}</td>
                  <td style={{padding:"7px 10px"}}><span style={{background:P.accent+"20",color:P.accent,padding:"1px 7px",borderRadius:6,fontSize:10,fontWeight:700}}>{r.region||"–"}</span></td>
                  <td style={{padding:"7px 10px",color:t.muted,fontSize:11}}>{r.cluster||"–"}</td>
                  <td style={{padding:"7px 10px",fontWeight:700,color:SC}}>{r.total}</td>
                  <td style={{padding:"7px 10px"}}>
                    {r.A1>0?<span onClick={()=>onCanvasserClick&&onCanvasserClick(r,"A1")} style={{color:P.a1,fontWeight:700,cursor:"pointer",borderBottom:"1px dotted "+P.a1}}>{r.A1}</span>:<span style={{color:t.muted}}>0</span>}
                  </td>
                  <td style={{padding:"7px 10px"}}>
                    {r.A2>0?<span onClick={()=>onCanvasserClick&&onCanvasserClick(r,"A2")} style={{color:P.a2,fontWeight:700,cursor:"pointer",borderBottom:"1px dotted "+P.a2}}>{r.A2}</span>:<span style={{color:t.muted}}>0</span>}
                  </td>
                  <td style={{padding:"7px 10px"}}>
                    {r.A3>0?<span onClick={()=>onCanvasserClick&&onCanvasserClick(r,"A3")} style={{color:P.a3,fontWeight:700,cursor:"pointer",borderBottom:"1px dotted "+P.a3}}>{r.A3}</span>:<span style={{color:t.muted}}>0</span>}
                  </td>
                  <td style={{padding:"7px 10px",color:t.muted,fontWeight:600}}>{r.outletCount}</td>
                  <td style={{padding:"7px 10px",color:t.muted,fontSize:10,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.outletList||"–"}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <Pagination page={pg} setPage={setPg} total={filt.length} pageSize={PG} t={t}/>
        </div>
      </div>
    </div>
  );
}

// ── OUTLET ACTIVITY PANEL ─────────────────────────────────────────────────────
function OutletActivityPanel({detail,onClose,t}){
  const [pg,setPg]=useState(0);
  const PG=10;
  if(!detail) return null;
  const {outletId,outletName,status,rows}=detail;
  const SC=status==="A1"?P.a1:status==="A2"?P.a2:status==="A3"?P.a3:"#06b6d4";
  const LABEL=status==="A1"?"A1 - Normal":status==="A2"?"A2 - Anomaly":status==="A3"?"A3 - Incomplete":"Semua Aktivitas";
  const fmtDate=v=>{if(!v)return"–";const d=new Date(v);return isNaN(d)?"–":d.toLocaleDateString("id-ID",{day:"2-digit",month:"short",year:"2-digit"});};
  const fmtDist=v=>{const n=parseFloat(v);return isNaN(n)?"–":n>=1000?(n/1000).toFixed(1)+"km":n.toFixed(0)+"m";};
  const fmtDur=v=>{const n=parseFloat(v);if(isNaN(n))return"–";if(n>=60)return(n/60).toFixed(1)+"j";if(n>=1)return n.toFixed(1)+"mnt";return Math.round(n*60)+"det";};
  const vsColor=v=>{const u=String(v||"").toUpperCase();return u==="VALID"?P.valid:u==="OBSERVE"?P.observe:u==="INVESTIGATE"?P.investigate:P.incomplete;};
  const getReason=r=>{
    const vs=String(r["_VS"]||r["Visit Status"]||"").toUpperCase();
    const dIn=parseFloat(r["Distance Check In (Meter)"])||0;
    const dOt=parseFloat(r["Distance Check Out (Meter)"])||0;
    const dur=parseFloat(r["Visit Duration (Menit)"]);
    const durSt=String(r["_DUR"]||r["Duration Status"]||"").toUpperCase();
    const loc=String(r["_LOC"]||r["Location Status"]||"").toUpperCase();
    const inR=String(r["In Range"]||"").toLowerCase();
    if(vs==="INCOMPLETE")return"❌ Checkout tidak ada";
    const f=[];
    if(durSt==="SHORT"||(!isNaN(dur)&&dur>0&&dur<2))f.push("⏱ Durasi singkat ("+fmtDur(dur)+")");
    else if(durSt==="LONG"||(!isNaN(dur)&&dur>60))f.push("⏱ Durasi panjang ("+fmtDur(dur)+")");
    if(dIn>5000)f.push("🚨 Check-in sangat jauh ("+fmtDist(dIn)+")");
    else if(dOt>5000)f.push("🚨 Check-out sangat jauh ("+fmtDist(dOt)+")");
    else if(dIn>500)f.push("📍 Check-in jauh ("+fmtDist(dIn)+")");
    else if(dOt>500)f.push("📍 Check-out jauh ("+fmtDist(dOt)+")");
    if(loc==="NOT MATCH")f.push("📌 Lokasi tidak match");
    if(inR==="no"||inR==="n")f.push("🎯 Out of range");
    return f.length>0?f.join(" · "):"✅ Normal";
  };
  const list=rows.slice(pg*PG,(pg+1)*PG);
  return(
    <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:1200,display:"flex",alignItems:"flex-end",background:"rgba(0,0,0,0.75)",backdropFilter:"blur(4px)"}} onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{width:"100%",maxHeight:"88vh",background:t.card,borderRadius:"20px 20px 0 0",border:`1px solid ${t.border}`,overflow:"hidden",display:"flex",flexDirection:"column",boxShadow:"0 -8px 40px rgba(0,0,0,0.6)",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
        <div style={{padding:"14px 18px 10px",borderBottom:`1px solid ${t.border}`,display:"flex",alignItems:"center",gap:12,flexShrink:0}}>
          <div style={{flex:1}}>
            <div style={{fontWeight:800,fontSize:15,color:t.text}}>🏪 {outletName}</div>
            <div style={{fontSize:11,color:t.muted,marginTop:3,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
              <span style={{color:t.muted,fontSize:10}}>ID: {outletId}</span>
              <span style={{background:SC+"20",color:SC,padding:"2px 10px",borderRadius:999,fontSize:10,fontWeight:700}}>{LABEL}</span>
              <span style={{color:t.text,fontWeight:800,fontSize:14}}>· {rows.length.toLocaleString()} aktivitas</span>
            </div>
          </div>
          <button onClick={onClose} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:13,fontWeight:700}}>✕</button>
        </div>
        <div style={{overflowY:"auto",flex:1,scrollbarWidth:"none",msOverflowStyle:"none"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
            <thead style={{position:"sticky",top:0,background:t.card,zIndex:1}}>
              <tr style={{background:t.cardAlt}}>
                {["#","Tanggal","Canvasser","Status","In Range","Jarak*","Durasi","Alasan"].map(h=>(
                  <th key={h} style={{padding:"9px 12px",textAlign:"left",fontSize:11,fontWeight:700,color:t.muted,whiteSpace:"nowrap",borderBottom:`1px solid ${t.border}`}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {list.map((r,i)=>{
                const vs=String(r["_VS"]||r["Visit Status"]||"").toUpperCase();
                const vc=vsColor(vs);
                const dist=parseFloat(r["Distance Check In (Meter)"])||0;
                const dur=parseFloat(r["Visit Duration (Menit)"]);
                const inR=String(r["In Range"]||"").toLowerCase();
                const isIn=inR==="yes"||inR==="y"||inR==="1";
                return(
                <tr key={i} style={{borderBottom:`1px solid ${t.border}`,background:i%2===0?"transparent":t.rowAlt}}>
                  <td style={{padding:"7px 10px",color:t.muted,fontSize:10}}>{pg*PG+i+1}</td>
                  <td style={{padding:"7px 10px",color:t.text,whiteSpace:"nowrap"}}>{fmtDate(r["Planned Visit Date"])}</td>
                  <td style={{padding:"7px 10px",fontWeight:600,color:t.text,whiteSpace:"nowrap"}}>{r["Canvasser"]||"–"}</td>
                  <td style={{padding:"7px 10px"}}><span style={{background:vc+"20",color:vc,padding:"2px 8px",borderRadius:999,fontSize:10,fontWeight:700}}>{vs||"–"}</span></td>
                  <td style={{padding:"7px 10px"}}>
                    {r["In Range"]!=null
                      ?<span style={{background:isIn?P.a1+"22":P.investigate+"22",color:isIn?P.a1:P.investigate,padding:"2px 8px",borderRadius:999,fontSize:10,fontWeight:700}}>{isIn?"✓ In":"✗ Out"}</span>
                      :<span style={{color:t.muted}}>–</span>}
                  </td>
                  <td style={{padding:"7px 10px",color:dist>500?P.investigate:dist>100?P.observe:t.muted,fontWeight:dist>500?700:400}}>{fmtDist(r["Distance Check In (Meter)"])}</td>
                  <td style={{padding:"7px 10px",color:!isNaN(dur)&&dur>0&&dur<2?P.short:t.muted}}>
                    {fmtDur(r["Visit Duration (Menit)"])}{!isNaN(dur)&&dur>0&&dur<1&&<span style={{fontSize:9,color:P.investigate,marginLeft:3}}>⚡</span>}
                  </td>
                  <td style={{padding:"6px 8px",textAlign:"center"}}>{["AVA Poster XL/AXIS Comply?","AVA Poster Smartfren Comply?","AVA SP XL Comply?","AVA SP AXIS Comply?","AVA SP Smartfren Comply?","AVA Voucher XL Comply?","AVA Voucher AXIS Comply?","AVA Voucher Smartfren Comply?"].filter(k=>String(r[k]||"").toLowerCase()==="yes").length===8?<span style={{background:"#22c55e22",color:"#22c55e",padding:"2px 6px",borderRadius:999,fontSize:10,fontWeight:700}}>8/8</span>:["AVA Poster XL/AXIS Comply?","AVA Poster Smartfren Comply?","AVA SP XL Comply?","AVA SP AXIS Comply?","AVA SP Smartfren Comply?","AVA Voucher XL Comply?","AVA Voucher AXIS Comply?","AVA Voucher Smartfren Comply?"].filter(k=>String(r[k]||"").toLowerCase()==="yes").length>=4?<span style={{background:"#f59e0b22",color:"#f59e0b",padding:"2px 6px",borderRadius:999,fontSize:10,fontWeight:700}}>{["AVA Poster XL/AXIS Comply?","AVA Poster Smartfren Comply?","AVA SP XL Comply?","AVA SP AXIS Comply?","AVA SP Smartfren Comply?","AVA Voucher XL Comply?","AVA Voucher AXIS Comply?","AVA Voucher Smartfren Comply?"].filter(k=>String(r[k]||"").toLowerCase()==="yes").length}/8</span>:<span style={{background:"#ef444422",color:"#ef4444",padding:"2px 6px",borderRadius:999,fontSize:10,fontWeight:700}}>{["AVA Poster XL/AXIS Comply?","AVA Poster Smartfren Comply?","AVA SP XL Comply?","AVA SP AXIS Comply?","AVA SP Smartfren Comply?","AVA Voucher XL Comply?","AVA Voucher AXIS Comply?","AVA Voucher Smartfren Comply?"].filter(k=>String(r[k]||"").toLowerCase()==="yes").length}/8</span>}</td>
                  <td style={{padding:"7px 10px",fontSize:11,color:t.muted}}>{getReason(r)}</td>
                </tr>
              );})}
            </tbody>
          </table>
          <Pagination page={pg} setPage={setPg} total={rows.length} pageSize={PG} t={t}/>
          <div style={{padding:"8px 16px",fontSize:10,color:t.muted,background:t.cardAlt,borderTop:`1px solid ${t.border}`}}>
            * Jarak = selisih GPS canvasser vs koordinat outlet terdaftar
          </div>
        </div>
      </div>
    </div>
  );
}

// ── OUTLET DRILL PANEL ───────────────────────────────────────────────────────
function OutletDrillPanel({drill,onClose,t,onDrill}){
  const [pg,setPg]=useState(0);
  const [sBy,setSBy]=useState("total");
  const [sDir,setSDir]=useState("desc");
  const [srch,setSrch]=useState("");
  const PG=10;
  const COLOR="#06b6d4";
  // Reset page + search when outlet type changes (useEffect = correct hooks pattern)
  useEffect(()=>{setPg(0);setSrch("");setSBy("total");setSDir("desc");},[drill?.outletType]);
  if(!drill) return null;
  const filt=[...drill.rows]
    .filter(r=>srch?r.name.toLowerCase().includes(srch.toLowerCase())||(r.cluster||"").toLowerCase().includes(srch.toLowerCase()):true)
    .sort((a,b)=>sDir==="desc"?b[sBy]-a[sBy]:a[sBy]-b[sBy]);
  const list=filt.slice(pg*PG,(pg+1)*PG);
  const total=drill.rows.reduce((s,r)=>s+r.total,0);
  const sortBtn=(label,sk)=>(
    <button onClick={()=>{if(sBy===sk)setSDir(d=>d==="desc"?"asc":"desc");else{setSBy(sk);setSDir("desc");setPg(0);}}}
      style={{background:sBy===sk?COLOR:t.cardAlt,color:sBy===sk?"#fff":t.muted,border:"1px solid "+t.border,borderRadius:6,padding:"3px 10px",fontSize:10,fontWeight:700,cursor:"pointer"}}>
      {label}{sBy===sk?(sDir==="desc"?" ↓":" ↑"):""}
    </button>
  );
  return(
    <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:1000,display:"flex",alignItems:"flex-end",background:"rgba(0,0,0,0.65)",backdropFilter:"blur(4px)"}} onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{width:"100%",maxHeight:"85vh",background:t.card,borderRadius:"20px 20px 0 0",border:`1px solid ${t.border}`,overflow:"hidden",display:"flex",flexDirection:"column",boxShadow:"0 -8px 40px rgba(0,0,0,0.5)",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
        <div style={{padding:"14px 18px 10px",borderBottom:`1px solid ${t.border}`,display:"flex",alignItems:"center",gap:12,flexShrink:0}}>
          <div style={{width:12,height:12,borderRadius:3,background:COLOR,flexShrink:0}}/>
          <div style={{flex:1}}>
            <div style={{fontWeight:800,fontSize:15,color:t.text}}>🏪 {drill.outletType}</div>
            <div style={{fontSize:11,color:t.muted,marginTop:1}}>{drill.rows.length} outlet · {total.toLocaleString()} kunjungan</div>
          </div>
          <button onClick={onClose} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:13,fontWeight:700}}>✕</button>
        </div>
        <div style={{padding:"8px 18px",borderBottom:`1px solid ${t.border}`,flexShrink:0}}>
          <input placeholder="🔍 Cari outlet / ID..." value={srch} onChange={e=>{setSrch(e.target.value);setPg(0);}}
            style={{width:"100%",background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"7px 12px",fontSize:12,outline:"none",boxSizing:"border-box"}}/>
          <div style={{display:"flex",gap:6,marginTop:6,flexWrap:"wrap",alignItems:"center"}}>
            <span style={{fontSize:10,color:t.muted,fontWeight:600}}>Sort:</span>
            {sortBtn("Total","total")}
            {sortBtn("A1","A1")}
            {sortBtn("A2","A2")}
            {sortBtn("A3","A3")}
            {sortBtn("Census","census")}
          </div>
        </div>
        <div style={{overflowY:"auto",flex:1,scrollbarWidth:"none",msOverflowStyle:"none"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
            <thead style={{position:"sticky",top:0,background:t.card,zIndex:1}}>
              <tr style={{background:t.cardAlt}}>
                {["#","Outlet ID","Outlet","Cluster","Total","A1","A2","A3","Census","Non-Census","Canvasser"].map(h=>(
                  <th key={h} style={{padding:"9px 10px",textAlign:"left",fontSize:11,fontWeight:700,color:t.muted,whiteSpace:"nowrap",borderBottom:`1px solid ${t.border}`}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {list.map((r,i)=>(
                <tr key={r.id||i} style={{borderBottom:`1px solid ${t.border}`,background:i%2===0?"transparent":t.rowAlt}}>
                  <td style={{padding:"7px 10px",color:t.muted,fontSize:10}}>{pg*PG+i+1}</td>
                  <td style={{padding:"7px 10px",color:t.muted,fontSize:10,whiteSpace:"nowrap"}}>{r.id||"–"}</td>
                  <td style={{padding:"7px 10px",fontWeight:600,color:t.text,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.name}</td>
                  <td style={{padding:"7px 10px",color:t.muted,fontSize:11,whiteSpace:"nowrap"}}>{r.cluster||"–"}</td>
                  <td style={{padding:"7px 10px"}}>
                    <span onClick={()=>onDrill&&onDrill(r,"ALL")} style={{color:COLOR,fontWeight:700,cursor:"pointer",borderBottom:"1px dotted "+COLOR}}>{r.total}</span>
                  </td>
                  <td style={{padding:"7px 10px"}}>
                    {r.A1>0
                      ?<span onClick={()=>onDrill&&onDrill(r,"A1")} style={{color:P.a1,fontWeight:700,cursor:"pointer",borderBottom:"1px dotted "+P.a1}}>{r.A1}</span>
                      :<span style={{color:t.muted}}>0</span>}
                  </td>
                  <td style={{padding:"7px 10px"}}>
                    {r.A2>0
                      ?<span onClick={()=>onDrill&&onDrill(r,"A2")} style={{color:P.a2,fontWeight:700,cursor:"pointer",borderBottom:"1px dotted "+P.a2}}>{r.A2}</span>
                      :<span style={{color:t.muted}}>0</span>}
                  </td>
                  <td style={{padding:"7px 10px"}}>
                    {r.A3>0
                      ?<span onClick={()=>onDrill&&onDrill(r,"A3")} style={{color:P.a3,fontWeight:700,cursor:"pointer",borderBottom:"1px dotted "+P.a3}}>{r.A3}</span>
                      :<span style={{color:t.muted}}>0</span>}
                  </td>
                  <td style={{padding:"7px 10px"}}>{r.census>0?<span style={{background:"#22c55e20",color:"#22c55e",padding:"1px 8px",borderRadius:999,fontSize:10,fontWeight:700}}>{r.census}</span>:<span style={{color:t.muted}}>–</span>}</td>
                  <td style={{padding:"7px 10px"}}>{r.nonCensus>0?<span style={{background:"#6366f120",color:"#6366f1",padding:"1px 8px",borderRadius:999,fontSize:10,fontWeight:700}}>{r.nonCensus}</span>:<span style={{color:t.muted}}>–</span>}</td>
                  <td style={{padding:"7px 10px",color:t.muted,fontSize:10,maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.canvasserList||"–"}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <Pagination page={pg} setPage={setPg} total={filt.length} pageSize={PG} t={t}/>
        </div>
      </div>
    </div>
  );
}

// ── DRILL DOWN PANEL ──────────────────────────────────────────────────────────
function DrillDownPanel({drill,onClose,t,onCanvasserClick}){
  const [search,setSearch]=useState("");
  const [pg,setPg]=useState(0);
  const [sBy,setSBy]=useState("count");
  const [sDir,setSDir]=useState("desc");
  const PG=10;
  if(!drill) return null;
  const toggleS=(k)=>{if(sBy===k)setSDir(d=>d==="desc"?"asc":"desc");else{setSBy(k);setSDir("desc");setPg(0);}};
  const filt=[...drill.rows]
    .filter(r=>search?r.name.toLowerCase().includes(search.toLowerCase())||(r.cluster||"").toLowerCase().includes(search.toLowerCase()):true)
    .sort((a,b)=>{const va=sBy==="pct"?a.count/a.total:a.count;const vb=sBy==="pct"?b.count/b.total:b.count;return sDir==="desc"?vb-va:va-vb;});
  const list=filt.slice(pg*PG,(pg+1)*PG);
  return(
    <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:1000,display:"flex",alignItems:"flex-end",background:"rgba(0,0,0,0.65)",backdropFilter:"blur(4px)"}}
      onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{width:"100%",maxHeight:"82vh",background:t.card,borderRadius:"20px 20px 0 0",border:`1px solid ${t.border}`,overflow:"hidden",display:"flex",flexDirection:"column",boxShadow:"0 -8px 40px rgba(0,0,0,0.5)",fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif"}}>
        <div style={{padding:"14px 18px 10px",borderBottom:`1px solid ${t.border}`,display:"flex",alignItems:"center",gap:12,flexShrink:0}}>
          <div style={{width:12,height:12,borderRadius:3,background:drill.color,flexShrink:0}}/>
          <div style={{flex:1}}>
            <div style={{fontWeight:800,fontSize:15,color:t.text}}>{drill.label}</div>
            <div style={{fontSize:14,color:t.text,fontWeight:800,marginTop:1}}>{drill.rows.length.toLocaleString()} canvasser · {drill.total.toLocaleString()} aktivitas</div>
          </div>
          <button onClick={onClose} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:13,fontWeight:700}}>✕</button>
        </div>
        <div style={{padding:"8px 18px",borderBottom:`1px solid ${t.border}`,flexShrink:0}}>
          <input placeholder="🔍 Cari canvasser / cluster..." value={search} onChange={e=>{setSearch(e.target.value);setPg(0);}}
            style={{width:"100%",background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"7px 12px",fontSize:12,outline:"none",boxSizing:"border-box"}}/>
          <div style={{fontSize:11,color:t.muted,marginTop:4,fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif"}}>💡 Klik nama untuk lihat detail aktivitas</div>
          <div style={{display:"flex",alignItems:"center",gap:6,marginTop:5,flexWrap:"wrap"}}>
            <span style={{fontSize:10,color:t.muted,fontWeight:600}}>Sort:</span>
            {[["Jumlah","count"],["% Total","pct"]].map(([label,key])=>(
              <button key={key} onClick={()=>toggleS(key)}
                style={{background:sBy===key?drill.color:t.cardAlt,color:sBy===key?"#fff":t.muted,border:"1px solid "+t.border,borderRadius:6,padding:"3px 10px",fontSize:10,fontWeight:700,cursor:"pointer"}}>
                {label}{sBy===key?(sDir==="desc"?" ↓":" ↑"):""}
              </button>
            ))}
            <span style={{marginLeft:"auto",fontSize:10,color:t.muted}}>{filt.length} canvasser</span>
          </div>
        </div>
        <div style={{overflowY:"auto",flex:1,scrollbarWidth:"none",msOverflowStyle:"none"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif"}}>
            <thead style={{position:"sticky",top:0,background:t.card,zIndex:1}}>
              <tr style={{background:t.cardAlt}}>
                {["#","Canvasser","Region","Cluster","Jumlah","% Total",""].map(h=>(
                  <th key={h} style={{padding:"9px 12px",textAlign:"left",fontSize:11,fontWeight:700,color:t.muted,whiteSpace:"nowrap",borderBottom:`1px solid ${t.border}`}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {list.map((r,i)=>(
                <tr key={i} style={{borderBottom:`1px solid ${t.border}`,background:i%2===0?"transparent":t.rowAlt,cursor:"pointer",transition:"background 0.1s"}}
                  onMouseEnter={e=>e.currentTarget.style.background=drill.color+"18"}
                  onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"transparent":t.rowAlt}
                  onClick={()=>onCanvasserClick&&onCanvasserClick(r)}>
                  <td style={{padding:"8px 10px",color:t.muted,fontSize:11}}>{i+1}</td>
                  <td style={{padding:"8px 10px",fontWeight:600,color:drill.color,whiteSpace:"nowrap",textDecoration:"none",fontSize:12}}>{r.name}</td>
                  <td style={{padding:"8px 10px"}}>
                    <span style={{background:P.accent+"20",color:P.accent,padding:"2px 8px",borderRadius:6,fontSize:10,fontWeight:700}}>{r.region||"–"}</span>
                  </td>
                  <td style={{padding:"8px 10px",color:t.muted,fontSize:11,whiteSpace:"nowrap",maxWidth:120,overflow:"hidden",textOverflow:"ellipsis"}}>{r.cluster||"–"}</td>
                  <td style={{padding:"8px 10px"}}>
                    <div style={{display:"flex",alignItems:"center",gap:6}}>
                      <div style={{width:36,height:5,borderRadius:3,background:t.border,flexShrink:0}}>
                        <div style={{width:Math.min(100,pct(r.count,drill.rows[0]?.count||1))+"%",height:"100%",borderRadius:3,background:drill.color}}/>
                      </div>
                      <span style={{fontWeight:700,color:drill.color}}>{r.count.toLocaleString()}</span>
                    </div>
                  </td>
                  <td style={{padding:"8px 10px",color:t.muted,fontWeight:600}}>{pctS(r.count,r.total)}</td>
                  <td style={{padding:"8px 10px"}}>
                    <button style={{background:drill.color+"20",border:"none",color:drill.color,borderRadius:6,padding:"3px 8px",fontSize:10,fontWeight:700,cursor:"pointer"}}>Detail ›</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <Pagination page={pg} setPage={setPg} total={filt.length} pageSize={PG} t={t}/>
        </div>
      </div>
    </div>
  );
}

// ── CANVASSER DETAIL PANEL ────────────────────────────────────────────────────
function CanvasserDetailPanel({detail,onClose,t}){
  const [pg,setPg]=useState(0);
  const [view,setView]=useState("list");
  const [oPg,setOPg]=useState(0);
  const [vtFilter,setVtFilter]=useState("ALL");
  const [outletFilter,setOutletFilter]=useState(null);
  const [outletFilterName,setOutletFilterName]=useState(null);
  const [statusFilter,setStatusFilter]=useState(null);
  const [sortCol,setSortCol]=useState("date");
  const [sortDir,setSortDir]=useState("asc");
  const [avaDrill,setAvaDrill]=useState(null);
  const [avaRowDetail,setAvaRowDetail]=useState(null);
  const [reasonDrill,setReasonDrill]=useState(null);
  const AVA_ITEMS=[
    {label:"Poster XL/AXIS",key:"AVA Poster XL/AXIS Comply?"},
    {label:"Poster Smartfren",key:"AVA Poster Smartfren Comply?"},
    {label:"SP XL",key:"AVA SP XL Comply?"},
    {label:"SP AXIS",key:"AVA SP AXIS Comply?"},
    {label:"SP Smartfren",key:"AVA SP Smartfren Comply?"},
    {label:"Voucher XL",key:"AVA Voucher XL Comply?"},
    {label:"Voucher AXIS",key:"AVA Voucher AXIS Comply?"},
    {label:"Voucher Smartfren",key:"AVA Voucher Smartfren Comply?"},
  ];
  const PG=10;
  useEffect(()=>{setPg(0);setOPg(0);setView("list");setVtFilter("ALL");setOutletFilter(null);setOutletFilterName(null);setStatusFilter(null);setSortCol("date");setSortDir("asc");setAvaDrill(null);setAvaRowDetail(null);},[detail?.sessionKey]);
  if(!detail) return null;
  const {canvasser,drillLabel,color,rows}=detail;
  const allRows=rows._all||rows;
  const fmtDate=v=>{if(!v)return"–";const d=new Date(v);return isNaN(d)?"–":d.toLocaleDateString("id-ID",{day:"2-digit",month:"short",year:"2-digit"});};
  const fmtDist=v=>{const n=parseFloat(v);return isNaN(n)?"–":n>=1000?(n/1000).toFixed(1)+"km":n.toFixed(0)+"m";};
  const fmtDur=v=>{const n=parseFloat(v);if(isNaN(n))return"–";if(n>=60)return(n/60).toFixed(1)+"j";if(n>=1)return n.toFixed(1)+"mnt";return Math.round(n*60)+"det";};
  const vsColor=vs=>{const u=String(vs||"").toUpperCase();return u==="VALID"?P.valid:u==="OBSERVE"?P.observe:u==="INVESTIGATE"?P.investigate:u==="INCOMPLETE"?P.incomplete:"#888";};
  const reason=r=>{
    const vs    = String(r["Visit Status"]||"").toUpperCase();
    const distIn= parseFloat(r["Distance Check In (Meter)"])||0;
    const distOut=parseFloat(r["Distance Check Out (Meter)"])||0;
    const dur   = parseFloat(r["Visit Duration (Menit)"]);
    const durSt = String(r["Duration Status"]||r["_DUR"]||"").toUpperCase();
    const loc   = String(r["Location Status"]||r["_LOC"]||"").toUpperCase();
    const inR   = String(r["In Range"]||"").toLowerCase();
    if(vs==="INCOMPLETE") return "❌ Checkout tidak ada";
    const f=[];
    // Duration - main factor for OBSERVE
    if(durSt==="SHORT"||(!isNaN(dur)&&dur>0&&dur<2))  f.push(`⏱ Durasi singkat (${fmtDur(dur)})`);
    else if(durSt==="LONG"||(!isNaN(dur)&&dur>60))    f.push(`⏱ Durasi panjang (${fmtDur(dur)})`);
    // Distance
    if(distIn>5000)   f.push(`🚨 Check-in sangat jauh dari outlet (${fmtDist(distIn)})`);
    else if(distOut>5000) f.push(`🚨 Check-out sangat jauh dari outlet (${fmtDist(distOut)})`);
    else if(distIn>500)   f.push(`📍 Check-in jauh dari outlet (${fmtDist(distIn)})`);
    else if(distOut>500)  f.push(`📍 Check-out jauh dari outlet (${fmtDist(distOut)})`);
    // Location
    if(loc==="NOT MATCH") f.push("📌 Lokasi tidak match");
    // In Range
    if(inR==="no"||inR==="n") f.push("🎯 Out of range");
    return f.length>0 ? f.join(" · ") : (vs==="VALID"?"✅ Normal":"❓ "+vs);
  };
  // Get unique visit types for filter buttons
  const vtTypes=["ALL",...new Set((rows||[]).map(r=>String(r["Activity Type"]||"Unknown").trim()))];
  // Apply filter
  // Saat filter outlet+status aktif, pakai allRows supaya count cocok dengan Per Outlet
  const baseRows = (outletFilter&&statusFilter) ? (allRows||rows||[]) : (rows||[]);
  const filteredRows=baseRows.filter(r=>{
    if(vtFilter!=="ALL"&&String(r["Activity Type"]||"Unknown").trim()!==vtFilter) return false;
    if(outletFilter&&String(r["Outlet ID"]||"").trim()!==outletFilter) return false;
    if(statusFilter&&(r["_CAS1"]||"")!==statusFilter) return false;
    return true;
  });
  // Apply sort
  const sortFn=(a,b)=>{
    const dir=sortDir==="asc"?1:-1;
    if(sortCol==="date") return dir*(new Date(a["Planned Visit Date"]||0)-new Date(b["Planned Visit Date"]||0));
    if(sortCol==="outlet") return dir*(String(a["Outlet"]||"").localeCompare(String(b["Outlet"]||"")));
    if(sortCol==="status") return dir*(String(a["_CAS1"]||"").localeCompare(String(b["_CAS1"]||"")));
    if(sortCol==="dist") return dir*((parseFloat(a["Distance Check In (Meter)"])||0)-(parseFloat(b["Distance Check In (Meter)"])||0));
    if(sortCol==="dur") return dir*((parseFloat(a["Visit Duration (Menit)"])||0)-(parseFloat(b["Visit Duration (Menit)"])||0));
    return 0;
  };
  const sorted=[...filteredRows].sort(sortFn);
  // Precompute drill count per outlet from rows (drill-filtered activities)
  const drillCountMap={};
  (rows||[]).forEach(r=>{
    const oid=String(r["Outlet ID"]||"").trim();
    if(!oid) return;
    if(!drillCountMap[oid]) drillCountMap[oid]={total:0,obs:0,inv:0};
    drillCountMap[oid].total++;
    const vs=String(r["_CVS"]||r["Visit Status"]||"").toUpperCase();
    if(vs==="OBSERVE") drillCountMap[oid].obs++;
    else if(vs==="INVESTIGATE") drillCountMap[oid].inv++;
  });

  const outletMap={};
  (allRows||[]).forEach(r=>{
    const oid=String(r["Outlet ID"]||r["Outlet"]||"").trim();
    const onm=String(r["Outlet"]||oid).trim();
    const as1=r["_CAS1"]||"";
    if(!outletMap[oid])outletMap[oid]={id:oid,name:onm,total:0,A1:0,A2:0,A3:0,drill:0,drillObs:0,drillInv:0};
    outletMap[oid].total++;
    if(as1==="A1 - NORMAL")outletMap[oid].A1++;
    else if(as1==="A2 - ANOMALY")outletMap[oid].A2++;
    else if(as1==="A3 - INCOMPLETE")outletMap[oid].A3++;
    // Assign drill count from precomputed map
    const dc=drillCountMap[oid]||{total:0,obs:0,inv:0};
    outletMap[oid].drill=dc.total;
    outletMap[oid].drillObs=dc.obs;
    outletMap[oid].drillInv=dc.inv;
  });
  const outletRows=Object.values(outletMap).sort((a,b)=>{
    // Sort by drill count first (outlets with drill activities on top), then by total
    if(b.drill!==a.drill) return b.drill-a.drill;
    return b.total-a.total;
  });

  return(
    <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:1100,display:"flex",alignItems:"flex-end",background:"rgba(0,0,0,0.75)",backdropFilter:"blur(4px)"}}
      onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{width:"100%",maxHeight:"88vh",background:t.card,borderRadius:"20px 20px 0 0",border:`1px solid ${t.border}`,overflow:"hidden",display:"flex",flexDirection:"column",boxShadow:"0 -8px 40px rgba(0,0,0,0.6)",fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif"}}>
        <div style={{padding:"14px 18px 10px",borderBottom:`1px solid ${t.border}`,display:"flex",alignItems:"center",gap:12,flexShrink:0}}>
          <div style={{flex:1}}>
            <div style={{fontWeight:800,fontSize:15,color:t.text}}>👤 {canvasser?.name}</div>
            {canvasser?.id&&<div style={{fontSize:10,color:t.muted,marginTop:1}}>ID: <span style={{color:t.text,fontWeight:600}}>{canvasser.id}</span></div>}
            <div style={{fontSize:11,color:t.muted,marginTop:3,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
              <span style={{background:P.accent+"20",color:P.accent,padding:"1px 7px",borderRadius:6,fontSize:10,fontWeight:700}}>{canvasser?.region}</span>
              <span style={{color:t.muted}}>{canvasser?.cluster}</span>
              <span style={{background:color+"20",color,padding:"2px 10px",borderRadius:999,fontSize:10,fontWeight:700}}>· {drillLabel}</span>
              <span style={{color:t.text,fontWeight:800,fontSize:14}}>
                {view==="outlet"
                  ?`· ${outletRows.length} outlet dikunjungi · ${allRows.length.toLocaleString()} total aktivitas`
                  :`· ${sorted.length.toLocaleString()} aktivitas sesuai filter`}
              </span>
            </div>
            <div style={{display:"flex",gap:4,marginTop:6}}>
              <button onClick={()=>setView("list")} style={{background:view==="list"?color:t.cardAlt,color:view==="list"?"#fff":t.muted,border:"1px solid "+t.border,borderRadius:6,padding:"3px 10px",fontSize:10,fontWeight:700,cursor:"pointer"}}>📋 Aktivitas ({filteredRows.length})</button>
              <button onClick={()=>setView("outlet")} style={{background:view==="outlet"?color:t.cardAlt,color:view==="outlet"?"#fff":t.muted,border:"1px solid "+t.border,borderRadius:6,padding:"3px 10px",fontSize:10,fontWeight:700,cursor:"pointer"}}>🏪 Per Outlet ({outletRows.length})</button>
              <button onClick={()=>setView("ava")} style={{background:view==="ava"?"#f59e0b":t.cardAlt,color:view==="ava"?"#fff":t.muted,border:"1px solid "+t.border,borderRadius:6,padding:"3px 10px",fontSize:10,fontWeight:700,cursor:"pointer"}}>🏷 AVA</button>
            </div>
          </div>
          <button onClick={onClose} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:13,fontWeight:700}}>✕</button>
        </div>
           {/* Visit Type Filter */}
          {view==="list"&&(
          <div style={{padding:"8px 16px",borderBottom:`1px solid ${t.border}`,display:"flex",gap:6,flexWrap:"wrap",alignItems:"center",flexShrink:0,background:t.cardAlt}}>
            <span style={{fontSize:10,color:t.muted,fontWeight:600}}>Filter:</span>
            {(outletFilter||statusFilter)&&<button onClick={()=>{setOutletFilter(null);setOutletFilterName(null);setStatusFilter(null);setPg(0);}} style={{background:"#ef444422",color:"#ef4444",border:"1px solid #ef444440",borderRadius:6,padding:"2px 8px",fontSize:10,fontWeight:700,cursor:"pointer"}}>✕ Reset Filter</button>}
            {outletFilter&&<span style={{background:P.accent+"22",color:P.accent,padding:"3px 10px",borderRadius:6,fontSize:10,fontWeight:700}}>
              📍 Outlet: {outletFilterName||outletFilter}
              <span style={{opacity:0.7,marginLeft:4}}>({outletFilter})</span>
            </span>}
            {statusFilter&&(()=>{
              const sc=statusFilter==="A1 - NORMAL"?P.a1:statusFilter==="A2 - ANOMALY"?P.a2:P.a3;
              const sl=statusFilter==="A1 - NORMAL"?"A1 - Normal":statusFilter==="A2 - ANOMALY"?"A2 - Anomaly":"A3 - Incomplete";
              return <span style={{background:sc+"22",color:sc,padding:"3px 10px",borderRadius:6,fontSize:10,fontWeight:700}}>Status: {sl}</span>;
            })()}
            {vtTypes.map(vt=>(
              <button key={vt} onClick={()=>{setVtFilter(vt);setPg(0);}}
                style={{background:vtFilter===vt?(vt==="ALL"?color:vt==="Regular Visit"?P.a1:vt==="Ad-Hoc Visit"?P.a2:"#06b6d4"):t.card,color:vtFilter===vt?"#fff":t.muted,border:"1px solid "+(vtFilter===vt?"transparent":t.border),borderRadius:6,padding:"2px 9px",fontSize:10,fontWeight:700,cursor:"pointer"}}>
                {vt==="ALL"?"Semua":vt}
              </button>
            ))}
            <span style={{marginLeft:"auto",fontSize:10,color:t.muted}}>{sorted.length} aktivitas</span>
          </div>
          )}
          {view==="list"&&(
          <div style={{padding:"6px 16px",borderBottom:`1px solid ${t.border}`,display:"flex",gap:4,flexWrap:"wrap",alignItems:"center",flexShrink:0}}>
            <span style={{fontSize:10,color:t.muted,fontWeight:600}}>Sort:</span>
            {[["date","Tanggal"],["outlet","Outlet"],["status","Status"],["dist","Jarak"],["dur","Durasi"]].map(([sk,lbl])=>(
              <button key={sk} onClick={()=>{if(sortCol===sk)setSortDir(d=>d==="asc"?"desc":"asc");else{setSortCol(sk);setSortDir("asc");setPg(0);}}}
                style={{background:sortCol===sk?color:t.cardAlt,color:sortCol===sk?"#fff":t.muted,border:"1px solid "+t.border,borderRadius:6,padding:"2px 8px",fontSize:10,fontWeight:700,cursor:"pointer"}}>
                {lbl}{sortCol===sk?(sortDir==="asc"?" ↑":" ↓"):""}
              </button>
            ))}
          </div>
          )}
       <div style={{overflowY:"auto",flex:1,scrollbarWidth:"none",msOverflowStyle:"none"}}>
          {view==="ava"?(
            <div style={{padding:"12px 16px"}}>
              <div style={{fontWeight:700,fontSize:13,marginBottom:8,color:t.text}}>🏷 AVA Compliance Summary</div>
              <div style={{fontSize:11,color:t.muted,marginBottom:12}}>Berdasarkan {(allRows||rows||[]).length.toLocaleString()} total kunjungan · Klik item untuk lihat detail outlet</div>
              {[
                {label:"AVA Tracking", key:"AVA Tracking?"},
                {label:"Poster XL/AXIS", key:"AVA Poster XL/AXIS Comply?"},
                {label:"Poster Smartfren", key:"AVA Poster Smartfren Comply?"},
                {label:"SP XL", key:"AVA SP XL Comply?"},
                {label:"SP AXIS", key:"AVA SP AXIS Comply?"},
                {label:"SP Smartfren", key:"AVA SP Smartfren Comply?"},
                {label:"Voucher XL", key:"AVA Voucher XL Comply?"},
                {label:"Voucher AXIS", key:"AVA Voucher AXIS Comply?"},
                {label:"Voucher Smartfren", key:"AVA Voucher Smartfren Comply?"},
              ].map(({label,key})=>{
                const src=allRows||rows||[];
                const tracked=src.filter(r=>r[key]!=null&&String(r[key]||"").trim()!=="");
                const yes=src.filter(r=>String(r[key]||"").toLowerCase()==="yes").length;
                const no=tracked.length-yes;
                const pctVal=tracked.length?Math.round(yes/tracked.length*100):0;
                const color=pctVal>=80?"#22c55e":pctVal>=50?"#f59e0b":"#ef4444";
                return(
                  <div key={key} style={{marginBottom:12,cursor:"pointer",borderRadius:8,padding:"8px 10px",background:t.cardAlt,border:`1px solid ${t.border}`}}
                    onClick={()=>{
                      // Build outlet breakdown for this AVA item
                      const outMap={};
                      src.forEach(r=>{
                        const oid=String(r["Outlet ID"]||r["Outlet"]||"").trim();
                        const onm=String(r["Outlet"]||oid).trim();
                        const val=String(r[key]||"").toLowerCase();
                        if(!outMap[oid]) outMap[oid]={id:oid,name:onm,yes:0,no:0,total:0};
                        outMap[oid].total++;
                        if(val==="yes") outMap[oid].yes++;
                        else if(val==="no") outMap[oid].no++;
                      });
                      setAvaDrill({label,key,outlets:Object.values(outMap).sort((a,b)=>b.total-a.total)});
                    }}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:5,alignItems:"center"}}>
                      <span style={{fontSize:11,fontWeight:700,color:t.text}}>{label}</span>
                      <div style={{display:"flex",gap:6,alignItems:"center"}}>
                        <span style={{background:"#22c55e22",color:"#22c55e",padding:"1px 7px",borderRadius:999,fontSize:10,fontWeight:700}}>✓ {yes}</span>
                        <span style={{background:"#ef444422",color:"#ef4444",padding:"1px 7px",borderRadius:999,fontSize:10,fontWeight:700}}>✗ {no}</span>
                        <span style={{fontSize:11,fontWeight:700,color,marginLeft:2}}>{pctVal}%</span>
                        <span style={{fontSize:10,color:t.muted}}>›</span>
                      </div>
                    </div>
                    <div style={{background:t.border,borderRadius:999,height:5}}>
                      <div style={{background:color,borderRadius:999,height:5,width:pctVal+"%",transition:"width 0.4s"}}/>
                    </div>
                  </div>
                );
              })}
              {avaDrill&&(
                <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:1300,display:"flex",alignItems:"flex-end",background:"rgba(0,0,0,0.65)",backdropFilter:"blur(4px)"}} onClick={()=>setAvaDrill(null)}>
                  <div onClick={e=>e.stopPropagation()} style={{width:"100%",maxHeight:"80vh",background:t.card,borderRadius:"16px 16px 0 0",border:`1px solid ${t.border}`,overflow:"hidden",display:"flex",flexDirection:"column"}}>
                    <div style={{padding:"14px 18px 10px",borderBottom:`1px solid ${t.border}`,display:"flex",alignItems:"center",gap:10}}>
                      <div style={{flex:1}}>
                        <div style={{fontWeight:800,fontSize:14,color:t.text}}>🏷 {avaDrill.label}</div>
                        <div style={{fontSize:11,color:t.muted,marginTop:2}}>{avaDrill.outlets.length} outlet dikunjungi</div>
                      </div>
                      <button onClick={()=>setAvaDrill(null)} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"4px 10px",cursor:"pointer",fontSize:12,fontWeight:700}}>✕</button>
                    </div>
                    <div style={{overflowY:"auto",flex:1,scrollbarWidth:"none"}}>
                      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                        <thead style={{position:"sticky",top:0,background:t.card}}>
                          <tr style={{background:t.cardAlt}}>
                            {["#","Outlet ID","Outlet","✓ Comply","✗ Tidak","% Comply"].map(h=>(
                              <th key={h} style={{padding:"8px 10px",textAlign:"left",fontSize:10,fontWeight:700,color:t.muted,borderBottom:`1px solid ${t.border}`,whiteSpace:"nowrap"}}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {avaDrill.outlets.map((o,i)=>{
                            const pct2=o.total?Math.round(o.yes/o.total*100):0;
                            const c2=pct2>=80?"#22c55e":pct2>=50?"#f59e0b":"#ef4444";
                            return(
                              <tr key={o.id||i} style={{borderBottom:`1px solid ${t.border}`,background:i%2===0?"transparent":t.rowAlt}}>
                                <td style={{padding:"7px 10px",color:t.muted,fontSize:10}}>{i+1}</td>
                                <td style={{padding:"7px 10px",color:t.muted,fontSize:10,whiteSpace:"nowrap"}}>{o.id||"–"}</td>
                                <td style={{padding:"7px 10px",fontWeight:600,color:t.text,maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{o.name||"–"}</td>
                                <td style={{padding:"7px 10px",color:"#22c55e",fontWeight:700}}>{o.yes}</td>
                                <td style={{padding:"7px 10px",color:o.no>0?"#ef4444":t.muted,fontWeight:700}}>{o.no}</td>
                                <td style={{padding:"7px 10px"}}>
                                  <span style={{background:c2+"22",color:c2,padding:"1px 7px",borderRadius:999,fontSize:10,fontWeight:700}}>{pct2}%</span>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}
            </div>
          ):view==="outlet"?(<>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
              <thead style={{position:"sticky",top:0,background:t.card,zIndex:1}}>
                <tr style={{background:t.cardAlt}}>
                  {["#","Outlet ID","Outlet","Total","A1","A2","A3",drillLabel||"Drill"].map(h=>(
                    <th key={h} style={{padding:"9px 12px",textAlign:"left",fontSize:11,fontWeight:700,color:t.muted,whiteSpace:"nowrap",borderBottom:`1px solid ${t.border}`}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {outletRows.slice(oPg*PG,(oPg+1)*PG).map((r,i)=>(
                  <tr key={r.id||i} style={{borderBottom:`1px solid ${t.border}`,background:r.drill>0?(color+"12"):i%2===0?"transparent":t.rowAlt}}>
                    <td style={{padding:"7px 10px",color:t.muted,fontSize:10}}>{oPg*PG+i+1}</td>
                    <td style={{padding:"7px 10px",color:t.muted,fontSize:10,whiteSpace:"nowrap"}}>{r.id||"–"}</td>
                    <td style={{padding:"7px 10px",fontWeight:600,color:t.text,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.name}</td>
                    <td style={{padding:"7px 10px",fontWeight:800,color:t.text}}>{r.total}</td>
                    <td style={{padding:"7px 10px"}}>
                      {r.A1>0?<span onClick={()=>{setOutletFilter(r.id);setOutletFilterName(r.name);setStatusFilter("A1 - NORMAL");setVtFilter("ALL");setPg(0);setView("list");}} style={{color:P.a1,fontWeight:700,cursor:"pointer",borderBottom:"1px dotted "+P.a1}}>{r.A1}</span>:<span style={{color:t.muted}}>0</span>}
                    </td>
                    <td style={{padding:"7px 10px"}}>
                      {r.A2>0?<span onClick={()=>{setOutletFilter(r.id);setOutletFilterName(r.name);setStatusFilter("A2 - ANOMALY");setVtFilter("ALL");setPg(0);setView("list");}} style={{color:P.a2,fontWeight:700,cursor:"pointer",borderBottom:"1px dotted "+P.a2}}>{r.A2}</span>:<span style={{color:t.muted}}>0</span>}
                    </td>
                    <td style={{padding:"7px 10px"}}>
                      {r.A3>0?<span onClick={()=>{setOutletFilter(r.id);setOutletFilterName(r.name);setStatusFilter("A3 - INCOMPLETE");setVtFilter("ALL");setPg(0);setView("list");}} style={{color:P.a3,fontWeight:700,cursor:"pointer",borderBottom:"1px dotted "+P.a3}}>{r.A3}</span>:<span style={{color:t.muted}}>0</span>}
                    </td>
                    <td style={{padding:"7px 10px"}}>
                      {r.drill>0?(
                        <div onClick={()=>{setOutletFilter(r.id);setOutletFilterName(r.name);setStatusFilter(null);setVtFilter("ALL");setPg(0);setView("list");}} style={{cursor:"pointer",display:"flex",flexDirection:"column",gap:2}}>
                          <span style={{background:color+"22",color,fontWeight:800,padding:"1px 7px",borderRadius:5,fontSize:11,textAlign:"center"}}>{r.drill}</span>
                          {(r.drillObs>0||r.drillInv>0)&&<div style={{display:"flex",gap:3,justifyContent:"center"}}>
                            {r.drillObs>0&&<span style={{background:"#f59e0b22",color:"#f59e0b",fontSize:9,padding:"0px 5px",borderRadius:4,fontWeight:700}}>{r.drillObs}obs</span>}
                            {r.drillInv>0&&<span style={{background:"#ef444422",color:"#ef4444",fontSize:9,padding:"0px 5px",borderRadius:4,fontWeight:700}}>{r.drillInv}inv</span>}
                          </div>}
                        </div>
                      ):<span style={{color:t.muted,fontSize:11}}>0</span>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <Pagination page={oPg} setPage={setOPg} total={outletRows.length} pageSize={PG} t={t}/>
          </>):(<>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif"}}>
            <thead style={{position:"sticky",top:0,background:t.card,zIndex:1}}>
              <tr style={{background:t.cardAlt}}>
                {["#","Tanggal","Visit Type","Outlet ID","Outlet","Status","In Range","Jarak ke Outlet*","Durasi","AVA","Alasan"].map(h=>(
                  <th key={h} style={{padding:"9px 12px",textAlign:"left",fontSize:11,fontWeight:700,color:t.muted,whiteSpace:"nowrap",borderBottom:`1px solid ${t.border}`}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {sorted.slice(pg*PG,(pg+1)*PG).map((r,i)=>{
                const vs=String(r["Visit Status"]||"").toUpperCase();
                const vc=vsColor(vs);
                const dist=parseFloat(r["Distance Check In (Meter)"])||0;
                const dur=parseFloat(r["Visit Duration (Menit)"])||0;
                return(
                <tr key={i} style={{borderBottom:`1px solid ${t.border}`,background:i%2===0?"transparent":t.rowAlt}}>
                  <td style={{padding:"7px 10px",color:t.muted,fontSize:10}}>{i+1}</td>
                  <td style={{padding:"7px 10px",color:t.text,whiteSpace:"nowrap"}}>{fmtDate(r["Planned Visit Date"])}</td>
                  <td style={{padding:"7px 10px"}}>{(()=>{const vt=r["Activity Type"]||"";const vc=vt==="Regular Visit"?P.a1:vt==="Ad-Hoc Visit"?P.a2:"#06b6d4";return vt?<span style={{background:vc+"22",color:vc,padding:"1px 7px",borderRadius:999,fontSize:9,fontWeight:700,whiteSpace:"nowrap"}}>{vt}</span>:<span style={{color:t.muted}}>–</span>;})()}</td>
                  <td style={{padding:"7px 10px",color:t.text,maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r["Outlet"]||"–"}</td>
                  <td style={{padding:"7px 10px",color:t.muted,fontSize:10,whiteSpace:"nowrap"}}>{r["Outlet ID"]||"–"}</td>
                  <td style={{padding:"7px 10px"}}>
                    <span style={{background:vc+"20",color:vc,padding:"2px 8px",borderRadius:999,fontSize:10,fontWeight:700}}>{vs||"–"}</span>
                  </td>
                  <td style={{padding:"7px 10px"}}>
                    {r["In Range"]!=null?(
                      <span style={{background:String(r["In Range"]).toLowerCase()==="yes"||String(r["In Range"])==="1"?P.a1+"22":P.investigate+"22",color:String(r["In Range"]).toLowerCase()==="yes"||String(r["In Range"])==="1"?P.a1:P.investigate,padding:"2px 8px",borderRadius:999,fontSize:10,fontWeight:700}}>
                        {String(r["In Range"]).toLowerCase()==="yes"||String(r["In Range"])==="1"?"✓ In":"✗ Out"}
                      </span>
                    ):<span style={{color:t.muted}}>–</span>}
                  </td>
                  <td style={{padding:"7px 10px",color:dist>500?P.investigate:dist>100?P.observe:t.muted,fontWeight:dist>500?700:400}}>{fmtDist(r["Distance Check In (Meter)"])}</td>
                  <td style={{padding:"7px 10px",color:dur>0&&dur<2?P.short:t.muted,fontWeight:dur>0&&dur<2?700:400}}>
                    <span>{fmtDur(r["Visit Duration (Menit)"])}</span>
                    {dur>0&&dur<1&&<span style={{fontSize:9,color:P.investigate,marginLeft:4,fontWeight:700}}>⚡</span>}
                  </td>
                  <td style={{padding:"6px 8px",textAlign:"center"}}>
                    {(()=>{const n=AVA_ITEMS.filter(a=>String(r[a.key]||"").toLowerCase()==="yes").length;const bg=n===8?"#22c55e22":n>=4?"#f59e0b22":"#ef444422";const fg=n===8?"#22c55e":n>=4?"#f59e0b":"#ef4444";return(
                      <span onClick={()=>setAvaRowDetail({outlet:r["Outlet"],date:fmtDate(r["Planned Visit Date"]),items:AVA_ITEMS.map(a=>({label:a.label,ok:String(r[a.key]||"").toLowerCase()==="yes"})),n})}
                        style={{background:bg,color:fg,padding:"2px 6px",borderRadius:999,fontSize:10,fontWeight:700,cursor:"pointer"}}
                        onMouseEnter={e=>e.currentTarget.style.opacity="0.7"} onMouseLeave={e=>e.currentTarget.style.opacity="1"}>{n}/8</span>
                    );})()}
                  </td>
                  <td style={{padding:"7px 10px",fontSize:11,color:t.muted}}>{reason(r)}</td>
                </tr>
              );})}
            </tbody>
          </table>
          <Pagination page={pg} setPage={setPg} total={sorted.length} pageSize={PG} t={t}/>
          </>)}
          {/* Footnote - always visible */}
          <div style={{padding:"10px 16px",borderTop:`1px solid ${t.border}`,fontSize:11,color:t.muted,lineHeight:1.7,background:t.cardAlt}}>
            <b>* Jarak ke Outlet</b> = selisih koordinat GPS HP canvasser vs koordinat outlet terdaftar saat check-in/out. Bukan jarak perjalanan — jarak besar berarti canvasser kemungkinan tidak berada di lokasi outlet.
          </div>
        </div>
      </div>
      {avaRowDetail&&(
        <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:1400,display:"flex",alignItems:"flex-end",background:"rgba(0,0,0,0.65)",backdropFilter:"blur(4px)"}} onClick={()=>setAvaRowDetail(null)}>
          <div onClick={e=>e.stopPropagation()} style={{width:"100%",maxHeight:"75vh",background:t.card,borderRadius:"16px 16px 0 0",border:`1px solid ${t.border}`,overflow:"hidden",display:"flex",flexDirection:"column",fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif"}}>
            <div style={{padding:"14px 18px 10px",borderBottom:`1px solid ${t.border}`,display:"flex",alignItems:"center",gap:10}}>
              <div style={{flex:1}}>
                <div style={{fontWeight:800,fontSize:14,color:t.text}}>🏷 AVA Compliance — {avaRowDetail.n}/8</div>
                <div style={{fontSize:11,color:t.muted,marginTop:2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{avaRowDetail.outlet} · {avaRowDetail.date}</div>
              </div>
              <button onClick={()=>setAvaRowDetail(null)} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"4px 10px",cursor:"pointer",fontSize:12,fontWeight:700}}>✕</button>
            </div>
            <div style={{overflowY:"auto",flex:1,padding:"10px 18px 18px",scrollbarWidth:"none"}}>
              {avaRowDetail.items.map((it,i)=>(
                <div key={i} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"9px 4px",borderBottom:i<avaRowDetail.items.length-1?"1px solid "+t.border:"none"}}>
                  <span style={{fontSize:12,color:t.text,fontWeight:600}}>{it.label}</span>
                  <span style={{background:it.ok?"#22c55e22":"#ef444422",color:it.ok?"#22c55e":"#ef4444",padding:"2px 9px",borderRadius:999,fontSize:10,fontWeight:700}}>{it.ok?"✓ Comply":"✗ Tidak"}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
// ─────────────────────────────────────────────────────────────────────────────
// UPLOAD SCREEN
// ─────────────────────────────────────────────────────────────────────────────
function UploadScreen({onLoad,roMap,onRoLoad,t}){
  const [drag,setDrag]=useState(false);
  const [loading,setLoading]=useState(false);
  const [error,setError]=useState(null);
  const [queue,setQueue]=useState([]);
  const fileRef=useRef();
  const folderRef=useRef();

  const handleRoFile=async(fileList)=>{
    if(!fileList||!fileList.length) return;
    try{
      const maps=await Promise.all(Array.from(fileList).map(async file=>{
        const buf=await file.arrayBuffer();
        const wb=XLSX.read(buf,{type:"array", cellHTML:false});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{defval:null});
        const map={};
        rows.forEach(r=>{
          const id=String(r["Outlet ID"]||"").trim(); if(!id) return;
          map[id]={
            lat:parseFloat(r["Latitude"])||null,
            lon:parseFloat(r["Longitude"])||null,
            census:String(r["From RO Census"]||"").toUpperCase()==="YES",
            type:String(r["Outlet Type"]||"").trim(),
            cluster:String(r["Sales Cluster"]||"").trim(),
            name:String(r["Outlet Name"]||"").trim(),
          };
        });
        return map;
      }));
      const merged=Object.assign({},...maps);
      onRoLoad(prev=>({...prev,...merged}));
    }catch(e){console.error("RO parse error:",e);}
  };

    const handleFiles=useCallback(async files=>{
    setLoading(true);setError(null);
    try{
      const validFiles=Array.from(files).filter(f=>/\.(xlsx|xls|csv)$/i.test(f.name));
      if(!validFiles.length)throw new Error("Tidak ada file .xlsx/.xls/.csv ditemukan");
      const bigFiles=validFiles.filter(f=>f.size>15*1024*1024);
      if(bigFiles.length){
        setError(`⚠️ File ${bigFiles.map(f=>f.name).join(", ")} berukuran besar (>15MB). Di HP/tablet ini berisiko gagal karena keterbatasan memori browser — kalau macet di "Membaca file...", coba convert ke CSV (jauh lebih ringan) atau upload dari laptop/desktop.`);
      }
      const results=await Promise.all(validFiles.map(f=>new Promise((res,rej)=>{
        const isCsv=/\.csv$/i.test(f.name);
        const reader=new FileReader();
        reader.onload=e=>{
          try{
            const fileResults=[];
            const buildEntries=(rows,sn)=>{
              if(!rows||!rows.length) return;
              const clusterNames=[...new Set(rows.map(r=>r["Cluster"]).filter(Boolean))];
              if(clusterNames.length>1){
                // Multi-cluster file: split rows per cluster
                clusterNames.forEach(cl=>{
                  const clRows=rows.filter(r=>String(r["Cluster"]||"").trim()===cl);
                  if(!clRows.length) return;
                  const rc=getRegionCode(cl);
                  fileResults.push({name:f.name+"|"+cl,label:cl,regionCode:rc,rows:clRows,originFile:f,filterCluster:cl});
                });
              } else {
                const label=clusterNames.length===1?clusterNames[0]:(sn?sn:f.name.replace(/\.[^.]+$/,""));
                const regionCode=getRegionCode(clusterNames[0]||sn||"");
                fileResults.push({name:sn?f.name+"|"+sn:f.name,label,regionCode,rows,originFile:f});
              }
            };
            if(isCsv){
              const rows=parseCSVText(e.target.result);
              buildEntries(rows,null);
            } else {
              const wb2=XLSX.read(e.target.result,{type:"array",cellDates:true,cellHTML:false});
              for(const sn of wb2.SheetNames){
                const ws2=wb2.Sheets[sn]; if(!ws2||!ws2["!ref"]) continue;
                const {rows}=readFileRows(wb2,sn);
                buildEntries(rows,wb2.SheetNames.length>1?sn:null);
              }
            }
            if(!fileResults.length) throw new Error(`${f.name}: tidak ada data`);
            res(fileResults.length===1?fileResults[0]:{multi:true,results:fileResults,name:f.name});
          }catch(err){rej(err);}
        };
        if(isCsv) reader.readAsText(f); else reader.readAsArrayBuffer(f);
      })));
      setQueue(prev=>{
        const m=[...prev];
        const merged=[];
        const flatResults=results.flatMap(r=>r.multi?r.results:[r]);
        flatResults.forEach(r=>{
          const byName=m.findIndex(x=>x.name===r.name);
          const byLabel=m.findIndex(x=>x.label===r.label);
          if(byName>=0){
            // Same filename → replace
            m[byName]=r;
          } else if(byLabel>=0){
            // Same cluster, different file → MERGE rows
            const existing=m[byLabel];
            const existIds=new Set((existing.rows||[]).map(row=>String(row["Activity ID"]||"")));
            const newRows=(r.rows||[]).filter(row=>!existIds.has(String(row["Activity ID"]||"")));
            m[byLabel]={
              ...existing,
              rows:[...(existing.rows||[]),...newRows],
              name:existing.name.includes("+")?existing.name:`${existing.name} + ${r.name.split("|").pop()}`,
            };
            if(newRows.length>0) merged.push(`${r.label} (+${newRows.length} baris)`);
          } else {
            m.push(r);
          }
        });
        if(merged.length>0){
          setError(`✅ Data di-merge: ${merged.join(", ")}`);
          setTimeout(()=>setError(null),5000);
        }
        return m;
      });
    }catch(err){setError(err.message);}
    setLoading(false);
  },[]);


  const regionGroups={};
  queue.forEach(f=>{if(!regionGroups[f.regionCode])regionGroups[f.regionCode]=[];regionGroups[f.regionCode].push(f);});

  return(
    <div style={{minHeight:"100vh",background:t.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:"'Segoe UI',system-ui,sans-serif",padding:24}}>
      <div style={{textAlign:"center",marginBottom:32}}>
        <div style={{display:"inline-flex",alignItems:"center",gap:12,background:t.card,border:`1px solid ${t.border}`,borderRadius:16,padding:"14px 28px",marginBottom:14,boxShadow:"0 0 40px rgba(37,99,235,0.15)"}}>
          <img src="/xlsmart-logo.png" alt="XLSMART" width="46" height="46" style={{objectFit:"contain",flexShrink:0}}/>
          <div>
            <div style={{fontSize:20,fontWeight:800,color:t.text}}>XL<span style={{color:"#3b82f6"}}>SMART</span> <span style={{color:"#3b82f6"}}>Analytics</span></div>
            <div style={{fontSize:11,color:t.muted,letterSpacing:"0.1em",textTransform:"uppercase"}}>Activity Quality Dashboard</div>
          </div>
        </div>
        <p style={{color:t.muted,fontSize:13,margin:0}}>Upload file per cluster — dashboard otomatis kelompokkan per region & nasional</p>
      </div>

      <div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
        onDrop={e=>{e.preventDefault();setDrag(false);handleFiles(e.dataTransfer.files);}}
        style={{width:"100%",maxWidth:520,border:`2px dashed ${drag?P.accent:t.border}`,borderRadius:20,padding:"36px 28px",textAlign:"center",background:drag?"rgba(37,99,235,0.06)":t.card,transition:"all 0.2s"}}>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" multiple style={{display:"none"}} onChange={e=>handleFiles(e.target.files)}/>
        <input ref={folderRef} type="file" accept=".xlsx,.xls,.csv" multiple webkitdirectory="" style={{display:"none"}} onChange={e=>handleFiles(e.target.files)}/>
        {loading
          ?<><div style={{fontSize:40,marginBottom:10}}>⚙️</div><div style={{color:"#60a5fa",fontWeight:700}}>Membaca file...</div></>
          :<>
            <div style={{fontSize:46,marginBottom:10}}>{drag?"📥":"📂"}</div>
            <div style={{color:t.text,fontSize:15,fontWeight:700,marginBottom:4}}>{drag?"Lepas di sini!":"Drag & drop file XLS/XLSX/CSV"}</div>
            <div style={{color:t.muted,fontSize:12,marginBottom:18}}>Atau pilih file / folder</div>
            <div style={{display:"flex",gap:8,justifyContent:"center",flexWrap:"wrap"}}>
              <button onClick={()=>fileRef.current.click()} style={{background:"linear-gradient(135deg,#1d5fc0,#2d8ef5)",color:"#fff",border:"none",padding:"9px 22px",borderRadius:10,fontWeight:700,fontSize:13,cursor:"pointer"}}>📄 Pilih File</button>
              <button onClick={()=>folderRef.current.click()} style={{background:"linear-gradient(135deg,#065f46,#059669)",color:"#fff",border:"none",padding:"9px 22px",borderRadius:10,fontWeight:700,fontSize:13,cursor:"pointer"}}>📁 Upload Folder</button>
            </div>
          </>}
      </div>

      {/* RO Master Data Upload */}
      <div style={{width:"100%",maxWidth:520,marginTop:14,padding:"12px 16px",background:t.card,borderRadius:12,border:`1px solid ${Object.keys(roMap).length>0?"#22c55e":t.border}`}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
          <span>🏪</span>
          <span style={{fontWeight:700,fontSize:12,color:t.text}}>Master Data Outlet (RO)</span>
          {Object.keys(roMap).length>0&&<span style={{background:"#22c55e20",color:"#22c55e",padding:"1px 8px",borderRadius:999,fontSize:10,fontWeight:700}}>✓ {Object.keys(roMap).length} outlet loaded</span>}
          <span style={{marginLeft:"auto",fontSize:10,color:t.muted}}>Opsional</span>
        </div>
        <div style={{fontSize:10,color:t.muted,marginBottom:8}}>Upload RO.xlsx → hitung jarak GPS & status Census otomatis</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          <label style={{display:"inline-flex",alignItems:"center",gap:6,background:Object.keys(roMap).length>0?"#22c55e22":t.cardAlt,color:Object.keys(roMap).length>0?"#22c55e":t.muted,borderRadius:8,padding:"6px 14px",cursor:"pointer",fontSize:11,fontWeight:700,border:`1px solid ${Object.keys(roMap).length>0?"#22c55e":t.border}`}}>
            📂 {Object.keys(roMap).length>0?"✓ Tambah / Ganti RO":"Upload File RO"}
            <input type="file" accept=".xlsx,.xls" multiple style={{display:"none"}} onClick={e=>e.target.value=""} onChange={e=>handleRoFile(e.target.files)}/>
          </label>
          {Object.keys(roMap).length>0&&(
            <button onClick={()=>onRoLoad({})} style={{background:"#ef444422",color:"#ef4444",border:"1px solid #ef444440",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:11,fontWeight:700}}>
              🗑 Clear RO
            </button>
          )}
        </div>
        {Object.keys(roMap).length>0&&(
          <div style={{fontSize:10,color:t.muted,marginTop:6}}>
            Upload file RO tambahan akan <b style={{color:t.text}}>merge</b> — Outlet ID sama akan di-overwrite dengan data terbaru
          </div>
        )}
      </div>

      {error&&<div style={{marginTop:12,background:error.startsWith("⚠️")?"rgba(245,158,11,0.1)":"rgba(239,68,68,0.1)",border:`1px solid ${error.startsWith("⚠️")?"#f59e0b":"#ef4444"}`,borderRadius:10,padding:"10px 18px",color:error.startsWith("⚠️")?"#fbbf24":"#f87171",fontSize:12,maxWidth:520}}>{error}</div>}

      {queue.length>0&&(
        <div style={{width:"100%",maxWidth:520,marginTop:18}}>
          {/* Group by region */}
          {Object.entries(regionGroups).map(([rgn,files],ri)=>(
            <div key={rgn} style={{marginBottom:14}}>
              <div style={{fontSize:11,color:t.muted,fontWeight:700,marginBottom:6,letterSpacing:"0.06em",display:"flex",alignItems:"center",gap:6}}>
                <span style={{background:P.regions[ri%P.regions.length]+"22",color:P.regions[ri%P.regions.length],padding:"2px 10px",borderRadius:999,fontWeight:800}}>Region {rgn}</span>
                <span>— {files.length} cluster</span>
              </div>
              {files.map((f,i)=>(
                <div key={f.name} style={{display:"flex",alignItems:"center",gap:10,background:t.card,border:`1px solid ${t.border}`,borderRadius:10,padding:"9px 14px",marginBottom:6}}>
                  <div style={{width:8,height:8,borderRadius:2,background:P.regions[ri%P.regions.length],flexShrink:0}}/>
                  <div style={{flex:1}}>
                    <div style={{fontSize:12,fontWeight:600,color:t.text}}>{f.label}</div>
                    <div style={{fontSize:11,color:t.muted}}>{f.rows.length.toLocaleString()} aktivitas</div>
                  </div>
                  <button onClick={()=>setQueue(q=>q.filter(x=>x.name!==f.name))} style={{background:"transparent",border:"none",color:"#f87171",cursor:"pointer",fontSize:16}}>×</button>
                </div>
              ))}
            </div>
          ))}
          <div style={{display:"flex",gap:8,marginTop:4,alignItems:"center"}}>
            <div style={{fontSize:11,color:t.muted,flex:1}}>{queue.length} cluster · {Object.keys(regionGroups).length} region · {queue.reduce((s,f)=>s+f.rows.length,0).toLocaleString()} aktivitas</div>
            <button onClick={()=>onLoad(queue)} style={{background:"linear-gradient(135deg,#1d5fc0,#2d8ef5)",color:"#fff",border:"none",borderRadius:10,padding:"9px 22px",fontWeight:700,fontSize:13,cursor:"pointer"}}>
              🚀 Buka Dashboard →
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────────────────────────────────────
function Dashboard({files,onReset,onAddFiles,dark,toggleDark,roMap={}}){
  const t=dark?DARK:LIGHT;
  const [params,setParams]=useState({...DEFAULT_PARAMS});
  const [showParams,setShowParams]=useState(false);
  const [selRegion,setSelRegion]=useState(null);
  const [selCluster,setSelCluster]=useState(null);
  const [tab,setTab]=useState("overview");
  const [kpiMode,setKpiMode]=useState("activity"); // "activity" | "canvasser"
  const [showGlossary,setShowGlossary]=useState(false);
  const [insightRankTab,setInsightRankTab]=useState("A2");
  const [outletChartMode,setOutletChartMode]=useState(0); // 0=Jumlah, 1=Persentase
  const [drill,setDrill]=useState(null);
  const [canvDetail,setCanvDetail]=useState(null);
  const [outletDrill,setOutletDrill]=useState(null);
  const [trendDrill,setTrendDrill]=useState(null);
  const [trendPeriod,setTrendPeriod]=useState("daily");
  const [addLoading,setAddLoading]=useState(null);
  const [showFileManager,setShowFileManager]=useState(false);
  const [outletTypeDrill,setOutletTypeDrill]=useState(null);
  const [vtDrill,setVtDrill]=useState(null);
  const [reasonDrill,setReasonDrill]=useState(null);
  const [outletActivity,setOutletActivity]=useState(null); // {canvasser, drillLabel, color}
  // Responsive
  const [winW,setWinW]=useState(typeof window!=="undefined"?window.innerWidth:1200);
  useState(()=>{
    const handler=()=>setWinW(window.innerWidth);
    window.addEventListener("resize",handler);
    return ()=>window.removeEventListener("resize",handler);
  });
  const isMobile=winW<640;
  const isTablet=winW<1024;
  const [tPg,setTPg]=useState(0); // trend table page
  const [cPg,setCPg]=useState(0); // canvasser table page
  const TPG=10, CPG=10;
  const [sk,setSk]=useState("total");
  const [sd,setSd]=useState("desc");
  const [search,setSearch]=useState("");
  const [fq,setFq]=useState("all");

  // ── Clusters (each file = one cluster) ───────────────────────────────────
  // Cache per-file: file yang SUDAH diproses & belum berubah gak perlu diproses ulang
  // (sebelumnya nambah 1 file baru = reprocess SEMUA file lama juga → berat/crash di HP)
  const clusterCacheRef=useRef(new Map()); // key: f.name → {result, paramsKey, roMapRef, rowsRef}
  const clusters=useMemo(()=>{
    const paramsKey=JSON.stringify(params);
    return files.map((f,i)=>{
      const cached=clusterCacheRef.current.get(f.name);
      if(cached&&cached.rowsRef===f.rows&&cached.paramsKey===paramsKey&&cached.roMapRef===roMap){
        return {...cached.result,color:P.regions[i%P.regions.length]};
      }
      // Re-apply validation with current params (supports param changes)
      const reRows=f.rows.map(r=>{
        // Enrich with RO master data if available
        const rid=String(r["Outlet ID"]||"").trim();
        const ro=roMap[rid];
        const enriched=ro?{
          ...r,
          "RO Latitude":  r["RO Latitude"]  ?? ro.lat,
          "RO Longitude": r["RO Longitude"] ?? ro.lon,
          "RO Census":    r["RO Census"]    ?? (ro.census?"YES":"NO"),
          "Outlet Type":  r["Outlet Type"]  || ro.type,
        }:r;
        return computeValidation(enriched,params);
      });
      const result={
        ...processRows(reRows),
        rawRows:reRows,  // kept for canvasser detail lookup
        label:f.label,regionCode:f.regionCode,fileName:f.name,
        originFile:f.originFile, // handle ke File asli — buat lazy-reload di tahap berikutnya
        filterCluster:f.filterCluster, // nama cluster buat filter ulang kalau originFile-nya file multi-cluster
      };
      clusterCacheRef.current.set(f.name,{result,paramsKey,roMapRef:roMap,rowsRef:f.rows});
      return {...result,color:P.regions[i%P.regions.length]};
    });
  },[files,params,roMap]);

  // ── FONDASI LAZY-LOAD (Tahap 1) ──────────────────────────────────────────
  // Cache terpisah utk raw rows per-cluster, keyed by label. Belum dipakai fitur
  // manapun (masih semua fitur baca dari cl.rawRows langsung, disimpan penuh).
  // Ini disiapkan supaya di tahap berikutnya, fitur bisa dipindah pelan-pelan
  // ke sini TANPA mengubah cara kerja yang sudah jalan sekarang.
  const lazyRowsCacheRef=useRef(new Map()); // label -> rows[]
  const ensureClusterRows=useCallback(async(clusterLabel)=>{
    if(lazyRowsCacheRef.current.has(clusterLabel)) return lazyRowsCacheRef.current.get(clusterLabel);
    const cl=clusters.find(c=>c.label===clusterLabel);
    if(!cl) return [];
    // Kalau rawRows udah ada di memori (cara kerja sekarang), pakai itu langsung —
    // gak perlu baca ulang file.
    if(cl.rawRows&&cl.rawRows.length){ lazyRowsCacheRef.current.set(clusterLabel,cl.rawRows); return cl.rawRows; }
    // Fallback: baca ulang dari File asli (dipakai nanti kalau rawRows udah gak disimpan permanen)
    if(!cl.originFile) return [];
    const buf=await cl.originFile.arrayBuffer();
    const isCsv=/\.csv$/i.test(cl.originFile.name);
    let rows;
    if(isCsv){ rows=parseCSVText(new TextDecoder().decode(buf)); }
    else{ const wb=XLSX.read(buf,{type:"array",cellDates:true,cellHTML:false}); rows=readFileRows(wb).rows; }
    if(cl.filterCluster) rows=rows.filter(r=>String(r["Cluster"]||"").trim()===cl.filterCluster);
    const enriched=rows.map(r=>{
      const rid=String(r["Outlet ID"]||"").trim();
      const ro=roMap[rid];
      const withRo=ro?{
        ...r,
        "RO Latitude":  r["RO Latitude"]  ?? ro.lat,
        "RO Longitude": r["RO Longitude"] ?? ro.lon,
        "RO Census":    r["RO Census"]    ?? (ro.census?"YES":"NO"),
        "Outlet Type":  r["Outlet Type"]  || ro.type,
      }:r;
      return computeValidation(withRo,params);
    });
    lazyRowsCacheRef.current.set(clusterLabel,enriched);
    return enriched;
  },[clusters,params,roMap]);

  // ── Group clusters by region ─────────────────────────────────────────────
  const regionGroups=useMemo(()=>{
    const g={};
    clusters.forEach(c=>{if(!g[c.regionCode])g[c.regionCode]=[];g[c.regionCode].push(c);});
    return g;
  },[clusters]);

  const regionCodes=useMemo(()=>Object.keys(regionGroups).sort(),[regionGroups]);

  // ── Region aggregates ────────────────────────────────────────────────────
  const regionAgg=useMemo(()=>{
    const r={};
    regionCodes.forEach((code,i)=>{
      r[code]={...aggregateList(regionGroups[code]),label:code,regionCode:code,color:P.regions[i%P.regions.length]};
    });
    return r;
  },[regionGroups,regionCodes]);

  // ── National aggregate ───────────────────────────────────────────────────
  const national=useMemo(()=>({
    ...aggregateList(clusters),
    label:regionCodes.length===1?`${regionCodes[0]} — ${regionFullName(regionCodes[0])}`:"Nasional",
    color:P.accent,
  }),[clusters,regionCodes]);

  // ── Current view ─────────────────────────────────────────────────────────
  const view=useMemo(()=>{
    if(selCluster) return clusters.find(c=>c.label===selCluster)||national;
    if(selRegion)  return regionAgg[selRegion]||national;
    return national;
  },[selCluster,selRegion,clusters,regionAgg,national]);

  // ── Comparison data for chart (auto-switch by level) ────────────────────
  const compData=useMemo(()=>{
    // CLUSTER LEVEL: outlet type breakdown
    if(selCluster){
      const cl=clusters.find(c=>c.label===selCluster);
      if(!cl||(cl.outletData||[]).length===0) return [];
      return (cl.outletData||[]).map(d=>({
        name:d.type.replace("RO ",""),fullName:d.type,
        A1:pct(d.A1,d.total),A2:pct(d.A2,d.total),A3:pct(d.A3,d.total),
        total:d.total,color:cl.color,
      }));
    }
    // REGION LEVEL: compare clusters in that region
    if(selRegion){
      return (regionGroups[selRegion]||[]).map(c=>({
        name:c.label.replace(new RegExp("^"+selRegion+"[-_ ]?","i"),"").trim()||c.label,
        fullName:c.label,
        A1:pct((c.actC||{})["A1 - NORMAL"],c.total),
        A2:pct((c.actC||{})["A2 - ANOMALY"],c.total),
        A3:pct((c.actC||{})["A3 - INCOMPLETE"],c.total),
        total:c.total,color:c.color,
      }));
    }
    // NATIONAL LEVEL: if >1 region → compare regions; if 1 region → compare clusters
    if(regionCodes.length>1){
      return regionCodes.map((code,i)=>{
        const ra=regionAgg[code]||{actC:{},total:0};
        return{
          name:code,fullName:"Region "+code,
          A1:pct((ra.actC||{})["A1 - NORMAL"],ra.total),
          A2:pct((ra.actC||{})["A2 - ANOMALY"],ra.total),
          A3:pct((ra.actC||{})["A3 - INCOMPLETE"],ra.total),
          total:ra.total,color:P.regions[i%P.regions.length],
        };
      });
    }
    // Only 1 region → show cluster comparison at national level
    return clusters.map(c=>({
      name:c.label.replace(new RegExp("^"+(regionCodes[0]||"")+"[-_ ]?","i"),"").trim()||c.label,
      fullName:c.label,
      A1:pct((c.actC||{})["A1 - NORMAL"],c.total),
      A2:pct((c.actC||{})["A2 - ANOMALY"],c.total),
      A3:pct((c.actC||{})["A3 - INCOMPLETE"],c.total),
      total:c.total,color:c.color,
    }));
  },[selRegion,selCluster,clusters,regionGroups,regionAgg,regionCodes]);

  // ── Canvasser sort/filter ────────────────────────────────────────────────
  const sorted=useMemo(()=>{
    let list=[...view.canvassers];
    if(search)list=list.filter(c=>c.name.toLowerCase().includes(search.toLowerCase())||c.cluster.toLowerCase().includes(search.toLowerCase()));
    if(fq==="high_a2")list=list.filter(c=>c.a2p>=30);
    if(fq==="high_a3")list=list.filter(c=>c.a3p>=20);
    if(fq==="top_a1") list=list.filter(c=>c.a1p>=80);
    if(fq==="inv")    list=list.filter(c=>c.invP>=5);
    list.sort((a,b)=>{const v1=a[sk]??-1,v2=b[sk]??-1;return sd==="desc"?v2-v1:v1-v2;});
    return list;
  },[view,search,fq,sk,sd]);

  const handleSort=key=>{if(sk===key)setSd(d=>d==="desc"?"asc":"desc");else{setSk(key);setSd("desc");}};
  const mkTip=p=><Tip {...p} t={t}/>;
  const computeReasonBreakdown=(statusKey,scope=null)=>{
    let scopedClusters=clusters;
    if(scope?.clusterLabel) scopedClusters=clusters.filter(cl=>cl.label===scope.clusterLabel);
    else if(scope?.regionCode) scopedClusters=clusters.filter(cl=>cl.regionCode===scope.regionCode);
    const src=scopedClusters.flatMap(cl=>(cl.rawRows||[]).map(r=>({...r,_clLabel:cl.label})));
    const filtered=src.filter(r=>String(r["_CAS1"]||r["Activity Status.1"]||"")===(statusKey==="A2"?"A2 - ANOMALY":statusKey==="A3"?"A3 - INCOMPLETE":"A1 - NORMAL"));
    if(!filtered.length) return {reasons:[],topCanvassers:[]};
    const tot=filtered.length;
    const cnt={};
    const canvByLbl={}; // {lbl: {"name|cluster": {name,cluster,count}}}
    const canvOverall={}; // {"name|cluster": {name,cluster,count}}
    const add=(l,r)=>{
      cnt[l]=(cnt[l]||0)+1;
      const nm=r["Canvasser"]||"Unknown";
      const cl=r["_clLabel"]||"";
      const key=nm+"|"+cl;
      if(!canvByLbl[l])canvByLbl[l]={};
      if(!canvByLbl[l][key])canvByLbl[l][key]={name:nm,cluster:cl,count:0};
      canvByLbl[l][key].count++;
    };
    filtered.forEach(r=>{
      const vs=String(r["_CVS"]||r["Visit Status"]||"").toUpperCase();
      const ds=String(r["_DUR"]||r["Duration Status"]||"").toUpperCase();
      const ls=String(r["_LOC"]||r["Location Status"]||"").toUpperCase();
      const ir=String(r["In Range"]||"").toLowerCase();
      const dur=parseFloat(r["Visit Duration (Menit)"]);
      const di=parseFloat(r["Distance Check In (Meter)"]);
      const do2=parseFloat(r["Distance Check Out (Meter)"]);
      const nm=r["Canvasser"]||"Unknown";
      const clbl=r["_clLabel"]||"";
      const okey=nm+"|"+clbl;
      if(!canvOverall[okey])canvOverall[okey]={name:nm,cluster:clbl,count:0};
      canvOverall[okey].count++;
      // ── Semua kriteria yang match tetap dicatat (1 row bisa kena beberapa reason) ──
      if(vs==="INCOMPLETE"||statusKey==="A3"){add("❌ Checkout tidak ada",r);return;}
      if(ls==="NOT MATCH") add("📌 Lokasi tidak match",r);
      if(di>5000||do2>5000) add("🚨 Jarak sangat jauh (>5km)",r);
      else if(di>200||do2>200) add("📍 Jarak jauh (>200m)",r);
      if(ds==="SHORT"||(dur>0&&dur<2)) add("⏱ Durasi singkat (<2 mnt)",r);
      else if(ds==="LONG"||(dur>30)) add("⏱ Durasi panjang (>30 mnt)",r);
      if(ir==="no"||ir==="n") add("🎯 Out of range",r);
      if(vs==="INVESTIGATE") add("🔍 Investigate",r); else if(vs==="OBSERVE") add("⚠️ Observe",r);
    });
    const grandTotal=Object.values(cnt).reduce((s,v)=>s+v,0)||1;
    const reasons=Object.entries(cnt).sort((a,b)=>b[1]-a[1]).map(([l,n])=>({
      lbl:l,cnt:n,pct:Math.round(n/grandTotal*100),
      top5:Object.values(canvByLbl[l]||{}).sort((a,b)=>b.count-a.count).slice(0,5),
    }));
    const topCanvassers=Object.values(canvOverall).sort((a,b)=>b.count-a.count).slice(0,5).map(x=>({name:x.name,cluster:x.cluster,[statusKey]:x.count}));
    return {reasons,topCanvassers};
  };
  const handleAddFiles=async(fileList)=>{
    if(!fileList?.length) return;
    const fArr=Array.from(fileList);
    setAddLoading({current:0,total:fArr.length,name:"Memulai..."});
    try{
      const results=await Promise.all(fArr.map((f,fi)=>new Promise((res,rej)=>{
        setAddLoading(p=>p?{...p,current:fi,name:f.name.split(".")[0]}:null);
        const rd=new FileReader();
        rd.onload=ev=>{
          try{
            const wb2=XLSX.read(ev.target.result,{type:"array",cellDates:true,cellHTML:false});
            const fileResults=[];
            for(const sn of wb2.SheetNames){
              const ws2=wb2.Sheets[sn]; if(!ws2||!ws2["!ref"]) continue;
              const {rows}=readFileRows(wb2,sn);
              if(!rows?.length) continue;
              const cls=[...new Set(rows.map(r=>r["Cluster"]).filter(Boolean))];
              if(cls.length>1){cls.forEach(cl=>{const r2=rows.filter(r=>String(r["Cluster"]||"").trim()===cl);if(r2.length)fileResults.push({name:f.name+"|"+cl,label:cl,regionCode:getRegionCode(cl),rows:r2,originFile:f,filterCluster:cl});});}
              else{const lbl=cls[0]||(wb2.SheetNames.length>1?sn:f.name.split(".")[0]);fileResults.push({name:wb2.SheetNames.length>1?f.name+"|"+sn:f.name,label:lbl,regionCode:getRegionCode(lbl),rows,originFile:f});}
            }
            if(!fileResults.length) throw new Error(f.name+": kosong");
            res(fileResults.length===1?fileResults[0]:{multi:true,results:fileResults,name:f.name});
          }catch(e){rej(e);}
        };
        rd.readAsArrayBuffer(f);
      })));
      const flat=results.flatMap(r=>r.multi?r.results:[r]);
      onAddFiles&&onAddFiles(prev=>{
        const m=[...(prev||[])];
        flat.forEach(r=>{
          const reRows=r.rows.map(row=>{const rid=String(row["Outlet ID"]||"").trim();const ro=roMap[rid];return ro?{...row,"RO Latitude":row["RO Latitude"]??ro.lat,"RO Longitude":row["RO Longitude"]??ro.lon,"RO Census":row["RO Census"]??(ro.census?"YES":"NO"),"Outlet Type":row["Outlet Type"]||ro.type}:row;});
          const byName=m.findIndex(x=>x.name===r.name);
          const byLabel=m.findIndex(x=>x.label===r.label);
          if(byName>=0){m[byName]={...r,rows:reRows};}
          else if(byLabel>=0){const ex=m[byLabel];const exIds=new Set((ex.rows||[]).map(x=>String(x["Activity ID"]||"")));const newR=reRows.filter(x=>!exIds.has(String(x["Activity ID"]||"")));m[byLabel]={...ex,rows:[...(ex.rows||[]),...newR]};}
          else{m.push({...r,rows:reRows});}
        });
        return m;
      });
    }catch(e){console.error(e);}
    setAddLoading({current:fArr.length,total:fArr.length,name:"Selesai!"});
    setTimeout(()=>setAddLoading(null),1500);
  };

  const openVtDrill=(visitType,statusFilter)=>{
    const map={};
    clusters.forEach(cl=>(cl.rawRows||[]).forEach(r=>{
      const vt=String(r["Activity Type"]||"Unknown").trim();
      if(vt!==visitType) return;
      const as1=r["_CAS1"]||"";
      if(statusFilter&&as1!==statusFilter) return;
      const cid=String(r["Canvasser ID"]||r["Canvasser"]||"").trim();
      const nm=String(r["Canvasser"]||"").trim();
      const outId=String(r["Outlet ID"]||"").trim();
      const outNm=String(r["Outlet"]||outId).trim();
      if(!map[cid])map[cid]={id:cid,name:nm,region:cl.regionCode||"",cluster:cl.label||"",total:0,A1:0,A2:0,A3:0,outlets:new Set()};
      map[cid].total++;
      if(as1==="A1 - NORMAL")map[cid].A1++;
      else if(as1==="A2 - ANOMALY")map[cid].A2++;
      else if(as1==="A3 - INCOMPLETE")map[cid].A3++;
      if(outId)map[cid].outlets.add(outNm||outId);
    }));
    const rows=Object.values(map).map(d=>({...d,outletList:[...d.outlets].join(", "),outletCount:d.outlets.size,outlets:undefined})).sort((a,b)=>b.total-a.total);
    const label=statusFilter?(statusFilter==="A1 - NORMAL"?"A1":statusFilter==="A2 - ANOMALY"?"A2":"A3"):"Semua";
    setVtDrill({visitType,statusFilter,label,rows});
  };

  const handleOutletActivity=(row,status)=>{
    if(!outletDrill||!outletDrill.rawByOutlet)return;
    const rawRows=outletDrill.rawByOutlet[row.id]||[];
    const filtered=status==="ALL"?rawRows:rawRows.filter(r=>{
      const ca=r["_CAS1"]||"";
      if(status==="A1")return ca==="A1 - NORMAL";
      if(status==="A2")return ca==="A2 - ANOMALY";
      if(status==="A3")return ca==="A3 - INCOMPLETE";
      return true;
    });
    setOutletActivity({outletId:row.id,outletName:row.name,status,rows:filtered});
  };

  // ── Compute outlet drill data ─────────────────────────────────────────────
  const openOutletDrill=(outletType)=>{
    const map={};
    const rawByOutlet={};
    clusters.forEach(cl=>(cl.rawRows||[]).forEach(r=>{
      const ot=(()=>{const v=String(r["Outlet Type"]||"").trim();return v.toUpperCase()==="RO"?"RO OTHER":v;})();
      if(ot.toLowerCase()!==outletType.toLowerCase())return;
      // Use _VS (direct cell read) for accurate Visit Status
      const cvs=r["_CVS"]||String(r["_VS"]||r["Visit Status"]||"").toUpperCase();
      const as1=r["_CAS1"]==="A1 - NORMAL"?"A1":r["_CAS1"]==="A2 - ANOMALY"?"A2":r["_CAS1"]==="A3 - INCOMPLETE"?"A3":"";
      const outId=String(r["Outlet ID"]||r["Outlet"]||"").trim();
      const outNm=String(r["Outlet"]||outId).trim();
      const cid=String(r["Canvasser ID"]||r["Canvasser"]||"").trim();
      const cnm=String(r["Canvasser"]||"").trim();
      const isCensus=["Y","YES","1","TRUE"].includes(String(r["RO Census"]||"").trim().toUpperCase());
      // Group by Outlet ID
      if(!map[outId])map[outId]={id:outId,name:outNm,cluster:cl.label||"",region:cl.regionCode||"",total:0,A1:0,A2:0,A3:0,census:0,nonCensus:0,canvassers:new Set()};
      map[outId].total++;
      if(as1==="A1")map[outId].A1++;else if(as1==="A2")map[outId].A2++;else if(as1==="A3")map[outId].A3++;
      if(isCensus)map[outId].census++;else map[outId].nonCensus++;
      map[outId].canvassers.add(cnm);
      if(!rawByOutlet[outId])rawByOutlet[outId]=[];
      rawByOutlet[outId].push(r);
    }));
    // Convert Set to count + list
    const rows=Object.values(map).map(d=>({...d,canvasserList:[...d.canvassers].join(", "),canvasserCount:d.canvassers.size,canvassers:undefined}));
    setOutletDrill({outletType,rows:rows.sort((a,b)=>b.total-a.total),rawByOutlet});
  };

  // ── Open drill-down panel ─────────────────────────────────────────────────
  const openDrill = useCallback((label, color, countKey) => {
    const canvList = view.canvassers
      .map(c => ({ name:c.name, region:c.region, cluster:c.cluster, count:c[countKey]||0, total:c.total }))
      .filter(r => r.count > 0)
      .sort((a,b) => b.count - a.count);
    if(!canvList.length) return;
    setDrill({ label, color, countKey, rows: canvList, total: canvList.reduce((s,r)=>s+r.count,0) });
  },[view.canvassers]);

  const getCanvasserRows = useCallback((canvasserName, clusterLabel, drillKey) => {
    const cl = clusters.find(c=>c.label===clusterLabel)||clusters.find(c=>(c.rawRows||[]).some(r=>r["Canvasser"]===canvasserName));
    if(!cl||!cl.rawRows) return [];
    const all = cl.rawRows.filter(r=>r["Canvasser"]===canvasserName);
    const ff = (r) => {
      const as1=r["_CAS1"]||""; // pre-computed in processRows
      const vs=r["_CVS"]||String(r["_VS"]||r["Visit Status"]||"").toUpperCase();
      const dur=parseFloat(r["Visit Duration (Menit)"]);
      const dIn=parseFloat(r["Distance Check In (Meter)"])||0;
      const dOt=parseFloat(r["Distance Check Out (Meter)"])||0;
      const loc=String(r["Location Status"]||r["_LOC"]||"").toUpperCase();
      const inR=String(r["In Range"]||"").toLowerCase();
      const durSt=String(r["Duration Status"]||r["_DUR"]||"").toUpperCase();
      const disSt=String(r["Distance Status"]||r["_DIS"]||"").toUpperCase();
      switch(drillKey){
        case "A1": return as1==="A1 - NORMAL";
        case "A2": return as1==="A2 - ANOMALY";
        case "A3": return as1==="A3 - INCOMPLETE";
        case "VALID": return vs==="VALID";
        case "OBSERVE": return vs==="OBSERVE";
        case "INVESTIGATE": return vs==="INVESTIGATE";
        case "INCOMPLETE": return vs==="INCOMPLETE";
        case "DUR_NORMAL": return durSt==="NORMAL"||(!isNaN(dur)&&dur>=3&&dur<=60);
        case "DUR_SHORT": return durSt==="SHORT"||(!isNaN(dur)&&dur>0&&dur<2);
        case "DUR_LONG": return durSt==="LONG"||(!isNaN(dur)&&dur>60);
        case "DIS_NEAR": return disSt==="NEAR"||(dIn<=100&&dIn>0);
        case "DIS_MID": return disSt==="MID"||(dIn>100&&dIn<=1000);
        case "DIS_FAR": return disSt==="FAR"||dIn>1000;
        case "DIS_INC": return vs==="INCOMPLETE";
        case "LOC_MATCH": return loc==="MATCH";
        case "LOC_NOTMATCH": return loc==="NOT MATCH";
        case "LOC_INC": return vs==="INCOMPLETE";
        case "IR_YES": return inR==="yes"||inR==="y"||inR==="1";
        case "IR_NO": return inR==="no"||inR==="n"||inR==="0"||inR==="false";
        default: return true;
      }
    };
    const rows = drillKey ? all.filter(ff) : all;
    const sorted=rows.sort((a,b)=>new Date(a["Planned Visit Date"]||0)-new Date(b["Planned Visit Date"]||0));sorted._all=all;return sorted;
  },[clusters]);
  const card=(x={})=>({background:t.card,border:`1px solid ${t.border}`,borderRadius:14,padding:20,...x});
  const ths=key=>({padding:"9px 10px",textAlign:"left",fontSize:10,fontWeight:700,color:sk===key?P.accent:t.muted,cursor:"pointer",letterSpacing:"0.05em",textTransform:"uppercase",whiteSpace:"nowrap",userSelect:"none",borderBottom:`2px solid ${sk===key?P.accent:t.border}`});

  const T=view.total||1;
  const ac=view.actC||{};
  const vc=view.visC||{};
  const dc=view.durC||{};
  const di=view.disC||{};
  const lc=view.locC||{};

  // Klasifikasi tiap CANVASSER ke A1/A2/A3 berdasarkan status yang PALING DOMINAN (terbanyak) di antara aktivitasnya
  const canvStatusCounts=useMemo(()=>{
    const cvs=view.canvassers||[];
    const counts={A1:0,A2:0,A3:0};
    const lists={A1:[],A2:[],A3:[]};
    cvs.forEach(c=>{
      const vals=[["A1",c.A1||0],["A2",c.A2||0],["A3",c.A3||0]];
      vals.sort((a,b)=>b[1]-a[1]);
      const dom=vals[0][0];
      counts[dom]++;
      lists[dom].push(c);
    });
    Object.keys(lists).forEach(k=>lists[k].sort((a,b)=>(b[k]||0)-(a[k]||0)));
    return {counts, lists, total:cvs.length};
  },[view.canvassers]);
  const [canvCategoryDrill,setCanvCategoryDrill]=useState(null); // {label,color,statusKey,list}

  // Current level label
  const levelLabel=selCluster?"Cluster":selRegion?"Region":regionCodes.length===1?`${regionCodes[0]} — ${regionFullName(regionCodes[0])}`:"Nasional";
  const compLabel=selCluster?`Outlet Type — ${selCluster}`:selRegion?`Cluster dalam ${selRegion}`:regionCodes.length>1?`Perbandingan Region`:`Cluster dalam Region ${regionCodes[0]||""}`;

  const tabs=[
    {id:"overview",label:"Overview"},
    {id:"trend",label:"Trend"},
    {id:"outlet",label:"Outlet"},
    {id:"detail",label:"Status Detail"},
    {id:"canvasser",label:"Canvasser"},
    {id:"findings",label:"Temuan"},
  ];

  // ── AUTO FINDINGS: temuan & rekomendasi otomatis dari data ──────────────
  const findings=useMemo(()=>{
    const out=[];
    // ── Scope data sesuai drill-down yang aktif (cluster/region/nasional) ──
    const scopedClusters=selCluster?clusters.filter(cl=>cl.label===selCluster)
      :selRegion?clusters.filter(cl=>cl.regionCode===selRegion)
      :clusters;
    const scopeLabel=selCluster||selRegion||"Nasional";
    const allRows=scopedClusters.flatMap(cl=>(cl.rawRows||[]).map(r=>({...r,_clLabel:cl.label})));
    if(!allRows.length) return out;

    // Kolom AVA yang tersedia (dipakai finding outlet compliance di bawah)
    const avaCols=["AVA Poster XL/AXIS Comply?","AVA Poster Smartfren Comply?","AVA SP XL Comply?","AVA SP AXIS Comply?","AVA SP Smartfren Comply?","AVA Voucher XL Comply?","AVA Voucher AXIS Comply?","AVA Voucher Smartfren Comply?"];
    const availAva=avaCols.filter(c=>allRows.some(r=>r[c]!=null&&String(r[c]).trim()!==""));

    // 0) RINGKASAN STATUS A1/A2/A3
    {
      const a=view.actC||{};
      const totS=Object.values(a).reduce((s,v)=>s+v,0)||1;
      const a1=a["A1 - NORMAL"]||0, a2=a["A2 - ANOMALY"]||0, a3=a["A3 - INCOMPLETE"]||0;
      const a1p=Math.round(a1/totS*100), a2p=Math.round(a2/totS*100), a3p=Math.round(a3/totS*100);
      const verdict=a1p>=70?"Kepatuhan baik":a1p>=40?"Kepatuhan sedang, perlu perhatian":"Kepatuhan rendah, perlu tindakan segera";
      out.push({severity:a1p>=70?"low":a1p>=40?"medium":"high",icon:"📊",
        title:`Ringkasan Status ${scopeLabel}: A1 ${a1p}% · A2 ${a2p}% · A3 ${a3p}%`,
        desc:`Dari ${totS.toLocaleString()} aktivitas: ${a1.toLocaleString()} Normal, ${a2.toLocaleString()} Anomaly, ${a3.toLocaleString()} Incomplete. ${verdict}.`,
        rec:a1p>=70?"Pertahankan, monitor rutin saja.":"Cek breakdown A2/A3 di tab Overview & Status Detail untuk tahu penyebab utamanya.",
        action:()=>setTab("overview")});
    }

    // ══ INDIKASI FRAUD ══════════════════════════════════════════════════
    // F1) GPS check-in sama persis dipakai utk banyak outlet berbeda (canvasser gak pindah tempat)
    {
      const byCanv={};
      allRows.forEach(r=>{
        const nm=r["Canvasser"]||"";
        const lat=parseFloat(r["Check-In Latitude"]), lon=parseFloat(r["Check-In Longitude"]);
        const oid=String(r["Outlet ID"]||r["Outlet"]||"").trim();
        const onm=r["Outlet"]||oid;
        if(!nm||isNaN(lat)||isNaN(lon)||!oid) return;
        const key=nm+"|"+r["_clLabel"];
        const coordKey=lat.toFixed(3)+","+lon.toFixed(3); // ~±110m
        if(!byCanv[key])byCanv[key]={name:nm,cluster:r["_clLabel"],coords:{}};
        if(!byCanv[key].coords[coordKey])byCanv[key].coords[coordKey]={outlets:new Set(),names:[]};
        if(!byCanv[key].coords[coordKey].outlets.has(oid)){
          byCanv[key].coords[coordKey].outlets.add(oid);
          byCanv[key].coords[coordKey].names.push(onm);
        }
      });
      let worst=null;
      Object.values(byCanv).forEach(c=>{
        Object.values(c.coords).forEach(co=>{
          if(co.outlets.size>=3&&(!worst||co.outlets.size>worst.n)) worst={name:c.name,cluster:c.cluster,n:co.outlets.size,sample:co.names.slice(0,3)};
        });
      });
      if(worst){
        out.push({severity:"fraud",icon:"📍",
          title:`${worst.name} — GPS sama di ${worst.n} outlet berbeda`,
          desc:`Titik koordinat check-in PERSIS SAMA (radius ±100m) dipakai untuk ${worst.n} outlet berbeda (${worst.cluster}), termasuk "${worst.sample.join('", "')}". Kemungkinan gak pernah pindah lokasi fisik saat check-in.`,
          rec:`Cek riwayat kunjungan ${worst.name} satu-satu — kemungkinan check-in dilakukan dari 1 tempat tetap (rumah/warung), bukan dari outlet.`,
          action:()=>{const cv=(view.canvassers||[]).find(c=>c.name===worst.name&&c.cluster===worst.cluster)||{name:worst.name,cluster:worst.cluster};const rows=getCanvasserRows(worst.name,worst.cluster,"A2");setCanvDetail({canvasser:cv,drillLabel:"A2 - Anomaly",color:P.a2,rows,drillKey:"A2",sessionKey:Date.now()});}});
      }
    }

    // F2) Kecepatan perpindahan antar kunjungan mustahil (impossible travel)
    {
      const byCanv={};
      allRows.forEach(r=>{
        const nm=r["Canvasser"]||"";
        const t=r["Actual Visit Time"]?new Date(r["Actual Visit Time"]):null;
        const lat=parseFloat(r["Check-In Latitude"]), lon=parseFloat(r["Check-In Longitude"]);
        if(!nm||!t||isNaN(t.getTime())||isNaN(lat)||isNaN(lon)) return;
        const key=nm+"|"+r["_clLabel"];
        if(!byCanv[key])byCanv[key]={name:nm,cluster:r["_clLabel"],visits:[]};
        byCanv[key].visits.push({t,lat,lon,outlet:r["Outlet"]||r["Outlet ID"]||"?"});
      });
      let worst=null;
      Object.values(byCanv).forEach(c=>{
        const vs=[...c.visits].sort((a,b)=>a.t-b.t);
        for(let i=1;i<vs.length;i++){
          const dtH=(vs[i].t-vs[i-1].t)/3600000;
          if(dtH<=0||dtH>3) continue; // beda hari / data gak wajar, skip
          const distKm=haversineM(vs[i-1].lat,vs[i-1].lon,vs[i].lat,vs[i].lon)/1000;
          const speed=distKm/dtH;
          if(speed>=150&&distKm>=5&&(!worst||speed>worst.speed)){
            worst={name:c.name,cluster:c.cluster,speed,distKm,dtH,from:vs[i-1].outlet,to:vs[i].outlet};
          }
        }
      });
      if(worst){
        out.push({severity:"fraud",icon:"🚀",
          title:`${worst.name} — kecepatan perjalanan mustahil`,
          desc:`Dari "${worst.from}" ke "${worst.to}" (${worst.cluster}) — jarak ${worst.distKm.toFixed(1)}km cuma dalam ${(worst.dtH*60).toFixed(0)} menit (≈${Math.round(worst.speed)} km/jam). Gak masuk akal untuk kunjungan lapangan.`,
          rec:`Cek 2 kunjungan ini manual — indikasi salah satu titik GPS-nya di-input palsu / kunjungan gak beneran terjadi.`,
          action:()=>{const cv=(view.canvassers||[]).find(c=>c.name===worst.name&&c.cluster===worst.cluster)||{name:worst.name,cluster:worst.cluster};const rows=getCanvasserRows(worst.name,worst.cluster,"A2");setCanvDetail({canvasser:cv,drillLabel:"A2 - Anomaly",color:P.a2,rows,drillKey:"A2",sessionKey:Date.now()});}});
      }
    }

    // F3) Lokasi hampir selalu gak match (persisten, bukan sesekali)
    {
      const byCanv={};
      allRows.forEach(r=>{
        const nm=r["Canvasser"]||"";
        if(!nm) return;
        const key=nm+"|"+r["_clLabel"];
        const di=parseFloat(r["Distance Check In (Meter)"]);
        if(!byCanv[key])byCanv[key]={name:nm,cluster:r["_clLabel"],far:0,tot:0};
        byCanv[key].tot++;
        if(!isNaN(di)&&di>200) byCanv[key].far++;
      });
      const cands=Object.values(byCanv).filter(c=>c.tot>=10).map(c=>({...c,rate:c.far/c.tot}));
      const worst=[...cands].sort((a,b)=>b.rate-a.rate)[0];
      if(worst&&worst.rate>=0.9){
        out.push({severity:"fraud",icon:"🛑",
          title:`${worst.name} — ${Math.round(worst.rate*100)}% kunjungan lokasi tidak match`,
          desc:`Dari ${worst.tot.toLocaleString()} kunjungan (${worst.cluster}), ${worst.far.toLocaleString()} di antaranya GPS jauh dari outlet (>200m) — hampir selalu, bukan sesekali.`,
          rec:`Ini pola persisten, bukan kebetulan — cek langsung ke ${worst.name}, kemungkinan besar tidak benar-benar mengunjungi outlet.`,
          action:()=>{const cv=(view.canvassers||[]).find(c=>c.name===worst.name&&c.cluster===worst.cluster)||{name:worst.name,cluster:worst.cluster};const rows=getCanvasserRows(worst.name,worst.cluster,"A2");setCanvDetail({canvasser:cv,drillLabel:"A2 - Anomaly",color:P.a2,rows,drillKey:"A2",sessionKey:Date.now()});}});
      }
    }


    // 2) Cluster dengan A1 Normal rate jauh di bawah rata-rata (dalam scope aktif)
    const clusterRates=scopedClusters.map(cl=>{const a=cl.actC||{};const tot=Object.values(a).reduce((s,v)=>s+v,0)||0;return {label:cl.label,rate:tot?((a["A1 - NORMAL"]||0)/tot):0,tot};}).filter(c=>c.tot>=30);
    if(clusterRates.length>=2){
      const avg=clusterRates.reduce((s,c)=>s+c.rate,0)/clusterRates.length;
      const worst=[...clusterRates].sort((a,b)=>a.rate-b.rate)[0];
      if(worst.rate<avg*0.5&&worst.rate<0.3){
        out.push({severity:"high",icon:"⚠️",
          title:`Cluster ${worst.label} — A1 rate terendah`,
          desc:`Cuma ${Math.round(worst.rate*100)}% normal, vs rata-rata cluster lain di ${scopeLabel} (${Math.round(avg*100)}%).`,
          rec:"Investigasi lapangan: supervisi canvasser, rute, akurasi GPS di cluster ini.",
          action:()=>{setSelRegion(null);setSelCluster(worst.label);setTab("overview");}});
      }
    }

    // 3) Canvasser dengan rasio A2 jauh di atas rata-rata
    const cvs=view.canvassers||[];
    const cvRates=cvs.map(c=>{const tot=(c.A1||0)+(c.A2||0)+(c.A3||0);return {...c,rate:tot?((c.A2||0)/tot):0,tot};}).filter(c=>c.tot>=15);
    if(cvRates.length>=3){
      const avg=cvRates.reduce((s,c)=>s+c.rate,0)/cvRates.length;
      const worst=[...cvRates].sort((a,b)=>b.rate-a.rate)[0];
      if(worst.rate>avg*1.8&&worst.rate>0.3){
        out.push({severity:"medium",icon:"🔍",
          title:`${worst.name} — rasio A2 tertinggi (% bukan jumlah)`,
          desc:`${Math.round(worst.rate*100)}% dari ${worst.tot.toLocaleString()} kunjungannya (${worst.cluster}) A2 — rata-rata canvasser lain ${Math.round(avg*100)}%.`,
          rec:"Cek riwayat kunjungannya manual — kemungkinan GPS HP, rute salah, atau perlu coaching.",
          action:()=>{const rows=getCanvasserRows(worst.name,worst.cluster,"A2");setCanvDetail({canvasser:worst,drillLabel:"A2 - Anomaly",color:P.a2,rows,drillKey:"A2",sessionKey:Date.now()});}});
      }
    }

    // 4) Penyebab anomali paling dominan (dalam scope aktif)
    const reasonScope=selCluster?{clusterLabel:selCluster}:selRegion?{regionCode:selRegion}:null;
    const bdA2=computeReasonBreakdown("A2",reasonScope);
    if(bdA2.reasons.length){
      const top=bdA2.reasons[0];
      if(top.pct>=40){
        const nm=top.lbl.split(" ").slice(1).join(" ");
        out.push({severity:"medium",icon:top.lbl.split(" ")[0],
          title:`Penyebab A2 dominan di ${scopeLabel}: ${nm}`,
          desc:`${top.pct}% kasus A2 (relatif ke reason lain) — ${top.cnt.toLocaleString()} kasus.`,
          rec:"Fokus perbaikan ke akar masalah ini dulu — dampaknya paling besar.",
          action:()=>setReasonDrill({statusKey:"A2",label:"A2 - Anomaly",color:P.a2,reasons:bdA2.reasons,topCanvassers:bdA2.topCanvassers})});
      }
    }

    // 5) Proporsi Incomplete (A3) tinggi (dalam scope aktif)
    const totA=Object.values(view.actC||{}).reduce((s,v)=>s+v,0)||1;
    const a3cnt=view.actC?.["A3 - INCOMPLETE"]||0;
    const a3rate=a3cnt/totA;
    if(a3rate>=0.15){
      out.push({severity:"high",icon:"🔵",
        title:`Incomplete rate ${scopeLabel} tinggi`,
        desc:`${Math.round(a3rate*100)}% aktivitas (${a3cnt.toLocaleString()} kunjungan) checkout tidak ada.`,
        rec:"Cek app canvasser: lupa checkout, sinyal putus, atau app crash sebelum submit.",
        action:()=>openDrill("A3 - Incomplete",P.a3,"A3")});
    }

    // 6) GPS Out of Range tinggi secara nasional
    const irRows=allRows.filter(r=>r["In Range"]!=null&&String(r["In Range"]).trim()!=="");
    if(irRows.length>=50){
      const outCnt=irRows.filter(r=>{const v=String(r["In Range"]).toLowerCase();return v==="no"||v==="n";}).length;
      const outRate=outCnt/irRows.length;
      if(outRate>=0.2){
        out.push({severity:"high",icon:"🎯",
          title:`GPS Out of Range tinggi di ${scopeLabel}`,
          desc:`${Math.round(outRate*100)}% kunjungan (${outCnt.toLocaleString()}) di luar radius outlet.`,
          rec:"Cek akurasi koordinat outlet (RO Master) & kalibrasi GPS HP canvasser.",
          action:()=>openDrill("Out of Range",P.investigate,"IR_NO")});
      }
    }

    // 7) Durasi sangat singkat (<5 menit) tinggi secara nasional
    const durRows=allRows.filter(r=>{const d=parseFloat(r["Visit Duration (Menit)"]);return !isNaN(d)&&d>=0;});
    if(durRows.length>=50){
      const shortCnt=durRows.filter(r=>parseFloat(r["Visit Duration (Menit)"])<5).length;
      const shortRate=shortCnt/durRows.length;
      if(shortRate>=0.25){
        out.push({severity:"medium",icon:"⏱",
          title:`Durasi kunjungan <5 menit tinggi di ${scopeLabel}`,
          desc:`${Math.round(shortRate*100)}% kunjungan (${shortCnt.toLocaleString()}) durasinya di bawah 5 menit.`,
          rec:"Indikasi kunjungan formalitas — cek pola canvasser/cluster yang paling banyak menyumbang.",
          action:()=>openDrill("Durasi Singkat (SHORT)","#f97316","DUR_SHORT")});
      }
    }

    // 8) Outlet dengan AVA compliance terendah
    if(availAva.length>=1){
      const outletMap={};
      allRows.forEach(r=>{
        const oid=String(r["Outlet ID"]||r["Outlet"]||"").trim();
        if(!oid) return;
        const onm=String(r["Outlet"]||oid).trim();
        let yes=0,tot2=0;
        availAva.forEach(c=>{const v=String(r[c]||"").toLowerCase();if(v==="yes"||v==="no"){tot2++;if(v==="yes")yes++;}});
        if(!outletMap[oid])outletMap[oid]={id:oid,name:onm,yes:0,tot:0};
        outletMap[oid].yes+=yes;outletMap[oid].tot+=tot2;
      });
      const outlets=Object.values(outletMap).map(o=>({...o,rate:o.tot?o.yes/o.tot:0})).filter(o=>o.tot>=16);
      if(outlets.length>=3){
        const worst=[...outlets].sort((a,b)=>a.rate-b.rate)[0];
        if(worst.rate<=0.1){
          out.push({severity:"medium",icon:"🏪",
            title:`Outlet ${worst.name} — AVA compliance ${Math.round(worst.rate*100)}%`,
            desc:`Dari ${worst.tot.toLocaleString()} pengecekan item AVA, cuma ${worst.yes} yang comply.`,
            rec:"Follow-up ke outlet ini — materi promosi kemungkinan gak terpasang sama sekali."});
        }
      }
    }

    // 9) Tren A1 rate menurun (paruh awal vs paruh akhir periode)
    const trend=view.trend||[];
    if(trend.length>=8){
      const half=Math.floor(trend.length/2);
      const first=trend.slice(0,half), second=trend.slice(half);
      const rate=arr=>{const t=arr.reduce((s,d)=>s+d.total,0)||1;const a1=arr.reduce((s,d)=>s+(d.A1||0),0);return a1/t;};
      const r1=rate(first), r2=rate(second);
      if(r1-r2>=0.1){
        out.push({severity:"high",icon:"📉",
          title:`Tren A1 Normal menurun di ${scopeLabel}`,
          desc:`A1 rate turun dari ${Math.round(r1*100)}% (paruh awal) jadi ${Math.round(r2*100)}% (paruh akhir periode).`,
          rec:"Cek apa yang berubah di paruh kedua — canvasser baru, rute berubah, atau musiman.",
          action:()=>setTab("trend")});
      }
    }

    return out;
  },[clusters,view,selCluster,selRegion]);


  return(
    <>
    <div style={{minHeight:"100vh",background:t.bg,color:t.text,fontFamily:"'Segoe UI',system-ui,sans-serif",transition:"background 0.3s"}}>

      {/* ── HEADER ── */}
      <div style={{padding:"18px 22px 14px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:14}}>
        <div style={{display:"flex",alignItems:"center",gap:20,flexWrap:"wrap"}}>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <img src="/xlsmart-logo.png" alt="XLSMART" width="36" height="36" style={{objectFit:"contain",flexShrink:0,borderRadius:9}}/>
            <div>
              <div style={{fontSize:16,fontWeight:800,color:t.text}}>XLSMART Analytics</div>
              <div style={{fontSize:9,color:t.muted,letterSpacing:"0.08em",textTransform:"uppercase"}}>Activity Quality Dashboard</div>
            </div>
          </div>
          <div style={{fontSize:11,color:t.muted,paddingLeft:20,borderLeft:`1px solid ${t.border}`,lineHeight:1.5}}>
            <b style={{color:t.text,fontWeight:700}}>{clusters.length}</b> cluster · <b style={{color:t.text,fontWeight:700}}>{regionCodes.length}</b> region · <b style={{color:t.text,fontWeight:700}}>{(national.total||0).toLocaleString()}</b> aktivitas
            {national.dateRange?.min&&<> &nbsp;·&nbsp; {fmtPeriod(national.dateRange)}</>}
          </div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <button onClick={()=>setShowParams(p=>!p)} title="Parameter" style={{width:34,height:34,background:showParams?P.accent+"22":t.cardAlt,border:"1px solid "+(showParams?P.accent:t.border),color:showParams?P.accent:t.muted,borderRadius:9,cursor:"pointer",fontSize:14,display:"flex",alignItems:"center",justifyContent:"center"}}>⚙️</button>
          <button onClick={toggleDark} title="Theme" style={{width:34,height:34,background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:9,cursor:"pointer",fontSize:14,display:"flex",alignItems:"center",justifyContent:"center"}}>{dark?"☀️":"🌙"}</button>
          <button onClick={()=>setShowFileManager(true)} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.muted,borderRadius:9,padding:"0 14px",height:34,fontSize:11,cursor:"pointer",fontWeight:700}}>↩ Ganti File ({clusters.length})</button>
          <label style={{background:P.accent,border:"none",color:"#1a1200",borderRadius:9,padding:"0 16px",height:34,fontSize:11,cursor:"pointer",fontWeight:800,display:"inline-flex",alignItems:"center",gap:5}}>
            + Add Files
            <input type="file" accept=".xlsx,.xls" multiple style={{display:"none"}} onClick={e=>e.target.value=""} onChange={e=>handleAddFiles(e.target.files)}/>
          </label>
        </div>
      </div>

      {/* ── NAVIGATION: Level 1 — Nasional + Region tabs ── */}
      <div style={{padding:"0 22px",borderBottom:`1px solid ${t.border}`}}>
        <div style={{display:"flex",gap:0,overflowX:"auto",scrollbarWidth:"none",msOverflowStyle:"none"}}>
          <button onClick={()=>{setSelRegion(null);setSelCluster(null);}} style={{padding:"0 16px 11px 0",marginRight:20,border:"none",cursor:"pointer",fontSize:13,fontWeight:700,background:"transparent",whiteSpace:"nowrap",color:!selRegion&&!selCluster?t.text:t.muted,borderBottom:`2px solid ${!selRegion&&!selCluster?P.accent:"transparent"}`,transition:"all 0.15s"}}>
            {regionCodes.length===1?`${regionCodes[0]} — ${regionFullName(regionCodes[0])}`:"Nasional"}
          </button>
          {regionCodes.map((code,i)=>{
            const isActive=selRegion===code&&!selCluster;
            const clCount=(regionGroups[code]||[]).length;
            return(
              <button key={code} onClick={()=>{setSelRegion(code);setSelCluster(null);}} style={{padding:"0 16px 11px 0",marginRight:20,border:"none",cursor:"pointer",fontSize:13,fontWeight:700,background:"transparent",whiteSpace:"nowrap",color:isActive?t.text:t.muted,borderBottom:`2px solid ${isActive?P.accent:"transparent"}`,transition:"all 0.15s"}}>
                <div style={{display:"flex",flexDirection:"column",alignItems:"flex-start",gap:1}}>
                  <span>{code}</span>
                  <span style={{fontSize:9,color:t.muted,fontWeight:500}}>{clCount} cluster</span>
                </div>
              </button>
            );
          })}
        </div>

        {/* ── Level 2 — Cluster tabs (shown when region selected) ── */}
        {selRegion&&(
          <div style={{padding:"0 0 0 16px",display:"flex",gap:0,overflowX:"auto",borderTop:`1px solid ${t.border}`}}>
            <button onClick={()=>setSelCluster(null)} style={{padding:"8px 14px 8px 0",marginRight:14,border:"none",cursor:"pointer",fontSize:11,fontWeight:700,background:"transparent",whiteSpace:"nowrap",color:!selCluster?P.accent:t.muted,borderBottom:`2px solid ${!selCluster?P.accent:"transparent"}`,transition:"all 0.15s"}}>
              ∑ {selRegion} Total
            </button>
            {(regionGroups[selRegion]||[]).map((cl,i)=>{
              const isActive=selCluster===cl.label;
              return(
                <button key={cl.label} onClick={()=>setSelCluster(cl.label)} style={{padding:"8px 14px 8px 0",marginRight:14,border:"none",cursor:"pointer",fontSize:11,fontWeight:700,background:"transparent",whiteSpace:"nowrap",color:isActive?P.accent:t.muted,borderBottom:`2px solid ${isActive?P.accent:"transparent"}`,transition:"all 0.15s"}}>
                  {cl.label.replace(new RegExp(`^${selRegion}[-_ ]?`,"i"),"").trim()||cl.label}
                </button>
              );
            })}
          </div>
        )}
      </div>

      <div style={{padding:"14px 22px"}}>

        {/* ── TOGGLE: By Activity ID vs By Canvasser ── */}
        <div style={{display:"flex",gap:20,marginBottom:14,alignItems:"center",flexWrap:"wrap",borderBottom:`1px solid ${t.border}`}}>
          {[{id:"activity",label:"By Activity ID"},{id:"canvasser",label:"By Canvasser"}].map(m=>(
            <button key={m.id} onClick={()=>setKpiMode(m.id)} style={{padding:"0 0 10px",border:"none",background:"none",cursor:"pointer",fontSize:12,fontWeight:700,color:kpiMode===m.id?t.text:t.muted,borderBottom:`2px solid ${kpiMode===m.id?P.accent:"transparent"}`,marginBottom:-1}}>{m.label}</button>
          ))}
          <button onClick={()=>setShowGlossary(v=>!v)} style={{marginLeft:"auto",marginBottom:8,background:"none",border:"none",cursor:"pointer",fontSize:11,fontWeight:700,color:showGlossary?P.accent:t.muted}}>ℹ️ Apa itu A1/A2/A3?</button>
        </div>
        {showGlossary&&(
          <div style={{...card(),marginBottom:14,fontSize:12,lineHeight:1.7}}>
            <div style={{fontWeight:800,fontSize:13,marginBottom:10,color:t.text}}>📖 Glossary Status Aktivitas</div>
            {[
              {c:P.a1,icon:"✅",title:"A1 — Normal",points:[
                "Durasi kunjungan sesuai standar",
                "Lokasi GPS sesuai dengan lokasi outlet",
                "Checkout tercatat dengan lengkap",
                "Tidak terdapat indikasi kejanggalan",
              ]},
              {c:P.a2,icon:"⚠️",title:"A2 — Anomaly",points:[
                "Durasi kunjungan terlalu singkat atau terlalu panjang",
                "Lokasi GPS tidak sesuai atau berjarak jauh dari outlet",
                "Status kunjungan tercatat sebagai Observe atau Investigate",
                "Checkout tetap tercatat, namun data mengindikasikan kejanggalan",
              ]},
              {c:P.a3,icon:"🔵",title:"A3 — Incomplete",points:[
                "Canvasser melakukan check-in namun tidak ada checkout yang tercatat",
                "Kunjungan tidak dapat divalidasi sepenuhnya",
                "Merupakan prioritas pemeriksaan tertinggi",
              ]},
            ].map((g,i)=>(
              <div key={i} style={{display:"flex",gap:10,marginBottom:i<2?10:14,paddingBottom:i<2?10:0,borderBottom:i<2?`1px solid ${t.border}`:"none"}}>
                <span style={{fontSize:16,flexShrink:0}}>{g.icon}</span>
                <div>
                  <div style={{fontWeight:700,color:g.c,marginBottom:5}}>{g.title}</div>
                  <ul style={{margin:0,paddingLeft:18,color:t.muted}}>
                    {g.points.map((p,pi)=><li key={pi} style={{marginBottom:3}}>{p}</li>)}
                  </ul>
                </div>
              </div>
            ))}
            <div style={{background:t.cardAlt,borderRadius:8,padding:"9px 11px",fontSize:11,color:t.muted,lineHeight:1.6}}>
              <div style={{marginBottom:6}}><b style={{color:t.text}}>📋 By Activity ID</b></div>
              <ul style={{margin:"0 0 8px",paddingLeft:16}}>
                <li>Persentase dan jumlah dihitung berdasarkan total aktivitas/kunjungan</li>
                <li>Satu canvasser dapat menyumbang lebih dari satu baris data</li>
              </ul>
              <div style={{marginBottom:6}}><b style={{color:t.text}}>👤 By Canvasser</b></div>
              <ul style={{margin:0,paddingLeft:16}}>
                <li>Persentase dan jumlah dihitung berdasarkan total orang (canvasser)</li>
                <li>Setiap canvasser diklasifikasikan ke A1/A2/A3 berdasarkan status yang paling sering muncul (dominan) di antara aktivitasnya</li>
              </ul>
            </div>
          </div>
        )}
        {/* ── KPI LIST (flat row style) ── */}
        <div style={{marginBottom:20}}>
          {(kpiMode==="canvasser"?[
            {label:"Total",val:canvStatusCounts.total.toLocaleString(),color:t.text},
            {label:"A1 — Normal",val:pctS(canvStatusCounts.counts.A1,canvStatusCounts.total||1),color:P.a1,sub:canvStatusCounts.counts.A1.toLocaleString()+" canvasser",drill:()=>setCanvCategoryDrill({label:"A1 - Normal",color:P.a1,statusKey:"A1",list:canvStatusCounts.lists.A1})},
            {label:"A2 — Anomaly",val:pctS(canvStatusCounts.counts.A2,canvStatusCounts.total||1),color:P.a2,sub:canvStatusCounts.counts.A2.toLocaleString()+" canvasser",drill:()=>setCanvCategoryDrill({label:"A2 - Anomaly",color:P.a2,statusKey:"A2",list:canvStatusCounts.lists.A2})},
            {label:"A3 — Incomplete",val:pctS(canvStatusCounts.counts.A3,canvStatusCounts.total||1),color:P.a3,sub:canvStatusCounts.counts.A3.toLocaleString()+" canvasser",drill:()=>setCanvCategoryDrill({label:"A3 - Incomplete",color:P.a3,statusKey:"A3",list:canvStatusCounts.lists.A3})},
            {label:"Investigate",val:pctS(vc["INVESTIGATE"],T),color:t.muted,sub:(vc["INVESTIGATE"]||0).toLocaleString()+" aktivitas",drill:()=>openDrill("Investigate",P.investigate,"INVESTIGATE")},
            {label:"Total Aktivitas",val:T.toLocaleString(),color:t.muted},
          ]:[
            {label:"Total Aktivitas",val:T.toLocaleString(),color:t.text},
            {label:"A1 — Normal",val:pctS(ac["A1 - NORMAL"],T),color:P.a1,sub:(ac["A1 - NORMAL"]||0).toLocaleString(),drill:()=>openDrill("A1 - Normal",P.a1,"A1")},
            {label:"A2 — Anomaly",val:pctS(ac["A2 - ANOMALY"],T),color:P.a2,sub:(ac["A2 - ANOMALY"]||0).toLocaleString(),drill:()=>openDrill("A2 - Anomaly",P.a2,"A2")},
            {label:"A3 — Incomplete",val:pctS(ac["A3 - INCOMPLETE"],T),color:P.a3,sub:(ac["A3 - INCOMPLETE"]||0).toLocaleString(),drill:()=>openDrill("A3 - Incomplete",P.a3,"A3")},
            {label:"Investigate",val:pctS(vc["INVESTIGATE"],T),color:t.muted,sub:(vc["INVESTIGATE"]||0).toLocaleString(),drill:()=>openDrill("Investigate",P.investigate,"INVESTIGATE")},
            {label:"Canvasser",val:view.canvassers.length.toLocaleString(),color:t.muted},
          ]).map((k,i)=>{
            const barPct=typeof k.val==="string"&&k.val.includes("%")?parseFloat(k.val):null;
            return(
              <div key={i} onClick={k.drill||undefined} style={{display:"flex",alignItems:"center",padding:"11px 0",borderBottom:i<5?`1px solid ${t.border}`:"none",cursor:k.drill?"pointer":"default"}}>
                <div style={{flex:1,minWidth:0,marginRight:12}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,fontSize:13,fontWeight:600,color:t.text}}>
                    {barPct!=null&&<span style={{width:7,height:7,borderRadius:"50%",background:k.color,flexShrink:0}}/>}
                    <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{k.label}</span>
                  </div>
                  {barPct!=null&&<div style={{width:"100%",height:3,background:t.border,borderRadius:99,marginTop:6}}><div style={{width:barPct+"%",height:3,background:k.color,borderRadius:99}}/></div>}
                </div>
                {k.sub&&<div style={{fontSize:11,color:t.muted,width:isMobile?70:90,textAlign:"right",flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{k.sub}</div>}
                <div style={{fontSize:isMobile?14:16,fontWeight:800,color:k.color,width:isMobile?58:70,textAlign:"right",flexShrink:0}}>{k.val}</div>
              </div>
            );
          })}
        </div>

        {/* ── TAB BUTTONS ── */}
        <div style={{display:"flex",gap:22,marginBottom:20,flexWrap:"wrap",borderBottom:`1px solid ${t.border}`,overflowX:"auto"}}>
          {tabs.map(tb=>(
            <button key={tb.id} onClick={()=>setTab(tb.id)} style={{padding:"0 0 11px",border:"none",background:"none",cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap",color:tab===tb.id?t.text:t.muted,borderBottom:`2px solid ${tab===tb.id?P.accent:"transparent"}`,marginBottom:-1,transition:"color 0.15s"}}>{tb.label}</button>
          ))}
        </div>

        {/* ════ OVERVIEW ════ */}
        {tab==="overview"&&(
          <div style={{display:"grid",gap:16}}>

            {/* ── ROW 1: Activity Status Pie (left) + Key Insights (right) ── */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(260px,1fr))",gap:16}}>
              <div style={card()}>
                <div style={{fontWeight:700,marginBottom:2,fontSize:11,letterSpacing:"0.06em",textTransform:"uppercase",color:t.muted}}>Activity Status</div>
                <div style={{fontSize:11,color:t.muted,marginBottom:10}}>
                  {view.label}
                  {view.label&&view.label.length<=3&&REGION_NAMES[view.label]&&<span style={{color:t.muted}}> ({REGION_NAMES[view.label]})</span>}
                  {view.label&&view.label.startsWith&&(()=>{const m=view.label.match(/^([A-Z]{2,3})-/);return m&&REGION_NAMES[m[1]]?<span style={{color:t.muted}}> · {REGION_NAMES[m[1]]}</span>:null;})()}
                </div>
                <ResponsiveContainer width="100%" height={190}>
                  <PieChart>
                    <Pie data={ACT.map(s=>({name:s.label,value:ac[s.key]||0,color:s.color,skey:s.key}))} cx="50%" cy="50%" outerRadius={75} innerRadius={32} dataKey="value" strokeWidth={0}
                      label={({value})=>value>0?pctS(value,T):""} labelLine={false}
                      onClick={d=>{const m={"A1 - NORMAL":"A1","A2 - ANOMALY":"A2","A3 - INCOMPLETE":"A3"};openDrill(d.name,d.color,m[d.skey]);}}>
                      {ACT.map((s,i)=><Cell key={i} fill={s.color} style={{cursor:"pointer"}}/>)}
                    </Pie>
                    <Tooltip content={mkTip}/>
                  </PieChart>
                </ResponsiveContainer>
                {ACT.map((s,i)=>{
                  const aKey={"A1 - NORMAL":"A1","A2 - ANOMALY":"A2","A3 - INCOMPLETE":"A3"}[s.key];
                  return(
                  <div key={i} style={{borderBottom:i<2?`1px solid ${t.border}`:"none"}}>
                    <div onClick={()=>openDrill(s.label,s.color,aKey)} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 0",cursor:"pointer"}} onMouseEnter={e=>e.currentTarget.style.opacity="0.75"} onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
                      <div style={{width:7,height:7,borderRadius:"50%",background:s.color,flexShrink:0}}/>
                      <span style={{fontSize:12,color:t.text,flex:1,fontWeight:600}}>{s.label}</span>
                      <span style={{fontSize:12,fontWeight:800,color:s.color}}>{pctS(ac[s.key],T)}</span>
                    </div>
                    {aKey!=="A1"&&<div onClick={e=>{e.stopPropagation();const bd=computeReasonBreakdown(aKey);setReasonDrill({statusKey:aKey,label:s.label,color:s.color,reasons:bd.reasons,topCanvassers:bd.topCanvassers});}} style={{fontSize:10,fontWeight:700,color:P.accent,cursor:"pointer",paddingBottom:8}}>🔍 Lihat penyebab ›</div>}
                  </div>
                  );
                })}
                <div style={{marginTop:10,padding:"8px 0 4px",fontSize:11,color:t.muted,fontWeight:700,letterSpacing:"0.06em",borderTop:`1px solid ${t.border}`}}>VISIT STATUS</div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))",gap:5,marginTop:4}}>
                  {VIS.map((s,i)=>(
                    <div key={i} onClick={()=>openDrill(s.label,s.color,s.key)} style={{display:"flex",alignItems:"center",gap:6,padding:"5px 8px",background:t.cardAlt,borderRadius:7,cursor:"pointer"}}
                      onMouseEnter={e=>e.currentTarget.style.opacity="0.75"} onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
                      <div style={{width:7,height:7,borderRadius:2,background:s.color}}/>
                      <span style={{fontSize:10,color:t.muted,flex:1}}>{s.label}</span>
                      <span style={{fontSize:11,fontWeight:700,color:s.color}}>{pctS(vc[s.key],T)}</span>
                    </div>
                  ))}
                </div>
              </div>

              <div style={card()}>
                <div style={{fontWeight:700,marginBottom:14,fontSize:11,letterSpacing:"0.06em",textTransform:"uppercase",color:t.muted}}>Key Insights</div>
                {(()=>{
                  const wA2=[...view.canvassers].filter(c=>c.A2>0).sort((a,b)=>b.A2-a.A2)[0];
                  const wA3=[...view.canvassers].filter(c=>c.A3>0).sort((a,b)=>b.A3-a.A3)[0];
                  const wInv=[...view.canvassers].filter(c=>c.INVESTIGATE>0).sort((a,b)=>b.INVESTIGATE-a.INVESTIGATE)[0];
                  const openInsight=(canv,key,label,color)=>{
                    if(!canv) return;
                    const rows=getCanvasserRows(canv.name,canv.cluster,key);
                    setCanvDetail({canvasser:canv,drillLabel:label,color,rows,drillKey:key,sessionKey:Date.now()});
                  };

                  // ── Section 1: Top Canvasser per Status ──
                  const topCanvasserItems=[
                    wA2&&{icon:"⚠️",color:P.a2,title:"A2 Anomaly terbanyak",name:wA2.name,cluster:wA2.cluster,val:wA2.A2,onClick:()=>openInsight(wA2,"A2","A2 - Anomaly",P.a2)},
                    wA3&&{icon:"🔵",color:P.a3,title:"A3 Incomplete terbanyak",name:wA3.name,cluster:wA3.cluster,val:wA3.A3,onClick:()=>openInsight(wA3,"A3","A3 - Incomplete",P.a3)},
                    wInv&&{icon:"🔍",color:P.investigate,title:"Investigate terbanyak",name:wInv.name,cluster:wInv.cluster,val:wInv.INVESTIGATE,onClick:()=>openInsight(wInv,"INVESTIGATE","Investigate",P.investigate)},
                  ].filter(Boolean);

                  // ── Section 2: Ranking Cluster & Region (bertab A1/A2/A3) ──
                  const rankKeyMap={A1:"A1 - NORMAL",A2:"A2 - ANOMALY",A3:"A3 - INCOMPLETE"};
                  const rankColorMap={A1:P.a1,A2:P.a2,A3:P.a3};
                  const rc=rankColorMap[insightRankTab];
                  const rk=rankKeyMap[insightRankTab];
                  const topRegions=regionCodes.length>1?regionCodes.map(code=>({code,regionCode:code,v:(regionAgg[code]?.actC||{})[rk]||0})).sort((a,b)=>b.v-a.v).slice(0,3):[];
                  const topClusters=clusters.length>1?[...clusters].sort((a,b)=>((b.actC||{})[rk]||0)-((a.actC||{})[rk]||0)).slice(0,3).map(cl=>({code:cl.label.split("-").slice(1).join("-")||cl.label,clusterLabel:cl.label,v:(cl.actC||{})[rk]||0})):[];
                  const openRank=(r)=>{
                    if(!r.v) return;
                    const scope=r.clusterLabel?{clusterLabel:r.clusterLabel}:r.regionCode?{regionCode:r.regionCode}:null;
                    const bd=computeReasonBreakdown(insightRankTab,scope);
                    const fullLbl=(ACT.find(a=>a.short===insightRankTab)||{}).label||insightRankTab;
                    setReasonDrill({statusKey:insightRankTab,label:`${fullLbl} — ${r.code}`,color:rc,reasons:bd.reasons,topCanvassers:bd.topCanvassers});
                  };

                  // ── Section 3: Pola Kunjungan ──
                  const vtData=view.visitTypeData||[];
                  const vtTotal=vtData.reduce((s,v)=>s+v.total,0)||1;
                  const regularVt=vtData.find(v=>v.type==="Regular Visit");
                  const adhocVt=vtData.find(v=>v.type==="Ad-Hoc Visit");
                  const shortCnt=dc["SHORT"]||0;

                  // ── Section 4: Penyebab Utama (dipisah per bucket biar jelas asalnya) ──
                  const rm=view.reasonMap||{};
                  const topReasonList=(bucket)=>Object.entries(rm[bucket]||{}).sort((a,b)=>b[1]-a[1]).slice(0,3);
                  const invReasons=topReasonList("investigate");
                  const obsReasons=topReasonList("observe");
                  const invTotal=Object.values(rm.investigate||{}).reduce((s,v)=>s+v,0);
                  const obsTotal=Object.values(rm.observe||{}).reduce((s,v)=>s+v,0);

                  return(<>
                    {/* Section 1 */}
                    {topCanvasserItems.length>0&&<>
                      <div style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:2}}>Top Canvasser per Status</div>
                      <div style={{fontSize:10,color:t.muted,marginBottom:8}}>Klik buat lihat detail kunjungannya</div>
                      {topCanvasserItems.map((it,i)=>(
                        <div key={i} onClick={it.onClick} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 0",borderBottom:`1px solid ${t.border}`,cursor:"pointer"}}
                          onMouseEnter={e=>e.currentTarget.style.opacity="0.75"} onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
                          <div style={{width:24,height:24,borderRadius:7,background:it.color+"1f",display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,flexShrink:0}}>{it.icon}</div>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontSize:10,color:t.muted}}>{it.title}</div>
                            <div style={{fontSize:12,fontWeight:700,color:t.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{it.name}</div>
                            <div style={{fontSize:9,color:t.muted}}>{it.cluster}</div>
                          </div>
                          <div style={{fontSize:13,fontWeight:800,color:it.color,flexShrink:0}}>{it.val.toLocaleString()}</div>
                        </div>
                      ))}
                    </>}

                    {/* Section 2 */}
                    {(topRegions.length>0||topClusters.length>0)&&<div style={{marginTop:18}}>
                      <div style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:8}}>Ranking Cluster & Region</div>
                      <div style={{display:"flex",gap:14,marginBottom:10}}>
                        {["A1","A2","A3"].map(k=>(
                          <span key={k} onClick={()=>setInsightRankTab(k)} style={{fontSize:11,fontWeight:700,color:insightRankTab===k?rankColorMap[k]:t.muted,paddingBottom:5,borderBottom:`2px solid ${insightRankTab===k?rankColorMap[k]:"transparent"}`,cursor:"pointer"}}>{k}</span>
                        ))}
                      </div>
                      <div style={{display:"flex",gap:20}}>
                        {topRegions.length>0&&<div style={{flex:1,minWidth:0}}>
                          <div style={{fontSize:9,color:t.muted,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:6}}>🗺 Top Region</div>
                          {topRegions.map((r,ri)=>(
                            <div key={ri} onClick={()=>openRank(r)} style={{display:"flex",justifyContent:"space-between",fontSize:11,padding:"5px 0",borderBottom:ri<topRegions.length-1?`1px solid ${t.border}`:"none",cursor:r.v?"pointer":"default"}}
                              onMouseEnter={e=>{if(r.v)e.currentTarget.style.opacity="0.7";}} onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
                              <span style={{color:t.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{["🥇","🥈","🥉"][ri]} {r.code}</span>
                              <span style={{color:t.muted,fontWeight:700,flexShrink:0,marginLeft:6}}>{r.v.toLocaleString()}</span>
                            </div>
                          ))}
                        </div>}
                        {topClusters.length>0&&<div style={{flex:1,minWidth:0}}>
                          <div style={{fontSize:9,color:t.muted,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:6}}>📍 Top Cluster</div>
                          {topClusters.map((r,ri)=>(
                            <div key={ri} onClick={()=>openRank(r)} style={{display:"flex",justifyContent:"space-between",fontSize:11,padding:"5px 0",borderBottom:ri<topClusters.length-1?`1px solid ${t.border}`:"none",cursor:r.v?"pointer":"default"}}
                              onMouseEnter={e=>{if(r.v)e.currentTarget.style.opacity="0.7";}} onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
                              <span style={{color:t.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{["🥇","🥈","🥉"][ri]} {r.code}</span>
                              <span style={{color:t.muted,fontWeight:700,flexShrink:0,marginLeft:6}}>{r.v.toLocaleString()}</span>
                            </div>
                          ))}
                        </div>}
                      </div>
                    </div>}

                    {/* Section 3 */}
                    <div style={{marginTop:18}}>
                      <div style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:8}}>Pola Kunjungan</div>
                      <div style={{display:"flex",border:`1px solid ${t.border}`,borderRadius:10,overflow:"hidden"}}>
                        {regularVt&&<div onClick={()=>openVtDrill("Regular Visit",null)} style={{flex:1,padding:"10px 6px",textAlign:"center",borderRight:`1px solid ${t.border}`,cursor:"pointer"}}>
                          <div style={{fontSize:14,fontWeight:800,color:P.accent}}>{pctS(regularVt.total,vtTotal)}</div>
                          <div style={{fontSize:8,color:t.muted,marginTop:2,textTransform:"uppercase"}}>Regular<br/>Visit</div>
                        </div>}
                        {adhocVt&&<div onClick={()=>openVtDrill("Ad-Hoc Visit",null)} style={{flex:1,padding:"10px 6px",textAlign:"center",borderRight:`1px solid ${t.border}`,cursor:"pointer"}}>
                          <div style={{fontSize:14,fontWeight:800,color:P.accent}}>{pctS(adhocVt.total,vtTotal)}</div>
                          <div style={{fontSize:8,color:t.muted,marginTop:2,textTransform:"uppercase"}}>Ad-Hoc<br/>Visit</div>
                        </div>}
                        <div onClick={()=>openDrill("Durasi Singkat (SHORT)","#f97316","DUR_SHORT")} style={{flex:1,padding:"10px 6px",textAlign:"center",cursor:"pointer"}}>
                          <div style={{fontSize:14,fontWeight:800,color:P.accent}}>{pctS(shortCnt,T)}</div>
                          <div style={{fontSize:8,color:t.muted,marginTop:2,textTransform:"uppercase"}}>Durasi<br/>&lt;2 Menit</div>
                        </div>
                      </div>
                    </div>

                    {/* Section 4 */}
                    {(invReasons.length>0||obsReasons.length>0)&&<div style={{marginTop:18}}>
                      <div style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:2}}>Alasan Utama di Balik Visit Status</div>
                      <div style={{fontSize:10,color:t.muted,marginBottom:10}}>Catatan lapangan yang paling sering disebut canvasser</div>
                      {invReasons.length>0&&<>
                        <div onClick={()=>openDrill("Investigate",P.investigate,"INVESTIGATE")} style={{fontSize:10,fontWeight:700,color:P.investigate,marginBottom:6,cursor:"pointer"}}>🔍 Investigate ({invTotal.toLocaleString()} kunjungan) ›</div>
                        {invReasons.map(([k,v],i)=>(
                          <div key={i} onClick={()=>openDrill("Investigate",P.investigate,"INVESTIGATE")} style={{display:"flex",alignItems:"center",gap:10,padding:"5px 0",cursor:"pointer"}}>
                            <div style={{fontSize:11,color:t.text,width:120,flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{k}</div>
                            <div style={{flex:1,height:5,background:t.border,borderRadius:99,overflow:"hidden"}}><div style={{width:(v/invReasons[0][1]*100)+"%",height:"100%",background:P.investigate,borderRadius:99}}/></div>
                            <div style={{fontSize:11,fontWeight:700,color:P.investigate,width:70,textAlign:"right",flexShrink:0}}>{v.toLocaleString()} kunjungan</div>
                          </div>
                        ))}
                      </>}
                      {obsReasons.length>0&&<>
                        <div onClick={()=>openDrill("Observe",P.a2,"OBSERVE")} style={{fontSize:10,fontWeight:700,color:P.a2,marginTop:14,marginBottom:6,cursor:"pointer"}}>⚠️ Observe ({obsTotal.toLocaleString()} kunjungan) ›</div>
                        {obsReasons.map(([k,v],i)=>(
                          <div key={i} onClick={()=>openDrill("Observe",P.a2,"OBSERVE")} style={{display:"flex",alignItems:"center",gap:10,padding:"5px 0",cursor:"pointer"}}>
                            <div style={{fontSize:11,color:t.text,width:120,flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{k}</div>
                            <div style={{flex:1,height:5,background:t.border,borderRadius:99,overflow:"hidden"}}><div style={{width:(v/obsReasons[0][1]*100)+"%",height:"100%",background:P.a2,borderRadius:99}}/></div>
                            <div style={{fontSize:11,fontWeight:700,color:P.a2,width:70,textAlign:"right",flexShrink:0}}>{v.toLocaleString()} kunjungan</div>
                          </div>
                        ))}
                      </>}
                    </div>}
                  </>);
                })()}
              </div>
            </div>

            {/* ── ROW 2: Comparison chart (region/cluster) OR Top 5 per status (cluster level) ── */}
            {selCluster?(
              // Cluster level: Top 5 canvassers per A1, A2, A3
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))",gap:24,marginBottom:24}}>
                {[
                  {label:"Top 5 — A1 Normal",color:P.a1,key:"A1",sort:"A1"},
                  {label:"Top 5 — A2 Anomaly",color:P.a2,key:"A2",sort:"A2"},
                  {label:"Top 5 — A3 Incomplete",color:P.a3,key:"A3",sort:"A3"},
                ].map((cat,ci)=>{
                  const top5=[...view.canvassers].sort((a,b)=>b[cat.sort]-a[cat.sort]).slice(0,5);
                  const maxV=top5[0]?.[cat.sort]||1;
                  return(
                  <div key={ci}>
                    <div style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:10}}>{cat.label}</div>
                    {top5.map((cv,i)=>(
                      <div key={i} style={{marginBottom:9}}>
                        <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}>
                          <span style={{fontWeight:600,color:t.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:"70%"}}>{cv.name}</span>
                          <span style={{fontWeight:700,color:cat.color}}>{(cv[cat.sort]||0).toLocaleString()}</span>
                        </div>
                        <div style={{height:3,borderRadius:99,background:t.border}}>
                          <div style={{width:pct(cv[cat.sort]||0,maxV)+"%",height:"100%",borderRadius:99,background:cat.color}}/>
                        </div>
                        <div style={{fontSize:10,color:t.muted,marginTop:2}}>{cv.cluster} · {pctS(cv[cat.sort]||0,cv.total)} dari total aktivitas dia</div>
                      </div>
                    ))}
                    <div onClick={()=>openDrill(cat.label,cat.color,cat.key)} style={{fontSize:11,fontWeight:700,color:cat.color,cursor:"pointer",marginTop:6}}>
                      Lihat semua ›
                    </div>
                  </div>
                );})}
              </div>
            ):(
              // National/Region level: comparison chart
              compData.length>0&&(
              <div style={{marginBottom:24}}>
                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
                  <div>
                    <div style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase"}}>{selRegion?`Perbandingan Cluster dalam Region ${selRegion}`:regionCodes.length>1?"Perbandingan Antar Region":`Perbandingan Cluster`}</div>
                    <div style={{fontSize:11,color:t.muted,marginTop:2}}>Klik bar atau baris tabel untuk lihat detail</div>
                  </div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))",gap:20}}>
                  <ResponsiveContainer width="100%" height={220}>
                    <BarChart data={compData} margin={{top:10,right:10,bottom:10,left:0}}
                      onClick={d=>{
                        if(!d?.activePayload) return;
                        const item=compData.find(x=>x.name===d.activeLabel);
                        if(!item) return;
                        if(selRegion){setSelCluster(item.fullName);}
                        else if(regionCodes.length>1){setSelRegion(item.fullName.replace("Region ",""));}
                        else{setSelRegion(regionCodes[0]);setSelCluster(item.fullName);}
                      }}>
                      <CartesianGrid strokeDasharray="3 3" stroke={t.border}/>
                      <XAxis dataKey="name" tick={{fill:t.muted,fontSize:11}}/>
                      <YAxis tick={{fill:t.muted,fontSize:10}} unit="%" domain={[0,100]}/>
                      <Tooltip content={mkTip}/>
                      <Legend formatter={v=><span style={{color:t.text,fontSize:11}}>{v}</span>}/>
                      <Bar dataKey="A1" name="A1 Normal"    stackId="a" fill={P.a1}><LabelList dataKey="A1" position="center" formatter={v=>v>0?v.toFixed(0)+"%":""} style={{fill:"#fff",fontSize:10,fontWeight:800,paintOrder:"stroke",stroke:"rgba(0,0,0,0.55)",strokeWidth:3}}/></Bar>
                      <Bar dataKey="A2" name="A2 Anomaly"   stackId="a" fill={P.a2}><LabelList dataKey="A2" position="center" formatter={v=>v>0?v.toFixed(0)+"%":""} style={{fill:"#fff",fontSize:10,fontWeight:800,paintOrder:"stroke",stroke:"rgba(0,0,0,0.55)",strokeWidth:3}}/></Bar>
                      <Bar dataKey="A3" name="A3 Incomplete" stackId="a" fill={P.a3} radius={[4,4,0,0]}><LabelList dataKey="A3" position="center" formatter={v=>v>0?v.toFixed(0)+"%":""} style={{fill:"#fff",fontSize:10,fontWeight:800,paintOrder:"stroke",stroke:"rgba(0,0,0,0.55)",strokeWidth:3}}/></Bar>
                    </BarChart>
                  </ResponsiveContainer>
                  <div style={{overflowY:"auto",maxHeight:240}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                      <thead style={{position:"sticky",top:0,background:t.bg,zIndex:1}}>
                        <tr>{[selRegion?"Cluster":"Region","Total Aktivitas","A1%","A2%","A3%"].map(h=>(
                          <th key={h} style={{padding:"0 8px 8px 0",textAlign:"left",fontSize:10,fontWeight:700,color:t.muted,textTransform:"uppercase",letterSpacing:"0.04em",borderBottom:`1px solid ${t.border}`,whiteSpace:"nowrap"}}>{h}</th>
                        ))}</tr>
                      </thead>
                      <tbody>
                        {compData.map((d,i)=>(
                          <tr key={i} style={{borderBottom:`1px solid ${t.border}`,cursor:"pointer"}}
                            onMouseEnter={e=>e.currentTarget.style.opacity="0.7"}
                            onMouseLeave={e=>e.currentTarget.style.opacity="1"}
                            onClick={()=>{if(selRegion){setSelCluster(d.fullName);}else if(regionCodes.length>1){setSelRegion(d.name);}else{setSelRegion(regionCodes[0]);setSelCluster(d.fullName);}}}>
                            <td style={{padding:"8px 8px 8px 0",fontWeight:700,color:t.text}}>{d.name}</td>
                            <td style={{padding:"8px",color:t.muted}}>{fmtK(d.total)}</td>
                            <td style={{padding:"8px",color:P.a1,fontWeight:700}}>{d.A1.toFixed(1)}%</td>
                            <td style={{padding:"8px",color:d.A2>=40?P.investigate:P.a2,fontWeight:700}}>{d.A2.toFixed(1)}%</td>
                            <td style={{padding:"8px",color:P.a3,fontWeight:700}}>{d.A3.toFixed(1)}%</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
              )
            )}


            {/* ── Visit Type list (dulu ada chart bar, dihapus karena skalanya beda jauh antar tipe bikin bar tipis numpuk) ── */}
            {(view.visitTypeData||[]).length>0&&(
            <div>
              <div style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:2}}>Volume per Tipe Kunjungan</div>
              <div style={{fontSize:11,color:t.muted,marginBottom:12}}>Perbandingan status A1/A2/A3 di tiap tipe kunjungan (Regular = rute terjadwal, Ad-Hoc = kunjungan tambahan di luar rute, Consignment = konsinyasi/titip barang)</div>
              <div>
                {(view.visitTypeData||[]).map((d,vi)=>{
                  const dTotal=d.total||1;
                  return(
                  <div key={d.type} onClick={()=>openVtDrill(d.type,null)} style={{padding:"11px 0",borderBottom:vi<(view.visitTypeData||[]).length-1?`1px solid ${t.border}`:"none",cursor:"pointer"}}>
                    <div style={{display:"flex",alignItems:"center",marginBottom:6}}>
                      <div style={{fontWeight:700,fontSize:12,color:t.text,flex:1}}>{d.type}</div>
                      <div style={{fontSize:13,fontWeight:800,color:t.text,flexShrink:0}}>{d.total.toLocaleString()}</div>
                    </div>
                    <div style={{display:"flex",height:6,borderRadius:3,overflow:"hidden",marginBottom:6}}>
                      <div style={{width:pctS(d.A1,dTotal),background:P.a1}}/>
                      <div style={{width:pctS(d.A2,dTotal),background:P.a2}}/>
                      <div style={{width:pctS(d.A3,dTotal),background:P.a3}}/>
                    </div>
                    <div style={{display:"flex",gap:14,fontSize:10,flexWrap:"wrap"}}>
                      <span onClick={e=>{e.stopPropagation();openVtDrill(d.type,"A1 - NORMAL");}} style={{color:P.a1,cursor:"pointer",fontWeight:700}}>A1: {d.A1.toLocaleString()} ({pctS(d.A1,dTotal)})</span>
                      <span onClick={e=>{e.stopPropagation();openVtDrill(d.type,"A2 - ANOMALY");}} style={{color:P.a2,cursor:"pointer",fontWeight:700}}>A2: {d.A2.toLocaleString()} ({pctS(d.A2,dTotal)})</span>
                      <span onClick={e=>{e.stopPropagation();openVtDrill(d.type,"A3 - INCOMPLETE");}} style={{color:P.a3,cursor:"pointer",fontWeight:700}}>A3: {d.A3.toLocaleString()} ({pctS(d.A3,dTotal)})</span>
                    </div>
                  </div>
                );})}
              </div>
            </div>
            )}

          </div>
        )}

        {/* ════ TREND ════ */}
        {tab==="trend"&&(
          <div>
            <div style={{marginBottom:28}}>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:2,flexWrap:"wrap"}}>
                <span style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase"}}>Ringkasan Tren</span>
                <div style={{display:"flex",gap:16,marginLeft:"auto",flexWrap:"wrap"}}>
                  {[["daily","Harian"],["weekly","Mingguan"],["monthly","Bulanan"],["quarterly","Kuartalan"],["half","Semesteran"],["yearly","Tahunan"]].map(([val,lbl])=>(
                    <button key={val} onClick={()=>setTrendPeriod(val)}
                      style={{background:"none",border:"none",color:trendPeriod===val?P.accent:t.muted,padding:0,fontSize:11,fontWeight:700,cursor:"pointer"}}>
                      {lbl}
                    </button>
                  ))}
                </div>
              </div>
              <div style={{fontSize:11,color:t.muted,marginBottom:16}}>Planned Visit Date · {view.label}</div>

              {(()=>{
                const trendData=groupTrend(view.trend,trendPeriod);
                if(!trendData.length) return <div style={{color:t.muted,fontSize:12,padding:"20px 0"}}>Belum ada data tren.</div>;

                const n=trendData.length;
                const avgPerDay=Math.round(trendData.reduce((s,d)=>s+d.total,0)/n);

                // Bagi 2 paruh berdasarkan urutan waktu (paruh awal & paruh akhir)
                const mid=Math.ceil(n/2);
                const firstHalf=trendData.slice(0,mid);
                const secondHalf=trendData.slice(mid);
                const avgRate=(arr,key)=>arr.length?arr.reduce((s,d)=>s+pct(d[key],d.total),0)/arr.length:0;
                const deltas=["A1","A2","A3"].map(key=>{
                  const r1=avgRate(firstHalf,key), r2=avgRate(secondHalf,key);
                  return {key,r1,r2,delta:r2-r1};
                });
                const rangeLabel=arr=>arr.length?(arr[0].name+"–"+arr[arr.length-1].name):"";

                // Hari menonjol (dicari dari data harian asli, bukan hasil grouping periode, biar presisi ke tanggal)
                const daily=[...view.trend];
                const bestA1=daily.length?[...daily].sort((a,b)=>pct(b.A1,b.total)-pct(a.A1,a.total))[0]:null;
                const worstA1=daily.length?[...daily].sort((a,b)=>pct(a.A1,a.total)-pct(b.A1,b.total))[0]:null;
                const worstA3=daily.length?[...daily].sort((a,b)=>pct(b.A3,b.total)-pct(a.A3,a.total))[0]:null;
                const topVolume=daily.length?[...daily].sort((a,b)=>b.total-a.total)[0]:null;

                const dNames={A1:"A1 Normal",A2:"A2 Anomaly",A3:"A3 Incomplete"};
                const dColors={A1:P.a1,A2:P.a2,A3:P.a3};
                const isGood=(key,delta)=>key==="A1"?delta>=0:delta<=0; // A1 naik=bagus, A2/A3 naik=jelek

                return(<>
                  <div style={{display:"flex",border:`1px solid ${t.border}`,borderRadius:12,overflow:"hidden",marginBottom:18}}>
                    <div style={{flex:1,padding:"14px 8px",textAlign:"center",borderRight:`1px solid ${t.border}`}}>
                      <div style={{fontSize:17,fontWeight:800,color:P.accent}}>{avgPerDay.toLocaleString()}</div>
                      <div style={{fontSize:9,color:t.muted,marginTop:3,textTransform:"uppercase"}}>Rata-rata Aktivitas<br/>per Hari</div>
                    </div>
                    <div style={{flex:1,padding:"14px 8px",textAlign:"center",borderRight:`1px solid ${t.border}`}}>
                      <div style={{fontSize:17,fontWeight:800,color:P.accent}}>{view.trend.length}</div>
                      <div style={{fontSize:9,color:t.muted,marginTop:3,textTransform:"uppercase"}}>Hari Terpantau<br/>dalam Periode Ini</div>
                    </div>
                    <div style={{flex:1,padding:"14px 8px",textAlign:"center"}}>
                      <div style={{fontSize:17,fontWeight:800,color:isGood("A1",deltas[0].delta)?P.a1:"#ef4444"}}>{deltas[0].delta>=0?"+":""}{deltas[0].delta.toFixed(0)}%</div>
                      <div style={{fontSize:9,color:t.muted,marginTop:3,textTransform:"uppercase"}}>Perubahan<br/>Rate A1</div>
                    </div>
                  </div>

                  <div style={{border:`1px solid ${t.border}`,borderRadius:12,padding:"14px 16px",marginBottom:18}}>
                    <div style={{fontSize:10,color:t.muted,marginBottom:10}}>Metode: rata-rata <b style={{color:t.text}}>rate harian</b> di paruh awal periode ({rangeLabel(firstHalf)}) dibandingkan paruh akhir ({rangeLabel(secondHalf)})</div>
                    {deltas.map((d,i)=>(
                      <div key={d.key} style={{marginBottom:i<2?10:0,fontSize:12,color:t.text,lineHeight:1.6}}>
                        {isGood(d.key,d.delta)?"✅":"⚠️"} <b style={{color:isGood(d.key,d.delta)?P.a1:"#ef4444"}}>Rate {dNames[d.key]} {isGood(d.key,d.delta)?"membaik":"memburuk"} {d.delta>=0?"+":""}{d.delta.toFixed(0)}%</b> — rata-rata {d.r1.toFixed(0)}% di paruh awal, jadi rata-rata {d.r2.toFixed(0)}% di paruh akhir.
                      </div>
                    ))}
                  </div>

                  <div style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:2}}>Hari Menonjol</div>
                  <div style={{fontSize:11,color:t.muted,marginBottom:12}}>Klik buat lihat breakdown canvasser di tanggal itu</div>
                  {[
                    bestA1&&{icon:"🏆",color:P.a1,title:"Rate A1 Normal tertinggi (hari ini)",date:bestA1.date,val:pctS(bestA1.A1,bestA1.total)},
                    worstA1&&{icon:"📉",color:"#ef4444",title:"Rate A1 Normal terendah (hari ini)",date:worstA1.date,val:pctS(worstA1.A1,worstA1.total)},
                    worstA3&&{icon:"🔵",color:P.a3,title:"Rate A3 Incomplete tertinggi (hari ini)",date:worstA3.date,val:pctS(worstA3.A3,worstA3.total)},
                    topVolume&&{icon:"📦",color:P.a2,title:"Jumlah aktivitas terbanyak (hari ini)",date:topVolume.date,val:topVolume.total.toLocaleString()},
                  ].filter(Boolean).map((h,i,arr)=>(
                    <div key={i} onClick={()=>setTrendDrill(h.date)} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 0",borderBottom:i<arr.length-1?`1px solid ${t.border}`:"none",cursor:"pointer"}}
                      onMouseEnter={e=>e.currentTarget.style.opacity="0.75"} onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
                      <div style={{width:26,height:26,borderRadius:7,background:h.color+"1f",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,flexShrink:0}}>{h.icon}</div>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontSize:10,color:t.muted}}>{h.title}</div>
                        <div style={{fontSize:12,fontWeight:700,color:t.text}}>{h.date}</div>
                      </div>
                      <div style={{fontSize:14,fontWeight:800,color:h.color,flexShrink:0}}>{h.val}</div>
                    </div>
                  ))}
                </>);
              })()}
            </div>
            {(()=>{const tRev=[...view.trend].reverse();return(
            <div>
              <div style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:2}}>Detail per Tanggal</div>
              <div style={{fontSize:11,color:t.muted,marginBottom:12}}>Klik tanggal untuk lihat breakdown canvasser</div>
              {tRev.slice(tPg*TPG,(tPg+1)*TPG).map((d,i)=>(
                <div key={i} onClick={()=>setTrendDrill(d.date)} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 0",borderBottom:`1px solid ${t.border}`,cursor:"pointer"}}
                  onMouseEnter={e=>e.currentTarget.style.opacity="0.75"} onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
                  <div style={{width:72,fontSize:12,fontWeight:700,color:t.text,flexShrink:0}}>{d.date}</div>
                  <div style={{flex:1,display:"flex",height:8,borderRadius:4,overflow:"hidden",minWidth:60}}>
                    <div style={{width:pctS(d.A1,d.total),background:P.a1}}/>
                    <div style={{width:pctS(d.A2,d.total),background:P.a2}}/>
                    <div style={{width:pctS(d.A3,d.total),background:P.a3}}/>
                  </div>
                  <div style={{display:"flex",gap:10,fontSize:10,color:t.muted,flexShrink:0}}>
                    <span style={{color:P.a1,fontWeight:700}}>{pctS(d.A1,d.total)}</span>
                    <span style={{color:P.a2,fontWeight:700}}>{pctS(d.A2,d.total)}</span>
                    <span style={{color:P.a3,fontWeight:700}}>{pctS(d.A3,d.total)}</span>
                  </div>
                  <div style={{fontSize:13,fontWeight:800,color:t.text,width:56,textAlign:"right",flexShrink:0}}>{d.total.toLocaleString()}</div>
                </div>
              ))}
              <div style={{marginTop:10}}><Pagination page={tPg} setPage={setTPg} total={tRev.length} pageSize={TPG} t={t}/></div>
            </div>
            );})()}
          </div>
        )}

        {/* ════ OUTLET ════ */}
        {tab==="outlet"&&(
          <div>
            <div style={{marginBottom:28}}>
              <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:2}}>
                <span style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase"}}>Volume per Outlet Type</span>
                <div style={{display:"flex",gap:14,marginLeft:"auto"}}>
                  {["Jumlah","Persentase"].map((lbl,mi)=>(
                    <button key={mi} onClick={()=>setOutletChartMode(mi)} style={{background:"none",border:"none",color:outletChartMode===mi?P.accent:t.muted,padding:0,fontSize:11,fontWeight:700,cursor:"pointer"}}>{lbl}</button>
                  ))}
                </div>
              </div>
              <ResponsiveContainer width="100%" height={260}>
                <BarChart data={view.outletData.map(d=>outletChartMode===0
                  ?{name:d.type.replace("RO ",""),A1:d.A1,A2:d.A2,A3:d.A3,_type:d.type}
                  :{name:d.type.replace("RO ",""),A1:pct(d.A1,d.total),A2:pct(d.A2,d.total),A3:pct(d.A3,d.total),_type:d.type}
                )} margin={{top:10,right:10,bottom:20,left:0}}
                onClick={d=>{if(d?.activePayload?.[0]?.payload?._type)openOutletDrill(d.activePayload[0].payload._type);}}>
                  <CartesianGrid strokeDasharray="3 3" stroke={t.border}/>
                  <XAxis dataKey="name" tick={{fill:t.muted,fontSize:11}}/>
                  <YAxis tick={{fill:t.muted,fontSize:10}} tickFormatter={outletChartMode===0?fmtK:v=>v+"%"} unit={outletChartMode===1?"%":""} domain={outletChartMode===1?[0,100]:undefined}/>
                  <Tooltip content={mkTip}/>
                  <Legend formatter={v=><span style={{color:t.text,fontSize:11}}>{v}</span>}/>
                  <Bar dataKey="A1" name="A1 Normal"    stackId="a" fill={P.a1}><LabelList dataKey="A1" position="center" formatter={v=>v>0?(outletChartMode===0?fmtK(v):v.toFixed(0)+"%"):""} style={{fill:"#fff",fontSize:9,fontWeight:800,paintOrder:"stroke",stroke:"rgba(0,0,0,0.55)",strokeWidth:3}}/></Bar>
                  <Bar dataKey="A2" name="A2 Anomaly"   stackId="a" fill={P.a2}><LabelList dataKey="A2" position="center" formatter={v=>v>0?(outletChartMode===0?fmtK(v):v.toFixed(0)+"%"):""} style={{fill:"#fff",fontSize:9,fontWeight:800,paintOrder:"stroke",stroke:"rgba(0,0,0,0.55)",strokeWidth:3}}/></Bar>
                  <Bar dataKey="A3" name="A3 Incomplete" stackId="a" fill={P.a3} radius={[4,4,0,0]}><LabelList dataKey="A3" position="center" formatter={v=>v>0?(outletChartMode===0?fmtK(v):v.toFixed(0)+"%"):""} style={{fill:"#fff",fontSize:9,fontWeight:800,paintOrder:"stroke",stroke:"rgba(0,0,0,0.55)",strokeWidth:3}}/></Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>

            {/* ── Outlet per Region ── */}
            {regionCodes.length>0&&(
            <div style={{marginBottom:28}}>
              <div style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:6}}>Perbandingan per Region</div>
              {(()=>{const rows=regionCodes.map(code=>{const ra=regionAgg[code]||{actC:{},total:0};return{code,a1p:pct((ra.actC||{})["A1 - NORMAL"],ra.total)};});
                const best=[...rows].sort((a,b)=>b.a1p-a.a1p)[0], worst=[...rows].sort((a,b)=>a.a1p-b.a1p)[0];
                return best&&worst&&rows.length>1?(
                <div style={{fontSize:11,color:t.muted,marginBottom:14}}>
                  💡 Region <b style={{color:t.text}}>{best.code}</b> paling sehat (<span style={{color:P.a1,fontWeight:700}}>{best.a1p}% A1</span>), <b style={{color:t.text}}>{worst.code}</b> paling perlu perhatian (<span style={{color:P.a1,fontWeight:700}}>{worst.a1p}% A1</span>).
                </div>):null;})()}

              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={regionCodes.map((code,i)=>{
                  const ra=regionAgg[code]||{actC:{},total:0};
                  return {
                    name:code,
                    fullName:regionFullName(code),
                    A1:(ra.actC||{})["A1 - NORMAL"]||0,
                    A2:(ra.actC||{})["A2 - ANOMALY"]||0,
                    A3:(ra.actC||{})["A3 - INCOMPLETE"]||0,
                    total:ra.total||0,
                  };
                })} margin={{top:5,right:10,bottom:5,left:0}}
                  onClick={d=>{if(d?.activePayload?.[0]?.payload?.name){setSelRegion(d.activePayload[0].payload.name);setSelCluster(null);setTab("overview");}}}>
                  <CartesianGrid strokeDasharray="3 3" stroke={t.border}/>
                  <XAxis dataKey="name" tick={{fill:t.muted,fontSize:10}} tickLine={false}/>
                  <YAxis tick={{fill:t.muted,fontSize:10}} tickLine={false} axisLine={false} tickFormatter={v=>v>=1000?Math.round(v/1000)+"K":v}/>
                  <Tooltip contentStyle={{background:t.card,border:`1px solid ${t.border}`,borderRadius:8,fontSize:11}}
                    formatter={(v,n,p)=>[v.toLocaleString()+" ("+pctS(v,p.payload.total)+")",n]}/>
                  <Legend iconSize={8} wrapperStyle={{fontSize:10}}/>
                  <Bar dataKey="A1" name="A1 Normal" fill={P.a1} stackId="a"><LabelList dataKey="A1" position="center" formatter={v=>v>0?fmtK(v):""} style={{fill:"#fff",fontSize:9,fontWeight:800,paintOrder:"stroke",stroke:"rgba(0,0,0,0.55)",strokeWidth:3}}/></Bar>
                  <Bar dataKey="A2" name="A2 Anomaly" fill={P.a2} stackId="a"><LabelList dataKey="A2" position="center" formatter={v=>v>0?fmtK(v):""} style={{fill:"#fff",fontSize:9,fontWeight:800,paintOrder:"stroke",stroke:"rgba(0,0,0,0.55)",strokeWidth:3}}/></Bar>
                  <Bar dataKey="A3" name="A3 Incomplete" fill={P.a3} stackId="a" radius={[3,3,0,0]}><LabelList dataKey="A3" position="center" formatter={v=>v>0?fmtK(v):""} style={{fill:"#fff",fontSize:9,fontWeight:800,paintOrder:"stroke",stroke:"rgba(0,0,0,0.55)",strokeWidth:3}}/></Bar>
                </BarChart>
              </ResponsiveContainer>
              <div style={{fontSize:10,color:t.muted,marginTop:8,marginBottom:16,textAlign:"center"}}>💡 Klik bar untuk drill-down ke region tersebut</div>

              {regionCodes.map((code,i)=>{
                const ra=regionAgg[code]||{actC:{},total:0};
                const rA1=(ra.actC||{})["A1 - NORMAL"]||0;
                const rA2=(ra.actC||{})["A2 - ANOMALY"]||0;
                const rA3=(ra.actC||{})["A3 - INCOMPLETE"]||0;
                return(
                  <div key={code} onClick={()=>{setSelRegion(code);setSelCluster(null);setTab("overview");}} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 0",borderBottom:i<regionCodes.length-1?`1px solid ${t.border}`:"none",cursor:"pointer"}}>
                    <div style={{width:36,fontSize:10,fontWeight:800,color:t.muted,flexShrink:0}}>{code}</div>
                    <div style={{flex:1,minWidth:0,fontSize:12,fontWeight:600,color:t.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{regionFullName(code)}</div>
                    <div style={{fontSize:11,color:t.muted,width:56,textAlign:"right",flexShrink:0}}>{fmtK(ra.total||0)}</div>
                    <div style={{fontSize:12,fontWeight:800,color:P.a1,width:52,textAlign:"right",flexShrink:0}}>{pctS(rA1,ra.total)}</div>
                    <div style={{fontSize:11,fontWeight:700,color:pct(rA2,ra.total)>=30?P.a2:t.muted,width:48,textAlign:"right",flexShrink:0}}>{pctS(rA2,ra.total)}</div>
                    <div style={{fontSize:11,fontWeight:700,color:t.muted,width:48,textAlign:"right",flexShrink:0}}>{pctS(rA3,ra.total)}</div>
                  </div>
                );
              })}
            </div>
            )}

            <div style={{marginBottom:28}}>
              <div style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:6}}>Detail per Outlet Type</div>
              {(()=>{const worst=[...view.outletData].sort((a,b)=>pct(a.A1,a.total)-pct(b.A1,b.total))[0];const best=[...view.outletData].sort((a,b)=>pct(b.A1,b.total)-pct(a.A1,a.total))[0];
                return worst&&best?(
                <div style={{fontSize:11,color:t.muted,marginBottom:12}}>
                  💡 <b style={{color:t.text}}>{best.type.replace("RO ","")}</b> paling sehat (<span style={{color:P.a1,fontWeight:700}}>{pctS(best.A1,best.total)} A1</span>), <b style={{color:t.text}}>{worst.type.replace("RO ","")}</b> paling perlu perhatian (<span style={{color:P.a1,fontWeight:700}}>{pctS(worst.A1,worst.total)} A1</span>).
                </div>):null;})()}
              <div style={{overflowX:"auto",scrollbarWidth:"none",msOverflowStyle:"none"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif"}}>
                  <thead><tr>
                    {["Outlet Type","Total","A1","A2","A3","Inv","Dist","A1%","A2%"].map(h=><th key={h} style={{padding:isMobile?"0 8px 8px 0":"0 12px 8px 0",textAlign:"left",fontSize:10,fontWeight:700,color:t.muted,textTransform:"uppercase",letterSpacing:"0.04em",whiteSpace:"nowrap",borderBottom:`1px solid ${t.border}`}}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {view.outletData.map((d,i)=>(
                      <tr key={i} style={{borderBottom:`1px solid ${t.border}`}}>
                        <td style={{padding:isMobile?"8px 8px 8px 0":"9px 12px 9px 0",fontWeight:700,color:t.text,cursor:"pointer",whiteSpace:"nowrap"}} onClick={()=>openOutletDrill(d.type)}>{d.type}</td>
                        <td style={{padding:"9px 12px 9px 0",fontWeight:700,color:t.muted}}>{d.total.toLocaleString()}</td>
                        <td style={{padding:"9px 12px 9px 0"}}>
                          {(d.A1||0)>0?<span onClick={()=>setOutletTypeDrill({type:d.type,status:"A1 - NORMAL",label:"A1 Normal"})} style={{color:P.a1,fontWeight:700,cursor:"pointer"}}>{(d.A1||0).toLocaleString()}</span>:<span style={{color:t.muted}}>0</span>}
                        </td>
                        <td style={{padding:"9px 12px 9px 0"}}>
                          {(d.A2||0)>0?<span onClick={()=>setOutletTypeDrill({type:d.type,status:"A2 - ANOMALY",label:"A2 Anomaly"})} style={{color:P.a2,fontWeight:700,cursor:"pointer"}}>{(d.A2||0).toLocaleString()}</span>:<span style={{color:t.muted}}>0</span>}
                        </td>
                        <td style={{padding:"9px 12px 9px 0"}}>
                          {(d.A3||0)>0?<span onClick={()=>setOutletTypeDrill({type:d.type,status:"A3 - INCOMPLETE",label:"A3 Incomplete"})} style={{color:P.a3,fontWeight:700,cursor:"pointer"}}>{(d.A3||0).toLocaleString()}</span>:<span style={{color:t.muted}}>0</span>}
                        </td>
                        <td style={{padding:"9px 12px 9px 0"}}>
                          {(d.INVESTIGATE||0)>0?<span onClick={()=>setOutletTypeDrill({type:d.type,status:"INVESTIGATE",label:"Investigate"})} style={{color:t.muted,fontWeight:600,cursor:"pointer"}}>{(d.INVESTIGATE||0).toLocaleString()}</span>:<span style={{color:t.muted}}>0</span>}
                        </td>
                        <td style={{padding:"9px 12px 9px 0",minWidth:90}}><Bar3 A1={d.A1||0} A2={d.A2||0} A3={d.A3||0} total={d.total}/></td>
                        <td style={{padding:"9px 12px 9px 0",color:P.a1,fontWeight:800}}>{pctS(d.A1,d.total)}</td>
                        <td style={{padding:"9px 0",color:pct(d.A2,d.total)>=40?P.investigate:t.muted,fontWeight:600}}>{pctS(d.A2,d.total)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* ── Census vs Non-Census ── */}
            {(view.censusData||[]).length>0&&(
            <div>
              <div style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:14}}>Census vs Non-Census</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))",gap:20}}>
                {(view.censusData||[]).map((d,i)=>{
                  const col=i===0?P.a1:"#6366f1";
                  return(
                  <div key={i}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                      <div style={{fontWeight:700,fontSize:13,color:col}}>{d.type}</div>
                      <div style={{fontSize:13,fontWeight:800,color:t.text}}>{d.total.toLocaleString()}</div>
                    </div>
                    <Bar3 A1={d.A1||0} A2={d.A2||0} A3={d.A3||0} total={d.total}/>
                    <div style={{display:"flex",gap:16,marginTop:10}}>
                      {[{l:"A1 Normal",v:d.A1||0,c:P.a1},{l:"A2 Anomaly",v:d.A2||0,c:P.a2},{l:"A3 Incomplete",v:d.A3||0,c:P.a3}].map((s,j)=>(
                        <div key={j}>
                          <div style={{fontSize:13,fontWeight:800,color:s.c}}>{pctS(s.v,d.total)}</div>
                          <div style={{fontSize:9,color:t.muted,marginTop:1}}>{s.l}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                );})}
              </div>
            </div>
            )}
          </div>
        )}
        {tab==="detail"&&(
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))",gap:"24px 40px"}}>
            {[
              {title:"Duration Status",icon:"⏱",data:[{l:"Normal",c:P.normal,v:dc["NORMAL"]||0,dk:"DUR_NORMAL"},{l:"Short",c:P.short,v:dc["SHORT"]||0,dk:"DUR_SHORT"},{l:"Long",c:P.long,v:dc["LONG"]||0,dk:"DUR_LONG"}]},
              {title:"Distance Status",icon:"📍",data:[{l:"Near",c:P.near,v:di["NEAR"]||0,dk:"DIS_NEAR"},{l:"Mid",c:P.mid,v:di["MID"]||0,dk:"DIS_MID"},{l:"Far",c:P.far,v:di["FAR"]||0,dk:"DIS_FAR"},{l:"Incomplete",c:P.a3,v:di["INCOMPLETE"]||0,dk:"DIS_INC"}]},
              {title:"Location Status",icon:"📌",data:[{l:"Match",c:P.match,v:lc["MATCH"]||0,dk:"LOC_MATCH"},{l:"Not Match",c:P.notmatch,v:lc["NOT MATCH"]||0,dk:"LOC_NOTMATCH"},{l:"Incomplete",c:P.a3,v:lc["INCOMPLETE"]||0,dk:"LOC_INC"}]},
              {title:"In Range Status",icon:"🎯",data:[{l:"In Range",c:P.a1,v:(view.inRangeC||{})["YES"]||0,dk:"IR_YES"},{l:"Out of Range",c:P.investigate,v:(view.inRangeC||{})["NO"]||0,dk:"IR_NO"}]},
            ].map((sec,si)=>{
              const hasData = sec.data.some(d=>d.v>0);
              const secTotal = sec.data.reduce((s,d)=>s+d.v,0)||1;
              return(
              <div key={si}>
                <div style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:10}}>{sec.icon} {sec.title}</div>
                {hasData?(
                  sec.data.map((d,i)=>(
                    <div key={i} onClick={()=>d.dk&&openDrill(sec.title+" · "+d.l,d.c,d.dk)}
                      style={{padding:"8px 0",borderBottom:i<sec.data.length-1?`1px solid ${t.border}`:"none",cursor:d.dk?"pointer":"default"}}>
                      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:5}}>
                        <div style={{width:7,height:7,borderRadius:"50%",background:d.c,flexShrink:0}}/>
                        <span style={{fontSize:12,color:t.text,fontWeight:600,flex:1}}>{d.l}</span>
                        <span style={{fontSize:10,color:t.muted}}>{d.v.toLocaleString()}</span>
                        <span style={{fontSize:12,fontWeight:800,color:d.c,minWidth:40,textAlign:"right"}}>{pctS(d.v,secTotal)}</span>
                      </div>
                      <div style={{height:3,borderRadius:99,background:t.border}}>
                        <div style={{width:pctS(d.v,secTotal),height:"100%",borderRadius:99,background:d.c}}/>
                      </div>
                    </div>
                  ))
                ):(
                  <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:140,color:t.muted,fontSize:12,gap:6}}>
                    <span style={{fontSize:28,opacity:0.3}}>📭</span>
                    <span>Data tidak tersedia</span>
                    <span style={{fontSize:10,opacity:0.6}}>Kolom tidak ada di file ini</span>
                  </div>
                )}
              </div>
            );})}

            <div style={{gridColumn:"1/-1",marginTop:8}}>
              <div style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase",marginBottom:10}}>🔍 Visit Status Breakdown</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:16}}>
                {VIS.map((s,i)=>(
                  <div key={i} onClick={()=>openDrill(s.label,s.color,s.key)} style={{cursor:"pointer"}}>
                    <div style={{fontSize:11,color:t.muted}}>{s.label}</div>
                    <div style={{fontSize:20,fontWeight:800,color:s.color,marginTop:2}}>{pctS(vc[s.key],T)}</div>
                    <div style={{fontSize:10,color:t.muted,marginTop:1}}>{(vc[s.key]||0).toLocaleString()} aktivitas</div>
                    <div style={{height:3,borderRadius:99,background:t.border,marginTop:6}}>
                      <div style={{width:pct(vc[s.key],T)+"%",height:"100%",borderRadius:99,background:s.color}}/>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}


        {/* ════ CANVASSER ════ */}
        {tab==="canvasser"&&(()=>{
          return(<div>
            <div style={{display:"flex",gap:16,marginBottom:10,flexWrap:"wrap",alignItems:"center"}}>
              <span style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase"}}>Sort:</span>
              {[["Jumlah A2","A2"],["% A2","a2p"],["Jumlah A3","A3"],["% A3","a3p"],["Jumlah Inv","INVESTIGATE"],["Total","total"]].map(([label,key])=>(
                <button key={key} onClick={()=>handleSort(key)}
                  style={{background:"none",border:"none",color:sk===key?P.accent:t.muted,padding:0,fontSize:11,fontWeight:700,cursor:"pointer"}}>
                  {label}{sk===key?(sd==="desc"?" ↓":" ↑"):""}
                </button>
              ))}
            </div>
            <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
              <input placeholder="🔍 Cari nama / cluster..." value={search} onChange={e=>setSearch(e.target.value)}
                style={{background:t.inputBg,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"8px 14px",fontSize:12,outline:"none",width:210}}/>
              {[
                {id:"all",label:"Semua"},{id:"high_a2",label:"A2 ≥30%"},
                {id:"high_a3",label:"A3 ≥20%"},{id:"top_a1",label:"A1 ≥80%"},{id:"inv",label:"Inv ≥5%"},
              ].map(f=>(
                <button key={f.id} onClick={()=>setFq(f.id)} style={{padding:"6px 12px",borderRadius:8,border:`1px solid ${fq===f.id?P.accent:t.border}`,cursor:"pointer",fontSize:11,fontWeight:700,background:fq===f.id?P.accent+"18":"none",color:fq===f.id?P.accent:t.muted}}>{f.label}</button>
              ))}
              <span style={{marginLeft:"auto",fontSize:11,color:t.muted}}>{sorted.length} canvasser</span>
            </div>
            <div style={{overflowX:"auto",scrollbarWidth:"none",msOverflowStyle:"none"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif"}}>
                <thead><tr>
                  {[["#",""],["Region","region"],["Cluster","cluster"],["Nama","name"],["Total","total"],
                    ["A1","A1"],["A2","A2"],["A3","A3"],["Inv","INVESTIGATE"],
                    ["A1%","a1p"],["A2%","a2p"],["A3%","a3p"],["Inv%","invP"],["Avg Dur","avgDur"],["Avg Dist","avgDis"]
                  ].map(([label,key])=>(
                    <th key={label} onClick={()=>key&&handleSort(key)} style={{...ths(key),borderBottom:`1px solid ${t.border}`}}>
                      {label} {key&&sk===key?(sd==="desc"?"↓":"↑"):""}
                    </th>
                  ))}
                </tr></thead>
                <tbody>
                  {sorted.slice(cPg*CPG,(cPg+1)*CPG).map((c,i)=>{
                    return(
                    <tr key={c.name+c.cluster} style={{borderBottom:`1px solid ${t.border}`}}>
                      <td style={{padding:"9px 10px 9px 0",color:t.muted,fontSize:10}}>{i+1}</td>
                      <td style={{padding:"9px 10px",color:t.muted,fontSize:11,fontWeight:600}}>{c.region||"–"}</td>
                      <td style={{padding:"9px 10px",color:t.muted,fontSize:11,whiteSpace:"nowrap",maxWidth:140,overflow:"hidden",textOverflow:"ellipsis"}}>{c.cluster||"–"}</td>
                      <td style={{padding:"9px 10px",fontWeight:700,color:t.text,whiteSpace:"nowrap"}}>{c.name}</td>
                      <td style={{padding:"9px 10px",fontWeight:700,color:t.muted}}>{c.total.toLocaleString()}</td>
                      <td style={{padding:"9px 10px"}}>
                        {c.A1>0
                          ?<span onClick={()=>{const rows=getCanvasserRows(c.name,c.cluster,"A1");setCanvDetail({canvasser:c,drillLabel:"A1 - Normal",color:P.a1,rows,drillKey:"A1",sessionKey:Date.now()});}} style={{color:P.a1,fontWeight:700,cursor:"pointer"}}>{c.A1.toLocaleString()}</span>
                          :<span style={{color:t.muted}}>0</span>}
                      </td>
                      <td style={{padding:"9px 10px"}}>
                        {c.A2>0
                          ?<span onClick={()=>{const rows=getCanvasserRows(c.name,c.cluster,"A2");setCanvDetail({canvasser:c,drillLabel:"A2 - Anomaly",color:P.a2,rows,drillKey:"A2",sessionKey:Date.now()});}} style={{color:P.a2,fontWeight:700,cursor:"pointer"}}>{c.A2.toLocaleString()}</span>
                          :<span style={{color:t.muted}}>0</span>}
                      </td>
                      <td style={{padding:"9px 10px"}}>
                        {c.A3>0
                          ?<span onClick={()=>{const rows=getCanvasserRows(c.name,c.cluster,"A3");setCanvDetail({canvasser:c,drillLabel:"A3 - Incomplete",color:P.a3,rows,drillKey:"A3",sessionKey:Date.now()});}} style={{color:P.a3,fontWeight:700,cursor:"pointer"}}>{c.A3.toLocaleString()}</span>
                          :<span style={{color:t.muted}}>0</span>}
                      </td>
                      <td style={{padding:"9px 10px",color:c.INVESTIGATE>0?P.investigate:t.muted,fontWeight:c.INVESTIGATE>0?700:400}}>{c.INVESTIGATE}</td>
                      <td style={{padding:"9px 10px",color:c.a1p>=70?P.a1:P.a2,fontWeight:700}}>{c.a1p.toFixed(1)}%</td>
                      <td style={{padding:"9px 10px",color:c.a2p>=40?P.investigate:P.a2}}>{c.a2p.toFixed(1)}%</td>
                      <td style={{padding:"9px 10px",color:t.muted}}>{c.a3p.toFixed(1)}%</td>
                      <td style={{padding:"9px 10px",color:c.invP>=5?P.investigate:t.muted,fontWeight:c.invP>=5?700:400}}>{c.invP.toFixed(1)}%</td>
                      <td style={{padding:"9px 10px",color:c.avgDur!=null&&c.avgDur<2?P.short:t.muted}}>{c.avgDur!=null?c.avgDur.toFixed(1)+" m":"—"}</td>
                      <td style={{padding:"9px 0",color:(c.avgDis||0)>500?P.investigate:(c.avgDis||0)>100?P.a2:t.muted}}>{c.avgDis!=null?c.avgDis.toFixed(0)+" m":"—"}</td>
                    </tr>
                  );})}
                </tbody>
              </table>
            </div>
            <div style={{marginTop:10}}><Pagination page={cPg} setPage={setCPg} total={sorted.length} pageSize={CPG} t={t}/></div>
          </div>);})()}

        {tab==="findings"&&(
          <div>
            <div style={{marginBottom:18}}>
              <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",marginBottom:4}}>
                <span style={{fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.05em",textTransform:"uppercase"}}>Temuan & Rekomendasi</span>
                <span style={{background:P.accent+"18",color:P.accent,fontSize:10,fontWeight:800,padding:"2px 9px",borderRadius:999}}>
                  📍 {selCluster||selRegion||"Nasional (semua data)"}
                </span>
              </div>
              <div style={{fontSize:11,color:t.muted}}>Dihasilkan otomatis dari pola di data yang di-upload — bukan pengganti verifikasi manual. Ganti scope lewat tab Overview (klik cluster/region).</div>
            </div>
            {findings.length===0?(
              <div style={{textAlign:"center",color:t.muted,padding:"32px 16px"}}>Belum ada temuan signifikan dari data saat ini. 👍</div>
            ):findings.map((f,i)=>{
              const sevColor=f.severity==="fraud"?"#dc2626":f.severity==="high"?"#ef4444":f.severity==="medium"?P.a2:f.severity==="low"?P.a1:"#3b82f6";
              const sevLabel=f.severity==="fraud"?"FRAUD":f.severity==="high"?"PERHATIAN":f.severity==="medium"?"CEK":f.severity==="low"?"BAIK":"INFO";
              return(
                <div key={i} onClick={f.action||undefined} style={{padding:"16px 0",borderBottom:i<findings.length-1?`1px solid ${t.border}`:"none",cursor:f.action?"pointer":"default"}}
                  onMouseEnter={e=>{if(f.action)e.currentTarget.style.opacity="0.8";}}
                  onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
                  <div style={{display:"flex",alignItems:"center",gap:9,marginBottom:6,flexWrap:"wrap"}}>
                    <span style={{fontSize:16}}>{f.icon}</span>
                    <span style={{fontWeight:700,fontSize:13,color:t.text,flex:1,minWidth:120}}>{f.title}</span>
                    <span style={{background:sevColor,color:"#fff",fontSize:9,fontWeight:800,padding:"3px 9px",borderRadius:999,letterSpacing:"0.03em",whiteSpace:"nowrap"}}>{sevLabel}</span>
                  </div>
                  <div style={{fontSize:12,color:t.muted,lineHeight:1.6,marginBottom:6}}>{f.desc}</div>
                  <div style={{fontSize:11,color:t.text,lineHeight:1.6}}><b style={{color:sevColor}}>Rekomendasi: </b>{f.rec}</div>
                  {f.action&&<div style={{fontSize:10,color:P.accent,fontWeight:700,marginTop:6}}>Lihat detail ›</div>}
                </div>
              );
            })}
          </div>
        )}
      </div>
      <div style={{textAlign:"center",fontSize:10,color:t.muted,padding:"14px 22px 28px",opacity:0.4}}>XLSMART Analytics Dashboard · Klik status di chart untuk lihat breakdown canvasser</div>
    </div>
    <OutletDrillPanel drill={outletDrill} onClose={()=>setOutletDrill(null)} t={t} onDrill={handleOutletActivity}/>
    {trendDrill&&(
      <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:1050,display:"flex",alignItems:"flex-end",background:"rgba(0,0,0,0.65)",backdropFilter:"blur(4px)"}} onClick={()=>setTrendDrill(null)}>
        <div onClick={e=>e.stopPropagation()} style={{width:"100%",maxHeight:"85vh",background:t.card,borderRadius:"20px 20px 0 0",border:`1px solid ${t.border}`,overflow:"hidden",display:"flex",flexDirection:"column",boxShadow:"0 -8px 40px rgba(0,0,0,0.5)",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
          <div style={{padding:"14px 18px 10px",borderBottom:`1px solid ${t.border}`,display:"flex",alignItems:"center",gap:12,flexShrink:0}}>
            <div style={{flex:1}}>
              <div style={{fontWeight:800,fontSize:15,color:t.text}}>📅 {trendDrill}</div>
              <div style={{fontSize:11,color:t.muted,marginTop:1}}>
                {(()=>{const d=view.trend.find(x=>x.date===trendDrill);return d?`${d.total.toLocaleString()} aktivitas · A1: ${pctS(d.A1,d.total)} · A2: ${pctS(d.A2,d.total)} · A3: ${pctS(d.A3,d.total)}`:""})()}
              </div>
            </div>
            <button onClick={()=>setTrendDrill(null)} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:13,fontWeight:700}}>✕</button>
          </div>
          <div style={{overflowY:"auto",flex:1,scrollbarWidth:"none",msOverflowStyle:"none"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
              <thead style={{position:"sticky",top:0,background:t.card,zIndex:1}}>
                <tr style={{background:t.cardAlt}}>
                  {["#","Canvasser","Cluster","A1","A2","A3","Total"].map(h=>(
                    <th key={h} style={{padding:"9px 12px",textAlign:"left",fontSize:11,fontWeight:700,color:t.muted,whiteSpace:"nowrap",borderBottom:`1px solid ${t.border}`}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {(()=>{
                  const dayRows=clusters.flatMap(cl=>(cl.rawRows||[]).filter(r=>{
                    const dt=typeof r["Planned Visit Date"]==="object"?r["Planned Visit Date"]?.toISOString?.()?.slice(0,10):String(r["Planned Visit Date"]||"").slice(0,10);
                    return dt===trendDrill;
                  }));
                  const map={};
                  dayRows.forEach(r=>{
                    const cid=String(r["Canvasser ID"]||r["Canvasser"]||"").trim();
                    const nm=String(r["Canvasser"]||"").trim();
                    const cl=String(r["Cluster"]||"").trim();
                    const as1=r["_CAS1"]||"";
                    if(!map[cid])map[cid]={name:nm,cluster:cl,A1:0,A2:0,A3:0,total:0};
                    map[cid].total++;
                    if(as1==="A1 - NORMAL")map[cid].A1++;else if(as1==="A2 - ANOMALY")map[cid].A2++;else if(as1==="A3 - INCOMPLETE")map[cid].A3++;
                  });
                  return Object.values(map).sort((a,b)=>b.total-a.total).map((r,i)=>(
                    <tr key={i} style={{borderBottom:`1px solid ${t.border}`,background:i%2===0?"transparent":t.rowAlt}}>
                      <td style={{padding:"7px 10px",color:t.muted,fontSize:10}}>{i+1}</td>
                      <td style={{padding:"7px 10px",fontWeight:600,color:t.text}}>{r.name}</td>
                      <td style={{padding:"7px 10px",color:t.muted,fontSize:11}}>{r.cluster}</td>
                      <td style={{padding:"7px 10px",color:r.A1>0?P.a1:t.muted,fontWeight:r.A1>0?700:400}}>{r.A1}</td>
                      <td style={{padding:"7px 10px",color:r.A2>0?P.a2:t.muted,fontWeight:r.A2>0?700:400}}>{r.A2}</td>
                      <td style={{padding:"7px 10px",color:r.A3>0?P.a3:t.muted,fontWeight:r.A3>0?700:400}}>{r.A3}</td>
                      <td style={{padding:"7px 10px",fontWeight:800,color:t.text}}>{r.total}</td>
                    </tr>
                  ));
                })()}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    )}

    {showParams&&(
      <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.7)",backdropFilter:"blur(4px)"}} onClick={()=>setShowParams(false)}>
        <div onClick={e=>e.stopPropagation()} style={{background:t.card,borderRadius:16,border:`1px solid ${t.border}`,padding:"24px 28px",minWidth:320,maxWidth:420,boxShadow:"0 8px 40px rgba(0,0,0,0.5)",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
          <div style={{fontWeight:800,fontSize:16,color:t.text,marginBottom:4}}>⚙️ Parameter Validasi</div>
          <div style={{fontSize:11,color:t.muted,marginBottom:10}}>Perubahan parameter memerlukan persetujuan management</div>
          <div style={{fontSize:11,color:"#f59e0b",background:"#f59e0b15",border:"1px solid #f59e0b40",borderRadius:8,padding:"8px 10px",marginBottom:16,lineHeight:1.6}}>
            ⚠️ Parameter ini <b>TIDAK mengubah status A1/A2/A3</b> (itu sudah ditentukan dari kolom Visit Status di file asli). Parameter ini cuma mempengaruhi breakdown Duration/Distance Status, filter drill-down, dan highlight tabel.
          </div>
          {[
            {key:"dur_short",label:"Durasi Minimal (menit)",desc:"Kunjungan di bawah ini = SHORT",unit:"menit"},
            {key:"dur_long", label:"Durasi Maksimal (menit)",desc:"Kunjungan di atas ini = LONG",unit:"menit"},
            {key:"dis_near", label:"Jarak NEAR (meter)",desc:"Jarak di bawah ini = NEAR (normal)",unit:"meter"},
            {key:"dis_far",  label:"Jarak FAR (meter)",desc:"Jarak di atas ini = FAR (anomali)",unit:"meter"},
          ].map(({key,label,desc,unit})=>(
            <div key={key} style={{marginBottom:16}}>
              <div style={{fontWeight:700,fontSize:12,color:t.text,marginBottom:2}}>{label}</div>
              <div style={{fontSize:10,color:t.muted,marginBottom:6}}>{desc}</div>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <input type="number" value={params[key]} min={0} step={key.startsWith("dur")?1:10}
                  onChange={e=>setParams(p=>({...p,[key]:parseFloat(e.target.value)||0}))}
                  style={{width:80,background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"6px 10px",fontSize:13,fontWeight:700,outline:"none"}}/>
                <span style={{fontSize:11,color:t.muted}}>{unit}</span>
                <span style={{fontSize:10,color:t.muted,marginLeft:4}}>Default: {DEFAULT_PARAMS[key]} {unit}</span>
              </div>
            </div>
          ))}
          <div style={{display:"flex",gap:8,marginTop:8}}>
            <button onClick={()=>setParams({...DEFAULT_PARAMS})} style={{flex:1,background:t.cardAlt,border:`1px solid ${t.border}`,color:t.muted,borderRadius:8,padding:"8px",cursor:"pointer",fontSize:12,fontWeight:700}}>↺ Reset Default</button>
            <button onClick={()=>setShowParams(false)} style={{flex:1,background:P.accent,border:"none",color:"#fff",borderRadius:8,padding:"8px",cursor:"pointer",fontSize:12,fontWeight:700}}>✓ Simpan</button>
          </div>
        </div>
      </div>
    )}

    {/* ── FILE MANAGER PANEL ── */}
    {showFileManager&&(
      <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:2500,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.65)",backdropFilter:"blur(4px)"}} onClick={()=>setShowFileManager(false)}>
        <div onClick={e=>e.stopPropagation()} style={{background:t.card,borderRadius:16,border:`1px solid ${t.border}`,width:"min(560px,95vw)",maxHeight:"80vh",overflow:"hidden",display:"flex",flexDirection:"column",boxShadow:"0 8px 40px rgba(0,0,0,0.5)",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
          <div style={{padding:"16px 20px",borderBottom:`1px solid ${t.border}`,display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontWeight:800,fontSize:15,color:t.text,flex:1}}>📋 File Manager</span>
            <span style={{fontSize:11,color:t.muted}}>{clusters.length} cluster loaded</span>
            <button onClick={()=>setShowFileManager(false)} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"4px 10px",cursor:"pointer",fontSize:12,fontWeight:700}}>✕</button>
          </div>
          <div style={{overflowY:"auto",flex:1,padding:"8px 0"}}>
            {clusters.map((cl,i)=>(
              <div key={cl.label} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 20px",borderBottom:`1px solid ${t.border}`,background:i%2===0?"transparent":t.rowAlt}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:2}}>
                    <span style={{background:cl.color+"22",color:cl.color,padding:"1px 8px",borderRadius:999,fontSize:10,fontWeight:700}}>{cl.regionCode||"?"}</span>
                    <span style={{fontWeight:700,fontSize:12,color:t.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{cl.label}</span>
                  </div>
                  <div style={{fontSize:10,color:t.muted,display:"flex",gap:8}}>
                    {cl.dateRange?.min&&<span>📅 {fmtPeriod(cl.dateRange)}</span>}
                    <span>📊 {(cl.rawRows||[]).length.toLocaleString()} baris</span>
                  </div>
                </div>
                <div style={{display:"flex",gap:6,flexShrink:0}}>
                  <label style={{background:P.accent+"22",color:P.accent,border:`1px solid ${P.accent}40`,borderRadius:7,padding:"4px 10px",fontSize:10,fontWeight:700,cursor:"pointer"}}>
                    🔄 Ganti
                    <input type="file" accept=".xlsx,.xls" style={{display:"none"}} onClick={e=>e.target.value=""} onChange={async e=>{
                      const file=e.target.files?.[0]; if(!file) return;
                      setAddLoading({current:0,total:1,name:file.name.replace(/\.[^.]+$/,"")});
                      const reader=new FileReader();
                      reader.onload=ev=>{
                        try{
                          const wb2=XLSX.read(ev.target.result,{type:"array",cellDates:true,cellHTML:false});
                          const sheets=wb2.SheetNames;
                          const newResults=[];
                          for(const sn of sheets){
                            const ws2=wb2.Sheets[sn]; if(!ws2||!ws2["!ref"]) continue;
                            const {rows}=readFileRows(wb2,sn);
                            if(!rows?.length) continue;
                            const clusterNames=[...new Set(rows.map(r=>r["Cluster"]).filter(Boolean))];
                            if(clusterNames.length>1){
                              clusterNames.forEach(cln=>{
                                const clRows=rows.filter(r=>String(r["Cluster"]||"").trim()===cln);
                                if(clRows.length) newResults.push({name:file.name+"|"+cln,label:cln,regionCode:getRegionCode(cln),rows:clRows,originFile:file});
                              });
                            } else {
                              const label=clusterNames[0]||sn;
                              newResults.push({name:file.name,label,regionCode:getRegionCode(label),rows,originFile:file});
                            }
                          }
                          onAddFiles(prev=>{
                            const m=[...(prev||[])];
                            newResults.forEach(r=>{
                              const reRows=r.rows.map(row=>{
                                const rid=String(row["Outlet ID"]||"").trim();
                                const ro=roMap[rid];
                                return ro?{...row,"RO Latitude":row["RO Latitude"]??ro.lat,"RO Longitude":row["RO Longitude"]??ro.lon,"RO Census":row["RO Census"]??(ro.census?"YES":"NO"),"Outlet Type":row["Outlet Type"]||ro.type}:row;
                              });
                              // REPLACE — find by label and overwrite
                              const byLabel=m.findIndex(x=>x.label===r.label||x.label===cl.label);
                              if(byLabel>=0) m[byLabel]={...r,rows:reRows};
                              else m.push({...r,rows:reRows});
                            });
                            return m;
                          });
                          setAddLoading({current:1,total:1,name:"Selesai!"});
                          setTimeout(()=>{setAddLoading(null);setShowFileManager(false);},1200);
                        }catch(err){setAddLoading(null);alert("Error: "+err.message);}
                      };
                      reader.readAsArrayBuffer(file);
                    }}/>
                  </label>
                  <button onClick={()=>{
                    onAddFiles(prev=>(prev||[]).filter(x=>x.label!==cl.label));
                    if(clusters.length<=1) setShowFileManager(false);
                  }} style={{background:"#ef444422",color:"#ef4444",border:"1px solid #ef444440",borderRadius:7,padding:"4px 10px",fontSize:10,fontWeight:700,cursor:"pointer"}}>
                    🗑
                  </button>
                </div>
              </div>
            ))}
          </div>
          <div style={{padding:"12px 20px",borderTop:`1px solid ${t.border}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontSize:10,color:t.muted}}>🔄 Ganti = replace data cluster tersebut · 🗑 = hapus cluster</span>
          </div>
        </div>
      </div>
    )}
{showFileManager&&(
      <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:2500,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.65)",backdropFilter:"blur(4px)"}} onClick={()=>setShowFileManager(false)}>
        <div onClick={e=>e.stopPropagation()} style={{background:t.card,borderRadius:16,border:`1px solid ${t.border}`,width:"min(560px,95vw)",maxHeight:"80vh",overflow:"hidden",display:"flex",flexDirection:"column",boxShadow:"0 8px 40px rgba(0,0,0,0.5)",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
          <div style={{padding:"16px 20px",borderBottom:`1px solid ${t.border}`,display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontWeight:800,fontSize:15,color:t.text,flex:1}}>📋 File Manager</span>
            <span style={{fontSize:11,color:t.muted}}>{clusters.length} cluster loaded</span>
            <button onClick={()=>setShowFileManager(false)} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"4px 10px",cursor:"pointer",fontSize:12,fontWeight:700}}>✕</button>
          </div>
          <div style={{overflowY:"auto",flex:1,padding:"8px 0"}}>
            {clusters.map((cl,i)=>(
              <div key={cl.label} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 20px",borderBottom:`1px solid ${t.border}`,background:i%2===0?"transparent":t.rowAlt}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:2}}>
                    <span style={{background:cl.color+"22",color:cl.color,padding:"1px 8px",borderRadius:999,fontSize:10,fontWeight:700}}>{cl.regionCode||"?"}</span>
                    <span style={{fontWeight:700,fontSize:12,color:t.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{cl.label}</span>
                  </div>
                  <div style={{fontSize:10,color:t.muted,display:"flex",gap:8,flexWrap:"wrap"}}>
                    {cl.dateRange?.min&&<span>📅 {fmtPeriod(cl.dateRange)}</span>}
                    <span>📊 {(cl.rawRows||[]).length.toLocaleString()} baris</span>
                  </div>
                </div>
                <div style={{display:"flex",gap:6,flexShrink:0}}>
                  <label style={{background:P.accent+"22",color:P.accent,border:`1px solid ${P.accent}40`,borderRadius:7,padding:"4px 10px",fontSize:10,fontWeight:700,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:4}}>
                    🔄 Ganti
                    <input type="file" accept=".xlsx,.xls" style={{display:"none"}} onClick={e=>e.target.value=""} onChange={async e=>{
                      const file=e.target.files?.[0]; if(!file) return;
                      const targetLabel=cl.label;
                      setAddLoading({current:0,total:1,name:file.name.split(".")[0]});
                      const rd=new FileReader();
                      rd.onload=ev=>{
                        try{
                          const wb2=XLSX.read(ev.target.result,{type:"array",cellDates:true,cellHTML:false});
                          const newR=[];
                          for(const sn of wb2.SheetNames){
                            const ws2=wb2.Sheets[sn]; if(!ws2||!ws2["!ref"]) continue;
                            const {rows}=readFileRows(wb2,sn);
                            if(!rows?.length) continue;
                            const cls2=[...new Set(rows.map(r=>r["Cluster"]).filter(Boolean))];
                            if(cls2.length>1){cls2.forEach(c2=>{const r2=rows.filter(r=>String(r["Cluster"]||"").trim()===c2);if(r2.length)newR.push({name:file.name+"|"+c2,label:c2,regionCode:getRegionCode(c2),rows:r2});});}
                            else{const lbl=cls2[0]||sn;newR.push({name:file.name,label:lbl,regionCode:getRegionCode(lbl),rows});}
                          }
                          if(!newR.length) throw new Error("File kosong");
                          onAddFiles(prev=>{
                            const m=[...(prev||[])];
                            newR.forEach(r=>{
                              const reRows=r.rows.map(row=>{const rid=String(row["Outlet ID"]||"").trim();const ro=roMap[rid];return ro?{...row,"RO Latitude":row["RO Latitude"]??ro.lat,"RO Longitude":row["RO Longitude"]??ro.lon}:row;});
                              const byLabel=m.findIndex(x=>x.label===targetLabel||x.label===r.label);
                              if(byLabel>=0)m[byLabel]={...r,rows:reRows};
                              else m.push({...r,rows:reRows});
                            });
                            return m;
                          });
                          setAddLoading({current:1,total:1,name:"Selesai!"});
                          setTimeout(()=>{setAddLoading(null);setShowFileManager(false);},1200);
                        }catch(err){setAddLoading(null);alert("Error: "+err.message);}
                      };
                      rd.readAsArrayBuffer(file);
                    }}/>
                  </label>
                  <button onClick={()=>{onAddFiles(prev=>(prev||[]).filter(x=>x.label!==cl.label));}} style={{background:"#ef444422",color:"#ef4444",border:"1px solid #ef444440",borderRadius:7,padding:"4px 10px",fontSize:10,fontWeight:700,cursor:"pointer"}}>🗑</button>
                </div>
              </div>
            ))}
          </div>
          <div style={{padding:"12px 20px",borderTop:`1px solid ${t.border}`,display:"flex",justifyContent:"space-between",alignItems:"center",gap:8}}>
            <span style={{fontSize:10,color:t.muted}}>🔄 Ganti = replace · 🗑 = hapus cluster</span>
            <label style={{display:"inline-flex",alignItems:"center",gap:6,background:"rgba(34,197,94,0.15)",border:"1px solid rgba(34,197,94,0.4)",color:"#4ade80",borderRadius:8,padding:"6px 14px",cursor:"pointer",fontSize:11,fontWeight:700}}>
              ➕ Add Files
              <input type="file" accept=".xlsx,.xls" multiple style={{display:"none"}} onClick={e=>e.target.value=""} onChange={e=>{setShowFileManager(false);handleAddFiles(e.target.files);}}/>
            </label>
          </div>
        </div>
      </div>
    )}
    {canvCategoryDrill&&(
      <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:1200,display:"flex",alignItems:"flex-end",background:"rgba(0,0,0,0.65)",backdropFilter:"blur(4px)"}} onClick={()=>setCanvCategoryDrill(null)}>
        <div onClick={e=>e.stopPropagation()} style={{width:"100%",maxHeight:"78vh",background:t.card,borderRadius:"18px 18px 0 0",border:"1px solid "+t.border,overflow:"hidden",display:"flex",flexDirection:"column",fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif"}}>
          <div style={{padding:"16px 18px 12px",borderBottom:"1px solid "+t.border,display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:12,height:12,borderRadius:3,background:canvCategoryDrill.color,flexShrink:0}}/>
            <div style={{flex:1}}>
              <div style={{fontWeight:800,fontSize:14,color:t.text}}>{canvCategoryDrill.label}</div>
              <div style={{fontSize:11,color:t.muted,marginTop:2}}>{canvCategoryDrill.list.length.toLocaleString()} canvasser (status dominan)</div>
            </div>
            <button onClick={()=>setCanvCategoryDrill(null)} style={{background:t.cardAlt,border:"1px solid "+t.border,color:t.text,borderRadius:8,padding:"4px 10px",cursor:"pointer",fontSize:12,fontWeight:700}}>✕</button>
          </div>
          <div style={{overflowY:"auto",flex:1,padding:"8px 18px 18px",scrollbarWidth:"none"}}>
            {canvCategoryDrill.list.map((c,i)=>(
              <div key={i} onClick={()=>{const rows=getCanvasserRows(c.name,c.cluster,canvCategoryDrill.statusKey);setCanvDetail({canvasser:c,drillLabel:canvCategoryDrill.label,color:canvCategoryDrill.color,rows,drillKey:canvCategoryDrill.statusKey,sessionKey:Date.now()});setCanvCategoryDrill(null);}}
                style={{display:"flex",alignItems:"center",gap:10,padding:"10px 6px",cursor:"pointer",borderBottom:i<canvCategoryDrill.list.length-1?"1px solid "+t.border:"none"}}
                onMouseEnter={e=>e.currentTarget.style.opacity="0.7"} onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
                <div style={{width:22,height:22,borderRadius:6,background:canvCategoryDrill.color+"22",color:canvCategoryDrill.color,fontSize:10,fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{i+1}</div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontSize:13,fontWeight:700,color:t.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name}</div>
                  <div style={{fontSize:10,color:t.muted,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.cluster}</div>
                </div>
                <div style={{fontSize:13,fontWeight:800,color:canvCategoryDrill.color,flexShrink:0}}>{(c[canvCategoryDrill.statusKey]||0).toLocaleString()}</div>
                <span style={{fontSize:12,color:t.muted,flexShrink:0}}>›</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    )}
    {reasonDrill&&(
      <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:1200,display:"flex",alignItems:"flex-end",background:"rgba(0,0,0,0.65)",backdropFilter:"blur(4px)"}} onClick={()=>setReasonDrill(null)}>
        <div onClick={e=>e.stopPropagation()} style={{width:"100%",maxHeight:"75vh",background:t.card,borderRadius:"18px 18px 0 0",border:"1px solid "+t.border,overflow:"hidden",display:"flex",flexDirection:"column",fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif"}}>
          <div style={{padding:"14px 18px 12px",borderBottom:"1px solid "+t.border,display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:12,height:12,borderRadius:3,background:reasonDrill.color,flexShrink:0}}/>
            <div style={{flex:1}}>
              <div style={{fontWeight:800,fontSize:14,color:t.text}}>{reasonDrill.label}</div>
              <div style={{fontSize:11,color:t.muted,marginTop:1}}>Breakdown penyebab anomali</div>
            </div>
            <button onClick={()=>setReasonDrill(null)} style={{background:t.cardAlt,border:"1px solid "+t.border,color:t.text,borderRadius:8,padding:"4px 10px",cursor:"pointer",fontSize:13,fontWeight:700}}>✕</button>
          </div>
          <div style={{overflowY:"auto",flex:1,padding:"16px 18px",scrollbarWidth:"none"}}>
            {!reasonDrill.reasons.length
              ?<div style={{textAlign:"center",color:t.muted,padding:"24px 0"}}>Data tidak tersedia</div>
              :reasonDrill.reasons.map((r,i)=>{
                const isSel=reasonDrill.selectedReasonIdx===i;
                return(
                <div key={i} onClick={()=>setReasonDrill(prev=>({...prev,selectedReasonIdx:prev.selectedReasonIdx===i?null:i}))}
                  style={{marginBottom:14,cursor:"pointer",padding:"8px",margin:"-8px -8px 6px",borderRadius:10,background:isSel?reasonDrill.color+"15":"transparent",border:isSel?"1px solid "+reasonDrill.color+"50":"1px solid transparent",transition:"all 0.15s"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:5}}>
                    <div style={{display:"flex",alignItems:"center",gap:7}}>
                      <span style={{fontSize:12,fontWeight:i===0?800:600,color:i===0?t.text:t.muted}}>{r.lbl}</span>
                      {i===0&&<span style={{background:reasonDrill.color+"25",color:reasonDrill.color,fontSize:9,fontWeight:700,padding:"1px 7px",borderRadius:999}}>TERBANYAK</span>}
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <span style={{fontSize:13,fontWeight:800,color:reasonDrill.color}}>{r.pct}%</span>
                      <span style={{fontSize:10,color:t.muted}}>({r.cnt.toLocaleString()})</span>
                      <span style={{fontSize:11,color:t.muted}}>{isSel?"▲":"▼"}</span>
                    </div>
                  </div>
                  <div style={{background:t.border,borderRadius:999,height:8}}>
                    <div style={{background:i===0?reasonDrill.color:reasonDrill.color+"77",borderRadius:999,height:8,width:r.pct+"%",transition:"width 0.5s"}}/>
                  </div>
                </div>
              );})
            }
            {(()=>{
              const selIdx=reasonDrill.selectedReasonIdx;
              const activeReason=selIdx!=null?reasonDrill.reasons[selIdx]:null;
              const list=activeReason
                ?(activeReason.top5||[]).map(cv=>({name:cv.name,cluster:cv.cluster,val:cv.count}))
                :(reasonDrill.topCanvassers||[]).map(cv=>({name:cv.name,cluster:cv.cluster,val:cv[reasonDrill.statusKey]||0}));
              const reasonName=activeReason?activeReason.lbl.split(" ").slice(1).join(" "):null;
              if(!list.length) return null;
              return(
              <div style={{marginTop:20,paddingTop:14,borderTop:"1px solid "+t.border}}>
                <div style={{fontSize:11,fontWeight:800,color:t.muted,letterSpacing:"0.05em",marginBottom:10}}>👤 TOP 5 CANVASSER{reasonName?` (${reasonName})`:""}</div>
                {list.map((cv,i)=>(
                  <div key={i} onClick={()=>{const rows=getCanvasserRows(cv.name,cv.cluster,reasonDrill.statusKey);setCanvDetail({canvasser:cv,drillLabel:reasonDrill.label,color:reasonDrill.color,rows,drillKey:reasonDrill.statusKey,sessionKey:Date.now()});setReasonDrill(null);}}
                    style={{display:"flex",alignItems:"center",gap:10,padding:"8px 4px",cursor:"pointer",borderBottom:i<list.length-1?"1px solid "+t.border:"none"}}
                    onMouseEnter={e=>e.currentTarget.style.opacity="0.7"} onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
                    <div style={{width:20,height:20,borderRadius:6,background:reasonDrill.color+"22",color:reasonDrill.color,fontSize:10,fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{i+1}</div>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:12,fontWeight:700,color:t.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{cv.name}</div>
                      <div style={{fontSize:10,color:t.muted,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{cv.cluster}</div>
                    </div>
                    <div style={{fontSize:13,fontWeight:800,color:reasonDrill.color,flexShrink:0}}>{cv.val.toLocaleString()}</div>
                    <span style={{fontSize:12,color:t.muted,flexShrink:0}}>›</span>
                  </div>
                ))}
              </div>
              );
            })()}
          </div>
        </div>
      </div>
    )}
    {addLoading&&(
      <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:3000,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.6)",backdropFilter:"blur(4px)"}}>
        <div style={{background:"#1e293b",borderRadius:16,padding:"28px 36px",minWidth:320,boxShadow:"0 8px 40px rgba(0,0,0,0.6)",fontFamily:"'Segoe UI',system-ui,sans-serif",textAlign:"center"}}>
          <div style={{fontSize:28,marginBottom:12}}>📂</div>
          <div style={{fontWeight:800,fontSize:15,color:"#f1f5f9",marginBottom:6}}>
            {addLoading.current>=addLoading.total?"✅ Selesai!":"Memproses File..."}
          </div>
          <div style={{fontSize:12,color:"#94a3b8",marginBottom:16,maxWidth:260,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
            {addLoading.name}
          </div>
          {/* Progress bar */}
          <div style={{background:"#334155",borderRadius:999,height:8,width:260,margin:"0 auto 10px"}}>
            <div style={{background:P.accent,borderRadius:999,height:8,width:`${Math.round((addLoading.current/addLoading.total)*100)}%`,transition:"width 0.3s ease"}}/>
          </div>
          <div style={{fontSize:11,color:"#64748b"}}>
            {addLoading.current} / {addLoading.total} file
            {addLoading.current>0&&<span style={{marginLeft:6,color:P.accent,fontWeight:600}}>({Math.round((addLoading.current/addLoading.total)*100)}%)</span>}
          </div>
        </div>
      </div>
    )}
    {outletTypeDrill&&(()=>{
      const {type,status,label}=outletTypeDrill;
      const isInv=status==="INVESTIGATE";
      const map={};
      clusters.forEach(cl=>(cl.rawRows||[]).forEach(r=>{
        const ot=String(r["Outlet Type"]||"").trim();
        if(ot!==type) return;
        const as1=r["_CAS1"]||"";
        const vs=String(r["_CVS"]||r["Visit Status"]||"").toUpperCase();
        if(isInv?vs!=="INVESTIGATE":as1!==status) return;
        const cid=String(r["Canvasser ID"]||r["Canvasser"]||"").trim();
        const nm=String(r["Canvasser"]||"").trim();
        if(!map[cid])map[cid]={id:cid,name:nm,region:cl.regionCode||"",cluster:cl.label||"",total:0};
        map[cid].total++;
      }));
      const rows2=Object.values(map).sort((a,b)=>b.total-a.total);
      const color=isInv?P.investigate:status==="A1 - NORMAL"?P.a1:status==="A2 - ANOMALY"?P.a2:P.a3;
      return(
        <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:1050,display:"flex",alignItems:"flex-end",background:"rgba(0,0,0,0.65)",backdropFilter:"blur(4px)"}} onClick={()=>setOutletTypeDrill(null)}>
          <div onClick={e=>e.stopPropagation()} style={{width:"100%",maxHeight:"85vh",background:t.card,borderRadius:"20px 20px 0 0",border:`1px solid ${t.border}`,overflow:"hidden",display:"flex",flexDirection:"column",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
            <div style={{padding:"14px 18px 10px",borderBottom:`1px solid ${t.border}`,display:"flex",alignItems:"center",gap:12}}>
              <div style={{flex:1}}>
                <div style={{fontWeight:800,fontSize:15,color:t.text}}>🏪 {type}</div>
                <div style={{fontSize:11,color:t.muted,marginTop:2,display:"flex",gap:8,alignItems:"center"}}>
                  <span style={{background:color+"20",color,padding:"1px 8px",borderRadius:999,fontSize:10,fontWeight:700}}>{label}</span>
                  <span>{rows2.length} canvasser · {rows2.reduce((s,r)=>s+r.total,0).toLocaleString()} aktivitas</span>
                </div>
              </div>
              <button onClick={()=>setOutletTypeDrill(null)} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:13,fontWeight:700}}>✕</button>
            </div>
            <div style={{overflowY:"auto",flex:1,scrollbarWidth:"none",msOverflowStyle:"none"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead style={{position:"sticky",top:0,background:t.card}}>
                  <tr style={{background:t.cardAlt}}>
                    {["#","Canvasser","Region","Cluster","Jumlah"].map(h=><th key={h} style={{padding:"9px 12px",textAlign:"left",fontSize:11,fontWeight:700,color:t.muted,borderBottom:`1px solid ${t.border}`,whiteSpace:"nowrap"}}>{h}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {rows2.slice(0,50).map((r,i)=>(
                    <tr key={r.id||i} style={{borderBottom:`1px solid ${t.border}`,background:i%2===0?"transparent":t.rowAlt,cursor:"pointer"}}
                      onClick={()=>{const rows3=getCanvasserRows(r.name,r.cluster,isInv?"INVESTIGATE":status==="A1 - NORMAL"?"A1":status==="A2 - ANOMALY"?"A2":"A3");setCanvDetail({canvasser:r,drillLabel:label,color,rows:rows3,drillKey:isInv?"INVESTIGATE":status==="A1 - NORMAL"?"A1":status==="A2 - ANOMALY"?"A2":"A3",sessionKey:Date.now()});}}>
                      <td style={{padding:"7px 10px",color:t.muted,fontSize:10}}>{i+1}</td>
                      <td style={{padding:"7px 10px",fontWeight:600,color:t.text}}>{r.name}</td>
                      <td style={{padding:"7px 10px"}}><span style={{background:P.accent+"20",color:P.accent,padding:"1px 7px",borderRadius:6,fontSize:10,fontWeight:700}}>{r.region||"–"}</span></td>
                      <td style={{padding:"7px 10px",color:t.muted,fontSize:11}}>{r.cluster||"–"}</td>
                      <td style={{padding:"7px 10px",fontWeight:800,color}}>{r.total.toLocaleString()}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      );
    })()}
    <VisitTypeDrillPanel drill={vtDrill} onClose={()=>setVtDrill(null)} t={t}
      onCanvasserClick={(r,key)=>{
        const drillMap={"A1":"A1 - Normal","A2":"A2 - Anomaly","A3":"A3 - Incomplete"};
        const colorMap={"A1":P.a1,"A2":P.a2,"A3":P.a3};
        const rows=getCanvasserRows(r.name,r.cluster,key);
        setCanvDetail({canvasser:r,drillLabel:drillMap[key],color:colorMap[key],rows,drillKey:key,sessionKey:Date.now()});
      }}/>
    <OutletActivityPanel detail={outletActivity} onClose={()=>setOutletActivity(null)} t={t}/>
    <DrillDownPanel drill={drill} onClose={()=>setDrill(null)} t={t}
      onCanvasserClick={(r)=>{
        const rows=getCanvasserRows(r.name,r.cluster,drill.countKey);
        setCanvDetail({canvasser:r,drillLabel:drill.label,color:drill.color,rows,drillKey:drill.countKey,sessionKey:Date.now()});
      }}/>
    <CanvasserDetailPanel detail={canvDetail} onClose={()=>setCanvDetail(null)} t={t}/>
    </>
  );
}

// ── ROOT ──────────────────────────────────────────────────────────────────────
export default function App(){
  const [files,setFiles]=useState(null);
  const [roMap,setRoMap]=useState({});
  const hideScrollbarStyle=`
    ::-webkit-scrollbar{display:none!important;width:0!important;height:0!important}
    *{scrollbar-width:none!important;-ms-overflow-style:none!important}
  `;
  const [dark,setDark]=useState(true);
  const t=dark?DARK:LIGHT;
  return(<><style>{hideScrollbarStyle}</style>{files
    ?<Dashboard files={files} onReset={()=>setFiles(null)} onAddFiles={setFiles} dark={dark} toggleDark={()=>setDark(d=>!d)} roMap={roMap}/>
    :<UploadScreen onLoad={setFiles} roMap={roMap} onRoLoad={setRoMap} t={t}/>}</>);
}