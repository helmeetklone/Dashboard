import { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  PieChart, Pie, Cell, ResponsiveContainer, LineChart, Line } from "recharts";

// ── THEMES ───────────────────────────────────────────────────────────────────
const DARK  = { bg:"#060d1a",card:"#0c1a2e",cardAlt:"#081422",border:"#162840",text:"#ddeeff",muted:"#5a7fa8",inputBg:"#0c1a2e",rowAlt:"rgba(255,255,255,0.02)",rowHover:"rgba(37,99,235,0.08)" };
const LIGHT = { bg:"#f0f4fa",card:"#ffffff",cardAlt:"#f5f8ff",border:"#d0dff0",text:"#0f2040",muted:"#6b8aad",inputBg:"#ffffff",rowAlt:"rgba(0,0,0,0.02)",rowHover:"rgba(37,99,235,0.05)" };


// ── DEFAULT VALIDATION PARAMETERS ────────────────────────────────────────────
const DEFAULT_PARAMS = {
  dur_short: 5,      // menit — di bawah ini = SHORT
  dur_long:  30,     // menit — di atas ini = LONG
  dis_near:  50,     // meter — di bawah ini = NEAR
  dis_far:   200,    // meter — di atas ini = FAR
};

const P = {
  a1:"#22c55e",a2:"#f59e0b",a3:"#3b82f6",
  valid:"#22c55e",observe:"#f59e0b",investigate:"#ef4444",incomplete:"#3b82f6",
  normal:"#22c55e",short:"#f97316",long:"#a855f7",
  near:"#22c55e",mid:"#f59e0b",far:"#ef4444",match:"#22c55e",notmatch:"#ef4444",
  accent:"#2563eb",
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

// ── READ FILE — reads critical cols by direct cell reference (100% reliable) ─────
function readFileRows(buf, sheetName=null) {
  const wb = XLSX.read(buf, {type:"array", cellDates:true});
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
  const total=rows.length;
  const actC={"A1 - NORMAL":0,"A2 - ANOMALY":0,"A3 - INCOMPLETE":0};
  const visC={VALID:0,OBSERVE:0,INVESTIGATE:0,INCOMPLETE:0};
  const durC={NORMAL:0,SHORT:0,LONG:0};
  const disC={NEAR:0,MID:0,FAR:0,INCOMPLETE:0};
  const locC={MATCH:0,"NOT MATCH":0,INCOMPLETE:0};
  const inRangeC={YES:0,NO:0};
  const outMap={},canvMap={},dateMap={},visitMap={},vtMap={};

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
      dur = isNaN(dm) ? "" : dm < 3 ? "SHORT" : dm > 60 ? "LONG" : "NORMAL";
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
        <div style={{overflowY:"auto",flex:1}}>
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
    if(durSt==="SHORT"||(!isNaN(dur)&&dur>0&&dur<3))f.push("⏱ Durasi singkat ("+fmtDur(dur)+")");
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
              <span style={{color:t.muted}}>· {rows.length} aktivitas</span>
            </div>
          </div>
          <button onClick={onClose} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:13,fontWeight:700}}>✕</button>
        </div>
        <div style={{overflowY:"auto",flex:1}}>
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
                  <td style={{padding:"7px 10px",color:!isNaN(dur)&&dur>0&&dur<3?P.short:t.muted}}>
                    {fmtDur(r["Visit Duration (Menit)"])}{!isNaN(dur)&&dur>0&&dur<1&&<span style={{fontSize:9,color:P.investigate,marginLeft:3}}>⚡</span>}
                  </td>
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
        <div style={{overflowY:"auto",flex:1}}>
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
            <div style={{fontSize:11,color:t.muted,marginTop:1}}>{drill.rows.length} canvasser · {drill.total.toLocaleString()} aktivitas</div>
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
        <div style={{overflowY:"auto",flex:1}}>
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
  const [sortCol,setSortCol]=useState("date");
  const [sortDir,setSortDir]=useState("asc");
  const PG=10;
  useEffect(()=>{setPg(0);setOPg(0);setView("list");setVtFilter("ALL");setSortCol("date");setSortDir("asc");},[detail?.canvasser?.id]);
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
    if(durSt==="SHORT"||(!isNaN(dur)&&dur>0&&dur<3))  f.push(`⏱ Durasi singkat (${fmtDur(dur)})`);
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
  const filteredRows=(rows||[]).filter(r=>{
    if(vtFilter==="ALL") return true;
    return String(r["Activity Type"]||"Unknown").trim()===vtFilter;
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
  const outletMap={};
  (allRows||[]).forEach(r=>{
    const oid=String(r["Outlet ID"]||r["Outlet"]||"").trim();
    const onm=String(r["Outlet"]||oid).trim();
    const as1=r["_CAS1"]||"";
    if(!outletMap[oid])outletMap[oid]={id:oid,name:onm,total:0,A1:0,A2:0,A3:0};
    outletMap[oid].total++;
    if(as1==="A1 - NORMAL")outletMap[oid].A1++;
    else if(as1==="A2 - ANOMALY")outletMap[oid].A2++;
    else if(as1==="A3 - INCOMPLETE")outletMap[oid].A3++;
  });
  const outletRows=Object.values(outletMap).sort((a,b)=>b.total-a.total);

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
              <span style={{color:t.muted}}>· {sorted.length} aktivitas sesuai filter</span>
            </div>
            <div style={{display:"flex",gap:4,marginTop:6}}>
              <button onClick={()=>setView("list")} style={{background:view==="list"?color:t.cardAlt,color:view==="list"?"#fff":t.muted,border:"1px solid "+t.border,borderRadius:6,padding:"3px 10px",fontSize:10,fontWeight:700,cursor:"pointer"}}>📋 Aktivitas ({filteredRows.length})</button>
              <button onClick={()=>setView("outlet")} style={{background:view==="outlet"?color:t.cardAlt,color:view==="outlet"?"#fff":t.muted,border:"1px solid "+t.border,borderRadius:6,padding:"3px 10px",fontSize:10,fontWeight:700,cursor:"pointer"}}>🏪 Per Outlet ({outletRows.length})</button>
            </div>
          </div>
          <button onClick={onClose} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:13,fontWeight:700}}>✕</button>
        </div>
           {/* Visit Type Filter */}
          {view==="list"&&(
          <div style={{padding:"8px 16px",borderBottom:`1px solid ${t.border}`,display:"flex",gap:6,flexWrap:"wrap",alignItems:"center",flexShrink:0,background:t.cardAlt}}>
            <span style={{fontSize:10,color:t.muted,fontWeight:600}}>Filter:</span>
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
       <div style={{overflowY:"auto",flex:1}}>
          {view==="outlet"?(<>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
              <thead style={{position:"sticky",top:0,background:t.card,zIndex:1}}>
                <tr style={{background:t.cardAlt}}>
                  {["#","Outlet ID","Outlet","Total","A1","A2","A3"].map(h=>(
                    <th key={h} style={{padding:"9px 12px",textAlign:"left",fontSize:11,fontWeight:700,color:t.muted,whiteSpace:"nowrap",borderBottom:`1px solid ${t.border}`}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {outletRows.slice(oPg*PG,(oPg+1)*PG).map((r,i)=>(
                  <tr key={r.id||i} style={{borderBottom:`1px solid ${t.border}`,background:i%2===0?"transparent":t.rowAlt}}>
                    <td style={{padding:"7px 10px",color:t.muted,fontSize:10}}>{oPg*PG+i+1}</td>
                    <td style={{padding:"7px 10px",color:t.muted,fontSize:10,whiteSpace:"nowrap"}}>{r.id||"–"}</td>
                    <td style={{padding:"7px 10px",fontWeight:600,color:t.text,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.name}</td>
                    <td style={{padding:"7px 10px",fontWeight:800,color:t.text}}>{r.total}</td>
                    <td style={{padding:"7px 10px",color:r.A1>0?P.a1:t.muted,fontWeight:r.A1>0?700:400}}>{r.A1}</td>
                    <td style={{padding:"7px 10px",color:r.A2>0?P.a2:t.muted,fontWeight:r.A2>0?700:400}}>{r.A2}</td>
                    <td style={{padding:"7px 10px",color:r.A3>0?P.a3:t.muted,fontWeight:r.A3>0?700:400}}>{r.A3}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <Pagination page={oPg} setPage={setOPg} total={outletRows.length} pageSize={PG} t={t}/>
          </>):(<>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif"}}>
            <thead style={{position:"sticky",top:0,background:t.card,zIndex:1}}>
              <tr style={{background:t.cardAlt}}>
                {["#","Tanggal","Visit Type","Outlet ID","Outlet","Status","In Range","Jarak ke Outlet*","Durasi","Alasan"].map(h=>(
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
                  <td style={{padding:"7px 10px",color:dur>0&&dur<3?P.short:t.muted,fontWeight:dur>0&&dur<3?700:400}}>
                    <span>{fmtDur(r["Visit Duration (Menit)"])}</span>
                    {dur>0&&dur<1&&<span style={{fontSize:9,color:P.investigate,marginLeft:4,fontWeight:700}}>⚡</span>}
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
        const wb=XLSX.read(buf,{type:"array"});
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
      const xlsFiles=Array.from(files).filter(f=>/\.(xlsx|xls)$/i.test(f.name));
      if(!xlsFiles.length)throw new Error("Tidak ada file .xlsx/.xls ditemukan");
      const results=await Promise.all(xlsFiles.map(f=>new Promise((res,rej)=>{
        const reader=new FileReader();
        reader.onload=e=>{
          try{
            const wb2=XLSX.read(e.target.result,{type:"array",cellDates:true});
            const sheets=wb2.SheetNames;
            const fileResults=[];
            for(const sn of sheets){
              const ws2=wb2.Sheets[sn]; if(!ws2||!ws2["!ref"]) continue;
              const {rows}=readFileRows(e.target.result,sn);
              if(!rows||!rows.length) continue;
              const clusterNames=[...new Set(rows.map(r=>r["Cluster"]).filter(Boolean))];
              const label=clusterNames.length===1?clusterNames[0]:(sheets.length>1?sn:f.name.replace(/\.[^.]+$/,""));
              const regionCode=getRegionCode(clusterNames[0]||sn||"");
              fileResults.push({name:sheets.length>1?f.name+"|"+sn:f.name,label,regionCode,rows});
            }
            if(!fileResults.length) throw new Error(`${f.name}: tidak ada data`);
            res(fileResults.length===1?fileResults[0]:{multi:true,results:fileResults,name:f.name});
          }catch(err){rej(err);}
        };
        reader.readAsArrayBuffer(f);
      })));
      setQueue(prev=>{
        const m=[...prev];
        const skipped=[];
        const flatResults=results.flatMap(r=>r.multi?r.results:[r]);
        flatResults.forEach(r=>{
          const byName=m.findIndex(x=>x.name===r.name);
          const byLabel=m.findIndex(x=>x.label===r.label&&x.name!==r.name);
          if(byName>=0){
            // same filename - silently replace (re-upload same file)
            m[byName]=r;
          } else if(byLabel>=0){
            // different filename but same cluster label - skip & warn
            skipped.push(r.label);
          } else {
            m.push(r);
          }
        });
        if(skipped.length>0){
          setError(`⚠️ Cluster berikut sudah ada, dilewati: ${skipped.join(", ")}`);
          setTimeout(()=>setError(null),4000);
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
        <input ref={fileRef} type="file" accept=".xlsx,.xls" multiple style={{display:"none"}} onChange={e=>handleFiles(e.target.files)}/>
        <input ref={folderRef} type="file" accept=".xlsx,.xls" multiple webkitdirectory="" style={{display:"none"}} onChange={e=>handleFiles(e.target.files)}/>
        {loading
          ?<><div style={{fontSize:40,marginBottom:10}}>⚙️</div><div style={{color:"#60a5fa",fontWeight:700}}>Membaca file...</div></>
          :<>
            <div style={{fontSize:46,marginBottom:10}}>{drag?"📥":"📂"}</div>
            <div style={{color:t.text,fontSize:15,fontWeight:700,marginBottom:4}}>{drag?"Lepas di sini!":"Drag & drop file XLS/XLSX"}</div>
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
function Dashboard({files,onReset,dark,toggleDark,roMap={}}){
  const t=dark?DARK:LIGHT;
  const [params,setParams]=useState({...DEFAULT_PARAMS});
  const [showParams,setShowParams]=useState(false);
  const [selRegion,setSelRegion]=useState(null);
  const [selCluster,setSelCluster]=useState(null);
  const [tab,setTab]=useState("overview");
  const [drill,setDrill]=useState(null);
  const [canvDetail,setCanvDetail]=useState(null);
  const [outletDrill,setOutletDrill]=useState(null);
  const [trendDrill,setTrendDrill]=useState(null);
  const [vtDrill,setVtDrill]=useState(null);
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
  const clusters=useMemo(()=>files.map((f,i)=>{
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
    return {
    ...processRows(reRows),
    rawRows:reRows,  // kept for canvasser detail lookup
    label:f.label,regionCode:f.regionCode,
    color:P.regions[i%P.regions.length],fileName:f.name,
  };}),[files,params]);

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
    label:"Nasional",
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
        case "DUR_SHORT": return durSt==="SHORT"||(!isNaN(dur)&&dur>0&&dur<3);
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
  const ths=key=>({padding:"9px 10px",textAlign:"left",fontSize:11,fontWeight:700,color:sk===key?"#60a5fa":t.muted,cursor:"pointer",letterSpacing:"0.05em",whiteSpace:"nowrap",userSelect:"none",borderBottom:`2px solid ${sk===key?P.accent:t.border}`});

  const T=view.total||1;
  const ac=view.actC||{};
  const vc=view.visC||{};
  const dc=view.durC||{};
  const di=view.disC||{};
  const lc=view.locC||{};

  // Current level label
  const levelLabel=selCluster?"Cluster":selRegion?"Region":"Nasional";
  const compLabel=selCluster?`Outlet Type — ${selCluster}`:selRegion?`Cluster dalam ${selRegion}`:regionCodes.length>1?`Perbandingan Region`:`Cluster dalam Region ${regionCodes[0]||""}`;

  const tabs=[
    {id:"overview",label:"📊 Overview"},
    {id:"trend",label:"📅 Trend"},
    {id:"outlet",label:"🏪 Outlet"},
    {id:"detail",label:"🔬 Status Detail"},
    {id:"canvasser",label:"👤 Canvasser"},
  ];

  return(
    <>
    <div style={{minHeight:"100vh",background:t.bg,color:t.text,fontFamily:"'Segoe UI',system-ui,sans-serif",transition:"background 0.3s"}}>

      {/* ── HEADER ── */}
      <div style={{background:dark?"linear-gradient(135deg,#040e1e,#071830)":t.card,borderBottom:`1px solid ${t.border}`,padding:"14px 22px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10,boxShadow:"0 4px 20px rgba(0,0,0,0.15)"}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <img src="/xlsmart-logo.png" alt="XLSMART" width="38" height="38" style={{objectFit:"contain",flexShrink:0}}/>
          <div>
            <div style={{fontSize:16,fontWeight:800,color:t.text}}>XLSMART <span style={{color:"#3b82f6"}}>Analytics</span></div>
            <div style={{fontSize:10,color:t.muted,letterSpacing:"0.08em",textTransform:"uppercase"}}>Activity Quality Dashboard</div>
          </div>
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          <div style={{fontSize:11,color:t.muted,background:t.cardAlt,border:`1px solid ${t.border}`,borderRadius:8,padding:"5px 10px"}}>
            {clusters.length} cluster · {regionCodes.length} region · {(national.total||0).toLocaleString()} aktivitas
          </div>
          <button onClick={()=>setShowParams(p=>!p)} style={{background:showParams?"#f59e0b22":t.cardAlt,border:"1px solid "+(showParams?"#f59e0b":t.border),color:showParams?"#f59e0b":t.muted,borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:11,fontWeight:700,marginRight:4}}>⚙️ Parameter</button>
          <button onClick={toggleDark} style={{background:t.cardAlt,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"5px 12px",fontSize:13,cursor:"pointer",fontWeight:700}}>
            {dark?"☀️":"🌙"}
          </button>
          <button onClick={onReset} style={{background:"rgba(239,68,68,0.15)",border:"1px solid rgba(239,68,68,0.3)",color:"#f87171",borderRadius:8,padding:"5px 12px",fontSize:11,cursor:"pointer",fontWeight:700}}>↩ Ganti File</button>
        </div>
      </div>

      {/* ── NAVIGATION: Level 1 — Nasional + Region tabs ── */}
      <div style={{background:t.card,borderBottom:`1px solid ${t.border}`}}>
        <div style={{padding:"0 20px",display:"flex",gap:0,overflowX:"auto"}}>
          <button onClick={()=>{setSelRegion(null);setSelCluster(null);}} style={{padding:"10px 16px",border:"none",cursor:"pointer",fontSize:12,fontWeight:700,background:"transparent",whiteSpace:"nowrap",color:!selRegion&&!selCluster?"#3b82f6":t.muted,borderBottom:`3px solid ${!selRegion&&!selCluster?"#3b82f6":"transparent"}`,transition:"all 0.15s"}}>
            🌐 Nasional
          </button>
          <div style={{width:1,background:t.border,margin:"8px 4px",flexShrink:0}}/>
          <span style={{padding:"10px 8px",fontSize:10,color:t.muted,fontWeight:700,letterSpacing:"0.08em",alignSelf:"center"}}>REGION:</span>
          {regionCodes.map((code,i)=>{
            const isActive=selRegion===code&&!selCluster;
            const rc=P.regions[i%P.regions.length];
            const clCount=(regionGroups[code]||[]).length;
            return(
              <button key={code} onClick={()=>{setSelRegion(code);setSelCluster(null);}} style={{padding:"10px 16px",border:"none",cursor:"pointer",fontSize:12,fontWeight:700,background:"transparent",whiteSpace:"nowrap",color:isActive?rc:t.muted,borderBottom:`3px solid ${isActive?rc:"transparent"}`,transition:"all 0.15s",display:"flex",alignItems:"center",gap:5}}>
                <span>{code}</span>
                <span style={{fontSize:10,background:isActive?rc+"22":t.cardAlt,color:isActive?rc:t.muted,padding:"1px 6px",borderRadius:999,fontWeight:600}}>{clCount}</span>
              </button>
            );
          })}
        </div>

        {/* ── Level 2 — Cluster tabs (shown when region selected) ── */}
        {selRegion&&(
          <div style={{padding:"0 20px 0 36px",display:"flex",gap:0,overflowX:"auto",borderTop:`1px solid ${t.border}`,background:t.cardAlt}}>
            <button onClick={()=>setSelCluster(null)} style={{padding:"8px 14px",border:"none",cursor:"pointer",fontSize:11,fontWeight:700,background:"transparent",whiteSpace:"nowrap",color:!selCluster?"#60a5fa":t.muted,borderBottom:`2px solid ${!selCluster?"#60a5fa":"transparent"}`,transition:"all 0.15s"}}>
              ∑ {selRegion} Total
            </button>
            {(regionGroups[selRegion]||[]).map((cl,i)=>{
              const isActive=selCluster===cl.label;
              return(
                <button key={cl.label} onClick={()=>setSelCluster(cl.label)} style={{padding:"8px 14px",border:"none",cursor:"pointer",fontSize:11,fontWeight:700,background:"transparent",whiteSpace:"nowrap",color:isActive?cl.color:t.muted,borderBottom:`2px solid ${isActive?cl.color:"transparent"}`,transition:"all 0.15s"}}>
                  {cl.label.replace(new RegExp(`^${selRegion}[-_ ]?`,"i"),"").trim()||cl.label}
                </button>
              );
            })}
          </div>
        )}
      </div>

      {/* ── Breadcrumb ── */}
      <div style={{padding:"8px 22px",background:t.bg,display:"flex",alignItems:"center",gap:6,fontSize:11,color:t.muted}}>
        <span style={{cursor:"pointer",color:P.accent}} onClick={()=>{setSelRegion(null);setSelCluster(null);}}>🌐 {national.label}</span>
        {selRegion&&<><span>›</span><span style={{cursor:"pointer",color:regionAgg[selRegion]?.color||t.muted}} onClick={()=>setSelCluster(null)}>Region {selRegion}</span></>}
        {selCluster&&<><span>›</span><span style={{color:view.color||t.text}}>{selCluster}</span></>}
        <span style={{marginLeft:"auto",background:`${P.accent}20`,color:P.accent,padding:"2px 10px",borderRadius:999,fontWeight:700,fontSize:10}}>
          {levelLabel} · {T.toLocaleString()} aktivitas
        </span>
      </div>

      <div style={{padding:"14px 22px"}}>

        {/* ── KPI CARDS ── */}
        <div style={{display:"grid",gridTemplateColumns:isMobile?"repeat(3,1fr)":isTablet?"repeat(4,1fr)":"repeat(auto-fit,minmax(120px,1fr))",gap:10,marginBottom:16}}>
          {[
            {label:"Total",val:T.toLocaleString(),icon:"📋",color:P.accent},
            {label:"A1 Normal",val:pctS(ac["A1 - NORMAL"],T),icon:"✅",color:P.a1,sub:(ac["A1 - NORMAL"]||0).toLocaleString(),drill:()=>openDrill("A1 - Normal",P.a1,"A1")},
            {label:"A2 Anomaly",val:pctS(ac["A2 - ANOMALY"],T),icon:"⚠️",color:P.a2,sub:(ac["A2 - ANOMALY"]||0).toLocaleString(),drill:()=>openDrill("A2 - Anomaly",P.a2,"A2")},
            {label:"A3 Incomplete",val:pctS(ac["A3 - INCOMPLETE"],T),icon:"🔵",color:P.a3,sub:(ac["A3 - INCOMPLETE"]||0).toLocaleString(),drill:()=>openDrill("A3 - Incomplete",P.a3,"A3")},
            {label:"Investigate",val:pctS(vc["INVESTIGATE"],T),icon:"🔍",color:P.investigate,sub:(vc["INVESTIGATE"]||0).toLocaleString(),drill:()=>openDrill("Investigate",P.investigate,"INVESTIGATE")},
            {label:"Canvasser",val:view.canvassers.length,icon:"👤",color:"#a78bfa"},

          ].map((k,i)=>(
            <div key={i} onClick={k.drill||undefined} style={{...card({borderTop:`3px solid ${k.color}`,cursor:k.drill?"pointer":"default"}),transition:"transform 0.15s,box-shadow 0.15s"}}
              onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-2px)";if(k.drill)e.currentTarget.style.boxShadow=`0 6px 20px ${k.color}30`;}}
              onMouseLeave={e=>{e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.boxShadow="none";}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div style={{fontSize:18}}>{k.icon}</div>
                {k.drill&&<span style={{fontSize:9,color:t.muted,background:t.cardAlt,padding:"1px 5px",borderRadius:4}}>›</span>}
              </div>
              <div style={{fontSize:19,fontWeight:800,color:k.color}}>{k.val}</div>
              {k.sub&&<div style={{fontSize:10,color:t.muted,marginTop:1}}>{k.sub}</div>}
              <div style={{fontSize:10,color:t.muted,marginTop:4}}>{k.label}</div>
            </div>
          ))}
        </div>

        {/* ── TAB BUTTONS ── */}
        <div style={{display:"flex",gap:6,marginBottom:16,flexWrap:"wrap"}}>
          {tabs.map(tb=>(
            <button key={tb.id} onClick={()=>setTab(tb.id)} style={{padding:"7px 14px",borderRadius:9,border:`1px solid ${tab===tb.id?"transparent":t.border}`,cursor:"pointer",fontSize:11,fontWeight:700,transition:"all 0.15s",background:tab===tb.id?"linear-gradient(135deg,#1d5fc0,#2d8ef5)":t.card,color:tab===tb.id?"#fff":t.muted,boxShadow:tab===tb.id?"0 4px 14px rgba(29,95,192,0.3)":"none"}}>{tb.label}</button>
          ))}
        </div>

        {/* ════ OVERVIEW ════ */}
        {tab==="overview"&&(
          <div style={{display:"grid",gap:16}}>

            {/* ── ROW 1: Activity Status Pie (left) + Key Insights (right) ── */}
            <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"minmax(240px,1fr) minmax(300px,1.5fr)",gap:16}}>
              <div style={card()}>
                <div style={{fontWeight:700,marginBottom:2}}>Activity Status</div>
                <div style={{fontSize:11,color:t.muted,marginBottom:10}}>{view.label}</div>
                <ResponsiveContainer width="100%" height={170}>
                  <PieChart>
                    <Pie data={ACT.map(s=>({name:s.label,value:ac[s.key]||0,color:s.color,skey:s.key}))} cx="50%" cy="50%" outerRadius={75} innerRadius={32} dataKey="value" strokeWidth={0}
                      onClick={d=>{const m={"A1 - NORMAL":"A1","A2 - ANOMALY":"A2","A3 - INCOMPLETE":"A3"};openDrill(d.name,d.color,m[d.skey]);}}>
                      {ACT.map((s,i)=><Cell key={i} fill={s.color} style={{cursor:"pointer"}}/>)}
                    </Pie>
                    <Tooltip content={mkTip}/>
                  </PieChart>
                </ResponsiveContainer>
                {ACT.map((s,i)=>{
                  const aKey={"A1 - NORMAL":"A1","A2 - ANOMALY":"A2","A3 - INCOMPLETE":"A3"}[s.key];
                  return(
                  <div key={i} onClick={()=>openDrill(s.label,s.color,aKey)} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 10px",background:t.cardAlt,borderRadius:8,marginBottom:5,cursor:"pointer",transition:"opacity 0.15s"}}
                    onMouseEnter={e=>e.currentTarget.style.opacity="0.75"}
                    onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
                    <div style={{width:10,height:10,borderRadius:3,background:s.color,flexShrink:0}}/>
                    <span style={{fontSize:11,color:t.muted,flex:1}}>{s.label}</span>
                    <span style={{fontSize:12,fontWeight:700,color:s.color}}>{(ac[s.key]||0).toLocaleString()}</span>
                    <span style={{fontSize:11,color:t.muted,minWidth:44,textAlign:"right"}}>{pctS(ac[s.key],T)}</span>
                    <span style={{fontSize:10,color:t.muted}}>›</span>
                  </div>
                );})}
                <div style={{marginTop:10,padding:"8px 0 4px",fontSize:11,color:t.muted,fontWeight:700,letterSpacing:"0.06em",borderTop:`1px solid ${t.border}`}}>VISIT STATUS</div>
                <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":isTablet?"1fr":"1fr 1fr",gap:5,marginTop:4}}>
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
                <div style={{fontWeight:700,marginBottom:12}}>📋 Key Insights</div>
                {(()=>{
                  // Sort by AMOUNT (count) not percentage - more meaningful
                  const wA2=[...view.canvassers].filter(c=>c.A2>0).sort((a,b)=>b.A2-a.A2)[0];
                  const wA3=[...view.canvassers].filter(c=>c.A3>0).sort((a,b)=>b.A3-a.A3)[0];
                  const wInv=[...view.canvassers].filter(c=>c.INVESTIGATE>0).sort((a,b)=>b.INVESTIGATE-a.INVESTIGATE)[0];
                  const topDay=[...(view.trend||[])].sort((a,b)=>b.total-a.total)[0];
                  const openInsight=(canv,key,label,color)=>{
                    if(!canv) return;
                    const rows=getCanvasserRows(canv.name,canv.cluster,key);
                    setCanvDetail({canvasser:canv,drillLabel:label,color,rows,drillKey:key});
                  };
                  return[
                    {icon:"✅",color:P.a1,title:"A1 Normal Rate",desc:`${pctS(ac["A1 - NORMAL"],T)} (${(ac["A1 - NORMAL"]||0).toLocaleString()}) aktivitas berjalan normal.`},
                    {icon:"⚠️",color:P.a2,title:"A2 Anomaly Terbanyak",onClick:wA2?()=>openInsight(wA2,"A2","A2 - Anomaly",P.a2):null,desc:wA2?`${wA2.name} [${wA2.cluster}]: ${wA2.A2.toLocaleString()} aktivitas (${wA2.a2p.toFixed(1)}% dari totalnya)`:"–"},
                    {icon:"🔵",color:P.a3,title:"A3 Incomplete Terbanyak",onClick:wA3?()=>openInsight(wA3,"A3","A3 - Incomplete",P.a3):null,desc:wA3?`${wA3.name} [${wA3.cluster}]: ${wA3.A3.toLocaleString()} aktivitas (${wA3.a3p.toFixed(1)}% dari totalnya)`:"–"},
                    {icon:"🔍",color:P.investigate,title:"Investigate Terbanyak",onClick:wInv?()=>openInsight(wInv,"INVESTIGATE","Investigate",P.investigate):null,desc:wInv?`${wInv.name} [${wInv.cluster}]: ${wInv.INVESTIGATE.toLocaleString()} aktivitas (${wInv.invP.toFixed(1)}% dari totalnya)`:"–"},
                    {icon:"⏱",color:"#f97316",title:"Durasi Singkat (SHORT)",desc:`${(dc["SHORT"]||0).toLocaleString()} aktivitas durasi singkat — perlu verifikasi.`},
                    ...((view.visitTypeData||[]).map(vt=>({
                      icon:vt.type==="Regular Visit"?"🚗":vt.type==="Ad-Hoc Visit"?"⚡":"📦",
                      color:vt.type==="Regular Visit"?P.a1:vt.type==="Ad-Hoc Visit"?P.a2:"#06b6d4",
                      title:vt.type,
                      onClick:()=>openVtDrill(vt.type,null),
                      desc:`Total: ${vt.total.toLocaleString()} · A1: ${vt.A1.toLocaleString()} · A2: ${vt.A2.toLocaleString()} · A3: ${vt.A3.toLocaleString()}`,
                    }))),
                    {icon:"📅",color:"#34d399",title:"Hari Tersibuk",onClick:topDay?()=>setTrendDrill(topDay.date):null,desc:topDay?`${topDay.date}: ${topDay.total.toLocaleString()} aktivitas (A1: ${pctS(topDay.A1,topDay.total)})`:"–"},
                  ].map((f,i)=>(
                    <div key={i} onClick={f.onClick} style={{display:"flex",gap:10,padding:"9px 12px",cursor:f.onClick?"pointer":"default",transition:"background 0.15s",background:f.onClick?"transparent":"transparent",borderRadius:10,background:t.cardAlt,border:`1px solid ${t.border}`,marginBottom:8,alignItems:"flex-start"}}>
                      <div style={{fontSize:14,width:28,height:28,borderRadius:7,background:f.color+"20",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{f.icon}</div>
                      <div>
                        <div style={{fontWeight:700,fontSize:11,color:f.color}}>{f.title}</div>
                        <div style={{fontSize:11,color:t.muted,marginTop:2,lineHeight:1.5}}>{f.desc}</div>
                      </div>
                    </div>
                  ));
                })()}
              </div>
            </div>

            {/* ── ROW 2: Comparison chart (region/cluster) OR Top 5 per status (cluster level) ── */}
            {selCluster?(
              // Cluster level: Top 5 canvassers per A1, A2, A3
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:14}}>
                {[
                  {label:"🏆 Top 5 A1 Normal",color:P.a1,key:"A1",sort:"A1"},
                  {label:"⚠️ Top 5 A2 Anomaly",color:P.a2,key:"A2",sort:"A2"},
                  {label:"🔵 Top 5 A3 Incomplete",color:P.a3,key:"A3",sort:"A3"},
                ].map((cat,ci)=>{
                  const top5=[...view.canvassers].sort((a,b)=>b[cat.sort]-a[cat.sort]).slice(0,5);
                  const maxV=top5[0]?.[cat.sort]||1;
                  return(
                  <div key={ci} style={card()}>
                    <div style={{fontWeight:700,fontSize:12,marginBottom:12,color:cat.color}}>{cat.label}</div>
                    {top5.map((cv,i)=>(
                      <div key={i} style={{marginBottom:10}}>
                        <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}>
                          <span style={{fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:"70%"}}>{cv.name}</span>
                          <span style={{fontWeight:700,color:cat.color}}>{(cv[cat.sort]||0).toLocaleString()}</span>
                        </div>
                        <div style={{height:5,borderRadius:3,background:t.border}}>
                          <div style={{width:pct(cv[cat.sort]||0,maxV)+"%",height:"100%",borderRadius:3,background:cat.color}}/>
                        </div>
                        <div style={{fontSize:10,color:t.muted,marginTop:1}}>{cv.cluster} · {pctS(cv[cat.sort]||0,cv.total)}</div>
                      </div>
                    ))}
                    <button onClick={()=>openDrill(cat.label,cat.color,cat.key)} style={{width:"100%",background:cat.color+"15",border:`1px solid ${cat.color}40`,color:cat.color,borderRadius:8,padding:"6px",fontSize:11,fontWeight:700,cursor:"pointer",marginTop:4}}>
                      Lihat Semua ›
                    </button>
                  </div>
                );})}
              </div>
            ):(
              // National/Region level: comparison chart
              compData.length>0&&(
              <div style={card()}>
                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14,flexWrap:"wrap",gap:8}}>
                  <div>
                    <div style={{fontWeight:700}}>{selRegion?"📊 Perbandingan Cluster":regionCodes.length>1?"📊 Perbandingan Region":"📊 Perbandingan Cluster"}</div>
                    <div style={{fontSize:11,color:t.muted,marginTop:2}}>{compLabel} — klik bar untuk drill-down</div>
                  </div>
                  <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                    {compData.map((d,i)=>(
                      <div key={i} style={{display:"flex",alignItems:"center",gap:4,cursor:"pointer",padding:"3px 8px",borderRadius:8,background:t.cardAlt,border:`1px solid ${t.border}`,fontSize:11,fontWeight:600}}
                        onClick={()=>{if(selRegion){setSelCluster(d.fullName);}else if(regionCodes.length>1){setSelRegion(d.name);}else{setSelRegion(regionCodes[0]);setSelCluster(d.fullName);}}}>
                        <div style={{width:7,height:7,borderRadius:2,background:d.color,flexShrink:0}}/>
                        <span style={{color:t.text}}>{d.name}</span>
                      </div>
                    ))}
                  </div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1.5fr 1fr",gap:16}}>
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
                      <Bar dataKey="A1" name="A1 Normal"    stackId="a" fill={P.a1}/>
                      <Bar dataKey="A2" name="A2 Anomaly"   stackId="a" fill={P.a2}/>
                      <Bar dataKey="A3" name="A3 Incomplete" stackId="a" fill={P.a3} radius={[4,4,0,0]}/>
                    </BarChart>
                  </ResponsiveContainer>
                  <div style={{overflowY:"auto",maxHeight:240}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                      <thead style={{position:"sticky",top:0,background:t.card,zIndex:1}}>
                        <tr>{[selRegion?"Cluster":"Region","Total","A1%","A2%","A3%"].map(h=>(
                          <th key={h} style={{padding:"6px 8px",textAlign:"left",fontWeight:700,color:t.muted,borderBottom:`1px solid ${t.border}`,whiteSpace:"nowrap"}}>{h}</th>
                        ))}</tr>
                      </thead>
                      <tbody>
                        {compData.map((d,i)=>(
                          <tr key={i} style={{borderBottom:`1px solid ${t.border}`,cursor:"pointer"}}
                            onMouseEnter={e=>e.currentTarget.style.background=d.color+"18"}
                            onMouseLeave={e=>e.currentTarget.style.background="transparent"}
                            onClick={()=>{if(selRegion){setSelCluster(d.fullName);}else if(regionCodes.length>1){setSelRegion(d.name);}else{setSelRegion(regionCodes[0]);setSelCluster(d.fullName);}}}>
                            <td style={{padding:"7px 8px",fontWeight:700,color:d.color}}>{d.name}</td>
                            <td style={{padding:"7px 8px",color:t.muted}}>{fmtK(d.total)}</td>
                            <td style={{padding:"7px 8px",color:P.a1,fontWeight:600}}>{d.A1.toFixed(1)}%</td>
                            <td style={{padding:"7px 8px",color:d.A2>=40?P.investigate:P.a2}}>{d.A2.toFixed(1)}%</td>
                            <td style={{padding:"7px 8px",color:P.a3}}>{d.A3.toFixed(1)}%</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
              )
            )}


            {/* ── Visit Type chart ── */}
            {(view.visitTypeData||[]).length>0&&(
            <div style={card()}>
              <div style={{fontWeight:700,fontSize:13,marginBottom:4}}>📊 Volume per Visit Type</div>
              <div style={{fontSize:11,color:t.muted,marginBottom:10}}>Perbandingan A1 · A2 · A3 per tipe kunjungan</div>
              <ResponsiveContainer width="100%" height={180}>
                <BarChart data={(view.visitTypeData||[]).map(d=>({name:d.type.replace(" Visit",""),A1:d.A1,A2:d.A2,A3:d.A3,_type:d.type}))} margin={{top:5,right:10,bottom:5,left:0}}
                  onClick={d=>{if(d?.activePayload?.[0]?.payload?._type){const p=d.activePayload[0];const sk=p.dataKey;const sfMap={"A1":"A1 - NORMAL","A2":"A2 - ANOMALY","A3":"A3 - INCOMPLETE"};openVtDrill(p.payload._type,sfMap[sk]||null);}}}>
                  <CartesianGrid strokeDasharray="3 3" stroke={t.border}/>
                  <XAxis dataKey="name" tick={{fill:t.muted,fontSize:11}} tickLine={false}/>
                  <YAxis tick={{fill:t.muted,fontSize:10}} tickLine={false} axisLine={false}/>
                  <Tooltip contentStyle={{background:t.card,border:`1px solid ${t.border}`,borderRadius:8,fontSize:11}}/>
                  <Legend iconSize={8} wrapperStyle={{fontSize:10}}/>
                  <Bar dataKey="A1" name="A1 Normal" fill={P.a1} stackId="a"/>
                  <Bar dataKey="A2" name="A2 Anomaly" fill={P.a2} stackId="a"/>
                  <Bar dataKey="A3" name="A3 Incomplete" fill={P.a3} stackId="a" radius={[3,3,0,0]}/>
                </BarChart>
              </ResponsiveContainer>
              <div style={{display:"flex",gap:8,marginTop:10,flexWrap:"wrap"}}>
                {(view.visitTypeData||[]).map(d=>(
                  <div key={d.type} onClick={()=>openVtDrill(d.type,null)} style={{flex:1,minWidth:100,background:t.cardAlt,borderRadius:8,padding:"8px 12px",border:`1px solid ${t.border}`,cursor:"pointer"}}>
                    <div style={{fontWeight:700,fontSize:11,color:t.text,marginBottom:2}}>{d.type}</div>
                    <div style={{fontSize:10,color:t.muted}}>Total: <b style={{color:t.text}}>{d.total.toLocaleString()}</b></div>
                    <div style={{display:"flex",gap:6,marginTop:3,fontSize:10,flexWrap:"wrap"}}>
                      <span onClick={e=>{e.stopPropagation();openVtDrill(d.type,"A1 - NORMAL");}} style={{color:P.a1,cursor:"pointer",borderBottom:"1px dotted "+P.a1}}>A1: {d.A1.toLocaleString()}</span>
                      <span onClick={e=>{e.stopPropagation();openVtDrill(d.type,"A2 - ANOMALY");}} style={{color:P.a2,cursor:"pointer",borderBottom:"1px dotted "+P.a2}}>A2: {d.A2.toLocaleString()}</span>
                      <span onClick={e=>{e.stopPropagation();openVtDrill(d.type,"A3 - INCOMPLETE");}} style={{color:P.a3,cursor:"pointer",borderBottom:"1px dotted "+P.a3}}>A3: {d.A3.toLocaleString()}</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>
            )}

          </div>
        )}

        {/* ════ TREND ════ */}
        {tab==="trend"&&(
          <div style={{display:"grid",gap:16}}>
            <div style={card()}>
              <div style={{fontWeight:700,marginBottom:4}}>Volume Aktivitas per Hari</div>
              <div style={{fontSize:11,color:t.muted,marginBottom:14}}>Planned Visit Date · {view.label}</div>
              <ResponsiveContainer width="100%" height={270}>
                <BarChart data={view.trend.map(d=>({name:d.date.slice(5),A1:d.A1,A2:d.A2,A3:d.A3,_date:d.date}))} margin={{top:10,right:10,bottom:20,left:0}}
                  onClick={d=>{if(d?.activePayload?.[0]?.payload?._date)setTrendDrill(d.activePayload[0].payload._date);}}>
                  <CartesianGrid strokeDasharray="3 3" stroke={t.border}/>
                  <XAxis dataKey="name" tick={{fill:t.muted,fontSize:10}}/>
                  <YAxis tick={{fill:t.muted,fontSize:10}} tickFormatter={fmtK}/>
                  <Tooltip content={mkTip}/>
                  <Legend formatter={v=><span style={{color:t.text,fontSize:11}}>{v}</span>}/>
                  <Bar dataKey="A1" name="A1 Normal"    stackId="a" fill={P.a1}/>
                  <Bar dataKey="A2" name="A2 Anomaly"   stackId="a" fill={P.a2}/>
                  <Bar dataKey="A3" name="A3 Incomplete" stackId="a" fill={P.a3} radius={[3,3,0,0]}/>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div style={card()}>
              <div style={{fontWeight:700,marginBottom:4}}>Tren Rate per Hari (%)</div>
              <ResponsiveContainer width="100%" height={200}>
                <LineChart data={view.trend.map(d=>({name:d.date.slice(5),a1:pct(d.A1,d.total),a2:pct(d.A2,d.total),a3:pct(d.A3,d.total)}))} margin={{top:10,right:10,bottom:20,left:0}}>
                  <CartesianGrid strokeDasharray="3 3" stroke={t.border}/>
                  <XAxis dataKey="name" tick={{fill:t.muted,fontSize:10}}/>
                  <YAxis tick={{fill:t.muted,fontSize:10}} unit="%" domain={[0,100]}/>
                  <Tooltip content={mkTip}/>
                  <Legend formatter={v=><span style={{color:t.text,fontSize:11}}>{v}</span>}/>
                  <Line type="monotone" dataKey="a1" name="A1 %" stroke={P.a1} strokeWidth={2} dot={{r:3}}/>
                  <Line type="monotone" dataKey="a2" name="A2 %" stroke={P.a2} strokeWidth={2} dot={{r:3}}/>
                  <Line type="monotone" dataKey="a3" name="A3 %" stroke={P.a3} strokeWidth={2} dot={{r:3}}/>
                </LineChart>
              </ResponsiveContainer>
            </div>
            {(()=>{const tRev=[...view.trend].reverse();return(
            <div style={card()}>
              <div style={{fontWeight:700,marginBottom:12}}>Detail per Tanggal</div>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif"}}>
                  <thead><tr style={{background:t.cardAlt}}>
                    {["Tanggal","Total","A1","A2","A3","A1%","A2%","A3%"].map(h=><th key={h} style={{padding:"9px 12px",textAlign:"left",fontSize:11,fontWeight:700,color:t.muted,whiteSpace:"nowrap",borderBottom:`1px solid ${t.border}`}}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {tRev.slice(tPg*TPG,(tPg+1)*TPG).map((d,i)=>(
                      <tr key={i} style={{borderBottom:`1px solid ${t.border}`,background:i%2===0?"transparent":t.rowAlt}}>
                        <td style={{padding:"8px 12px",fontWeight:700,color:"#3b82f6"}}>{d.date}</td>
                        <td style={{padding:"8px 12px",fontWeight:600}}>{d.total.toLocaleString()}</td>
                        <td style={{padding:"8px 12px",color:P.a1}}>{(d.A1||0).toLocaleString()}</td>
                        <td style={{padding:"8px 12px",color:P.a2}}>{(d.A2||0).toLocaleString()}</td>
                        <td style={{padding:"8px 12px",color:P.a3}}>{(d.A3||0).toLocaleString()}</td>
                        <td style={{padding:"8px 12px",color:P.a1,fontWeight:700}}>{pctS(d.A1,d.total)}</td>
                        <td style={{padding:"8px 12px",color:pct(d.A2,d.total)>=40?P.investigate:P.a2}}>{pctS(d.A2,d.total)}</td>
                        <td style={{padding:"8px 12px",color:P.a3}}>{pctS(d.A3,d.total)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <Pagination page={tPg} setPage={setTPg} total={tRev.length} pageSize={TPG} t={t}/>
            </div>
            );})()}
          </div>
        )}

        {/* ════ OUTLET ════ */}
        {tab==="outlet"&&(
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":isTablet?"1fr":"1fr 1fr",gap:16}}>
            {["Volume","Persentase"].map((title,mi)=>(
              <div key={mi} style={card()}>
                <div style={{fontWeight:700,marginBottom:14}}>{title} per Outlet Type</div>
                <ResponsiveContainer width="100%" height={250}>
                  <BarChart data={view.outletData.map(d=>mi===0
                    ?{name:d.type.replace("RO ",""),A1:d.A1,A2:d.A2,A3:d.A3,_type:d.type}
                    :{name:d.type.replace("RO ",""),A1:pct(d.A1,d.total),A2:pct(d.A2,d.total),A3:pct(d.A3,d.total),_type:d.type}
                  )} margin={{top:10,right:10,bottom:20,left:0}}
                  onClick={d=>{if(d?.activePayload?.[0]?.payload?._type)openOutletDrill(d.activePayload[0].payload._type);}}>
                    <CartesianGrid strokeDasharray="3 3" stroke={t.border}/>
                    <XAxis dataKey="name" tick={{fill:t.muted,fontSize:11}}/>
                    <YAxis tick={{fill:t.muted,fontSize:10}} tickFormatter={mi===0?fmtK:v=>v+"%"} unit={mi===1?"%":""} domain={mi===1?[0,100]:undefined}/>
                    <Tooltip content={mkTip}/>
                    <Legend formatter={v=><span style={{color:t.text,fontSize:11}}>{v}</span>}/>
                    <Bar dataKey="A1" name="A1 Normal"    stackId="a" fill={P.a1}/>
                    <Bar dataKey="A2" name="A2 Anomaly"   stackId="a" fill={P.a2}/>
                    <Bar dataKey="A3" name="A3 Incomplete" stackId="a" fill={P.a3} radius={[4,4,0,0]}/>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            ))}
            <div style={{...card(),gridColumn:"1/-1"}}>
              <div style={{fontWeight:700,marginBottom:12}}>Detail per Outlet Type</div>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif"}}>
                  <thead><tr style={{background:t.cardAlt}}>
                    {["Outlet Type","Total","A1","A2","A3","Investigate","Distribusi","A1%","A2%","A3%"].map(h=><th key={h} style={{padding:"9px 12px",textAlign:"left",fontSize:11,fontWeight:700,color:t.muted,whiteSpace:"nowrap",borderBottom:`1px solid ${t.border}`}}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {view.outletData.map((d,i)=>(
                      <tr key={i} style={{borderBottom:`1px solid ${t.border}`,background:i%2===0?"transparent":t.rowAlt}}>
                        <td style={{padding:"9px 12px",fontWeight:700,color:P.accent,cursor:"pointer"}} onClick={()=>openOutletDrill(d.type)}>{d.type}</td>
                        <td style={{padding:"9px 12px",fontWeight:600}}>{d.total.toLocaleString()}</td>
                        <td style={{padding:"9px 12px",color:P.a1}}>{(d.A1||0).toLocaleString()}</td>
                        <td style={{padding:"9px 12px",color:P.a2}}>{(d.A2||0).toLocaleString()}</td>
                        <td style={{padding:"9px 12px",color:P.a3}}>{(d.A3||0).toLocaleString()}</td>
                        <td style={{padding:"9px 12px",color:(d.INVESTIGATE||0)>0?P.investigate:t.muted}}>{(d.INVESTIGATE||0).toLocaleString()}</td>
                        <td style={{padding:"9px 12px",minWidth:90}}><Bar3 A1={d.A1||0} A2={d.A2||0} A3={d.A3||0} total={d.total}/></td>
                        <td style={{padding:"9px 12px",color:P.a1,fontWeight:700}}>{pctS(d.A1,d.total)}</td>
                        <td style={{padding:"9px 12px",color:pct(d.A2,d.total)>=40?P.investigate:P.a2}}>{pctS(d.A2,d.total)}</td>
                        <td style={{padding:"9px 12px",color:P.a3}}>{pctS(d.A3,d.total)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* ── Census vs Non-Census ── */}
            {(view.censusData||[]).length>0&&(
            <div style={{...card(),gridColumn:"1/-1"}}>
              <div style={{fontWeight:700,marginBottom:14}}>🏘 Census vs Non-Census</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))",gap:14}}>
                {(view.censusData||[]).map((d,i)=>{
                  const col=i===0?"#22c55e":"#6366f1";
                  return(
                  <div key={i} style={{background:t.cardAlt,border:`1px solid ${t.border}`,borderRadius:12,padding:16,borderLeft:`4px solid ${col}`}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                      <div style={{fontWeight:700,fontSize:14,color:col}}>{d.type}</div>
                      <div style={{fontSize:13,fontWeight:800}}>{d.total.toLocaleString()}</div>
                    </div>
                    <Bar3 A1={d.A1||0} A2={d.A2||0} A3={d.A3||0} total={d.total}/>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginTop:10}}>
                      {[{l:"A1 Normal",v:d.A1||0,c:P.a1},{l:"A2 Anomaly",v:d.A2||0,c:P.a2},{l:"A3 Incomplete",v:d.A3||0,c:P.a3}].map((s,j)=>(
                        <div key={j} style={{textAlign:"center",padding:"8px 4px",background:t.card,borderRadius:8}}>
                          <div style={{fontSize:15,fontWeight:800,color:s.c}}>{s.v.toLocaleString()}</div>
                          <div style={{fontSize:11,fontWeight:600,color:s.c}}>{pctS(s.v,d.total)}</div>
                          <div style={{fontSize:10,color:t.muted,marginTop:1}}>{s.l}</div>
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
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16}}>
            {[
              {title:"Duration Status",icon:"⏱",data:[{l:"Normal",c:P.normal,v:dc["NORMAL"]||0,dk:"DUR_NORMAL"},{l:"Short",c:P.short,v:dc["SHORT"]||0,dk:"DUR_SHORT"},{l:"Long",c:P.long,v:dc["LONG"]||0,dk:"DUR_LONG"}]},
              {title:"Distance Status",icon:"📍",data:[{l:"Near",c:P.near,v:di["NEAR"]||0,dk:"DIS_NEAR"},{l:"Mid",c:P.mid,v:di["MID"]||0,dk:"DIS_MID"},{l:"Far",c:P.far,v:di["FAR"]||0,dk:"DIS_FAR"},{l:"Incomplete",c:P.a3,v:di["INCOMPLETE"]||0,dk:"DIS_INC"}]},
              {title:"Location Status",icon:"📌",data:[{l:"Match",c:P.match,v:lc["MATCH"]||0,dk:"LOC_MATCH"},{l:"Not Match",c:P.notmatch,v:lc["NOT MATCH"]||0,dk:"LOC_NOTMATCH"},{l:"Incomplete",c:P.a3,v:lc["INCOMPLETE"]||0,dk:"LOC_INC"}]},
            ].map((sec,si)=>{
              const hasData = sec.data.some(d=>d.v>0);
              return(
              <div key={si} style={card()}>
                <div style={{fontWeight:700,marginBottom:12}}>{sec.icon} {sec.title}</div>
                {hasData?(
                  <>
                  <ResponsiveContainer width="100%" height={150}>
                    <PieChart>
                      <Pie data={sec.data.map(d=>({name:d.l,value:d.v,color:d.c,dk:d.dk}))} cx="50%" cy="50%" outerRadius={65} innerRadius={26} dataKey="value" strokeWidth={0}
                        onClick={d=>{if(d&&d.dk)openDrill(sec.title+" · "+d.name, d.color||d.fill||"#888", d.dk);}}>
                        {sec.data.map((d,i)=><Cell key={i} fill={d.c} style={{cursor:d.dk?"pointer":"default"}}/>)}
                      </Pie>
                      <Tooltip content={mkTip}/>
                    </PieChart>
                  </ResponsiveContainer>
                  {sec.data.map((d,i)=>(
                    <div key={i} onClick={()=>d.dk&&openDrill(sec.title+" · "+d.l,d.c,d.dk)}
                      style={{display:"flex",alignItems:"center",gap:8,padding:"6px 10px",background:t.cardAlt,borderRadius:8,marginBottom:5,cursor:d.dk?"pointer":"default",transition:"opacity 0.15s"}}
                      onMouseEnter={e=>{if(d.dk)e.currentTarget.style.opacity="0.75";}}
                      onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
                      <div style={{width:8,height:8,borderRadius:2,background:d.c,flexShrink:0}}/>
                      <span style={{fontSize:11,color:t.muted,flex:1}}>{d.l}</span>
                      <span style={{fontSize:12,fontWeight:700,color:d.c}}>{d.v.toLocaleString()}</span>
                      <span style={{fontSize:10,color:t.muted,minWidth:40,textAlign:"right"}}>{pctS(d.v,T)}</span>
                      {d.dk&&<span style={{fontSize:10,color:t.muted}}>›</span>}
                    </div>
                  ))}
                  </>
                ):(
                  <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:180,color:t.muted,fontSize:12,gap:8}}>
                    <span style={{fontSize:32,opacity:0.3}}>📭</span>
                    <span>Data tidak tersedia</span>
                    <span style={{fontSize:10,opacity:0.6}}>Kolom tidak ada di file ini</span>
                  </div>
                )}
              </div>
            );})}
            <div style={{...card(),gridColumn:"1/-1"}}>
              <div style={{fontWeight:700,marginBottom:14}}>📍 In Range Status</div>
              {(()=>{
                const irC=view.inRangeC||{};
                const irYes=irC["YES"]||0, irNo=irC["NO"]||0, irTotal=irYes+irNo;
                if(!irTotal) return <div style={{textAlign:"center",padding:20,color:t.muted,fontSize:12}}>Data tidak tersedia</div>;
                return(
                  <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:16}}>
                    <div>
                      <ResponsiveContainer width="100%" height={160}>
                        <PieChart>
                          <Pie data={[{name:"In Range",value:irYes,color:P.a1},{name:"Out of Range",value:irNo,color:P.investigate}]}
                            cx="50%" cy="50%" outerRadius={68} innerRadius={28} dataKey="value" strokeWidth={0}
                            onClick={d=>openDrill(d.name,d.color,d.name==="In Range"?"IR_YES":"IR_NO")}>
                            {[P.a1,P.investigate].map((col,i)=><Cell key={i} fill={col} style={{cursor:"pointer"}}/>)}
                          </Pie>
                          <Tooltip content={mkTip}/>
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                    <div style={{display:"flex",flexDirection:"column",gap:10,justifyContent:"center"}}>
                      {[{l:"✓ In Range",v:irYes,c:P.a1,dk:"IR_YES"},{l:"✗ Out of Range",v:irNo,c:P.investigate,dk:"IR_NO"}].map((d,i)=>(
                        <div key={i} onClick={()=>openDrill(d.l,d.c,d.dk)} style={{display:"flex",alignItems:"center",gap:10,padding:"12px 14px",background:t.cardAlt,borderRadius:10,cursor:"pointer",border:`1px solid ${t.border}`,transition:"opacity 0.15s"}}
                          onMouseEnter={e=>e.currentTarget.style.opacity="0.75"} onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
                          <div style={{width:12,height:12,borderRadius:3,background:d.c,flexShrink:0}}/>
                          <div style={{flex:1}}>
                            <div style={{fontWeight:700,fontSize:13,color:d.c}}>{d.l}</div>
                            <div style={{fontSize:11,color:t.muted,marginTop:1}}>{d.v.toLocaleString()} aktivitas</div>
                          </div>
                          <div style={{textAlign:"right"}}>
                            <div style={{fontSize:18,fontWeight:800,color:d.c}}>{pctS(d.v,irTotal)}</div>
                            <div style={{fontSize:10,color:t.muted}}>Klik untuk detail ›</div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })()}
            </div>

            <div style={{...card(),gridColumn:"1/-1"}}>
              <div style={{fontWeight:700,marginBottom:14}}>🔍 Visit Status Breakdown</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:10}}>
                {VIS.map((s,i)=>(
                  <div key={i} onClick={()=>openDrill(s.label,s.color,s.key)} style={{background:t.cardAlt,border:`1px solid ${t.border}`,borderRadius:10,padding:"14px 16px",borderTop:`3px solid ${s.color}`,cursor:"pointer",transition:"transform 0.15s,box-shadow 0.15s"}}
                    onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 6px 20px rgba(0,0,0,0.3)";}}
                    onMouseLeave={e=>{e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.boxShadow="none";}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                      <div style={{fontSize:22,fontWeight:800,color:s.color}}>{(vc[s.key]||0).toLocaleString()}</div>
                      <span style={{fontSize:10,color:t.muted,background:t.card,padding:"2px 6px",borderRadius:6}}>Klik ›</span>
                    </div>
                    <div style={{fontSize:11,color:t.muted,marginTop:2}}>{s.label}</div>
                    <div style={{fontSize:17,fontWeight:700,color:s.color,marginTop:4}}>{pctS(vc[s.key],T)}</div>
                    <div style={{height:4,borderRadius:3,background:t.border,marginTop:8}}>
                      <div style={{width:pct(vc[s.key],T)+"%",height:"100%",borderRadius:3,background:s.color}}/>
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
            <div style={{display:"flex",gap:8,marginBottom:8,flexWrap:"wrap",alignItems:"center"}}>
              <span style={{fontSize:11,color:t.muted,fontWeight:600}}>Sort:</span>
              {[["Jumlah A2","A2"],["% A2","a2p"],["Jumlah A3","A3"],["% A3","a3p"],["Jumlah Inv","INVESTIGATE"],["Total","total"]].map(([label,key])=>(
                <button key={key} onClick={()=>handleSort(key)}
                  style={{background:sk===key?P.accent:t.cardAlt,color:sk===key?"#fff":t.muted,border:"1px solid "+t.border,borderRadius:6,padding:"4px 10px",fontSize:10,fontWeight:700,cursor:"pointer"}}>
                  {label}{sk===key?(sd==="desc"?" ↓":" ↑"):""}
                </button>
              ))}
            </div>
            <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
              <input placeholder="🔍 Cari nama / cluster..." value={search} onChange={e=>setSearch(e.target.value)}
                style={{background:t.inputBg,border:`1px solid ${t.border}`,color:t.text,borderRadius:8,padding:"8px 14px",fontSize:12,outline:"none",width:210}}/>
              {[
                {id:"all",label:"Semua"},{id:"high_a2",label:"⚠️ A2 ≥30%"},
                {id:"high_a3",label:"🔵 A3 ≥20%"},{id:"top_a1",label:"✅ A1 ≥80%"},{id:"inv",label:"🔍 Inv ≥5%"},
              ].map(f=>(
                <button key={f.id} onClick={()=>setFq(f.id)} style={{padding:"7px 14px",borderRadius:8,border:`1px solid ${fq===f.id?"transparent":t.border}`,cursor:"pointer",fontSize:11,fontWeight:700,background:fq===f.id?"linear-gradient(135deg,#1d5fc0,#2d8ef5)":t.card,color:fq===f.id?"#fff":t.muted}}>{f.label}</button>
              ))}
              <span style={{marginLeft:"auto",fontSize:11,color:t.muted}}>{sorted.length} canvasser</span>
            </div>
            <div style={{...card({padding:0}),overflow:"hidden"}}>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Segoe UI',system-ui,-apple-system,sans-serif"}}>
                  <thead><tr style={{background:t.cardAlt}}>
                    {[["#",""],["Region","region"],["Cluster","cluster"],["Nama","name"],["Total","total"],
                      ["A1","A1"],["A2","A2"],["A3","A3"],["Inv","INVESTIGATE"],
                      ["A1%","a1p"],["A2%","a2p"],["A3%","a3p"],["Inv%","invP"],["Avg Dur","avgDur"],["Avg Dist","avgDis"]
                    ].map(([label,key])=>(
                      <th key={label} onClick={()=>key&&handleSort(key)} style={ths(key)}>
                        {label} {key&&sk===key?(sd==="desc"?"↓":"↑"):""}
                      </th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {sorted.slice(cPg*CPG,(cPg+1)*CPG).map((c,i)=>(
                      <tr key={c.name+c.cluster} style={{borderBottom:`1px solid ${t.border}`,background:i%2===0?"transparent":t.rowAlt,transition:"background 0.1s"}}
                        onMouseEnter={e=>e.currentTarget.style.background=t.rowHover}
                        onMouseLeave={e=>e.currentTarget.style.background=i%2===0?"transparent":t.rowAlt}>
                        <td style={{padding:"7px 10px",color:t.muted,fontSize:10}}>{i+1}</td>
                        <td style={{padding:"7px 10px"}}>
                          <span style={{background:P.accent+"20",color:P.accent,padding:"2px 8px",borderRadius:6,fontSize:10,fontWeight:700}}>{c.region||"–"}</span>
                        </td>
                        <td style={{padding:"7px 10px",color:t.muted,fontSize:11,whiteSpace:"nowrap",maxWidth:140,overflow:"hidden",textOverflow:"ellipsis"}}>{c.cluster||"–"}</td>
                        <td style={{padding:"7px 10px",fontWeight:600,whiteSpace:"nowrap"}}>{c.name}</td>
                        <td style={{padding:"7px 10px",fontWeight:700}}>{c.total.toLocaleString()}</td>
                        <td style={{padding:"7px 10px"}}>
                          {c.A1>0
                            ?<span onClick={()=>{const rows=getCanvasserRows(c.name,c.cluster,"A1");setCanvDetail({canvasser:c,drillLabel:"A1 - Normal",color:P.a1,rows,drillKey:"A1"});}} style={{color:P.a1,fontWeight:700,cursor:"pointer",borderBottom:"1px dotted "+P.a1}}>{c.A1.toLocaleString()}</span>
                            :<span style={{color:t.muted}}>0</span>}
                        </td>
                        <td style={{padding:"7px 10px"}}>
                          {c.A2>0
                            ?<span onClick={()=>{const rows=getCanvasserRows(c.name,c.cluster,"A2");setCanvDetail({canvasser:c,drillLabel:"A2 - Anomaly",color:P.a2,rows,drillKey:"A2"});}} style={{color:P.a2,fontWeight:700,cursor:"pointer",borderBottom:"1px dotted "+P.a2}}>{c.A2.toLocaleString()}</span>
                            :<span style={{color:t.muted}}>0</span>}
                        </td>
                        <td style={{padding:"7px 10px"}}>
                          {c.A3>0
                            ?<span onClick={()=>{const rows=getCanvasserRows(c.name,c.cluster,"A3");setCanvDetail({canvasser:c,drillLabel:"A3 - Incomplete",color:P.a3,rows,drillKey:"A3"});}} style={{color:P.a3,fontWeight:700,cursor:"pointer",borderBottom:"1px dotted "+P.a3}}>{c.A3.toLocaleString()}</span>
                            :<span style={{color:t.muted}}>0</span>}
                        </td>
                        <td style={{padding:"7px 10px",color:c.INVESTIGATE>0?P.investigate:t.muted,fontWeight:c.INVESTIGATE>0?700:400}}>{c.INVESTIGATE}</td>
                        <td style={{padding:"7px 10px"}}>
                          <div style={{display:"flex",alignItems:"center",gap:4}}>
                            <div style={{width:34,height:4,borderRadius:3,background:t.border,flexShrink:0}}>
                              <div style={{width:c.a1p+"%",height:"100%",borderRadius:3,background:c.a1p>=70?P.a1:P.a2}}/>
                            </div>
                            <span style={{color:c.a1p>=70?P.a1:P.a2,fontWeight:700,fontSize:11}}>{c.a1p.toFixed(1)}%</span>
                          </div>
                        </td>
                        <td style={{padding:"7px 10px",color:c.a2p>=40?P.investigate:P.a2}}>{c.a2p.toFixed(1)}%</td>
                        <td style={{padding:"7px 10px",color:P.a3}}>{c.a3p.toFixed(1)}%</td>
                        <td style={{padding:"7px 10px",color:c.invP>=5?P.investigate:t.muted,fontWeight:c.invP>=5?700:400}}>{c.invP.toFixed(1)}%</td>
                        <td style={{padding:"7px 10px",color:c.avgDur!=null&&c.avgDur<2?P.short:t.muted}}>{c.avgDur!=null?c.avgDur.toFixed(1)+" m":"—"}</td>
                        <td style={{padding:"7px 10px",color:(c.avgDis||0)>500?P.investigate:(c.avgDis||0)>100?P.a2:t.muted}}>{c.avgDis!=null?c.avgDis.toFixed(0)+" m":"—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <Pagination page={cPg} setPage={setCPg} total={sorted.length} pageSize={CPG} t={t}/>
            </div>
          </div>);})()}

      </div>
      <div style={{textAlign:"center",fontSize:10,color:t.muted,padding:"14px 22px 28px",opacity:0.4}}>XLSMART Analytics · Klik status di chart untuk lihat breakdown canvasser</div>
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
          <div style={{overflowY:"auto",flex:1}}>
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
          <div style={{fontSize:11,color:t.muted,marginBottom:20}}>Perubahan parameter memerlukan persetujuan management</div>
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
    <VisitTypeDrillPanel drill={vtDrill} onClose={()=>setVtDrill(null)} t={t}
      onCanvasserClick={(r,key)=>{
        const drillMap={"A1":"A1 - Normal","A2":"A2 - Anomaly","A3":"A3 - Incomplete"};
        const colorMap={"A1":P.a1,"A2":P.a2,"A3":P.a3};
        const rows=getCanvasserRows(r.name,r.cluster,key);
        setCanvDetail({canvasser:r,drillLabel:drillMap[key],color:colorMap[key],rows,drillKey:key});
      }}/>
    <OutletActivityPanel detail={outletActivity} onClose={()=>setOutletActivity(null)} t={t}/>
    <DrillDownPanel drill={drill} onClose={()=>setDrill(null)} t={t}
      onCanvasserClick={(r)=>{
        const rows=getCanvasserRows(r.name,r.cluster,drill.countKey);
        setCanvDetail({canvasser:r,drillLabel:drill.label,color:drill.color,rows,drillKey:drill.countKey});
      }}/>
    <CanvasserDetailPanel detail={canvDetail} onClose={()=>setCanvDetail(null)} t={t}/>
    </>
  );
}

// ── ROOT ──────────────────────────────────────────────────────────────────────
export default function App(){
  const [files,setFiles]=useState(null);
  const [roMap,setRoMap]=useState({});
  const [dark,setDark]=useState(true);
  const t=dark?DARK:LIGHT;
  return files
    ?<Dashboard files={files} onReset={()=>setFiles(null)} dark={dark} toggleDark={()=>setDark(d=>!d)} roMap={roMap}/>
    :<UploadScreen onLoad={setFiles} roMap={roMap} onRoLoad={setRoMap} t={t}/>;
}
