import { useState, useCallback, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell, AreaChart, Area, ReferenceLine
} from "recharts";

// ── PARSER ────────────────────────────────────────────────────────────────────
function parseWorkbook(wb) {
  const sessions = [];
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header:1, defval:"" });
  const headerIdxs = [];
  rows.forEach((row,i) => { if (String(row[0]).includes("FITFOCUS TRACKER")) headerIdxs.push(i); });

  headerIdxs.forEach(hIdx => {
    const dateRow = rows[hIdx+1] || [];
    const rawDate = String(dateRow[1]||dateRow[0]||"").replace(/^Date\s*:\s*/i,"").trim();
    const nextHeaderIdx = headerIdxs.find(i=>i>hIdx) || rows.length;
    const gymBlocks = [];
    for (let r=hIdx+1; r<nextHeaderIdx; r++) {
      if (rows[r].some(c=>String(c)==="Dititip")) gymBlocks.push({gymHeaderIdx:r,gymHeaderRow:rows[r]});
    }
    if (!gymBlocks.length) return;
    const allGyms = [];
    gymBlocks.forEach(({gymHeaderIdx:ghIdx,gymHeaderRow}) => {
      const gymNamesRow = rows[ghIdx-1]||[];
      const gymCols = [];
      gymHeaderRow.forEach((cell,ci) => { if (String(cell)==="Dititip") gymCols.push(ci); });
      const gymMap = {};
      gymCols.forEach(ci => {
        let name="";
        for (let back=ci; back>=Math.max(0,ci-5); back--) {
          const val=String(gymNamesRow[back]||"").trim();
          if (val&&val!=="VARIAN"&&val!=="Dititip"&&val!=="Sisa"&&val!=="Habis"&&val!=="Notes"&&val!=="PRODUCTION"&&val!=="SALES"&&val!=="SELL RATE") { name=val; break; }
        }
        gymMap[ci]=name||`Gym ${ci}`;
      });
      const flavorRows=[];
      for (let r=ghIdx+1; r<rows.length; r++) {
        const row=rows[r];
        const first=String(row[0]||"").trim();
        if (!first) continue;
        if (first==="TOTAL"||first==="SELL RATE"||first.includes("FITFOCUS")||first.includes("CARA PAKAI")) break;
        if (row.some(c=>String(c)==="Dititip")) break;
        flavorRows.push(row);
      }
      gymCols.forEach(ci => {
        const flavors={};
        flavorRows.forEach(row => {
          const fn=String(row[0]||"").trim();
          if (!fn) return;
          flavors[fn]=[parseFloat(row[ci])||0,parseFloat(row[ci+1])||0];
        });
        allGyms.push({gym:gymMap[ci],flavors});
      });
    });
    // Parse the DELIVERY block for this date (columns "Driver 1", "Driver 2", ... under a
    // "DELIVERY" section header). Sum whatever driver columns exist — new drivers added later
    // (Driver 2, Driver 3...) get picked up automatically, no code change needed.
    let deliveryTotal=0, deliveryFound=false;
    for (let r=hIdx+1; r<nextHeaderIdx; r++) {
      const row = rows[r] || [];
      if (row.some(c=>String(c).trim().toUpperCase()==="DELIVERY")) {
        const subHeaderRow = rows[r+1] || [];
        const driverCols = [];
        subHeaderRow.forEach((cell,ci) => { if (/^driver\s*\d*/i.test(String(cell).trim())) driverCols.push(ci); });
        const valueRow = rows[r+2] || [];
        driverCols.forEach(ci => { deliveryTotal += parseFloat(valueRow[ci]) || 0; });
        deliveryFound = driverCols.length>0;
        break;
      }
    }
    if (allGyms.length) sessions.push({date:rawDate,gyms:allGyms,delivery:deliveryTotal,deliveryFound});
  });
  return sessions;
}

function parseSessionDate(raw) {
  const months={jan:0,januari:0,feb:1,februari:1,mar:2,maret:2,apr:3,april:3,mei:4,may:4,jun:5,juni:5,jul:6,juli:6,aug:7,agustus:7,sep:8,september:8,okt:9,oktober:9,oct:9,nov:10,november:10,des:11,desember:11,dec:11};
  const clean=raw.toLowerCase().replace(/[^a-z0-9\s]/g,"");
  const parts=clean.split(/\s+/).filter(Boolean);
  let day=null,month=null,year=null;
  parts.forEach(p=>{
    if (/^\d{4}$/.test(p)) year=parseInt(p);
    else if (/^\d{1,2}$/.test(p)&&!day) day=parseInt(p);
    else if (months[p]!==undefined) month=months[p];
  });
  if (day&&month!==null) return new Date(year||2026,month,day);
  return null;
}

function formatDateShort(raw) {
  const d=parseSessionDate(raw);
  if (!d) return raw.replace(/^(senin|kamis|selasa|rabu|jumat|sabtu|minggu)\s+/i,"");
  return d.toLocaleDateString("en-US",{day:"numeric",month:"short"});
}

function getMonthKey(raw) {
  const d=parseSessionDate(raw);
  if (!d) return "Unknown";
  return d.toLocaleDateString("en-US",{month:"short",year:"numeric"});
}

function getYearKey(raw) {
  const d=parseSessionDate(raw);
  if (!d) return "Unknown";
  return String(d.getFullYear());
}

// ── FINANCIAL ─────────────────────────────────────────────────────────────────
const HPP_BY_FLAVOR={"Choco Forest":12420,"Mixed Berry":11500,"Pink Banana":11500,"Milky Dew":11500};
const HPP_DEFAULT=12000;
const SELL_PRICE=25000;
// Fallback only — used if a session's sheet is missing the DELIVERY block entirely.
// Real delivery cost is read per-session from the "Driver N" columns in the sheet (see parseWorkbook).
const DELIVERY_COST_FALLBACK=70000;
function getHPP(f){return HPP_BY_FLAVOR[f]??HPP_DEFAULT;}
function formatRp(n){if(Math.abs(n)>=1000000)return`Rp ${(n/1000000).toFixed(1)}M`;if(Math.abs(n)>=1000)return`Rp ${(n/1000).toFixed(0)}K`;return`Rp ${n.toLocaleString("en-US")}`;}
function formatRpFull(n){return"Rp "+Math.round(n).toLocaleString("en-US");}
// BUGFIX: only strip generic suffixes (GYM / FITNESS CENTER). The old regex also stripped
// "Solo Baru"/"Manahan"/"Gentan" which are the ONLY thing distinguishing e.g. "RPM Solo Baru"
// from "RPM Manahan" — both collapsed to "RPM" and became indistinguishable in charts/legends.
function gymShortName(g){const s=String(g).replace(/\s*(GYM|FITNESS CENTER)\b/gi,"").trim();return s||g;}
// Fixed display order for invoice cards (not sorted by quantity) — flavors not in this
// list fall to the end, in whatever order they first appear.
const FLAVOR_ORDER=["Choco Forest","Milky Dew","Pink Banana","Mixed Berry"];
function flavorSortIndex(f){const i=FLAVOR_ORDER.indexOf(f);return i===-1?FLAVOR_ORDER.length:i;}

// ── ANALYTICS ─────────────────────────────────────────────────────────────────
function buildMetrics(sessions, timeframe) {
  if (!sessions.length) return null;
  const sold=(d,s)=>Math.max(0,d-s);

  // SKIP sessions where total production = 0
  const validSessions = sessions.filter(s => {
    const totalProd = s.gyms.reduce((sum,g)=>sum+Object.values(g.flavors).reduce((a,[d])=>a+d,0),0);
    return totalProd > 0;
  });

  if (!validSessions.length) return null;

  // BUGFIX: previously d-s was silently clamped to 0 with Math.max(0,d-s) whenever
  // Remaining > Delivered (an impossible real-world value — a data-entry typo). That clamp
  // still happens below for the actual math, but now we also surface it as a warning
  // instead of hiding it, so bad input in the source Excel gets caught, not buried.
  const anomalies=[];
  sessions.forEach(s=>s.gyms.forEach(g=>Object.entries(g.flavors).forEach(([f,[d,si]])=>{
    if(si>d) anomalies.push({date:formatDateShort(s.date),gym:g.gym,flavor:f,delivered:d,remaining:si});
  })));

  // Sessions where no DELIVERY/Driver block was found in the sheet — these fall back to
  // DELIVERY_COST_FALLBACK, which can silently understate/overstate real cost. Surfaced as
  // a warning rather than hidden.
  const deliveryFallbackSessions = validSessions
    .filter(s=>!s.deliveryFound)
    .map(s=>formatDateShort(s.date));

  const allGyms=[...new Set(validSessions.flatMap(s=>s.gyms.map(g=>g.gym)))];
  const allFlavors=[...new Set(validSessions.flatMap(s=>s.gyms.flatMap(g=>Object.keys(g.flavors))))];

  // Group sessions by timeframe
  const groupKey = (s) => {
    if (timeframe==="weekly") return formatDateShort(s.date);
    if (timeframe==="monthly") return getMonthKey(s.date);
    return getYearKey(s.date);
  };

  // Build byDate (per session for charts)
  const byDate = validSessions.map(s => {
    let p=0,w=0,sl=0;
    s.gyms.forEach(g=>Object.values(g.flavors).forEach(([d,si])=>{p+=d;w+=si;sl+=sold(d,si);}));
    return {date:formatDateShort(s.date),produced:p,waste:w,sold:sl,sellRate:p>0?Math.round((sl/p)*100):0,rawDate:s.date};
  });

  // Build grouped data for timeframe view
  const grouped = {};
  validSessions.forEach(s => {
    const key = groupKey(s);
    if (!grouped[key]) grouped[key] = {produced:0,waste:0,sold:0};
    s.gyms.forEach(g=>Object.values(g.flavors).forEach(([d,si])=>{
      grouped[key].produced+=d; grouped[key].waste+=si; grouped[key].sold+=sold(d,si);
    }));
  });
  const byTimeframe = Object.entries(grouped).map(([date,v])=>({
    date, ...v, sellRate:v.produced>0?Math.round((v.sold/v.produced)*100):0
  }));

  // Per gym totals
  const gymMap={};
  allGyms.forEach(g=>{gymMap[g]={produced:0,waste:0,sold:0};});
  validSessions.forEach(s=>s.gyms.forEach(g=>{
    if(!gymMap[g.gym])gymMap[g.gym]={produced:0,waste:0,sold:0};
    Object.values(g.flavors).forEach(([d,si])=>{gymMap[g.gym].produced+=d;gymMap[g.gym].waste+=si;gymMap[g.gym].sold+=sold(d,si);});
  }));
  const gymData=allGyms.map(g=>({
    gym:g,gymShort:gymShortName(g),
    ...gymMap[g],sellRate:gymMap[g].produced>0?Math.round((gymMap[g].sold/gymMap[g].produced)*100):0
  })).sort((a,b)=>b.sellRate-a.sellRate);

  // Per gym per session (for separate chart)
  const gymPerSession = validSessions.map(s => {
    const row = {date:formatDateShort(s.date)};
    allGyms.forEach(gymName => {
      const gymEntry = s.gyms.find(g=>g.gym===gymName);
      if (gymEntry) {
        let p=0,sl=0;
        Object.values(gymEntry.flavors).forEach(([d,si])=>{p+=d;sl+=sold(d,si);});
        row[gymName] = p>0?Math.round((sl/p)*100):null;
      } else {
        row[gymName]=null;
      }
    });
    return row;
  });

  // Active vs cut gyms — a gym not seen in the last 6 sessions is treated as discontinued.
  // Its historical numbers still count everywhere (totals, trends), but it's excluded
  // from "current state" views like rankings, best/worst insights, and reorder recs.
  const RECENT_WINDOW=Math.min(6,validSessions.length);
  const recentSessions=validSessions.slice(-RECENT_WINDOW);
  const activeGymSet=new Set(recentSessions.flatMap(s=>s.gyms.map(g=>g.gym)));
  const newGymSet=new Set(validSessions.slice(-3).flatMap(s=>s.gyms.map(g=>g.gym)).filter(gym=>!validSessions.slice(0,-3).some(s=>s.gyms.some(g=>g.gym===gym))));
  gymData.forEach(g=>{g.active=activeGymSet.has(g.gym);g.isNew=newGymSet.has(g.gym);});
  const activeGymCount=gymData.filter(g=>g.active).length;
  const cutGymCount=gymData.length-activeGymCount;

  // Gym trend: compare first-half vs second-half average sell rate per gym
  const gymTrend={};
  allGyms.forEach(gym=>{
    const rates=gymPerSession.map(row=>row[gym]).filter(v=>v!==null&&v!==undefined);
    if(rates.length<2){gymTrend[gym]={direction:"flat",delta:0};return;}
    const mid=Math.ceil(rates.length/2);
    const firstAvg=rates.slice(0,mid).reduce((a,b)=>a+b,0)/mid;
    const secondRates=rates.slice(mid);
    const secondAvg=secondRates.length?secondRates.reduce((a,b)=>a+b,0)/secondRates.length:firstAvg;
    const delta=Math.round(secondAvg-firstAvg);
    gymTrend[gym]={direction:delta>3?"up":delta<-3?"down":"flat",delta};
  });
  gymData.forEach(g=>{g.trend=gymTrend[g.gym]||{direction:"flat",delta:0};});

  // Per flavor totals
  const flavorMap={};
  allFlavors.forEach(f=>{flavorMap[f]={produced:0,waste:0,sold:0};});
  validSessions.forEach(s=>s.gyms.forEach(g=>allFlavors.forEach(f=>{
    if(g.flavors[f]){const[d,si]=g.flavors[f];flavorMap[f].produced+=d;flavorMap[f].waste+=si;flavorMap[f].sold+=sold(d,si);}
  })));
  const flavorData=allFlavors.map(f=>({flavor:f,flavorShort:f.split(" ")[0],...flavorMap[f],sellRate:flavorMap[f].produced>0?Math.round((flavorMap[f].sold/flavorMap[f].produced)*100):0}));

  // Flavor per session breakdown
  const flavorPerSession = validSessions.map(s => {
    const row={date:formatDateShort(s.date)};
    allFlavors.forEach(f=>{
      let p=0,sl=0;
      s.gyms.forEach(g=>{if(g.flavors[f]){const[d,si]=g.flavors[f];p+=d;sl+=sold(d,si);}});
      row[f]=p>0?Math.round((sl/p)*100):null;
    });
    return row;
  });

  const totalProduced=byDate.reduce((a,b)=>a+b.produced,0);
  const totalSold=byDate.reduce((a,b)=>a+b.sold,0);
  const totalWaste=byDate.reduce((a,b)=>a+b.waste,0);
  const avgSellRate=totalProduced>0?Math.round((totalSold/totalProduced)*100):0;

  // Financial
  const byDateFinancial=validSessions.map((s,idx)=>{
    let revenue=0,hpp_total=0,soldCount=0,wasteCount=0,wasteCost=0;
    s.gyms.forEach(g=>Object.entries(g.flavors).forEach(([f,[d,si]])=>{
      const sl=Math.max(0,d-si);revenue+=sl*SELL_PRICE;hpp_total+=d*getHPP(f);soldCount+=sl;wasteCount+=si;wasteCost+=si*getHPP(f);
    }));
    const delivery = s.deliveryFound ? s.delivery : DELIVERY_COST_FALLBACK;
    return{date:byDate[idx].date,sold:soldCount,waste:wasteCount,revenue,hpp:hpp_total,delivery,wasteCost,profit:revenue-hpp_total-delivery};
  });
  const totalRevenue=byDateFinancial.reduce((a,b)=>a+b.revenue,0);
  const totalHPP=byDateFinancial.reduce((a,b)=>a+b.hpp,0);
  const totalDelivery=byDateFinancial.reduce((a,b)=>a+b.delivery,0);
  const totalProfit=byDateFinancial.reduce((a,b)=>a+b.profit,0);
  const totalWasteCost=byDateFinancial.reduce((a,b)=>a+b.wasteCost,0);

  // Cumulative profit — running total across sessions, so you can see growth not just per-session swings
  let runningProfit=0;
  const cumulativeProfit=byDateFinancial.map(d=>{runningProfit+=d.profit;return{date:d.date,profit:runningProfit};});

  // Gross profit per flavor (revenue − HPP). Excludes delivery since delivery isn't
  // naturally attributable to a single flavor — it's a per-trip, per-gym cost.
  const flavorFinancialMap={};
  allFlavors.forEach(f=>{flavorFinancialMap[f]={revenue:0,hpp:0};});
  validSessions.forEach(s=>s.gyms.forEach(g=>allFlavors.forEach(f=>{
    if(g.flavors[f]){const[d,si]=g.flavors[f];const sl=Math.max(0,d-si);flavorFinancialMap[f].revenue+=sl*SELL_PRICE;flavorFinancialMap[f].hpp+=d*getHPP(f);}
  })));
  flavorData.forEach(f=>{
    const fin=flavorFinancialMap[f.flavor]||{revenue:0,hpp:0};
    f.revenue=fin.revenue;f.hpp=fin.hpp;f.grossProfit=fin.revenue-fin.hpp;
    f.margin=fin.revenue>0?Math.round((f.grossProfit/fin.revenue)*100):0;
  });

  // Gym financial — delivery is one shared trip per session, so that session's actual
  // delivery cost (parsed from the sheet) is split only across the gyms visited that day.
  const gymFinancial=allGyms.map((gym)=>{
    let revenue=0,hpp_total=0,soldCount=0,wasteCount=0,gymDelivery=0;
    validSessions.forEach(s=>{
      const gymEntry=s.gyms.find(g=>g.gym===gym);
      if(!gymEntry)return;
      Object.entries(gymEntry.flavors).forEach(([f,[d,si]])=>{
        const sl=Math.max(0,d-si);revenue+=sl*SELL_PRICE;hpp_total+=d*getHPP(f);soldCount+=sl;wasteCount+=si;
      });
      gymDelivery+=(s.deliveryFound?s.delivery:DELIVERY_COST_FALLBACK)/s.gyms.length;
    });
    const gymShort=gymShortName(gym);
    return{gym,gymShort,revenue,hpp:hpp_total,delivery:gymDelivery,profit:revenue-hpp_total-gymDelivery,sold:soldCount,waste:wasteCount,active:activeGymSet.has(gym)};
  });

  // Qty recommendations — only for gyms still active. This is a multi-factor model, not a
  // single average, per gym PER FLAVOR (waste/trend can differ a lot between flavors in the
  // same gym):
  //   1. Recency-weighted average sold over the last 30 days (14-day half-life — a session
  //      from 2 weeks ago counts half as much as one from today).
  //   2. Linear regression of sold-over-time within that window, projected forward to the
  //      next likely session date, blended with #1 (more weight to regression the more data
  //      points we have).
  //   3. Decline-start detection — finds the peak in the recent trend, then the date the
  //      2-session rolling average first dropped to ≤70% of that peak and stayed there —
  //      i.e. the actual date a flavor started sliding, not just "trending down".
  //   4. Waste-rate-aware buffer: waste >20% of what's delivered → buffer trimmed below the
  //      raw estimate even if demand looks okay on paper (waste is prioritized over
  //      avoiding stockouts). Waste <5% AND sold out ≥30% of sessions → small buffer added
  //      (likely under-stocked, leaving money on the table).
  //   5. Confidence: high/medium/low based on how many data points and how volatile they are.
  const WINDOW_DAYS=30;
  const parsedDates=validSessions.map(s=>parseSessionDate(s.date)).filter(Boolean);
  const asOf=parsedDates.length?new Date(Math.max(...parsedDates.map(d=>d.getTime()))):null;
  function linreg(pts){
    const n=pts.length;
    if(n<2) return{slope:0,intercept:pts[0]?pts[0].y:0};
    const mx=pts.reduce((a,p)=>a+p.x,0)/n, my=pts.reduce((a,p)=>a+p.y,0)/n;
    let num=0,den=0;
    pts.forEach(p=>{num+=(p.x-mx)*(p.y-my);den+=(p.x-mx)*(p.x-mx);});
    const slope=den===0?0:num/den;
    return{slope,intercept:my-slope*mx};
  }
  const qtyRec=allGyms.filter(gym=>activeGymSet.has(gym)).map(gym=>{
    const trend=gymTrend[gym]||{direction:"flat",delta:0};
    const series={};
    validSessions.forEach(s=>{
      const d=parseSessionDate(s.date);
      if(!d||!asOf) return;
      const dayOffset=(d-asOf)/86400000;
      if(dayOffset<-WINDOW_DAYS) return;
      const g=s.gyms.find(x=>x.gym===gym);
      if(!g) return;
      Object.entries(g.flavors).forEach(([f,[dt,si]])=>{
        (series[f]=series[f]||[]).push({x:dayOffset,date:d,sold:Math.max(0,dt-si),dititip:dt,sisa:si});
      });
    });
    const flavors=allFlavors.map(f=>{
      const pts=(series[f]||[]).sort((a,b)=>a.x-b.x);
      if(!pts.length) return null;

      let wSum=0,wTot=0;
      pts.forEach(p=>{const w=Math.pow(0.5,Math.abs(p.x)/14);wSum+=p.sold*w;wTot+=w;});
      const weightedAvg=wTot>0?wSum/wTot:0;

      const gaps=[];for(let i=1;i<pts.length;i++)gaps.push(pts[i].x-pts[i-1].x);
      const medianGap=gaps.length?gaps.slice().sort((a,b)=>a-b)[Math.floor(gaps.length/2)]:4;
      const{slope,intercept}=linreg(pts.map(p=>({x:p.x,y:p.sold})));
      const nextX=pts[pts.length-1].x+(medianGap||4);
      const regProjection=Math.max(0,slope*nextX+intercept);
      const regWeight=pts.length>=4?0.45:pts.length>=3?0.3:0.1;
      const base=weightedAvg*(1-regWeight)+regProjection*regWeight;

      const totalDititip=pts.reduce((a,p)=>a+p.dititip,0);
      const totalSisa=pts.reduce((a,p)=>a+p.sisa,0);
      const wasteRate=totalDititip>0?totalSisa/totalDititip:0;
      const stockoutRate=pts.filter(p=>p.sisa===0&&p.dititip>0).length/pts.length;
      // Below 4 data points, waste%/stockout% swing too wildly on 1-2 sessions to mean
      // anything (e.g. one sold-out session out of 3 is already 33%) — treat as "not enough
      // history yet" instead of flagging a false signal. 4 matches the minimum the
      // decline-detector below also requires, so the two stay consistent.
      const enoughData=pts.length>=4;
      let buffer=0,flag=null;
      if(enoughData&&wasteRate>0.20){buffer=-1;flag="high waste — trimmed";}
      else if(enoughData&&wasteRate<0.05&&stockoutRate>0.30){buffer=1;flag="often sells out";}

      let declineSince=null;
      if(pts.length>=4){
        const roll=pts.map((p,i)=>{const win=pts.slice(Math.max(0,i-1),i+1);return{x:p.x,date:p.date,avg:win.reduce((a,q)=>a+q.sold,0)/win.length};});
        let peakIdx=0;
        roll.forEach((r,i)=>{if(r.avg>roll[peakIdx].avg)peakIdx=i;});
        for(let i=peakIdx+1;i<roll.length-1;i++){
          if(roll[i].avg<=roll[peakIdx].avg*0.7&&roll[i+1].avg<=roll[peakIdx].avg*0.7){declineSince=roll[i].date;break;}
        }
      }

      const mean=pts.reduce((a,p)=>a+p.sold,0)/pts.length;
      const variance=pts.reduce((a,p)=>a+Math.pow(p.sold-mean,2),0)/pts.length;
      const cv=mean>0?Math.sqrt(variance)/mean:0;
      const confidence=pts.length>=5&&cv<0.4?"high":pts.length>=3&&cv<0.7?"medium":"low";

      // Action-oriented signal instead of a flat severity level — each code says what to DO,
      // not just how alarming it is, and each gets its own color instead of red-means-anything:
      //   trim    = waste is high, cut the delivery back (red — cost is being lost)
      //   boost   = selling out with almost no waste, likely under-stocked (blue — opportunity,
      //             not a problem, so it should never read as an alarm)
      //   watch   = a real declining trend, worth keeping an eye on (amber)
      //   lowdata = under 4 sessions of history, recommendation is a rough estimate (gray)
      //   stable  = no signal — filtered out of the "needs attention" views entirely
      let signal;
      if(flag==="high waste — trimmed") signal={code:"trim",label:"TRIM",color:"#FF3B30"};
      else if(flag==="often sells out") signal={code:"boost",label:"BOOST",color:"#0A84FF"};
      else if(declineSince) signal={code:"watch",label:"WATCH",color:"#FF9500"};
      else if(!enoughData) signal={code:"lowdata",label:"LOW DATA",color:"#8E8E93"};
      else signal={code:"stable",label:null,color:null};

      return{
        flavor:f,flavorShort:f.split(" ")[0],
        avgSold:Math.round(weightedAvg*10)/10,
        recommended:weightedAvg>0?Math.max(1,Math.round(base)+buffer):Math.max(0,Math.round(base)+buffer),
        confidence,flag,signal,
        wasteRatePct:Math.round(wasteRate*100),
        stockoutRatePct:Math.round(stockoutRate*100),
        declineSince:declineSince?declineSince.toLocaleDateString("en-US",{day:"numeric",month:"short"}):null,
        wasteUnits:totalSisa,
        wasteCost:totalSisa*getHPP(f),
        dataPoints:pts.length
      };
    }).filter(f=>f&&(f.avgSold>0||f.wasteUnits>0));
    return{gym,gymShort:gymShortName(gym),trend,flavors};
  });

  // Top waste hotspots across active gyms (30-day window) — where over-delivery is costing
  // the most, sorted by cost so the highest-impact fixes surface first.
  const topWaste=qtyRec.flatMap(g=>g.flavors.map(f=>({gym:g.gym,gymShort:g.gymShort,flavor:f.flavor,flavorShort:f.flavorShort,wasteUnits:f.wasteUnits,wasteCost:f.wasteCost})))
    .filter(w=>w.wasteUnits>0).sort((a,b)=>b.wasteCost-a.wasteCost).slice(0,5);

  // Next-session projection — recency-weighted average of the last few sessions' actual
  // revenue/profit, compared against the segment before it to show trend direction. No
  // model call needed: this is just the same trend the sell-rate/profit charts already show,
  // surfaced as a forward-looking number.
  const projWindow=Math.min(3,byDateFinancial.length);
  const recentFin=byDateFinancial.slice(-projWindow);
  const priorFin=byDateFinancial.slice(-projWindow*2,-projWindow);
  const avg=(arr,key)=>arr.length?arr.reduce((a,b)=>a+b[key],0)/arr.length:0;
  const projRevenue=Math.round(avg(recentFin,"revenue"));
  const projProfit=Math.round(avg(recentFin,"profit"));
  const priorAvgProfit=avg(priorFin,"profit");
  const projDeltaPct=priorFin.length&&priorAvgProfit!==0?Math.round(((projProfit-priorAvgProfit)/Math.abs(priorAvgProfit))*100):null;
  const nextSessionProjection={
    revenue:projRevenue,profit:projProfit,
    direction:projDeltaPct===null?"flat":projDeltaPct>5?"up":projDeltaPct<-5?"down":"flat",
    deltaPct:projDeltaPct,basisSessions:projWindow
  };

  const n=byDate.length;

  return{byDate,byTimeframe,byDateFinancial,gymData,gymPerSession,gymFinancial,flavorData,flavorPerSession,cumulativeProfit,
    totalProduced,totalSold,totalWaste,avgSellRate,
    totalRevenue,totalProfit,totalWasteCost,totalHPP,totalDelivery,
    qtyRec,topWaste,nextSessionProjection,allGyms,allFlavors,activeGymCount,cutGymCount,recentWindow:RECENT_WINDOW,anomalies,deliveryFallbackSessions,
    sessionCount:n,gymCount:allGyms.length};
}

// ── PALETTE ──────────────────────────────────────────────────────────────────
// ── ICONS ────────────────────────────────────────────────────────────────────
// Minimal stroke-based line icons instead of OS emoji — emoji render as full-color cartoon
// glyphs that clash with the glass/neumorphic look no matter what's wrapped around them.
function Glyph({name,size=16,color="currentColor"}){
  const c={width:size,height:size,viewBox:"0 0 24 24",fill:"none",stroke:color,strokeWidth:1.8,strokeLinecap:"round",strokeLinejoin:"round"};
  switch(name){
    case"box":return(<svg {...c}><path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/></svg>);
    case"check":return(<svg {...c}><circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/></svg>);
    case"trash":return(<svg {...c}><path d="M4 7h16"/><path d="M6 7l1 12a2 2 0 002 2h6a2 2 0 002-2l1-12"/><path d="M9 7V4a1 1 0 011-1h4a1 1 0 011 1v3"/><path d="M10 11v6"/><path d="M14 11v6"/></svg>);
    case"trendUp":return(<svg {...c}><path d="M3 17l6-6 4 4 8-8"/><path d="M15 7h6v6"/></svg>);
    case"trendDown":return(<svg {...c}><path d="M3 7l6 6 4-4 8 8"/><path d="M15 17h6v-6"/></svg>);
    case"trophy":return(<svg {...c}><path d="M8 4h8v4a4 4 0 01-8 0V4z"/><path d="M8 5H5a3 3 0 003 3"/><path d="M16 5h3a3 3 0 01-3 3"/><path d="M12 12v5"/><path d="M9 20h6"/></svg>);
    case"alert":return(<svg {...c}><path d="M12 3l10 18H2L12 3z"/><path d="M12 10v4"/><path d="M12 17h.01"/></svg>);
    case"star":return(<svg {...c}><path d="M12 3l2.9 6 6.6.6-5 4.4 1.5 6.5L12 17l-5.9 3.5L7.6 14 2.6 9.6l6.6-.6L12 3z"/></svg>);
    case"cash":return(<svg {...c}><rect x="3" y="6" width="18" height="13" rx="2"/><path d="M3 10h18"/><circle cx="17" cy="14.5" r="1.3" fill={color} stroke="none"/></svg>);
    case"bars":return(<svg {...c}><path d="M4 20V10"/><path d="M10 20V4"/><path d="M16 20v-7"/><path d="M20 20v-3"/></svg>);
    case"calendar":return(<svg {...c}><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4"/><path d="M8 3v4"/><path d="M3 10h18"/></svg>);
    case"folder":return(<svg {...c}><path d="M3 7a2 2 0 012-2h4l2 2h8a2 2 0 012 2v9a2 2 0 01-2 2H5a2 2 0 01-2-2V7z"/></svg>);
    case"upload":return(<svg {...c}><path d="M12 16V4"/><path d="M7 9l5-5 5 5"/><path d="M4 16v3a2 2 0 002 2h12a2 2 0 002-2v-3"/></svg>);
    default:return null;
  }
}

// ── PALETTE ──────────────────────────────────────────────────────────────────
const P={sold:"#34C759",waste:"#FF9500",produced:"#5E5CE6",rate:"#0A84FF",forecast:"#AF52DE",
  gyms:["#5E5CE6","#34C759","#FF9500","#0A84FF","#FF2D55","#FFD60A","#30B0C7","#BF5AF2"],
  flavors:["#FF2D55","#5E5CE6","#FFD60A","#0A84FF","#34C759","#FF9500"]};

// ── UI PARTS ─────────────────────────────────────────────────────────────────
function StatCard({label,value,color,icon,span=1,hero=false}){
  return(
    <div style={{gridColumn:`span ${span}`,position:"relative",overflow:"hidden",background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",border:`1px solid ${color}33`,borderRadius:hero?26:20,padding:hero?"22px 24px":"18px 20px"}}>
      {hero&&<div style={{position:"absolute",top:-36,right:-36,width:130,height:130,borderRadius:"50%",background:color,opacity:0.14,filter:"blur(34px)"}}/>}
      <div style={{position:"relative",width:hero?42:32,height:hero?42:32,borderRadius:hero?13:10,background:`${color}1c`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:hero?20:16,marginBottom:hero?14:8}}>{icon}</div>
      <div style={{position:"relative",color,fontSize:hero?34:26,fontWeight:700,lineHeight:1}}>{value}</div>
      <div style={{position:"relative",color:"#8E8E93",fontSize:hero?13:12,marginTop:6,fontWeight:500}}>{label}</div>
    </div>
  );
}
function InsightCard({icon,label,title,sub,color,span=1}){
  return(
    <div style={{gridColumn:`span ${span}`,background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"14px 16px",border:`1px solid ${color}33`,display:"flex",gap:12,alignItems:"flex-start"}}>
      <div style={{width:32,height:32,minWidth:32,borderRadius:10,background:`${color}1c`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,flexShrink:0}}>{icon}</div>
      <div style={{minWidth:0}}>
        <div style={{fontSize:10,color,fontWeight:700,letterSpacing:1,marginBottom:4}}>{label}</div>
        <div style={{fontSize:15,fontWeight:800,color:"#1D1D1F",marginBottom:2}}>{title}</div>
        <div style={{fontSize:11,color:"#8E8E93"}}>{sub}</div>
      </div>
    </div>
  );
}
function TrendBadge({trend}){
  if(!trend||trend.direction==="flat") return <span style={{fontSize:10,color:"#AEAEB2",fontWeight:700}}>▬ steady</span>;
  const up=trend.direction==="up";
  return <span style={{fontSize:10,color:up?"#34C759":"#FF3B30",fontWeight:700}}>{up?"▲":"▼"} {Math.abs(trend.delta)}pt {up?"improving":"declining"}</span>;
}
const SIGNAL_RANK={trim:0,watch:1,boost:2,lowdata:3,stable:4};
function GymReorderCard({g,color}){
  const [showAll,setShowAll]=useState(false);
  const sorted=[...g.flavors].sort((a,b)=>SIGNAL_RANK[a.signal.code]-SIGNAL_RANK[b.signal.code]);
  const actionable=sorted.filter(f=>f.signal.code!=="stable");
  const shown=showAll?sorted:(actionable.length?actionable:sorted.slice(0,3));
  const hasMore=shown.length<sorted.length;
  return(
    <div style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:20,border:`1px solid ${color}33`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
        <div style={{fontSize:14,fontWeight:700,color}}>{g.gymShort}</div>
        <TrendBadge trend={g.trend}/>
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:13}}>
        {shown.map(f=>(
          <div key={f.flavor} style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontSize:13,color:"#1D1D1F"}}>{f.flavor}</span>
            <div style={{display:"flex",alignItems:"center",gap:9}}>
              {f.signal.label&&<span style={{fontSize:9,fontWeight:700,color:f.signal.color,background:f.signal.color+"16",padding:"3px 8px",borderRadius:20,letterSpacing:0.3}}>{f.signal.label}</span>}
              <span style={{fontSize:14,fontWeight:700,color:"#1D1D1F",minWidth:46,textAlign:"right"}}>{f.recommended} pcs</span>
            </div>
          </div>
        ))}
      </div>
      {hasMore&&<button onClick={()=>setShowAll(true)} style={{marginTop:16,background:"none",border:"none",color:"#5E5CE6",fontSize:12,fontWeight:600,cursor:"pointer",padding:0}}>View all {sorted.length} flavors</button>}
      {showAll&&actionable.length<sorted.length&&<button onClick={()=>setShowAll(false)} style={{marginTop:16,background:"none",border:"none",color:"#AEAEB2",fontSize:12,fontWeight:600,cursor:"pointer",padding:0}}>Show less</button>}
    </div>
  );
}
function SectionTitle({children}){
  return(
    <div style={{display:"flex",alignItems:"center",gap:10,margin:"26px 0 12px"}}>
      <div style={{width:4,height:18,background:"linear-gradient(180deg,#5E5CE6,#34C759)",borderRadius:2}}/>
      <h2 style={{margin:0,fontSize:15,fontWeight:700,color:"#1D1D1F"}}>{children}</h2>
    </div>
  );
}
const Tip=({active,payload,label})=>{
  if(!active||!payload?.length)return null;
  return(
    <div style={{background:"#E5E5EA",border:"1px solid #D1D1D6",borderRadius:10,padding:"9px 13px",fontSize:12}}>
      <p style={{margin:"0 0 5px",color:"#3A3A3C",fontWeight:600}}>{label}</p>
      {payload.map((p,i)=>(<p key={i} style={{margin:"2px 0",color:p.color}}>{p.name}: <strong>{p.value}{p.name?.toLowerCase().includes("rate")||p.name?.includes("%")?"%":""}</strong></p>))}
    </div>
  );
};

// ── TIMEFRAME SELECTOR ────────────────────────────────────────────────────────
function TimeframeSelector({value,onChange}){
  return(
    <div style={{display:"flex",gap:6,marginBottom:16}}>
      {[["weekly","Weekly"],["monthly","Monthly"],["yearly","Yearly"]].map(([v,l])=>(
        <button key={v} onClick={()=>onChange(v)}
          style={{padding:"6px 14px",borderRadius:980,border:"none",cursor:"pointer",fontSize:12,fontWeight:600,fontFamily:"inherit",transition:"background 0.15s,color 0.15s",
            background:value===v?"#5E5CE6":"#E9E9EB",color:value===v?"#fff":"#6E6E73"}}>
          {l}
        </button>
      ))}
    </div>
  );
}

// ── DATE FILTER ────────────────────────────────────────────────────────────────
function DateFilter({cutoff,onChange}){
  return(
    <div style={{display:"flex",alignItems:"center",gap:10,background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",border:"1px solid rgba(255,255,255,0.5)",borderRadius:12,padding:"10px 16px",marginBottom:16}}>
      <span style={{width:22,height:22,minWidth:22,borderRadius:7,background:"#5E5CE61c",display:"flex",alignItems:"center",justifyContent:"center"}}><Glyph name="calendar" size={12} color="#5E5CE6"/></span>
      <span style={{fontSize:12,color:"#8E8E93",fontWeight:600,whiteSpace:"nowrap"}}>Up to:</span>
      <input type="date" value={cutoff} onChange={e=>onChange(e.target.value)}
        style={{background:"#F5F5F7",border:"1px solid #E5E5EA",borderRadius:8,padding:"6px 10px",color:"#1D1D1F",fontSize:12,fontFamily:"inherit",outline:"none",flex:1}}/>
      {cutoff&&<button onClick={()=>onChange("")}
        style={{background:"#E9E9EB",border:"none",borderRadius:980,padding:"6px 12px",color:"#6E6E73",fontSize:11,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>
        Clear
      </button>}
    </div>
  );
}

// ── UPLOAD ────────────────────────────────────────────────────────────────────
function UploadScreen({onFile,dragging,onDragOver,onDragLeave,onDrop,error,syncLoading,syncError}){
  const bgStyle={minHeight:"100vh",position:"relative",background:"linear-gradient(160deg,#F1F0F8 0%,#F6F6FA 45%,#F3F8F4 100%)",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:32,fontFamily:"-apple-system,BlinkMacSystemFont,'SF Pro Display','SF Pro Text',system-ui,sans-serif",letterSpacing:"-0.01em",overflow:"hidden"};
  const blobs=(<>
    <div style={{position:"fixed",top:"-15%",left:"-12%",width:420,height:420,borderRadius:"50%",background:"#5E5CE6",opacity:0.18,filter:"blur(100px)",pointerEvents:"none"}}/>
    <div style={{position:"fixed",bottom:"-10%",right:"-12%",width:400,height:400,borderRadius:"50%",background:"#34C759",opacity:0.14,filter:"blur(110px)",pointerEvents:"none"}}/>
  </>);
  if(syncLoading){
    return(
      <div style={bgStyle}>
        {blobs}
        <style>{`
          @keyframes ffBreathe{0%,100%{transform:scale(1);opacity:0.8;}50%{transform:scale(1.1);opacity:1;}}
          @keyframes ffSpin{0%{transform:rotate(0deg);}100%{transform:rotate(360deg);}}
        `}</style>
        <div style={{position:"relative",width:84,height:84,marginBottom:20}}>
          <div style={{position:"absolute",inset:0,borderRadius:"50%",background:"linear-gradient(135deg,#5E5CE6,#34C759)",filter:"blur(20px)",animation:"ffBreathe 2.4s ease-in-out infinite"}}/>
          <div style={{position:"absolute",inset:0,borderRadius:"50%",border:"3px solid transparent",borderTopColor:"#5E5CE6",borderRightColor:"#5E5CE633",animation:"ffSpin 1s linear infinite"}}/>
          <div style={{position:"absolute",inset:13,borderRadius:"50%",background:"rgba(255,255,255,0.65)",backdropFilter:"blur(12px)",WebkitBackdropFilter:"blur(12px)",display:"flex",alignItems:"center",justifyContent:"center",boxShadow:"inset 0 1px 0 rgba(255,255,255,0.8)"}}><Glyph name="bars" size={22} color="#5E5CE6"/></div>
        </div>
        <div style={{position:"relative",color:"#1D1D1F",fontWeight:700,fontSize:15}}>Syncing your data…</div>
      </div>
    );
  }
  return(
    <div style={bgStyle}>
      {blobs}
      <h1 style={{position:"relative",color:"#1D1D1F",fontSize:28,fontWeight:700,letterSpacing:"-0.02em",margin:"0 0 6px"}}>FitFocus Analytics</h1>
      <p style={{position:"relative",color:"#8E8E93",fontSize:15,margin:"0 0 36px",textAlign:"center"}}>Upload your sales tracker to generate the full dashboard</p>
      {syncError&&<div style={{position:"relative",marginBottom:20,background:"#FFF6E5",color:"#B25E00",padding:"10px 20px",borderRadius:10,fontSize:12,maxWidth:380,textAlign:"center"}}>Couldn't auto-sync ({syncError}) — upload manually below, or check the Apps Script setup.</div>}
      <label onDragOver={onDragOver} onDragLeave={onDragLeave} onDrop={onDrop}
        style={{position:"relative",cursor:"pointer",display:"block",width:"100%",maxWidth:420,border:`2px dashed ${dragging?"#5E5CE6":"rgba(0,0,0,0.12)"}`,borderRadius:20,padding:"44px 28px",textAlign:"center",background:dragging?"rgba(94,92,230,0.08)":"rgba(255,255,255,0.55)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:dragging?"none":"0 10px 34px rgba(0,0,0,0.08),inset 0 1px 0 rgba(255,255,255,0.5)",transition:"all 0.2s"}}>
        <input type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>{if(e.target.files[0])onFile(e.target.files[0]);}}/>
        <div style={{marginBottom:12,color:"#5E5CE6"}}><Glyph name={dragging?"folder":"upload"} size={34}/></div>
        <div style={{color:"#1D1D1F",fontWeight:700,fontSize:15,marginBottom:6}}>{dragging?"Drop it!":"Drag & drop your Excel file"}</div>
        <div style={{color:"#AEAEB2",fontSize:13}}>or click to browse · .xlsx / .xls</div>
      </label>
      {error&&<div style={{position:"relative",marginTop:20,background:"#FFF1F0",color:"#D70015",padding:"10px 20px",borderRadius:10,fontSize:13,maxWidth:380,textAlign:"center"}}>{error}</div>}
    </div>
  );
}

// ── APPS SCRIPT CONFIG ───────────────────────────────────────────────────────
// Paste your deployed Apps Script Web App URL here (see api/AppsScript.gs setup instructions).
// It looks like: https://script.google.com/macros/s/AKfycb.../exec
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyjae8ljZwXp3pdcxqV5B-MhiQc3PCwEAvf2MMYV29E0qMprWulUZwa4dlCpMZJ9tkc/exec";

// Apps Script Web Apps proxy every request through a second googleusercontent.com redirect
// before your code runs. Browsers automatically convert a POST into a GET on that redirect
// (per the fetch spec) — which silently drops the request body. Query-string params survive
// the redirect fine, so this uses GET with everything in the URL instead of a POST body.
async function callAppsScript(payload){
  const params = new URLSearchParams();
  Object.entries(payload).forEach(([key,value])=>{
    if(value!==undefined && value!==null){
      params.append(key,value);
    }
  });
  const res = await fetch(`${APPS_SCRIPT_URL}?${params.toString()}`, {method:"GET"});
  if(!res.ok){
    throw new Error(`HTTP ${res.status}`);
  }
  return await res.json();
}

// ── LOGO ──────────────────────────────────────────────────────────────────────
// Put your logo file at public/logo.png in the repo — Vercel serves anything in public/
// straight from the root URL, so /logo.png just works once deployed. No logo uploaded yet?
// This quietly falls back to the "FF" text mark instead of showing a broken image icon.
function LogoMark({size=52}){
  const [imgError,setImgError]=useState(false);
  return(
    <div style={{width:size,height:size,borderRadius:size*0.3,background:"rgba(94,92,230,0.12)",border:"1px solid rgba(94,92,230,0.25)",display:"flex",alignItems:"center",justifyContent:"center",overflow:"hidden",flexShrink:0}}>
      {!imgError
        ? <img src="/logo.png" alt="FitFocus" onError={()=>setImgError(true)} style={{width:"100%",height:"100%",objectFit:"cover"}}/>
        : <span style={{color:"#5E5CE6",fontSize:size*0.38,fontWeight:700,letterSpacing:"-0.03em"}}>FF</span>}
    </div>
  );
}

function GlassIconButton({onClick,disabled,icon,accent="#5E5CE6",children}){
  const [pressed,setPressed]=useState(false);
  return(
    <button onClick={onClick} disabled={disabled}
      onMouseDown={()=>setPressed(true)} onMouseUp={()=>setPressed(false)} onMouseLeave={()=>setPressed(false)}
      onTouchStart={()=>setPressed(true)} onTouchEnd={()=>setPressed(false)}
      style={{
        cursor:disabled?"default":"pointer",display:"flex",alignItems:"center",gap:8,
        background:"rgba(255,255,255,0.55)",backdropFilter:"blur(16px) saturate(180%)",WebkitBackdropFilter:"blur(16px) saturate(180%)",
        border:"1px solid rgba(255,255,255,0.6)",borderRadius:980,padding:"6px 16px 6px 6px",
        fontSize:13,fontWeight:600,color:"#1D1D1F",fontFamily:"inherit",
        boxShadow:pressed?"inset 3px 3px 8px rgba(148,148,180,0.25),inset -3px -3px 8px rgba(255,255,255,0.6)":"4px 4px 10px rgba(148,148,180,0.18),-4px -4px 10px rgba(255,255,255,0.75)",
        transform:pressed?"scale(0.97)":"scale(1)",
        transition:"transform 0.12s,box-shadow 0.12s"}}>
      <span style={{width:24,height:24,borderRadius:"50%",background:`linear-gradient(135deg,${accent},${accent}aa)`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,color:"#fff",flexShrink:0}}>{icon}</span>
      {children}
    </button>
  );
}
// ── LOGIN ─────────────────────────────────────────────────────────────────────
function LoginScreen({onSubmit,loading,error}){
  const [email,setEmail]=useState("");
  const [password,setPassword]=useState("");
  return(
    <div style={{minHeight:"100vh",position:"relative",background:"linear-gradient(160deg,#F1F0F8 0%,#F6F6FA 45%,#F3F8F4 100%)",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:32,fontFamily:"-apple-system,BlinkMacSystemFont,'SF Pro Display','SF Pro Text',system-ui,sans-serif",letterSpacing:"-0.01em",overflow:"hidden"}}>
      <div style={{position:"fixed",top:"-15%",left:"-12%",width:420,height:420,borderRadius:"50%",background:"#5E5CE6",opacity:0.2,filter:"blur(100px)",pointerEvents:"none"}}/>
      <div style={{position:"fixed",bottom:"-10%",right:"-12%",width:400,height:400,borderRadius:"50%",background:"#34C759",opacity:0.16,filter:"blur(110px)",pointerEvents:"none"}}/>
      <div style={{position:"relative",background:"rgba(255,255,255,0.55)",backdropFilter:"blur(30px) saturate(180%)",WebkitBackdropFilter:"blur(30px) saturate(180%)",border:"1px solid rgba(255,255,255,0.6)",borderRadius:28,padding:"40px 32px",boxShadow:"0 20px 60px rgba(0,0,0,0.1),inset 0 1px 0 rgba(255,255,255,0.6)",width:"100%",maxWidth:340,display:"flex",flexDirection:"column",alignItems:"center"}}>
        <div style={{marginBottom:22}}><LogoMark size={52}/></div>
        <h1 style={{color:"#1D1D1F",fontSize:19,fontWeight:600,letterSpacing:"-0.02em",margin:"0 0 28px"}}>Sign in to FitFocus</h1>
        <form onSubmit={e=>{e.preventDefault();if(email&&password&&!loading)onSubmit(email,password);}} style={{width:"100%"}}>
          <input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="Email" autoFocus autoCapitalize="none" autoCorrect="off"
            style={{width:"100%",boxSizing:"border-box",padding:"13px 16px",borderRadius:12,border:"1px solid rgba(0,0,0,0.08)",background:"rgba(255,255,255,0.7)",fontSize:15,fontFamily:"inherit",marginBottom:10,outline:"none"}}/>
          <input type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="Password"
            style={{width:"100%",boxSizing:"border-box",padding:"13px 16px",borderRadius:12,border:"1px solid rgba(0,0,0,0.08)",background:"rgba(255,255,255,0.7)",fontSize:15,fontFamily:"inherit",marginBottom:18,outline:"none"}}/>
          <button type="submit" disabled={loading||!email||!password}
            style={{width:"100%",padding:"13px 16px",borderRadius:12,border:"none",cursor:loading?"default":"pointer",fontSize:15,fontWeight:600,fontFamily:"inherit",
              background:loading?"#C7C7E6":"#5E5CE6",color:"#fff",transition:"background 0.15s",boxShadow:"0 4px 16px rgba(94,92,230,0.35)"}}>
            {loading?"Signing in…":"Sign in"}
          </button>
        </form>
        {error&&<div style={{marginTop:16,background:"#FFF1F0",color:"#D70015",padding:"10px 20px",borderRadius:10,fontSize:13,textAlign:"center"}}>{error}</div>}
      </div>
    </div>
  );
}

// ── MAIN ──────────────────────────────────────────────────────────────────────
export default function FitFocusDashboard(){
  const [authToken,setAuthToken]=useState(null);
  const [loginLoading,setLoginLoading]=useState(false);
  const [loginError,setLoginError]=useState("");
  const [rawSessions,setRawSessions]=useState(null);
  const [fileName,setFileName]=useState("");
  const [dragging,setDragging]=useState(false);
  const [activeTab,setActiveTab]=useState("overview");
  const [error,setError]=useState("");
  const [cutoff,setCutoff]=useState("");
  const [timeframe,setTimeframe]=useState("weekly");
  const [invoiceDates,setInvoiceDates]=useState([]);
  const [expandedMonths,setExpandedMonths]=useState(null);
  const [syncLoading,setSyncLoading]=useState(false);
  const [syncError,setSyncError]=useState("");
  const [lastSynced,setLastSynced]=useState(null);
  const [uploadPressed,setUploadPressed]=useState(false);

  const handleLogin=useCallback(async(email,password)=>{
    setLoginLoading(true);setLoginError("");
    try{
      const data=await callAppsScript({action:"login",email,password});
      if(data.ok) setAuthToken(data.token);
      else setLoginError(data.error||"wrong email or password");
    }catch(err){
      setLoginError(String((err&&err.message)||err));
    }finally{
      setLoginLoading(false);
    }
  },[]);

  // Shared parse path for both the Apps-Script-synced file and a manually uploaded one, so
  // there's exactly one place that turns raw bytes into session data.
  const parseArrayBuffer=useCallback((arrayBuffer,label)=>{
    try{
      const wb=XLSX.read(arrayBuffer,{type:"array"});
      const data=parseWorkbook(wb);
      if(!data.length){setError("No session data found.");return false;}
      setRawSessions(data);setFileName(label);setActiveTab("overview");setCutoff("");setError("");
      return true;
    }catch(err){setError("Could not read the file.");return false;}
  },[]);

  const processFile=useCallback((file)=>{
    setError("");
    const reader=new FileReader();
    reader.onload=(e)=>{parseArrayBuffer(e.target.result,file.name);};
    reader.readAsArrayBuffer(file);
  },[parseArrayBuffer]);

  // Auto-sync via Apps Script — the .xlsx never leaves Google's own servers except as bytes
  // sent to this authenticated session; the Drive file itself stays completely private
  // (Apps Script reads it server-side under your own account, not via a public link).
  const syncFromScript=useCallback(async()=>{
    if(!authToken) return;
    setSyncLoading(true);setSyncError("");
    try{
      const data=await callAppsScript({action:"getData",token:authToken});
      if(!data.ok){
        if(/session|token/i.test(data.error||"")) setAuthToken(null); // expired — back to login
        setSyncError(data.error||"sync failed");
        return;
      }
      const binary=atob(data.fileBase64);
      const bytes=new Uint8Array(binary.length);
      for(let i=0;i<binary.length;i++) bytes[i]=binary.charCodeAt(i);
      const ok=parseArrayBuffer(bytes.buffer,data.fileName||"Google Sheets (via Apps Script)");
      if(ok) setLastSynced(new Date());
    }catch(err){
      setSyncError(String((err&&err.message)||err));
    }finally{
      setSyncLoading(false);
    }
  },[authToken,parseArrayBuffer]);

  useEffect(()=>{if(authToken)syncFromScript();},[authToken]); // eslint-disable-line react-hooks/exhaustive-deps

  const onDragOver=(e)=>{e.preventDefault();setDragging(true);};
  const onDragLeave=()=>setDragging(false);
  const onDrop=(e)=>{e.preventDefault();setDragging(false);if(e.dataTransfer.files[0])processFile(e.dataTransfer.files[0]);};

  const sessions=useMemo(()=>{
    if(!rawSessions) return [];
    return cutoff
      ?rawSessions.filter(s=>{const d=parseSessionDate(s.date);return!d||d<=new Date(cutoff+"T23:59:59");})
      :rawSessions;
  },[rawSessions,cutoff]);

  const M=useMemo(()=>rawSessions?buildMetrics(sessions,timeframe):null,[rawSessions,sessions,timeframe]);

  // Invoicing — sum sold (Delivered-Remaining) per gym per flavor across whichever sessions
  // are checked. Independent of buildMetrics' "valid session" filter on purpose: even a
  // zero-total session should still be selectable/countable here for an invoice.
  const invoiceData=useMemo(()=>{
    const chosen=sessions.filter(s=>invoiceDates.includes(s.date));
    const gymMap={};
    chosen.forEach(s=>s.gyms.forEach(g=>{
      if(!gymMap[g.gym])gymMap[g.gym]={};
      Object.entries(g.flavors).forEach(([f,[d,si]])=>{
        const sl=Math.max(0,d-si);
        gymMap[g.gym][f]=(gymMap[g.gym][f]||0)+sl;
      });
    }));
    return Object.entries(gymMap).map(([gym,flavorMap])=>{
      const flavors=Object.entries(flavorMap).map(([flavor,qty])=>({flavor,qty})).sort((a,b)=>flavorSortIndex(a.flavor)-flavorSortIndex(b.flavor));
      return{gym,gymShort:gymShortName(gym),flavors,total:flavors.reduce((a,b)=>a+b.qty,0)};
    }).sort((a,b)=>b.total-a.total);
  },[sessions,invoiceDates]);

  if(!authToken){ return <LoginScreen onSubmit={handleLogin} loading={loginLoading} error={loginError}/>; }
  if(!rawSessions){ return <UploadScreen onFile={processFile} dragging={dragging} onDragOver={onDragOver} onDragLeave={onDragLeave} onDrop={onDrop} error={error} syncLoading={syncLoading} syncError={syncError}/>; }

  if(!M)return(
    <div style={{minHeight:"100vh",background:"#F5F5F7",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{color:"#8E8E93",fontSize:14}}>No valid sessions. All sessions have zero production.</div>
    </div>
  );

  const activeGymData=M.gymData.filter(g=>g.active);
  const cutGymData=M.gymData.filter(g=>!g.active);
  const worstGym=[...activeGymData].sort((a,b)=>b.waste-a.waste)[0]||M.gymData[0];
  const bestGym=[...activeGymData].sort((a,b)=>b.sellRate-a.sellRate)[0]||M.gymData[0];
  const worstFlavor=[...M.flavorData].sort((a,b)=>b.waste-a.waste)[0];
  const bestFlavor=[...M.flavorData].sort((a,b)=>b.sellRate-a.sellRate)[0];

  const tabs=[
    {id:"overview",label:"Overview"},
    {id:"gyms",label:"By Gym"},
    {id:"flavors",label:"By Flavor"},
    {id:"trends",label:"Trends"},
    {id:"financial",label:"Financial"},
    {id:"profit",label:"Profit"},
    {id:"invoice",label:"Invoice"},
    {id:"forecast",label:"🔮"},
    {id:"review",label:"Gym Review"},
  ];

  const GYM_COLORS=P.gyms;
  const FLAVOR_COLORS=P.flavors;

  return(
    <div style={{minHeight:"100vh",position:"relative",background:"linear-gradient(160deg,#F1F0F8 0%,#F6F6FA 45%,#F3F8F4 100%)",color:"#1D1D1F",fontFamily:"-apple-system,BlinkMacSystemFont,'SF Pro Display','SF Pro Text',system-ui,sans-serif",letterSpacing:"-0.01em",paddingBottom:60}}>
      {/* Ambient blurred color — glass panels need something colorful behind them to actually read as glass, not just gray */}
      <div style={{position:"fixed",top:"-12%",left:"-10%",width:440,height:440,borderRadius:"50%",background:"#5E5CE6",opacity:0.16,filter:"blur(100px)",pointerEvents:"none",zIndex:0}}/>
      <div style={{position:"fixed",bottom:"0%",right:"-10%",width:400,height:400,borderRadius:"50%",background:"#34C759",opacity:0.13,filter:"blur(110px)",pointerEvents:"none",zIndex:0}}/>
      <div style={{position:"fixed",top:"38%",right:"18%",width:280,height:280,borderRadius:"50%",background:"#FF9500",opacity:0.11,filter:"blur(90px)",pointerEvents:"none",zIndex:0}}/>
      <div style={{position:"relative",zIndex:1}}>
      {/* Header */}
      <div style={{background:"rgba(255,255,255,0.55)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",borderBottom:"1px solid rgba(255,255,255,0.5)",padding:"20px 26px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <LogoMark size={34}/>
          <div>
          <div style={{display:"flex",alignItems:"center",gap:9,marginBottom:4}}>
            <span style={{fontSize:19,fontWeight:700,letterSpacing:"-0.02em"}}>FitFocus</span>
            <span style={{fontSize:10,color:"#5E5CE6",fontWeight:700,letterSpacing:1.2,background:"#5E5CE614",padding:"3px 7px",borderRadius:6}}>ANALYTICS</span>
          </div>
          <div style={{color:"#8E8E93",fontSize:12,fontWeight:500}}>{fileName} · {M.sessionCount} sessions · {M.activeGymCount} active gyms{M.cutGymCount>0?` (+${M.cutGymCount} discontinued)`:""}{cutoff?` · up to ${new Date(cutoff).toLocaleDateString("en-US",{day:"numeric",month:"short",year:"numeric"})}`:""}
          </div>
          {lastSynced&&<div style={{color:"#AEAEB2",fontSize:11,marginTop:2}}>Synced {lastSynced.toLocaleTimeString([],{hour:"2-digit",minute:"2-digit"})}{syncError?` · last refresh failed (${syncError})`:""}</div>}
        </div>
        </div>
        <div style={{display:"flex",gap:8}}>
          <GlassIconButton onClick={syncFromScript} disabled={syncLoading} icon={syncLoading?"⏳":"↻"} accent="#5E5CE6">
            {syncLoading?"Syncing…":"Refresh"}
          </GlassIconButton>
          <label onDragOver={onDragOver} onDragLeave={onDragLeave} onDrop={onDrop}
            onMouseDown={()=>setUploadPressed(true)} onMouseUp={()=>setUploadPressed(false)} onMouseLeave={()=>setUploadPressed(false)}
            onTouchStart={()=>setUploadPressed(true)} onTouchEnd={()=>setUploadPressed(false)}
            style={{cursor:"pointer",display:"flex",alignItems:"center",gap:8,background:"rgba(255,255,255,0.55)",backdropFilter:"blur(16px) saturate(180%)",WebkitBackdropFilter:"blur(16px) saturate(180%)",border:"1px solid rgba(255,255,255,0.6)",borderRadius:980,padding:"6px 16px 6px 6px",fontSize:13,fontWeight:600,color:"#1D1D1F",
              boxShadow:uploadPressed?"inset 3px 3px 8px rgba(148,148,180,0.25),inset -3px -3px 8px rgba(255,255,255,0.6)":"4px 4px 10px rgba(148,148,180,0.18),-4px -4px 10px rgba(255,255,255,0.75)",
              transform:uploadPressed?"scale(0.97)":"scale(1)",transition:"transform 0.12s,box-shadow 0.12s"}}>
            <input type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>{if(e.target.files[0])processFile(e.target.files[0]);}}/>
            <span style={{width:24,height:24,borderRadius:"50%",background:"linear-gradient(135deg,#34C759,#34C759aa)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,color:"#fff",flexShrink:0}}>⇧</span>
            Upload new file
          </label>
        </div>
      </div>

      {/* Tabs — segmented control */}
      <div style={{padding:"16px 26px 0"}}>
        <div style={{display:"inline-flex",gap:2,background:"rgba(255,255,255,0.4)",backdropFilter:"blur(20px) saturate(180%)",WebkitBackdropFilter:"blur(20px) saturate(180%)",border:"1px solid rgba(255,255,255,0.5)",borderRadius:13,padding:3,overflowX:"auto",maxWidth:"100%"}}>
          {tabs.map(t=>(
            <button key={t.id} onClick={()=>setActiveTab(t.id)}
              style={{padding:"7px 16px",borderRadius:10,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit",whiteSpace:"nowrap",
                background:activeTab===t.id?"rgba(255,255,255,0.9)":"transparent",
                color:activeTab===t.id?"#1D1D1F":"#6E6E73",
                boxShadow:activeTab===t.id?"0 2px 8px rgba(0,0,0,0.08),inset 0 1px 0 rgba(255,255,255,0.8)":"none",
                transition:"background 0.2s,color 0.2s,box-shadow 0.2s"}}>
              {t.label}
            </button>
          ))}
        </div>
      </div>

      <div style={{padding:"20px 26px"}}>
        <DateFilter cutoff={cutoff} onChange={setCutoff}/>

        {/* OVERVIEW */}
        {activeTab==="overview"&&(<>
          {M.anomalies.length>0&&(
            <div style={{background:"#FF3B3011",border:"1px solid #FF3B3033",borderRadius:18,padding:"14px 18px",marginBottom:12}}>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                <span style={{width:18,height:18,minWidth:18,borderRadius:6,background:"#FF3B3022",color:"#FF3B30",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800}}>!</span>
                <span style={{fontSize:11,color:"#FF3B30",fontWeight:700}}>{M.anomalies.length} DATA ANEH DI EXCEL</span>
              </div>
              <div style={{fontSize:12,color:"#6E6E73",marginBottom:8}}>Remaining exceeds Delivered (data entry error) — counted as 0 sold for now, but should be verified in the source file:</div>
              {M.anomalies.slice(0,6).map((a,i)=>(
                <div key={i} style={{fontSize:11,color:"#8E8E93",padding:"2px 0"}}>{a.date} · {gymShortName(a.gym)} · {a.flavor} — Delivered {a.delivered}, Remaining {a.remaining}</div>
              ))}
              {M.anomalies.length>6&&<div style={{fontSize:11,color:"#AEAEB2",marginTop:2}}>+{M.anomalies.length-6} more</div>}
            </div>
          )}
          {M.deliveryFallbackSessions.length>0&&(
            <div style={{background:"#FF950011",border:"1px solid #FF950033",borderRadius:18,padding:"14px 18px",marginBottom:12}}>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                <span style={{width:18,height:18,minWidth:18,borderRadius:6,background:"#FF950022",color:"#FF9500",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800}}>!</span>
                <span style={{fontSize:11,color:"#FF9500",fontWeight:700}}>{M.deliveryFallbackSessions.length} SESI PAKAI DELIVERY COST DEFAULT</span>
              </div>
              <div style={{fontSize:12,color:"#6E6E73"}}>No "Driver" column found in the sheet for: {M.deliveryFallbackSessions.join(", ")}. Defaulted to Rp {DELIVERY_COST_FALLBACK.toLocaleString("en-US")} — check the sheet if the real cost differed.</div>
            </div>
          )}
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gridAutoRows:"minmax(0,auto)",gap:12,marginBottom:12}}>
            <StatCard icon={<Glyph name="box" color="#5E5CE6" size={20}/>} label="Total Produced" value={M.totalProduced} color="#5E5CE6" span={2} hero/>
            <StatCard icon={<Glyph name="check" color="#34C759"/>} label="Total Sold" value={M.totalSold} color="#34C759"/>
            <StatCard icon={<Glyph name="trash" color="#FF9500"/>} label="Total Waste" value={M.totalWaste} color="#FF9500"/>
            <StatCard icon={<Glyph name="trendUp" color="#0A84FF"/>} label="Avg Sell Rate" value={`${M.avgSellRate}%`} color="#0A84FF" span={2}/>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:12,marginBottom:4}}>
            <InsightCard icon={<Glyph name="trophy" color="#34C759"/>} label="BEST GYM" color="#34C759" title={bestGym.gym} sub={`${bestGym.sellRate}% sell rate · ${bestGym.sold} sold`} span={2}/>
            <InsightCard icon={<Glyph name="alert" color="#FF9500"/>} label="MOST WASTE GYM" color="#FF9500" title={worstGym.gym} sub={`${worstGym.waste} units wasted`}/>
            <InsightCard icon={<Glyph name="star" color="#34C759"/>} label="TOP FLAVOR" color="#34C759" title={bestFlavor.flavor} sub={`${bestFlavor.sellRate}% sell rate`}/>
            {(()=>{
              const improving=[...activeGymData].sort((a,b)=>(b.trend?.delta||0)-(a.trend?.delta||0))[0];
              const declining=[...activeGymData].sort((a,b)=>(a.trend?.delta||0)-(b.trend?.delta||0))[0];
              const show=improving&&improving.trend?.direction==="up"?improving:declining;
              const isUp=show===improving&&improving?.trend?.direction==="up";
              return show&&show.trend&&show.trend.direction!=="flat"
                ?<InsightCard icon={<Glyph name={isUp?"trendUp":"trendDown"} color={isUp?"#34C759":"#FF3B30"}/>} label={isUp?"MOST IMPROVED":"MOST DECLINING"} color={isUp?"#34C759":"#FF3B30"} title={show.gym} sub={`${isUp?"+":""}${show.trend.delta}pt sell rate vs earlier sessions`}/>
                :<InsightCard icon={<Glyph name="alert" color="#FF9500"/>} label="MOST WASTED FLAVOR" color="#FF9500" title={worstFlavor.flavor} sub={`${worstFlavor.waste} units wasted`}/>;
            })()}
          </div>
          {cutGymData.length>0&&(
            <div style={{fontSize:11,color:"#AEAEB2",marginBottom:8,padding:"0 2px"}}>
              ℹ️ {cutGymData.map(g=>g.gymShort).join(", ")} are no longer supplied (not present in the last {M.recentWindow} sessions) — excluded from rankings & recommendations, though their data remains in history & trends.
            </div>
          )}
          <SectionTitle>Sold vs Waste</SectionTitle>
          <TimeframeSelector value={timeframe} onChange={setTimeframe}/>
          <div style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"16px 4px 8px"}}>
            <ResponsiveContainer width="100%" height={230}>
              <BarChart data={M.byTimeframe}>
                <CartesianGrid strokeDasharray="3 3" stroke="#E5E5EA"/>
                <XAxis dataKey="date" tick={{fill:"#8E8E93",fontSize:10}}/>
                <YAxis tick={{fill:"#8E8E93",fontSize:10}}/>
                <Tooltip content={<Tip/>}/>
                <Legend wrapperStyle={{color:"#6E6E73",fontSize:11}}/>
                <Bar dataKey="sold" name="Sold" fill={P.sold} radius={[3,3,0,0]}/>
                <Bar dataKey="waste" name="Waste" fill={P.waste} radius={[3,3,0,0]}/>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <SectionTitle>Overall Sell Rate</SectionTitle>
          <TimeframeSelector value={timeframe} onChange={setTimeframe}/>
          <div style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"16px 4px 8px"}}>
            <ResponsiveContainer width="100%" height={190}>
              <AreaChart data={M.byTimeframe}>
                <defs><linearGradient id="rg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#0A84FF" stopOpacity={0.3}/><stop offset="95%" stopColor="#0A84FF" stopOpacity={0}/></linearGradient></defs>
                <CartesianGrid strokeDasharray="3 3" stroke="#E5E5EA"/>
                <XAxis dataKey="date" tick={{fill:"#8E8E93",fontSize:10}}/>
                <YAxis domain={[0,100]} tick={{fill:"#8E8E93",fontSize:10}} tickFormatter={v=>`${v}%`}/>
                <Tooltip content={<Tip/>}/>
                <Area type="monotone" dataKey="sellRate" name="Sell Rate" stroke="#0A84FF" fill="url(#rg)" strokeWidth={2} dot={{r:3,fill:"#0A84FF"}}/>
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </>)}

        {/* BY GYM */}
        {activeTab==="gyms"&&(<>
          <SectionTitle>Gym Ranking — Best to Worst</SectionTitle>
          {cutGymData.length>0&&(
            <div style={{fontSize:11,color:"#AEAEB2",marginBottom:10,padding:"0 2px"}}>
              ℹ️ {cutGymData.map(g=>g.gymShort).join(", ")} have been discontinued (inactive for the last {M.recentWindow} sessions) — hidden from this ranking, but still shown in the trend chart below.
            </div>
          )}
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(170px,1fr))",gap:12,marginBottom:8}}>
            {activeGymData.map((g,i)=>(
              <div key={g.gym} style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"15px 16px",border:`1px solid ${GYM_COLORS[i%GYM_COLORS.length]}44`}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                  <div style={{fontSize:12,fontWeight:700,color:GYM_COLORS[i%GYM_COLORS.length]}}>{g.gym}</div>
                  <div style={{display:"flex",alignItems:"center",gap:6}}>
                    {g.isNew&&<span style={{fontSize:9,color:"#0A84FF",fontWeight:700,background:"#0A84FF1a",padding:"2px 6px",borderRadius:20}}>NEW</span>}
                    <span style={{fontSize:10,color:"#AEAEB2",fontWeight:700}}>#{i+1}</span>
                  </div>
                </div>
                {[["Produced",g.produced,"#6E6E73"],["Sold",g.sold,"#34C759"],["Waste",g.waste,"#FF9500"]].map(([l,v,c])=>(
                  <div key={l} style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                    <span style={{color:"#8E8E93",fontSize:12}}>{l}</span>
                    <span style={{fontWeight:700,fontSize:12,color:c}}>{v}</span>
                  </div>
                ))}
                <div style={{background:"#E5E5EA",borderRadius:20,height:5,overflow:"hidden",marginTop:8}}>
                  <div style={{width:`${g.sellRate}%`,height:"100%",background:GYM_COLORS[i%GYM_COLORS.length],borderRadius:20}}/>
                </div>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:3}}>
                  <TrendBadge trend={g.trend}/>
                  <span style={{fontSize:11,color:GYM_COLORS[i%GYM_COLORS.length],fontWeight:700}}>{g.sellRate}%</span>
                </div>
              </div>
            ))}
          </div>

          <SectionTitle>Sell Rate per Gym per Session</SectionTitle>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:10,marginBottom:8}}>
            {M.allGyms.map((gym,i)=>(
              <div key={gym} style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"10px 6px 4px",border:`1px solid ${GYM_COLORS[i%GYM_COLORS.length]}33`}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"0 10px 2px"}}>
                  <span style={{fontSize:11,fontWeight:700,color:GYM_COLORS[i%GYM_COLORS.length]}}>{gymShortName(gym)}</span>
                  <TrendBadge trend={M.gymData.find(g=>g.gym===gym)?.trend}/>
                </div>
                <ResponsiveContainer width="100%" height={90}>
                  <LineChart data={M.gymPerSession} margin={{top:4,right:8,left:8,bottom:0}}>
                    <YAxis domain={[0,100]} hide/>
                    <XAxis dataKey="date" hide/>
                    <Tooltip content={<Tip/>}/>
                    <Line type="monotone" dataKey={gym} name={gymShortName(gym)}
                      stroke={GYM_COLORS[i%GYM_COLORS.length]} strokeWidth={2} dot={{r:2}}
                      connectNulls={false}/>
                  </LineChart>
                </ResponsiveContainer>
              </div>
            ))}
          </div>

        </>)}

        {/* BY FLAVOR */}
        {activeTab==="flavors"&&(<>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12}}>
            {M.flavorData.map((f,i)=>(
              <div key={f.flavor} style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"15px 16px",border:`1px solid ${FLAVOR_COLORS[i%FLAVOR_COLORS.length]}44`}}>
                <div style={{fontSize:12,fontWeight:700,color:FLAVOR_COLORS[i%FLAVOR_COLORS.length],marginBottom:7}}>{f.flavor}</div>
                <div style={{fontSize:24,fontWeight:700,color:"#1D1D1F",lineHeight:1}}>{f.sellRate}%</div>
                <div style={{fontSize:11,color:"#8E8E93",marginBottom:7}}>sell rate</div>
                <div style={{fontSize:12,color:"#6E6E73"}}>Sold: <strong style={{color:"#34C759"}}>{f.sold}</strong></div>
                <div style={{fontSize:12,color:"#6E6E73"}}>Waste: <strong style={{color:"#FF9500"}}>{f.waste}</strong></div>
                <div style={{fontSize:12,color:"#6E6E73"}}>Produced: <strong>{f.produced}</strong></div>
                <div style={{fontSize:12,color:"#6E6E73",marginTop:4,paddingTop:4,borderTop:"1px solid #F0F0F2"}}>Gross profit: <strong style={{color:f.grossProfit>=0?"#AF52DE":"#FF3B30"}}>{formatRp(f.grossProfit)}</strong> <span style={{color:"#AEAEB2"}}>({f.margin}% margin)</span></div>
              </div>
            ))}
          </div>

          <SectionTitle>Flavor Sell Rate per Session</SectionTitle>
          <div style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"16px 4px 8px",marginBottom:8}}>
            <ResponsiveContainer width="100%" height={230}>
              <LineChart data={M.flavorPerSession}>
                <CartesianGrid strokeDasharray="3 3" stroke="#E5E5EA"/>
                <XAxis dataKey="date" tick={{fill:"#8E8E93",fontSize:9}}/>
                <YAxis domain={[0,100]} tick={{fill:"#8E8E93",fontSize:10}} tickFormatter={v=>`${v}%`}/>
                <Tooltip content={<Tip/>}/>
                <Legend wrapperStyle={{color:"#6E6E73",fontSize:11}}/>
                {M.allFlavors.map((f,i)=>(
                  <Line key={f} type="monotone" dataKey={f} name={f.split(" ")[0]}
                    stroke={FLAVOR_COLORS[i%FLAVOR_COLORS.length]} strokeWidth={2} dot={{r:2}} connectNulls={false}/>
                ))}
              </LineChart>
            </ResponsiveContainer>
          </div>

          <SectionTitle>Waste Share by Flavor</SectionTitle>
          <div style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"16px 4px 8px"}}>
            <ResponsiveContainer width="100%" height={220}>
              <PieChart>
                <Pie data={M.flavorData} dataKey="waste" nameKey="flavor" cx="50%" cy="50%" outerRadius={88}
                  label={({flavor,percent})=>`${flavor.split(" ")[0]} ${(percent*100).toFixed(0)}%`} labelLine={false}>
                  {M.flavorData.map((_,i)=><Cell key={i} fill={FLAVOR_COLORS[i%FLAVOR_COLORS.length]}/>)}
                </Pie>
                <Tooltip content={<Tip/>}/>
              </PieChart>
            </ResponsiveContainer>
          </div>
        </>)}

        {/* TRENDS */}
        {activeTab==="trends"&&(<>
          <SectionTitle>Cumulative Profit</SectionTitle>
          <div style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"16px 4px 8px",marginBottom:8}}>
            <ResponsiveContainer width="100%" height={220}>
              <AreaChart data={M.cumulativeProfit}>
                <defs><linearGradient id="cp" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#AF52DE" stopOpacity={0.35}/><stop offset="95%" stopColor="#AF52DE" stopOpacity={0}/></linearGradient></defs>
                <CartesianGrid strokeDasharray="3 3" stroke="#E5E5EA"/>
                <XAxis dataKey="date" tick={{fill:"#8E8E93",fontSize:9}}/>
                <YAxis tick={{fill:"#8E8E93",fontSize:10}} tickFormatter={v=>formatRp(v)}/>
                <Tooltip content={<Tip/>} formatter={v=>formatRpFull(v)}/>
                <Area type="monotone" dataKey="profit" name="Cumulative Profit" stroke="#AF52DE" fill="url(#cp)" strokeWidth={2} dot={{r:3,fill:"#AF52DE"}}/>
              </AreaChart>
            </ResponsiveContainer>
            <div style={{fontSize:11,color:"#8E8E93",padding:"0 10px 4px"}}>Running total of profit session by session — this is the number that tells you if the business is actually growing.</div>
          </div>

          <SectionTitle>Session History</SectionTitle>
          <div style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,overflow:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:560}}>
              <thead>
                <tr style={{background:"#E5E5EA"}}>
                  {["Date","Sold","Waste","Sell Rate","Revenue","Profit"].map(h=>(
                    <th key={h} style={{padding:"10px 12px",textAlign:"left",color:"#8E8E93",fontWeight:600,fontSize:11,whiteSpace:"nowrap"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {M.byDate.map((d,i)=>{
                  const fin=M.byDateFinancial[i];
                  return(
                  <tr key={i} style={{borderTop:"1px solid #E5E5EA"}}>
                    <td style={{padding:"8px 12px",color:"#3A3A3C",fontWeight:600,whiteSpace:"nowrap"}}>{d.date}</td>
                    <td style={{padding:"8px 12px",color:"#34C759",fontWeight:600}}>{d.sold}</td>
                    <td style={{padding:"8px 12px",color:"#FF9500",fontWeight:600}}>{d.waste}</td>
                    <td style={{padding:"8px 12px"}}>
                      <span style={{background:d.sellRate>=80?"#34C75922":d.sellRate>=50?"#0A84FF22":"#FF950022",color:d.sellRate>=80?"#34C759":d.sellRate>=50?"#0A84FF":"#FF9500",padding:"2px 9px",borderRadius:20,fontWeight:700,fontSize:11}}>{d.sellRate}%</span>
                    </td>
                    <td style={{padding:"8px 12px",color:"#34C759"}}>{formatRpFull(fin.revenue)}</td>
                    <td style={{padding:"8px 12px",fontWeight:700,color:fin.profit>=0?"#AF52DE":"#FF3B30"}}>{formatRpFull(fin.profit)}</td>
                  </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr style={{borderTop:"2px solid #D1D1D6",background:"#E5E5EA"}}>
                  <td style={{padding:"10px 12px",color:"#1D1D1F",fontWeight:700}}>TOTAL</td>
                  <td style={{padding:"10px 12px",color:"#1D1D1F",fontWeight:700}}>{M.totalSold}</td>
                  <td style={{padding:"10px 12px",color:"#1D1D1F",fontWeight:700}}>{M.totalWaste}</td>
                  <td style={{padding:"10px 12px",color:"#1D1D1F",fontWeight:700}}>{M.avgSellRate}%</td>
                  <td style={{padding:"10px 12px",color:"#34C759",fontWeight:700}}>{formatRpFull(M.totalRevenue)}</td>
                  <td style={{padding:"10px 12px",fontWeight:700,color:M.totalProfit>=0?"#AF52DE":"#FF3B30"}}>{formatRpFull(M.totalProfit)}</td>
                </tr>
              </tfoot>
            </table>
          </div>
        </>)}

        {/* FINANCIAL */}
        {activeTab==="financial"&&(<>
          <div style={{background:"#5E5CE611",border:"1px solid #5E5CE633",borderRadius:18,padding:"14px 18px",marginBottom:12}}>
            <div style={{fontSize:11,color:"#5E5CE6",fontWeight:700,marginBottom:4}}>FORMULA</div>
            <div style={{fontSize:12,color:"#6E6E73"}}>
              <span style={{color:"#34C759"}}>Revenue</span> − <span style={{color:"#FF9500"}}>HPP (all produced)</span> − <span style={{color:"#0A84FF"}}>Delivery (actual per session)</span> = <span style={{color:"#AF52DE",fontWeight:700}}>Real Profit</span>
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:12,marginBottom:12}}>
            {[
              {label:"Total Revenue",value:M.totalRevenue,color:"#34C759",sub:`${M.totalSold} bottles × Rp 25K`},
              {label:"Total HPP",value:M.totalHPP,color:"#FF9500",sub:"All units delivered"},
              {label:"Total Delivery",value:M.totalDelivery,color:"#0A84FF",sub:`${M.sessionCount} sessions · avg Rp ${M.sessionCount>0?Math.round(M.totalDelivery/M.sessionCount).toLocaleString("en-US"):0}/session`},
              {label:"Real Profit",value:M.totalProfit,color:M.totalProfit>=0?"#AF52DE":"#FF3B30",sub:"Revenue − HPP − Delivery"},
            ].map(({label,value,color,sub})=>(
              <div key={label} style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"14px 16px",border:`1px solid ${color}33`}}>
                <div style={{fontSize:11,color:"#8E8E93",marginBottom:4}}>{label}</div>
                <div style={{fontSize:18,fontWeight:700,color,lineHeight:1}}>{formatRpFull(value)}</div>
                <div style={{fontSize:10,color:"#AEAEB2",marginTop:5}}>{sub}</div>
              </div>
            ))}
          </div>

          <SectionTitle>Waste Cost in Rp</SectionTitle>
          <div style={{background:"#FF3B3011",border:"1px solid #FF3B3033",borderRadius:18,padding:"14px 18px",marginBottom:12}}>
            <div style={{fontSize:11,color:"#FF3B30",fontWeight:700,marginBottom:4}}>TOTAL CAPITAL LOST TO WASTE</div>
            <div style={{fontSize:24,fontWeight:700,color:"#FF3B30"}}>{formatRpFull(M.totalWasteCost)}</div>
            <div style={{fontSize:12,color:"#8E8E93",marginTop:4}}>{M.totalWaste} bottles · calculated using each flavor's cost of goods (HPP)</div>
          </div>

          <SectionTitle>Profit by Flavor</SectionTitle>
          <div style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"16px 4px 8px",marginBottom:8}}>
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={M.flavorData.map(f=>({...f,flavor:f.flavorShort}))}>
                <CartesianGrid strokeDasharray="3 3" stroke="#E5E5EA"/>
                <XAxis dataKey="flavor" tick={{fill:"#8E8E93",fontSize:11}}/>
                <YAxis tick={{fill:"#8E8E93",fontSize:10}} tickFormatter={v=>formatRp(v)}/>
                <Tooltip content={<Tip/>} formatter={v=>formatRpFull(v)}/>
                <Bar dataKey="grossProfit" name="Gross Profit" radius={[5,5,0,0]}>
                  {M.flavorData.map((f,i)=><Cell key={i} fill={f.grossProfit>=0?FLAVOR_COLORS[i%FLAVOR_COLORS.length]:"#FF3B30"}/>)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
            <div style={{fontSize:11,color:"#8E8E93",padding:"0 10px 4px"}}>Revenue minus HPP for all units produced of that flavor. Excludes delivery, since delivery is per-trip, not per-flavor.</div>
          </div>

          <SectionTitle>Gym Profitability</SectionTitle>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12,marginBottom:8}}>
            {M.gymFinancial.map((g,i)=>(
              <div key={g.gym} style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"14px 16px",border:`1px solid ${g.profit>=0?"#34C75933":"#FF3B3033"}`}}>
                <div style={{fontSize:12,fontWeight:700,color:GYM_COLORS[i%GYM_COLORS.length],marginBottom:8}}>{g.gym}</div>
                {[["Revenue",g.revenue,"#34C759"],["HPP",g.hpp,"#FF9500"],["Delivery",g.delivery,"#0A84FF"],["Profit",g.profit,g.profit>=0?"#AF52DE":"#FF3B30"]].map(([l,v,c])=>(
                  <div key={l} style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                    <span style={{fontSize:11,color:"#8E8E93"}}>{l}</span>
                    <span style={{fontSize:11,fontWeight:700,color:c}}>{formatRpFull(v)}</span>
                  </div>
                ))}
                <div style={{marginTop:8,padding:"4px 8px",borderRadius:6,textAlign:"center",background:g.profit>=0?"#34C75922":"#FF3B3022",color:g.profit>=0?"#34C759":"#FF3B30",fontSize:10,fontWeight:700}}>
                  {g.profit>=0?"✓ Profitable":"⚠ Bleeding"}
                </div>
              </div>
            ))}
          </div>

        </>)}

        {/* PROFIT PER SESSION */}
        {activeTab==="profit"&&(()=>{
          const pf=M.byDateFinancial;
          const losingSessions=pf.filter(d=>d.profit<0);
          const bestSession=[...pf].sort((a,b)=>b.profit-a.profit)[0];
          const worstSession=[...pf].sort((a,b)=>a.profit-b.profit)[0];
          const avgProfit=Math.round(pf.reduce((a,b)=>a+b.profit,0)/pf.length);
          return(<>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:12,marginBottom:12}}>
              {[
                {label:"Avg Profit/Session",value:avgProfit,color:avgProfit>=0?"#AF52DE":"#FF3B30"},
                {label:"Best Session",value:bestSession.profit,color:"#34C759",sub:bestSession.date},
                {label:"Worst Session",value:worstSession.profit,color:"#FF3B30",sub:worstSession.date},
                {label:"Losing Sessions",value:losingSessions.length,color:"#FF9500",isCount:true,sub:`of ${pf.length} sessions`},
              ].map(({label,value,color,sub,isCount})=>(
                <div key={label} style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"14px 16px",border:`1px solid ${color}33`}}>
                  <div style={{fontSize:11,color:"#8E8E93",marginBottom:4}}>{label}</div>
                  <div style={{fontSize:18,fontWeight:700,color,lineHeight:1}}>{isCount?value:formatRpFull(value)}</div>
                  {sub&&<div style={{fontSize:10,color:"#AEAEB2",marginTop:5}}>{sub}</div>}
                </div>
              ))}
            </div>

            <SectionTitle>Profit per Session</SectionTitle>
            <div style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"16px 4px 8px",marginBottom:12}}>
              <ResponsiveContainer width="100%" height={260}>
                <BarChart data={pf}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E5E5EA"/>
                  <XAxis dataKey="date" tick={{fill:"#8E8E93",fontSize:9}}/>
                  <YAxis tick={{fill:"#8E8E93",fontSize:10}} tickFormatter={v=>formatRp(v)}/>
                  <Tooltip content={<Tip/>} formatter={v=>formatRpFull(v)}/>
                  <ReferenceLine y={0} stroke="#8E8E93" strokeWidth={1.5}/>
                  <Bar dataKey="profit" name="Profit" radius={[4,4,4,4]}>
                    {pf.map((d,i)=><Cell key={i} fill={d.profit>=0?"#34C759":"#FF3B30"}/>)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
              <div style={{fontSize:11,color:"#8E8E93",padding:"0 10px 4px"}}>Red indicates a session that was genuinely unprofitable (not just a low sell rate). Unlike Cumulative Profit in the Trends tab, which shows the running total, this view highlights exactly which sessions are dragging performance down.</div>
            </div>

            {losingSessions.length>0&&(
              <div style={{background:"#FF3B3011",border:"1px solid #FF3B3033",borderRadius:18,padding:"14px 18px"}}>
                <div style={{fontSize:11,color:"#FF3B30",fontWeight:700,marginBottom:6}}>LOSING SESSIONS</div>
                {losingSessions.map(d=>(
                  <div key={d.date} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"4px 0",color:"#6E6E73"}}>
                    <span>{d.date}</span>
                    <span style={{fontWeight:700,color:"#FF3B30"}}>{formatRpFull(d.profit)}</span>
                  </div>
                ))}
              </div>
            )}
          </>);
        })()}

        {/* INVOICING */}
        {activeTab==="invoice"&&(()=>{
          const totalBottles=invoiceData.reduce((a,g)=>a+g.total,0);
          const toggle=(d)=>setInvoiceDates(prev=>prev.includes(d)?prev.filter(x=>x!==d):[...prev,d]);
          const selectLastN=(n)=>setInvoiceDates(sessions.slice(-n).map(s=>s.date));
          return(<>
            <div style={{background:"#5E5CE611",border:"1px solid #5E5CE633",borderRadius:18,padding:"14px 18px",marginBottom:12}}>
              <div style={{fontSize:11,color:"#5E5CE6",fontWeight:700,marginBottom:4}}>SELECT SESSIONS</div>
              <p style={{margin:0,color:"#6E6E73",fontSize:13,lineHeight:1.6}}>Select the sessions to combine — the result shows total bottles sold per gym per flavor across those sessions only.</p>
            </div>

            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:12}}>
              {[1,3,5,10].map(n=>(
                <button key={n} onClick={()=>selectLastN(n)} style={{padding:"7px 14px",borderRadius:20,border:"1px solid #5E5CE633",background:"#5E5CE611",color:"#5E5CE6",fontSize:12,fontWeight:600,cursor:"pointer"}}>Last {n} sessions</button>
              ))}
              <button onClick={()=>setInvoiceDates(sessions.map(s=>s.date))} style={{padding:"7px 14px",borderRadius:20,border:"1px solid #34C75933",background:"#34C75911",color:"#34C759",fontSize:12,fontWeight:600,cursor:"pointer"}}>All</button>
              <button onClick={()=>setInvoiceDates([])} style={{padding:"7px 14px",borderRadius:20,border:"1px solid #FF3B3033",background:"#FF3B3011",color:"#FF3B30",fontSize:12,fontWeight:600,cursor:"pointer"}}>Clear</button>
            </div>

            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16}}>
              {(()=>{
                const groups={};
                sessions.forEach(s=>{
                  const key=getMonthKey(s.date);
                  if(!groups[key])groups[key]=[];
                  groups[key].push(s);
                });
                const monthKeys=Object.keys(groups); // sessions are already in sheet order (oldest→newest)
                const latestKey=monthKeys[monthKeys.length-1];
                const isExpanded=(key)=>expandedMonths===null?key===latestKey:expandedMonths.includes(key);
                const toggleMonth=(key)=>setExpandedMonths(prev=>{
                  const base=prev===null?[latestKey]:prev;
                  return base.includes(key)?base.filter(k=>k!==key):[...base,key];
                });
                const selectMonth=(key)=>setInvoiceDates(prev=>[...new Set([...prev,...groups[key].map(s=>s.date)])]);
                return monthKeys.map(key=>{
                  const monthSessions=groups[key];
                  const monthSelectedCount=monthSessions.filter(s=>invoiceDates.includes(s.date)).length;
                  const open=isExpanded(key);
                  return(
                    <div key={key} style={{width:"100%",background:"#FFFFFF",borderRadius:14,border:"1px solid #E5E5EA",overflow:"hidden"}}>
                      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 14px",cursor:"pointer"}} onClick={()=>toggleMonth(key)}>
                        <div style={{display:"flex",alignItems:"center",gap:8}}>
                          <span style={{fontSize:12,color:"#8E8E93"}}>{open?"▾":"▸"}</span>
                          <span style={{fontSize:13,fontWeight:700,color:"#1D1D1F"}}>{key}</span>
                          <span style={{fontSize:11,color:"#AEAEB2"}}>({monthSessions.length} sessions{monthSelectedCount>0?`, ${monthSelectedCount} selected`:""})</span>
                        </div>
                        <button onClick={(e)=>{e.stopPropagation();selectMonth(key);}} style={{padding:"4px 10px",borderRadius:14,border:"1px solid #5E5CE633",background:"#5E5CE611",color:"#5E5CE6",fontSize:11,fontWeight:600,cursor:"pointer"}}>Select this month</button>
                      </div>
                      {open&&(
                        <div style={{display:"flex",gap:8,flexWrap:"wrap",padding:"0 14px 12px"}}>
                          {monthSessions.map(s=>{
                            const active=invoiceDates.includes(s.date);
                            return(
                              <button key={s.date} onClick={()=>toggle(s.date)} style={{padding:"7px 13px",borderRadius:20,border:`1px solid ${active?"#5E5CE6":"#E5E5EA"}`,background:active?"#5E5CE6":"#F7F7F9",color:active?"#FFFFFF":"#3A3A3C",fontSize:12,fontWeight:600,cursor:"pointer"}}>
                                {active?"✓ ":""}{formatDateShort(s.date)}
                              </button>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  );
                });
              })()}
            </div>

            {invoiceDates.length===0?(
              <div style={{textAlign:"center",padding:"40px 20px",color:"#AEAEB2",fontSize:13}}>No sessions selected yet.</div>
            ):(<>
              <div style={{fontSize:12,color:"#8E8E93",marginBottom:12,padding:"0 2px"}}>{invoiceDates.length} sessions selected · {totalBottles} bottles total · {invoiceData.length} gyms</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:12}}>
                {invoiceData.map((g,i)=>(
                  <div key={g.gym} style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"15px 16px",border:`1px solid ${GYM_COLORS[i%GYM_COLORS.length]}44`}}>
                    <div style={{fontSize:13,fontWeight:700,color:GYM_COLORS[i%GYM_COLORS.length],marginBottom:8}}>{g.gymShort}</div>
                    {g.flavors.map(f=>(
                      <div key={f.flavor} style={{display:"flex",justifyContent:"space-between",fontSize:12,color:"#3A3A3C",padding:"3px 0"}}>
                        <span>{f.flavor}</span>
                        <span style={{fontWeight:700}}>{f.qty} bottles</span>
                      </div>
                    ))}
                    <div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:"#1D1D1F",fontWeight:700,marginTop:6,paddingTop:6,borderTop:"1px solid #F0F0F2"}}>
                      <span>Total</span>
                      <span>{g.total} bottles</span>
                    </div>
                  </div>
                ))}
              </div>
            </>)}
          </>);
        })()}

        {activeTab==="forecast"&&(()=>{
          const proj=M.nextSessionProjection;
          const dirColor=proj.direction==="up"?"#34C759":proj.direction==="down"?"#FF3B30":"#8E8E93";
          const dirArrow=proj.direction==="up"?"▲":proj.direction==="down"?"▼":"▬";
          return(
          <div>
            <SectionTitle>Next Session — Projection</SectionTitle>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:12,marginBottom:8}}>
              <InsightCard icon={<Glyph name="cash" color="#34C759"/>} label="PROJECTED REVENUE" color="#34C759" title={formatRpFull(proj.revenue)} sub={`avg of last ${proj.basisSessions} sessions`}/>
              <InsightCard icon={<Glyph name="bars" color="#AF52DE"/>} label="PROJECTED PROFIT" color="#AF52DE" title={formatRpFull(proj.profit)} sub={proj.deltaPct===null?"not enough history to compare":`${dirArrow} ${Math.abs(proj.deltaPct)}% vs prior sessions`}/>
            </div>

            <SectionTitle>Recommended Quantity</SectionTitle>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))",gap:14,marginBottom:8}}>
              {(M.qtyRec||[]).map((g,i)=><GymReorderCard key={g.gym} g={g} color={GYM_COLORS[i%GYM_COLORS.length]}/>)}
            </div>

            <SectionTitle>Sell Rate by Gym</SectionTitle>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:10,marginBottom:8}}>
              {(M.gymData||[]).filter(g=>g.active).map((g,i)=>(
                <div key={g.gym} style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"12px 14px",border:`1px solid ${GYM_COLORS[i%GYM_COLORS.length]}33`}}>
                  <div style={{fontSize:12,fontWeight:700,color:GYM_COLORS[i%GYM_COLORS.length],marginBottom:4}}>{g.gymShort}</div>
                  <div style={{fontSize:18,fontWeight:700,color:"#1D1D1F"}}>{g.sellRate}%</div>
                  <TrendBadge trend={g.trend}/>
                </div>
              ))}
            </div>

            <SectionTitle>Top Waste Hotspots</SectionTitle>
            {(M.topWaste||[]).length===0 ? (
              <div style={{textAlign:"center",padding:"24px 20px",color:"#AEAEB2",fontSize:13}}>No meaningful waste in the recent window — nice.</div>
            ) : (
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:12}}>
                {M.topWaste.map((w,i)=>(
                  <div key={w.gym+w.flavor} style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"14px 16px",border:"1px solid #FF950044"}}>
                    <div style={{fontSize:12,fontWeight:700,color:"#FF9500",marginBottom:4}}>{w.gymShort} · {w.flavor}</div>
                    <div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:"#3A3A3C"}}>
                      <span>{w.wasteUnits} pcs unsold</span>
                      <span style={{fontWeight:700,color:"#FF9500"}}>{formatRpFull(w.wasteCost)}</span>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
          );
        })()}

        {activeTab==="review"&&(()=>{
          const gymsWithIssues=(M.qtyRec||[])
            .map(g=>({...g,issues:g.flavors.filter(f=>f.signal.code!=="stable")}))
            .filter(g=>g.issues.length>0)
            .map(g=>({...g,topSignal:g.issues.slice().sort((a,b)=>SIGNAL_RANK[a.signal.code]-SIGNAL_RANK[b.signal.code])[0].signal,totalWasteCost:g.issues.reduce((a,f)=>a+f.wasteCost,0)}))
            .sort((a,b)=>SIGNAL_RANK[a.topSignal.code]-SIGNAL_RANK[b.topSignal.code]||b.totalWasteCost-a.totalWasteCost);
          return(
          <div>
            <SectionTitle>Gym Review</SectionTitle>
            {gymsWithIssues.length===0?(
              <div style={{textAlign:"center",padding:"60px 20px",color:"#AEAEB2",fontSize:13}}>Nothing needs attention right now — every gym and flavor looks stable.</div>
            ):(
              <div style={{display:"flex",flexDirection:"column",gap:12}}>
                {gymsWithIssues.map(g=>(
                  <div key={g.gym} style={{background:"rgba(255,255,255,0.6)",backdropFilter:"blur(24px) saturate(180%)",WebkitBackdropFilter:"blur(24px) saturate(180%)",boxShadow:"7px 7px 16px rgba(148,148,180,0.16),-7px -7px 16px rgba(255,255,255,0.75),inset 0 1px 0 rgba(255,255,255,0.5)",borderRadius:18,padding:"18px 20px",borderLeft:`3px solid ${g.topSignal.color}`}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:14}}>
                      <div style={{fontSize:15,fontWeight:700,color:"#1D1D1F"}}>{g.gymShort}</div>
                      <span style={{fontSize:9,fontWeight:700,color:g.topSignal.color,background:g.topSignal.color+"16",padding:"3px 8px",borderRadius:20,letterSpacing:0.3}}>{g.issues.length} flagged</span>
                    </div>
                    {g.issues.map((f,i)=>(
                      <div key={f.flavor} style={{paddingTop:i===0?0:12,marginTop:i===0?0:12,borderTop:i===0?"none":"1px solid rgba(0,0,0,0.07)"}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                          <span style={{fontSize:13,fontWeight:600,color:"#1D1D1F"}}>{f.flavor}</span>
                          <span style={{fontSize:9,fontWeight:700,color:f.signal.color,background:f.signal.color+"16",padding:"2px 7px",borderRadius:20,letterSpacing:0.3}}>{f.signal.label}</span>
                        </div>
                        <div style={{fontSize:12,color:"#6E6E73",lineHeight:1.5}}>
                          {f.signal.code==="trim"&&<div>Waste rate is {f.wasteRatePct}% of what's delivered — recommendation trimmed below the raw average to cut over-delivery.</div>}
                          {f.signal.code==="boost"&&<div>Sells out in {f.stockoutRatePct}% of recent sessions with almost no waste — likely under-stocked, an opportunity to increase.</div>}
                          {f.signal.code==="watch"&&<div>Sales have been declining since {f.declineSince}.</div>}
                          {f.signal.code==="lowdata"&&<div>Only {f.dataPoints} session{f.dataPoints===1?"":"s"} of recent history — recommendation is a rough estimate until more data comes in.</div>}
                        </div>
                        {f.signal.code==="trim"&&f.wasteCost>0&&<div style={{marginTop:4,fontSize:11,color:"#FF3B30",fontWeight:600}}>{formatRpFull(f.wasteCost)} lost to waste this window</div>}
                      </div>
                    ))}
                  </div>
                ))}
              </div>
            )}
          </div>
          );
        })()}

      </div>
      </div>
    </div>
  );
}
