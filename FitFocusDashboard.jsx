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
      let buffer=0,flag=null;
      if(wasteRate>0.20){buffer=-1;flag="high waste — trimmed";}
      else if(wasteRate<0.05&&stockoutRate>0.30){buffer=1;flag="often sells out";}

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

      return{
        flavor:f,flavorShort:f.split(" ")[0],
        avgSold:Math.round(weightedAvg*10)/10,
        recommended:weightedAvg>0?Math.max(1,Math.round(base)+buffer):Math.max(0,Math.round(base)+buffer),
        confidence,flag,
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
const P={sold:"#34C759",waste:"#FF9500",produced:"#5E5CE6",rate:"#0A84FF",forecast:"#AF52DE",
  gyms:["#5E5CE6","#34C759","#FF9500","#0A84FF","#FF2D55","#FFD60A","#30B0C7","#BF5AF2"],
  flavors:["#FF2D55","#5E5CE6","#FFD60A","#0A84FF","#34C759","#FF9500"]};

// ── UI PARTS ─────────────────────────────────────────────────────────────────
function StatCard({label,value,color,icon}){
  return(
    <div style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",border:`1px solid ${color}33`,borderRadius:20,padding:"18px 20px"}}>
      <div style={{fontSize:24,marginBottom:4}}>{icon}</div>
      <div style={{color,fontSize:28,fontWeight:700,lineHeight:1}}>{value}</div>
      <div style={{color:"#8E8E93",fontSize:12,marginTop:5,fontWeight:500}}>{label}</div>
    </div>
  );
}
function InsightCard({icon,label,title,sub,color}){
  return(
    <div style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"14px 16px",border:`1px solid ${color}33`}}>
      <div style={{fontSize:10,color,fontWeight:700,letterSpacing:1,marginBottom:5}}>{icon} {label}</div>
      <div style={{fontSize:16,fontWeight:800,color:"#1D1D1F",marginBottom:2}}>{title}</div>
      <div style={{fontSize:11,color:"#8E8E93"}}>{sub}</div>
    </div>
  );
}
function TrendBadge({trend}){
  if(!trend||trend.direction==="flat") return <span style={{fontSize:10,color:"#AEAEB2",fontWeight:700}}>▬ steady</span>;
  const up=trend.direction==="up";
  return <span style={{fontSize:10,color:up?"#34C759":"#FF3B30",fontWeight:700}}>{up?"▲":"▼"} {Math.abs(trend.delta)}pt {up?"improving":"declining"}</span>;
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
    <div style={{display:"flex",alignItems:"center",gap:10,background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",border:"1px solid #E5E5EA",borderRadius:12,padding:"10px 16px",marginBottom:16}}>
      <span style={{fontSize:12,color:"#8E8E93",fontWeight:600,whiteSpace:"nowrap"}}>📅 Up to:</span>
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
  if(syncLoading){
    return(
      <div style={{minHeight:"100vh",background:"#F5F5F7",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:32,fontFamily:"-apple-system,BlinkMacSystemFont,'SF Pro Display','SF Pro Text',system-ui,sans-serif",letterSpacing:"-0.01em"}}>
        <div style={{fontSize:38,marginBottom:12}}>☁️</div>
        <div style={{color:"#1D1D1F",fontWeight:700,fontSize:15}}>Syncing your data…</div>
      </div>
    );
  }
  return(
    <div style={{minHeight:"100vh",background:"#F5F5F7",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:32,fontFamily:"-apple-system,BlinkMacSystemFont,'SF Pro Display','SF Pro Text',system-ui,sans-serif",letterSpacing:"-0.01em"}}>
      <h1 style={{color:"#1D1D1F",fontSize:28,fontWeight:700,letterSpacing:"-0.02em",margin:"0 0 6px"}}>FitFocus Analytics</h1>
      <p style={{color:"#8E8E93",fontSize:15,margin:"0 0 36px",textAlign:"center"}}>Upload your sales tracker to generate the full dashboard</p>
      {syncError&&<div style={{marginBottom:20,background:"#FFF6E5",color:"#B25E00",padding:"10px 20px",borderRadius:10,fontSize:12,maxWidth:380,textAlign:"center"}}>Couldn't auto-sync ({syncError}) — upload manually below, or check the Apps Script setup.</div>}
      <label onDragOver={onDragOver} onDragLeave={onDragLeave} onDrop={onDrop}
        style={{cursor:"pointer",display:"block",width:"100%",maxWidth:420,border:`2px dashed ${dragging?"#5E5CE6":"#D1D1D6"}`,borderRadius:20,padding:"44px 28px",textAlign:"center",background:dragging?"#5E5CE614":"#FFFFFF",boxShadow:dragging?"none":"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",transition:"all 0.2s"}}>
        <input type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>{if(e.target.files[0])onFile(e.target.files[0]);}}/>
        <div style={{fontSize:38,marginBottom:12}}>{dragging?"📂":"📊"}</div>
        <div style={{color:"#1D1D1F",fontWeight:700,fontSize:15,marginBottom:6}}>{dragging?"Drop it!":"Drag & drop your Excel file"}</div>
        <div style={{color:"#AEAEB2",fontSize:13}}>or click to browse · .xlsx / .xls</div>
      </label>
      {error&&<div style={{marginTop:20,background:"#FFF1F0",color:"#D70015",padding:"10px 20px",borderRadius:10,fontSize:13,maxWidth:380,textAlign:"center"}}>{error}</div>}
    </div>
  );
}

// ── APPS SCRIPT CONFIG ───────────────────────────────────────────────────────
// Paste your deployed Apps Script Web App URL here (see api/AppsScript.gs setup instructions).
// It looks like: https://script.google.com/macros/s/AKfycb.../exec
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyB1F0gCST-evlmstKQi71QxIi1fpe8AaqrhrUcY3Mi__IQ2xCQGRYbtx_DFJ9j2GOI/exec";

// Content-Type must stay "text/plain" (not "application/json") — Apps Script Web Apps don't
// respond to CORS preflight (OPTIONS) requests, and application/json triggers one. text/plain
// keeps this a "simple request" so the browser skips preflight entirely. Apps Script still
// parses the body as JSON on its end regardless of this header.
async function callAppsScript(payload){

  const params = new URLSearchParams();

  Object.entries(payload).forEach(([key,value])=>{
    if(value!==undefined && value!==null){
      params.append(key,value);
    }
  });

  const res = await fetch(
    `${APPS_SCRIPT_URL}?${params.toString()}`,
    {
      method:"GET"
    }
  );

  if(!res.ok){
    throw new Error(`HTTP ${res.status}`);
  }

  return await res.json();
}

// ── LOGIN ─────────────────────────────────────────────────────────────────────
function LoginScreen({onSubmit,loading,error}){
  const [email,setEmail]=useState("");
  const [password,setPassword]=useState("");
  return(
    <div style={{minHeight:"100vh",background:"#F5F5F7",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:32,fontFamily:"-apple-system,BlinkMacSystemFont,'SF Pro Display','SF Pro Text',system-ui,sans-serif",letterSpacing:"-0.01em"}}>
      <div style={{width:44,height:44,borderRadius:12,background:"#1D1D1F",display:"flex",alignItems:"center",justifyContent:"center",marginBottom:20}}>
        <span style={{color:"#fff",fontSize:18,fontWeight:800,letterSpacing:"-0.03em"}}>FF</span>
      </div>
      <h1 style={{color:"#1D1D1F",fontSize:20,fontWeight:700,letterSpacing:"-0.02em",margin:"0 0 28px"}}>Sign in to FitFocus</h1>
      <form onSubmit={e=>{e.preventDefault();if(email&&password&&!loading)onSubmit(email,password);}} style={{width:"100%",maxWidth:320}}>
        <input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="Email" autoFocus autoCapitalize="none" autoCorrect="off"
          style={{width:"100%",boxSizing:"border-box",padding:"12px 16px",borderRadius:12,border:"1px solid #D1D1D6",fontSize:15,fontFamily:"inherit",marginBottom:10,outline:"none"}}/>
        <input type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="Password"
          style={{width:"100%",boxSizing:"border-box",padding:"12px 16px",borderRadius:12,border:"1px solid #D1D1D6",fontSize:15,fontFamily:"inherit",marginBottom:14,outline:"none"}}/>
        <button type="submit" disabled={loading||!email||!password}
          style={{width:"100%",padding:"12px 16px",borderRadius:12,border:"none",cursor:loading?"default":"pointer",fontSize:15,fontWeight:700,fontFamily:"inherit",
            background:loading?"#AEAEB2":"#1D1D1F",color:"#fff"}}>
          {loading?"Signing in…":"Sign in"}
        </button>
      </form>
      {error&&<div style={{marginTop:16,background:"#FFF1F0",color:"#D70015",padding:"10px 20px",borderRadius:10,fontSize:13,maxWidth:320,textAlign:"center"}}>{error}</div>}
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
    {id:"profit",label:"Profit/Session"},
    {id:"invoice",label:"Invoice"},
    {id:"forecast",label:"🔮"},
  ];

  const GYM_COLORS=P.gyms;
  const FLAVOR_COLORS=P.flavors;

  return(
    <div style={{minHeight:"100vh",background:"#F5F5F7",color:"#1D1D1F",fontFamily:"-apple-system,BlinkMacSystemFont,'SF Pro Display','SF Pro Text',system-ui,sans-serif",letterSpacing:"-0.01em",paddingBottom:60}}>
      {/* Header */}
      <div style={{background:"#FFFFFF",borderBottom:"1px solid #E5E5EA",padding:"20px 26px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12,position:"sticky",top:0,zIndex:10,backdropFilter:"saturate(180%) blur(20px)"}}>
        <div>
          <div style={{display:"flex",alignItems:"center",gap:9,marginBottom:4}}>
            <span style={{fontSize:19,fontWeight:700,letterSpacing:"-0.02em"}}>FitFocus</span>
            <span style={{fontSize:10,color:"#5E5CE6",fontWeight:700,letterSpacing:1.2,background:"#5E5CE614",padding:"3px 7px",borderRadius:6}}>ANALYTICS</span>
          </div>
          <div style={{color:"#8E8E93",fontSize:12,fontWeight:500}}>{fileName} · {M.sessionCount} sessions · {M.activeGymCount} active gyms{M.cutGymCount>0?` (+${M.cutGymCount} discontinued)`:""}{cutoff?` · up to ${new Date(cutoff).toLocaleDateString("en-US",{day:"numeric",month:"short",year:"numeric"})}`:""}
          </div>
          {lastSynced&&<div style={{color:"#AEAEB2",fontSize:11,marginTop:2}}>Synced {lastSynced.toLocaleTimeString([],{hour:"2-digit",minute:"2-digit"})}{syncError?` · last refresh failed (${syncError})`:""}</div>}
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={syncFromScript} disabled={syncLoading}
            style={{cursor:syncLoading?"default":"pointer",background:"#F5F5F7",border:"1px solid #E5E5EA",borderRadius:980,padding:"8px 14px",fontSize:13,fontWeight:600,color:"#1D1D1F",display:"flex",alignItems:"center",gap:6}}>
            {syncLoading?"⏳":"🔄"} {syncLoading?"Syncing…":"Refresh"}
          </button>
          <label onDragOver={onDragOver} onDragLeave={onDragLeave} onDrop={onDrop}
            style={{cursor:"pointer",background:"#F5F5F7",border:"1px solid #E5E5EA",borderRadius:980,padding:"8px 16px",fontSize:13,fontWeight:600,color:"#1D1D1F",display:"flex",alignItems:"center",gap:7,transition:"background 0.15s"}}>
            <input type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>{if(e.target.files[0])processFile(e.target.files[0]);}}/>
            Upload new file
          </label>
        </div>
      </div>

      {/* Tabs — segmented control */}
      <div style={{padding:"16px 26px 0"}}>
        <div style={{display:"inline-flex",gap:2,background:"#E9E9EB",borderRadius:12,padding:3,overflowX:"auto",maxWidth:"100%"}}>
          {tabs.map(t=>(
            <button key={t.id} onClick={()=>setActiveTab(t.id)}
              style={{padding:"7px 16px",borderRadius:9,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit",whiteSpace:"nowrap",
                background:activeTab===t.id?"#FFFFFF":"transparent",
                color:activeTab===t.id?"#1D1D1F":"#6E6E73",
                boxShadow:activeTab===t.id?"0 1px 3px rgba(0,0,0,0.1)":"none",
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
              <div style={{fontSize:11,color:"#FF3B30",fontWeight:700,marginBottom:6}}>⚠️ {M.anomalies.length} DATA ANEH DI EXCEL</div>
              <div style={{fontSize:12,color:"#6E6E73",marginBottom:8}}>Remaining exceeds Delivered (data entry error) — counted as 0 sold for now, but should be verified in the source file:</div>
              {M.anomalies.slice(0,6).map((a,i)=>(
                <div key={i} style={{fontSize:11,color:"#8E8E93",padding:"2px 0"}}>{a.date} · {gymShortName(a.gym)} · {a.flavor} — Delivered {a.delivered}, Remaining {a.remaining}</div>
              ))}
              {M.anomalies.length>6&&<div style={{fontSize:11,color:"#AEAEB2",marginTop:2}}>+{M.anomalies.length-6} more</div>}
            </div>
          )}
          {M.deliveryFallbackSessions.length>0&&(
            <div style={{background:"#FF950011",border:"1px solid #FF950033",borderRadius:18,padding:"14px 18px",marginBottom:12}}>
              <div style={{fontSize:11,color:"#FF9500",fontWeight:700,marginBottom:6}}>⚠️ {M.deliveryFallbackSessions.length} SESI PAKAI DELIVERY COST DEFAULT</div>
              <div style={{fontSize:12,color:"#6E6E73"}}>No "Driver" column found in the sheet for: {M.deliveryFallbackSessions.join(", ")}. Defaulted to Rp {DELIVERY_COST_FALLBACK.toLocaleString("en-US")} — check the sheet if the real cost differed.</div>
            </div>
          )}
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:12,marginBottom:12}}>
            <StatCard icon="📦" label="Total Produced" value={M.totalProduced} color="#5E5CE6"/>
            <StatCard icon="✅" label="Total Sold" value={M.totalSold} color="#34C759"/>
            <StatCard icon="🗑️" label="Total Waste" value={M.totalWaste} color="#FF9500"/>
            <StatCard icon="📈" label="Avg Sell Rate" value={`${M.avgSellRate}%`} color="#0A84FF"/>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:12,marginBottom:4}}>
            <InsightCard icon="🏆" label="BEST GYM" color="#34C759" title={bestGym.gym} sub={`${bestGym.sellRate}% sell rate · ${bestGym.sold} sold`}/>
            <InsightCard icon="⚠️" label="MOST WASTE GYM" color="#FF9500" title={worstGym.gym} sub={`${worstGym.waste} units wasted`}/>
            <InsightCard icon="⭐" label="TOP FLAVOR" color="#34C759" title={bestFlavor.flavor} sub={`${bestFlavor.sellRate}% sell rate`}/>
            {(()=>{
              const improving=[...activeGymData].sort((a,b)=>(b.trend?.delta||0)-(a.trend?.delta||0))[0];
              const declining=[...activeGymData].sort((a,b)=>(a.trend?.delta||0)-(b.trend?.delta||0))[0];
              const show=improving&&improving.trend?.direction==="up"?improving:declining;
              const isUp=show===improving&&improving?.trend?.direction==="up";
              return show&&show.trend&&show.trend.direction!=="flat"
                ?<InsightCard icon={isUp?"📈":"📉"} label={isUp?"MOST IMPROVED":"MOST DECLINING"} color={isUp?"#34C759":"#FF3B30"} title={show.gym} sub={`${isUp?"+":""}${show.trend.delta}pt sell rate vs earlier sessions`}/>
                :<InsightCard icon="⚠️" label="MOST WASTED FLAVOR" color="#FF9500" title={worstFlavor.flavor} sub={`${worstFlavor.waste} units wasted`}/>;
            })()}
          </div>
          {cutGymData.length>0&&(
            <div style={{fontSize:11,color:"#AEAEB2",marginBottom:8,padding:"0 2px"}}>
              ℹ️ {cutGymData.map(g=>g.gymShort).join(", ")} are no longer supplied (not present in the last {M.recentWindow} sessions) — excluded from rankings & recommendations, though their data remains in history & trends.
            </div>
          )}
          <SectionTitle>Sold vs Waste</SectionTitle>
          <TimeframeSelector value={timeframe} onChange={setTimeframe}/>
          <div style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"16px 4px 8px"}}>
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
          <div style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"16px 4px 8px"}}>
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
              <div key={g.gym} style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"15px 16px",border:`1px solid ${GYM_COLORS[i%GYM_COLORS.length]}44`}}>
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
              <div key={gym} style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"10px 6px 4px",border:`1px solid ${GYM_COLORS[i%GYM_COLORS.length]}33`}}>
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
              <div key={f.flavor} style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"15px 16px",border:`1px solid ${FLAVOR_COLORS[i%FLAVOR_COLORS.length]}44`}}>
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
          <div style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"16px 4px 8px",marginBottom:8}}>
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
          <div style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"16px 4px 8px"}}>
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
          <div style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"16px 4px 8px",marginBottom:8}}>
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
          <div style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,overflow:"auto"}}>
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
              <div key={label} style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"14px 16px",border:`1px solid ${color}33`}}>
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
          <div style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"16px 4px 8px",marginBottom:8}}>
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
              <div key={g.gym} style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"14px 16px",border:`1px solid ${g.profit>=0?"#34C75933":"#FF3B3033"}`}}>
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

          <SectionTitle>Optimal Quantity Recommendation</SectionTitle>
          {M.qtyRec.map((g,i)=>(
            <div key={g.gym} style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"14px 16px",marginBottom:12,border:`1px solid ${GYM_COLORS[i%GYM_COLORS.length]}33`}}>
              <div style={{fontSize:12,fontWeight:700,color:GYM_COLORS[i%GYM_COLORS.length],marginBottom:10}}>{g.gym}</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(110px,1fr))",gap:8}}>
                {g.flavors.map(f=>(
                  <div key={f.flavor} style={{background:"#F5F5F7",borderRadius:10,padding:"10px 12px",border:"1px solid #E5E5EA"}}>
                    <div style={{fontSize:11,color:"#8E8E93",marginBottom:4}}>{f.flavor}</div>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <div>
                        <div style={{fontSize:9,color:"#AEAEB2"}}>Avg sold</div>
                        <div style={{fontSize:13,fontWeight:700,color:"#6E6E73"}}>{f.avgSold}</div>
                      </div>
                      <div style={{textAlign:"right"}}>
                        <div style={{fontSize:9,color:"#AEAEB2"}}>Rec.</div>
                        <div style={{fontSize:16,fontWeight:700,color:"#0A84FF"}}>{f.recommended}</div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ))}

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
                <div key={label} style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"14px 16px",border:`1px solid ${color}33`}}>
                  <div style={{fontSize:11,color:"#8E8E93",marginBottom:4}}>{label}</div>
                  <div style={{fontSize:18,fontWeight:700,color,lineHeight:1}}>{isCount?value:formatRpFull(value)}</div>
                  {sub&&<div style={{fontSize:10,color:"#AEAEB2",marginTop:5}}>{sub}</div>}
                </div>
              ))}
            </div>

            <SectionTitle>Profit per Session</SectionTitle>
            <div style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"16px 4px 8px",marginBottom:12}}>
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
                  <div key={g.gym} style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"15px 16px",border:`1px solid ${GYM_COLORS[i%GYM_COLORS.length]}44`}}>
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
              <InsightCard icon="💰" label="PROJECTED REVENUE" color="#34C759" title={formatRpFull(proj.revenue)} sub={`avg of last ${proj.basisSessions} sessions`}/>
              <InsightCard icon="📊" label="PROJECTED PROFIT" color="#AF52DE" title={formatRpFull(proj.profit)} sub={proj.deltaPct===null?"not enough history to compare":`${dirArrow} ${Math.abs(proj.deltaPct)}% vs prior sessions`}/>
            </div>

            <SectionTitle>Recommended Qty per Gym</SectionTitle>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(210px,1fr))",gap:12,marginBottom:8}}>
              {(M.qtyRec||[]).map((g,i)=>(
                <div key={g.gym} style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"15px 16px",border:`1px solid ${GYM_COLORS[i%GYM_COLORS.length]}44`}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                    <div style={{fontSize:13,fontWeight:700,color:GYM_COLORS[i%GYM_COLORS.length]}}>{g.gymShort}</div>
                    <TrendBadge trend={g.trend}/>
                  </div>
                  {g.flavors.map(f=>{
                    const confColor=f.confidence==="high"?"#34C759":f.confidence==="medium"?"#FF9500":"#AEAEB2";
                    return(
                    <div key={f.flavor} style={{padding:"5px 0",borderTop:"1px solid #F2F2F7"}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:12,color:"#3A3A3C"}}>
                        <span>{f.flavor}</span>
                        <span style={{display:"flex",alignItems:"center",gap:6}}>
                          <span style={{fontSize:9,fontWeight:700,color:confColor,background:confColor+"18",padding:"2px 6px",borderRadius:20,textTransform:"uppercase"}}>{f.confidence}</span>
                          <span style={{fontWeight:700}}>{f.recommended} pcs</span>
                        </span>
                      </div>
                      {(f.flag||f.declineSince)&&(
                        <div style={{fontSize:10,color:f.flag&&f.flag.includes("waste")?"#FF3B30":"#FF9500",marginTop:1}}>
                          {f.flag}{f.flag&&f.declineSince?" · ":""}{f.declineSince?`declining since ${f.declineSince}`:""}
                        </div>
                      )}
                    </div>
                    );
                  })}
                </div>
              ))}
            </div>

            <SectionTitle>Sell Rate by Gym</SectionTitle>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:10,marginBottom:8}}>
              {(M.gymData||[]).filter(g=>g.active).map((g,i)=>(
                <div key={g.gym} style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"12px 14px",border:`1px solid ${GYM_COLORS[i%GYM_COLORS.length]}33`}}>
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
                  <div key={w.gym+w.flavor} style={{background:"#FFFFFF",boxShadow:"0 1px 2px rgba(0,0,0,0.03),0 6px 20px rgba(0,0,0,0.05)",borderRadius:18,padding:"14px 16px",border:"1px solid #FF950044"}}>
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

      </div>
    </div>
  );
}
