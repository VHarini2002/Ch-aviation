import { useState, useCallback, useEffect, useRef } from "react";
import {
  BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, AreaChart, Area
} from "recharts";

// SheetJS loaded dynamically at runtime (ESM CDN, not available as npm package in CRA)
const XLSX_SCRIPT = "https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs";

/* ─── DESIGN TOKENS ─── */
const T = {
  bg: "#04080F",
  surface: "#080E1A",
  card: "#0C1524",
  cardBorder: "#132036",
  cardHover: "#111E30",
  blue: "#0070F3",
  blueGlow: "#0070F344",
  cyan: "#00D4FF",
  orange: "#FF6B00",
  orangeGlow: "#FF6B0033",
  green: "#00E676",
  text: "#E8F2FF",
  muted: "#5A7A9A",
  dim: "#2A3F5F",
  font: "'IBM Plex Mono', 'Courier New', monospace",
  fontDisplay: "'Bebas Neue', 'Impact', sans-serif",
  fontBody: "'IBM Plex Sans', 'Segoe UI', sans-serif",
};

const STATUS_COLORS = {
  "Active": "#00E676", "Not yet delivered": "#0070F3", "Stored": "#FFD600",
  "Maintenance": "#FF9800", "Scrapped": "#546E7A", "Retired": "#7B1FA2",
  "Not built (order cancelled)": "#F44336", "On display": "#00BCD4",
  "Beyond repair": "#B71C1C", "Crashed": "#E53935",
};

const OP_STATUS_COLORS = { Active: "#34d399", Stored: "#fbbf24", Maintenance: "#fb923c", Repair: "#f87171", Retired: "#a78bfa", "Not yet delivered": "#60a5fa", Unknown: "#94a3b8" };
const OP_PIE_COLORS = ["#38bdf8", "#34d399", "#a78bfa", "#fbbf24", "#f87171", "#fb923c", "#c084fc"];
const OP_ENGINE_TYPE_COLORS = { CFM: "#38bdf8", LEAP: "#f97316" };
const OP_MONTHS = { jan:1,january:1,feb:2,february:2,mar:3,march:3,apr:4,april:4,may:5,jun:6,june:6,jul:7,july:7,aug:8,august:8,sep:9,sept:9,september:9,oct:10,october:10,nov:11,november:11,dec:12,december:12 };

/* ─── UTILS ─── */
function countBy(arr, key) {
  return arr.reduce((acc, r) => {
    const v = r[key] || "Unknown";
    acc[v] = (acc[v] || 0) + 1;
    return acc;
  }, {});
}
function toNumber(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const s = String(v).trim().replace(/,/g, "");
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}
function cleanDimValue(v) {
  const s = String(v ?? "").trim();
  if (!s) return "";
  // Treat these as empty/invalid dimension values.
  if (["–", "undefined", "Unknown", "Unassigned", "null", "NaN"].includes(s)) return "";
  return s;
}
function toChartData(obj, limit = 10) {
  return Object.entries(obj)
    .filter(([k]) => !["–","undefined","Unknown","Unassigned"].includes(k))
    .sort((a, b) => b[1] - a[1])
    .slice(0, limit)
    .map(([name, value]) => ({ name, value }));
}
function processExcel(wb) {
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const XLSX = window.__XLSX__;
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const headers = raw[0];
  const rows = raw.slice(1).map(r => {
    let obj = {};
    headers.forEach((h, i) => { obj[h] = r[i]; });
    return obj;
  });
  const mapWithEngines = (arr) => arr.map(r => ({ ...r, __engines: toNumber(r["Number of Engines"]) }));
  const cfm56 = mapWithEngines(rows.filter(r => r["Engine Full Name"] && String(r["Engine Full Name"]).includes("CFM56")));
  const leap  = mapWithEngines(rows.filter(r => r["Engine Full Name"] && String(r["Engine Full Name"]).includes("LEAP")));

  const cfm56Fam = {}, leapFam = {};
  cfm56.forEach(r => {
    const n = String(r["Engine Full Name"]);
    const f = n.includes("CFM56-7B") ? "CFM56-7B" : n.includes("CFM56-5B") ? "CFM56-5B"
            : n.includes("CFM56-5C") ? "CFM56-5C" : n.includes("CFM56-5A") ? "CFM56-5A" : "CFM56-3";
    cfm56Fam[f] = (cfm56Fam[f] || 0) + 1;
  });
  leap.forEach(r => {
    const n = String(r["Engine Full Name"]);
    const f = n.includes("LEAP-1A") ? "LEAP-1A (A320neo)" : n.includes("LEAP-1B") ? "LEAP-1B (B737 MAX)" : "LEAP-1C (C919)";
    leapFam[f] = (leapFam[f] || 0) + 1;
  });

  const cfm56OperatorFull = Object.entries(countBy(cfm56, "Operator Name"))
    .filter(([k]) => !["–","undefined","Unknown","Unassigned"].includes(k))
    .sort((a, b) => b[1] - a[1])
    .map(([name, value]) => ({ name, value }));
  const leapOperatorFull = Object.entries(countBy(leap, "Operator Name"))
    .filter(([k]) => !["–","undefined","Unknown","Unassigned"].includes(k))
    .sort((a, b) => b[1] - a[1])
    .map(([name, value]) => ({ name, value }));

  return {
    total: rows.length, cfm56Total: cfm56.length, leapTotal: leap.length,
    cfm56Active: cfm56.filter(r => r.Status === "Active").length,
    leapActive:  leap.filter(r => r.Status === "Active").length,
    leapPending: leap.filter(r => r.Status === "Not yet delivered").length,
    cfm56Status: toChartData(countBy(cfm56, "Status"), 8),
    leapStatus:  toChartData(countBy(leap,  "Status"), 8),
    cfm56Families: Object.entries(cfm56Fam).sort((a,b)=>b[1]-a[1]).map(([name,value])=>({name,value})),
    leapFamilies:  Object.entries(leapFam).sort((a,b)=>b[1]-a[1]).map(([name,value])=>({name,value})),
    cfm56Aircraft:   toChartData(countBy(cfm56, "Aircraft Varient"), 8),
    leapAircraft:    toChartData(countBy(leap,  "Aircraft Varient"), 8),
    cfm56Continents: toChartData(countBy(cfm56, "Operator Continent"), 7),
    leapContinents:  toChartData(countBy(leap,  "Operator Continent"), 7),
    cfm56Operators:  toChartData(countBy(cfm56, "Operator Name"), 8),
    leapOperators:   toChartData(countBy(leap,  "Operator Name"), 8),
    cfm56OperatorsFull: cfm56OperatorFull,
    leapOperatorsFull:  leapOperatorFull,
    cfm56Ownership:  toChartData(countBy(cfm56, "Ownership Type"), 6),
    leapOwnership:   toChartData(countBy(leap,  "Ownership Type"), 6),
    cfm56Variants:   toChartData(countBy(cfm56, "Engine Full Name"), 10),
    leapVariants:    toChartData(countBy(leap,  "Engine Full Name"), 10),
    cfm56Rows: cfm56,
    leapRows: leap,
  };
}

const opHeader = (v) => String(v || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, " ");
const opNum = (v) => { const n = Number(String(v ?? "").replace(/,/g, "").trim()); return Number.isFinite(n) ? n : 0; };
const opText = (v, f = "Unknown") => String(v ?? "").trim() || f;
const opMonthSort = (v) => OP_MONTHS[String(v || "").toLowerCase()] || OP_MONTHS[String(v || "").toLowerCase().slice(0, 3)] || Number.MAX_SAFE_INTEGER;
const opFmt = (v) => Math.round(v || 0).toLocaleString();
function opPick(obj, keys) { for (const k of keys) if (obj[k] !== undefined && obj[k] !== null && obj[k] !== "") return obj[k]; return ""; }
function opTypeFrom(model) { return String(model || "").toUpperCase().includes("LEAP") ? "LEAP" : "CFM"; }
function opFamilyFrom(model) { const m = String(model || "").toUpperCase(); if (m.includes("CFM56-3")) return "CFM56-3"; if (m.includes("CFM56-5")) return "CFM56-5"; if (m.includes("CFM56-7")) return "CFM56-7"; if (m.includes("LEAP-1A")) return "LEAP-1A"; if (m.includes("LEAP-1B")) return "LEAP-1B"; if (m.includes("LEAP-1C")) return "LEAP-1C"; if (m.includes("LEAP")) return "LEAP"; if (m.includes("CFM56")) return "CFM56"; return "Other"; }
function opDashFrom(model) { const m = String(model || "").toUpperCase().match(/-(\d[A-Z]?)/); return m ? m[1] : "Unknown"; }
function opStatusFrom(v) { const s = String(v || "").trim(); const l = s.toLowerCase(); if (!s) return "Unknown"; if (l.includes("active")) return "Active"; if (l.includes("store")) return "Stored"; if (l.includes("maint")) return "Maintenance"; if (l.includes("repair")) return "Repair"; if (l.includes("retired")) return "Retired"; if (l.includes("not yet delivered")) return "Not yet delivered"; return s; }
function opMonthFrom(raw, file) { const d = String(raw || "").trim().toLowerCase(); if (OP_MONTHS[d]) return d.slice(0,3).toUpperCase(); const f = String(file || "").toLowerCase(); const k = Object.keys(OP_MONTHS).find((x) => f.includes(x)); return k ? k.slice(0,3).toUpperCase() : "Unknown"; }
function opParseWorkbook(file, wb) {
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = window.__XLSX__.utils.sheet_to_json(sheet, { defval: "" });
  return rows.map((source, i) => {
    const x = {};
    Object.entries(source).forEach(([k, v]) => { x[opHeader(k)] = v; });
    const model = opPick(x, ["engine full name", "engine model", "engine", "model"]);
    return {
      id: `${file.name}-${i}`,
      month: opMonthFrom(opPick(x, ["month", "period"]), file.name),
      operator: opText(opPick(x, ["operator name", "operator", "airline"])),
      continent: opText(opPick(x, ["operator continent", "continent", "region"])),
      country: opText(opPick(x, ["operator country", "country"])),
      engineModel: opText(model),
      engineType: opTypeFrom(model),
      family: opFamilyFrom(model),
      dash: opDashFrom(model),
      status: opStatusFrom(opPick(x, ["status", "engine status"])),
      engines: opNum(opPick(x, ["number of engines", "engines", "engine count", "count"])) || 1
    };
  });
}

/* ─── REUSABLE COMPONENTS ─── */
const Tooltip_ = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background:"#0C1524", border:`1px solid ${T.dim}`, borderRadius:8, padding:"10px 14px", fontFamily:T.font, fontSize:12 }}>
      <div style={{ color:T.muted, marginBottom:4 }}>{label}</div>
      {payload.map((p,i) => <div key={i} style={{ color: p.fill||p.color, fontWeight:700 }}>{p.name}: {p.value?.toLocaleString()}</div>)}
    </div>
  );
};

function Card({ children, style, glow }) {
  return (
    <div style={{
      background: T.card, border:`1px solid ${T.cardBorder}`,
      borderRadius:12, padding:20,
      boxShadow: glow ? `0 0 24px ${glow}` : "none",
      ...style
    }}>{children}</div>
  );
}

function ChartLabel({ children }) {
  return <div style={{ fontFamily:T.font, fontSize:11, color:T.muted, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:14 }}>{children}</div>;
}

function Tag({ children, color }) {
  return (
    <span style={{
      fontFamily:T.font, fontSize:10, fontWeight:700, letterSpacing:"0.08em",
      padding:"3px 9px", borderRadius:4, border:`1px solid ${color}55`, color, background:`${color}11`
    }}>{children}</span>
  );
}

/* ─── NAV ─── */
function Nav({ activeSection, onNav, fileName, onReset }) {
  const [scrolled, setScrolled] = useState(false);
  useEffect(() => {
    const h = () => setScrolled(window.scrollY > 40);
    window.addEventListener("scroll", h);
    return () => window.removeEventListener("scroll", h);
  }, []);
  return (
    <nav style={{
      position:"fixed", top:0, left:0, right:0, zIndex:1000,
      background: scrolled ? "rgba(4,8,15,0.96)" : "transparent",
      backdropFilter: scrolled ? "blur(12px)" : "none",
      borderBottom: scrolled ? `1px solid ${T.cardBorder}` : "1px solid transparent",
      transition:"all 0.3s", padding:"0 40px",
      display:"flex", alignItems:"center", height:64,
    }}>
      {/* Logo */}
      <div style={{ display:"flex", alignItems:"center", gap:12, cursor:"pointer" }} onClick={() => onNav("hero")}>
        <div style={{
          width:36, height:36, borderRadius:8, background:`linear-gradient(135deg, ${T.blue}, ${T.orange})`,
          display:"flex", alignItems:"center", justifyContent:"center", fontSize:18
        }}>✈</div>
        <div>
          <div style={{ fontFamily:T.fontDisplay, fontSize:20, letterSpacing:2, color:T.text, lineHeight:1 }}>CH AVIATION</div>
          <div style={{ fontFamily:T.font, fontSize:9, color:T.muted, letterSpacing:"0.15em" }}>ENGINE INTELLIGENCE</div>
        </div>
      </div>
      <div style={{ flex:1 }} />
      {/* Links */}
      {["about","insights","dashboard"].map(s => (
        <button key={s} onClick={() => onNav(s)} style={{
          background:"none", border:"none", cursor:"pointer", fontFamily:T.font,
          fontSize:11, letterSpacing:"0.12em", textTransform:"uppercase",
          color: activeSection === s ? T.cyan : T.muted,
          padding:"8px 18px",
          borderBottom: activeSection === s ? `2px solid ${T.cyan}` : "2px solid transparent",
          transition:"color 0.2s",
        }}>{s}</button>
      ))}
      {fileName && (
        <div style={{
          fontFamily:T.font, fontSize:10, color:T.muted, background:T.surface,
          border:`1px solid ${T.dim}`, borderRadius:6, padding:"5px 12px", marginLeft:16,
          display:"flex", alignItems:"center", gap:8
        }}>
          <span style={{ color:T.green }}>●</span> {fileName.length > 22 ? fileName.slice(0,22)+"…" : fileName}
          <span style={{ color:T.muted, cursor:"pointer", marginLeft:4 }} onClick={onReset}>✕</span>
        </div>
      )}
    </nav>
  );
}

/* ─── HERO ─── */
function Hero({ onNav }) {
  return (
    <section id="hero" style={{
      minHeight:"100vh", display:"flex", flexDirection:"column",
      alignItems:"center", justifyContent:"center",
      background: T.bg, position:"relative", overflow:"hidden",
      padding:"80px 24px 40px"
    }}>
      {/* Grid bg */}
      <div style={{
        position:"absolute", inset:0, zIndex:0,
        backgroundImage:`
          linear-gradient(rgba(0,112,243,0.07) 1px, transparent 1px),
          linear-gradient(90deg, rgba(0,112,243,0.07) 1px, transparent 1px)`,
        backgroundSize:"40px 40px",
      }} />
      {/* Glow orbs */}
      <div style={{ position:"absolute", top:"20%", left:"15%", width:400, height:400, borderRadius:"50%", background:`radial-gradient(circle, ${T.blueGlow} 0%, transparent 70%)`, zIndex:0 }} />
      <div style={{ position:"absolute", bottom:"15%", right:"12%", width:320, height:320, borderRadius:"50%", background:`radial-gradient(circle, ${T.orangeGlow} 0%, transparent 70%)`, zIndex:0 }} />

      <div style={{ position:"relative", zIndex:1, textAlign:"center", maxWidth:860 }}>
        {/* Badge */}
        <div style={{ display:"inline-flex", alignItems:"center", gap:8, background:`${T.blue}15`, border:`1px solid ${T.blue}44`, borderRadius:20, padding:"6px 16px", marginBottom:32 }}>
          <span style={{ width:6, height:6, borderRadius:"50%", background:T.cyan, display:"inline-block", boxShadow:`0 0 8px ${T.cyan}` }} />
          <span style={{ fontFamily:T.font, fontSize:11, color:T.cyan, letterSpacing:"0.12em" }}>AVIATION ENGINE MASTER DATA PLATFORM</span>
        </div>

        {/* Title */}
        <h1 style={{
          fontFamily:T.fontDisplay, fontSize:"clamp(52px,9vw,108px)",
          letterSpacing:6, color:T.text, lineHeight:0.92, margin:"0 0 8px",
        }}>CH AVIATION</h1>
        <div style={{
          fontFamily:T.fontDisplay, fontSize:"clamp(22px,4vw,48px)",
          letterSpacing:8, background:`linear-gradient(90deg, ${T.blue}, ${T.cyan})`,
          WebkitBackgroundClip:"text", WebkitTextFillColor:"transparent", marginBottom:32
        }}>ENGINE INTELLIGENCE</div>

        <p style={{ fontFamily:T.fontBody, fontSize:17, color:T.muted, maxWidth:600, margin:"0 auto 48px", lineHeight:1.7 }}>
          The world's most comprehensive CFM56 & LEAP engine fleet database.
          Upload your master data and unlock real-time insights across status,
          operators, geography and ownership.
        </p>

        {/* CTAs */}
        <div style={{ display:"flex", gap:16, justifyContent:"center", flexWrap:"wrap" }}>
          <button onClick={() => onNav("dashboard")} style={{
            fontFamily:T.font, fontSize:12, letterSpacing:"0.12em", textTransform:"uppercase",
            background:`linear-gradient(135deg, ${T.blue}, #0050C0)`,
            color:"#fff", border:"none", borderRadius:8, padding:"14px 32px", cursor:"pointer",
            boxShadow:`0 0 20px ${T.blueGlow}`,
          }}>Launch Dashboard →</button>
          <button onClick={() => onNav("insights")} style={{
            fontFamily:T.font, fontSize:12, letterSpacing:"0.12em", textTransform:"uppercase",
            background:"transparent", color:T.cyan,
            border:`1px solid ${T.cyan}55`, borderRadius:8, padding:"14px 32px", cursor:"pointer",
          }}>View Insights</button>
        </div>

        {/* Stats row */}
        <div style={{ display:"flex", gap:0, justifyContent:"center", marginTop:72, flexWrap:"wrap" }}>
          {[
            { n:"20,922+", l:"Engine Records" },
            { n:"CFM56 & LEAP", l:"Engine Families" },
            { n:"6", l:"Global Regions" },
            { n:"Real-Time", l:"Data Processing" },
          ].map((s, i) => (
            <div key={i} style={{
              padding:"20px 36px", borderLeft: i===0?"none":`1px solid ${T.dim}`,
              textAlign:"center"
            }}>
              <div style={{ fontFamily:T.fontDisplay, fontSize:28, color:T.text, letterSpacing:2 }}>{s.n}</div>
              <div style={{ fontFamily:T.font, fontSize:10, color:T.muted, letterSpacing:"0.1em", marginTop:4 }}>{s.l}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Scroll hint */}
      <div style={{ position:"absolute", bottom:32, left:"50%", transform:"translateX(-50%)", display:"flex", flexDirection:"column", alignItems:"center", gap:6 }}>
        <div style={{ fontFamily:T.font, fontSize:9, color:T.muted, letterSpacing:"0.2em" }}>SCROLL</div>
        <div style={{ width:1, height:40, background:`linear-gradient(to bottom, ${T.muted}, transparent)` }} />
      </div>
    </section>
  );
}

/* ─── ABOUT ─── */
function About() {
  const features = [
    { icon:"🔵", title:"CFM56 Coverage", desc:"Complete registry of the CFM56-3, -5A, -5B, -5C, and -7B sub-families powering the B737 Classic/NG and A320ceo fleets worldwide.", tag:"LEGACY FLEET" },
    { icon:"🟠", title:"LEAP Engine Tracking", desc:"Full visibility into LEAP-1A, LEAP-1B, and LEAP-1C variants for the A320neo, B737 MAX, and COMAC C919 next-gen programmes.", tag:"NEO FLEET" },
    { icon:"🌍", title:"Global Operator Data", desc:"Operator intelligence across 6 continents — scheduled carriers, charters, cargo operators, and government fleets.", tag:"WORLDWIDE" },
    { icon:"📊", title:"Fleet Status Monitoring", desc:"Track active, stored, in-maintenance, scrapped, and on-order engines with ownership and lease classification.", tag:"LIVE STATUS" },
    { icon:"✈️", title:"Aircraft Type Mapping", desc:"Precise engine-to-airframe assignments covering all variants of the B737 and A320 family across operators.", tag:"AIRFRAME LINK" },
    { icon:"🔐", title:"Ownership Intelligence", desc:"Distinguish operating leases, financial leases, company-owned, and remarketed assets with manager and owner data.", tag:"ASSET DATA" },
  ];
  return (
    <section id="about" style={{ background:T.surface, padding:"100px 40px", position:"relative" }}>
      <div style={{ maxWidth:1100, margin:"0 auto" }}>
        <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:16 }}>
          <div style={{ width:3, height:32, background:T.blue, borderRadius:2 }} />
          <div style={{ fontFamily:T.font, fontSize:11, color:T.blue, letterSpacing:"0.2em" }}>ABOUT THE PLATFORM</div>
        </div>
        <h2 style={{ fontFamily:T.fontDisplay, fontSize:"clamp(36px,5vw,64px)", letterSpacing:3, color:T.text, margin:"0 0 16px" }}>
          BUILT FOR AVIATION<br />
          <span style={{ color:T.muted }}>PROFESSIONALS</span>
        </h2>
        <p style={{ fontFamily:T.fontBody, fontSize:16, color:T.muted, maxWidth:580, lineHeight:1.8, marginBottom:64 }}>
          CH Aviation's Engine Intelligence platform transforms raw master data into actionable fleet insights — designed for lessors, MROs, analysts, and operators who demand precision.
        </p>

        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(300px,1fr))", gap:16 }}>
          {features.map((f,i) => (
            <Card key={i} style={{ borderTop:`2px solid ${i%2===0?T.blue:T.orange}`, padding:"28px 24px" }}>
              <div style={{ fontSize:28, marginBottom:12 }}>{f.icon}</div>
              <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10 }}>
                <div style={{ fontFamily:T.fontBody, fontSize:15, fontWeight:700, color:T.text }}>{f.title}</div>
                <Tag color={i%2===0?T.cyan:T.orange}>{f.tag}</Tag>
              </div>
              <p style={{ fontFamily:T.fontBody, fontSize:13, color:T.muted, lineHeight:1.7, margin:0 }}>{f.desc}</p>
            </Card>
          ))}
        </div>

        {/* Methodology strip */}
        <div style={{ marginTop:64, background:T.card, border:`1px solid ${T.cardBorder}`, borderRadius:12, padding:"32px 36px", display:"flex", gap:32, flexWrap:"wrap", alignItems:"center" }}>
          <div style={{ flex:1, minWidth:240 }}>
            <div style={{ fontFamily:T.fontDisplay, fontSize:28, color:T.blue, letterSpacing:2 }}>HOW IT WORKS</div>
            <p style={{ fontFamily:T.fontBody, fontSize:14, color:T.muted, lineHeight:1.7, margin:"12px 0 0" }}>
              Upload any CH Aviation Master Data Excel export — the platform parses it entirely in-browser, with no data sent to external servers. All analysis is instant and private.
            </p>
          </div>
          <div style={{ display:"flex", gap:0, flexWrap:"wrap" }}>
            {["01 Upload .xlsx","02 Parse & Classify","03 Visualise Fleet","04 Export Insights"].map((s,i) => (
              <div key={i} style={{ padding:"12px 24px", textAlign:"center", borderLeft: i===0?"none":`1px solid ${T.dim}` }}>
                <div style={{ fontFamily:T.fontDisplay, fontSize:22, color:T.text, letterSpacing:2 }}>{s.split(" ")[0]}</div>
                <div style={{ fontFamily:T.font, fontSize:10, color:T.muted, letterSpacing:"0.1em", marginTop:4 }}>{s.split(" ").slice(1).join(" ")}</div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </section>
  );
}

/* ─── STATIC INSIGHTS ─── */
function Insights() {
  const cfmVariants = [
    { name:"CFM56-7B26", value:1663, pct:12.8 },
    { name:"CFM56-5B4/P", value:1449, pct:11.1 },
    { name:"CFM56-7B26/3", value:1151, pct:8.9 },
    { name:"CFM56-3C1", value:938, pct:7.2 },
    { name:"CFM56-7B24", value:864, pct:6.6 },
  ];
  const leapVariants = [
    { name:"LEAP-1A26", value:2371, pct:29.9 },
    { name:"LEAP-1A32", value:1741, pct:21.9 },
    { name:"LEAP-1C28", value:789, pct:10.0 },
    { name:"LEAP-1B28", value:704, pct:8.9 },
    { name:"LEAP-1B28G05", value:636, pct:8.0 },
  ];
  const geoData = [
    { name:"Asia", CFM56:3682, LEAP:3694 },
    { name:"Europe", CFM56:2709, LEAP:1686 },
    { name:"N. America", CFM56:2662, LEAP:1457 },
    { name:"S. America", CFM56:455, LEAP:290 },
    { name:"Africa", CFM56:425, LEAP:111 },
    { name:"Oceania", CFM56:217, LEAP:108 },
  ];
  const trendData = [
    { year:"2019", CFM56:11800, LEAP:2100 },
    { year:"2020", CFM56:11200, LEAP:3000 },
    { year:"2021", CFM56:11600, LEAP:4200 },
    { year:"2022", CFM56:12100, LEAP:5500 },
    { year:"2023", CFM56:12800, LEAP:7000 },
    { year:"2024", CFM56:12993, LEAP:7929 },
  ];

  return (
    <section id="insights" style={{ background:T.bg, padding:"100px 40px" }}>
      <div style={{ maxWidth:1100, margin:"0 auto" }}>
        <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:16 }}>
          <div style={{ width:3, height:32, background:T.orange, borderRadius:2 }} />
          <div style={{ fontFamily:T.font, fontSize:11, color:T.orange, letterSpacing:"0.2em" }}>DATA INSIGHTS</div>
        </div>
        <h2 style={{ fontFamily:T.fontDisplay, fontSize:"clamp(36px,5vw,64px)", letterSpacing:3, color:T.text, margin:"0 0 16px" }}>
          GLOBAL FLEET<br />
          <span style={{ color:T.muted }}>AT A GLANCE</span>
        </h2>
        <p style={{ fontFamily:T.fontBody, fontSize:15, color:T.muted, maxWidth:560, lineHeight:1.8, marginBottom:56 }}>
          Derived from the January 2026 Master Data export — 20,922 engine records across the CFM56 and LEAP families.
        </p>

        {/* Big number cards */}
        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(180px,1fr))", gap:12, marginBottom:40 }}>
          {[
            { v:"20,922", l:"Total Engines", c:T.cyan },
            { v:"12,993", l:"CFM56 Fleet", c:T.blue },
            { v:"7,929", l:"LEAP Fleet", c:T.orange },
            { v:"13,287", l:"Active Units", c:T.green },
            { v:"3,202", l:"On Order (LEAP)", c:"#60A5FA" },
            { v:"6", l:"Regions Covered", c:"#A78BFA" },
          ].map((k,i) => (
            <Card key={i} style={{ textAlign:"center", borderTop:`2px solid ${k.c}`, padding:"24px 16px" }}>
              <div style={{ fontFamily:T.fontDisplay, fontSize:34, color:k.c, letterSpacing:1 }}>{k.v}</div>
              <div style={{ fontFamily:T.font, fontSize:10, color:T.muted, letterSpacing:"0.1em", marginTop:6 }}>{k.l}</div>
            </Card>
          ))}
        </div>

        {/* Fleet trend + geo */}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:16 }}>
          <Card>
            <ChartLabel>Fleet Growth Trend (Estimated)</ChartLabel>
            <ResponsiveContainer width="100%" height={220}>
              <AreaChart data={trendData} margin={{ left:-20, right:10 }}>
                <defs>
                  <linearGradient id="cfmGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor={T.blue} stopOpacity={0.3}/>
                    <stop offset="95%" stopColor={T.blue} stopOpacity={0}/>
                  </linearGradient>
                  <linearGradient id="leapGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor={T.orange} stopOpacity={0.3}/>
                    <stop offset="95%" stopColor={T.orange} stopOpacity={0}/>
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke={T.dim} />
                <XAxis dataKey="year" tick={{ fill:T.muted, fontSize:11, fontFamily:T.font }} />
                <YAxis tick={{ fill:T.muted, fontSize:10, fontFamily:T.font }} />
                <Tooltip content={<Tooltip_ />} />
                <Legend />
                <Area type="monotone" dataKey="CFM56" stroke={T.blue} fill="url(#cfmGrad)" strokeWidth={2} />
                <Area type="monotone" dataKey="LEAP"  stroke={T.orange} fill="url(#leapGrad)" strokeWidth={2} />
              </AreaChart>
            </ResponsiveContainer>
          </Card>

          <Card>
            <ChartLabel>Geographic Distribution</ChartLabel>
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={geoData} margin={{ left:-20, right:10 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={T.dim} />
                <XAxis dataKey="name" tick={{ fill:T.muted, fontSize:10, fontFamily:T.font }} />
                <YAxis tick={{ fill:T.muted, fontSize:10, fontFamily:T.font }} />
                <Tooltip content={<Tooltip_ />} />
                <Legend />
                <Bar dataKey="CFM56" fill={T.blue} radius={[4,4,0,0]} />
                <Bar dataKey="LEAP"  fill={T.orange} radius={[4,4,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          </Card>
        </div>

        {/* Top variants */}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:16 }}>
          {[
            { title:"Top CFM56 Variants", data:cfmVariants, color:T.blue },
            { title:"Top LEAP Variants", data:leapVariants, color:T.orange },
          ].map(({ title, data, color }) => (
            <Card key={title}>
              <ChartLabel>{title}</ChartLabel>
              {data.map(d => (
                <div key={d.name} style={{ display:"flex", alignItems:"center", gap:10, marginBottom:12 }}>
                  <div style={{ fontFamily:T.font, fontSize:11, color:T.text, width:130, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{d.name}</div>
                  <div style={{ flex:1, height:8, background:T.surface, borderRadius:4, overflow:"hidden" }}>
                    <div style={{ height:"100%", width:`${d.pct*5}%`, background:`linear-gradient(90deg, ${color}, ${color}88)`, borderRadius:4 }} />
                  </div>
                  <div style={{ fontFamily:T.font, fontSize:11, color, minWidth:44, textAlign:"right", fontWeight:700 }}>{d.value.toLocaleString()}</div>
                </div>
              ))}
            </Card>
          ))}
        </div>

        {/* CTA band */}
        <div style={{
          background:`linear-gradient(135deg, ${T.blue}18, ${T.orange}10)`,
          border:`1px solid ${T.blue}33`, borderRadius:16, padding:"48px 40px",
          display:"flex", alignItems:"center", justifyContent:"space-between", gap:32, flexWrap:"wrap"
        }}>
          <div>
            <div style={{ fontFamily:T.fontDisplay, fontSize:32, color:T.text, letterSpacing:2 }}>UPLOAD YOUR DATA</div>
            <p style={{ fontFamily:T.fontBody, fontSize:14, color:T.muted, margin:"8px 0 0", maxWidth:460, lineHeight:1.7 }}>
              These insights are derived from static sample data. Upload your own CH Aviation export to see live, personalised analysis of your fleet data.
            </p>
          </div>
          <button
            onClick={() => document.getElementById("dashboard")?.scrollIntoView({ behavior:"smooth" })}
            style={{
              fontFamily:T.font, fontSize:12, letterSpacing:"0.12em", textTransform:"uppercase",
              background:`linear-gradient(135deg, ${T.blue}, #0050C0)`,
              color:"#fff", border:"none", borderRadius:8,
              padding:"14px 36px", cursor:"pointer", whiteSpace:"nowrap",
              boxShadow:`0 0 24px ${T.blueGlow}`,
            }}
          >Open Dashboard →</button>
        </div>
      </div>
    </section>
  );
}

/* ─── DASHBOARD SECTION ─── */
function DashboardSection({ data, setData, fileName, setFileName }) {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [tab, setTab] = useState("overview");
  const [mode, setMode] = useState("classic");
  const [xlsxReady, setXlsxReady] = useState(!!window.__XLSX__);

  useEffect(() => {
    if (window.__XLSX__) { setXlsxReady(true); return; }
    const script = document.createElement("script");
    script.type = "module";
    script.textContent = `import * as XLSX from "${XLSX_SCRIPT}"; window.__XLSX__ = XLSX; window.dispatchEvent(new Event("xlsxready"));`;
    document.head.appendChild(script);
    window.addEventListener("xlsxready", () => setXlsxReady(true), { once:true });
  }, []);

  const handleFile = useCallback(async (file) => {
    if (!file || !xlsxReady) return;
    setLoading(true); setError("");
    setFileName(file.name);
    try {
      const buf = await file.arrayBuffer();
      const wb = window.__XLSX__.read(buf, { type:"array" });
      setData(processExcel(wb));
    } catch { setError("Failed to parse file. Please upload a valid CH Aviation Excel file."); }
    setLoading(false);
  }, [xlsxReady]);

  const onDrop = useCallback((e) => { e.preventDefault(); handleFile(e.dataTransfer.files[0]); }, [handleFile]);

  return (
    <section id="dashboard" style={{ background:T.surface, padding:"100px 40px 80px" }}>
      <div style={{ maxWidth:1100, margin:"0 auto" }}>
        <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:16 }}>
          <div style={{ width:3, height:32, background:T.cyan, borderRadius:2 }} />
          <div style={{ fontFamily:T.font, fontSize:11, color:T.cyan, letterSpacing:"0.2em" }}>LIVE DASHBOARD</div>
        </div>
        <h2 style={{ fontFamily:T.fontDisplay, fontSize:"clamp(36px,5vw,64px)", letterSpacing:3, color:T.text, margin:"0 0 40px" }}>
          ENGINE FLEET<br />
          <span style={{ color:T.muted }}>ANALYSER</span>
        </h2>

        {!data ? (
          /* Upload zone */
          <div
            onDrop={onDrop} onDragOver={e => e.preventDefault()}
            onClick={() => document.getElementById("xlsxInput").click()}
            style={{
              border:`2px dashed ${T.dim}`, borderRadius:16, padding:"80px 40px",
              textAlign:"center", cursor:"pointer", background:T.card,
              transition:"border-color 0.2s",
            }}
            onMouseEnter={e => e.currentTarget.style.borderColor = T.blue}
            onMouseLeave={e => e.currentTarget.style.borderColor = T.dim}
          >
            <div style={{ fontSize:56, marginBottom:20 }}>📂</div>
            <div style={{ fontFamily:T.fontDisplay, fontSize:32, letterSpacing:3, color:T.text, marginBottom:10 }}>
              DROP YOUR EXCEL FILE
            </div>
            <div style={{ fontFamily:T.font, fontSize:12, color:T.muted, marginBottom:32 }}>
              CH_Aviation_MasterData_*.xlsx · Processed entirely in-browser
            </div>
            <div style={{
              display:"inline-block", fontFamily:T.font, fontSize:11, letterSpacing:"0.12em", textTransform:"uppercase",
              background:`linear-gradient(135deg, ${T.blue}, #004ACC)`,
              color:"#fff", border:"none", borderRadius:8, padding:"12px 32px",
              boxShadow:`0 0 20px ${T.blueGlow}`,
            }}>
              {xlsxReady ? "Select File" : "Loading parser…"}
            </div>
            <input id="xlsxInput" type="file" accept=".xlsx,.xls" style={{ display:"none" }}
              onChange={e => handleFile(e.target.files[0])} />
          </div>
        ) : (
          /* Dashboard UI */
          <div>
            <div style={{ display:"flex", gap:8, marginBottom:14 }}>
              {[
                { id:"classic", label:"Classic Dashboard" },
                { id:"operator-leap", label:"Operator Search + Leap Engine" }
              ].map((m) => (
                <button key={m.id} onClick={() => setMode(m.id)} style={{
                  background: mode === m.id ? `${T.cyan}22` : "none",
                  border:`1px solid ${mode === m.id ? T.cyan : T.dim}`,
                  color: mode === m.id ? T.cyan : T.muted,
                  borderRadius:8, cursor:"pointer", fontFamily:T.font, fontSize:10,
                  letterSpacing:"0.08em", textTransform:"uppercase", padding:"8px 12px"
                }}>{m.label}</button>
              ))}
            </div>
            {/* Sub-tabs */}
            {mode === "classic" ? (
              <>
                <div style={{ display:"flex", gap:4, borderBottom:`1px solid ${T.dim}`, marginBottom:24 }}>
                  {["overview","cfm56","leap","comparison"].map(t => (
                    <button key={t} onClick={() => setTab(t)} style={{
                      background:"none", border:"none", cursor:"pointer", fontFamily:T.font,
                      fontSize:10, letterSpacing:"0.12em", textTransform:"uppercase",
                      color: tab===t ? T.text : T.muted,
                      padding:"10px 18px",
                      borderBottom:`2px solid ${tab===t ? T.cyan : "transparent"}`,
                      transition:"all 0.15s"
                    }}>
                      {t === "overview" ? "⬡ Overview" : t === "cfm56" ? "● CFM56" : t === "leap" ? "● LEAP" : "⇄ Compare"}
                    </button>
                  ))}
                  <div style={{ flex:1 }} />
                  <button onClick={() => { setData(null); setFileName(""); }} style={{
                    background:"none", border:`1px solid ${T.dim}`, borderRadius:6, cursor:"pointer",
                    fontFamily:T.font, fontSize:10, color:T.muted, padding:"6px 14px", marginBottom:4
                  }}>↩ New Upload</button>
                </div>
                <DashboardContent data={data} tab={tab} />
              </>
            ) : (
              <OperatorLeapPanel xlsxReady={xlsxReady} />
            )}
          </div>
        )}

        {loading && (
          <div style={{ textAlign:"center", marginTop:32, fontFamily:T.font, fontSize:13, color:T.cyan, display:"flex", alignItems:"center", justifyContent:"center", gap:10 }}>
            <div style={{ width:16, height:16, border:`2px solid ${T.blue}`, borderTopColor:"transparent", borderRadius:"50%", animation:"spin 0.8s linear infinite" }} />
            Parsing {fileName}…
          </div>
        )}
        {error && <div style={{ marginTop:20, fontFamily:T.font, fontSize:12, color:"#EF4444", textAlign:"center" }}>{error}</div>}
      </div>
      <style>{`@keyframes spin { to { transform:rotate(360deg); } }`}</style>
    </section>
  );
}

/* ─── DASHBOARD CONTENT ─── */
function DashboardContent({ data, tab }) {
  const { cfm56Total, leapTotal, cfm56Active, leapActive, leapPending, total } = data;

  const cfm56OperatorsFull = data.cfm56OperatorsFull || [];
  const leapOperatorsFull = data.leapOperatorsFull || [];

  const [cfmPage, setCfmPage] = useState(1);
  const [leapPage, setLeapPage] = useState(1);
  const PAGE_SIZE = 20;

  const cfmTotalPages = Math.max(1, Math.ceil(cfm56OperatorsFull.length / PAGE_SIZE));
  const leapTotalPages = Math.max(1, Math.ceil(leapOperatorsFull.length / PAGE_SIZE));

  const cfmStartIndex = (cfmPage - 1) * PAGE_SIZE;
  const leapStartIndex = (leapPage - 1) * PAGE_SIZE;

  const cfmPageItems = cfm56OperatorsFull.slice(cfmStartIndex, cfmStartIndex + PAGE_SIZE);
  const leapPageItems = leapOperatorsFull.slice(leapStartIndex, leapStartIndex + PAGE_SIZE);

  const KPI = ({ label, value, sub, color }) => (
    <Card style={{ borderTop:`2px solid ${color}`, padding:"18px 20px", flex:1, minWidth:130 }}>
      <div style={{ fontFamily:"'Bebas Neue',sans-serif", fontSize:32, color, letterSpacing:1 }}>
        {typeof value === "number" ? value.toLocaleString() : value}
      </div>
      <div style={{ fontFamily:T.font, fontSize:10, color:T.text, letterSpacing:"0.08em", marginTop:2 }}>{label}</div>
      {sub && <div style={{ fontFamily:T.font, fontSize:9, color:T.muted, marginTop:3 }}>{sub}</div>}
    </Card>
  );

  if (tab === "overview") return (
    <div>
      <div style={{ display:"flex", gap:12, flexWrap:"wrap", marginBottom:20 }}>
        <KPI label="TOTAL ENGINES" value={total} color={T.cyan} />
        <KPI label="CFM56 FLEET" value={cfm56Total} color={T.blue} sub={`${((cfm56Total/total)*100).toFixed(1)}% of total`} />
        <KPI label="LEAP FLEET" value={leapTotal} color={T.orange} sub={`${((leapTotal/total)*100).toFixed(1)}% of total`} />
        <KPI label="CFM56 ACTIVE" value={cfm56Active} color={T.green} sub={`${((cfm56Active/cfm56Total)*100).toFixed(1)}% rate`} />
        <KPI label="LEAP ACTIVE" value={leapActive} color="#34D399" sub={`${((leapActive/leapTotal)*100).toFixed(1)}% rate`} />
        <KPI label="ON ORDER" value={leapPending} color="#60A5FA" sub="Not yet delivered" />
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:14 }}>
        <Card>
          <ChartLabel>Fleet Composition</ChartLabel>
          <ResponsiveContainer width="100%" height={230}>
            <PieChart>
              <Pie data={[{name:"CFM56",value:cfm56Total}, {name:"LEAP",value:leapTotal}]}
                cx="50%" cy="50%" innerRadius={60} outerRadius={95} paddingAngle={4} dataKey="value">
                <Cell fill={T.blue}/><Cell fill={T.orange}/>
              </Pie>
              <Tooltip content={<Tooltip_ />}/><Legend/>
            </PieChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <ChartLabel>Geographic Distribution</ChartLabel>
          <ResponsiveContainer width="100%" height={230}>
            <BarChart data={data.cfm56Continents.map(c=>{
              const l=data.leapContinents.find(x=>x.name===c.name);
              return { name:c.name.replace("North America","N.Am").replace("South America","S.Am"), CFM56:c.value, LEAP:l?l.value:0 };
            })} margin={{left:-20,right:5}}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.dim}/>
              <XAxis dataKey="name" tick={{fill:T.muted,fontSize:10,fontFamily:T.font}}/>
              <YAxis tick={{fill:T.muted,fontSize:10,fontFamily:T.font}}/>
              <Tooltip content={<Tooltip_/>}/><Legend/>
              <Bar dataKey="CFM56" fill={T.blue} radius={[3,3,0,0]}/>
              <Bar dataKey="LEAP" fill={T.orange} radius={[3,3,0,0]}/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>

      {/* Status bars */}
      <Card>
        <ChartLabel>Fleet Status Overview</ChartLabel>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:24 }}>
          {[{label:"CFM56",arr:data.cfm56Status,tot:cfm56Total,color:T.blue},{label:"LEAP",arr:data.leapStatus,tot:leapTotal,color:T.orange}].map(({label,arr,tot,color})=>(
            <div key={label}>
              <div style={{fontFamily:T.font,fontSize:10,color,letterSpacing:"0.1em",marginBottom:10}}>— {label} STATUS</div>
              {arr.map(s=>(
                <div key={s.name} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                  <div style={{width:8,height:8,borderRadius:2,background:STATUS_COLORS[s.name]||T.muted,flexShrink:0}}/>
                  <div style={{fontSize:11,fontFamily:T.font,color:T.muted,width:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.name}</div>
                  <div style={{flex:1,height:6,background:T.surface,borderRadius:3}}>
                    <div style={{height:"100%",width:`${(s.value/tot*100).toFixed(1)}%`,background:STATUS_COLORS[s.name]||T.muted,borderRadius:3}}/>
                  </div>
                  <div style={{fontFamily:T.font,fontSize:11,color:T.text,fontWeight:700,minWidth:48,textAlign:"right"}}>{s.value.toLocaleString()}</div>
                </div>
              ))}
            </div>
          ))}
        </div>
      </Card>
    </div>
  );

  if (tab === "cfm56") return (
    <div>
      <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:20}}>
        {[
          {l:"TOTAL CFM56",v:cfm56Total,c:T.blue},
          {l:"ACTIVE",v:cfm56Active,c:T.green,s:`${((cfm56Active/cfm56Total)*100).toFixed(1)}%`},
          {l:"STORED",v:data.cfm56Status.find(s=>s.name==="Stored")?.value||0,c:"#FFD600"},
          {l:"SCRAPPED",v:data.cfm56Status.find(s=>s.name==="Scrapped")?.value||0,c:T.muted},
          {l:"MAINTENANCE",v:data.cfm56Status.find(s=>s.name==="Maintenance")?.value||0,c:"#FF9800"},
        ].map((k,i)=><KPI key={i} label={k.l} value={k.v} color={k.c} sub={k.s}/>)}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
        <Card>
          <ChartLabel>Sub-family Breakdown</ChartLabel>
          <ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie data={data.cfm56Families} cx="50%" cy="50%" outerRadius={90} dataKey="value"
                label={({name,value})=>`${name}: ${value.toLocaleString()}`} labelLine={false}>
                {data.cfm56Families.map((_,i)=><Cell key={i} fill={[T.blue,"#3399FF","#66B2FF","#0044AA","#99CCFF"][i%5]}/>)}
              </Pie>
              <Tooltip content={<Tooltip_/>}/>
            </PieChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <ChartLabel>Top Aircraft Types</ChartLabel>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={data.cfm56Aircraft} layout="vertical" margin={{left:8,right:8}}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.dim} horizontal={false}/>
              <XAxis type="number" tick={{fill:T.muted,fontSize:10,fontFamily:T.font}}/>
              <YAxis type="category" dataKey="name" tick={{fill:T.muted,fontSize:10,fontFamily:T.font}} width={88}/>
              <Tooltip content={<Tooltip_/>}/>
              <Bar dataKey="value" fill={T.blue} radius={[0,4,4,0]} name="Engines"/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <Card>
          <ChartLabel>Top Operators</ChartLabel>
          {data.cfm56Operators.map((op,i)=>(
            <div key={op.name} style={{display:"flex",alignItems:"center",gap:8,marginBottom:9}}>
              <div style={{fontFamily:T.font,fontSize:10,color:T.muted,width:16,textAlign:"right"}}>{i+1}</div>
              <div style={{flex:1,fontFamily:T.font,fontSize:11,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{op.name}</div>
              <div style={{width:80,height:5,background:T.surface,borderRadius:3}}>
                <div style={{height:"100%",width:`${(op.value/data.cfm56Operators[0].value)*100}%`,background:T.blue,borderRadius:3}}/>
              </div>
              <div style={{fontFamily:T.font,fontSize:10,color:T.blue,fontWeight:700,minWidth:36,textAlign:"right"}}>{op.value}</div>
            </div>
          ))}
        </Card>
        <Card>
          <ChartLabel>Top Variants</ChartLabel>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={data.cfm56Variants} margin={{left:-15,right:5,bottom:40}}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.dim}/>
              <XAxis dataKey="name" tick={{fill:T.muted,fontSize:9,fontFamily:T.font}} angle={-40} textAnchor="end" interval={0}/>
              <YAxis tick={{fill:T.muted,fontSize:10,fontFamily:T.font}}/>
              <Tooltip content={<Tooltip_/>}/>
              <Bar dataKey="value" radius={[4,4,0,0]} name="Count">
                {data.cfm56Variants.map((_,i)=><Cell key={i} fill={`hsl(${210+i*5},80%,${48+i*2}%)`}/>)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>
      <Card style={{ marginTop:14 }}>
        <ChartLabel>All Operators — CFM56</ChartLabel>
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:8 }}>
          <div style={{ fontFamily:T.font, fontSize:10, color:T.muted }}>
            {cfm56OperatorsFull.length.toLocaleString()} operators total
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <button
              onClick={() => setCfmPage(p => Math.max(1, p - 1))}
              disabled={cfmPage === 1}
              style={{
                fontFamily:T.font,
                fontSize:10,
                padding:"4px 10px",
                borderRadius:4,
                border:`1px solid ${T.dim}`,
                background: cfmPage === 1 ? T.surface : T.card,
                color: cfmPage === 1 ? T.dim : T.text,
                cursor: cfmPage === 1 ? "default" : "pointer"
              }}
            >
              ‹ Prev
            </button>
            <div style={{ fontFamily:T.font, fontSize:10, color:T.muted }}>
              Page {cfmPage} of {cfmTotalPages}
            </div>
            <button
              onClick={() => setCfmPage(p => Math.min(cfmTotalPages, p + 1))}
              disabled={cfmPage === cfmTotalPages}
              style={{
                fontFamily:T.font,
                fontSize:10,
                padding:"4px 10px",
                borderRadius:4,
                border:`1px solid ${T.dim}`,
                background: cfmPage === cfmTotalPages ? T.surface : T.card,
                color: cfmPage === cfmTotalPages ? T.dim : T.text,
                cursor: cfmPage === cfmTotalPages ? "default" : "pointer"
              }}
            >
              Next ›
            </button>
          </div>
        </div>
        <div style={{ maxHeight:260, overflowY:"auto", borderTop:`1px solid ${T.cardBorder}`, marginTop:4 }}>
          {cfmPageItems.map((op, i) => (
            <div
              key={op.name || i}
              style={{
                display:"flex",
                alignItems:"center",
                justifyContent:"space-between",
                padding:"6px 4px",
                borderBottom:`1px solid ${T.cardBorder}`
              }}
            >
              <div style={{ display:"flex", alignItems:"center", gap:8, minWidth:0 }}>
                <div style={{ fontFamily:T.font, fontSize:10, color:T.muted, width:18, textAlign:"right" }}>
                  {cfmStartIndex + i + 1}
                </div>
                <div
                  style={{
                    fontFamily:T.font,
                    fontSize:11,
                    color:T.text,
                    overflow:"hidden",
                    textOverflow:"ellipsis",
                    whiteSpace:"nowrap"
                  }}
                  title={op.name}
                >
                  {op.name || "Unknown"}
                </div>
              </div>
              <div style={{ fontFamily:T.font, fontSize:11, color:T.blue, fontWeight:700, minWidth:60, textAlign:"right" }}>
                {op.value.toLocaleString()}
              </div>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );

  if (tab === "leap") return (
    <div>
      <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:20}}>
        {[
          {l:"TOTAL LEAP",v:leapTotal,c:T.orange},
          {l:"ACTIVE",v:leapActive,c:T.green,s:`${((leapActive/leapTotal)*100).toFixed(1)}%`},
          {l:"ON ORDER",v:leapPending,c:"#60A5FA",s:"Not yet delivered"},
          {l:"MAINTENANCE",v:data.leapStatus.find(s=>s.name==="Maintenance")?.value||0,c:"#FF9800"},
          {l:"STORED",v:data.leapStatus.find(s=>s.name==="Stored")?.value||0,c:"#FFD600"},
        ].map((k,i)=><KPI key={i} label={k.l} value={k.v} color={k.c} sub={k.s}/>)}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
        <Card>
          <ChartLabel>Family Breakdown</ChartLabel>
          <ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie data={data.leapFamilies} cx="50%" cy="50%" outerRadius={90} dataKey="value"
                label={({name,value})=>`${name.split(" ")[0]}: ${value.toLocaleString()}`} labelLine={false}>
                {data.leapFamilies.map((_,i)=><Cell key={i} fill={[T.orange,T.cyan,"#FF9944"][i%3]}/>)}
              </Pie>
              <Tooltip content={<Tooltip_/>}/><Legend/>
            </PieChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <ChartLabel>Top Aircraft Types</ChartLabel>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={data.leapAircraft} layout="vertical" margin={{left:8,right:8}}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.dim} horizontal={false}/>
              <XAxis type="number" tick={{fill:T.muted,fontSize:10,fontFamily:T.font}}/>
              <YAxis type="category" dataKey="name" tick={{fill:T.muted,fontSize:10,fontFamily:T.font}} width={88}/>
              <Tooltip content={<Tooltip_/>}/>
              <Bar dataKey="value" fill={T.orange} radius={[0,4,4,0]} name="Engines"/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <Card>
          <ChartLabel>Top Operators</ChartLabel>
          {data.leapOperators.filter(o=>o.name!=="Unassigned").slice(0,8).map((op,i)=>(
            <div key={op.name} style={{display:"flex",alignItems:"center",gap:8,marginBottom:9}}>
              <div style={{fontFamily:T.font,fontSize:10,color:T.muted,width:16,textAlign:"right"}}>{i+1}</div>
              <div style={{flex:1,fontFamily:T.font,fontSize:11,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{op.name}</div>
              <div style={{width:80,height:5,background:T.surface,borderRadius:3}}>
                <div style={{height:"100%",width:`${(op.value/data.leapOperators[0].value)*100}%`,background:T.orange,borderRadius:3}}/>
              </div>
              <div style={{fontFamily:T.font,fontSize:10,color:T.orange,fontWeight:700,minWidth:36,textAlign:"right"}}>{op.value}</div>
            </div>
          ))}
        </Card>
        <Card>
          <ChartLabel>Top Variants</ChartLabel>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={data.leapVariants} margin={{left:-15,right:5,bottom:40}}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.dim}/>
              <XAxis dataKey="name" tick={{fill:T.muted,fontSize:9,fontFamily:T.font}} angle={-40} textAnchor="end" interval={0}/>
              <YAxis tick={{fill:T.muted,fontSize:10,fontFamily:T.font}}/>
              <Tooltip content={<Tooltip_/>}/>
              <Bar dataKey="value" radius={[4,4,0,0]} name="Count">
                {data.leapVariants.map((_,i)=><Cell key={i} fill={`hsl(${25+i*8},85%,${48+i*1.5}%)`}/>)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>
      <Card style={{ marginTop:14 }}>
        <ChartLabel>All Operators — LEAP</ChartLabel>
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:8 }}>
          <div style={{ fontFamily:T.font, fontSize:10, color:T.muted }}>
            {leapOperatorsFull.length.toLocaleString()} operators total
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <button
              onClick={() => setLeapPage(p => Math.max(1, p - 1))}
              disabled={leapPage === 1}
              style={{
                fontFamily:T.font,
                fontSize:10,
                padding:"4px 10px",
                borderRadius:4,
                border:`1px solid ${T.dim}`,
                background: leapPage === 1 ? T.surface : T.card,
                color: leapPage === 1 ? T.dim : T.text,
                cursor: leapPage === 1 ? "default" : "pointer"
              }}
            >
              ‹ Prev
            </button>
            <div style={{ fontFamily:T.font, fontSize:10, color:T.muted }}>
              Page {leapPage} of {leapTotalPages}
            </div>
            <button
              onClick={() => setLeapPage(p => Math.min(leapTotalPages, p + 1))}
              disabled={leapPage === leapTotalPages}
              style={{
                fontFamily:T.font,
                fontSize:10,
                padding:"4px 10px",
                borderRadius:4,
                border:`1px solid ${T.dim}`,
                background: leapPage === leapTotalPages ? T.surface : T.card,
                color: leapPage === leapTotalPages ? T.dim : T.text,
                cursor: leapPage === leapTotalPages ? "default" : "pointer"
              }}
            >
              Next ›
            </button>
          </div>
        </div>
        <div style={{ maxHeight:260, overflowY:"auto", borderTop:`1px solid ${T.cardBorder}`, marginTop:4 }}>
          {leapPageItems.map((op, i) => (
            <div
              key={op.name || i}
              style={{
                display:"flex",
                alignItems:"center",
                justifyContent:"space-between",
                padding:"6px 4px",
                borderBottom:`1px solid ${T.cardBorder}`
              }}
            >
              <div style={{ display:"flex", alignItems:"center", gap:8, minWidth:0 }}>
                <div style={{ fontFamily:T.font, fontSize:10, color:T.muted, width:18, textAlign:"right" }}>
                  {leapStartIndex + i + 1}
                </div>
                <div
                  style={{
                    fontFamily:T.font,
                    fontSize:11,
                    color:T.text,
                    overflow:"hidden",
                    textOverflow:"ellipsis",
                    whiteSpace:"nowrap"
                  }}
                  title={op.name}
                >
                  {op.name || "Unknown"}
                </div>
              </div>
              <div style={{ fontFamily:T.font, fontSize:11, color:T.orange, fontWeight:700, minWidth:60, textAlign:"right" }}>
                {op.value.toLocaleString()}
              </div>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );

  // Comparison tab
  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
        {[
          {label:"Total Fleet",cfm:cfm56Total,leap:leapTotal},
          {label:"Active Units",cfm:cfm56Active,leap:leapActive},
          {label:"Active Rate",cfm:`${((cfm56Active/cfm56Total)*100).toFixed(1)}%`,leap:`${((leapActive/leapTotal)*100).toFixed(1)}%`},
          {label:"Unique Variants",cfm:data.cfm56Variants.length+"+",leap:data.leapVariants.length+"+"},
        ].map(row=>(
          <Card key={row.label} style={{padding:"16px 20px"}}>
            <div style={{fontFamily:T.font,fontSize:9,color:T.muted,marginBottom:12,letterSpacing:"0.12em",textTransform:"uppercase"}}>{row.label}</div>
            <div style={{display:"flex",alignItems:"center",gap:12}}>
              <div style={{flex:1,textAlign:"center"}}>
                <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:28,color:T.blue}}>{typeof row.cfm==="number"?row.cfm.toLocaleString():row.cfm}</div>
                <div style={{fontFamily:T.font,fontSize:9,color:T.muted}}>CFM56</div>
              </div>
              <div style={{color:T.dim,fontSize:16}}>vs</div>
              <div style={{flex:1,textAlign:"center"}}>
                <div style={{fontFamily:"'7Bebas Neue',sans-serif",fontSize:28,color:T.orange}}>{typeof row.leap==="number"?row.leap.toLocaleString():row.leap}</div>
                <div style={{fontFamily:T.font,fontSize:9,color:T.muted}}>LEAP</div>
              </div>
            </div>
          </Card>
        ))}
      </div>
      <Card style={{marginBottom:14}}>
        <ChartLabel>Regional Deployment — CFM56 vs LEAP</ChartLabel>
        <ResponsiveContainer width="100%" height={240}>
          <BarChart data={["Asia","Europe","North America","South America","Africa","Oceania"].map(cont=>{
            const c=data.cfm56Continents.find(x=>x.name===cont);
            const l=data.leapContinents.find(x=>x.name===cont);
            return {name:cont.replace("North America","N.Am").replace("South America","S.Am"),CFM56:c?.value||0,LEAP:l?.value||0};
          })} margin={{left:-15,right:10}}>
            <CartesianGrid strokeDasharray="3 3" stroke={T.dim}/>
            <XAxis dataKey="name" tick={{fill:T.muted,fontSize:11,fontFamily:T.font}}/>
            <YAxis tick={{fill:T.muted,fontSize:10,fontFamily:T.font}}/>
            <Tooltip content={<Tooltip_/>}/><Legend/>
            <Bar dataKey="CFM56" fill={T.blue} radius={[4,4,0,0]}/>
            <Bar dataKey="LEAP" fill={T.orange} radius={[4,4,0,0]}/>
          </BarChart>
        </ResponsiveContainer>
      </Card>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <Card>
          <ChartLabel>Top Operators — CFM56</ChartLabel>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={data.cfm56Operators.slice(0,6)} layout="vertical" margin={{left:8}}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.dim} horizontal={false}/>
              <XAxis type="number" tick={{fill:T.muted,fontSize:10,fontFamily:T.font}}/>
              <YAxis type="category" dataKey="name" tick={{fill:T.muted,fontSize:10,fontFamily:T.font}} width={110}/>
              <Tooltip content={<Tooltip_/>}/>
              <Bar dataKey="value" fill={T.blue} radius={[0,4,4,0]} name="CFM56"/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <ChartLabel>Top Operators — LEAP</ChartLabel>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={data.leapOperators.filter(o=>o.name!=="Unassigned").slice(0,6)} layout="vertical" margin={{left:8}}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.dim} horizontal={false}/>
              <XAxis type="number" tick={{fill:T.muted,fontSize:10,fontFamily:T.font}}/>
              <YAxis type="category" dataKey="name" tick={{fill:T.muted,fontSize:10,fontFamily:T.font}} width={110}/>
              <Tooltip content={<Tooltip_/>}/>
              <Bar dataKey="value" fill={T.orange} radius={[0,4,4,0]} name="LEAP"/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>
    </div>
  );
}

function OpSectionTitle({ title, subtitle }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <h2 style={{ margin: 0, fontSize: 20 }}>{title}</h2>
      <p style={{ margin: "4px 0 0", color: "#94a3b8", fontSize: 13 }}>{subtitle}</p>
    </div>
  );
}

function OperatorLeapPanel({ xlsxReady }) {
  const folderRef = useRef(null);
  const [rows, setRows] = useState([]);
  const [trendRows, setTrendRows] = useState([]);
  const [activeTab, setActiveTab] = useState("operator-search");
  const [singleFileName, setSingleFileName] = useState("");
  const [folderInfo, setFolderInfo] = useState({ count: 0, months: [] });
  const [loadingSingle, setLoadingSingle] = useState(false);
  const [loadingFolder, setLoadingFolder] = useState(false);
  const [error, setError] = useState("");
  const [openFilterKey, setOpenFilterKey] = useState(null);
  const [filters, setFilters] = useState({ family: [], continent: [], status: [], dash: [], country: [], operator: [] });
  const [filterSearch, setFilterSearch] = useState({ family: "", continent: "", status: "", dash: "", country: "", operator: "" });

  const includesFilter = (values, value) => !values.length || values.includes(value);
  const baseRows = rows.filter((r) => includesFilter(filters.family, r.family) && includesFilter(filters.continent, r.continent) && includesFilter(filters.status, r.status) && includesFilter(filters.dash, r.dash));
  const countryOptions = Array.from(new Set(baseRows.map((r) => r.country))).sort();
  const operatorOptions = Array.from(new Set(baseRows.filter((r) => includesFilter(filters.country, r.country)).map((r) => r.operator))).sort();
  const filtered = baseRows.filter((r) => includesFilter(filters.country, r.country) && includesFilter(filters.operator, r.operator));
  const options = (() => { const get = (k) => Array.from(new Set(rows.map((r) => r[k]))).sort(); return { family: get("family"), continent: get("continent"), status: get("status"), dash: get("dash") }; })();

  const grouped = (() => {
    const map = new Map();
    filtered.forEach((r) => {
      const key = `${r.operator}||${r.country}||${r.continent}`;
      if (!map.has(key)) map.set(key, { key, operator: r.operator, country: r.country, continent: r.continent, family: {}, dash: {}, status: {}, engines: 0 });
      const g = map.get(key);
      g.family[r.family] = (g.family[r.family] || 0) + r.engines;
      g.dash[r.dash] = (g.dash[r.dash] || 0) + r.engines;
      g.status[r.status] = (g.status[r.status] || 0) + r.engines;
      g.engines += r.engines;
    });
    return Array.from(map.values()).sort((a, b) => b.engines - a.engines);
  })();

  const kpi = {
    total: filtered.reduce((s, r) => s + r.engines, 0),
    active: filtered.filter((r) => r.status === "Active").reduce((s, r) => s + r.engines, 0),
    activeCFM: filtered.filter((r) => r.status === "Active" && r.engineType === "CFM").reduce((s, r) => s + r.engines, 0),
    activeLEAP: filtered.filter((r) => r.status === "Active" && r.engineType === "LEAP").reduce((s, r) => s + r.engines, 0),
  };
  const operatorRanking = grouped.slice(0, 12).map((o) => ({ name: o.operator, engines: o.engines }));
  const byContinent = Object.entries(filtered.reduce((m, r) => {
    m[r.continent] = (m[r.continent] || 0) + r.engines;
    return m;
  }, {})).map(([name, engines]) => ({ name, engines })).sort((a, b) => b.engines - a.engines);
  const byCountry = Object.entries(filtered.reduce((m, r) => {
    m[r.country] = (m[r.country] || 0) + r.engines;
    return m;
  }, {})).map(([name, engines]) => ({ name, engines })).sort((a, b) => b.engines - a.engines);
  const byFamily = Object.entries(filtered.reduce((m, r) => {
    m[r.family] = (m[r.family] || 0) + r.engines;
    return m;
  }, {})).map(([name, value]) => ({ name, value }));
  const byDash = Object.entries(filtered.reduce((m, r) => {
    m[r.dash] = (m[r.dash] || 0) + r.engines;
    return m;
  }, {})).map(([dash, engines]) => ({ dash, engines }));
  const byType = Object.entries(filtered.reduce((m, r) => {
    m[r.engineType] = (m[r.engineType] || 0) + r.engines;
    return m;
  }, {})).map(([name, value]) => ({ name, value }));
  const byStatus = Object.entries(filtered.reduce((m, r) => {
    m[r.status] = (m[r.status] || 0) + r.engines;
    return m;
  }, {})).map(([status, value]) => ({ status, value }));
  const opportunity = Object.entries(filtered.reduce((m, r) => {
    if (["Stored", "Maintenance", "Repair"].includes(r.status)) {
      m[r.operator] = (m[r.operator] || 0) + r.engines;
    }
    return m;
  }, {})).map(([operator, value]) => ({ operator, value })).sort((a, b) => b.value - a.value).slice(0, 12);
  const trendByMonth = (() => { const m = {}; trendRows.forEach((r) => { if (!m[r.month]) m[r.month] = { month: r.month, LEAP: 0, CFM: 0 }; m[r.month][r.engineType] += r.engines; }); return Object.values(m).sort((a, b) => opMonthSort(a.month) - opMonthSort(b.month)); })();
  const leapGrowth = (() => { if (trendByMonth.length < 2) return 0; const p = trendByMonth[trendByMonth.length - 2].LEAP; const c = trendByMonth[trendByMonth.length - 1].LEAP; return p ? ((c - p) / p) * 100 : c ? 100 : 0; })();
  const leapRegion = (() => {
    const months = Array.from(new Set(trendRows.map((r) => r.month))).sort((a, b) => opMonthSort(a) - opMonthSort(b));
    const continents = Array.from(new Set(trendRows.map((r) => r.continent)));
    return months.map((m) => {
      const row = { month: m };
      continents.forEach((c) => {
        row[c] = trendRows.filter((r) => r.month === m && r.continent === c && r.engineType === "LEAP").reduce((s, r) => s + r.engines, 0);
      });
      return row;
    });
  })();

  async function parseFiles(fileList) {
    if (!xlsxReady || !window.__XLSX__) throw new Error("Excel parser is still loading.");
    const files = Array.from(fileList || []).filter((f) => /\.(xlsx|xls)$/i.test(f.name));
    if (!files.length) throw new Error("No Excel files found.");
    const out = [];
    for (const file of files) {
      const data = await file.arrayBuffer();
      out.push(...opParseWorkbook(file, window.__XLSX__.read(data, { type: "array" })));
    }
    return { files, rows: out };
  }

  async function onSingle(e) {
    setError("");
    setLoadingSingle(true);
    try {
      const file = e.target.files?.[0];
      if (!file) return;
      const result = await parseFiles([file]);
      setRows(result.rows);
      setSingleFileName(file.name);
      setFilters({ family: [], continent: [], status: [], dash: [], country: [], operator: [] });
      setFilterSearch({ family: "", continent: "", status: "", dash: "", country: "", operator: "" });
    } catch (err) {
      setError(err.message || "Failed to parse file.");
    } finally {
      setLoadingSingle(false);
    }
  }

  async function onFolder(e) {
    setError("");
    setLoadingFolder(true);
    try {
      const result = await parseFiles(e.target.files);
      const months = Array.from(new Set(result.rows.map((r) => r.month))).sort((a, b) => opMonthSort(a) - opMonthSort(b));
      setTrendRows(result.rows);
      setFolderInfo({ count: result.files.length, months });
      setActiveTab("leap-engine");
    } catch (err) {
      setError(err.message || "Failed to parse folder files.");
    } finally {
      setLoadingFolder(false);
    }
  }

  return (
    <div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 16 }}>
        {[["operator-search", "Operator Search"], ["leap-engine", "Leap Engine"]].map(([id, label]) => (
          <button key={id} onClick={() => setActiveTab(id)} style={{ border: "1px solid rgba(148,163,184,0.25)", background: activeTab === id ? "rgba(56,189,248,0.2)" : "rgba(15,23,42,0.55)", color: "#e2e8f0", borderRadius: 999, padding: "8px 14px", cursor: "pointer" }}>{label}</button>
        ))}
      </div>

      <div style={{ display: "grid", gap: 14, gridTemplateColumns: "repeat(auto-fit,minmax(280px,1fr))", marginBottom: 16 }}>
        <Card>
          <h3>Upload Data (Operator Search)</h3>
          <input type="file" accept=".xlsx,.xls" onChange={onSingle} />
          <div style={{ marginTop: 8, color: "#94a3b8" }}>
            {loadingSingle ? "Parsing..." : singleFileName ? `Loaded: ${singleFileName}` : "No file loaded"}
          </div>
        </Card>
        <Card>
          <h3>LEAP Demand Analysis (Folder)</h3>
          <input
            ref={folderRef}
            type="file"
            multiple
            onClick={() => {
              if (folderRef.current) {
                folderRef.current.setAttribute("webkitdirectory", "");
                folderRef.current.setAttribute("directory", "");
              }
            }}
            onChange={onFolder}
            accept=".xlsx,.xls"
          />
          <div style={{ marginTop: 8, color: "#94a3b8" }}>
            {loadingFolder ? "Reading files..." : folderInfo.count ? `${folderInfo.count} files | Months: ${folderInfo.months.join(", ")}` : "No folder analyzed"}
          </div>
        </Card>
      </div>

      {error ? <Card style={{ borderColor: "rgba(248,113,113,0.45)", color: "#fda4af", marginBottom: 14 }}>{error}</Card> : null}

      {activeTab === "operator-search" ? (
        rows.length ? (
          <div style={{ display: "grid", gap: 16 }}>
            <Card style={{ display: "grid", gap: 10, gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))" }}>
              {[
                ["Engine Family", "family", options.family],
                ["Operator Continent", "continent", options.continent],
                ["Status", "status", options.status],
                ["Dash", "dash", options.dash],
                ["Operator Country", "country", countryOptions],
                ["Operator Name", "operator", operatorOptions],
              ].map(([label, key, list]) => (
                <div key={key}>
                  <div
                    onClick={() => setOpenFilterKey((prev) => prev === key ? null : key)}
                    style={{ width: "100%", borderRadius: 8, border: "1px solid rgba(148,163,184,0.3)", background: "rgba(15,23,42,0.7)", color: "#e2e8f0", minHeight: 36, padding: "8px 10px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" }}
                  >
                    <div>
                      <div style={{ fontSize: 11, color: "#94a3b8" }}>{label}</div>
                      <div style={{ fontSize: 12 }}>{filters[key].length ? `${filters[key].length} selected` : "All selected"}</div>
                    </div>
                    <div style={{ fontSize: 12, color: "#94a3b8" }}>{openFilterKey === key ? "▲" : "▼"}</div>
                  </div>
                  {openFilterKey === key ? (
                    <div style={{ marginTop: 6, border: "1px solid rgba(148,163,184,0.2)", borderRadius: 8, padding: 8, background: "rgba(2,6,23,0.35)" }}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
                        <div style={{ fontSize: 12, color: "#94a3b8" }}>{label}</div>
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            setFilters((prev) => ({ ...prev, [key]: [] }));
                          }}
                          style={{ background: "none", border: "none", color: "#60a5fa", cursor: "pointer", fontSize: 11, padding: 0 }}
                        >
                          Clear
                        </button>
                      </div>
                      <input
                        type="text"
                        value={filterSearch[key]}
                        onChange={(e) => setFilterSearch((prev) => ({ ...prev, [key]: e.target.value }))}
                        placeholder={`Search ${label.toLowerCase()}`}
                        style={{ width: "100%", borderRadius: 8, border: "1px solid rgba(148,163,184,0.3)", background: "rgba(15,23,42,0.7)", color: "#e2e8f0", height: 32, padding: "0 8px", marginBottom: 6 }}
                      />
                      <div style={{ maxHeight: 120, overflowY: "auto", paddingRight: 2 }}>
                        {list.filter((opt) => String(opt).toLowerCase().includes(filterSearch[key].toLowerCase())).map((opt) => (
                          <label key={opt} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: "#e2e8f0", marginBottom: 4 }}>
                            <input
                              type="checkbox"
                              checked={filters[key].includes(opt)}
                              onChange={() => setFilters((prev) => {
                                const selected = prev[key].includes(opt) ? prev[key].filter((x) => x !== opt) : [...prev[key], opt];
                                const next = { ...prev, [key]: selected };
                                if (["family", "continent", "status", "dash"].includes(key)) {
                                  next.country = [];
                                  next.operator = [];
                                }
                                if (key === "country") next.operator = [];
                                return next;
                              })}
                            />
                            <span>{opt}</span>
                          </label>
                        ))}
                        {!list.filter((opt) => String(opt).toLowerCase().includes(filterSearch[key].toLowerCase())).length ? (
                          <div style={{ color: "#94a3b8", fontSize: 12 }}>No matches</div>
                        ) : null}
                      </div>
                    </div>
                  ) : null}
                </div>
              ))}
            </Card>

            <div style={{ display: "grid", gap: 12, gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))" }}>
              {[["Total Engines", kpi.total, "#38bdf8"], ["Active Engines", kpi.active, "#34d399"], ["Active CFM", kpi.activeCFM, "#60a5fa"], ["Active LEAP", kpi.activeLEAP, "#f97316"]].map(([l, v, c]) => (
                <Card key={l} style={{ borderColor: `${c}55` }}>
                  <div style={{ color: "#94a3b8" }}>{l}</div>
                  <div style={{ fontSize: 30, fontWeight: 700, color: c }}>{opFmt(v)}</div>
                </Card>
              ))}
            </div>

            <Card style={{ overflowX: "auto" }}>
              <OpSectionTitle title="Operator Engine Distribution" subtitle="Grouped rows with expandable status breakdown" />
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ color: "#94a3b8", textAlign: "left" }}>
                    <th style={{ padding: 8 }}>Country</th><th style={{ padding: 8 }}>Operator</th><th style={{ padding: 8 }}>Continent</th><th style={{ padding: 8 }}>Family</th><th style={{ padding: 8 }}>Dash</th><th style={{ padding: 8 }}>Status</th><th style={{ padding: 8 }}>Engines</th>
                  </tr>
                </thead>
                <tbody>
                  {grouped.slice(0, 40).map((r) => (
                    <tr key={r.key} style={{ borderTop: "1px solid rgba(148,163,184,0.16)" }}>
                      <td style={{ padding: 8 }}>{r.country}</td><td style={{ padding: 8 }}>{r.operator}</td><td style={{ padding: 8 }}>{r.continent}</td><td style={{ padding: 8 }}>{Object.keys(r.family).slice(0,2).join(", ") || "Mixed"}</td><td style={{ padding: 8 }}>{Object.keys(r.dash).slice(0,3).join(", ") || "Mixed"}</td><td style={{ padding: 8 }}>{Object.entries(r.status).map(([k,v]) => `${k}: ${opFmt(v)}`).join(" | ")}</td><td style={{ padding: 8 }}>{opFmt(r.engines)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </Card>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(360px,1fr))", gap: 16 }}>
              <Card style={{ minHeight: 340 }}>
                <OpSectionTitle title="Operator Ranking" subtitle="Top operators by engine count" />
                <ResponsiveContainer width="100%" height={270}><BarChart data={operatorRanking}><CartesianGrid strokeDasharray="3 3" stroke="rgba(148,163,184,0.2)" /><XAxis dataKey="name" hide /><YAxis /><Tooltip /><Bar dataKey="engines" fill="#38bdf8" radius={[6,6,0,0]} /></BarChart></ResponsiveContainer>
              </Card>
              <Card style={{ minHeight: 340 }}>
                <OpSectionTitle title="Status Breakdown" subtitle="Lifecycle mix" />
                <ResponsiveContainer width="100%" height={270}><PieChart><Pie data={byStatus} dataKey="value" nameKey="status" outerRadius={95} label>{byStatus.map((x) => <Cell key={x.status} fill={OP_STATUS_COLORS[x.status] || OP_STATUS_COLORS.Unknown} />)}</Pie><Tooltip /></PieChart></ResponsiveContainer>
              </Card>
            </div>

            <div style={{ display: "grid", gap: 16, gridTemplateColumns: "repeat(auto-fit,minmax(320px,1fr))" }}>
              <Card style={{ minHeight: 320 }}><OpSectionTitle title="Continent Breakdown" subtitle="Engines per continent" /><ResponsiveContainer width="100%" height={240}><BarChart data={byContinent}><CartesianGrid strokeDasharray="3 3" stroke="rgba(148,163,184,0.2)" /><XAxis dataKey="name" /><YAxis /><Tooltip /><Bar dataKey="engines" fill="#34d399" /></BarChart></ResponsiveContainer></Card>
              <Card style={{ minHeight: 320 }}><OpSectionTitle title="Country Drilldown" subtitle="Top countries" /><ResponsiveContainer width="100%" height={240}><BarChart data={byCountry.slice(0,12)}><CartesianGrid strokeDasharray="3 3" stroke="rgba(148,163,184,0.2)" /><XAxis dataKey="name" hide /><YAxis /><Tooltip /><Bar dataKey="engines" fill="#a78bfa" /></BarChart></ResponsiveContainer></Card>
              <Card style={{ minHeight: 330 }}><OpSectionTitle title="Family Distribution" subtitle="CFM56 / LEAP families" /><ResponsiveContainer width="100%" height={250}><PieChart><Pie data={byFamily} dataKey="value" nameKey="name" outerRadius={90} label>{byFamily.map((x, i) => <Cell key={x.name} fill={OP_PIE_COLORS[i % OP_PIE_COLORS.length]} />)}</Pie><Tooltip /></PieChart></ResponsiveContainer></Card>
              <Card style={{ minHeight: 330 }}><OpSectionTitle title="Dash Analysis" subtitle="Most used dash" /><ResponsiveContainer width="100%" height={250}><BarChart data={byDash}><CartesianGrid strokeDasharray="3 3" stroke="rgba(148,163,184,0.2)" /><XAxis dataKey="dash" /><YAxis /><Tooltip /><Bar dataKey="engines" fill="#f97316" /></BarChart></ResponsiveContainer></Card>
              <Card style={{ minHeight: 330 }}><OpSectionTitle title="Engine Type Split" subtitle="CFM vs LEAP" /><ResponsiveContainer width="100%" height={250}><PieChart><Pie data={byType} dataKey="value" nameKey="name" outerRadius={90} label>{byType.map((x) => <Cell key={x.name} fill={OP_ENGINE_TYPE_COLORS[x.name] || "#94a3b8"} />)}</Pie><Tooltip /></PieChart></ResponsiveContainer></Card>
              <Card style={{ minHeight: 330 }}><OpSectionTitle title="Opportunity Score" subtitle="Stored + Maintenance + Repair" /><ResponsiveContainer width="100%" height={250}><BarChart data={opportunity}><CartesianGrid strokeDasharray="3 3" stroke="rgba(148,163,184,0.2)" /><XAxis dataKey="operator" hide /><YAxis /><Tooltip /><Bar dataKey="value" fill="#f59e0b" /></BarChart></ResponsiveContainer></Card>
            </div>
          </div>
        ) : (
          <Card style={{ color: "#94a3b8" }}>Upload a single file in Operator Search to view the dashboard.</Card>
        )
      ) : (
        trendRows.length ? (
          <div style={{ display: "grid", gap: 16 }}>
            <div style={{ display: "grid", gap: 12, gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))" }}>
              <Card style={{ borderColor: "rgba(249,115,22,0.45)" }}><div style={{ color: "#94a3b8" }}>LEAP MoM Growth</div><div style={{ fontSize: 32, color: leapGrowth >= 0 ? "#34d399" : "#f87171", fontWeight: 700 }}>{leapGrowth >= 0 ? "+" : ""}{leapGrowth.toFixed(1)}%</div></Card>
              <Card><div style={{ color: "#94a3b8" }}>Months</div><div style={{ fontSize: 24, fontWeight: 700 }}>{folderInfo.months.join(", ")}</div></Card>
              <Card><div style={{ color: "#94a3b8" }}>Files Processed</div><div style={{ fontSize: 32, fontWeight: 700 }}>{folderInfo.count}</div></Card>
            </div>
            <div style={{ display: "grid", gap: 16, gridTemplateColumns: "repeat(auto-fit,minmax(360px,1fr))" }}>
              <Card style={{ minHeight: 330 }}><OpSectionTitle title="Monthly LEAP Trend" subtitle="LEAP by month" /><ResponsiveContainer width="100%" height={250}><AreaChart data={trendByMonth}><CartesianGrid strokeDasharray="3 3" stroke="rgba(148,163,184,0.2)" /><XAxis dataKey="month" /><YAxis /><Tooltip /><Area type="monotone" dataKey="LEAP" stroke="#f97316" fill="#f9731633" strokeWidth={3} /></AreaChart></ResponsiveContainer></Card>
              <Card style={{ minHeight: 330 }}><OpSectionTitle title="CFM vs LEAP Growth" subtitle="Two-line trend" /><ResponsiveContainer width="100%" height={250}><BarChart data={trendByMonth}><CartesianGrid strokeDasharray="3 3" stroke="rgba(148,163,184,0.2)" /><XAxis dataKey="month" /><YAxis /><Tooltip /><Legend /><Bar dataKey="CFM" fill="#38bdf8" /><Bar dataKey="LEAP" fill="#f97316" /></BarChart></ResponsiveContainer></Card>
            </div>
            <Card style={{ minHeight: 350 }}>
              <OpSectionTitle title="Region-wise LEAP Growth" subtitle="Continent adoption" />
              <ResponsiveContainer width="100%" height={270}>
                <BarChart data={leapRegion}><CartesianGrid strokeDasharray="3 3" stroke="rgba(148,163,184,0.2)" /><XAxis dataKey="month" /><YAxis /><Tooltip /><Legend />{Object.keys(leapRegion[0] || {}).filter((k) => k !== "month").map((c, i) => <Bar key={c} dataKey={c} fill={OP_PIE_COLORS[i % OP_PIE_COLORS.length]} />)}</BarChart>
              </ResponsiveContainer>
            </Card>
          </div>
        ) : (
          <Card style={{ color: "#94a3b8" }}>To analyze LEAP engine trends, select a folder with monthly Excel files.</Card>
        )
      )}
    </div>
  );
}

/* ─── FOOTER ─── */
function Footer() {
  return (
    <footer style={{ background:"#02050C", borderTop:`1px solid ${T.cardBorder}`, padding:"60px 40px 40px" }}>
      <div style={{ maxWidth:1100, margin:"0 auto" }}>
        <div style={{ display:"flex", gap:40, flexWrap:"wrap", marginBottom:48 }}>
          <div style={{ flex:2, minWidth:240 }}>
            <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:16 }}>
              <div style={{ width:32, height:32, borderRadius:7, background:`linear-gradient(135deg,${T.blue},${T.orange})`, display:"flex", alignItems:"center", justifyContent:"center" }}>✈</div>
              <div style={{ fontFamily:T.fontDisplay, fontSize:18, letterSpacing:2, color:T.text }}>CH AVIATION</div>
            </div>
            <p style={{ fontFamily:T.fontBody, fontSize:13, color:T.muted, lineHeight:1.8, maxWidth:300 }}>
              Aviation engine intelligence platform — turning raw master data into fleet insight for the global aviation industry.
            </p>
          </div>
          {[
            { title:"PLATFORM", links:["Dashboard","Data Insights","Upload Guide","API Access"] },
            { title:"ENGINE FAMILIES", links:["CFM56 Series","LEAP-1A","LEAP-1B","LEAP-1C"] },
            { title:"COMPANY", links:["About CH Aviation","Methodology","Data Sources","Contact"] },
          ].map(col => (
            <div key={col.title} style={{ minWidth:140 }}>
              <div style={{ fontFamily:T.font, fontSize:10, color:T.muted, letterSpacing:"0.15em", marginBottom:16 }}>{col.title}</div>
              {col.links.map(l => (
                <div key={l} style={{ fontFamily:T.fontBody, fontSize:13, color:T.dim, marginBottom:10, cursor:"pointer" }}>{l}</div>
              ))}
            </div>
          ))}
        </div>
        <div style={{ borderTop:`1px solid ${T.cardBorder}`, paddingTop:24, display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:12 }}>
          <div style={{ fontFamily:T.font, fontSize:10, color:T.muted }}>© 2026 CH AVIATION ENGINE INTELLIGENCE · ALL RIGHTS RESERVED</div>
          <div style={{ display:"flex", gap:12 }}>
            <Tag color={T.blue}>CFM56</Tag>
            <Tag color={T.orange}>LEAP</Tag>
            <Tag color={T.cyan}>MASTER DATA</Tag>
          </div>
        </div>
      </div>
    </footer>
  );
}

/* ─── ROOT APP ─── */
export default function App() {
  const [activeSection, setActiveSection] = useState("hero");
  const [data, setData] = useState(null);
  const [fileName, setFileName] = useState("");

  const onNav = (section) => {
    setActiveSection(section);
    document.getElementById(section)?.scrollIntoView({ behavior:"smooth" });
  };

  // Track scroll position for nav highlight
  useEffect(() => {
    const handler = () => {
      const sections = ["hero","about","insights","dashboard"];
      for (let i = sections.length-1; i >= 0; i--) {
        const el = document.getElementById(sections[i]);
        if (el && window.scrollY >= el.offsetTop - 120) {
          setActiveSection(sections[i]);
          break;
        }
      }
    };
    window.addEventListener("scroll", handler);
    return () => window.removeEventListener("scroll", handler);
  }, []);

  return (
    <div style={{ fontFamily:T.fontBody, background:T.bg, color:T.text, minHeight:"100vh" }}>
      {/* Google Fonts */}
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=IBM+Plex+Mono:wght@400;700&family=IBM+Plex+Sans:wght@400;600;700&display=swap');
        * { box-sizing:border-box; }
        html { scroll-behavior:smooth; }
        ::-webkit-scrollbar { width:6px; }
        ::-webkit-scrollbar-track { background:#04080F; }
        ::-webkit-scrollbar-thumb { background:#132036; border-radius:3px; }
        body { margin:0; }
      `}</style>

      <Nav activeSection={activeSection} onNav={onNav} fileName={fileName} onReset={() => { setData(null); setFileName(""); }} />
      <Hero onNav={onNav} />
      <About />
      <Insights />
      <DashboardSection data={data} setData={setData} fileName={fileName} setFileName={setFileName} />
      <Footer />
    </div>
  );
}
