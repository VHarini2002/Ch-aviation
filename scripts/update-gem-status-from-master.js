/**
 * Updates Gem Status.xlsx region/country/GEM fields from CH Aviation MasterData MASTER INFO.
 * Preserves Gem Status row order and structure (no merge into master workbook).
 */
const XLSX = require("xlsx");
const path = require("path");

const ROOT = path.join(__dirname, "..");
const GEM_PATHS = [
  path.join(ROOT, "Gem Status.xlsx"),
  path.join(ROOT, "public", "Gem Status.xlsx"),
];
const MASTER_PATH = path.join(ROOT, "CH Aviation MasterData APR 13 2026 JR.xlsx");

const norm = (s) =>
  String(s ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ");
const normRegion = (s) => String(s ?? "").trim().toLowerCase();

function comboKey(op, country, region) {
  return `${norm(op)}|||${norm(country)}|||${normRegion(region)}`;
}

function pickConsensus(set) {
  const arr = [...set].filter((x) => x !== "" && x != null);
  if (!arr.length) return "";
  const counts = {};
  for (const g of arr) {
    const key = String(g).trim();
    counts[key] = (counts[key] || 0) + 1;
  }
  return Object.entries(counts).sort((a, b) => b[1] - a[1])[0][0];
}

function pickTop(obj) {
  const entries = Object.entries(obj).filter(([k]) => k && k !== "–");
  if (!entries.length) return "";
  return entries.sort((a, b) => b[1] - a[1])[0][0];
}

function buildMasterLookups(rows) {
  const byCombo = new Map();
  const byOpCountry = new Map();
  const byOp = new Map();

  for (const r of rows) {
    const op = String(r["Operator Name"] ?? "").trim();
    if (!op || op === "–") continue;

    const region = String(r["GEM REGION"] ?? "").trim();
    const country = String(r["Operator Country"] ?? "").trim();
    const gem = String(r["GEM STATUS"] ?? "").trim();
    const comment = String(r["COMMENTS"] ?? "").trim();

    const add = (map, key) => {
      if (!map.has(key)) {
        map.set(key, {
          regionCounts: {},
          countryCounts: {},
          gems: new Set(),
          comments: new Set(),
        });
      }
      const bucket = map.get(key);
      if (region) bucket.regionCounts[region] = (bucket.regionCounts[region] || 0) + 1;
      if (country && country !== "–") {
        bucket.countryCounts[country] = (bucket.countryCounts[country] || 0) + 1;
      }
      if (gem !== "") bucket.gems.add(gem);
      if (comment) bucket.comments.add(comment);
    };

    add(byCombo, comboKey(op, country, region));
    add(byOpCountry, `${norm(op)}|||${norm(country)}`);
    add(byOp, norm(op));
  }

  const resolve = (bucket) => ({
    region: pickTop(bucket.regionCounts),
    country: pickTop(bucket.countryCounts),
    gemStatus: pickConsensus(bucket.gems),
    comments: [...bucket.comments].join(" | "),
  });

  return {
    resolveCombo: (op, country, region) => {
      const exact = byCombo.get(comboKey(op, country, region));
      if (exact) return resolve(exact);
      const oc = byOpCountry.get(`${norm(op)}|||${norm(country)}`);
      if (oc) return resolve(oc);
      const o = byOp.get(norm(op));
      if (o) return resolve(o);
      return null;
    },
  };
}

function updateGemFile(gemPath, lookups) {
  const wb = XLSX.readFile(gemPath);
  const sheetName = wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" });

  let updated = 0;
  const out = rows.map((row) => {
    const op = String(row["Operator Name"] ?? "").trim();
    if (!op || op === "–") return row;

    const country = String(row["Operator Country"] ?? "").trim();
    const region = String(row["GEM REGION"] ?? "").trim();
    const resolved = lookups.resolveCombo(op, country, region);
    if (!resolved) return row;

    const next = {
      "GEM REGION": resolved.region || region,
      "Operator Country": resolved.country || country,
      "Operator Name": op,
      "GEM STATUS": resolved.gemStatus,
      COMMENTS: resolved.comments,
    };

    const changed =
      next["GEM REGION"] !== region ||
      next["Operator Country"] !== country ||
      String(row["GEM STATUS"] ?? "").trim() !== next["GEM STATUS"] ||
      String(row["COMMENTS"] ?? "").trim() !== next["COMMENTS"];

    if (changed) updated++;
    return next;
  });

  const newSheet = XLSX.utils.json_to_sheet(out, {
    header: ["GEM REGION", "Operator Country", "Operator Name", "GEM STATUS", "COMMENTS"],
  });
  wb.Sheets[sheetName] = newSheet;
  XLSX.writeFile(wb, gemPath);
  return { rows: out.length, updated };
}

function main() {
  const masterWb = XLSX.readFile(MASTER_PATH);
  const masterSheet = masterWb.SheetNames.includes("MASTER INFO")
    ? "MASTER INFO"
    : masterWb.SheetNames[0];
  const masterRows = XLSX.utils.sheet_to_json(masterWb.Sheets[masterSheet], {
    defval: "",
  });
  const lookups = buildMasterLookups(masterRows);

  for (const gemPath of GEM_PATHS) {
    const result = updateGemFile(gemPath, lookups);
    console.log(`Updated ${gemPath}: ${result.updated}/${result.rows} rows changed`);
  }
}

main();
