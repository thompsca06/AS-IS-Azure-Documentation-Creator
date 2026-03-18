#!/usr/bin/env node
/**
 * Generate-StandaloneDiagrams.js
 *
 * Generates Management Group Hierarchy and Network Topology diagrams
 * as standalone PNG files from the JSON output of Invoke-AzureTenancyAudit.ps1.
 *
 * Usage:
 *   node Generate-StandaloneDiagrams.js <path-to-json> [output-dir] [--width=1200] [--format=png|svg]
 *
 * Examples:
 *   node Generate-StandaloneDiagrams.js data.json
 *   node Generate-StandaloneDiagrams.js data.json ./diagrams --width=1600
 *   node Generate-StandaloneDiagrams.js data.json ./diagrams --format=svg
 *
 * Output:
 *   <output-dir>/MG_Hierarchy_Diagram.png
 *   <output-dir>/Network_Topology_Diagram.png
 *   (or .svg if --format=svg)
 */

const fs = require("fs");
const path = require("path");
const { generateMgHierarchyDiagram, generateNetworkDiagram } = require("./Generate-Diagrams");

// ============================================================
// PARSE ARGS
// ============================================================
const args = process.argv.slice(2);
const flags = {};
const positional = [];

for (const arg of args) {
  if (arg.startsWith("--")) {
    const [key, val] = arg.slice(2).split("=");
    flags[key] = val || "true";
  } else {
    positional.push(arg);
  }
}

const jsonPath = positional[0];
if (!jsonPath) {
  console.error("Usage: node Generate-StandaloneDiagrams.js <json-path> [output-dir] [--width=1200] [--format=png|svg]");
  console.error("");
  console.error("Options:");
  console.error("  --width=N       Diagram width in pixels (default: 1200)");
  console.error("  --format=FORMAT Output format: png or svg (default: png)");
  console.error("  --mg-only       Generate only the Management Group diagram");
  console.error("  --net-only      Generate only the Network Topology diagram");
  process.exit(1);
}

if (!fs.existsSync(jsonPath)) {
  console.error(`File not found: ${jsonPath}`);
  process.exit(1);
}

const outputDir = positional[1] || path.dirname(jsonPath);
const diagramWidth = parseInt(flags.width || "1200", 10);
const outputFormat = (flags.format || "png").toLowerCase();
const mgOnly = flags["mg-only"] === "true";
const netOnly = flags["net-only"] === "true";

// ============================================================
// BRAND COLORS (same as main document generator)
// ============================================================
const BRAND_DEFAULTS = {
  primary: "2DA58E", accent: "B5E5E2", dark: "1A6B5A", light: "E0F5F3",
  text: "333333", grey: "666666", white: "FFFFFF",
};

let colors = { ...BRAND_DEFAULTS };
try {
  const cfgPath = path.join(path.dirname(process.argv[1] || __filename), "branding-config.json");
  if (fs.existsSync(cfgPath)) {
    const raw = JSON.parse(fs.readFileSync(cfgPath, "utf8"));
    if (raw.colors) colors = { ...colors, ...raw.colors };
    console.log("Branding: loaded from branding-config.json");
  }
} catch (err) {
  // Use defaults
}

// ============================================================
// MAIN
// ============================================================
(async () => {
  console.log("========================================");
  console.log("  Standalone Diagram Generator");
  console.log("========================================");
  console.log(`Input:   ${jsonPath}`);
  console.log(`Output:  ${outputDir}`);
  console.log(`Width:   ${diagramWidth}px`);
  console.log(`Format:  ${outputFormat}`);
  console.log("");

  const raw = fs.readFileSync(jsonPath, "utf8").replace(/^\uFEFF/, "");
  const data = JSON.parse(raw);
  const S = data.Sections || {};
  const customerName = data.CustomerName || "Azure";

  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  let generated = 0;

  // Helper to write diagram output (handles SVG fallback when sharp unavailable)
  function writeDiagram(result, baseName) {
    const isSvgOnly = result.isSvg || outputFormat === "svg";
    const ext = isSvgOnly ? "svg" : "png";
    const filename = `${customerName}_${baseName}.${ext}`;
    const outPath = path.join(outputDir, filename);

    if (isSvgOnly && result.svgString) {
      fs.writeFileSync(outPath, result.svgString, "utf8");
    } else {
      fs.writeFileSync(outPath, result.pngBuffer);
    }

    const note = result.isSvg && outputFormat !== "svg" ? " (SVG fallback - sharp not available)" : "";
    console.log(`  [OK] ${filename} (${result.widthPx}x${result.heightPx}px)${note}`);
    return outPath;
  }

  // ---- Management Group Hierarchy Diagram ----
  if (!netOnly) {
    console.log("Generating Management Group Hierarchy diagram...");
    try {
      const result = await generateMgHierarchyDiagram(
        S.ManagementGroups, S.Subscriptions, colors, diagramWidth
      );
      if (result) {
        writeDiagram(result, "MG_Hierarchy_Diagram");
        generated++;
      } else {
        console.log("  [SKIP] No management group data available.");
      }
    } catch (err) {
      console.error(`  [ERROR] MG diagram: ${err.message}`);
    }
  }

  // ---- Network Topology Diagram ----
  if (!mgOnly) {
    console.log("Generating Network Topology diagram...");
    try {
      const result = await generateNetworkDiagram(
        S.Networking, colors, diagramWidth
      );
      if (result) {
        writeDiagram(result, "Network_Topology_Diagram");
        generated++;
      } else {
        console.log("  [SKIP] No networking data available.");
      }
    } catch (err) {
      console.error(`  [ERROR] Network diagram: ${err.message}`);
    }
  }

  console.log("");
  console.log(`Generated ${generated} diagram(s) in: ${outputDir}`);
  if (generated === 0) {
    console.log("No diagrams generated. Check that the JSON contains ManagementGroups and/or Networking data.");
  }
})().catch(err => {
  console.error("Fatal error:", err);
  process.exit(1);
});
