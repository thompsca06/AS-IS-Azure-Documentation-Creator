#!/usr/bin/env node
/**
 * Generate-BuildDocument.js
 * 
 * Reads the JSON output from Invoke-AzureTenancyAudit.ps1
 * and generates a fully populated AS-IS Build Document (.docx)
 * 
 * Usage:
 *   node Generate-BuildDocument.js <path-to-json> [output-path]
 *   
 * Example:
 *   node Generate-BuildDocument.js .\AzureBuildDoc_Output\20260226\Sense_Azure_ASIS_Data.json
 *   node Generate-BuildDocument.js data.json "C:\Reports\Sense_ASIS_Build.docx"
 * 
 * Prerequisites:
 *   npm install -g docx
 */

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, PageOrientation,
  TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, TabStopType, TabStopPosition
} = require("docx");
const { generateMgHierarchyDiagram, generateNetworkDiagram } = require("./Generate-Diagrams");

// ============================================================
// ARGS
// ============================================================
const jsonPath = process.argv[2];
if (!jsonPath) {
  console.error("Usage: node Generate-BuildDocument.js <json-path> [output-docx-path]");
  console.error("  json-path:   Path to the JSON file from Invoke-AzureTenancyAudit.ps1");
  console.error("  output-path: (Optional) Output .docx path. Defaults to same dir as JSON.");
  process.exit(1);
}

if (!fs.existsSync(jsonPath)) {
  console.error(`File not found: ${jsonPath}`);
  process.exit(1);
}

(async () => {

const raw = fs.readFileSync(jsonPath, "utf8").replace(/^\uFEFF/, ""); // strip BOM
const data = JSON.parse(raw);
const S = data.Sections || {};

const customerName = data.CustomerName || "[Customer Name]";
const outputDocx = process.argv[3] || path.join(
  path.dirname(jsonPath),
  `${customerName}_Azure_ASIS_Build_Document.docx`
);

const entraIncluded = data.EntraIncluded || false;
const entraStatus = data.EntraStatus || (entraIncluded ? "Unknown" : "Not in scope");

console.log(`Customer:  ${customerName}`);
console.log(`Entra ID:  ${entraStatus}`);
console.log(`Input:     ${jsonPath}`);
console.log(`Output:    ${outputDocx}`);


// ============================================================
// BRAND & CONSTANTS (loaded from branding-config.json if available)
// ============================================================
const BRAND_DEFAULTS = {
  colors: {
    primary: "2DA58E", accent: "B5E5E2", dark: "1A6B5A", light: "E0F5F3",
    lighter: "F2F2F2", lighterAlt: "EAEAEA", text: "333333", grey: "666666",
    white: "FFFFFF", border: "FFFFFF", red: "C00000", amber: "ED7D31",
    green: "548235", yellow: "FFF2CC", redBg: "FCE4EC", amberBg: "FFF3E0", greenBg: "E8F5E9",
  },
  page: { width: 11906, height: 16838, margin: 1134 },
  fonts: { heading: "Segoe UI Semibold", body: "Segoe UI" },
  company: { name: "Tieva Ltd", website: "tieva.co.uk", classification: "Confidential" },
};

let brandCfg = BRAND_DEFAULTS;
try {
  const cfgPath = path.join(path.dirname(process.argv[1] || __filename), "branding-config.json");
  if (fs.existsSync(cfgPath)) {
    const raw = JSON.parse(fs.readFileSync(cfgPath, "utf8"));
    brandCfg = {
      colors: { ...BRAND_DEFAULTS.colors, ...(raw.colors || {}) },
      page: { ...BRAND_DEFAULTS.page, ...(raw.page || {}) },
      fonts: { ...BRAND_DEFAULTS.fonts, ...(raw.fonts || {}) },
      company: { ...BRAND_DEFAULTS.company, ...(raw.company || {}) },
    };
    console.log("Branding: loaded from branding-config.json");
  }
} catch (err) {
  console.warn("Warning: Could not load branding-config.json, using defaults:", err.message);
}

const B = brandCfg.colors;
const BRAND_FONTS = brandCfg.fonts;
const BRAND_COMPANY = brandCfg.company;

const PAGE_W = brandCfg.page.width;
const PAGE_H = brandCfg.page.height;
const MARGIN = brandCfg.page.margin;
const CW = PAGE_W - 2 * MARGIN; // content width ~9638
const LANDSCAPE_CW_PX = 940; // landscape content width in px (~10 in at 96 DPI)

const bdr = { style: BorderStyle.SINGLE, size: 4, color: B.border };
const borders = { top: bdr, bottom: bdr, left: bdr, right: bdr };
const cm = { top: 50, bottom: 50, left: 80, right: 80 };

// ============================================================
// HELPERS
// ============================================================
function s(v) { return v === null || v === undefined ? "" : String(v); }

function hCell(text, width) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: B.accent, type: ShadingType.CLEAR },
    margins: cm,
    children: [new Paragraph({ children: [new TextRun({ text: s(text), bold: true, color: "000000", font: "Segoe UI Semibold", size: 18 })] })],
  });
}

function dCell(text, width, shade) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: shade ? B.lighterAlt : B.lighter, type: ShadingType.CLEAR },
    margins: cm,
    children: [new Paragraph({ children: [new TextRun({ text: s(text), font: "Segoe UI", size: 18, color: B.text })] })],
  });
}

function statusCell(text, width) {
  let fill = B.white;
  const t = s(text).toUpperCase();
  if (t.includes("NOT CONFIGURED") || t.includes("NOT ENABLED") || t.includes("CRITICAL") || t.includes("NO POLICIES")) fill = B.redBg;
  else if (t.includes("NOT ASSESSED") || t.includes("NEEDS IMPROVEMENT") || t.includes("INCONSISTENT") || t.includes("PARTIALLY") || t.includes("HIGH")) fill = B.amberBg;
  else if (t.includes("ENABLED") || t.includes("ADEQUATE") || t.includes("CONSISTENT") || t.includes("CONFIGURED")) fill = B.greenBg;

  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: cm,
    children: [new Paragraph({ children: [new TextRun({ text: s(text), font: "Segoe UI", size: 18, color: B.text, bold: true })] })],
  });
}

function makeTable(headers, rows, colWidths) {
  const tw = colWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: tw, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({ children: headers.map((h, i) => hCell(h, colWidths[i])) }),
      ...rows.map((row, ri) =>
        new TableRow({ children: row.map((c, ci) => dCell(c, colWidths[ci], ri % 2 === 1)) })
      ),
    ],
  });
}

function makeStatusTable(headers, rows, colWidths, statusColIndex) {
  const tw = colWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: tw, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({ children: headers.map((h, i) => hCell(h, colWidths[i])) }),
      ...rows.map((row, ri) =>
        new TableRow({
          children: row.map((c, ci) =>
            ci === statusColIndex ? statusCell(c, colWidths[ci]) :
            dCell(c, colWidths[ci], ri % 2 === 1)
          ),
        })
      ),
    ],
  });
}

function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text, font: "Segoe UI Semibold" })] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text, font: "Segoe UI Semibold" })] }); }
function h3(text) { return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text, font: "Segoe UI Semibold" })] }); }
function p(text) { return new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: s(text), font: "Segoe UI", size: 20, color: B.text })] }); }
function pBold(text) { return new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: s(text), font: "Segoe UI", size: 20, color: B.text, bold: true })] }); }
function note(text) { return new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: s(text), font: "Segoe UI", size: 18, color: B.grey, italics: true })] }); }
function gap() { return new Paragraph({ spacing: { after: 160 }, children: [] }); }
function pb() { return new Paragraph({ children: [new PageBreak()] }); }

function countOrNA(arr) { return Array.isArray(arr) ? arr.length : 0; }

// ============================================================
// BUILD CHILDREN ARRAY
// ============================================================
// ============================================================
// PRE-GENERATE DIAGRAMS (async — requires sharp)
// ============================================================
const DIAGRAM_WIDTH_PX = LANDSCAPE_CW_PX; // diagrams render on landscape pages
let mgDiagramResult = null;
let networkDiagramResult = null;
let mgDiagramContent = null;   // saved separately for landscape section
let netDiagramContent = null;
let mgDiagramSplitIdx = -1;    // children index where landscape section starts
let netDiagramSplitIdx = -1;

try {
  mgDiagramResult = await generateMgHierarchyDiagram(S.ManagementGroups, S.Subscriptions, B, DIAGRAM_WIDTH_PX);
  if (mgDiagramResult) console.log(`MG diagram:  ${mgDiagramResult.widthPx}x${mgDiagramResult.heightPx}px`);
} catch (err) {
  console.warn("Warning: MG hierarchy diagram generation failed:", err.message);
}

try {
  networkDiagramResult = await generateNetworkDiagram(S.Networking, B, DIAGRAM_WIDTH_PX);
  if (networkDiagramResult) console.log(`Net diagram: ${networkDiagramResult.widthPx}x${networkDiagramResult.heightPx}px`);
} catch (err) {
  console.warn("Warning: Network diagram generation failed:", err.message);
}

// ============================================================
// BUILD CHILDREN ARRAY
// ============================================================
const children = [];

// ---- COVER PAGE ----
children.push(
  // Teal top border bar
  new Paragraph({ spacing: { before: 0, after: 0 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: B.primary, space: 1 } }, children: [] }),
  new Paragraph({ spacing: { before: 2400 }, children: [] }),
  // Customer name
  new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 200 },
    children: [new TextRun({ text: customerName, font: "Segoe UI Semibold", size: 64, bold: true, color: B.primary })] }),
  // Subtitle line 1
  new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 100 },
    children: [new TextRun({ text: "Azure Tenancy", font: "Segoe UI", size: 44, color: B.dark })] }),
  // Subtitle line 2
  new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 100 },
    children: [new TextRun({ text: "AS-IS Build Document", font: "Segoe UI", size: 36, color: B.dark })] }),
  // Light turquoise separator
  new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 600 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: B.accent, space: 1 } }, children: [] }),
  // Metadata
  new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 80 },
    children: [new TextRun({ text: `Prepared by: Tieva Ltd`, font: "Segoe UI", size: 22, color: B.grey })] }),
  new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 80 },
    children: [new TextRun({ text: `Date: ${data.CollectionDate || new Date().toLocaleDateString("en-GB")}`, font: "Segoe UI", size: 22, color: B.grey })] }),
  new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 80 },
    children: [new TextRun({ text: `Tenant ID: ${data.TenantId || "N/A"}`, font: "Segoe UI", size: 20, color: B.grey })] }),
  new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 80 },
    children: [new TextRun({ text: "Classification: Confidential", font: "Segoe UI", size: 20, color: B.grey })] }),
  // Teal bottom border bar
  new Paragraph({ spacing: { before: 1200, after: 0 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: B.primary, space: 1 } }, children: [] }),
  pb()
);

// ---- DOCUMENT CONTROL ----
children.push(
  h1("Document Control"),
  gap(),
  h2("Document Information"),
  makeTable(
    ["Field", "Details"],
    [
      ["Document Title", `${customerName} - Azure AS-IS Build Document`],
      ["Creation Date", data.CollectionDate || "[DD/MM/YYYY]"],
      ["Owner / Author", data.CollectedBy || "[Name]"],
      ["Customer Name", customerName],
      ["Tenant ID", data.TenantId || "[GUID]"],
      ["Entra ID Scope", entraIncluded ? `Included (${entraStatus})` : "Not in scope"],
      ["Internal / External Audience", "Internal / External"],
      ["Document Classification", "Confidential"],
      ["Data Source", "Automated collection via Invoke-AzureTenancyAudit.ps1"],
    ],
    [3200, CW - 3200]
  ),
  gap(),
  h2("TIEVA Approval History"),
  makeTable(
    ["Version", "Approver", "Position", "Date"],
    [
      ["1.0", "[Name]", "Technical Consultant", "[DD/MM/YYYY]"],
      ["", "", "", ""],
      ["", "", "", ""],
    ],
    [1200, 2800, 2800, 2838]
  ),
  gap(),
  h2("Revision History"),
  makeTable(
    ["Version", "Revision", "Revised By", "Date"],
    [
      ["1.0", "Initial automated collection", data.CollectedBy || "[Name]", data.CollectionDate || "[Date]"],
      ["", "", "", ""],
    ],
    [1200, 3500, 2700, 2238]
  ),
  pb()
);

// ---- TOC ----
children.push(
  h1("Contents"),
  new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
  pb()
);

// ============================================================
// 1. EXECUTIVE SUMMARY
// ============================================================
children.push(
  h1("1. Executive Summary"),
  p(`This document provides an AS-IS record of the Azure tenancy configuration for ${customerName}. All data was collected programmatically on ${data.CollectionDate || "N/A"} using the Invoke-AzureTenancyAudit.ps1 script running as ${data.CollectedBy || "N/A"}.`),
  gap(),
  pBold("Environment at a Glance:"),
  makeTable(
    ["Metric", "Value"],
    [
      ["Tenant ID", data.TenantId || "N/A"],
      ["Entra ID (Identity)", entraIncluded ? (entraStatus === "Collected" ? "Collected" : "Attempted - " + entraStatus) : "Not in scope"],
      ["Active Subscriptions", String(countOrNA(S.Subscriptions))],
      ["Management Groups", String(countOrNA(S.ManagementGroups))],
      ["Virtual Networks", String(countOrNA(S.Networking?.VirtualNetworks))],
      ["Subnets", String(countOrNA(S.Networking?.Subnets))],
      ["Virtual Machines", String(countOrNA(S.Compute?.VirtualMachines))],
      ["Storage Accounts", String(countOrNA(S.Storage))],
      ["Recovery Services Vaults", String(countOrNA(S.Backup?.RecoveryVaults))],
      ["Policy Assignments", String(countOrNA(S.Policy?.PolicyAssignments))],
      ["RBAC Assignments", String(countOrNA(S.RBAC))],
      ["Firewalls", String(countOrNA(S.Networking?.Firewalls))],
      ["Application Gateways", String(countOrNA(S.Networking?.ApplicationGateways))],
      ["Load Balancers", String(countOrNA(S.Networking?.LoadBalancers))],
      ["Private Endpoints", String(countOrNA(S.Networking?.PrivateEndpoints))],
      ["Managed Disks", String(countOrNA(S.Compute?.ManagedDisks))],
      ["VM Scale Sets", String(countOrNA(S.Compute?.VMScaleSets))],
      ["App Service Plans", String(countOrNA(S.AppServices?.AppServicePlans))],
      ["Web Apps", String(countOrNA(S.AppServices?.WebApps))],
      ["Function Apps", String(countOrNA(S.AppServices?.FunctionApps))],
      ["SQL Servers", String(countOrNA(S.Databases?.SQLServers))],
      ["Cosmos DB Accounts", String(countOrNA(S.Databases?.CosmosDB))],
      ["AKS Clusters", String(countOrNA(S.Containers?.AKSClusters))],
      ["Container Registries", String(countOrNA(S.Containers?.ContainerRegistries))],
      ["Automation Accounts", String(countOrNA(S.Automation?.AutomationAccounts))],
      ["Resource Types (Total)", String(countOrNA(S.ResourceSummary))],
    ],
    [4500, CW - 4500]
  ),
  pb()
);

// ============================================================
// 2. ENTRA ID & IDENTITY
// ============================================================
children.push(h1("2. Entra ID (Azure AD) & Identity"));

if (S.Identity) {
  const ti = S.Identity.TenantInfo;
  if (ti) {
    children.push(
      h2("2.1. Tenant Overview"),
      makeTable(
        ["Property", "Value"],
        [
          ["Tenant Name", ti.TenantName],
          ["Tenant ID", ti.TenantId],
          ["Primary Domain", ti.PrimaryDomain],
          ["Verified Domains", ti.VerifiedDomains],
          ["Created", ti.CreatedDateTime],
        ],
        [3200, CW - 3200]
      ),
      gap()
    );
  }

  // Conditional Access
  const ca = S.Identity.ConditionalAccess;
  if (ca && ca.length > 0) {
    children.push(
      h2("2.2. Conditional Access Policies"),
      p(`${ca.length} Conditional Access policies found. ${ca.filter(x => x.State === "enabled").length} enabled, ${ca.filter(x => x.State === "enabledForReportingButNotEnforced").length} report-only, ${ca.filter(x => x.State === "disabled").length} disabled.`),
      gap()
    );

    const caRows = ca.map(c => [
      c.DisplayName, c.State,
      [c.IncludeUsers, c.IncludeGroups].filter(Boolean).join("; ") || "N/A",
      c.IncludeApps || "N/A",
      c.GrantControls || "N/A"
    ]);
    children.push(
      makeStatusTable(
        ["Policy Name", "State", "Users / Groups", "Applications", "Grant Controls"],
        caRows,
        [2400, 1200, 1800, 2119, 2119],
        1
      ),
      gap()
    );
  } else {
    children.push(
      h2("2.2. Conditional Access Policies"),
      p("No Conditional Access policies were found in this tenant."),
      gap()
    );
  }

  // Directory Roles
  const dr = S.Identity.DirectoryRoles;
  if (dr && dr.length > 0) {
    children.push(h2("2.3. Directory Role Assignments"));
    
    // Group by role
    const roleGroups = {};
    dr.forEach(r => {
      if (!roleGroups[r.RoleName]) roleGroups[r.RoleName] = [];
      roleGroups[r.RoleName].push(r.MemberName || r.MemberId);
    });
    
    const roleRows = Object.entries(roleGroups).map(([role, members]) => [
      role, String(members.length), members.slice(0, 5).join(", ") + (members.length > 5 ? ` (+${members.length - 5} more)` : "")
    ]);
    
    children.push(
      makeTable(["Role", "Member Count", "Members"], roleRows, [3000, 1200, CW - 4200]),
      gap()
    );

    // Flag Global Admin count
    const gaCount = (roleGroups["Global Administrator"] || []).length;
    if (gaCount > 0) {
      children.push(p(`${gaCount} Global Administrator(s) assigned.`));
    }
  }

  // App Registrations
  const apps = S.Identity.AppRegistrations;
  if (apps && apps.length > 0) {
    children.push(
      h2("2.4. App Registrations"),
      p(`${apps.length} app registrations found.`)
    );
    const expired = apps.filter(a => a.ExpiredCreds && a.ExpiredCreds !== "None");
    if (expired.length > 0) {
      children.push(p(`${expired.length} app registration(s) have expired credentials.`));
    }
    
    // Show top 20 or all if fewer
    const appRows = apps.slice(0, 20).map(a => [
      a.DisplayName, a.AppId, a.SignInAudience, String(a.SecretCount || 0), a.ExpiredCreds || "None"
    ]);
    children.push(
      makeTable(
        ["Application", "App ID", "Sign-In Audience", "Secrets", "Expired Credentials"],
        appRows,
        [2500, 2300, 1600, 900, 2338]
      ),
      gap()
    );
    if (apps.length > 20) children.push(note(`Showing first 20 of ${apps.length} app registrations. See Excel for full list.`));
  }

  // Named Locations
  const nls = S.Identity.NamedLocations;
  if (nls && nls.length > 0) {
    children.push(
      h2("2.5. Named Locations"),
      p(`${nls.length} named locations configured.`)
    );
    const nlRows = nls.map(nl => [nl.Name, nl.Type, nl.IsTrusted ? "Yes" : "No", nl.Details || ""]);
    children.push(
      makeTable(["Name", "Type", "Trusted", "Details"], nlRows, [2500, 1500, 1200, 4438]),
      gap()
    );
  }

  // Groups Summary
  const gs = S.Identity.GroupSummary;
  if (gs && gs.TotalGroups) {
    children.push(
      h2("2.6. Groups Summary"),
      makeTable(
        ["Group Type", "Count"],
        [
          ["Total Groups", String(gs.TotalGroups || 0)],
          ["Security Groups", String(gs.SecurityGroups || 0)],
          ["Microsoft 365 Groups", String(gs.M365Groups || 0)],
          ["Dynamic Membership Groups", String(gs.DynamicGroups || 0)],
        ],
        [3200, CW - 3200]
      ),
      gap()
    );
  }

  // License Summary
  const lics = S.Identity.Licenses;
  if (lics && lics.length > 0) {
    children.push(
      h2("2.7. License Summary"),
      p(`${lics.length} license SKUs found in tenant.`)
    );
    const licRows = lics.map(l => [l.SKUName, String(l.Total || 0), String(l.Consumed || 0), String(l.Available || 0)]);
    children.push(
      makeTable(["SKU Name", "Total", "Consumed", "Available"], licRows, [3500, 1500, 1500, 3138]),
      gap()
    );
  }
} else {
  if (!entraIncluded) {
    // Entra was intentionally excluded from scope
    children.push(
      p("Entra ID was not in scope for this audit."),
      gap()
    );
  } else {
    // Entra was in scope but collection failed (e.g. auth cancelled, permissions issue)
    children.push(
      p("Entra ID data collection was attempted but failed. This may be due to authentication being cancelled or insufficient Graph API permissions."),
      gap()
    );
  }
}

children.push(pb());

// ============================================================
// 3. MANAGEMENT GROUPS & SUBSCRIPTIONS
// ============================================================
children.push(h1("3. Management Groups & Subscriptions"));

// Management Groups
const mgs = S.ManagementGroups;
if (mgs && mgs.length > 0) {
  children.push(
    h2("3.1. Management Group Hierarchy"),
    p(`${mgs.length} management groups found in the tenant hierarchy.`)
  );
  const mgRows = mgs.map(m => [
    m.DisplayName, m.ParentName || "Root",
    Array.isArray(m.Subscriptions) ? m.Subscriptions.join(", ") : (m.Subscriptions || "None"),
    Array.isArray(m.ChildMGs) ? m.ChildMGs.join(", ") : (m.ChildMGs || "None")
  ]);
  children.push(
    makeTable(["Management Group", "Parent", "Subscriptions", "Child MGs"], mgRows, [2800, 2000, 2619, 2219]),
    gap()
  );

  // Insert MG hierarchy diagram on its own landscape page
  if (mgDiagramResult) {
    mgDiagramSplitIdx = children.length; // mark split point
    mgDiagramContent = [
      h2("3.1a. Management Group Hierarchy Diagram"),
      new Paragraph({
        spacing: { before: 200, after: 200 },
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            type: "png",
            data: mgDiagramResult.pngBuffer,
            transformation: { width: DIAGRAM_WIDTH_PX, height: mgDiagramResult.heightPx },
            altText: { title: "Management Group Hierarchy", description: "Auto-generated diagram showing the Azure management group hierarchy and subscription placement.", name: "MG Hierarchy Diagram" },
          }),
        ],
      }),
      note("Figure 3.1: Management Group Hierarchy (auto-generated)"),
    ];
  }
} else {
  children.push(
    h2("3.1. Management Group Hierarchy"),
    p("No management groups exist beyond the tenant root group."),
    gap()
  );
}

// Subscriptions
const subs = S.Subscriptions;
if (subs && subs.length > 0) {
  children.push(
    h2("3.2. Subscriptions"),
    p(`${subs.length} active subscriptions found.`)
  );
  const subRows = subs.map(sub => [
    sub.SubscriptionName, sub.SubscriptionId, sub.State, sub.OfferType || "N/A"
  ]);
  children.push(
    makeTable(["Subscription Name", "Subscription ID", "State", "Offer Type"], subRows, [2800, 3200, 1200, 2438]),
    gap()
  );
}

children.push(pb());

// ============================================================
// 4. NETWORKING
// ============================================================
// Insert network heading + topology diagram on landscape page (or portrait if no diagram)
if (networkDiagramResult) {
  netDiagramSplitIdx = children.length; // mark split point (before heading)
  netDiagramContent = [
    h1("4. Networking"),
    p("The following diagram illustrates the network topology based on collected VNet, peering, and connectivity data."),
    new Paragraph({
      spacing: { before: 200, after: 200 },
      alignment: AlignmentType.CENTER,
      children: [
        new ImageRun({
          type: "png",
          data: networkDiagramResult.pngBuffer,
          transformation: { width: DIAGRAM_WIDTH_PX, height: networkDiagramResult.heightPx },
          altText: { title: "Network Topology", description: "Auto-generated diagram showing Azure virtual networks, peering connections, and network appliances.", name: "Network Topology Diagram" },
        }),
      ],
    }),
    note("Figure 4.1: Network Topology (auto-generated)"),
  ];
} else {
  children.push(h1("4. Networking"));
}

const net = S.Networking || {};

// 4.1 VNets
if (net.VirtualNetworks && net.VirtualNetworks.length > 0) {
  children.push(
    h2("4.1. Virtual Networks"),
    p(`${net.VirtualNetworks.length} virtual networks deployed.`)
  );
  const vnetRows = net.VirtualNetworks.map(v => [
    v.VNetName, v.ResourceGroup, v.Subscription, v.Location, v.AddressSpace, v.DNSServers || "Azure Default"
  ]);
  children.push(
    makeTable(
      ["VNet Name", "Resource Group", "Subscription", "Location", "Address Space", "DNS Servers"],
      vnetRows,
      [2000, 1600, 1500, 1000, 1500, 2038]
    ),
    gap()
  );
}

// 4.2 Subnets
if (net.Subnets && net.Subnets.length > 0) {
  children.push(
    h2("4.2. Subnets"),
    p(`${net.Subnets.length} subnets configured.`)
  );
  const snetRows = net.Subnets.map(sn => [
    sn.SubnetName, sn.VNetName, sn.AddressPrefix, sn.NSG || "None", sn.RouteTable || "None",
    sn.ServiceEndpoints || "None"
  ]);
  children.push(
    makeTable(
      ["Subnet", "VNet", "Address Range", "NSG", "Route Table", "Service Endpoints"],
      snetRows,
      [2000, 1600, 1300, 1500, 1500, 1738]
    ),
    gap()
  );

  // Note subnets without NSGs
  const noNSG = net.Subnets.filter(sn =>
    (!sn.NSG || sn.NSG === "None") &&
    !["GatewaySubnet", "AzureBastionSubnet", "AzureFirewallSubnet", "RouteServerSubnet"].includes(sn.SubnetName)
  );
  if (noNSG.length > 0) {
    children.push(p(`${noNSG.length} subnet(s) have no NSG associated: ${noNSG.map(n => n.SubnetName).join(", ")}`));
  }
}

// 4.3 VNet Peering
if (net.VNetPeerings && net.VNetPeerings.length > 0) {
  children.push(
    h2("4.3. VNet Peering"),
    p(`${net.VNetPeerings.length} peering connections.`)
  );
  const peerRows = net.VNetPeerings.map(pr => [
    pr.PeeringName, pr.SourceVNet, pr.RemoteVNet, pr.PeeringState,
    pr.AllowGatewayTransit ? "Yes" : "No", pr.UseRemoteGateways ? "Yes" : "No"
  ]);
  children.push(
    makeStatusTable(
      ["Peering Name", "Source VNet", "Remote VNet", "State", "Gateway Transit", "Remote Gateway"],
      peerRows,
      [2200, 1600, 1600, 1100, 1100, 2038],
      3
    ),
    gap()
  );
}

// 4.4 NSG Rules
if (net.NSGRules && net.NSGRules.length > 0) {
  children.push(
    h2("4.4. Network Security Group Rules"),
    p(`${net.NSGRules.length} custom NSG rules across all NSGs.`)
  );

  // Group by NSG name
  const nsgGroups = {};
  net.NSGRules.forEach(r => {
    if (!nsgGroups[r.NSGName]) nsgGroups[r.NSGName] = [];
    nsgGroups[r.NSGName].push(r);
  });

  for (const [nsgName, rules] of Object.entries(nsgGroups)) {
    children.push(h3(nsgName));
    const ruleRows = rules.map(r => [
      r.RuleName, r.Direction, String(r.Priority), r.Access, r.SourceAddress, r.DestPort || "*"
    ]);
    children.push(
      makeTable(
        ["Rule Name", "Direction", "Priority", "Action", "Source", "Dest Port"],
        ruleRows,
        [2000, 1000, 800, 800, 2800, 2238]
      ),
      gap()
    );
  }
}

// 4.5 Route Tables
if (net.RouteTables && net.RouteTables.length > 0) {
  children.push(
    h2("4.5. Route Tables (UDRs)"),
    p(`${net.RouteTables.length} custom routes configured.`)
  );
  const rtRows = net.RouteTables.map(r => [
    r.RouteTableName, r.RouteName, r.AddressPrefix, r.NextHopType, r.NextHopIP || "N/A", r.AssociatedSubnets || "None"
  ]);
  children.push(
    makeTable(
      ["Route Table", "Route Name", "Address Prefix", "Next Hop Type", "Next Hop IP", "Subnets"],
      rtRows,
      [1800, 1500, 1300, 1500, 1500, 2038]
    ),
    gap()
  );
}

// 4.6 VPN Gateways
if (net.VPNGateways && net.VPNGateways.length > 0) {
  children.push(h2("4.6. VPN Gateways"));
  net.VPNGateways.forEach(gw => {
    children.push(
      makeTable(
        ["Property", "Value"],
        [
          ["Gateway Name", gw.GatewayName],
          ["Type", gw.GatewayType],
          ["VPN Type", gw.VpnType],
          ["SKU", gw.SKU],
          ["Active-Active", gw.ActiveActive ? "Yes" : "No"],
          ["BGP Enabled", gw.EnableBGP ? `Yes (ASN: ${gw.BGPAsn || "N/A"})` : "No"],
          ["Location", gw.Location],
          ["Subscription", gw.Subscription],
        ],
        [3200, CW - 3200]
      ),
      gap()
    );
  });
}

// 4.7 VPN Connections
if (net.VPNConnections && net.VPNConnections.length > 0) {
  children.push(
    h2("4.7. VPN Connections"),
  );
  const connRows = net.VPNConnections.map(c => [
    c.ConnectionName, c.ConnectionType, c.ConnectionStatus, c.VPNGateway || "N/A", c.LocalNetGateway || "N/A", c.Protocol || "N/A"
  ]);
  children.push(
    makeStatusTable(
      ["Connection", "Type", "Status", "VPN Gateway", "Local Net Gateway", "Protocol"],
      connRows,
      [2000, 1200, 1200, 1700, 1700, 1738],
      2
    ),
    gap()
  );
}

// 4.8 Local Network Gateways
if (net.LocalNetGateways && net.LocalNetGateways.length > 0) {
  children.push(
    h2("4.8. Local Network Gateways"),
  );
  const lngRows = net.LocalNetGateways.map(l => [
    l.Name, l.GatewayIPAddress, l.AddressSpaces || "N/A", l.Subscription
  ]);
  children.push(
    makeTable(
      ["Name", "Gateway IP", "On-Prem Address Spaces", "Subscription"],
      lngRows,
      [2500, 1800, 3138, 2200]
    ),
    gap()
  );
}

// 4.9 Bastion
if (net.Bastions && net.Bastions.length > 0) {
  children.push(h2("4.9. Azure Bastion"));
  const bastRows = net.Bastions.map(b => [b.BastionName, b.SKU || "N/A", b.Location, b.Subscription]);
  children.push(
    makeTable(["Bastion Name", "SKU", "Location", "Subscription"], bastRows, [2800, 1500, 2700, 2638]),
    gap()
  );
}

// 4.10 Private DNS
if (net.PrivateDNSZones && net.PrivateDNSZones.length > 0) {
  children.push(
    h2("4.10. Private DNS Zones"),
    p(`${net.PrivateDNSZones.length} private DNS zones configured.`)
  );
  const dnsRows = net.PrivateDNSZones.map(z => [
    z.ZoneName, z.LinkedVNets || "None", z.AutoRegistration || "None", String(z.RecordSetCount || 0)
  ]);
  children.push(
    makeTable(["Zone Name", "Linked VNets", "Auto-Registration VNets", "Record Sets"], dnsRows, [3000, 2500, 2500, 1638]),
    gap()
  );
}

// 4.11 Public IPs
if (net.PublicIPs && net.PublicIPs.length > 0) {
  children.push(
    h2("4.11. Public IP Addresses"),
    p(`${net.PublicIPs.length} public IPs allocated.`)
  );
  const pipRows = net.PublicIPs.map(pip => [
    pip.Name, pip.IPAddress || "Dynamic", pip.SKU || "N/A", pip.AllocationMethod, pip.AssociatedTo || "Unassociated"
  ]);
  children.push(
    makeTable(["Name", "IP Address", "SKU", "Allocation", "Associated To"], pipRows, [2200, 1600, 1000, 1200, 3638]),
    gap()
  );
  
  const unassoc = net.PublicIPs.filter(p => p.AssociatedTo === "Unassociated");
  if (unassoc.length > 0) {
    children.push(p(`${unassoc.length} public IP(s) are not associated to a resource: ${unassoc.map(u => u.Name).join(", ")}`));
  }
}

// 4.12 Azure Firewalls
if (net.Firewalls && net.Firewalls.length > 0) {
  children.push(
    h2("4.12. Azure Firewalls"),
    p(`${net.Firewalls.length} Azure Firewall(s) found.`)
  );
  const rows = net.Firewalls.map(x => [s(x.FirewallName), s(x.ResourceGroup), s(x.SKU), s(x.ThreatIntelMode), s(x.PolicyName), s(x.Zones)]);
  children.push(
    makeTable(["Firewall Name", "Resource Group", "SKU", "Threat Intel", "Policy", "Zones"], rows, [1800, 1500, 1000, 1200, 1500, 1638]),
    gap()
  );
}

// 4.13 Application Gateways
if (net.ApplicationGateways && net.ApplicationGateways.length > 0) {
  children.push(
    h2("4.13. Application Gateways"),
    p(`${net.ApplicationGateways.length} Application Gateway(s) found.`)
  );
  const rows = net.ApplicationGateways.map(x => [s(x.Name), s(x.SKU), s(x.Tier), x.WAFEnabled ? "Yes" : "No", s(x.ListenerCount), s(x.BackendPoolCount), s(x.Zones)]);
  children.push(
    makeTable(["Name", "SKU", "Tier", "WAF Enabled", "Listeners", "Backend Pools", "Zones"], rows, [1600, 1100, 1100, 900, 1000, 1000, 1938]),
    gap()
  );
}

// 4.14 Load Balancers
if (net.LoadBalancers && net.LoadBalancers.length > 0) {
  children.push(
    h2("4.14. Load Balancers"),
    p(`${net.LoadBalancers.length} Load Balancer(s) found.`)
  );
  const rows = net.LoadBalancers.map(x => [s(x.Name), s(x.SKU), s(x.Type), s(x.FrontendCount), s(x.BackendPoolCount), s(x.RuleCount), s(x.ProbeCount)]);
  children.push(
    makeTable(["Name", "SKU", "Type", "Frontends", "Backends", "Rules", "Probes"], rows, [1800, 1000, 1000, 1100, 1100, 1100, 1538]),
    gap()
  );
}

// 4.15 NAT Gateways
if (net.NATGateways && net.NATGateways.length > 0) {
  children.push(
    h2("4.15. NAT Gateways"),
    p(`${net.NATGateways.length} NAT Gateway(s) found.`)
  );
  const rows = net.NATGateways.map(x => [s(x.Name), s(x.ResourceGroup), s(x.PublicIPs), s(x.IdleTimeout), s(x.Zones)]);
  children.push(
    makeTable(["Name", "Resource Group", "Public IPs", "Idle Timeout", "Zones"], rows, [2000, 1800, 2500, 1200, 2138]),
    gap()
  );
}

// 4.16 ExpressRoute Circuits
if (net.ExpressRoutes && net.ExpressRoutes.length > 0) {
  children.push(
    h2("4.16. ExpressRoute Circuits"),
    p(`${net.ExpressRoutes.length} ExpressRoute Circuit(s) found.`)
  );
  const rows = net.ExpressRoutes.map(x => [s(x.Name), s(x.Provider), s(x.Bandwidth), s(x.SKU), s(x.CircuitState), s(x.PeeringCount)]);
  children.push(
    makeStatusTable(["Name", "Provider", "Bandwidth", "SKU", "State", "Peerings"], rows, [2000, 1800, 1200, 1500, 1200, 1938], 4),
    gap()
  );
}

// 4.17 DDoS Protection Plans
if (net.DDoSPlans && net.DDoSPlans.length > 0) {
  children.push(
    h2("4.17. DDoS Protection Plans"),
    p(`${net.DDoSPlans.length} DDoS Protection Plan(s) found.`)
  );
  const rows = net.DDoSPlans.map(x => [s(x.Name), s(x.ResourceGroup), s(x.Location), s(x.ProtectedVNets)]);
  children.push(
    makeTable(["Name", "Resource Group", "Location", "Protected VNets"], rows, [2500, 2500, 2000, 2638]),
    gap()
  );
}

// 4.18 Private Endpoints
if (net.PrivateEndpoints && net.PrivateEndpoints.length > 0) {
  children.push(
    h2("4.18. Private Endpoints"),
    p(`${net.PrivateEndpoints.length} Private Endpoint(s) found.`)
  );
  const rows = net.PrivateEndpoints.map(x => [s(x.Name), s(x.TargetType), s(x.TargetResource), s(x.Subnet), s(x.Status)]);
  children.push(
    makeStatusTable(["Name", "Target Type", "Target Resource", "Subnet", "Status"], rows, [2000, 1500, 2000, 1500, 2638], 4),
    gap()
  );
}

// 4.19 NSG Flow Logs
if (net.NSGFlowLogs && net.NSGFlowLogs.length > 0) {
  children.push(
    h2("4.19. NSG Flow Logs"),
    p(`${net.NSGFlowLogs.length} NSG Flow Log(s) found.`)
  );
  const rows = net.NSGFlowLogs.map(x => [s(x.FlowLogName), s(x.NSG), x.Enabled ? "Yes" : "No", s(x.StorageAccount), s(x.RetentionDays), x.TrafficAnalytics ? "Yes" : "No"]);
  children.push(
    makeTable(["Flow Log", "NSG", "Enabled", "Storage", "Retention Days", "Traffic Analytics"], rows, [1500, 1500, 900, 1800, 1100, 1838]),
    gap()
  );
}

children.push(pb());

// ============================================================
// 5. COMPUTE
// ============================================================
children.push(h1("5. Compute Resources"));

const comp = S.Compute || {};

if (comp.VirtualMachines && comp.VirtualMachines.length > 0) {
  children.push(
    h2("5.1. Virtual Machines"),
    p(`${comp.VirtualMachines.length} virtual machines deployed.`)
  );
  const vmRows = comp.VirtualMachines.map(vm => [
    vm.VMName, vm.Subscription, vm.PrivateIPs || "N/A", vm.VMSize,
    vm.OSType || "N/A", vm.PowerState || "Unknown", vm.Zone || "None"
  ]);
  children.push(
    makeTable(
      ["VM Name", "Subscription", "Private IP(s)", "VM Size", "OS", "Power State", "Zone"],
      vmRows,
      [1500, 1400, 1200, 1400, 900, 1200, 1038]
    ),
    gap()
  );

  // Disk details table
  children.push(h2("5.2. VM Disk Configuration"));
  const diskRows = comp.VirtualMachines.map(vm => [
    vm.VMName, vm.DiskConfig || "N/A", String(vm.DataDiskCount || 0), vm.BootDiagnostics || "N/A"
  ]);
  children.push(
    makeTable(
      ["VM Name", "Disk Configuration", "Data Disks", "Boot Diagnostics"],
      diskRows,
      [2000, 4438, 1200, 2000]
    ),
    gap()
  );
}

// VM Extensions
if (comp.VMExtensions && comp.VMExtensions.length > 0) {
  children.push(
    h2("5.3. VM Extensions"),
    p(`${comp.VMExtensions.length} extensions installed across all VMs.`)
  );

  // Group by VM
  const extGroups = {};
  comp.VMExtensions.forEach(e => {
    if (!extGroups[e.VMName]) extGroups[e.VMName] = [];
    extGroups[e.VMName].push(e);
  });

  const extSummary = Object.entries(extGroups).map(([vm, exts]) => [
    vm, String(exts.length), exts.map(e => e.ExtensionName).join(", ")
  ]);
  children.push(
    makeTable(["VM Name", "Extension Count", "Extensions Installed"], extSummary, [2000, 1200, 6438]),
    gap()
  );
}

// AVD
if (comp.AVDHostPools && comp.AVDHostPools.length > 0) {
  children.push(
    h2("5.4. Azure Virtual Desktop"),
    p(`${comp.AVDHostPools.length} host pool(s) deployed.`)
  );
  comp.AVDHostPools.forEach(hp => {
    children.push(
      h3(hp.HostPoolName),
      makeTable(
        ["Property", "Value"],
        [
          ["Host Pool Type", hp.HostPoolType],
          ["Load Balancer", hp.LoadBalancerType],
          ["Max Session Limit", String(hp.MaxSessionLimit || "N/A")],
          ["Start VM on Connect", hp.StartVMOnConnect ? "Yes" : "No"],
          ["Validation Environment", hp.ValidationEnv ? "Yes" : "No"],
          ["Subscription", hp.Subscription],
        ],
        [3200, CW - 3200]
      ),
      gap()
    );
  });

  if (comp.AVDSessionHosts && comp.AVDSessionHosts.length > 0) {
    children.push(h3("Session Hosts"));
    const shRows = comp.AVDSessionHosts.map(sh => [
      sh.SessionHostName, sh.HostPoolName, sh.Status || "Unknown",
      String(sh.Sessions || 0), sh.AllowNewSession ? "Yes" : "No"
    ]);
    children.push(
      makeStatusTable(
        ["Session Host", "Host Pool", "Status", "Sessions", "Allow New"],
        shRows,
        [2800, 2000, 1500, 1000, 2338],
        2
      ),
      gap()
    );
  }
}

// 5.5 Managed Disks
if (comp.ManagedDisks && comp.ManagedDisks.length > 0) {
  children.push(
    h2("5.5. Managed Disks"),
    p(`${comp.ManagedDisks.length} Managed Disk(s) found.`)
  );
  const rows = comp.ManagedDisks.map(x => [s(x.DiskName), s(x.SizeGB), s(x.SKU), s(x.State), s(x.EncryptionType), s(x.AttachedTo), s(x.Zone)]);
  children.push(
    makeTable(["Disk Name", "Size (GB)", "SKU", "State", "Encryption", "Attached To", "Zone"], rows, [1600, 900, 1200, 1100, 1400, 1500, 938]),
    gap()
  );
}

// 5.6 Snapshots
if (comp.Snapshots && comp.Snapshots.length > 0) {
  children.push(
    h2("5.6. Snapshots"),
    p(`${comp.Snapshots.length} Snapshot(s) found.`)
  );
  const rows = comp.Snapshots.map(x => [s(x.SnapshotName), s(x.SourceDisk), s(x.SizeGB), x.Incremental ? "Yes" : "No", s(x.TimeCreated)]);
  children.push(
    makeTable(["Snapshot", "Source Disk", "Size (GB)", "Incremental", "Created"], rows, [2200, 2200, 1100, 1200, 2938]),
    gap()
  );
}

// 5.7 VM Scale Sets
if (comp.VMScaleSets && comp.VMScaleSets.length > 0) {
  children.push(
    h2("5.7. VM Scale Sets"),
    p(`${comp.VMScaleSets.length} VM Scale Set(s) found.`)
  );
  const rows = comp.VMScaleSets.map(x => [s(x.Name), s(x.SKU), s(x.Capacity), s(x.UpgradePolicy), s(x.Zones), s(x.OrchestrationMode)]);
  children.push(
    makeTable(["Name", "SKU", "Capacity", "Upgrade Policy", "Zones", "Orchestration"], rows, [2000, 1600, 1000, 1400, 1000, 1638]),
    gap()
  );
}

children.push(pb());

// ============================================================
// 6. STORAGE
// ============================================================
children.push(h1("6. Storage"));

const stor = S.Storage;
if (stor && stor.length > 0) {
  children.push(
    h2("6.1. Storage Accounts"),
    p(`${stor.length} storage accounts deployed.`)
  );
  const saRows = stor.map(sa => [
    sa.AccountName, sa.Subscription, sa.SKU || "N/A", sa.Kind || "N/A",
    sa.NetworkRuleSet || "Allow", String(sa.PrivateEndpoints || 0), sa.MinTLS || "N/A"
  ]);
  children.push(
    makeTable(
      ["Account", "Subscription", "SKU", "Kind", "Default Network Action", "Private Endpoints", "Min TLS"],
      saRows,
      [1700, 1400, 1200, 1100, 1200, 1000, 1038]
    ),
    gap()
  );
  
  const publicAccess = stor.filter(sa => sa.PublicAccess === true || sa.PublicAccess === "true");
  if (publicAccess.length > 0) {
    children.push(p(`${publicAccess.length} storage account(s) have public blob access enabled: ${publicAccess.map(s => s.AccountName).join(", ")}`));
  }
}

children.push(pb());

// ============================================================
// 7. BACKUP & RECOVERY
// ============================================================
children.push(h1("7. Backup & Recovery"));

const bk = S.Backup || {};

if (bk.RecoveryVaults && bk.RecoveryVaults.length > 0) {
  children.push(
    h2("7.1. Recovery Services Vaults"),
    p(`${bk.RecoveryVaults.length} Recovery Services vault(s) deployed.`)
  );
  const vltRows = bk.RecoveryVaults.map(v => [
    v.VaultName, v.Subscription, v.Location, v.Redundancy || "N/A",
    v.SoftDelete || "Unknown", v.Immutability || "Unknown"
  ]);
  children.push(
    makeTable(
      ["Vault Name", "Subscription", "Location", "Redundancy", "Soft Delete", "Immutability"],
      vltRows,
      [2200, 1600, 1200, 1200, 1200, 2238]
    ),
    gap()
  );
}

if (bk.BackupPolicies && bk.BackupPolicies.length > 0) {
  children.push(
    h2("7.2. Backup Policies"),
    p(`${bk.BackupPolicies.length} backup policies configured.`)
  );
  const polRows = bk.BackupPolicies.map(pol => [
    pol.PolicyName, pol.VaultName, pol.WorkloadType || "N/A",
    pol.Frequency || "N/A", pol.DailyRetention || "N/A",
    pol.WeeklyRetention || "N/A", pol.MonthlyRetention || "N/A"
  ]);
  children.push(
    makeTable(
      ["Policy", "Vault", "Workload", "Frequency", "Daily", "Weekly", "Monthly"],
      polRows,
      [1600, 1500, 1100, 1100, 1100, 1100, 2138]
    ),
    gap()
  );
}

if (bk.BackupItems && bk.BackupItems.length > 0) {
  children.push(
    h2("7.3. Protected Items"),
    p(`${bk.BackupItems.length} resources under backup protection.`)
  );
  const biRows = bk.BackupItems.map(bi => [
    bi.VMName, bi.VaultName, bi.PolicyName || "N/A",
    bi.ProtectionStatus || "N/A", bi.LastBackupStatus || "N/A",
    bi.LastBackupTime || "N/A"
  ]);
  children.push(
    makeStatusTable(
      ["Resource", "Vault", "Policy", "Protection Status", "Last Backup", "Last Backup Time"],
      biRows,
      [2000, 1500, 1200, 1300, 1200, 2438],
      3
    ),
    gap()
  );
}

// Unprotected VMs
if (bk.UnprotectedVMs && bk.UnprotectedVMs.length > 0) {
  children.push(
    h2("7.4. UNPROTECTED VIRTUAL MACHINES"),
    pBold(`${bk.UnprotectedVMs.length} running VMs have NO backup configured:`),
  );
  const upRows = bk.UnprotectedVMs.map(u => [u.VMName, u.Subscription, u.VMSize || "N/A"]);
  children.push(
    makeTable(["VM Name", "Subscription", "VM Size"], upRows, [3500, 3500, 2638]),
    gap()
  );
} else {
  children.push(
    h2("7.4. Backup Coverage"),
    p("All running VMs have backup protection configured."),
    gap()
  );
}

// 7.5 ASR Replicated Items
if (bk.ASRReplicatedItems && bk.ASRReplicatedItems.length > 0) {
  children.push(
    h2("7.5. ASR Replicated Items"),
    p(`${bk.ASRReplicatedItems.length} ASR Replicated Item(s) found.`)
  );
  const rows = bk.ASRReplicatedItems.map(x => [s(x.FriendlyName), s(x.Vault), s(x.ReplicationHealth), s(x.ProtectionState), s(x.ActiveLocation), s(x.TestFailoverState)]);
  children.push(
    makeStatusTable(["VM", "Vault", "Replication Health", "Protection State", "Active Location", "Test Failover"], rows, [2000, 1600, 1300, 1300, 1200, 1338], 2),
    gap()
  );
}

// 7.6 ASR Policies
if (bk.ASRPolicies && bk.ASRPolicies.length > 0) {
  children.push(
    h2("7.6. ASR Policies"),
    p(`${bk.ASRPolicies.length} ASR Replication Polic(ies) found.`)
  );
  const rows = bk.ASRPolicies.map(x => [s(x.PolicyName), s(x.Vault), s(x.ReplicationProvider), s(x.RecoveryPointRetention), s(x.AppConsistentFreq)]);
  children.push(
    makeTable(["Policy", "Vault", "Provider", "Recovery Point Retention", "App Consistent Freq"], rows, [2000, 1600, 1800, 2100, 2138]),
    gap()
  );
}

children.push(pb());

// ============================================================
// 8. SECURITY
// ============================================================
children.push(h1("8. Security"));

const sec = S.Security || {};

if (sec.DefenderPlans && sec.DefenderPlans.length > 0) {
  children.push(h2("8.1. Microsoft Defender for Cloud"));
  
  // Group by subscription
  const defBySub = {};
  sec.DefenderPlans.forEach(d => {
    if (!defBySub[d.Subscription]) defBySub[d.Subscription] = [];
    defBySub[d.Subscription].push(d);
  });

  for (const [sub, plans] of Object.entries(defBySub)) {
    children.push(h3(sub));
    const planRows = plans.map(pl => [pl.ResourceType, pl.PricingTier, pl.SubPlan || "N/A"]);
    children.push(
      makeStatusTable(["Resource Type", "Pricing Tier", "Sub-Plan"], planRows, [3500, 3069, 3069], 1),
      gap()
    );
  }
}

if (sec.SecureScores && sec.SecureScores.length > 0) {
  children.push(h2("8.2. Secure Scores"));
  const ssRows = sec.SecureScores.map(ss => [
    ss.Subscription, String(ss.CurrentScore || 0), String(ss.MaxScore || 0), `${ss.Percentage || 0}%`
  ]);
  children.push(
    makeTable(["Subscription", "Current Score", "Max Score", "Percentage"], ssRows, [3500, 1500, 1500, 3138]),
    gap()
  );
}

if (sec.KeyVaults && sec.KeyVaults.length > 0) {
  children.push(
    h2("8.3. Key Vaults"),
    p(`${sec.KeyVaults.length} Key Vault(s) deployed.`)
  );
  const kvRows = sec.KeyVaults.map(kv => [
    kv.VaultName, kv.Subscription, kv.SoftDelete ? "Yes" : "No",
    kv.PurgeProtection ? "Yes" : "No", kv.RBACAuth ? "RBAC" : "Access Policy",
    kv.PublicAccess || "N/A"
  ]);
  children.push(
    makeTable(
      ["Vault", "Subscription", "Soft Delete", "Purge Protection", "Auth Model", "Public Access"],
      kvRows,
      [2000, 1500, 1100, 1200, 1200, 2638]
    ),
    gap()
  );
}

// Defender Recommendations
if (sec.DefenderRecommendations && sec.DefenderRecommendations.length > 0) {
  children.push(
    h2("8.4. Defender Recommendations"),
    p(`${sec.DefenderRecommendations.length} non-healthy security assessments identified.`)
  );
  const drRows = sec.DefenderRecommendations.slice(0, 30).map(r => [
    r.AssessmentName, r.Severity || "N/A", r.Status || "N/A", r.ResourceType || "N/A"
  ]);
  children.push(
    makeStatusTable(
      ["Assessment", "Severity", "Status", "Resource Type"],
      drRows,
      [3500, 1200, 1500, 3438],
      2
    ),
    gap()
  );
  if (sec.DefenderRecommendations.length > 30) children.push(note(`Showing first 30 of ${sec.DefenderRecommendations.length} assessments. See Excel for full list.`));
}

children.push(pb());

// ============================================================
// 9. RBAC
// ============================================================
children.push(h1("9. Role-Based Access Control (RBAC)"));

const rbac = S.RBAC;
if (rbac && rbac.length > 0) {
  children.push(p(`${rbac.length} role assignments found across all subscriptions.`));
  
  // Summary by role
  const roleSummary = {};
  rbac.forEach(r => {
    if (!roleSummary[r.RoleName]) roleSummary[r.RoleName] = { total: 0, users: 0, groups: 0, spns: 0 };
    roleSummary[r.RoleName].total++;
    if (r.PrincipalType === "User") roleSummary[r.RoleName].users++;
    else if (r.PrincipalType === "Group") roleSummary[r.RoleName].groups++;
    else roleSummary[r.RoleName].spns++;
  });

  const rSumRows = Object.entries(roleSummary)
    .sort((a, b) => b[1].total - a[1].total)
    .slice(0, 20)
    .map(([role, counts]) => [role, String(counts.total), String(counts.users), String(counts.groups), String(counts.spns)]);

  children.push(
    h2("9.1. Role Assignment Summary"),
    makeTable(
      ["Role", "Total Assignments", "Users", "Groups", "Service Principals"],
      rSumRows,
      [3000, 1500, 1200, 1200, 2738]
    ),
    gap()
  );
  
  const directUsers = rbac.filter(r => r.PrincipalType === "User" && (r.ScopeLevel === "Subscription" || r.ScopeLevel === "ManagementGroup"));
  if (directUsers.length > 0) {
    children.push(
      p(`${directUsers.length} direct user role assignment(s) at subscription or management group level.`),
      gap()
    );
  }
}
children.push(pb());

// ============================================================
// 10. GOVERNANCE & POLICY
// ============================================================
children.push(h1("10. Governance & Policy"));

const pol = S.Policy || {};

if (pol.PolicyAssignments && pol.PolicyAssignments.length > 0) {
  children.push(
    h2("10.1. Azure Policy Assignments"),
    p(`${pol.PolicyAssignments.length} policy assignments found.`)
  );

  const paRows = pol.PolicyAssignments.map(pa => [
    pa.DisplayName || pa.AssignmentName, pa.IsInitiative ? "Initiative" : "Policy",
    pa.EnforcementMode || "Default", pa.Effect || "N/A",
    pa.Scope ? pa.Scope.split("/").slice(-2).join("/") : "N/A"
  ]);
  children.push(
    makeTable(
      ["Policy / Initiative", "Type", "Enforcement", "Effect", "Scope"],
      paRows,
      [3000, 1000, 1200, 1200, 3238]
    ),
    gap()
  );
}

// Missing standard policies
if (pol.MissingStdPolicies && pol.MissingStdPolicies.length > 0) {
  children.push(
    h2("10.2. Standard Policies Not Present"),
    p("The following standard policies were not found applied to any scope:"),
  );
  const mpRows = pol.MissingStdPolicies.map(mp => [mp, "Not Applied"]);
  children.push(
    makeStatusTable(
      ["Standard Policy", "Status"],
      mpRows,
      [5500, 4138],
      1
    ),
    gap()
  );
}

// Policy Compliance
if (pol.PolicyCompliance && pol.PolicyCompliance.length > 0) {
  children.push(
    h2("10.3. Policy Compliance Summary"),
  );
  const pcRows = pol.PolicyCompliance.map(pc => [
    pc.Subscription, pc.PolicyAssignment, String(pc.NonCompliantResources || 0), String(pc.NonCompliantPolicies || 0)
  ]);
  children.push(
    makeTable(
      ["Subscription", "Policy Assignment", "Non-Compliant Resources", "Non-Compliant Policies"],
      pcRows,
      [2500, 3000, 2069, 2069]
    ),
    gap()
  );
}

// Resource Locks
if (pol.ResourceLocks && pol.ResourceLocks.length > 0) {
  children.push(
    h2("10.4. Resource Locks"),
    p(`${pol.ResourceLocks.length} resource lock(s) configured.`)
  );
  const rlRows = pol.ResourceLocks.map(r => [
    r.LockName, r.LockLevel || "N/A", r.ScopeType || "N/A", r.Subscription || "N/A", r.Notes || ""
  ]);
  children.push(
    makeTable(
      ["Lock Name", "Lock Level", "Scope Type", "Subscription", "Notes"],
      rlRows,
      [2000, 1200, 1400, 1800, 3238]
    ),
    gap()
  );
}

// Custom RBAC Roles
if (pol.CustomRoles && pol.CustomRoles.length > 0) {
  children.push(
    h2("10.5. Custom RBAC Roles"),
    p(`${pol.CustomRoles.length} custom RBAC role(s) defined.`)
  );
  const crRows = pol.CustomRoles.map(r => [
    r.RoleName, r.Description || "N/A", r.AssignableScopes || "N/A", String(r.PermissionCount || 0)
  ]);
  children.push(
    makeTable(
      ["Role Name", "Description", "Assignable Scopes", "Permissions"],
      crRows,
      [2200, 3000, 2800, 1638]
    ),
    gap()
  );
}

// Policy Exemptions
if (pol.PolicyExemptions && pol.PolicyExemptions.length > 0) {
  children.push(
    h2("10.6. Policy Exemptions"),
    p(`${pol.PolicyExemptions.length} policy exemption(s) configured.`)
  );
  const peRows = pol.PolicyExemptions.map(r => [
    r.ExemptionName, r.Subscription || "N/A", r.Category || "N/A", r.ExpiresOn || "N/A"
  ]);
  children.push(
    makeTable(
      ["Exemption", "Subscription", "Category", "Expires On"],
      peRows,
      [3000, 2200, 1800, 2638]
    ),
    gap()
  );
}

children.push(pb());

// ============================================================
// 11. TAGGING
// ============================================================
children.push(h1("11. Tagging Standard"));

const tags = S.Tags || {};

if (tags.TagAuditSummary && tags.TagAuditSummary.length > 0) {
  children.push(
    h2("11.1. Tagging Coverage"),
  );
  const tagRows = tags.TagAuditSummary.map(t => [
    t.Subscription, String(t.TotalResources), String(t.TaggedResources),
    String(t.UntaggedResources), `${t.TaggingPercentage}%`
  ]);
  children.push(
    makeTable(
      ["Subscription", "Total Resources", "Tagged", "Untagged", "Coverage %"],
      tagRows,
      [2500, 1500, 1500, 1500, 2638]
    ),
    gap()
  );
}

if (tags.TagKeyUsage && tags.TagKeyUsage.length > 0) {
  children.push(
    h2("11.2. Tag Keys in Use"),
  );
  const tkRows = tags.TagKeyUsage.map(tk => [tk.TagKey, String(tk.UsageCount), tk.SampleValues || ""]);
  children.push(
    makeTable(["Tag Key", "Usage Count", "Sample Values"], tkRows, [2500, 1500, 5638]),
    gap()
  );
}

if (tags.MissingStdTags && tags.MissingStdTags.length > 0) {
  children.push(
    h2("11.3. Standard Tags Not Present"),
    p("The following standard tags were not found on any resources:"),
  );
  const mtRows = tags.MissingStdTags.map(mt => [mt, "Not Found"]);
  children.push(
    makeStatusTable(["Expected Tag", "Status"], mtRows, [5000, 4638], 1),
    gap()
  );
}

children.push(pb());

// ============================================================
// 12. MONITORING & LOGGING
// ============================================================
children.push(h1("12. Monitoring & Logging"));

const mon = S.Monitoring || {};

if (mon.LogWorkspaces && mon.LogWorkspaces.length > 0) {
  children.push(
    h2("12.1. Log Analytics Workspaces"),
  );
  const lwRows = mon.LogWorkspaces.map(w => [
    w.WorkspaceName, w.Subscription, w.Location, String(w.RetentionDays || "N/A"), w.SKU || "N/A"
  ]);
  children.push(
    makeTable(["Workspace", "Subscription", "Location", "Retention (Days)", "SKU"], lwRows, [2500, 1800, 1200, 1200, 2938]),
    gap()
  );
}

children.push(h2("12.2. Alert Rules"));
if (mon.AlertRules && mon.AlertRules.length > 0) {
  children.push(
    p(`${mon.AlertRules.length} metric alert rules configured.`)
  );
  const arRows = mon.AlertRules.map(a => [
    a.AlertName, String(a.Severity), a.Enabled ? "Enabled" : "Disabled",
    a.TargetResource || "N/A", a.ActionGroups || "None"
  ]);
  children.push(
    makeTable(["Alert Name", "Severity", "Status", "Target", "Action Group"], arRows, [2500, 900, 1000, 2700, 2538]),
    gap()
  );
} else {
  children.push(
    p("No metric alert rules are configured."),
    gap()
  );
}

children.push(h2("12.3. Action Groups"));
if (mon.ActionGroups && mon.ActionGroups.length > 0) {
  const agRows = mon.ActionGroups.map(ag => [
    ag.ActionGroupName, ag.EmailReceivers || "None", ag.SMSReceivers || "None", ag.WebhookReceivers || "None"
  ]);
  children.push(
    makeTable(["Action Group", "Email Receivers", "SMS Receivers", "Webhook Receivers"], agRows, [2500, 2800, 2169, 2169]),
    gap()
  );
} else {
  children.push(
    p("No action groups are configured."),
    gap()
  );
}

// Application Insights
if (mon.AppInsights && mon.AppInsights.length > 0) {
  children.push(
    h2("12.4. Application Insights"),
    p(`${mon.AppInsights.length} Application Insights instance(s) configured.`)
  );
  const aiRows = mon.AppInsights.map(a => [
    a.Name, a.AppType || "N/A", a.WorkspaceId || "N/A", String(a.RetentionDays || "N/A"), a.IngestionMode || "N/A"
  ]);
  children.push(
    makeTable(
      ["Name", "App Type", "Workspace", "Retention Days", "Ingestion Mode"],
      aiRows,
      [2000, 1500, 2000, 1200, 2938]
    ),
    gap()
  );
}

// Diagnostic Settings
if (mon.DiagnosticSettings && mon.DiagnosticSettings.length > 0) {
  children.push(
    h2("12.5. Diagnostic Settings"),
    p(`${mon.DiagnosticSettings.length} diagnostic setting(s) configured.`)
  );
  const dsRows = mon.DiagnosticSettings.map(d => [
    d.ResourceName || "N/A", d.ResourceType || "N/A", d.DiagSettingName || "N/A",
    d.WorkspaceId || "N/A", d.StorageAccount || "N/A", d.EventHub || "N/A"
  ]);
  children.push(
    makeTable(
      ["Resource", "Resource Type", "Setting Name", "Workspace", "Storage", "Event Hub"],
      dsRows,
      [1600, 1400, 1600, 1700, 1600, 1738]
    ),
    gap()
  );
}

// Scheduled Query Rules
if (mon.ScheduledQueryRules && mon.ScheduledQueryRules.length > 0) {
  children.push(
    h2("12.6. Scheduled Query Rules"),
    p(`${mon.ScheduledQueryRules.length} scheduled query rule(s) configured.`)
  );
  const sqrRows = mon.ScheduledQueryRules.map(r => [
    r.Name, String(r.Severity || "N/A"), r.Enabled ? "Yes" : "No", r.Description || "N/A"
  ]);
  children.push(
    makeTable(
      ["Name", "Severity", "Enabled", "Description"],
      sqrRows,
      [2500, 1200, 1000, 4938]
    ),
    gap()
  );
}

// Activity Log Alerts
if (mon.ActivityLogAlerts && mon.ActivityLogAlerts.length > 0) {
  children.push(
    h2("12.7. Activity Log Alerts"),
    p(`${mon.ActivityLogAlerts.length} activity log alert(s) configured.`)
  );
  const alaRows = mon.ActivityLogAlerts.map(a => [
    a.Name, a.Enabled ? "Yes" : "No", a.Description || "N/A"
  ]);
  children.push(
    makeTable(
      ["Name", "Enabled", "Description"],
      alaRows,
      [3000, 1200, 5438]
    ),
    gap()
  );
}

children.push(pb());

// ============================================================
// 13. COST MANAGEMENT
// ============================================================
children.push(h1("13. Cost Management"));

const costs = S.CostManagement;
if (costs && costs.length > 0) {
  children.push(
    h2("13.1. Budgets"),
  );
  const budRows = costs.map(b => [
    b.BudgetName, b.Subscription, String(b.Amount), b.TimeGrain || "Monthly",
    String(b.CurrentSpend || "N/A"), b.Notifications || "None"
  ]);
  children.push(
    makeTable(
      ["Budget", "Subscription", "Amount", "Period", "Current Spend", "Notifications"],
      budRows,
      [1800, 1500, 1200, 1000, 1200, 2938]
    ),
    gap()
  );
} else {
  children.push(
    p("No Azure budgets are configured on any subscription."),
    gap()
  );
}

children.push(pb());

// ============================================================
// 14. RESOURCE GROUPS
// ============================================================
children.push(h1("14. Resource Groups"));

const rgs = S.ResourceGroups;
if (rgs && rgs.length > 0) {
  children.push(
    p(`${rgs.length} resource groups across all subscriptions.`)
  );
  const rgRows = rgs.map(rg => [
    rg.ResourceGroupName, rg.Subscription, rg.Location, String(rg.ResourceCount || 0)
  ]);
  children.push(
    makeTable(
      ["Resource Group", "Subscription", "Location", "Resource Count"],
      rgRows,
      [3500, 2500, 1500, 2138]
    ),
    gap()
  );
}

children.push(pb());

// ============================================================
// 15. APPLICATION SERVICES
// ============================================================
children.push(h1("15. Application Services"));

const apps = S.AppServices || {};

if (apps.AppServicePlans && apps.AppServicePlans.length > 0) {
  children.push(
    h2("15.1. App Service Plans"),
    p(`${apps.AppServicePlans.length} App Service plan(s) deployed.`)
  );
  const aspRows = apps.AppServicePlans.map(a => [
    a.PlanName, a.Subscription, a.SKU, a.Tier, a.OS, String(a.WorkerCount || 0), a.ZoneRedundant ? "Yes" : "No"
  ]);
  children.push(
    makeTable(
      ["Plan Name", "Subscription", "SKU", "Tier", "OS", "Workers", "Zone Redundant"],
      aspRows,
      [1800, 1500, 900, 1000, 900, 900, 1638]
    ),
    gap()
  );
}

if (apps.WebApps && apps.WebApps.length > 0) {
  children.push(
    h2("15.2. Web Apps"),
    p(`${apps.WebApps.length} web application(s) deployed.`)
  );
  const waRows = apps.WebApps.map(w => [
    w.AppName, w.AppServicePlan || "N/A", w.State || "N/A",
    w.HTTPSOnly ? "Yes" : "No", w.MinTLS || "N/A", w.Runtime || "N/A",
    w.ManagedIdentity || "None"
  ]);
  children.push(
    makeTable(
      ["App Name", "App Service Plan", "State", "HTTPS Only", "Min TLS", "Runtime", "Identity"],
      waRows,
      [1600, 1400, 900, 900, 800, 1800, 1238]
    ),
    gap()
  );
  if (apps.WebApps.length > 20) children.push(note(`Showing all ${apps.WebApps.length} web apps. See Excel for additional details.`));
}

if (apps.FunctionApps && apps.FunctionApps.length > 0) {
  children.push(
    h2("15.3. Function Apps"),
    p(`${apps.FunctionApps.length} Function App(s) deployed.`)
  );
  const faRows = apps.FunctionApps.map(f => [
    f.AppName, f.Runtime || "N/A", f.RuntimeVersion || "N/A",
    f.OSType || "N/A", f.PlanType || "N/A", f.ManagedIdentity || "None"
  ]);
  children.push(
    makeTable(
      ["App Name", "Runtime", "Version", "OS", "Plan", "Identity"],
      faRows,
      [2000, 1200, 1200, 1000, 1800, 1438]
    ),
    gap()
  );
}

if ((!apps.AppServicePlans || apps.AppServicePlans.length === 0) &&
    (!apps.WebApps || apps.WebApps.length === 0) &&
    (!apps.FunctionApps || apps.FunctionApps.length === 0)) {
  children.push(p("No App Services, Web Apps, or Function Apps are deployed."), gap());
}

children.push(pb());

// ============================================================
// 16. DATABASES
// ============================================================
children.push(h1("16. Databases"));

const dbs = S.Databases || {};

if (dbs.SQLServers && dbs.SQLServers.length > 0) {
  children.push(
    h2("16.1. Azure SQL Servers"),
    p(`${dbs.SQLServers.length} Azure SQL Server(s) deployed.`)
  );
  const sqlSrvRows = dbs.SQLServers.map(srv => [
    srv.ServerName, srv.Subscription, srv.Version || "N/A",
    srv.AdminLogin || "N/A", srv.PublicAccess || "N/A", srv.MinTLS || "N/A"
  ]);
  children.push(
    makeTable(
      ["Server Name", "Subscription", "Version", "Admin Login", "Public Access", "Min TLS"],
      sqlSrvRows,
      [2000, 1600, 900, 1500, 1200, 2438]
    ),
    gap()
  );
}

if (dbs.SQLDatabases && dbs.SQLDatabases.length > 0) {
  children.push(
    h2("16.2. Azure SQL Databases"),
    p(`${dbs.SQLDatabases.length} SQL database(s) found (excluding master).`)
  );
  const sqlDbRows = dbs.SQLDatabases.map(db => [
    db.DatabaseName, db.ServerName, db.Edition || "N/A",
    db.ServiceTier || "N/A", String(db.MaxSizeGB || "N/A"),
    db.ZoneRedundant ? "Yes" : "No", db.Status || "N/A"
  ]);
  children.push(
    makeStatusTable(
      ["Database", "Server", "Edition", "Service Tier", "Max Size (GB)", "Zone Redundant", "Status"],
      sqlDbRows,
      [1600, 1400, 1000, 1200, 1000, 1000, 1438],
      6
    ),
    gap()
  );
}

if (dbs.CosmosDB && dbs.CosmosDB.length > 0) {
  children.push(
    h2("16.3. Cosmos DB Accounts"),
    p(`${dbs.CosmosDB.length} Cosmos DB account(s) deployed.`)
  );
  const cosmosRows = dbs.CosmosDB.map(c => [
    c.AccountName, c.Subscription, c.Kind || "N/A",
    c.ConsistencyLevel || "N/A", c.Locations || "N/A",
    c.AutoFailover ? "Yes" : "No"
  ]);
  children.push(
    makeTable(
      ["Account", "Subscription", "API", "Consistency", "Locations", "Auto-Failover"],
      cosmosRows,
      [1800, 1500, 1000, 1200, 2200, 938]
    ),
    gap()
  );
}

if ((!dbs.SQLServers || dbs.SQLServers.length === 0) &&
    (!dbs.SQLDatabases || dbs.SQLDatabases.length === 0) &&
    (!dbs.CosmosDB || dbs.CosmosDB.length === 0)) {
  children.push(p("No Azure SQL or Cosmos DB resources are deployed."), gap());
}

children.push(pb());

// ============================================================
// 17. CONTAINERS
// ============================================================
children.push(h1("17. Containers"));

const cont = S.Containers || {};

if (cont.AKSClusters && cont.AKSClusters.length > 0) {
  children.push(
    h2("17.1. AKS Clusters"),
    p(`${cont.AKSClusters.length} AKS cluster(s) deployed.`)
  );
  const aksRows = cont.AKSClusters.map(a => [
    a.ClusterName, a.Subscription, a.Version || "N/A",
    a.NodePools || "N/A", a.NetworkPlugin || "N/A",
    a.RBACEnabled ? "Yes" : "No", a.PrivateCluster ? "Yes" : "No"
  ]);
  children.push(
    makeTable(
      ["Cluster", "Subscription", "Version", "Node Pools", "Network Plugin", "RBAC", "Private"],
      aksRows,
      [1600, 1400, 900, 2200, 1200, 800, 1538]
    ),
    gap()
  );
}

if (cont.ContainerRegistries && cont.ContainerRegistries.length > 0) {
  children.push(
    h2("17.2. Container Registries"),
    p(`${cont.ContainerRegistries.length} container registr${cont.ContainerRegistries.length === 1 ? "y" : "ies"} deployed.`)
  );
  const acrRows = cont.ContainerRegistries.map(r => [
    r.RegistryName, r.Subscription, r.SKU || "N/A",
    r.AdminEnabled ? "Yes" : "No", r.LoginServer || "N/A",
    r.PublicAccess || "N/A"
  ]);
  children.push(
    makeTable(
      ["Registry", "Subscription", "SKU", "Admin Enabled", "Login Server", "Public Access"],
      acrRows,
      [1800, 1500, 1000, 1100, 2400, 1838]
    ),
    gap()
  );
}

if ((!cont.AKSClusters || cont.AKSClusters.length === 0) &&
    (!cont.ContainerRegistries || cont.ContainerRegistries.length === 0)) {
  children.push(p("No AKS clusters or container registries are deployed."), gap());
}

children.push(pb());

// ============================================================
// 18. AUTOMATION & ADVISOR
// ============================================================
children.push(h1("18. Automation & Advisor"));

const auto = S.Automation || {};

if (auto.AutomationAccounts && auto.AutomationAccounts.length > 0) {
  children.push(
    h2("18.1. Automation Accounts"),
    p(`${auto.AutomationAccounts.length} Automation Account(s) deployed.`)
  );
  const aaRows = auto.AutomationAccounts.map(a => [
    a.AccountName, a.ResourceGroup, a.Subscription,
    a.Location || "N/A", a.State || "N/A", a.ManagedIdentity || "None"
  ]);
  children.push(
    makeTable(
      ["Account", "Resource Group", "Subscription", "Location", "State", "Identity"],
      aaRows,
      [1800, 1500, 1500, 1200, 1000, 1638]
    ),
    gap()
  );
}

if (auto.Runbooks && auto.Runbooks.length > 0) {
  children.push(
    h2("18.2. Runbooks"),
    p(`${auto.Runbooks.length} runbook(s) found.`)
  );
  const rbRows = auto.Runbooks.slice(0, 30).map(r => [
    r.RunbookName, r.AutomationAccount || "N/A", r.Subscription,
    r.RunbookType || "N/A", r.State || "N/A", r.LastModified || "N/A"
  ]);
  children.push(
    makeTable(
      ["Runbook", "Automation Account", "Subscription", "Type", "State", "Last Modified"],
      rbRows,
      [2000, 1800, 1500, 1200, 1000, 1138]
    ),
    gap()
  );
  if (auto.Runbooks.length > 30) children.push(note(`Showing first 30 of ${auto.Runbooks.length} runbooks. See Excel for full list.`));
}

const advisor = S.Advisor;
if (advisor && advisor.length > 0) {
  children.push(
    h2("18.3. Azure Advisor Recommendations"),
    p(`${advisor.length} Advisor recommendation(s) retrieved.`)
  );

  // Group by category
  const advCats = {};
  advisor.forEach(r => {
    const cat = r.Category || "Other";
    if (!advCats[cat]) advCats[cat] = 0;
    advCats[cat]++;
  });
  const catRows = Object.entries(advCats).sort((a, b) => b[1] - a[1]).map(([cat, count]) => [cat, String(count)]);
  children.push(
    makeTable(["Category", "Count"], catRows, [4800, CW - 4800]),
    gap()
  );

  // Detail table (top 30)
  const advRows = advisor.slice(0, 30).map(r => [
    r.Category || "N/A", r.Impact || "N/A", r.Problem || "N/A", r.ResourceType || "N/A"
  ]);
  children.push(
    makeTable(
      ["Category", "Impact", "Problem", "Resource Type"],
      advRows,
      [1500, 1000, 4638, 2500]
    ),
    gap()
  );
  if (advisor.length > 30) children.push(note(`Showing first 30 of ${advisor.length} recommendations. See Excel for full list.`));
}

if ((!auto.AutomationAccounts || auto.AutomationAccounts.length === 0) &&
    (!auto.Runbooks || auto.Runbooks.length === 0) &&
    (!advisor || advisor.length === 0)) {
  children.push(p("No Automation Accounts or Advisor recommendations found."), gap());
}

children.push(pb());

// ============================================================
// 19. RESOURCE SUMMARY
// ============================================================
children.push(h1("19. Resource Summary"));

const resSummary = S.ResourceSummary;
if (resSummary && resSummary.length > 0) {
  children.push(
    p(`${resSummary.length} distinct resource types found across all subscriptions via Azure Resource Graph.`)
  );
  const resRows = resSummary.map(r => [r.ResourceType || "N/A", String(r.Count || 0)]);
  children.push(
    makeTable(
      ["Resource Type", "Count"],
      resRows,
      [6500, CW - 6500]
    ),
    gap()
  );
} else {
  children.push(p("Resource Graph summary not available."), gap());
}

// ============================================================
// ASSEMBLE DOCUMENT
// ============================================================
console.log("Assembling document...");

// ---- Reusable header / footer factories (each section needs its own instance) ----
function makeHeader() {
  return new Header({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: `${customerName} AS-IS Build Document  |  ${new Date().toLocaleString("en-GB", { month: "long", year: "numeric" })}`,
            font: "Segoe UI", size: 16, color: B.primary,
          }),
        ],
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: B.accent, space: 2 } },
        spacing: { after: 100 },
      }),
    ],
  });
}
function makeFooter() {
  return new Footer({
    children: [
      new Paragraph({
        children: [
          new TextRun({ text: "Confidential", font: "Segoe UI", size: 16, color: B.grey }),
          new TextRun({ text: "\ttieva.co.uk", font: "Segoe UI", size: 16, color: B.primary }),
          new TextRun({ text: "\tPage ", font: "Segoe UI", size: 16, color: B.grey }),
          new TextRun({ children: [PageNumber.CURRENT], font: "Segoe UI", size: 16, color: B.grey }),
          new TextRun({ text: " of ", font: "Segoe UI", size: 16, color: B.grey }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Segoe UI", size: 16, color: B.grey }),
        ],
        tabStops: [
          { type: TabStopType.CENTER, position: Math.round(TabStopPosition.MAX / 2) },
          { type: TabStopType.RIGHT, position: TabStopPosition.MAX },
        ],
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: B.accent, space: 2 } },
        spacing: { before: 100 },
      }),
    ],
  });
}

function makeSection(landscape, sectionChildren) {
  // Note: docx library auto-swaps width/height when orientation is LANDSCAPE,
  // so always pass dimensions in portrait order (PAGE_W × PAGE_H).
  return {
    properties: {
      page: {
        size: landscape
          ? { width: PAGE_W, height: PAGE_H, orientation: PageOrientation.LANDSCAPE }
          : { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
      },
    },
    headers: { default: makeHeader() },
    footers: { default: makeFooter() },
    children: sectionChildren,
  };
}

// ---- Split children into portrait / landscape sections at diagram points ----
const splitPoints = [];
if (mgDiagramContent && mgDiagramSplitIdx >= 0)
  splitPoints.push({ idx: mgDiagramSplitIdx, content: mgDiagramContent });
if (netDiagramContent && netDiagramSplitIdx >= 0)
  splitPoints.push({ idx: netDiagramSplitIdx, content: netDiagramContent });
splitPoints.sort((a, b) => a.idx - b.idx);

const docSections = [];
let cursor = 0;
for (const sp of splitPoints) {
  if (sp.idx > cursor) docSections.push(makeSection(false, children.slice(cursor, sp.idx)));
  docSections.push(makeSection(true, sp.content));
  cursor = sp.idx;
}
if (cursor < children.length) docSections.push(makeSection(false, children.slice(cursor)));
// Fallback: no diagrams at all → single portrait section
if (docSections.length === 0) docSections.push(makeSection(false, children));

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Segoe UI", size: 20 } },
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Segoe UI Semibold", color: B.primary },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Segoe UI Semibold", color: B.primary },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 },
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Segoe UI Semibold", color: B.text },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 },
      },
    ],
  },
  sections: docSections,
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync(outputDocx, buffer);
console.log(`\nDocument generated successfully!`);
console.log(`Output: ${outputDocx}`);
console.log(`Sections: 19`);
console.log(`\nRemember to right-click the Table of Contents in Word and select "Update Field" to populate it.`);

})().catch(err => {
  console.error("Fatal error:", err);
  process.exit(1);
});
