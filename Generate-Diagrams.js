/**
 * Generate-Diagrams.js
 * Generates SVG diagrams for Management Group hierarchy and Network topology.
 * Converts to PNG via sharp when available; falls back to raw SVG for
 * environments without native modules (e.g. Azure Cloud Shell).
 *
 * Exports: generateMgHierarchyDiagram, generateNetworkDiagram
 * Both return Promise<{ pngBuffer: Buffer, svgString: string, widthPx: number, heightPx: number } | null>
 */
let sharp;
try {
  sharp = require("sharp");
} catch (err) {
  // sharp not available (e.g. Azure Cloud Shell) - will output SVG only
  console.warn("Warning: sharp module not available. Diagrams will be SVG only (no PNG conversion).");
}

const COLORS = {
  primary: "2DA58E", accent: "B5E5E2", dark: "1A6B5A", light: "E0F5F3",
  text: "333333", grey: "666666", white: "FFFFFF",
};
const FONT = "'Segoe UI', Arial, sans-serif";
const SVG_DPI = 300; // render SVGs at 300 DPI for high quality output (default is 72)
const $ = (hex) => `#${hex}`;

/** XML-escape a string for safe SVG embedding. */
function esc(str) {
  if (str == null) return "";
  return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** Truncate to maxLen characters, appending ellipsis when needed. */
function trunc(text, max) {
  if (!text) return "";
  return text.length <= max ? text : text.slice(0, max - 1) + "\u2026";
}

// ═══════════════════════════════════════════════════════════════════════════════
//  MANAGEMENT GROUP HIERARCHY DIAGRAM
// ═══════════════════════════════════════════════════════════════════════════════
async function generateMgHierarchyDiagram(managementGroups, subscriptions, colors, maxWidth) {
  try {
    const col = { ...COLORS, ...(colors || {}) };
    maxWidth = maxWidth || 640;
    if (!managementGroups || managementGroups.length === 0) return null;

    // Step 1 --- Build tree from flat MG list --------------------------------
    const lookup = {};
    managementGroups.forEach((mg) => {
      lookup[mg.Name] = {
        displayName: mg.DisplayName || mg.Name,
        subs: Array.isArray(mg.Subscriptions) ? mg.Subscriptions : [],
        childMGs: [], parentKey: null,
      };
    });
    managementGroups.forEach((mg) => {
      if (mg.ParentId && mg.ParentId !== "Root") {
        const parentKey = mg.ParentId.split("/").pop();
        if (lookup[parentKey]) {
          lookup[parentKey].childMGs.push(lookup[mg.Name]);
          lookup[mg.Name].parentKey = parentKey;
        }
      }
    });
    // Find root (no parent)
    let root = Object.values(lookup).find((n) => !n.parentKey);
    if (!root) {
      const rm = managementGroups.find((mg) => mg.ParentId === "Root" || !mg.ParentId);
      if (rm) root = lookup[rm.Name];
    }
    if (!root) return null;

    // Layout constants
    const MG_BOX_WIDTH = 130, MG_BOX_HEIGHT = 30, SUB_BOX_WIDTH = 115, SUB_BOX_HEIGHT = 22;
    const HORIZONTAL_GAP = 12, VERTICAL_GAP = 45, PADDING = 20, TITLE_HEIGHT = 30, LEGEND_HEIGHT = 40;
    const SUB_COLLAPSE_THRESHOLD = 4; // collapse subs above this count

    // Step 2 --- Recursive layout (bottom-up widths, top-down positions) ------
    function buildNode(mg) {
      const kids = mg.childMGs.map(buildNode);
      // Subscription leaf nodes
      if (mg.subs.length > SUB_COLLAPSE_THRESHOLD) {
        kids.push({ type: "sub", label: `${mg.subs.length} Subscriptions`,
          w: SUB_BOX_WIDTH, h: SUB_BOX_HEIGHT, children: [], subtreeW: SUB_BOX_WIDTH });
      } else {
        for (const s of mg.subs) {
          const lbl = typeof s === "string" ? s : (s.DisplayName || s.Name || s.SubscriptionId || "Sub");
          kids.push({ type: "sub", label: lbl, w: SUB_BOX_WIDTH, h: SUB_BOX_HEIGHT, children: [], subtreeW: SUB_BOX_WIDTH });
        }
      }
      const childTotal = kids.length > 0
        ? kids.reduce((s, k) => s + k.subtreeW, 0) + HORIZONTAL_GAP * (kids.length - 1) : 0;
      return { type: "mg", label: mg.displayName, w: MG_BOX_WIDTH, h: MG_BOX_HEIGHT,
        children: kids, subtreeW: Math.max(MG_BOX_WIDTH, childTotal) };
    }
    const tree = buildNode(root);
    const naturalW = tree.subtreeW + PADDING * 2;
    const canvasW = Math.max(maxWidth, naturalW);

    function placeNodes(node, cx, cy) {
      node.x = cx - node.w / 2; node.y = cy; node.cx = cx;
      if (!node.children.length) return;
      const total = node.children.reduce((s, k) => s + k.subtreeW, 0) + HORIZONTAL_GAP * (node.children.length - 1);
      let sx = cx - total / 2;
      for (const ch of node.children) {
        placeNodes(ch, sx + ch.subtreeW / 2, cy + node.h + VERTICAL_GAP);
        sx += ch.subtreeW + HORIZONTAL_GAP;
      }
    }
    placeNodes(tree, canvasW / 2, PADDING + TITLE_HEIGHT);

    // Find max Y for canvas height
    let maxY = 0;
    (function walk(n) { maxY = Math.max(maxY, n.y + n.h); n.children.forEach(walk); })(tree);
    const canvasH = maxY + PADDING + LEGEND_HEIGHT;

    // Step 3 --- Generate SVG elements ----------------------------------------
    const lines = [], boxes = [];
    function collect(node, parent) {
      if (parent) {
        lines.push(`<line x1="${parent.cx}" y1="${parent.y + parent.h}" ` +
          `x2="${node.cx}" y2="${node.y}" stroke="${$(col.grey)}" stroke-width="1.5"/>`);
      }
      if (node.type === "mg") {
        boxes.push(`<rect x="${node.x}" y="${node.y}" width="${node.w}" height="${node.h}" ` +
          `rx="6" fill="${$(col.white)}" stroke="${$(col.primary)}" stroke-width="2"/>` +
          `<text x="${node.cx}" y="${node.y + node.h / 2 + 4}" text-anchor="middle" ` +
          `font-family="${FONT}" font-size="10" font-weight="bold" ` +
          `fill="${$(col.dark)}">${esc(trunc(node.label, 18))}</text>`);
      } else {
        boxes.push(`<rect x="${node.x}" y="${node.y}" width="${node.w}" height="${node.h}" ` +
          `rx="4" fill="${$(col.light)}" stroke="${$(col.accent)}" stroke-width="1.5"/>` +
          `<text x="${node.cx}" y="${node.y + node.h / 2 + 3}" text-anchor="middle" ` +
          `font-family="${FONT}" font-size="8" fill="${$(col.text)}">${esc(trunc(node.label, 18))}</text>`);
      }
      node.children.forEach((ch) => collect(ch, node));
    }
    collect(tree, null);

    // Legend
    const ly = canvasH - LEGEND_HEIGHT + 8;
    const legItems = [
      { fill: $(col.white), stroke: $(col.primary), sw: 2, rx: 6, label: "Management Group" },
      { fill: $(col.light), stroke: $(col.accent), sw: 1.5, rx: 4, label: "Subscription" },
    ];
    const legParts = []; let lx = PADDING;
    for (const it of legItems) {
      legParts.push(`<rect x="${lx}" y="${ly}" width="14" height="10" rx="${it.rx}" ` +
        `fill="${it.fill}" stroke="${it.stroke}" stroke-width="${it.sw}"/>` +
        `<text x="${lx + 18}" y="${ly + 8}" font-family="${FONT}" font-size="8" ` +
        `fill="${$(col.grey)}">${esc(it.label)}</text>`);
      lx += 18 + it.label.length * 5 + 16;
    }

    const svg = [
      `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${canvasW} ${canvasH}" width="${canvasW}" height="${canvasH}">`,
      `<rect width="100%" height="100%" fill="${$(col.white)}"/>`,
      `<text x="${canvasW / 2}" y="${PADDING + 14}" text-anchor="middle" font-family="${FONT}" ` +
        `font-size="14" font-weight="bold" fill="${$(col.primary)}">Management Group Hierarchy</text>`,
      ...lines, ...boxes, ...legParts, `</svg>`,
    ].join("\n");

    // Step 4 --- Convert to PNG (if sharp available) or return SVG -------------
    const svgString = svg;
    if (!sharp) {
      return { pngBuffer: Buffer.from(svg), svgString, widthPx: canvasW, heightPx: canvasH, isSvg: true };
    }
    const dpiScale = SVG_DPI / 72;
    let pngBuffer;
    if (canvasW > maxWidth) {
      const targetW = Math.round(maxWidth * dpiScale);
      const targetH = Math.round(canvasH * (maxWidth / canvasW) * dpiScale);
      pngBuffer = await sharp(Buffer.from(svg), { density: SVG_DPI })
        .resize({ width: targetW, height: targetH, fit: "fill" })
        .png().toBuffer();
      return { pngBuffer, svgString, widthPx: maxWidth, heightPx: Math.round(canvasH * (maxWidth / canvasW)) };
    }
    const hiW = Math.round(canvasW * dpiScale);
    const hiH = Math.round(canvasH * dpiScale);
    pngBuffer = await sharp(Buffer.from(svg), { density: SVG_DPI })
      .resize({ width: hiW, height: hiH, fit: "fill" })
      .png().toBuffer();
    return { pngBuffer, svgString, widthPx: canvasW, heightPx: canvasH };
  } catch (err) {
    console.error("generateMgHierarchyDiagram error:", err.message || err);
    return null;
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
//  NETWORK TOPOLOGY DIAGRAM
// ═══════════════════════════════════════════════════════════════════════════════
async function generateNetworkDiagram(networking, colors, maxWidth) {
  try {
    const col = { ...COLORS, ...(colors || {}) };
    maxWidth = maxWidth || 640;
    if (!networking) return null;

    const vnets = networking.VirtualNetworks || [];
    const subnets = networking.Subnets || [];
    const peerings = networking.VNetPeerings || [];
    const bastions = networking.Bastions || [];
    const firewalls = networking.Firewalls || [];
    const vpnGWs = networking.VPNGateways || [];
    const appGWs = networking.ApplicationGateways || [];
    const lbs = networking.LoadBalancers || [];
    const exRoutes = networking.ExpressRouteCircuits || [];
    const localGWs = networking.LocalNetworkGateways || [];
    if (vnets.length === 0) return null;

    // Step 1 --- Prepare data -------------------------------------------------
    // Helper: extract name from various PS property conventions
    const vn = (v) => v.VNetName || v.Name || v.name || "";
    const sn = (s) => s.SubnetName || s.Name || s.name || "";
    const rn = (r) => r.BastionName || r.FirewallName || r.GatewayName || r.Name || r.name || "";

    const vnetNames = new Set(vnets.map(vn));

    // Deduplicate bidirectional peerings
    const peerSet = new Set(), uniqPeers = [];
    for (const p of peerings) {
      const src = p.SourceVNet || p.sourceVNet || p.VNetName || "";
      const rem = p.RemoteVNet || p.remoteVNet || p.RemoteVNetName || "";
      if (!src || !rem) continue;
      const key = [src, rem].sort().join("|||");
      if (!peerSet.has(key)) {
        peerSet.add(key);
        uniqPeers.push({ source: src, remote: rem, state: p.PeeringState || p.State || "Connected" });
      }
    }

    // Count connections per VNet to detect hub
    const peerCnt = {};
    for (const p of uniqPeers) {
      peerCnt[p.source] = (peerCnt[p.source] || 0) + 1;
      peerCnt[p.remote] = (peerCnt[p.remote] || 0) + 1;
    }
    let hubName = null, hubMax = 1;
    for (const [nm, cnt] of Object.entries(peerCnt)) {
      if (cnt > hubMax && vnetNames.has(nm)) { hubMax = cnt; hubName = nm; }
    }

    // External VNets (in peerings but not in VirtualNetworks)
    const extVNets = new Set();
    for (const p of uniqPeers) {
      if (!vnetNames.has(p.source)) extVNets.add(p.source);
      if (!vnetNames.has(p.remote)) extVNets.add(p.remote);
    }

    // Map subnets to VNets
    const vnetSubs = {};
    for (const s of subnets) {
      const svn = s.VNetName || s.vnetName || "";
      (vnetSubs[svn] = vnetSubs[svn] || []).push(s);
    }

    // Map appliances to VNets via subnet names and explicit resource arrays
    const vnetAppl = {};
    const addA = (vnn, t) => { (vnetAppl[vnn] = vnetAppl[vnn] || new Set()).add(t); };
    for (const name of vnetNames) {
      for (const s of (vnetSubs[name] || [])) {
        const subName = sn(s).toLowerCase();
        if (subName.includes("azurebastionsubnet")) addA(name, "Bastion");
        if (subName.includes("azurefirewallsubnet")) addA(name, "Firewall");
        if (subName === "gatewaysubnet") addA(name, "VPN GW");
      }
    }
    const mapRes = (arr, type) => arr.forEach((r) => { const v = r.VNetName || r.vnetName; if (v) addA(v, type); });
    mapRes(bastions, "Bastion"); mapRes(firewalls, "Firewall"); mapRes(vpnGWs, "VPN GW");
    mapRes(appGWs, "AppGW"); mapRes(lbs, "LB");

    // Step 2 & 3 --- Layout constants and VNet box computation ----------------
    const HEADER_HEIGHT = 28, ADDRESS_ROW_HEIGHT = 16, SUBNET_LINE_HEIGHT = 14, SUBNET_DETAIL_HEIGHT = 10;
    const APPLIANCE_HEIGHT = 16, BOX_PADDING = 10;
    const MAX_VISIBLE_SUBNETS = 12, TITLE_HEIGHT = 30, LEGEND_HEIGHT = 55, TOP_PADDING = 20;

    function makeBox(vnet, isHub, isExt) {
      const name = typeof vnet === "string" ? vnet : vn(vnet);
      const addrRaw = typeof vnet === "string" ? ""
        : (vnet.AddressSpace || vnet.addressSpace || vnet.AddressPrefixes || "");
      const addr = Array.isArray(addrRaw) ? addrRaw.join(", ") : String(addrRaw || "");
      const dnsRaw = typeof vnet === "string" ? ""
        : (vnet.DNSServers || vnet.dnsServers || "");
      const dns = dnsRaw && dnsRaw !== "Azure Default" ? dnsRaw : "";
      const subs = vnetSubs[name] || [];
      const appls = vnetAppl[name] ? Array.from(vnetAppl[name]) : [];
      // Wider boxes for detailed subnet info
      const w = isHub ? 380 : isExt ? 220 : 320;
      const shown = subs.slice(0, MAX_VISIBLE_SUBNETS), extra = subs.length - shown.length;
      // Each subnet now gets a main line + detail line (NSG/IP)
      let h = BOX_PADDING + HEADER_HEIGHT + (addr ? ADDRESS_ROW_HEIGHT : 0) + (dns ? SUBNET_DETAIL_HEIGHT : 0);
      h += shown.length * (SUBNET_LINE_HEIGHT + SUBNET_DETAIL_HEIGHT);
      if (extra > 0) h += SUBNET_LINE_HEIGHT;
      if (appls.length > 0) h += APPLIANCE_HEIGHT + 4;
      h += BOX_PADDING;
      if (isExt) h = Math.max(h, 50);
      return { name, addr, dns, subs: shown, extra, appls, w, h, isHub, isExt, x: 0, y: 0, cx: 0, cy: 0 };
    }

    const vboxes = {}, spokes = [];
    for (const v of vnets) {
      const nm = vn(v), hub = nm === hubName;
      vboxes[nm] = makeBox(v, hub, false);
      if (!hub) spokes.push(nm);
    }
    for (const ext of extVNets) {
      vboxes[ext] = makeBox(ext, false, true);
      spokes.push(ext);
    }

    // Position boxes (hub-spoke or grid)
    const hubTopY = TOP_PADDING + TITLE_HEIGHT + 20;
    let canvasH;

    if (hubName && vboxes[hubName]) {
      const hb = vboxes[hubName];
      hb.cx = maxWidth / 2; hb.cy = hubTopY + hb.h / 2;
      hb.x = hb.cx - hb.w / 2; hb.y = hb.cy - hb.h / 2;

      const intSpokes = spokes.filter((n) => !extVNets.has(n));
      const extList = spokes.filter((n) => extVNets.has(n));
      // Gap between hub bottom and spoke tops for clean routing
      const spokeTopY = hb.y + hb.h + 60;

      // Calculate total width needed for spokes with proper gaps
      const SPOKE_GAP = 20;
      const totalSpokeW = intSpokes.reduce((s, n) => s + (vboxes[n] ? vboxes[n].w : 0), 0) + SPOKE_GAP * Math.max(0, intSpokes.length - 1);
      const spokeStartX = Math.max(TOP_PADDING, (maxWidth - totalSpokeW) / 2);

      // Position spokes side by side with gaps (not overlapping)
      let sx = spokeStartX;
      for (let i = 0; i < intSpokes.length; i++) {
        const b = vboxes[intSpokes[i]];
        b.x = sx; b.y = spokeTopY;
        b.cx = sx + b.w / 2; b.cy = spokeTopY + b.h / 2;
        sx += b.w + SPOKE_GAP;
      }

      // If spokes overflow the canvas, widen it
      if (sx > maxWidth) maxWidth = sx + TOP_PADDING;
      // External VNets further below - calculate based on tallest spoke
      if (extList.length > 0) {
        const maxSpokeBottom = intSpokes.reduce((m, n) => {
          const bx = vboxes[n]; return bx ? Math.max(m, bx.y + bx.h) : m;
        }, spokeTopY);
        const ey = maxSpokeBottom + 60;
        const extTotalW = extList.reduce((s, n) => s + (vboxes[n] ? vboxes[n].w : 0), 0) + SPOKE_GAP * Math.max(0, extList.length - 1);
        let ex = Math.max(TOP_PADDING, (maxWidth - extTotalW) / 2);
        for (let i = 0; i < extList.length; i++) {
          const b = vboxes[extList[i]];
          b.x = ex; b.y = ey;
          b.cx = ex + b.w / 2; b.cy = ey + b.h / 2;
          ex += b.w + SPOKE_GAP;
        }
      }
      let mxB = 0, mxR = 0;
      for (const b of Object.values(vboxes)) {
        mxB = Math.max(mxB, b.y + b.h);
        mxR = Math.max(mxR, b.x + b.w);
      }
      canvasH = mxB + TOP_PADDING + LEGEND_HEIGHT;
      // Ensure canvas is wide enough for all boxes
      if (mxR + TOP_PADDING > maxWidth) maxWidth = mxR + TOP_PADDING;
    } else {
      // Grid layout
      const all = [...vnetNames, ...extVNets];
      let gx = TOP_PADDING, gy = TOP_PADDING + TITLE_HEIGHT + 10, rowH = 0;
      for (const nm of all) {
        const b = vboxes[nm]; if (!b) continue;
        if (gx + b.w > maxWidth - TOP_PADDING) { gx = TOP_PADDING; gy += rowH + 16; rowH = 0; }
        b.x = gx; b.y = gy; b.cx = gx + b.w / 2; b.cy = gy + b.h / 2;
        gx += b.w + 16; rowH = Math.max(rowH, b.h);
      }
      canvasH = gy + rowH + TOP_PADDING + LEGEND_HEIGHT;
    }
    canvasH = Math.max(canvasH, 300);
    // Re-center hub if canvas widened
    if (hubName && vboxes[hubName]) {
      const hb = vboxes[hubName];
      hb.cx = maxWidth / 2; hb.x = hb.cx - hb.w / 2;
    }
    const canvasW = maxWidth;

    // Step 4 --- Generate SVG -------------------------------------------------
    const sL = [], sB = []; // svg lines, svg boxes

    // Peering connection lines with connection dots
    const DOT_R = 4; // radius of connection dot
    for (const p of uniqPeers) {
      const a = vboxes[p.source], b = vboxes[p.remote];
      if (!a || !b) continue;
      const conn = /^connected$/i.test(p.state);
      const clr = conn ? $(col.primary) : "#E57373";
      const dash = conn ? "" : ` stroke-dasharray="6,3"`;

      const upper = a.cy < b.cy ? a : b;
      const lower = a.cy < b.cy ? b : a;

      if (Math.abs(a.y - b.y) < 20) {
        // Same row: connect side edges
        const left = a.cx < b.cx ? a : b;
        const right = a.cx < b.cx ? b : a;
        const lx1 = left.x + left.w, ly1 = left.cy;
        const lx2 = right.x, ly2 = right.cy;
        sL.push(`<line x1="${lx1}" y1="${ly1}" x2="${lx2}" y2="${ly2}" stroke="${clr}" stroke-width="1.5"${dash}/>`);
        sL.push(`<circle cx="${lx1}" cy="${ly1}" r="${DOT_R}" fill="${clr}"/>`);
        sL.push(`<circle cx="${lx2}" cy="${ly2}" r="${DOT_R}" fill="${clr}"/>`);
      } else {
        // Different rows: bezier from bottom of upper to top of lower
        const x1 = upper.cx, y1 = upper.y + upper.h;
        const x2 = lower.cx, y2 = lower.y;
        const cp1y = y1 + (y2 - y1) * 0.5;
        const cp2y = y2 - (y2 - y1) * 0.5;
        sL.push(`<path d="M ${x1} ${y1} C ${x1} ${cp1y}, ${x2} ${cp2y}, ${x2} ${y2}" ` +
          `fill="none" stroke="${clr}" stroke-width="1.5"${dash}/>`);
        // Connection dots at both ends
        sL.push(`<circle cx="${x1}" cy="${y1}" r="${DOT_R}" fill="${clr}"/>`);
        sL.push(`<circle cx="${x2}" cy="${y2}" r="${DOT_R}" fill="${clr}"/>`);
      }
    }

    // Local Network Gateways (on-prem boxes) and VPN lines
    if (localGWs.length > 0) {
      const lgy = TOP_PADDING + TITLE_HEIGHT - 10;
      const step = localGWs.length > 1 ? (canvasW - TOP_PADDING * 2 - 120) / (localGWs.length - 1) : 0;
      for (let i = 0; i < localGWs.length; i++) {
        const gw = localGWs[i], gwNm = gw.GatewayName || gw.Name || gw.name || "On-Prem";
        const gwIp = gw.GatewayIpAddress || gw.gatewayIpAddress || "";
        const lx = localGWs.length === 1 ? canvasW / 2 - 60 : TOP_PADDING + step * i;
        sB.push(`<rect x="${lx}" y="${lgy}" width="120" height="32" rx="6" ` +
          `fill="${$(col.dark)}" stroke="${$(col.primary)}" stroke-width="1.5"/>` +
          `<text x="${lx + 60}" y="${lgy + 13}" text-anchor="middle" font-family="${FONT}" ` +
          `font-size="8" font-weight="bold" fill="${$(col.white)}">${esc(trunc(gwNm, 18))}</text>` +
          (gwIp ? `<text x="${lx + 60}" y="${lgy + 25}" text-anchor="middle" font-family="${FONT}" ` +
            `font-size="7" fill="${$(col.accent)}">${esc(gwIp)}</text>` : ""));
        const tgt = hubName || (vnets[0] && vn(vnets[0]));
        if (tgt && vboxes[tgt]) {
          const tb = vboxes[tgt];
          sL.push(`<path d="M ${lx + 60} ${lgy + 32} C ${lx + 60} ${tb.y - 10}, ${tb.cx} ${lgy + 50}, ${tb.cx} ${tb.y}" ` +
            `fill="none" stroke="#FFB74D" stroke-width="1.5" stroke-dasharray="6,3"/>`);
          sL.push(`<circle cx="${lx + 60}" cy="${lgy + 32}" r="${DOT_R}" fill="#FFB74D"/>`);
          sL.push(`<circle cx="${tb.cx}" cy="${tb.y}" r="${DOT_R}" fill="#FFB74D"/>`);
        }
      }
    }

    // ExpressRoute lines (thick purple) - connect from top edge of box
    for (const er of exRoutes) {
      const vnName = er.VNetName || er.vnetName || "";
      if (vnName && vboxes[vnName]) {
        const bx = vboxes[vnName];
        sL.push(`<line x1="${bx.cx}" y1="${bx.y}" x2="${bx.cx}" y2="${TOP_PADDING}" ` +
          `stroke="#BA68C8" stroke-width="3"/>`);
        sL.push(`<circle cx="${bx.cx}" cy="${bx.y}" r="5" fill="#BA68C8"/>`);
      }
    }

    // Appliance badge colour map
    const AC = {
      Bastion: { bg: "#64B5F6", fg: $(col.white) },
      Firewall: { bg: "#E57373", fg: $(col.white) },
      "VPN GW": { bg: "#FFB74D", fg: $(col.text) },
      AppGW: { bg: "#BA68C8", fg: $(col.white) },
      LB: { bg: "#4DB6AC", fg: $(col.white) },
    };

    // Draw VNet boxes
    for (const bx of Object.values(vboxes)) {
      const { x, y, w, h, name, addr, dns, subs, extra, appls, isHub, isExt } = bx;
      // Main rect with subtle shadow
      if (isExt) {
        sB.push(`<rect x="${x + 2}" y="${y + 2}" width="${w}" height="${h}" rx="8" fill="#00000010"/>` +
          `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="8" ` +
          `fill="#F5F5F5" stroke="${$(col.grey)}" stroke-width="1.5" stroke-dasharray="5,3"/>`);
      } else {
        sB.push(`<rect x="${x + 2}" y="${y + 2}" width="${w}" height="${h}" rx="8" fill="#00000010"/>` +
          `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="8" ` +
          `fill="${$(col.light)}" stroke="${$(col.primary)}" stroke-width="2"/>`);
      }
      // Header bar (top rounded, bottom square to merge with body)
      const hFill = isExt ? $(col.grey) : $(col.primary);
      sB.push(`<rect x="${x}" y="${y}" width="${w}" height="${HEADER_HEIGHT}" rx="8" fill="${hFill}"/>` +
        `<rect x="${x}" y="${y + 12}" width="${w}" height="${HEADER_HEIGHT - 12}" fill="${hFill}"/>`);
      // Header text — truncation based on box width
      const hdrMax = Math.floor((w - 20) / 5.5);
      sB.push(`<text x="${x + w / 2}" y="${y + 18}" text-anchor="middle" font-family="${FONT}" ` +
        `font-size="10" font-weight="bold" fill="${$(col.white)}">${esc(trunc(name, hdrMax))}</text>`);

      let cy = y + HEADER_HEIGHT + 4;
      // Address space (show all prefixes)
      if (addr) {
        sB.push(`<text x="${x + w / 2}" y="${cy + 10}" text-anchor="middle" font-family="${FONT}" ` +
          `font-size="8" font-weight="bold" fill="${$(col.dark)}">${esc(trunc(addr, 50))}</text>`);
        cy += ADDRESS_ROW_HEIGHT;
      }
      // DNS servers
      if (dns) {
        sB.push(`<text x="${x + w / 2}" y="${cy + 8}" text-anchor="middle" font-family="${FONT}" ` +
          `font-size="7" fill="${$(col.grey)}">DNS: ${esc(trunc(dns, 40))}</text>`);
        cy += SUBNET_DETAIL_HEIGHT;
      }
      // Subnets — show name + address prefix on main line, NSG + details on sub-line
      // Calculate max chars based on box width (approx 4.5px per char at font-size 7.5)
      const subMaxChars = Math.floor((w - BOX_PADDING * 2) / 4.5);
      const detailMaxChars = Math.floor((w - BOX_PADDING * 2 - 4) / 3.8);
      for (const s of subs) {
        const subNm = sn(s), sa = s.AddressPrefix || s.addressPrefix || "";
        // Main subnet line: name and CIDR
        const lbl = subNm && sa ? `${subNm}  ${sa}` : (subNm || sa);
        // Subnet background stripe for readability
        sB.push(`<rect x="${x + 4}" y="${cy}" width="${w - 8}" height="${SUBNET_LINE_HEIGHT + SUBNET_DETAIL_HEIGHT - 2}" rx="2" fill="${$(col.white)}" opacity="0.5"/>`);
        sB.push(`<text x="${x + BOX_PADDING}" y="${cy + 10}" font-family="${FONT}" font-size="7.5" ` +
          `font-weight="bold" fill="${$(col.text)}">${esc(trunc(lbl, subMaxChars))}</text>`);
        cy += SUBNET_LINE_HEIGHT;
        // Detail line: NSG, service endpoints, delegations
        const nsgName = s.NSG || s.nsg || "None";
        const svcEndpoints = s.ServiceEndpoints || s.serviceEndpoints || "";
        const delegations = s.Delegations || s.delegations || "";
        let detail = `NSG: ${nsgName === "None" ? "\u2718 None" : nsgName}`;
        if (svcEndpoints && svcEndpoints !== "None") detail += ` | SE: ${svcEndpoints}`;
        if (delegations && delegations !== "None") detail += ` | Del: ${delegations}`;
        const nsgColor = nsgName === "None" ? "#C00000" : $(col.grey);
        sB.push(`<text x="${x + BOX_PADDING + 4}" y="${cy + 7}" font-family="${FONT}" font-size="6" ` +
          `fill="${nsgColor}">${esc(trunc(detail, detailMaxChars))}</text>`);
        cy += SUBNET_DETAIL_HEIGHT;
      }
      if (extra > 0) {
        sB.push(`<text x="${x + BOX_PADDING}" y="${cy + 9}" font-family="${FONT}" font-size="7.5" ` +
          `font-style="italic" fill="${$(col.grey)}">+${extra} more subnet(s)</text>`);
        cy += SUBNET_LINE_HEIGHT;
      }
      // Appliance badges
      if (appls.length > 0) {
        cy += 4; let bxx = x + BOX_PADDING;
        for (const ap of appls) {
          const ac = AC[ap] || { bg: $(col.grey), fg: $(col.white) };
          const bw = ap.length * 5.5 + 12;
          sB.push(`<rect x="${bxx}" y="${cy}" width="${bw}" height="14" rx="3" fill="${ac.bg}"/>` +
            `<text x="${bxx + bw / 2}" y="${cy + 10}" text-anchor="middle" font-family="${FONT}" ` +
            `font-size="7" font-weight="bold" fill="${ac.fg}">${esc(ap)}</text>`);
          bxx += bw + 4;
        }
      }
    }

    // Legend -- line types
    const legY = canvasH - LEGEND_HEIGHT + 5, legP = [];
    // Legend background
    legP.push(`<rect x="0" y="${legY - 5}" width="${canvasW}" height="${LEGEND_HEIGHT}" fill="${$(col.white)}"/>` +
      `<line x1="${TOP_PADDING}" y1="${legY - 4}" x2="${canvasW - TOP_PADDING}" y2="${legY - 4}" stroke="${$(col.accent)}" stroke-width="0.5"/>`);
    const legItems = [
      { color: $(col.primary), dash: false, label: "VNet Peering" },
      { color: "#E57373", dash: true, label: "Disconnected" },
      { color: "#FFB74D", dash: true, label: "VPN / S2S" },
      { color: "#BA68C8", dash: false, label: "ExpressRoute" },
    ];
    let lx = TOP_PADDING;
    for (const it of legItems) {
      const d = it.dash ? ` stroke-dasharray="4,2"` : "";
      legP.push(`<line x1="${lx}" y1="${legY + 6}" x2="${lx + 20}" y2="${legY + 6}" stroke="${it.color}" stroke-width="2"${d}/>` +
        `<text x="${lx + 24}" y="${legY + 9}" font-family="${FONT}" font-size="8" fill="${$(col.grey)}">${esc(it.label)}</text>`);
      lx += 24 + it.label.length * 5 + 16;
    }
    // NSG status indicator
    legP.push(`<text x="${lx}" y="${legY + 9}" font-family="${FONT}" font-size="8" fill="#C00000">\u2718 No NSG</text>`);
    lx += 60;
    // Legend -- appliance badges
    const aLeg = [{ label: "Bastion", bg: "#64B5F6" }, { label: "Firewall", bg: "#E57373" }, { label: "VPN GW", bg: "#FFB74D" }, { label: "AppGW", bg: "#BA68C8" }, { label: "LB", bg: "#4DB6AC" }];
    const alY = legY + 20; let alx = TOP_PADDING;
    for (const it of aLeg) {
      legP.push(`<rect x="${alx}" y="${alY}" width="10" height="10" rx="2" fill="${it.bg}"/>` +
        `<text x="${alx + 14}" y="${alY + 8}" font-family="${FONT}" font-size="8" fill="${$(col.grey)}">${esc(it.label)}</text>`);
      alx += 14 + it.label.length * 5 + 14;
    }
    // Subnet detail key
    legP.push(`<text x="${alx + 10}" y="${alY + 8}" font-family="${FONT}" font-size="7" fill="${$(col.grey)}">SE=Service Endpoints  Del=Delegations</text>`);

    const svg = [
      `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${canvasW} ${canvasH}" width="${canvasW}" height="${canvasH}">`,
      `<rect width="100%" height="100%" fill="${$(col.white)}"/>`,
      `<text x="${canvasW / 2}" y="${TOP_PADDING + 14}" text-anchor="middle" font-family="${FONT}" ` +
        `font-size="14" font-weight="bold" fill="${$(col.primary)}">Network Topology</text>`,
      ...sL, ...sB, ...legP, `</svg>`,
    ].join("\n");

    // Step 5 --- Convert to PNG (if sharp available) or return SVG -------------
    const svgString = svg;
    if (!sharp) {
      return { pngBuffer: Buffer.from(svg), svgString, widthPx: canvasW, heightPx: canvasH, isSvg: true };
    }
    const dpiScale = SVG_DPI / 72;
    const hiW = Math.round(canvasW * dpiScale);
    const hiH = Math.round(canvasH * dpiScale);
    const pngBuffer = await sharp(Buffer.from(svg), { density: SVG_DPI })
      .resize({ width: hiW, height: hiH, fit: "fill" })
      .png().toBuffer();
    return { pngBuffer, svgString, widthPx: canvasW, heightPx: canvasH };
  } catch (err) {
    console.error("generateNetworkDiagram error:", err.message || err);
    return null;
  }
}

module.exports = { generateMgHierarchyDiagram, generateNetworkDiagram };
