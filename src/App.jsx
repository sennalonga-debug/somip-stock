import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  LayoutDashboard, Factory, ArrowDownCircle, ArrowUpCircle, ClipboardList,
  Truck, AlertTriangle, Plus, X, Trash2, Pencil, Fuel, RotateCcw, Check,
  Users, History, Loader2, CheckCircle2, AlertCircle, CloudOff, Thermometer,
  FileBarChart, Download, Printer, TrendingDown, TrendingUp, LogOut, Lock, Mail, Menu, ImagePlus, Palette,
} from "lucide-react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import pptxgen from "pptxgenjs";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell, ReferenceLine, Legend,
} from "recharts";
import { supabase, SUPABASE_CONFIGURED } from "./supabaseClient.js";

/* ------------------------------------------------------------------ */
/* Design tokens                                                       */
/* ------------------------------------------------------------------ */
const C = {
  blue: "#0071BD",
  navy: "#0A1F33",
  navyLight: "#12304C",
  orange: "#F16B16",
  ink: "#16212D",
  sub: "#5B6B7A",
  bg: "#F4F6F8",
  border: "#E2E6EA",
  success: "#1E8A5F",
  danger: "#C63C3C",
  warning: "#D98B12",
};
// Applique les 2 couleurs principales personnalisées (Superviseur, page Réglages) à toute
// l'application — C est un objet muté volontairement, pas remplacé, pour que chaque usage
// existant de C.blue/C.orange dans tout le fichier reflète automatiquement le nouveau thème.
function applyTheme(primary, accent) {
  if (primary && /^#[0-9A-Fa-f]{6}$/.test(primary)) C.blue = primary;
  if (accent && /^#[0-9A-Fa-f]{6}$/.test(accent)) C.orange = accent;
}
// Logo personnalisé (Superviseur, page Réglages) — lu directement par ReportHeader et les
// exports PDF/PowerPoint, sans avoir à faire transiter les réglages dans tous les rapports.
let CURRENT_LOGO_URL = null;
function setCurrentLogoUrl(url) { CURRENT_LOGO_URL = url || null; }
let CURRENT_LOGO_TOTAL_URL = null;
function setCurrentLogoTotalUrl(url) { CURRENT_LOGO_TOTAL_URL = url || null; }

/* ------------------------------------------------------------------ */
/* Seed / reference data                                               */
/* ------------------------------------------------------------------ */
const SITES_SEED = [
  { id: "okouma", code: "OKM", name: "Okouma", capacity: 60000, stockInitial: 38000 },
  { id: "prehomo", code: "PRH", name: "Comilog Prehomo", capacity: 50000, stockInitial: 21000 },
  { id: "lipaka", code: "LPK", name: "AMD Lipaka", capacity: 30000, stockInitial: 9500 },
  { id: "fcv", code: "FCV", name: "Setrag FCV", capacity: 25000, stockInitial: 14200 },
  { id: "benguia", code: "GSB", name: "Setrag GSEZ Benguia", capacity: 25000, stockInitial: 6100 },
  { id: "cim", code: "CIM", name: "CIM", capacity: 20000, stockInitial: 11800 },
  { id: "traction", code: "GTR", name: "Gare Traction", capacity: 15000, stockInitial: 3200 },
  { id: "cmm", code: "CMM", name: "CMM (C2M)", capacity: 20000, stockInitial: 16000 },
];

const TRUCKS = ["FK253AA", "FK254AA", "JJ751AA", "JJ752AA", "JJ763AA", "JL232AA"];
// Lubrifiants gérés sur Prehomo et Okouma uniquement, en plus du gasoil.
const LUBRICANTS = [
  { id: "ac30", label: "AC30", densite: 0.89 },
  { id: "ac50", label: "AC50", densite: 0.90 },
  { id: "sw10", label: "SW10", densite: 0.88 },
  { id: "rubia_tir7400", label: "Rubia Tir7400", densite: 0.883 },
];
const LUBRICANT_SITE_IDS = ["prehomo", "okouma"];
// Certains sites ont plusieurs compteurs physiques distincts pour les sorties, identifiés
// par le code du site (plus fiable que l'id interne). Par défaut : un seul compteur.
const SITE_METERS_BY_CODE = {
  PRH: ["Poste 1", "Poste 2"],
  OKM: ["Poste 1", "Poste 2"],
  LPK: ["Compteur 1", "Compteur 2"],
  CMM: ["Compteur 1", "Compteur 2"],
  GTR: ["Compteur 1", "Compteur 2"],
  FCV: ["Compteur"],
  GSB: ["Compteur"],
  CIM: ["Compteur 1", "Compteur 2", "Compteur Agglo"],
};
function metersForSite(site, dynamicMeters) {
  if (!site) return ["Compteur"];
  const dyn = (dynamicMeters || []).filter((m) => m.siteId === site.id).map((m) => m.name);
  if (dyn.length) return dyn;
  if (site.isMobile) return ["Poste 1", "Poste 2"];
  return SITE_METERS_BY_CODE[site.code] || ["Compteur"];
}
const PRODUCTS = [{ id: "gasoil", label: "Gasoil" }, ...LUBRICANTS];

const MOVEMENTS_SEED = [
  { id: "m1", siteId: "okouma", type: "reception", date: "2026-08-28", quantity: 15000, delta: 15000, ref: "BL-2891", commentaire: "Livraison TotalEnergies", isDemo: true },
  { id: "m2", siteId: "okouma", type: "sortie_camion", date: "2026-08-30", quantity: 4200, delta: -4200, camion: "Camion Laitier 2", destination: "Carrière Nord — Engins", commentaire: "", isDemo: true },
  { id: "m3", siteId: "prehomo", type: "sortie", date: "2026-09-01", quantity: 2600, delta: -2600, destinataire: "Atelier mécanique", commentaire: "", isDemo: true },
  { id: "m4", siteId: "lipaka", type: "reception", date: "2026-09-02", quantity: 8000, delta: 8000, ref: "BL-2903", commentaire: "", isDemo: true },
  { id: "m5", siteId: "traction", type: "sortie", date: "2026-09-02", quantity: 1800, delta: -1800, destinataire: "Locomotive 12", commentaire: "", isDemo: true },
];

const SETTINGS_SEED = { objectifFreinte: 3, logoUrl: null, logoTotalUrl: null, colorPrimary: "#0071BD", colorAccent: "#F16B16" };
const rowToSettings = (r) => r ? {
  objectifFreinte: Number(r.objectif_freinte), logoUrl: r.logo_url || null, logoTotalUrl: r.logo_total_url || null,
  colorPrimary: r.color_primary || "#0071BD", colorAccent: r.color_accent || "#F16B16",
} : SETTINGS_SEED;

const TYPE_META = {
  reception: { label: "Réception", color: C.success, sign: "+" },
  sortie: { label: "Vente", color: C.ink, sign: "" },
  sortie_camion: { label: "Sortie vers camion laitier", color: C.orange, sign: "−" },
  retour_camion: { label: "Retour camion (cuve)", color: C.success, sign: "+" },
  retour_cuve_camion: { label: "Retour cuve (camion)", color: C.danger, sign: "−" },
  ajustement: { label: "Ajustement d'inventaire", color: C.blue, sign: "±" },
};

const todayStr = () => new Date().toISOString().slice(0, 10);

// ---- Saisie hors-connexion : file d'attente locale, synchronisée dès le retour du réseau ----
const OFFLINE_QUEUE_KEY = "somip_offline_queue_v1";
function readOfflineQueue() {
  try { return JSON.parse(localStorage.getItem(OFFLINE_QUEUE_KEY) || "[]"); } catch (e) { return []; }
}
function writeOfflineQueue(q) {
  try { localStorage.setItem(OFFLINE_QUEUE_KEY, JSON.stringify(q)); } catch (e) { /* stockage indisponible : tant pis, pas bloquant */ }
}
function isNetworkError(e) {
  if (!navigator.onLine) return true;
  const msg = String(e?.message || e || "");
  return /fetch|network|Failed to fetch|NetworkError|ERR_INTERNET/i.test(msg);
}
const FRENCH_MONTHS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"];
function formatDateLong(dateStr) {
  const [y, m, d] = dateStr.split("-").map(Number);
  return `${d} ${FRENCH_MONTHS[m - 1]} ${y}`;
}
function formatDateShort(dateStr) {
  const [y, m, d] = dateStr.split("-").map(Number);
  return `${pad2(d)}/${pad2(m)}/${y}`;
}
// Les camions ne livrent/reprennent que par multiples de 5000 L (5000/15000/20000/35000) :
// toute demande d'approvisionnement est arrondie à l'inférieur, au multiple de 5000 le plus proche.
function roundDown5000(n) {
  return Math.floor(Math.max(0, n) / 5000) * 5000;
}
const currentMonth = () => new Date().toISOString().slice(0, 7);
const fmt = (n) => {
  const r = Math.round(n);
  const neg = r < 0;
  const s = Math.abs(r).toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
  return neg ? `-${s}` : s;
};
const uid = () => Math.random().toString(36).slice(2, 10);

/**
 * Classification d'un écart d'inventaire.
 * Stock théorique = Stock initial + Réceptions − Sorties (déjà porté par stockOf()).
 * Écart = Stock physique − Stock théorique.
 *  - écart < 0  -> PERTE (le physique est inférieur au théorique)
 *  - écart > 0  -> GAIN (le physique est supérieur au théorique)
 *  - écart = 0  -> aucun écart
 * Taux de freinte = |écart| ÷ Stock théorique × 1000, exprimé en ‰, et ne
 * s'applique qu'aux PERTES (le terme "freinte" désigne un manquant, pas un excédent).
 * Objectif = seuil de tolérance (‰) fixé par SOMIP (par défaut 3/1000).
 */
function classifyEcart(ecartL, theorique, objectif) {
  const ecartPermille = theorique !== 0 ? (ecartL / theorique) * 1000 : 0;
  const nature = ecartL === 0 ? "neutre" : ecartL < 0 ? "perte" : "gain";
  const tauxFreinte = nature === "perte" ? Math.abs(ecartPermille) : 0;
  let conformite;
  if (nature === "neutre") conformite = "conforme";
  else if (nature === "perte") conformite = tauxFreinte <= objectif ? "conforme" : "non_conforme";
  else conformite = Math.abs(ecartPermille) <= objectif ? "conforme" : "a_verifier";
  return { ecartL, ecartPermille, nature, tauxFreinte, objectif, conformite };
}

const NATURE_META = {
  perte: { label: "Perte", color: C.danger },
  gain: { label: "Gain", color: C.success },
  neutre: { label: "Aucun écart", color: C.sub },
};

const CONFORMITE_META = {
  conforme: { label: "Conforme à l'objectif", color: C.success },
  non_conforme: { label: "Hors objectif (freinte)", color: C.danger },
  a_verifier: { label: "Excédent à vérifier", color: C.warning },
};

/* ------------------------------------------------------------------ */
/* Correction de volume à 15°C — ASTM D1250 / API MPMS Chapitre 11.1   */
/* Table 53B (densité observée -> densité à 15°C) et Table 54B         */
/* (volume -> volume à 15°C), "produits généralisés", plage 653–1075   */
/* kg/m³. Formule officielle : VCF = e^(-ALPHA·ΔT·(1+0.8·ALPHA·ΔT))    */
/* avec ALPHA = f(densité à 15°C), déterminé par zone de densité.      */
/* ------------------------------------------------------------------ */
const VCF_MIN_DENSITY = 653;
const VCF_MAX_DENSITY = 1075;

function vcfAlpha(den15) {
  if (den15 < 770) return (346.42278 + 0.43884 * den15) / (den15 * den15);
  if (den15 < 778) return -0.0033612 + 2680.32 / (den15 * den15); // zone de transition
  if (den15 < 839) return 594.5418 / (den15 * den15);
  return (186.9696 + 0.48618 * den15) / (den15 * den15);
}
function vcfFactor(den15, tempC) {
  const dT = tempC - 15;
  const alpha = vcfAlpha(den15);
  return { vcf: Math.exp(-alpha * dT * (1 + 0.8 * alpha * dT)), alpha };
}
// La densité à 15°C est l'inconnue de sa propre formule (ALPHA en dépend) :
// on résout par itération, convergence en quelques passes (écart de température modéré).
function densityAt15FromObserved(densiteObservee, tempC) {
  let d15 = densiteObservee;
  for (let i = 0; i < 12; i++) {
    const { vcf } = vcfFactor(d15, tempC);
    d15 = densiteObservee / vcf;
  }
  return d15;
}
function correctVolumeTo15({ volumeAmbiant, tempC, densiteObservee }) {
  const vOk = typeof volumeAmbiant === "number" && volumeAmbiant > 0;
  const tOk = typeof tempC === "number" && !isNaN(tempC);
  const dOk = typeof densiteObservee === "number" && densiteObservee >= VCF_MIN_DENSITY && densiteObservee <= VCF_MAX_DENSITY;
  if (!vOk || !tOk || !dOk) return null;
  const densite15 = densityAt15FromObserved(densiteObservee, tempC);
  const { vcf, alpha } = vcfFactor(densite15, tempC);
  return { densite15, alpha, vcf, volume15: volumeAmbiant * vcf };
}
// Quantité "officielle" d'un mouvement pour la comptabilisation à 15°C :
// le volume corrigé s'il a été renseigné (température + densité), sinon le
// volume ambiant en repli (mouvement sans mesure de température/densité).
function movementQty15(m) {
  return m.volumeCorrige15 !== undefined ? m.volumeCorrige15 : m.quantity;
}

/* ------------------------------------------------------------------ */
/* Rapports — fonctions de calcul par période                          */
/* ------------------------------------------------------------------ */
// Calcule le stock à une date donnée en s'ancrant sur la DERNIÈRE JAUGE MESURÉE connue
// avant cette date (le stock fin d'un jour devient le stock début du jour suivant), et non
// sur un cumul théorique indépendant depuis le stock initial. Si aucune jauge n'a encore été
// mesurée pour ce site/produit, on retombe sur le cumul depuis le stock initial du site.
function anchoredStock(siteId, stockInitialFallback, movements, inventaires, product, cutoffDate, inclusive) {
  const priorInv = pickLatestInv(
    inventaires.filter((i) => i.siteId === siteId && (i.product || "gasoil") === product && (inclusive ? i.date <= cutoffDate : i.date < cutoffDate))
  );
  if (priorInv) {
    const afterMovs = movements.filter((m) => m.siteId === siteId && (m.product || "gasoil") === product && m.date > priorInv.date && (inclusive ? m.date <= cutoffDate : m.date < cutoffDate));
    return afterMovs.reduce((a, m) => a + m.delta, priorInv.stockPhysique);
  }
  const movs = movements.filter((m) => m.siteId === siteId && (m.product || "gasoil") === product && (inclusive ? m.date <= cutoffDate : m.date < cutoffDate));
  return movs.reduce((a, m) => a + m.delta, stockInitialFallback);
}
function stockBeforeDate(site, movements, dateExclusive, inventaires) {
  return anchoredStock(site.id, site.stockInitial, movements, inventaires || [], "gasoil", dateExclusive, false);
}
function stockThroughDate(site, movements, dateInclusive, inventaires) {
  return anchoredStock(site.id, site.stockInitial, movements, inventaires || [], "gasoil", dateInclusive, true);
}
function anchoredStock15(siteId, stockInitialFallback, movements, inventaires, product, cutoffDate, inclusive) {
  const priorInv = pickLatestInv(
    inventaires.filter((i) => i.siteId === siteId && (i.product || "gasoil") === product && i.stockPhysique15 !== undefined && (inclusive ? i.date <= cutoffDate : i.date < cutoffDate))
  );
  if (priorInv) {
    const afterMovs = movements.filter((m) => m.siteId === siteId && (m.product || "gasoil") === product && m.date > priorInv.date && (inclusive ? m.date <= cutoffDate : m.date < cutoffDate));
    return afterMovs.reduce((a, m) => a + Math.sign(m.delta) * movementQty15(m), priorInv.stockPhysique15);
  }
  const movs = movements.filter((m) => m.siteId === siteId && (m.product || "gasoil") === product && (inclusive ? m.date <= cutoffDate : m.date < cutoffDate));
  return movs.reduce((a, m) => a + Math.sign(m.delta) * movementQty15(m), stockInitialFallback);
}
function stockBeforeDate15(site, movements, dateExclusive, inventaires) {
  return anchoredStock15(site.id, site.stockInitial, movements, inventaires || [], "gasoil", dateExclusive, false);
}
function stockBeforeDateProduct(stockInitial, movements, siteId, product, dateExclusive, inventaires) {
  return anchoredStock(siteId, stockInitial, movements, inventaires || [], product, dateExclusive, false);
}
function movementsInRange(movements, siteId, startInclusive, endInclusive, product = "gasoil") {
  return movements.filter((m) => m.siteId === siteId && (m.product || "gasoil") === product && m.date >= startInclusive && m.date <= endInclusive);
}
function sumQty15(list, types) {
  return list.filter((m) => types.includes(m.type)).reduce((a, m) => a + movementQty15(m), 0);
}
// Retourne l'inventaire le plus récent d'une liste (date la plus tardive, puis heure de
// saisie la plus tardive en cas d'égalité) — "la dernière jauge saisie" pour cette journée.
function pickLatestInv(list) {
  return list.reduce((best, cur) => {
    if (!best) return cur;
    if (cur.date !== best.date) return cur.date > best.date ? cur : best;
    return (cur.createdAt || "") > (best.createdAt || "") ? cur : best;
  }, null);
}
function sumQty(list, types) {
  return list.filter((m) => types.includes(m.type)).reduce((a, m) => a + m.quantity, 0);
}
function pad2(n) { return String(n).padStart(2, "0"); }
function decadeBounds(dateStr) {
  const [y, m, d] = dateStr.split("-").map(Number);
  let startDay, endDay, label;
  if (d <= 10) { startDay = 1; endDay = 10; label = "1ère décade"; }
  else if (d <= 20) { startDay = 11; endDay = 20; label = "2e décade"; }
  else { startDay = 21; endDay = new Date(y, m, 0).getDate(); label = "3e décade"; }
  return { start: `${y}-${pad2(m)}-${pad2(startDay)}`, end: `${y}-${pad2(m)}-${pad2(endDay)}`, label, monthLabel: `${pad2(m)}/${y}` };
}
function monthBounds(monthStr) {
  const [y, m] = monthStr.split("-").map(Number);
  const lastDay = new Date(y, m, 0).getDate();
  return { start: `${monthStr}-01`, end: `${monthStr}-${pad2(lastDay)}` };
}

function exportToExcel(filename, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, rows }) => {
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  XLSX.writeFile(wb, filename);
}

// Génère un PDF téléchargeable avec l'en-tête SOMIP (bandeau bleu/orange) et un tableau —
// pour un envoi direct par mail, sans passer par la boîte de dialogue d'impression.
function hexToRgb(hex) {
  const h = hex.replace("#", "");
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}
async function loadImageDataUrl(url) {
  const res = await fetch(url);
  const blob = await res.blob();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}
async function exportToPdf({ filename, title, period, columns, rows, totalsRow, sections, sideBySide = false, centerTitle = false, subtitle, bigPeriod = false }) {
  const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const [pR, pG, pB] = hexToRgb(C.blue), [aR, aG, aB] = hexToRgb(C.orange);
  const marginX = 32;

  // Bandeau bicolore SOMIP.
  doc.setFillColor(pR, pG, pB);
  doc.rect(0, 0, pageWidth * 0.6, 7, "F");
  doc.setFillColor(aR, aG, aB);
  doc.rect(pageWidth * 0.6, 0, pageWidth * 0.4, 7, "F");

  let textX = marginX;
  if (CURRENT_LOGO_URL) {
    try {
      const dataUrl = await loadImageDataUrl(CURRENT_LOGO_URL);
      const fmt = dataUrl.includes("image/png") ? "PNG" : "JPEG";
      doc.addImage(dataUrl, fmt, marginX, 18, 30, 30);
      textX = marginX + 40;
    } catch (e) { /* logo indisponible : on continue sans */ }
  }
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.setTextColor(pR, pG, pB);
  doc.text("SOMIP", textX, 32);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9.5);
  doc.setTextColor(110, 120, 130);
  doc.text("Stock Gasoil · Zone Sud-Est · Gabon", textX, 46);
  doc.setDrawColor(226, 230, 234);
  doc.setLineWidth(0.75);
  doc.line(marginX, 62, pageWidth - marginX, 62);

  const titleX = centerTitle ? pageWidth / 2 : marginX;
  const titleOpts = centerTitle ? { align: "center" } : {};
  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.setTextColor(20, 30, 40);
  doc.text(title, titleX, 86, titleOpts);
  let afterTitleY = 86;
  if (subtitle) {
    doc.setFont("helvetica", "normal");
    doc.setFontSize(11);
    doc.setTextColor(90, 100, 110);
    doc.text(subtitle, titleX, 102, titleOpts);
    afterTitleY = 102;
  }
  if (period) {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(bigPeriod ? 26 : 11);
    doc.setTextColor(aR, aG, aB);
    doc.text(period, titleX, afterTitleY + (bigPeriod ? 30 : 17), titleOpts);
    afterTitleY += bigPeriod ? 30 : 17;
  }

  const tableSections = sections && sections.length ? sections : [{ columns, rows, totalsRow }];
  let startY = afterTitleY + 20;
  const fullWidth = pageWidth - marginX * 2;

  const drawSection = (sec, x, w, y) => {
    let sy = y;
    if (sec.heading) {
      doc.setFillColor(pR, pG, pB);
      doc.roundedRect(x, sy, w, 22, 3, 3, "F");
      doc.setFont("helvetica", "bold");
      doc.setFontSize(10.5);
      doc.setTextColor(255, 255, 255);
      doc.text(sec.heading.toUpperCase(), x + 10, sy + 15);
      sy += 22 + 8;
    }
    autoTable(doc, {
      startY: sy,
      head: [sec.columns],
      body: sec.rows,
      foot: sec.totalsRow ? [sec.totalsRow] : undefined,
      theme: "grid",
      headStyles: { fillColor: [pR, pG, pB], textColor: 255, fontStyle: "bold", fontSize: sideBySide ? 8.5 : 10, halign: "right", cellPadding: sideBySide ? 5 : 7 },
      footStyles: { fillColor: [246, 248, 249], textColor: [20, 30, 40], fontStyle: "bold", fontSize: sideBySide ? 8.5 : 10, lineWidth: { top: 1.2 }, lineColor: [aR, aG, aB] },
      bodyStyles: { fontSize: sideBySide ? 8 : 9.5, cellPadding: sideBySide ? 4.5 : 6.5, textColor: [40, 48, 56] },
      alternateRowStyles: { fillColor: [249, 250, 251] },
      styles: { font: "helvetica", lineColor: [226, 230, 234], lineWidth: 0.5 },
      columnStyles: { 0: { halign: "left", fontStyle: "bold" } },
      margin: { left: x, right: pageWidth - x - w },
      tableWidth: w,
      didParseCell: (data) => {
        if (data.column.index === 0) data.cell.styles.halign = "left";
      },
    });
    return doc.lastAutoTable.finalY;
  };

  // Un "groupe" peut être une section unique, ou plusieurs sous-tableaux empilés dans une
  // même colonne (ex : Lubrifiant — Prehomo puis Lubrifiant — Okouma, l'un sous l'autre).
  const drawGroup = (group, x, w, y) => {
    if (!Array.isArray(group)) return drawSection(group, x, w, y);
    let sy = y;
    group.forEach((sec) => { sy = drawSection(sec, x, w, sy) + 16; });
    return sy - 16;
  };

  if (sideBySide && tableSections.length === 2) {
    // Deux colonnes juxtaposées (au lieu d'empilées) : tient sur une seule page dans la
    // plupart des cas, plus lisible pour comparer Gasoil et Lubrifiants côte à côte.
    const gap = 18;
    const halfWidth = (fullWidth - gap) / 2;
    const leftX = marginX, rightX = marginX + halfWidth + gap;
    const yLeft = drawGroup(tableSections[0], leftX, halfWidth, startY);
    const yRight = drawGroup(tableSections[1], rightX, halfWidth, startY);
    startY = Math.max(yLeft, yRight) + 26;
  } else {
    tableSections.forEach((sec) => {
      startY = drawGroup(sec, marginX, fullWidth, startY) + 26;
    });
  }

  // Pied de page : date d'édition discrète + numéro de page (pas dans l'en-tête, pour ne pas
  // surcharger le titre — l'utilisateur ne veut pas de date/heure en haut sur certains rapports).
  const pageCount = doc.internal.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(150, 158, 165);
    doc.text(`Édité le ${new Date().toLocaleDateString("fr-FR")} à ${new Date().toLocaleTimeString("fr-FR")}`, marginX, pageHeight - 18);
    doc.text(`Page ${i} / ${pageCount}`, pageWidth - marginX, pageHeight - 18, { align: "right" });
  }

  doc.save(filename);
}

// Procès-verbal d'inventaire officiel (inopiné/mensuel), cuve par cuve, avec 3 cases de
// signature (SOMIP / Opérateur / TotalEnergies Marketing) — document portrait, à imprimer et signer.
// Exposition — reproduction fidèle du modèle Excel fourni : bandeau orange, logo, en-tête
// EXPOSITION AU [date] - ZONE SUD-EST / DECADE N°[n] - [mois année], tableau Gasoil (8 sites,
// dans l'ordre fourni) à gauche, tableau Lubrifiant vrac (4 produits × Prehomo/Okouma) à droite.
// Exposition — reproduction fidèle du modèle Excel fourni : UNE SEULE grille A:L, dessinée
// cellule par cellule (avec fusions), pas deux tableaux séparés. Bandeau orange, logo, en-tête
// EXPOSITION AU [date] - ZONE SUD-EST / DECADE N°[n] - [mois année]. Colonnes A-D : Gasoil
// (8 sites, dans l'ordre fourni, + TOTAL). Colonnes E-L : Lubrifiant vrac (4 produits, une
// ligne Prehomo puis une ligne Okouma), à l'intérieur de la même grille.
async function exportExpositionModelPdf({ dateStr, decadeNum, monthLabel, gasoilRows, totalVentes, totalStock, totalDemande, productOrder, lubFor, filename }) {
  const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 28;
  const [pR, pG, pB] = hexToRgb(C.blue);
  const [aR, aG, aB] = hexToRgb(C.orange);
  const [gR, gG, gB] = hexToRgb("#2E9B5C");
  const [brR, brG, brB] = hexToRgb("#8B5E34");
  const GREY_HEAD = [60, 68, 76];
  const GREY_SUB = [240, 242, 244];
  const BORDER = [190, 197, 204];

  // Ligne orange pleine largeur (rangée 1 du modèle).
  doc.setFillColor(aR, aG, aB);
  doc.rect(0, 0, pageWidth, 10, "F");

  let y = 30;
  if (CURRENT_LOGO_URL) {
    try {
      const dataUrl = await loadImageDataUrl(CURRENT_LOGO_URL);
      const fmtImg = dataUrl.includes("image/png") ? "PNG" : "JPEG";
      doc.addImage(dataUrl, fmtImg, marginX, y, 40, 40);
    } catch (e) { /* logo indisponible : on continue sans */ }
  }
  doc.setFont("helvetica", "bold");
  doc.setFontSize(14);
  doc.setTextColor(pR, pG, pB);
  doc.text(`EXPOSITION AU ${formatDateShort(dateStr)} - ZONE SUD-EST`, marginX + 125, y + 17);
  doc.setFontSize(12.5);
  doc.setTextColor(aR, aG, aB);
  doc.text(`DECADE N°${decadeNum} - ${monthLabel}`, pageWidth - marginX, y + 17, { align: "right" });
  y += 56;

  // ---- Grille A:L, une seule table, dessinée cellule par cellule ----
  // 12 colonnes : A (site, large) + B,C,D (gasoil) + E..L (8 colonnes lubrifiant : 4 produits × Stock/Ventes).
  const colW = [118, 78, 82, 84]; // A, B, C, D
  const fullWidth = pageWidth - marginX * 2;
  const lubWidth = fullWidth - colW.reduce((a, w) => a + w, 0);
  const lubColW = lubWidth / 8;
  for (let i = 0; i < 8; i++) colW.push(lubColW);
  const colX = [marginX];
  colW.forEach((w) => colX.push(colX[colX.length - 1] + w));
  const totalW = colX[colX.length - 1] - marginX;

  const PRODUCT_LABELS = { rubia_tir7400: "TIR-7400", ac50: "AC50", ac30: "AC30", sw10: "SW10" };
  const colColors = [[pR, pG, pB], [aR, aG, aB], [gR, gG, gB], [brR, brG, brB]];

  const cell = (c1, c2, rowY, rowH, opts = {}) => {
    const x = colX[c1], w = colX[c2 + 1] - colX[c1];
    if (opts.fill) { doc.setFillColor(...opts.fill); doc.rect(x, rowY, w, rowH, "F"); }
    doc.setDrawColor(...BORDER);
    doc.setLineWidth(0.6);
    doc.rect(x, rowY, w, rowH, "S");
    if (opts.text !== undefined) {
      doc.setFont("helvetica", opts.bold ? "bold" : "normal");
      doc.setFontSize(opts.fontSize || 8.5);
      doc.setTextColor(...(opts.color || [30, 38, 46]));
      const align = opts.align || "center";
      const tx = align === "left" ? x + 6 : align === "right" ? x + w - 6 : x + w / 2;
      doc.text(String(opts.text), tx, rowY + rowH / 2 + 3, { align });
    }
  };

  const rowHHeader = 20, rowHData = 19;
  let ry = y;

  // Rangée "LUBRIFIANT VRAC" (E:L fusionné) — rien à afficher sur A:D à cette hauteur.
  cell(0, 3, ry, rowHHeader, { fill: [255, 255, 255] });
  cell(4, 11, ry, rowHHeader, { fill: [pR, pG, pB], text: "LUBRIFIANT VRAC", bold: true, color: [255, 255, 255], fontSize: 10 });
  ry += rowHHeader;

  // Rangée des 4 produits (2 colonnes fusionnées chacun).
  cell(0, 3, ry, rowHHeader, { fill: [255, 255, 255] });
  productOrder.forEach((p, i) => {
    cell(4 + i * 2, 4 + i * 2 + 1, ry, rowHHeader, { fill: colColors[i], text: PRODUCT_LABELS[p], bold: true, color: [255, 255, 255], fontSize: 9 });
  });
  ry += rowHHeader;

  // Rangée d'en-têtes : SITES/CONSIGNATION/DEMANDE D'APPRO/VENTES DECADEn + STOCK/VENTES ×4.
  cell(0, 0, ry, rowHHeader, { fill: GREY_HEAD, text: "SITES", bold: true, color: [255, 255, 255], align: "left", fontSize: 8.5 });
  cell(1, 1, ry, rowHHeader, { fill: GREY_HEAD, text: "CONSIGNATION", bold: true, color: [255, 255, 255], fontSize: 7.5 });
  cell(2, 2, ry, rowHHeader, { fill: GREY_HEAD, text: "DEMANDE D'APPRO", bold: true, color: [255, 255, 255], fontSize: 7 });
  cell(3, 3, ry, rowHHeader, { fill: GREY_HEAD, text: `VENTES DECADE${decadeNum}`, bold: true, color: [255, 255, 255], fontSize: 7 });
  for (let i = 0; i < 8; i++) {
    cell(4 + i, 4 + i, ry, rowHHeader, { fill: GREY_SUB, text: i % 2 === 0 ? "STOCK" : "VENTES", bold: true, fontSize: 7.5 });
  }
  ry += rowHHeader;

  // Lignes de données : 8 sites Gasoil (dans l'ordre fourni). La ligne Lubrifiant Prehomo
  // s'aligne sur la ligne où "PREHOMO" apparaît côté Gasoil, et la ligne Lubrifiant Okouma sur
  // celle où "OKOUMA" apparaît — pas juste les 2 premières lignes de la liste.
  const lubDataFor = (siteId) => productOrder.flatMap((p) => {
    const r = lubFor(siteId, p);
    return [r ? fmt(r.stockConsignation) : "—", r ? fmt(r.ventes) : "—"];
  });
  const prehomoRowIdx = gasoilRows.findIndex((r) => r.label === "PREHOMO");
  const okoumaRowIdx = gasoilRows.findIndex((r) => r.label === "OKOUMA");
  const lubRowsByIdx = {};
  if (prehomoRowIdx >= 0) lubRowsByIdx[prehomoRowIdx] = lubDataFor("prehomo");
  if (okoumaRowIdx >= 0) lubRowsByIdx[okoumaRowIdx] = lubDataFor("okouma");

  gasoilRows.forEach((r, idx) => {
    cell(0, 0, ry, rowHData, { text: r.label, bold: true, align: "left", fontSize: 8.5 });
    cell(1, 1, ry, rowHData, { text: fmt(r.stockConsignation), fontSize: 8.5, align: "right" });
    cell(2, 2, ry, rowHData, { text: fmt(r.demandeAppro), fontSize: 8.5, align: "right" });
    cell(3, 3, ry, rowHData, { text: fmt(r.ventesCumulees), fontSize: 8.5, align: "right" });
    if (lubRowsByIdx[idx]) {
      lubRowsByIdx[idx].forEach((val, i) => cell(4 + i, 4 + i, ry, rowHData, { text: val, fontSize: 8, align: "right" }));
    } else {
      cell(4, 11, ry, rowHData, { fill: [255, 255, 255] });
    }
    ry += rowHData;
  });

  // Ligne TOTAL — bandeau bleu.
  cell(0, 0, ry, rowHData, { fill: [pR, pG, pB], text: "TOTAL", bold: true, color: [255, 255, 255], align: "left", fontSize: 8.5 });
  cell(1, 1, ry, rowHData, { fill: [pR, pG, pB], text: fmt(totalStock), bold: true, color: [255, 255, 255], fontSize: 8.5, align: "right" });
  cell(2, 2, ry, rowHData, { fill: [pR, pG, pB], text: fmt(totalDemande), bold: true, color: [255, 255, 255], fontSize: 8.5, align: "right" });
  cell(3, 3, ry, rowHData, { fill: [pR, pG, pB], text: fmt(totalVentes), bold: true, color: [255, 255, 255], fontSize: 8.5, align: "right" });
  cell(4, 11, ry, rowHData, { fill: [255, 255, 255] });
  ry += rowHData;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.setTextColor(150, 158, 165);
  doc.text(`Édité le ${new Date().toLocaleDateString("fr-FR")} à ${new Date().toLocaleTimeString("fr-FR")}`, marginX, pageHeight - 16);

  doc.save(filename);
}

async function exportInventaireOfficielToPdf(inv, site) {
  const doc = new jsPDF({ orientation: "portrait", unit: "pt", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const [pR, pG, pB] = hexToRgb(C.blue), [aR, aG, aB] = hexToRgb(C.orange);
  const marginX = 40;
  const fullWidth = pageWidth - marginX * 2;

  // Bandeau bicolore SOMIP.
  doc.setFillColor(pR, pG, pB);
  doc.rect(0, 0, pageWidth * 0.6, 7, "F");
  doc.setFillColor(aR, aG, aB);
  doc.rect(pageWidth * 0.6, 0, pageWidth * 0.4, 7, "F");

  let textX = marginX;
  if (CURRENT_LOGO_URL) {
    try {
      const dataUrl = await loadImageDataUrl(CURRENT_LOGO_URL);
      const fmtImg = dataUrl.includes("image/png") ? "PNG" : "JPEG";
      doc.addImage(dataUrl, fmtImg, marginX, 20, 32, 32);
      textX = marginX + 42;
    } catch (e) { /* logo indisponible : on continue sans */ }
  }
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.setTextColor(pR, pG, pB);
  doc.text("SOMIP", textX, 34);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9.5);
  doc.setTextColor(110, 120, 130);
  doc.text("Sites externalisés — Zone Sud-Est · Gabon", textX, 48);

  // Logo TotalEnergies, en haut à droite (document conjoint SOMIP / TotalEnergies).
  if (CURRENT_LOGO_TOTAL_URL) {
    try {
      const dataUrlT = await loadImageDataUrl(CURRENT_LOGO_TOTAL_URL);
      const fmtT = dataUrlT.includes("image/png") ? "PNG" : "JPEG";
      doc.addImage(dataUrlT, fmtT, pageWidth - marginX - 60, 18, 60, 34);
    } catch (e) { /* logo indisponible : on continue sans */ }
  }

  doc.setDrawColor(226, 230, 234);
  doc.setLineWidth(0.75);
  doc.line(marginX, 66, pageWidth - marginX, 66);

  // Titre dynamique : "Inventaire Inopiné" ou "Inventaire fin <mois>" (mensuel).
  const titre = inv.type === "mensuel"
    ? `Inventaire fin ${FRENCH_MONTHS[Number(inv.date.slice(5, 7)) - 1]}`
    : "Inventaire Inopiné";
  doc.setFont("helvetica", "bold");
  doc.setFontSize(17);
  doc.setTextColor(20, 30, 40);
  doc.text(titre, pageWidth / 2, 92, { align: "center" });
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11.5);
  doc.setTextColor(aR, aG, aB);
  doc.text(`${site?.name || inv.siteId} — ${PRODUIT_INVENTAIRE_LABELS[inv.produit] || "Gasoil"} — ${formatDateLong(inv.date)}`, pageWidth / 2, 110, { align: "center" });

  let y = 138;
  const infoRow = (label1, value1, label2, value2) => {
    const colW = fullWidth / 2;
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(110, 120, 130);
    doc.text(label1.toUpperCase(), marginX, y);
    doc.text(label2.toUpperCase(), marginX + colW, y);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11.5);
    doc.setTextColor(20, 30, 40);
    doc.text(value1 || "—", marginX, y + 15);
    doc.text(value2 || "—", marginX + colW, y + 15);
    y += 34;
  };
  infoRow("Inventoriste", inv.inventoriste, "Opérateur", inv.operateur);
  doc.setDrawColor(226, 230, 234);
  doc.line(marginX, y - 14, pageWidth - marginX, y - 14);
  y += 6;

  const drawBanner = (text) => {
    doc.setFillColor(pR, pG, pB);
    doc.roundedRect(marginX, y, fullWidth, 22, 3, 3, "F");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10.5);
    doc.setTextColor(255, 255, 255);
    doc.text(text, marginX + 10, y + 15);
    y += 22 + 8;
  };

  // Tableau 1 — Cuves : hauteur, densité, température, présence d'eau, volumes ambiant et 15°C.
  drawBanner("CUVES — RELEVÉ DÉTAILLÉ");
  autoTable(doc, {
    startY: y,
    head: [["Cuve", "Hauteur", "Densité", "Temp. (°C)", "Eau", "Vol. ambiant (L)", "Vol. 15°C (L)"]],
    body: (inv.cuves || []).map((c) => [
      c.cuve,
      c.hauteur !== null && c.hauteur !== undefined ? fmt(c.hauteur) : "—",
      c.densite !== null && c.densite !== undefined ? String(c.densite) : "—",
      c.temperatureC !== null && c.temperatureC !== undefined ? String(c.temperatureC) : "—",
      c.eau ? "Oui" : "Non",
      `${fmt(c.stockAmbiant)} L`,
      c.volume15 !== null && c.volume15 !== undefined ? `${fmt(c.volume15)} L` : "—",
    ]),
    theme: "grid",
    headStyles: { fillColor: [pR, pG, pB], textColor: 255, fontStyle: "bold", fontSize: 8.5, cellPadding: 6 },
    bodyStyles: { fontSize: 9, cellPadding: 6, textColor: [40, 48, 56] },
    alternateRowStyles: { fillColor: [249, 250, 251] },
    styles: { font: "helvetica", lineColor: [226, 230, 234], lineWidth: 0.5, halign: "right" },
    columnStyles: { 0: { halign: "left", fontStyle: "bold" }, 4: { halign: "center" } },
    margin: { left: marginX, right: marginX },
    didParseCell: (data) => {
      if (data.column.index === 4 && data.section === "body" && data.cell.raw === "Oui") {
        data.cell.styles.textColor = [aR, aG, aB];
        data.cell.styles.fontStyle = "bold";
      }
    },
  });
  y = doc.lastAutoTable.finalY + 20;

  // Tableau 2 — Index, indépendant des cuves (les compteurs de vente habituels du site).
  if (inv.indexCompteurs && inv.indexCompteurs.length > 0) {
    drawBanner("INDEX — RELEVÉ COMPTEUR");
    autoTable(doc, {
      startY: y,
      head: [["Compteur", "Index fin"]],
      body: inv.indexCompteurs.map((d) => [d.compteur, d.indexFin !== null && d.indexFin !== undefined ? fmt(d.indexFin) : "—"]),
      theme: "grid",
      headStyles: { fillColor: [pR, pG, pB], textColor: 255, fontStyle: "bold", fontSize: 10, cellPadding: 7 },
      bodyStyles: { fontSize: 10, cellPadding: 7, textColor: [40, 48, 56] },
      alternateRowStyles: { fillColor: [249, 250, 251] },
      styles: { font: "helvetica", lineColor: [226, 230, 234], lineWidth: 0.5, halign: "right" },
      columnStyles: { 0: { halign: "left", fontStyle: "bold" } },
      margin: { left: marginX, right: marginX },
      tableWidth: fullWidth / 2,
    });
    y = doc.lastAutoTable.finalY + 20;
  }

  // Tableau 3 — Index des compteurs de dépotage (livraison), séparé lui aussi.
  if (inv.depotage && inv.depotage.length > 0) {
    drawBanner("INDEX — COMPTEURS DE DÉPOTAGE");
    autoTable(doc, {
      startY: y,
      head: [["Compteur de dépotage", "Index fin"]],
      body: inv.depotage.map((d) => [d.compteur, d.indexFin !== null && d.indexFin !== undefined ? fmt(d.indexFin) : "—"]),
      theme: "grid",
      headStyles: { fillColor: [pR, pG, pB], textColor: 255, fontStyle: "bold", fontSize: 10, cellPadding: 7 },
      bodyStyles: { fontSize: 10, cellPadding: 7, textColor: [40, 48, 56] },
      alternateRowStyles: { fillColor: [249, 250, 251] },
      styles: { font: "helvetica", lineColor: [226, 230, 234], lineWidth: 0.5, halign: "right" },
      columnStyles: { 0: { halign: "left", fontStyle: "bold" } },
      margin: { left: marginX, right: marginX },
      tableWidth: fullWidth / 2,
    });
    y = doc.lastAutoTable.finalY + 20;
  }

  // Totaux : stock ambiant et à 15°C, côte à côte.
  const halfW = (fullWidth - 16) / 2;
  const boxH = 46;
  doc.setFillColor(246, 248, 249);
  doc.roundedRect(marginX, y, halfW, boxH, 4, 4, "F");
  doc.roundedRect(marginX + halfW + 16, y, halfW, boxH, 4, 4, "F");
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(110, 120, 130);
  doc.text("STOCK TOTAL — BASE AMBIANTE", marginX + 12, y + 17);
  doc.text("STOCK TOTAL — BASE 15°C", marginX + halfW + 28, y + 17);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.setTextColor(pR, pG, pB);
  doc.text(`${fmt(inv.stockAmbiant)} L`, marginX + 12, y + 36);
  doc.setTextColor(inv.stock15 !== undefined ? aR : 170, inv.stock15 !== undefined ? aG : 175, inv.stock15 !== undefined ? aB : 180);
  doc.text(inv.stock15 !== undefined ? `${fmt(inv.stock15)} L` : "Non calculé", marginX + halfW + 28, y + 36);
  y += boxH + 24;

  if (inv.commentaire) {
    doc.setFont("helvetica", "italic");
    doc.setFontSize(9.5);
    doc.setTextColor(90, 100, 110);
    const lines = doc.splitTextToSize(`Commentaire : ${inv.commentaire}`, fullWidth);
    doc.text(lines, marginX, y);
    y += lines.length * 12 + 16;
  }

  // Cases de signature, en bas de page (fixe, pas juste après le contenu, pour un rendu
  // toujours propre même si le tableau des cuves est court ou long). Si une signature
  // numérique a déjà été apposée, son image remplace la case vide.
  const sigY = Math.max(y + 20, pageHeight - 150);
  const sigW = (fullWidth - 32) / 3;
  const sigSlots = [
    { label: "SOMIP", url: inv.signatureSomipUrl, by: inv.signatureSomipBy, at: inv.signatureSomipAt },
    { label: "Opérateur", url: inv.signatureOperateurUrl, by: inv.signatureOperateurBy, at: inv.signatureOperateurAt },
    { label: "TotalEnergies Marketing", url: inv.signatureTotalUrl, by: inv.signatureTotalBy, at: inv.signatureTotalAt },
  ];
  for (let i = 0; i < sigSlots.length; i++) {
    const slot = sigSlots[i];
    const x = marginX + i * (sigW + 16);
    doc.setDrawColor(180, 188, 195);
    doc.setLineWidth(0.75);
    doc.rect(x, sigY, sigW, 90);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9.5);
    doc.setTextColor(20, 30, 40);
    doc.text(slot.label.toUpperCase(), x + sigW / 2, sigY + 14, { align: "center" });
    if (slot.url) {
      try {
        const sigDataUrl = await loadImageDataUrl(slot.url);
        const sigFmt = sigDataUrl.includes("image/png") ? "PNG" : "JPEG";
        doc.addImage(sigDataUrl, sigFmt, x + 10, sigY + 20, sigW - 20, 45, undefined, "FAST");
      } catch (e) { /* signature indisponible : on continue sans */ }
      doc.setFont("helvetica", "normal");
      doc.setFontSize(7);
      doc.setTextColor(110, 120, 130);
      const dateLabel = slot.at ? new Date(slot.at).toLocaleDateString("fr-FR") : "";
      doc.text(`${slot.by || ""}${dateLabel ? " — " + dateLabel : ""}`, x + sigW / 2, sigY + 82, { align: "center" });
    } else {
      doc.setFont("helvetica", "normal");
      doc.setFontSize(8);
      doc.setTextColor(150, 158, 165);
      doc.text("Signature", x + sigW / 2, sigY + 82, { align: "center" });
    }
  }

  doc.save(`SOMIP_${titre.replace(/\s+/g, "_")}_${site?.code || inv.siteId}_${inv.date}.pdf`);
}

// Envoie une ou plusieurs photos vers Supabase Storage (bucket "somip-photos") et renvoie
// leurs URLs publiques. Utilisé pour justifier une perte (Stock fin) ou illustrer un Bilan Matières.
async function uploadPhotos(files, folder) {
  const urls = [];
  for (const file of files) {
    const ext = (file.name.split(".").pop() || "jpg").toLowerCase().replace(/[^a-z0-9]/g, "") || "jpg";
    const path = `${folder}/${Date.now()}-${Math.random().toString(36).slice(2, 8)}.${ext}`;
    const { error } = await supabase.storage.from("somip-photos").upload(path, file, { cacheControl: "3600", upsert: false });
    if (error) throw new Error(`Échec de l'envoi de la photo "${file.name}" : ${error.message}`);
    const { data } = supabase.storage.from("somip-photos").getPublicUrl(path);
    urls.push(data.publicUrl);
  }
  return urls;
}
function dataUrlToFile(dataUrl, filename) {
  const [header, base64] = dataUrl.split(",");
  const mime = /data:(.*?);base64/.exec(header)?.[1] || "image/png";
  const bin = atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new File([bytes], filename, { type: mime });
}

// Pavé de signature tactile/souris — dessin libre, effacer, valider (renvoie un data URL PNG).
function SignaturePad({ onSave, onCancel }) {
  const canvasRef = useRef(null);
  const drawingRef = useRef(false);
  const hasDrawnRef = useRef(false);
  const lastPos = useRef({ x: 0, y: 0 });

  const getPos = (e) => {
    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const clientY = e.touches ? e.touches[0].clientY : e.clientY;
    return { x: (clientX - rect.left) * (canvas.width / rect.width), y: (clientY - rect.top) * (canvas.height / rect.height) };
  };
  const start = (e) => {
    e.preventDefault();
    drawingRef.current = true;
    lastPos.current = getPos(e);
  };
  const move = (e) => {
    if (!drawingRef.current) return;
    e.preventDefault();
    const ctx = canvasRef.current.getContext("2d");
    const pos = getPos(e);
    ctx.strokeStyle = "#1A2733";
    ctx.lineWidth = 2.5;
    ctx.lineCap = "round";
    ctx.beginPath();
    ctx.moveTo(lastPos.current.x, lastPos.current.y);
    ctx.lineTo(pos.x, pos.y);
    ctx.stroke();
    lastPos.current = pos;
    hasDrawnRef.current = true;
  };
  const end = () => { drawingRef.current = false; };
  const clear = () => {
    const canvas = canvasRef.current;
    canvas.getContext("2d").clearRect(0, 0, canvas.width, canvas.height);
    hasDrawnRef.current = false;
  };
  const save = () => {
    if (!hasDrawnRef.current) return;
    onSave(canvasRef.current.toDataURL("image/png"));
  };

  return (
    <div>
      <canvas
        ref={canvasRef} width={500} height={200}
        style={{ width: "100%", maxWidth: 500, height: 160, background: "#fff", border: `1px solid ${C.border}`, borderRadius: 8, touchAction: "none", cursor: "crosshair" }}
        onMouseDown={start} onMouseMove={move} onMouseUp={end} onMouseLeave={end}
        onTouchStart={start} onTouchMove={move} onTouchEnd={end}
      />
      <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
        <button className="somip-btn somip-btn-primary" onClick={save}><Check size={14} /> Valider la signature</button>
        <button className="somip-btn somip-btn-secondary" onClick={clear}><RotateCcw size={14} /> Effacer</button>
        {onCancel && <button onClick={onCancel} style={{ border: "none", background: "none", cursor: "pointer", padding: "6px 10px" }}><X size={16} color={C.sub} /></button>}
      </div>
    </div>
  );
}

// Petit sélecteur de photos réutilisable : aperçus en miniature + bouton de suppression avant envoi.
function PhotoPicker({ files, setFiles, existingUrls, onRemoveExisting }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <label className="somip-btn somip-btn-secondary" style={{ fontSize: 12, padding: "6px 12px", cursor: "pointer", display: "inline-flex" }}>
        <ImagePlus size={14} /> Ajouter des photos
        <input type="file" accept="image/*" multiple style={{ display: "none" }} onChange={(e) => setFiles((prev) => [...prev, ...Array.from(e.target.files || [])])} />
      </label>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 8 }}>
        {(existingUrls || []).map((url, i) => (
          <div key={`existing-${i}`} style={{ position: "relative" }}>
            <img src={url} alt="" style={{ width: 64, height: 64, objectFit: "cover", borderRadius: 6, border: `1px solid ${C.border}` }} />
            {onRemoveExisting && (
              <button onClick={() => onRemoveExisting(i)} style={{ position: "absolute", top: -6, right: -6, background: C.danger, border: "none", borderRadius: "50%", width: 18, height: 18, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <X size={11} color="#fff" />
              </button>
            )}
          </div>
        ))}
        {files.map((f, i) => (
          <div key={`new-${i}`} style={{ position: "relative" }}>
            <img src={URL.createObjectURL(f)} alt="" style={{ width: 64, height: 64, objectFit: "cover", borderRadius: 6, border: `1px solid ${C.blue}` }} />
            <button onClick={() => setFiles((prev) => prev.filter((_, idx) => idx !== i))} style={{ position: "absolute", top: -6, right: -6, background: C.danger, border: "none", borderRadius: "50%", width: 18, height: 18, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
              <X size={11} color="#fff" />
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}

// Export PowerPoint du Bilan Matières : titre, tableau, diagramme natif (modifiable dans
// PowerPoint), et une diapositive photos pour chaque période qui en possède.
async function exportBilanToPptx(history, periodType, siteName) {
  const pptx = new pptxgen();
  pptx.defineLayout({ name: "SOMIP", width: 10, height: 5.63 });
  pptx.layout = "SOMIP";
  const BLUE = C.blue.replace("#", ""), ORANGE = C.orange.replace("#", ""), INK = "1A2733", SUB = "5A6470";
  const PERIOD_LABEL = { mensuel: "Synthèse mensuelle", trimestriel: "Synthèse trimestrielle", decadaire: "Synthèse décadaire" };

  // Diapositive de titre.
  const s1 = pptx.addSlide();
  s1.background = { color: "FFFFFF" };
  s1.addShape("rect", { x: 0, y: 0, w: 6, h: 0.12, fill: { color: BLUE } });
  s1.addShape("rect", { x: 6, y: 0, w: 4, h: 0.12, fill: { color: ORANGE } });
  if (CURRENT_LOGO_URL) {
    try { s1.addImage({ path: CURRENT_LOGO_URL, x: 0.5, y: 0.5, w: 1, h: 1, sizing: { type: "contain", w: 1, h: 1 } }); } catch (e) { /* logo indisponible : on continue sans */ }
  }
  s1.addText(`SOMIP — Bilan Matières${siteName ? ` — ${siteName}` : ""}`, { x: 0.5, y: 2.0, w: 9, h: 0.7, fontSize: 26, bold: true, color: BLUE });
  s1.addText(PERIOD_LABEL[periodType] || "Synthèse", { x: 0.5, y: 2.7, w: 9, h: 0.5, fontSize: 16, color: ORANGE, bold: true });
  s1.addText(`Édité le ${new Date().toLocaleDateString("fr-FR")}`, { x: 0.5, y: 3.2, w: 9, h: 0.4, fontSize: 11, color: SUB });

  // Diapositive tableau.
  const s2 = pptx.addSlide();
  s2.addShape("rect", { x: 0, y: 0, w: 6, h: 0.08, fill: { color: BLUE } });
  s2.addShape("rect", { x: 6, y: 0, w: 4, h: 0.08, fill: { color: ORANGE } });
  s2.addText("Historique des périodes", { x: 0.4, y: 0.25, w: 9, h: 0.4, fontSize: 18, bold: true, color: INK });
  const header = ["Période", "Stock début", "Réception", "Ventes", "Transferts", "Théorique", "Stock fin", "Gain/Perte"].map((t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: BLUE }, fontSize: 9 } }));
  const rows = [header, ...history.map((b) => [
    String(b.periodKey), `${fmt(b.stockDebut)} L`, `+${fmt(b.reception15)} L`, `${fmt(b.ventes15)} L`,
    `${b.transferts15 >= 0 ? "+" : ""}${fmt(b.transferts15)} L`, `${fmt(b.stockTheorique)} L`, `${fmt(b.stockFin15)} L`,
    { text: `${b.ecart >= 0 ? "+" : ""}${fmt(b.ecart)} L`, options: { color: b.ecart < 0 ? "D64545" : b.ecart > 0 ? "2E9B5C" : INK, bold: true } },
  ].map((c) => (typeof c === "string" ? { text: c, options: { fontSize: 9, color: INK } } : c)))];
  s2.addTable(rows, { x: 0.4, y: 0.8, w: 9.2, autoPage: true, border: { type: "solid", color: "E2E6E9", pt: 0.5 }, fontFace: "Arial" });

  // Diapositive diagramme (natif PowerPoint, modifiable).
  if (history.length > 0) {
    const s3 = pptx.addSlide();
    s3.addShape("rect", { x: 0, y: 0, w: 6, h: 0.08, fill: { color: BLUE } });
    s3.addShape("rect", { x: 6, y: 0, w: 4, h: 0.08, fill: { color: ORANGE } });
    s3.addText("Évolution des volumes", { x: 0.4, y: 0.25, w: 9, h: 0.4, fontSize: 18, bold: true, color: INK });
    const labels = history.map((b) => String(b.periodKey));
    const chartData = [
      { name: "Réception", labels, values: history.map((b) => b.reception15) },
      { name: "Ventes", labels, values: history.map((b) => b.ventes15) },
      { name: "Stock fin", labels, values: history.map((b) => b.stockFin15) },
    ];
    s3.addChart(pptx.ChartType.bar, chartData, {
      x: 0.4, y: 0.9, w: 9.2, h: 4.3, barDir: "col",
      chartColors: [BLUE, ORANGE, "2E9B5C"], showLegend: true, legendPos: "b",
      showValue: false, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9,
    });
  }

  // Diapositives photos, pour chaque période qui en possède.
  for (const b of history) {
    if (!b.photoUrls || b.photoUrls.length === 0) continue;
    const s = pptx.addSlide();
    s.addShape("rect", { x: 0, y: 0, w: 6, h: 0.08, fill: { color: BLUE } });
    s.addShape("rect", { x: 6, y: 0, w: 4, h: 0.08, fill: { color: ORANGE } });
    s.addText(`Photos justificatives — ${b.periodKey}`, { x: 0.4, y: 0.25, w: 9, h: 0.4, fontSize: 16, bold: true, color: INK });
    if (b.commentaire) s.addText(b.commentaire, { x: 0.4, y: 0.65, w: 9.2, h: 0.35, fontSize: 10, color: SUB, italic: true });
    const positions = [{ x: 0.4, y: 1.1 }, { x: 3.55, y: 1.1 }, { x: 6.7, y: 1.1 }, { x: 0.4, y: 3.4 }, { x: 3.55, y: 3.4 }, { x: 6.7, y: 3.4 }];
    for (let i = 0; i < Math.min(b.photoUrls.length, 6); i++) {
      try {
        s.addImage({ path: b.photoUrls[i], x: positions[i].x, y: positions[i].y, w: 3, h: 2.1, sizing: { type: "cover", w: 3, h: 2.1 } });
      } catch (e) { /* une image indisponible ne doit pas bloquer tout l'export */ }
    }
  }

  await pptx.writeFile({ fileName: `SOMIP_Bilan_Matieres_${siteName ? siteName.replace(/\s+/g, "") + "_" : ""}${periodType}_${new Date().toISOString().slice(0, 10)}.pptx` });
}

// Export combiné "Tous les sites" pour une période donnée : un seul document (PDF ou
// PowerPoint) rassemblant la ligne de chaque site, plutôt que par site séparément.
async function exportBilanAllSitesToPptx(rows, periodType, periodKey) {
  const pptx = new pptxgen();
  pptx.defineLayout({ name: "SOMIP", width: 10, height: 5.63 });
  pptx.layout = "SOMIP";
  const BLUE = C.blue.replace("#", ""), ORANGE = C.orange.replace("#", ""), INK = "1A2733", SUB = "5A6470";

  const s1 = pptx.addSlide();
  s1.background = { color: "FFFFFF" };
  s1.addShape("rect", { x: 0, y: 0, w: 6, h: 0.12, fill: { color: BLUE } });
  s1.addShape("rect", { x: 6, y: 0, w: 4, h: 0.12, fill: { color: ORANGE } });
  if (CURRENT_LOGO_URL) {
    try { s1.addImage({ path: CURRENT_LOGO_URL, x: 0.5, y: 0.5, w: 1, h: 1, sizing: { type: "contain", w: 1, h: 1 } }); } catch (e) { /* logo indisponible : on continue sans */ }
  }
  s1.addText("SOMIP — Bilan Matières — Tous les sites", { x: 0.5, y: 2.0, w: 9, h: 0.7, fontSize: 24, bold: true, color: BLUE });
  s1.addText(`${PERIOD_TYPE_LABELS[periodType]} — ${periodKey}`, { x: 0.5, y: 2.7, w: 9, h: 0.5, fontSize: 16, color: ORANGE, bold: true });
  s1.addText(`Édité le ${new Date().toLocaleDateString("fr-FR")}`, { x: 0.5, y: 3.2, w: 9, h: 0.4, fontSize: 11, color: SUB });

  const s2 = pptx.addSlide();
  s2.addShape("rect", { x: 0, y: 0, w: 6, h: 0.08, fill: { color: BLUE } });
  s2.addShape("rect", { x: 6, y: 0, w: 4, h: 0.08, fill: { color: ORANGE } });
  s2.addText(`Détail par site — ${periodKey}`, { x: 0.4, y: 0.25, w: 9, h: 0.4, fontSize: 18, bold: true, color: INK });
  const header = ["Site", "Stock début", "Réception", "Ventes", "Transferts", "Théorique", "Stock fin", "Gain/Perte"].map((t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: BLUE }, fontSize: 9 } }));
  const dataRows = rows.map((r) => [
    { text: r.site?.name || "?", options: { fontSize: 9, color: INK } },
    { text: `${fmt(r.stockDebut)} L`, options: { fontSize: 9, color: INK } },
    { text: `+${fmt(r.reception15)} L`, options: { fontSize: 9, color: INK } },
    { text: `${fmt(r.ventes15)} L`, options: { fontSize: 9, color: INK } },
    { text: `${r.transferts15 >= 0 ? "+" : ""}${fmt(r.transferts15)} L`, options: { fontSize: 9, color: INK } },
    { text: `${fmt(r.stockTheorique)} L`, options: { fontSize: 9, color: INK } },
    { text: `${fmt(r.stockFin15)} L`, options: { fontSize: 9, color: INK } },
    { text: `${r.ecart >= 0 ? "+" : ""}${fmt(r.ecart)} L`, options: { fontSize: 9, bold: true, color: r.ecart < 0 ? "D64545" : r.ecart > 0 ? "2E9B5C" : INK } },
  ]);
  const totalStockDebut = rows.reduce((a, r) => a + r.stockDebut, 0), totalReception = rows.reduce((a, r) => a + r.reception15, 0);
  const totalVentes = rows.reduce((a, r) => a + r.ventes15, 0), totalTransferts = rows.reduce((a, r) => a + r.transferts15, 0);
  const totalTheorique = rows.reduce((a, r) => a + r.stockTheorique, 0), totalStockFin = rows.reduce((a, r) => a + r.stockFin15, 0);
  const totalEcart = rows.reduce((a, r) => a + r.ecart, 0);
  const totalRow = [
    { text: "TOTAL", options: { fontSize: 9, bold: true, fill: { color: "F4F6F8" } } },
    { text: `${fmt(totalStockDebut)} L`, options: { fontSize: 9, bold: true, fill: { color: "F4F6F8" } } },
    { text: `+${fmt(totalReception)} L`, options: { fontSize: 9, bold: true, fill: { color: "F4F6F8" } } },
    { text: `${fmt(totalVentes)} L`, options: { fontSize: 9, bold: true, fill: { color: "F4F6F8" } } },
    { text: `${totalTransferts >= 0 ? "+" : ""}${fmt(totalTransferts)} L`, options: { fontSize: 9, bold: true, fill: { color: "F4F6F8" } } },
    { text: `${fmt(totalTheorique)} L`, options: { fontSize: 9, bold: true, fill: { color: "F4F6F8" } } },
    { text: `${fmt(totalStockFin)} L`, options: { fontSize: 9, bold: true, fill: { color: "F4F6F8" } } },
    { text: `${totalEcart >= 0 ? "+" : ""}${fmt(totalEcart)} L`, options: { fontSize: 9, bold: true, fill: { color: "F4F6F8" } } },
  ];
  const tableRows = [header, ...dataRows, totalRow];
  s2.addTable(tableRows, { x: 0.3, y: 0.8, w: 9.4, autoPage: true, border: { type: "solid", color: "E2E6E9", pt: 0.5 }, fontFace: "Arial" });

  const s3 = pptx.addSlide();
  s3.addShape("rect", { x: 0, y: 0, w: 6, h: 0.08, fill: { color: BLUE } });
  s3.addShape("rect", { x: 6, y: 0, w: 4, h: 0.08, fill: { color: ORANGE } });
  s3.addText("Comparaison par site", { x: 0.4, y: 0.25, w: 9, h: 0.4, fontSize: 18, bold: true, color: INK });
  const labels = rows.map((r) => r.site?.name || "?");
  const chartData = [
    { name: "Réception", labels, values: rows.map((r) => r.reception15) },
    { name: "Ventes", labels, values: rows.map((r) => r.ventes15) },
    { name: "Stock fin", labels, values: rows.map((r) => r.stockFin15) },
  ];
  s3.addChart(pptx.ChartType.bar, chartData, {
    x: 0.4, y: 0.9, w: 9.2, h: 4.3, barDir: "col",
    chartColors: [BLUE, ORANGE, "2E9B5C"], showLegend: true, legendPos: "b",
    showValue: false, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9,
  });

  for (const r of rows) {
    if (!r.photoUrls || r.photoUrls.length === 0) continue;
    const s = pptx.addSlide();
    s.addShape("rect", { x: 0, y: 0, w: 6, h: 0.08, fill: { color: BLUE } });
    s.addShape("rect", { x: 6, y: 0, w: 4, h: 0.08, fill: { color: ORANGE } });
    s.addText(`Photos justificatives — ${r.site?.name || "?"}`, { x: 0.4, y: 0.25, w: 9, h: 0.4, fontSize: 16, bold: true, color: INK });
    if (r.commentaire) s.addText(r.commentaire, { x: 0.4, y: 0.65, w: 9.2, h: 0.35, fontSize: 10, color: SUB, italic: true });
    const positions = [{ x: 0.4, y: 1.1 }, { x: 3.55, y: 1.1 }, { x: 6.7, y: 1.1 }, { x: 0.4, y: 3.4 }, { x: 3.55, y: 3.4 }, { x: 6.7, y: 3.4 }];
    for (let i = 0; i < Math.min(r.photoUrls.length, 6); i++) {
      try { s.addImage({ path: r.photoUrls[i], x: positions[i].x, y: positions[i].y, w: 3, h: 2.1, sizing: { type: "cover", w: 3, h: 2.1 } }); } catch (e) { /* ignore */ }
    }
  }

  await pptx.writeFile({ fileName: `SOMIP_Bilan_Matieres_TousSites_${periodType}_${periodKey}.pptx` });
}

function ReportHeader({ title, period, showEditedDate = true }) {
  return (
    <div className="somip-print-only" style={{ marginBottom: 16 }}>
      <div style={{ height: 5, background: `linear-gradient(90deg, ${C.blue} 0%, ${C.blue} 60%, ${C.orange} 60%, ${C.orange} 100%)`, borderRadius: 3, marginBottom: 12 }} />
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", borderBottom: `2px solid ${C.blue}`, paddingBottom: 10, marginBottom: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          {CURRENT_LOGO_URL && <img src={CURRENT_LOGO_URL} alt="" style={{ height: 34, width: 34, objectFit: "cover", borderRadius: 6 }} />}
          <div>
            <div style={{ fontWeight: 800, fontSize: 17, color: C.blue, letterSpacing: 0.3 }}>SOMIP <span style={{ color: C.orange }}>—</span> Stock Gasoil</div>
            <div style={{ fontSize: 11, color: C.sub }}>Zone Sud-Est · Gabon</div>
          </div>
        </div>
        {showEditedDate && (
          <div style={{ textAlign: "right", fontSize: 11, color: C.sub }}>
            Édité le {new Date().toLocaleDateString("fr-FR")} à {new Date().toLocaleTimeString("fr-FR")}
          </div>
        )}
      </div>
      <h2 style={{ margin: "0 0 2px", fontSize: 16, color: C.navy }}>{title}</h2>
      {period && <div style={{ fontSize: 12.5, color: C.orange, fontWeight: 600 }}>{period}</div>}
    </div>
  );
}

function ReportToolbar({ onExcel, onPrint, onPdf }) {
  return (
    <div className="somip-no-print" style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
      <button className="somip-btn somip-btn-ghost" onClick={onExcel}><Download size={14} /> Export Excel</button>
      {onPdf && <button className="somip-btn somip-btn-primary" onClick={onPdf}><Download size={14} /> Télécharger PDF</button>}
      <button className="somip-btn somip-btn-ghost" onClick={onPrint}><Printer size={14} /> Imprimer</button>
    </div>
  );
}



/* ------------------------------------------------------------------ */
/* Rôles & permissions                                                  */
/* ------------------------------------------------------------------ */
const ROLE_VALUES = ["superviseur", "operateur", "chauffeur", "lecture", "totalenergies"];
const ROLE_LABELS = { superviseur: "Superviseur", operateur: "Opérateur", chauffeur: "Chauffeur", lecture: "Lecture", totalenergies: "TotalEnergies" };
const PERIOD_TYPE_LABELS = { mensuel: "Mensuel", trimestriel: "Trimestriel", decadaire: "Décadaire" };
const TYPE_INVENTAIRE_LABELS = { inopine: "Inopiné", mensuel: "Mensuel" };
const PRODUIT_INVENTAIRE_LABELS = { gasoil: "Gasoil", lubrifiant_vrac: "Lubrifiant vrac" };
// canManage : sites, utilisateurs, réglages, modification/suppression, historique.
// canWrite  : peut ajouter des réceptions/sorties/inventaires (saisie).
function permsFor(role) {
  return {
    canManage: role === "superviseur",
    canWrite: role === "superviseur" || role === "operateur" || role === "chauffeur",
    canInventaireOfficiel: role === "superviseur" || role === "operateur",
    // Qui peut apposer chaque signature. TotalEnergies : uniquement sa propre case.
    canSignSomip: role === "superviseur" || role === "operateur",
    canSignOperateur: role === "superviseur" || role === "operateur",
    canSignTotal: role === "totalenergies" || role === "superviseur",
    isTotalEnergiesOnly: role === "totalenergies",
  };
}

/* ------------------------------------------------------------------ */
/* Base de données partagée (Supabase) — conversion lignes <-> objets  */
/* ------------------------------------------------------------------ */
const numOrUndef = (v) => (v === null || v === undefined ? undefined : Number(v));

const rowToSite = (r) => ({ id: r.id, code: r.code, name: r.name, capacity: Number(r.capacity), stockInitial: Number(r.stock_initial), isMobile: !!r.is_mobile, active: r.active !== false });
const siteToRow = (s) => ({ id: s.id, code: s.code, name: s.name, capacity: s.capacity, stock_initial: s.stockInitial, is_mobile: !!s.isMobile });

const rowToMovement = (r) => ({
  id: r.id, siteId: r.site_id, type: r.type, date: r.date, quantity: Number(r.quantity), delta: Number(r.delta),
  product: r.product || "gasoil",
  ref: r.ref || undefined, commentaire: r.commentaire || "", destinataire: r.destinataire || undefined,
  camion: r.camion || undefined, destination: r.destination || undefined, compteur: r.compteur || undefined, isDemo: !!r.is_demo,
  temperatureC: numOrUndef(r.temperature_c), densiteObservee: numOrUndef(r.densite_observee),
  densite15: numOrUndef(r.densite15), vcf: numOrUndef(r.vcf), volumeCorrige15: numOrUndef(r.volume_corrige15),
  indexAvant: numOrUndef(r.index_avant), indexApres: numOrUndef(r.index_apres),
  createdBy: r.created_by, createdAt: r.created_at,
});
const movementToRow = (m) => ({
  site_id: m.siteId, type: m.type, date: m.date, quantity: m.quantity, delta: m.delta,
  product: m.product || "gasoil",
  ref: m.ref ?? null, commentaire: m.commentaire ?? null, destinataire: m.destinataire ?? null,
  camion: m.camion ?? null, destination: m.destination ?? null, compteur: m.compteur ?? null, is_demo: !!m.isDemo,
  temperature_c: m.temperatureC ?? null, densite_observee: m.densiteObservee ?? null,
  densite15: m.densite15 ?? null, vcf: m.vcf ?? null, volume_corrige15: m.volumeCorrige15 ?? null,
  index_avant: m.indexAvant ?? null, index_apres: m.indexApres ?? null,
  created_by: m.createdBy ?? null,
});

const rowToInventaire = (r) => ({
  id: r.id, siteId: r.site_id, date: r.date, stockPhysique: Number(r.stock_physique), commentaire: r.commentaire || "",
  product: r.product || "gasoil",
  basisEcart: r.basis_ecart, stockTheoriqueAmbiant: numOrUndef(r.stock_theorique_ambiant), stockTheorique15: numOrUndef(r.stock_theorique15),
  stockTheorique: numOrUndef(r.stock_theorique), stockPhysiqueUsed: numOrUndef(r.stock_physique_used),
  ecart: Number(r.ecart), ecartPermille: Number(r.ecart_permille), nature: r.nature, tauxFreinte: Number(r.taux_freinte),
  objectifUtilise: numOrUndef(r.objectif_utilise), conformite: r.conformite, adjustmentId: r.adjustment_id,
  temperatureC: numOrUndef(r.temperature_c), densiteObservee: numOrUndef(r.densite_observee), densite15: numOrUndef(r.densite15),
  vcf: numOrUndef(r.vcf), stockPhysique15: numOrUndef(r.stock_physique15), photoUrls: r.photo_urls || [], createdBy: r.created_by, createdAt: r.created_at,
});
const inventaireToRow = (i) => ({
  site_id: i.siteId, date: i.date, stock_physique: i.stockPhysique, commentaire: i.commentaire ?? null,
  product: i.product || "gasoil",
  basis_ecart: i.basisEcart, stock_theorique_ambiant: i.stockTheoriqueAmbiant ?? null, stock_theorique15: i.stockTheorique15 ?? null,
  stock_theorique: i.stockTheorique ?? null, stock_physique_used: i.stockPhysiqueUsed ?? null, ecart: i.ecart, ecart_permille: i.ecartPermille,
  nature: i.nature, taux_freinte: i.tauxFreinte, objectif_utilise: i.objectifUtilise ?? null, conformite: i.conformite,
  adjustment_id: i.adjustmentId ?? null, temperature_c: i.temperatureC ?? null, densite_observee: i.densiteObservee ?? null,
  densite15: i.densite15 ?? null, vcf: i.vcf ?? null, stock_physique15: i.stockPhysique15 ?? null, photo_urls: i.photoUrls || [], created_by: i.createdBy ?? null,
});

const rowToProductStock = (r) => ({ id: r.id, siteId: r.site_id, product: r.product, capacity: Number(r.capacity), stockInitial: Number(r.stock_initial) });
const productStockToRow = (p) => ({ site_id: p.siteId, product: p.product, capacity: p.capacity, stock_initial: p.stockInitial });

const rowToSiteMeter = (r) => ({ id: r.id, siteId: r.site_id, name: r.name });
const siteMeterToRow = (m) => ({ site_id: m.siteId, name: m.name });

const rowToBilan = (r) => ({
  id: r.id, siteId: r.site_id, periodType: r.period_type, periodKey: r.period_key,
  reception15: Number(r.reception15), ventes15: Number(r.ventes15), transferts15: Number(r.transferts15), stockFin15: Number(r.stock_fin15),
  commentaire: r.commentaire || "", photoUrls: r.photo_urls || [], createdBy: r.created_by, createdAt: r.created_at,
});
const bilanToRow = (b) => ({
  site_id: b.siteId, period_type: b.periodType, period_key: b.periodKey,
  reception15: b.reception15, ventes15: b.ventes15, transferts15: b.transferts15, stock_fin15: b.stockFin15,
  commentaire: b.commentaire ?? null, photo_urls: b.photoUrls || [], created_by: b.createdBy ?? null,
});

const rowToInventaireOfficiel = (r) => ({
  id: r.id, siteId: r.site_id, date: r.date, type: r.type, produit: r.produit || "gasoil", inventoriste: r.inventoriste || "", operateur: r.operateur || "",
  cuves: r.cuves || [], depotage: r.depotage || [], indexCompteurs: r.index_compteurs || [], stockAmbiant: Number(r.stock_ambiant || 0), stock15: numOrUndef(r.stock15), commentaire: r.commentaire || "",
  signatureSomipUrl: r.signature_somip_url || null, signatureSomipBy: r.signature_somip_by || null, signatureSomipAt: r.signature_somip_at || null,
  signatureOperateurUrl: r.signature_operateur_url || null, signatureOperateurBy: r.signature_operateur_by || null, signatureOperateurAt: r.signature_operateur_at || null,
  signatureTotalUrl: r.signature_total_url || null, signatureTotalBy: r.signature_total_by || null, signatureTotalAt: r.signature_total_at || null,
  createdBy: r.created_by, createdAt: r.created_at,
});
const inventaireOfficielToRow = (i) => ({
  site_id: i.siteId, date: i.date, type: i.type, produit: i.produit || "gasoil", inventoriste: i.inventoriste ?? null, operateur: i.operateur ?? null,
  cuves: i.cuves || [], depotage: i.depotage || [], index_compteurs: i.indexCompteurs || [], stock_ambiant: i.stockAmbiant || 0, stock15: i.stock15 ?? null, commentaire: i.commentaire ?? null,
  created_by: i.createdBy ?? null,
});
const rowToSiteTank = (r) => ({ id: r.id, siteId: r.site_id, name: r.name });
const siteTankToRow = (t) => ({ site_id: t.siteId, name: t.name });
const rowToSiteDepotageMeter = (r) => ({ id: r.id, siteId: r.site_id, name: r.name });
const siteDepotageMeterToRow = (t) => ({ site_id: t.siteId, name: t.name });

const rowToAssignment = (r) => ({ id: r.id, truckId: r.truck_id, stationId: r.station_id, startDate: r.start_date, endDate: r.end_date || null });
const assignmentToRow = (a) => ({ truck_id: a.truckId, station_id: a.stationId, start_date: a.startDate, end_date: a.endDate ?? null });


const rowToAudit = (r) => ({ id: r.id, ts: r.ts, user: r.user_name, action: r.action, detail: r.detail });
const rowToProfile = (r) => ({ id: r.id, name: r.full_name, role: r.role, lastSeenAt: r.last_seen_at || null, assignedSiteId: r.assigned_site_id || null, assignedSiteIds: r.assigned_site_ids || [], active: r.active !== false });

async function fetchTable(table, mapper, orderCol, ascending) {
  if (!SUPABASE_CONFIGURED) return [];
  let q = supabase.from(table).select("*");
  if (orderCol) q = q.order(orderCol, { ascending: !!ascending });
  const { data, error } = await q;
  if (error) return [];
  return (data || []).map(mapper);
}



/* ------------------------------------------------------------------ */
/* Small reusable UI                                                    */
/* ------------------------------------------------------------------ */
function ConfirmIconButton({ onConfirm, title }) {
  const [armed, setArmed] = useState(false);
  const timer = useRef(null);
  useEffect(() => () => clearTimeout(timer.current), []);
  return (
    <button
      title={armed ? "Cliquer pour confirmer" : title || "Supprimer"}
      onClick={(e) => {
        e.stopPropagation();
        if (!armed) { setArmed(true); timer.current = setTimeout(() => setArmed(false), 2500); }
        else { clearTimeout(timer.current); setArmed(false); onConfirm(); }
      }}
      style={{
        display: "inline-flex", alignItems: "center", justifyContent: "center",
        width: 28, height: 28, borderRadius: 6, border: "none", cursor: "pointer",
        background: armed ? C.danger : "transparent", color: armed ? "#fff" : C.sub,
        transition: "background .15s, color .15s",
      }}
    >
      {armed ? <Check size={14} /> : <Trash2 size={14} />}
    </button>
  );
}

function ConfirmTextButton({ onConfirm, label, confirmLabel, className }) {
  const [armed, setArmed] = useState(false);
  const timer = useRef(null);
  useEffect(() => () => clearTimeout(timer.current), []);
  return (
    <button
      className={className || "somip-btn somip-btn-ghost"}
      style={armed ? { background: C.danger, color: "#fff", borderColor: C.danger } : undefined}
      onClick={() => {
        if (!armed) { setArmed(true); timer.current = setTimeout(() => setArmed(false), 2800); }
        else { clearTimeout(timer.current); setArmed(false); onConfirm(); }
      }}
    >
      {armed ? (confirmLabel || "Confirmer ?") : label}
    </button>
  );
}

function Badge({ color, children }) {
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 5, padding: "3px 9px",
      borderRadius: 999, fontSize: 11.5, fontWeight: 600, color,
      background: color + "1A", whiteSpace: "nowrap",
    }}>
      <span style={{ width: 6, height: 6, borderRadius: "50%", background: color }} />
      {children}
    </span>
  );
}

function DemoBadge() {
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", padding: "2px 7px", borderRadius: 5,
      fontSize: 10, fontWeight: 700, letterSpacing: 0.3, color: C.sub,
      background: "#EDF0F3", border: `1px dashed #C7CED6`, marginLeft: 6,
    }}>
      DÉMO
    </span>
  );
}

function StatCard({ label, value, unit, accent, icon: Icon }) {
  return (
    <div className="somip-panel" style={{ padding: "16px 18px", borderLeft: `3px solid ${accent}`, flex: 1, minWidth: 190 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
        <span style={{ fontSize: 12.5, fontWeight: 600, color: C.sub }}>{label}</span>
        <Icon size={16} color={accent} />
      </div>
      <div className="somip-mono" style={{ fontSize: 24, fontWeight: 600, color: C.ink }}>
        {value}{unit && <span style={{ fontSize: 12.5, fontWeight: 500, color: C.sub, marginLeft: 5 }}>{unit}</span>}
      </div>
    </div>
  );
}

function GaugeBar({ pct, color }) {
  const clamped = Math.max(0, Math.min(100, pct));
  return (
    <div style={{ background: "#EDF0F3", borderRadius: 999, height: 7, width: "100%", overflow: "hidden" }}>
      <div style={{ width: `${clamped}%`, height: "100%", background: color, borderRadius: 999, transition: "width .3s" }} />
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <label className="somip-label">{label}</label>
      {children}
    </div>
  );
}

function EmptyRow({ colSpan, text }) {
  return (
    <tr><td colSpan={colSpan} style={{ padding: "26px 12px", textAlign: "center", color: C.sub, fontSize: 13 }}>{text}</td></tr>
  );
}

function VcfMiniPanel({ tempC, densite, onTempC, onDensite, result, compact }) {
  return (
    <div style={{ background: C.bg, borderRadius: 8, padding: compact ? 10 : 12, marginBottom: 12 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8 }}>
        <Thermometer size={13} color={C.blue} />
        <span style={{ fontSize: 11.5, fontWeight: 700, color: C.sub }}>Correction à 15°C (optionnel)</span>
      </div>
      <div style={{ display: "flex", gap: 8, marginBottom: result ? 8 : 0 }}>
        <div style={{ flex: 1 }}>
          <label className="somip-label" style={{ fontSize: 11 }}>Température (°C)</label>
          <input type="number" step="0.1" className="somip-input" value={tempC} onChange={(e) => onTempC(e.target.value)} placeholder="Ex : 28.5" />
        </div>
        <div style={{ flex: 1 }}>
          <label className="somip-label" style={{ fontSize: 11 }}>Densité observée (kg/m³)</label>
          <input type="number" step="0.1" className="somip-input" value={densite} onChange={(e) => onDensite(e.target.value)} placeholder="Ex : 845" />
        </div>
      </div>
      {tempC !== "" && densite !== "" && !result && (
        <p style={{ margin: "6px 0 0", fontSize: 11, color: C.warning }}>
          Densité hors plage {VCF_MIN_DENSITY}–{VCF_MAX_DENSITY} kg/m³, ou valeurs incomplètes — correction non calculée.
        </p>
      )}
      {result && (
        <div style={{ fontSize: 12, marginTop: 4 }}>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
            <span style={{ color: C.sub }}>Densité à 15°C</span>
            <span className="somip-mono" style={{ fontWeight: 600 }}>{result.densite15.toFixed(2)} kg/m³</span>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
            <span style={{ color: C.sub }}>VCF</span>
            <span className="somip-mono" style={{ fontWeight: 600 }}>{result.vcf.toFixed(5)}</span>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <span style={{ color: C.sub }}>Volume corrigé à 15°C</span>
            <span className="somip-mono" style={{ fontWeight: 700, color: C.blue }}>{fmt(result.volume15)} L</span>
          </div>
        </div>
      )}
    </div>
  );
}

function SyncIndicator({ status, lastSync }) {
  const map = {
    saving: { icon: Loader2, color: C.blue, label: "Enregistrement...", spin: true },
    ok: { icon: CheckCircle2, color: C.success, label: "Données sauvegardées" },
    error: { icon: AlertCircle, color: C.danger, label: "Erreur de sauvegarde — nouvelle tentative au prochain changement" },
    unavailable: { icon: CloudOff, color: C.warning, label: "Stockage persistant indisponible dans cet environnement" },
  };
  const m = map[status] || map.ok;
  const Icon = m.icon;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 11.5, color: m.color, fontWeight: 600 }}>
      <Icon size={13} style={m.spin ? { animation: "somipSpin .8s linear infinite" } : undefined} />
      {m.label}
      {status === "ok" && lastSync && (
        <span style={{ color: C.sub, fontWeight: 500 }}>· {lastSync.toLocaleTimeString("fr-FR")}</span>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Écran de connexion / inscription                                     */
/* ------------------------------------------------------------------ */
function AuthScreen() {
  const [mode, setMode] = useState("login"); // "login" | "signup"
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [fullName, setFullName] = useState("");
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [busy, setBusy] = useState(false);
  const [logoUrl, setLogoUrl] = useState(null);

  useEffect(() => {
    supabase.from("settings").select("logo_url").eq("id", 1).maybeSingle()
      .then(({ data }) => { if (data?.logo_url) setLogoUrl(data.logo_url); })
      .catch(() => {});
  }, []);

  const submit = async () => {
    setError(""); setInfo("");
    if (!email || !password) { setError("Adresse e-mail et mot de passe requis."); return; }
    setBusy(true);
    if (mode === "login") {
      const { error: err } = await supabase.auth.signInWithPassword({ email, password });
      if (err) setError("Connexion impossible : " + err.message);
    } else {
      const { error: err } = await supabase.auth.signUp({
        email, password, options: { data: { full_name: fullName || email } },
      });
      if (err) setError("Inscription impossible : " + err.message);
      else setInfo("Compte créé. Un Superviseur doit maintenant t'attribuer un rôle depuis la page Utilisateurs avant que tu puisses saisir des données. Connecte-toi dès que c'est fait.");
    }
    setBusy(false);
  };

  return (
    <div style={{ minHeight: "100%", display: "flex", alignItems: "center", justifyContent: "center", background: C.bg, padding: 24, fontFamily: "'Inter', -apple-system, sans-serif" }}>
      <div style={{ background: "#fff", border: `1px solid ${C.border}`, borderRadius: 12, padding: 30, width: "100%", maxWidth: 380 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 9, marginBottom: 22 }}>
          {logoUrl ? (
            <img src={logoUrl} alt="Logo" style={{ width: 34, height: 34, borderRadius: 8, objectFit: "cover" }} />
          ) : (
            <div style={{ width: 34, height: 34, borderRadius: 8, background: C.blue, display: "flex", alignItems: "center", justifyContent: "center" }}>
              <Fuel size={18} color="#fff" />
            </div>
          )}
          <div>
            <div style={{ fontWeight: 700, fontSize: 15 }}>SOMIP</div>
            <div style={{ color: C.sub, fontSize: 11 }}>Stock Gasoil</div>
          </div>
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 20 }}>
          <button className={`somip-tab ${mode === "login" ? "active" : ""}`} style={{ flex: 1, textAlign: "center" }} onClick={() => { setMode("login"); setError(""); setInfo(""); }}>Connexion</button>
          <button className={`somip-tab ${mode === "signup" ? "active" : ""}`} style={{ flex: 1, textAlign: "center" }} onClick={() => { setMode("signup"); setError(""); setInfo(""); }}>Créer un compte</button>
        </div>

        {mode === "signup" && (
          <Field label="Nom complet">
            <input className="somip-input" value={fullName} onChange={(e) => setFullName(e.target.value)} placeholder="Ex : Jean Mabiala" />
          </Field>
        )}
        <Field label="E-mail ou identifiant">
          <input type="email" className="somip-input" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="prenom.nom@somip-sarl.ga" />
        </Field>
        <Field label="Mot de passe">
          <input type="password" className="somip-input" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" />
        </Field>

        {error && <p style={{ color: C.danger, fontSize: 12.5, margin: "0 0 12px" }}>{error}</p>}
        {info && <p style={{ color: C.success, fontSize: 12.5, margin: "0 0 12px" }}>{info}</p>}

        <button className="somip-btn somip-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={submit} disabled={busy}>
          {mode === "login" ? <Lock size={15} /> : <Mail size={15} />}
          {mode === "login" ? "Se connecter" : "Créer mon compte"}
        </button>

        {mode === "signup" && (
          <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
            Par défaut, un nouveau compte n'a que des droits de consultation. Un Superviseur doit t'accorder le droit de saisie depuis la page Utilisateurs.
          </p>
        )}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* App                                                                   */
/* ------------------------------------------------------------------ */
export default function App() {
  const [authLoading, setAuthLoading] = useState(true);
  const [session, setSession] = useState(null);
  const [profile, setProfile] = useState(null);

  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState(null);
  const [retryKey, setRetryKey] = useState(0);
  const [sites, setSites] = useState([]);
  const [movements, setMovements] = useState([]);
  const [offlineQueueCount, setOfflineQueueCount] = useState(() => readOfflineQueue().length);
  const [inventaires, setInventaires] = useState([]);
  const [productStocks, setProductStocks] = useState([]);
  const [siteMeters, setSiteMeters] = useState([]);
  const [bilans, setBilans] = useState([]);
  const [inventairesOfficiels, setInventairesOfficiels] = useState([]);
  const [siteTanks, setSiteTanks] = useState([]);
  const [siteDepotageMeters, setSiteDepotageMeters] = useState([]);
  const [truckAssignments, setTruckAssignments] = useState([]);
  const [profiles, setProfiles] = useState([]);
  const [audit, setAudit] = useState([]);
  const [settings, setSettings] = useState(SETTINGS_SEED);
  const [view, setView] = useState("dashboard");
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const [notice, setNotice] = useState(null);
  const [noticeType, setNoticeType] = useState("success");
  const [syncStatus, setSyncStatus] = useState(SUPABASE_CONFIGURED ? "ok" : "unavailable");
  const [lastSync, setLastSync] = useState(null);
  const noticeTimer = useRef(null);

  const flash = (msg, type = "success") => {
    clearTimeout(noticeTimer.current);
    setNotice(msg);
    setNoticeType(type);
    if (type !== "error") {
      noticeTimer.current = setTimeout(() => setNotice(null), 3000);
    }
  };

  /* ---- authentification ---- */
  useEffect(() => {
    if (!SUPABASE_CONFIGURED) { setAuthLoading(false); return; }
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session);
      setAuthLoading(false);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => listener.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) { setProfile(null); return; }
    let cancelled = false;
    (async () => {
      try {
        const timeout = (ms) => new Promise((_, reject) => setTimeout(() => reject(new Error("Délai dépassé lors du chargement du profil (le serveur ne répond pas).")), ms));
        const fetchProfile = supabase.from("profiles").select("*").eq("id", session.user.id).maybeSingle();
        const { data, error } = await Promise.race([fetchProfile, timeout(15000)]);
        if (cancelled) return;
        if (error) throw error;
        if (!data) {
          setLoadError("Aucun profil trouvé pour ce compte. Contacte un Superviseur pour vérifier ta fiche dans la table 'profiles'.");
          setLoading(false);
          return;
        }
        if (data.active === false) {
          await supabase.auth.signOut();
          setLoadError("Ce compte a été désactivé par un Superviseur. Contacte-le si tu penses qu'il s'agit d'une erreur.");
          setLoading(false);
          return;
        }
        setProfile(rowToProfile(data));
      } catch (e) {
        if (cancelled) return;
        setLoadError(e?.message || "Erreur lors du chargement du profil.");
        setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [session]);

  const currentUserName = profile?.name || session?.user?.email || "Utilisateur";
  const currentRole = profile?.role || "lecture";
  const perms = permsFor(currentRole);
  const signOut = () => supabase.auth.signOut();

  /* ---- chargement initial des données (une fois connecté) ---- */
  useEffect(() => {
    if (!session || !profile) return;
    let cancelled = false;
    (async () => {
      try {
        const timeout = (ms) => new Promise((_, reject) => setTimeout(() => reject(new Error("Délai dépassé (le serveur ne répond pas)")), ms));
        const load = (async () => {
          const [sitesData, movementsData, inventairesData, profilesData, auditData, productStocksData, assignmentsData, siteMetersData, bilansData, invOffData, siteTanksData, siteDepotageMetersData] = await Promise.all([
            fetchTable("sites", rowToSite),
            fetchTable("movements", rowToMovement, "date"),
            fetchTable("inventaires", rowToInventaire, "date"),
            fetchTable("profiles", rowToProfile),
            fetchTable("audit", rowToAudit, "ts", false),
            fetchTable("product_stocks", rowToProductStock),
            fetchTable("truck_assignments", rowToAssignment, "start_date"),
            fetchTable("site_meters", rowToSiteMeter, "name"),
            fetchTable("bilan_matieres", rowToBilan, "period_key"),
            fetchTable("inventaires_officiels", rowToInventaireOfficiel, "date"),
            fetchTable("site_tanks", rowToSiteTank, "name"),
            fetchTable("site_depotage_meters", rowToSiteDepotageMeter, "name"),
          ]);
          let settingsRow = null;
          try {
            const res = await supabase.from("settings").select("*").eq("id", 1).maybeSingle();
            settingsRow = res.data;
          } catch (e) { /* réglages optionnels : on garde la valeur par défaut si ça échoue */ }
          return { sitesData, movementsData, inventairesData, profilesData, auditData, productStocksData, assignmentsData, siteMetersData, bilansData, invOffData, siteTanksData, siteDepotageMetersData, settingsRow };
        })();
        const result = await Promise.race([load, timeout(15000)]);
        if (cancelled) return;
        setSites(result.sitesData);
        setMovements(result.movementsData);
        setInventaires(result.inventairesData);
        setProfiles(result.profilesData);
        setAudit(result.auditData);
        setProductStocks(result.productStocksData);
        setTruckAssignments(result.assignmentsData);
        setSiteMeters(result.siteMetersData);
        setBilans(result.bilansData);
        setInventairesOfficiels(result.invOffData);
        setSiteTanks(result.siteTanksData);
        setSiteDepotageMeters(result.siteDepotageMetersData);
        setSettings(rowToSettings(result.settingsRow));
        setLastSync(new Date());
        setLoadError(null);
        setLoading(false);
      } catch (e) {
        if (cancelled) return;
        setLoadError(e?.message || "Erreur de chargement inconnue.");
        setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [session, profile, retryKey]);

  /* ---- synchronisation périodique : voir les changements des autres utilisateurs ---- */
  const heartbeatRef = useRef(0);
  useEffect(() => {
    if (loading || !session || !profile) return;
    const interval = setInterval(async () => {
      if (navigator.onLine) flushOfflineQueue();
      const [s, m, i, p, a, ps, ta, sm, bl, io, st, sdm] = await Promise.all([
        fetchTable("sites", rowToSite),
        fetchTable("movements", rowToMovement, "date"),
        fetchTable("inventaires", rowToInventaire, "date"),
        fetchTable("profiles", rowToProfile),
        fetchTable("audit", rowToAudit, "ts", false),
        fetchTable("product_stocks", rowToProductStock),
        fetchTable("truck_assignments", rowToAssignment, "start_date"),
        fetchTable("site_meters", rowToSiteMeter, "name"),
        fetchTable("bilan_matieres", rowToBilan, "period_key"),
        fetchTable("inventaires_officiels", rowToInventaireOfficiel, "date"),
        fetchTable("site_tanks", rowToSiteTank, "name"),
        fetchTable("site_depotage_meters", rowToSiteDepotageMeter, "name"),
      ]);
      setSites(s); setMovements(m); setInventaires(i); setProfiles(p); setAudit(a); setProductStocks(ps); setTruckAssignments(ta); setSiteMeters(sm); setBilans(bl); setInventairesOfficiels(io); setSiteTanks(st); setSiteDepotageMeters(sdm);
      const { data: se } = await supabase.from("settings").select("*").eq("id", 1).maybeSingle();
      if (se) setSettings(rowToSettings(se));
      setLastSync(new Date());
      // Présence en ligne : met à jour la dernière activité connue, au plus toutes les 30s.
      const now = Date.now();
      if (now - heartbeatRef.current > 30000) {
        heartbeatRef.current = now;
        supabase.from("profiles").update({ last_seen_at: new Date().toISOString() }).eq("id", session.user.id);
      }
    }, 7000);
    return () => clearInterval(interval);
  }, [loading, session, profile]);

  // Vide la file d'attente hors-connexion vers Supabase, dès qu'une tentative semble possible.
  // Ce qui échoue encore (toujours hors-ligne) reste en attente pour la prochaine tentative.
  const flushingRef = useRef(false);
  const flushOfflineQueue = async () => {
    if (flushingRef.current) return;
    const q = readOfflineQueue();
    if (q.length === 0) return;
    flushingRef.current = true;
    const remaining = [];
    let anySucceeded = false;
    for (const item of q) {
      try {
        const { error } = await supabase.from(item.table).insert(item.row);
        if (error) throw error;
        anySucceeded = true;
      } catch (e) {
        remaining.push(item);
      }
    }
    writeOfflineQueue(remaining);
    setOfflineQueueCount(remaining.length);
    flushingRef.current = false;
    if (anySucceeded) {
      // Au moins une entrée a été synchronisée : les remplace localement par les vraies
      // données du serveur (retire les entrées optimistes temporaires du même coup).
      const [m, i] = await Promise.all([fetchTable("movements", rowToMovement, "date"), fetchTable("inventaires", rowToInventaire, "date")]);
      setMovements(m); setInventaires(i);
      flash(remaining.length === 0 ? "Toutes les saisies hors-connexion ont été synchronisées." : `${q.length - remaining.length} saisie(s) synchronisée(s), ${remaining.length} en attente.`);
    }
  };
  useEffect(() => {
    const onOnline = () => flushOfflineQueue();
    window.addEventListener("online", onOnline);
    return () => window.removeEventListener("online", onOnline);
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const appendAudit = async (action, detail) => {
    const { data } = await supabase.from("audit").insert({ user_name: currentUserName, action, detail }).select().maybeSingle();
    if (data) setAudit((prev) => [rowToAudit(data), ...prev].slice(0, 300));
  };

  const withSync = async (fn) => {
    setSyncStatus("saving");
    try {
      await fn();
      setSyncStatus("ok");
      setLastSync(new Date());
      return true;
    } catch (e) {
      setSyncStatus("error");
      console.error(e);
      flash(e?.message ? `Erreur : ${e.message}` : "Action refusée ou erreur de sauvegarde.", "error");
      return false;
    }
  };

  /* ---- dérivés ---- */
  const stockOf = (siteId, product = "gasoil") => {
    if (product === "gasoil") {
      const site = sites.find((s) => s.id === siteId);
      if (!site) return 0;
      return movements.filter((m) => m.siteId === siteId && (m.product || "gasoil") === "gasoil").reduce((acc, m) => acc + m.delta, site.stockInitial);
    }
    const ps = productStocks.find((p) => p.siteId === siteId && p.product === product);
    if (!ps) return 0;
    return movements.filter((m) => m.siteId === siteId && m.product === product).reduce((acc, m) => acc + m.delta, ps.stockInitial);
  };
  const stockOf15 = (siteId, product = "gasoil") => {
    if (product === "gasoil") {
      const site = sites.find((s) => s.id === siteId);
      if (!site) return 0;
      return movements.filter((m) => m.siteId === siteId && (m.product || "gasoil") === "gasoil").reduce((acc, m) => acc + Math.sign(m.delta) * movementQty15(m), site.stockInitial);
    }
    const ps = productStocks.find((p) => p.siteId === siteId && p.product === product);
    if (!ps) return 0;
    return movements.filter((m) => m.siteId === siteId && m.product === product).reduce((acc, m) => acc + Math.sign(m.delta) * movementQty15(m), ps.stockInitial);
  };

  /* ---- mutations : stocks de lubrifiants (Superviseur uniquement) ---- */
  const saveProductStock = ({ siteId, product, capacity, stockInitial }) => withSync(async () => {
    const row = productStockToRow({ siteId, product, capacity: Number(capacity), stockInitial: Number(stockInitial) || 0 });
    const { data, error } = await supabase.from("product_stocks").upsert(row, { onConflict: "site_id,product" }).select().maybeSingle();
    if (error) throw error;
    if (!data) throw new Error("La mise à jour n'a pas pu être confirmée par le serveur — réessaie.");
    const saved = rowToProductStock(data);
    setProductStocks((prev) => {
      const exists = prev.some((p) => p.siteId === siteId && p.product === product);
      return exists ? prev.map((p) => (p.siteId === siteId && p.product === product ? saved : p)) : [...prev, saved];
    });
    appendAudit("Réglage lubrifiant", `${LUBRICANTS.find((l) => l.id === product)?.label || product} — ${sites.find((s) => s.id === siteId)?.name || ""}`);
    flash("Stock de lubrifiant mis à jour.");
  });

  /* ---- mutations : compteurs par site (Superviseur uniquement) ---- */
  const addSiteMeter = ({ siteId, name }) => withSync(async () => {
    const cleanName = name.trim();
    if (!cleanName) throw new Error("Le nom du compteur ne peut pas être vide.");
    const { data, error } = await supabase.from("site_meters").insert(siteMeterToRow({ siteId, name: cleanName })).select().maybeSingle();
    if (error) throw error;
    if (!data) throw new Error("Le compteur n'a pas pu être confirmé par le serveur — réessaie.");
    setSiteMeters((prev) => [...prev, rowToSiteMeter(data)]);
    appendAudit("Ajout compteur", `${cleanName} — ${sites.find((s) => s.id === siteId)?.name || ""}`);
    flash("Compteur ajouté.");
  });
  const removeSiteMeter = (meter) => withSync(async () => {
    const { data, error } = await supabase.from("site_meters").delete().eq("id", meter.id).select();
    if (error) throw error;
    if (!data || data.length === 0) throw new Error("Suppression refusée par la base de données — le compteur n'a pas été retiré.");
    setSiteMeters((prev) => prev.filter((m) => m.id !== meter.id));
    appendAudit("Suppression compteur", `${meter.name} — ${sites.find((s) => s.id === meter.siteId)?.name || ""}`);
    flash("Compteur supprimé.");
  });

  /* ---- mutations : Bilan Matières (Superviseur uniquement) ---- */
  const saveBilan = ({ siteId, periodType, periodKey, reception15, ventes15, transferts15, stockFin15, commentaire, photoFiles = [], existingPhotoUrls = [] }) => withSync(async () => {
    const newUrls = photoFiles.length ? await uploadPhotos(photoFiles, `bilans/${siteId}/${periodType}`) : [];
    const photoUrls = [...existingPhotoUrls, ...newUrls];
    const row = bilanToRow({ siteId, periodType, periodKey, reception15: Number(reception15) || 0, ventes15: Number(ventes15) || 0, transferts15: Number(transferts15) || 0, stockFin15: Number(stockFin15) || 0, commentaire, photoUrls, createdBy: currentUserName });
    const { data, error } = await supabase.from("bilan_matieres").upsert(row, { onConflict: "site_id,period_type,period_key" }).select().maybeSingle();
    if (error) throw error;
    if (!data) throw new Error("Le Bilan Matières n'a pas pu être confirmé par le serveur — réessaie.");
    const saved = rowToBilan(data);
    setBilans((prev) => {
      const exists = prev.some((b) => b.siteId === siteId && b.periodType === periodType && b.periodKey === periodKey);
      return exists ? prev.map((b) => (b.siteId === siteId && b.periodType === periodType && b.periodKey === periodKey ? saved : b)) : [...prev, saved];
    });
    appendAudit("Bilan Matières", `${sites.find((s) => s.id === siteId)?.name || siteId} — ${PERIOD_TYPE_LABELS[periodType]} ${periodKey}`);
    flash("Bilan Matières enregistré.");
  });
  const deleteBilan = (bilan) => withSync(async () => {
    const { data, error } = await supabase.from("bilan_matieres").delete().eq("id", bilan.id).select();
    if (error) throw error;
    if (!data || data.length === 0) throw new Error("Suppression refusée par la base de données.");
    setBilans((prev) => prev.filter((b) => b.id !== bilan.id));
    flash("Bilan supprimé.");
  });

  /* ---- mutations : Inventaires officiels (inopinés/mensuels), cuve par cuve ---- */
  const addInventaireOfficiel = ({ siteId, date, type, produit, inventoriste, operateur, cuves, depotage, indexCompteurs, commentaire }) => withSync(async () => {
    // Densité et température sont désormais propres à CHAQUE cuve : le volume à 15°C se calcule
    // cuve par cuve, puis on additionne. Le total à 15°C n'est complet que si toutes les cuves
    // ont une densité + température renseignées. Les index (relevé compteur) sont indépendants
    // des cuves — une liste à part, comme les compteurs de dépotage.
    const cuvesComputed = (cuves || []).map((c) => {
      const amb = Number(c.stockAmbiant) || 0;
      let volume15 = null;
      if (c.temperatureC !== undefined && c.temperatureC !== "" && c.temperatureC !== null && c.densite !== undefined && c.densite !== "" && c.densite !== null) {
        const corr = correctVolumeTo15({ volumeAmbiant: amb, tempC: Number(c.temperatureC), densiteObservee: Number(c.densite) });
        if (corr) volume15 = corr.volume15;
      }
      return { ...c, stockAmbiant: amb, volume15 };
    });
    const stockAmbiant = cuvesComputed.reduce((a, c) => a + c.stockAmbiant, 0);
    const stock15 = cuvesComputed.every((c) => c.volume15 !== null) ? cuvesComputed.reduce((a, c) => a + c.volume15, 0) : undefined;
    const row = inventaireOfficielToRow({ siteId, date, type, produit, inventoriste, operateur, cuves: cuvesComputed, depotage: depotage || [], indexCompteurs: indexCompteurs || [], stockAmbiant, stock15, commentaire, createdBy: currentUserName });
    const { data, error } = await supabase.from("inventaires_officiels").insert(row).select().maybeSingle();
    if (error) throw error;
    if (!data) throw new Error("L'inventaire officiel n'a pas pu être confirmé par le serveur — réessaie.");
    const saved = rowToInventaireOfficiel(data);
    setInventairesOfficiels((prev) => [...prev, saved]);
    appendAudit("Inventaire officiel", `${TYPE_INVENTAIRE_LABELS[type]} — ${sites.find((s) => s.id === siteId)?.name || siteId} — ${date}`);
    flash("Inventaire officiel enregistré.");
  });
  const deleteInventaireOfficiel = (inv) => withSync(async () => {
    const { data, error } = await supabase.from("inventaires_officiels").delete().eq("id", inv.id).select();
    if (error) throw error;
    if (!data || data.length === 0) throw new Error("Suppression refusée par la base de données.");
    setInventairesOfficiels((prev) => prev.filter((i) => i.id !== inv.id));
    flash("Inventaire officiel supprimé.");
  });
  const signInventaireOfficiel = (inv, role, dataUrl) => withSync(async () => {
    const file = dataUrlToFile(dataUrl, `signature-${role}.png`);
    const [url] = await uploadPhotos([file], `signatures/${inv.id}`);
    const col = role === "somip" ? "signature_somip" : role === "operateur" ? "signature_operateur" : "signature_total";
    const nowIso = new Date().toISOString();
    const patch = { [`${col}_url`]: url, [`${col}_by`]: currentUserName, [`${col}_at`]: nowIso };
    const { data, error } = await supabase.from("inventaires_officiels").update(patch).eq("id", inv.id).select().maybeSingle();
    if (error) throw error;
    if (!data) throw new Error("La signature n'a pas pu être confirmée par le serveur — réessaie.");
    const saved = rowToInventaireOfficiel(data);
    setInventairesOfficiels((prev) => prev.map((i) => (i.id === inv.id ? saved : i)));
    appendAudit("Signature inventaire officiel", `${role} — ${sites.find((s) => s.id === inv.siteId)?.name || inv.siteId} — ${inv.date}`);
    flash("Signature enregistrée.");
  });

  /* ---- mutations : cuves par site (Superviseur uniquement) ---- */
  const addSiteTank = ({ siteId, name }) => withSync(async () => {
    const cleanName = name.trim();
    if (!cleanName) throw new Error("Le nom de la cuve ne peut pas être vide.");
    const { data, error } = await supabase.from("site_tanks").insert(siteTankToRow({ siteId, name: cleanName })).select().maybeSingle();
    if (error) throw error;
    if (!data) throw new Error("La cuve n'a pas pu être confirmée par le serveur — réessaie.");
    setSiteTanks((prev) => [...prev, rowToSiteTank(data)]);
    appendAudit("Ajout cuve", `${cleanName} — ${sites.find((s) => s.id === siteId)?.name || ""}`);
    flash("Cuve ajoutée.");
  });
  const removeSiteTank = (tank) => withSync(async () => {
    const { data, error } = await supabase.from("site_tanks").delete().eq("id", tank.id).select();
    if (error) throw error;
    if (!data || data.length === 0) throw new Error("Suppression refusée par la base de données — la cuve n'a pas été retirée.");
    setSiteTanks((prev) => prev.filter((t) => t.id !== tank.id));
    appendAudit("Suppression cuve", `${tank.name} — ${sites.find((s) => s.id === tank.siteId)?.name || ""}`);
    flash("Cuve supprimée.");
  });
  const addSiteDepotageMeter = ({ siteId, name }) => withSync(async () => {
    const cleanName = name.trim();
    if (!cleanName) throw new Error("Le nom du compteur ne peut pas être vide.");
    const { data, error } = await supabase.from("site_depotage_meters").insert(siteDepotageMeterToRow({ siteId, name: cleanName })).select().maybeSingle();
    if (error) throw error;
    if (!data) throw new Error("Le compteur n'a pas pu être confirmé par le serveur — réessaie.");
    setSiteDepotageMeters((prev) => [...prev, rowToSiteDepotageMeter(data)]);
    appendAudit("Ajout compteur de dépotage", `${cleanName} — ${sites.find((s) => s.id === siteId)?.name || ""}`);
    flash("Compteur de dépotage ajouté.");
  });
  const removeSiteDepotageMeter = (meter) => withSync(async () => {
    const { data, error } = await supabase.from("site_depotage_meters").delete().eq("id", meter.id).select();
    if (error) throw error;
    if (!data || data.length === 0) throw new Error("Suppression refusée par la base de données — le compteur n'a pas été retiré.");
    setSiteDepotageMeters((prev) => prev.filter((m) => m.id !== meter.id));
    appendAudit("Suppression compteur de dépotage", `${meter.name} — ${sites.find((s) => s.id === meter.siteId)?.name || ""}`);
    flash("Compteur de dépotage supprimé.");
  });

  /* ---- mutations : affectation des camions aux stations (Superviseur uniquement) ---- */
  const assignTruck = ({ truckId, stationId, startDate }) => withSync(async () => {
    // Ferme toute affectation encore ouverte pour ce camion, juste avant la nouvelle date de début.
    const openAssignment = truckAssignments.find((a) => a.truckId === truckId && !a.endDate);
    if (openAssignment) {
      const prevDay = new Date(startDate);
      prevDay.setDate(prevDay.getDate() - 1);
      const endDate = prevDay.toISOString().slice(0, 10);
      const { error: e1 } = await supabase.from("truck_assignments").update({ end_date: endDate }).eq("id", openAssignment.id);
      if (e1) throw e1;
      setTruckAssignments((prev) => prev.map((a) => (a.id === openAssignment.id ? { ...a, endDate } : a)));
    }
    const record = { truckId, stationId, startDate, endDate: null };
    const { data, error } = await supabase.from("truck_assignments").insert(assignmentToRow(record)).select().maybeSingle();
    if (error) throw error;
    if (!data) throw new Error("L'affectation n'a pas pu être confirmée par le serveur — réessaie.");
    const saved = rowToAssignment(data);
    setTruckAssignments((prev) => [...prev, saved]);
    appendAudit("Affectation camion", `${sites.find((s) => s.id === truckId)?.name || truckId} → ${sites.find((s) => s.id === stationId)?.name || stationId} (depuis le ${startDate})`);
    flash("Camion affecté.");
  });

  /* ---- mutations : sites (Superviseur uniquement) ---- */
  const addSite = (form) => withSync(async () => {
    const site = { id: uid(), name: form.name.trim(), code: form.code.trim().toUpperCase(), capacity: Number(form.capacity), stockInitial: Number(form.stockInitial) || 0, isMobile: !!form.isMobile };
    const { error } = await supabase.from("sites").insert(siteToRow(site));
    if (error) throw error;
    setSites((prev) => [...prev, site]);
    appendAudit("Ajout site", `${site.name} (${site.code})`);
    flash("Site ajouté.");
  });
  const editSite = (id, patch) => withSync(async () => {
    const updated = { ...patch, capacity: Number(patch.capacity), stockInitial: Number(patch.stockInitial), isMobile: !!patch.isMobile };
    const { error } = await supabase.from("sites").update(siteToRow({ id, ...updated })).eq("id", id);
    if (error) throw error;
    setSites((prev) => prev.map((s) => (s.id === id ? { ...s, ...updated } : s)));
    appendAudit("Modification site", `${patch.name} (${patch.code})`);
    flash("Site mis à jour.");
  });
  const removeSite = (site) => withSync(async () => {
    if (movements.some((m) => m.siteId === site.id)) { flash(`Impossible : "${site.name}" a des mouvements enregistrés.`); return; }
    const { error } = await supabase.from("sites").delete().eq("id", site.id);
    if (error) throw error;
    setSites((prev) => prev.filter((s) => s.id !== site.id));
    appendAudit("Suppression site", site.name);
    flash("Site supprimé.");
  });

  /* ---- mutations : mouvements ---- */
  const addMovement = (payload) => withSync(async () => {
    const record = { id: uid(), createdBy: currentUserName, createdAt: new Date().toISOString(), isDemo: false, ...payload };
    const row = movementToRow(record);
    try {
      const { data, error } = await supabase.from("movements").insert(row).select().maybeSingle();
      if (error) throw error;
      const saved = data ? rowToMovement(data) : record;
      setMovements((prev) => [...prev, saved]);
      appendAudit(TYPE_META[payload.type].label, `${fmt(payload.quantity)} L — ${sites.find((s) => s.id === payload.siteId)?.name || ""}`);
    } catch (e) {
      if (!isNetworkError(e)) throw e;
      // Pas de réseau : on garde la saisie en local (visible immédiatement) et on la met en
      // file d'attente, synchronisée automatiquement dès le retour de la connexion.
      const q = readOfflineQueue();
      q.push({ id: uid(), table: "movements", row, createdAt: new Date().toISOString() });
      writeOfflineQueue(q);
      setOfflineQueueCount(q.length);
      setMovements((prev) => [...prev, record]);
      flash("Pas de connexion — saisie enregistrée sur l'appareil, sera synchronisée automatiquement dès le retour du réseau.");
    }
    // Chargement et retour cuve se saisissent indépendamment des deux côtés (site et camion) —
    // aucune création automatique de miroir, pour éviter tout risque de double comptage.
  });
  const deleteMovement = (id) => withSync(async () => {
    const m = movements.find((mm) => mm.id === id);
    const { data, error } = await supabase.from("movements").delete().eq("id", id).select();
    if (error) throw error;
    if (!data || data.length === 0) throw new Error("Suppression refusée par la base de données (droits insuffisants ou ligne déjà supprimée) — le mouvement n'a pas été retiré.");
    setMovements((prev) => prev.filter((mm) => mm.id !== id));
    if (m) appendAudit("Suppression mouvement", `${TYPE_META[m.type]?.label || m.type} — ${fmt(m.quantity)} L`);
  });
  const purgeDemoMovements = () => withSync(async () => {
    const demoIds = movements.filter((m) => m.isDemo).map((m) => m.id);
    if (demoIds.length === 0) return;
    const { error } = await supabase.from("movements").delete().in("id", demoIds);
    if (error) throw error;
    setMovements((prev) => prev.filter((m) => !m.isDemo));
    appendAudit("Purge données démo", `${demoIds.length} écriture(s) d'exemple supprimée(s)`);
    flash("Écritures de démonstration supprimées.");
  });

  /* ---- mutations : inventaires ---- */
  const addInventaire = ({ siteId, product = "gasoil", date, stockPhysique, commentaire, temperatureC, densiteObservee, densite15, vcf, stockPhysique15, photoFiles = [] }) => withSync(async () => {
    // Le Stock théorique reste toujours un calcul PUR (Réceptions + Retours − Ventes), jamais
    // recalé automatiquement par l'écart mesuré : aucun mouvement "Ajustement" n'est créé ici.
    // Le Gain/Perte (physique − théorique) est uniquement enregistré et affiché, sans jamais
    // modifier le calcul du théorique des jours suivants.
    const theoriqueAmbiant = stockOf(siteId, product);
    const theorique15 = stockOf15(siteId, product);
    const basisEcart = "ambiant";
    const theoriqueUsed = theoriqueAmbiant;
    const physiqueUsed = stockPhysique;
    const ecart = physiqueUsed - theoriqueUsed;
    const cls = classifyEcart(ecart, theoriqueUsed, settings.objectifFreinte);
    const has15 = stockPhysique15 !== undefined;
    const vcfFields = has15 ? { temperatureC, densiteObservee, densite15, vcf, stockPhysique15 } : {};
    const offline = !navigator.onLine;
    const photoUrls = (!offline && photoFiles.length) ? await uploadPhotos(photoFiles, `inventaires/${siteId}`) : [];
    const invDraft = {
      siteId, product, date, stockPhysique, commentaire, basisEcart,
      stockTheoriqueAmbiant: theoriqueAmbiant, stockTheorique15: theorique15,
      stockTheorique: theoriqueUsed, stockPhysiqueUsed: physiqueUsed,
      ecart: cls.ecartL, ecartPermille: cls.ecartPermille, nature: cls.nature,
      tauxFreinte: cls.tauxFreinte, objectifUtilise: cls.objectif, conformite: cls.conformite,
      adjustmentId: null, photoUrls, createdBy: currentUserName, createdAt: new Date().toISOString(), ...vcfFields,
    };
    const row = inventaireToRow(invDraft);
    try {
      const { data: dataI, error: e2 } = await supabase.from("inventaires").insert(row).select().maybeSingle();
      if (e2) throw e2;
      if (!dataI) throw new Error("L'inventaire n'a pas pu être confirmé par le serveur — réessaie.");
      const invRecord = rowToInventaire(dataI);
      setInventaires((prev) => [...prev, invRecord]);
      appendAudit("Inventaire", `${sites.find((s) => s.id === siteId)?.name || ""} — base ${has15 ? "15°C" : "ambiante"} — ${NATURE_META[cls.nature].label} ${ecart >= 0 ? "+" : ""}${fmt(ecart)} L (${cls.ecartPermille >= 0 ? "+" : ""}${cls.ecartPermille.toFixed(2)} ‰)`);
      flash("Inventaire enregistré.");
    } catch (e) {
      if (!isNetworkError(e)) throw e;
      const q = readOfflineQueue();
      q.push({ id: uid(), table: "inventaires", row, createdAt: new Date().toISOString() });
      writeOfflineQueue(q);
      setOfflineQueueCount(q.length);
      setInventaires((prev) => [...prev, { id: uid(), ...invDraft }]);
      flash(photoFiles.length
        ? "Pas de connexion — inventaire enregistré sur l'appareil SANS les photos (à rajouter une fois reconnecté). Synchronisation automatique dès le retour du réseau."
        : "Pas de connexion — inventaire enregistré sur l'appareil, sera synchronisé automatiquement dès le retour du réseau.");
    }
  });
  const deleteInventaire = (inv) => withSync(async () => {
    const { data, error: e1 } = await supabase.from("inventaires").delete().eq("id", inv.id).select();
    if (e1) throw e1;
    if (!data || data.length === 0) throw new Error("Suppression refusée par la base de données (droits insuffisants ou ligne déjà supprimée) — l'inventaire n'a pas été retiré.");
    if (inv.adjustmentId) await supabase.from("movements").delete().eq("id", inv.adjustmentId);
    setInventaires((prev) => prev.filter((i) => i.id !== inv.id));
    setMovements((prev) => prev.filter((m) => m.id !== inv.adjustmentId));
    appendAudit("Suppression inventaire", `${sites.find((s) => s.id === inv.siteId)?.name || ""} — ${inv.date}`);
  });
  const updateSettings = (patch) => withSync(async () => {
    const next = { ...settings, ...patch };
    const { error } = await supabase.from("settings").update({
      objectif_freinte: next.objectifFreinte, logo_url: next.logoUrl, color_primary: next.colorPrimary, color_accent: next.colorAccent,
    }).eq("id", 1);
    if (error) throw error;
    setSettings(next);
    appendAudit("Modification des réglages", patch.objectifFreinte !== undefined ? `Nouvel objectif : ${next.objectifFreinte} ‰` : "Personnalisation (logo/couleurs)");
    flash("Réglages mis à jour.");
  });
  const updateTheme = ({ logoFile, logoTotalFile, colorPrimary, colorAccent }) => withSync(async () => {
    let logoUrl = settings.logoUrl;
    let logoTotalUrl = settings.logoTotalUrl;
    if (logoFile) {
      const urls = await uploadPhotos([logoFile], "branding");
      logoUrl = urls[0];
    }
    if (logoTotalFile) {
      const urls2 = await uploadPhotos([logoTotalFile], "branding-total");
      logoTotalUrl = urls2[0];
    }
    const next = { ...settings, logoUrl, logoTotalUrl, colorPrimary: colorPrimary || settings.colorPrimary, colorAccent: colorAccent || settings.colorAccent };
    const { error } = await supabase.from("settings").update({
      logo_url: next.logoUrl, logo_total_url: next.logoTotalUrl, color_primary: next.colorPrimary, color_accent: next.colorAccent,
    }).eq("id", 1);
    if (error) throw error;
    setSettings(next);
    appendAudit("Personnalisation", "Logo et/ou couleurs mis à jour");
    flash("Personnalisation enregistrée.");
  });
  useEffect(() => { applyTheme(settings.colorPrimary, settings.colorAccent); setCurrentLogoUrl(settings.logoUrl); setCurrentLogoTotalUrl(settings.logoTotalUrl); }, [settings.colorPrimary, settings.colorAccent, settings.logoUrl, settings.logoTotalUrl]);

  /* ---- mutations : rôle d'un utilisateur (Superviseur uniquement) ---- */
  const updateUserRole = (userId, role) => withSync(async () => {
    const { error } = await supabase.from("profiles").update({ role }).eq("id", userId);
    if (error) throw error;
    const target = profiles.find((u) => u.id === userId);
    setProfiles((prev) => prev.map((u) => (u.id === userId ? { ...u, role } : u)));
    appendAudit("Modification rôle utilisateur", `${target?.name || ""} → ${ROLE_LABELS[role]}`);
    flash("Rôle mis à jour.");
  });
  const updateUserSites = (userId, assignedSiteIds) => withSync(async () => {
    const { error } = await supabase.from("profiles").update({ assigned_site_ids: assignedSiteIds || [] }).eq("id", userId);
    if (error) throw error;
    const target = profiles.find((u) => u.id === userId);
    setProfiles((prev) => prev.map((u) => (u.id === userId ? { ...u, assignedSiteIds: assignedSiteIds || [] } : u)));
    const label = !assignedSiteIds || assignedSiteIds.length === 0 ? "Tous les sites" : assignedSiteIds.map((id) => sites.find((s) => s.id === id)?.name || id).join(", ");
    appendAudit("Modification sites assignés", `${target?.name || ""} → ${label}`);
    flash("Sites assignés mis à jour.");
  });
  const toggleUserActive = (userId, active) => withSync(async () => {
    const { error } = await supabase.from("profiles").update({ active }).eq("id", userId);
    if (error) throw error;
    const target = profiles.find((u) => u.id === userId);
    setProfiles((prev) => prev.map((u) => (u.id === userId ? { ...u, active } : u)));
    appendAudit(active ? "Réactivation compte" : "Désactivation compte", target?.name || "");
    flash(active ? "Compte réactivé." : "Compte désactivé.");
  });
  const toggleSiteActive = (siteId, active) => withSync(async () => {
    const { error } = await supabase.from("sites").update({ active }).eq("id", siteId);
    if (error) throw error;
    const target = sites.find((s) => s.id === siteId);
    setSites((prev) => prev.map((s) => (s.id === siteId ? { ...s, active } : s)));
    appendAudit(active ? "Réactivation site" : "Désactivation site", target?.name || "");
    flash(active ? "Site réactivé." : "Site désactivé.");
  });

  const NAV = [
    { id: "dashboard", label: "Tableau de bord", icon: LayoutDashboard, show: !perms.isTotalEnergiesOnly },
    { id: "sites", label: "Sites", icon: Factory, show: perms.canManage },
    { id: "saisie", label: "Saisie journalière", icon: ClipboardList, show: !perms.isTotalEnergiesOnly },
    { id: "inventaires", label: "Inventaires", icon: ClipboardList, show: true },
    { id: "vcf", label: "Correction 15°C", icon: Thermometer, show: !perms.isTotalEnergiesOnly },
    { id: "rapports", label: "Rapports", icon: FileBarChart, show: !perms.isTotalEnergiesOnly },
    { id: "utilisateurs", label: "Utilisateurs", icon: Users, show: perms.canManage },
    { id: "personnalisation", label: "Personnalisation", icon: Palette, show: perms.canManage },
    { id: "historique", label: "Historique", icon: History, show: perms.canManage },
  ].filter((n) => n.show);
  const viewTitle = NAV.find((n) => n.id === view)?.label || "";
  useEffect(() => { if (perms.isTotalEnergiesOnly && view === "dashboard") setView("inventaires"); }, [perms.isTotalEnergiesOnly]); // eslint-disable-line react-hooks/exhaustive-deps

  if (!SUPABASE_CONFIGURED) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100%", minHeight: 480, background: C.bg, fontFamily: "Inter, sans-serif", padding: 24, textAlign: "center" }}>
        <div>
          <CloudOff size={28} color={C.warning} />
          <h2 style={{ margin: "12px 0 6px" }}>Configuration manquante</h2>
          <p style={{ color: C.sub, maxWidth: 420 }}>
            Les clés Supabase (VITE_SUPABASE_URL et VITE_SUPABASE_ANON_KEY) ne sont pas définies. Ajoute-les dans les variables d'environnement du projet puis redéploie.
          </p>
        </div>
      </div>
    );
  }

  if (authLoading) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100%", minHeight: 480, background: C.bg }}>
        <Loader2 size={22} style={{ animation: "somipSpin .8s linear infinite" }} color={C.sub} />
        <style>{`@keyframes somipSpin { to { transform: rotate(360deg); } }`}</style>
      </div>
    );
  }

  if (!session) return <AuthScreen />;

  if (loadError) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100%", minHeight: 480, background: C.bg, fontFamily: "Inter, sans-serif", padding: 24 }}>
        <div style={{ textAlign: "center", maxWidth: 420 }}>
          <AlertCircle size={26} color={C.danger} />
          <h2 style={{ margin: "12px 0 6px", fontSize: 16 }}>Impossible de charger les données</h2>
          <p style={{ color: C.sub, fontSize: 13, marginBottom: 16 }}>{loadError}</p>
          <button className="somip-btn somip-btn-primary" onClick={() => { setLoading(true); setLoadError(null); setRetryKey((k) => k + 1); }}>
            <RotateCcw size={15} /> Réessayer
          </button>
        </div>
      </div>
    );
  }

  if (!profile || loading) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100%", minHeight: 480, background: C.bg, fontFamily: "Inter, sans-serif" }}>
        <div style={{ textAlign: "center", color: C.sub }}>
          <Loader2 size={22} style={{ animation: "somipSpin .8s linear infinite" }} />
          <div style={{ marginTop: 10, fontSize: 13, fontWeight: 600 }}>Chargement des données SOMIP...</div>
        </div>
        <style>{`@keyframes somipSpin { to { transform: rotate(360deg); } }`}</style>
      </div>
    );
  }

  return (
    <div className="somip-app" style={{ display: "flex", height: "100%", minHeight: 640, background: C.bg }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
        .somip-app * { box-sizing: border-box; }
        .somip-app { font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; color: ${C.ink}; }
        .somip-mono { font-family: 'IBM Plex Mono', 'SFMono-Regular', Menlo, Consolas, monospace; font-variant-numeric: tabular-nums; }
        .somip-scroll { overflow-y: auto; }
        .somip-scroll::-webkit-scrollbar { width: 8px; height: 8px; }
        .somip-scroll::-webkit-scrollbar-thumb { background: #C7CED6; border-radius: 4px; }
        .somip-nav-item { display:flex; align-items:center; gap:10px; padding:9px 14px; border-radius:8px; color:#AEBBC8; cursor:pointer; font-size:13px; font-weight:500; transition: background .15s, color .15s; border:none; background:transparent; width:100%; text-align:left; }
        .somip-nav-item:hover { background: rgba(255,255,255,0.07); color:#fff; }
        .somip-nav-item.active { background:${C.blue}; color:#fff; }
        .somip-btn { display:inline-flex; align-items:center; gap:6px; padding:9px 16px; border-radius:7px; font-size:13.5px; font-weight:600; cursor:pointer; border:1px solid transparent; transition:opacity .15s, background .15s; }
        .somip-btn:hover { opacity:0.9; }
        .somip-btn:disabled { opacity:0.45; cursor:not-allowed; }
        .somip-btn-primary { background:${C.blue}; color:#fff; }
        .somip-btn-secondary { background:${C.orange}; color:#fff; }
        .somip-btn-ghost { background:#fff; color:${C.ink}; border-color:${C.border}; }
        .somip-btn-ghost:hover { background:${C.bg}; opacity:1; }
        .somip-input, .somip-select, .somip-textarea { width:100%; padding:9px 11px; border-radius:7px; border:1px solid ${C.border}; font-size:13.5px; font-family:inherit; color:${C.ink}; background:#fff; }
        .somip-input:focus, .somip-select:focus, .somip-textarea:focus { outline:none; border-color:${C.blue}; box-shadow:0 0 0 3px rgba(0,113,189,0.12); }
        .somip-label { font-size:12px; font-weight:600; color:${C.sub}; margin-bottom:5px; display:block; }
        .somip-table { width:100%; border-collapse:collapse; }
        .somip-table th { text-align:left; font-size:11px; font-weight:600; color:${C.sub}; padding:9px 12px; border-bottom:1px solid ${C.border}; white-space:nowrap; }
        .somip-table td { padding:11px 12px; font-size:13px; border-bottom:1px solid #EEF1F3; }
        .somip-table tr:hover td { background:#FAFBFC; }
        .somip-panel { background:#fff; border:1px solid ${C.border}; border-radius:10px; }
        .somip-tab { padding:8px 16px; border-radius:7px; font-size:13px; font-weight:600; cursor:pointer; border:1px solid ${C.border}; background:#fff; color:${C.sub}; }
        .somip-tab.active { background:${C.ink}; color:#fff; border-color:${C.ink}; }
        .somip-fade { animation: somipFade .2s ease; }
        @keyframes somipFade { from { opacity:0; transform:translateY(3px);} to {opacity:1; transform:none;} }
        @keyframes somipSpin { to { transform: rotate(360deg); } }
        .somip-print-only { display: none; }
        .somip-mobile-toggle { display: none; }
        .somip-mobile-backdrop { position: fixed; inset: 0; background: rgba(10,20,30,0.5); z-index: 35; }
        @media (max-width: 860px) {
          .somip-mobile-toggle { display: inline-flex !important; }
          .somip-sidebar { position: fixed !important; top: 0; left: 0; bottom: 0; z-index: 40; transform: translateX(-105%); transition: transform .22s ease; box-shadow: 6px 0 28px rgba(0,0,0,0.28); }
          .somip-sidebar.open { transform: translateX(0); }
          .somip-header { padding: 12px 14px !important; }
          .somip-scroll { padding: 14px !important; }
        }
        @media print {
          .somip-no-print, .somip-sidebar, .somip-header { display: none !important; }
          .somip-print-only { display: block !important; }
          .somip-scroll { overflow: visible !important; height: auto !important; padding: 0 !important; }
          body, .somip-app { background: #fff !important; }
          .somip-panel { border: none !important; }
        }
      `}</style>

      {mobileNavOpen && <div className="somip-mobile-backdrop" onClick={() => setMobileNavOpen(false)} />}

      {/* Sidebar */}
      <aside className={`somip-sidebar ${mobileNavOpen ? "open" : ""}`} style={{ width: 226, background: `linear-gradient(180deg, ${C.navy}, ${C.navyLight})`, display: "flex", flexDirection: "column", padding: "20px 14px", flexShrink: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 9, padding: "0 6px 22px" }}>
          {settings.logoUrl ? (
            <img src={settings.logoUrl} alt="Logo" style={{ width: 32, height: 32, borderRadius: 8, objectFit: "cover" }} />
          ) : (
            <div style={{ width: 32, height: 32, borderRadius: 8, background: C.blue, display: "flex", alignItems: "center", justifyContent: "center" }}>
              <Fuel size={17} color="#fff" />
            </div>
          )}
          <div>
            <div style={{ color: "#fff", fontWeight: 700, fontSize: 14.5, letterSpacing: 0.2 }}>SOMIP</div>
            <div style={{ color: "#8CA0B4", fontSize: 10.5, fontWeight: 500 }}>Stock Gasoil</div>
          </div>
          <button className="somip-mobile-toggle" onClick={() => setMobileNavOpen(false)} style={{ marginLeft: "auto", border: "none", background: "none", cursor: "pointer", color: "#fff" }}>
            <X size={20} />
          </button>
        </div>
        <nav style={{ display: "flex", flexDirection: "column", gap: 3 }}>
          {NAV.map((n) => (
            <button key={n.id} className={`somip-nav-item ${view === n.id ? "active" : ""}`} onClick={() => { setView(n.id); setMobileNavOpen(false); }}>
              <n.icon size={16} />{n.label}
            </button>
          ))}
        </nav>
        <div style={{ flex: 1 }} />
        <button className="somip-nav-item" onClick={signOut}>
          <LogOut size={16} /> Se déconnecter
        </button>
        <div style={{ color: "#5C7288", fontSize: 10.5, padding: "10px 6px 0" }}>Zone Sud-Est · Gabon</div>
      </aside>

      {/* Main */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", minWidth: 0 }}>
        <header className="somip-header" style={{ padding: "16px 28px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "space-between", background: "#fff", gap: 16, flexWrap: "wrap" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <button className="somip-mobile-toggle" onClick={() => setMobileNavOpen(true)} style={{ border: "none", background: "none", cursor: "pointer", padding: 4, color: C.ink }}>
              <Menu size={22} />
            </button>
            <div>
              <h1 style={{ margin: 0, fontSize: 18, fontWeight: 700 }}>{viewTitle}</h1>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <SyncIndicator status={syncStatus} lastSync={lastSync} />
                {offlineQueueCount > 0 && (
                  <span style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 11, fontWeight: 700, color: C.warning, background: "#FFF6E5", padding: "2px 8px", borderRadius: 12 }}>
                    <CloudOff size={11} /> {offlineQueueCount} en attente
                  </span>
                )}
              </div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <div style={{ fontSize: 12, color: C.sub, textAlign: "right" }}>
              <div style={{ fontWeight: 700, color: C.ink }}>{currentUserName}</div>
              <Badge color={C.blue}>{ROLE_LABELS[currentRole]}</Badge>
            </div>
            <div style={{ fontSize: 12.5, color: C.sub, textAlign: "right" }}>
              {new Date().toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long", year: "numeric" })}
            </div>
          </div>
        </header>

        {notice && (
          <div className="somip-no-print" style={{
            display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12,
            padding: "10px 20px", fontSize: 13, fontWeight: 600,
            background: noticeType === "error" ? "#FCEAEA" : "#E9F7EF",
            color: noticeType === "error" ? C.danger : C.success,
            borderBottom: `1px solid ${noticeType === "error" ? "#F3C6C6" : "#C9EBD7"}`,
          }}>
            <span>{notice}</span>
            <button onClick={() => setNotice(null)} style={{ border: "none", background: "none", cursor: "pointer", color: "inherit", flexShrink: 0 }}>
              <X size={16} />
            </button>
          </div>
        )}

        <div className="somip-scroll" style={{ flex: 1, padding: "24px 28px" }}>
          {view === "dashboard" && <Dashboard sites={sites} movements={movements} inventaires={inventaires} stockOf={stockOf} purgeDemoMovements={purgeDemoMovements} canManage={perms.canManage} truckAssignments={truckAssignments} />}
          {view === "sites" && perms.canManage && <SitesView sites={sites} movements={movements} stockOf={stockOf} addSite={addSite} editSite={editSite} removeSite={removeSite} toggleSiteActive={toggleSiteActive} productStocks={productStocks} saveProductStock={saveProductStock} truckAssignments={truckAssignments} assignTruck={assignTruck} siteMeters={siteMeters} addSiteMeter={addSiteMeter} removeSiteMeter={removeSiteMeter} siteTanks={siteTanks} addSiteTank={addSiteTank} removeSiteTank={removeSiteTank} siteDepotageMeters={siteDepotageMeters} addSiteDepotageMeter={addSiteDepotageMeter} removeSiteDepotageMeter={removeSiteDepotageMeter} />}
          {view === "saisie" && <DailyEntryView sites={sites} movements={movements} inventaires={inventaires} productStocks={productStocks} siteMeters={siteMeters} saveProductStock={saveProductStock} addMovement={addMovement} addInventaire={addInventaire} deleteMovement={deleteMovement} deleteInventaire={deleteInventaire} settings={settings} canWrite={perms.canWrite} canManage={perms.canManage} assignedSiteIds={profile?.assignedSiteIds} truckAssignments={truckAssignments} />}
          {view === "inventaires" && <InventairesView sites={sites} inventaires={inventaires} stockOf={stockOf} stockOf15={stockOf15} addInventaire={addInventaire} deleteInventaire={deleteInventaire} settings={settings} updateSettings={updateSettings} canWrite={perms.canWrite} canManage={perms.canManage} canInventaireOfficiel={perms.canInventaireOfficiel} inventairesOfficiels={inventairesOfficiels} addInventaireOfficiel={addInventaireOfficiel} deleteInventaireOfficiel={deleteInventaireOfficiel} siteTanks={siteTanks} siteDepotageMeters={siteDepotageMeters} siteMeters={siteMeters} signInventaireOfficiel={signInventaireOfficiel} canSignSomip={perms.canSignSomip} canSignOperateur={perms.canSignOperateur} canSignTotal={perms.canSignTotal} isTotalEnergiesOnly={perms.isTotalEnergiesOnly} />}
          {view === "vcf" && <VcfView />}
          {view === "rapports" && <ReportsView sites={sites} movements={movements} inventaires={inventaires} productStocks={productStocks} truckAssignments={truckAssignments} settings={settings} stockOf={stockOf} bilans={bilans} saveBilan={saveBilan} deleteBilan={deleteBilan} canManage={perms.canManage} />}
          {view === "utilisateurs" && perms.canManage && <UsersView profiles={profiles} updateUserRole={updateUserRole} updateUserSites={updateUserSites} toggleUserActive={toggleUserActive} sites={sites} session={session} />}
          {view === "personnalisation" && perms.canManage && <BrandingView settings={settings} updateTheme={updateTheme} />}
          {view === "historique" && perms.canManage && <HistoryView audit={audit} />}
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Dashboard                                                            */
/* ------------------------------------------------------------------ */
function Dashboard({ sites, movements, inventaires, stockOf, purgeDemoMovements, canManage, truckAssignments }) {
  const month = currentMonth();
  const rows = sites.map((s) => {
    const stock = stockOf(s.id);
    const pct = s.capacity ? (stock / s.capacity) * 100 : 0;
    const status = pct < 20 ? "danger" : pct < 35 ? "warning" : "ok";
    return { ...s, stock, pct, status };
  });
  const totalStock = rows.reduce((a, r) => a + r.stock, 0);
  const totalCapacity = rows.reduce((a, r) => a + r.capacity, 0);
  const alerts = rows.filter((r) => r.status !== "ok");
  const receptionsMonth = movements.filter((m) => m.type === "reception" && (m.product || "gasoil") === "gasoil" && m.date.startsWith(month)).reduce((a, m) => a + m.quantity, 0);
  const sortiesMonth = movements.filter((m) => (m.type === "sortie" || m.type === "sortie_camion") && (m.product || "gasoil") === "gasoil" && m.date.startsWith(month)).reduce((a, m) => a + m.quantity, 0);
  const recent = [...movements].sort((a, b) => (a.date < b.date ? 1 : -1)).slice(0, 8);
  const demoCount = movements.filter((m) => m.isDemo).length;
  const statusColor = { ok: C.blue, warning: C.warning, danger: C.danger };

  const latestInventaireBySite = {};
  [...inventaires].sort((a, b) => (a.date < b.date ? -1 : 1)).forEach((i) => { latestInventaireBySite[i.siteId] = i; });
  const horsObjectif = Object.values(latestInventaireBySite).filter((i) => i.conformite === "non_conforme").length;

  // Gain/Perte cumulé du mois en cours, pour chaque site fixe ET chaque camion (calcul propre
  // à chacun, indépendant — comme demandé, tout dans un même tableau).
  const monthStartD = `${month}-01`;
  const todayD = todayStr();
  const bigLosses = [];
  const ecartRows = sites.map((s) => {
    let cur = new Date(monthStartD);
    const end = new Date(todayD);
    let ecartCumule = 0, daysWithJauge = 0;
    while (cur <= end) {
      const d = `${cur.getFullYear()}-${pad2(cur.getMonth() + 1)}-${pad2(cur.getDate())}`;
      const stockDebut = stockBeforeDate(s, movements, d, inventaires);
      const dayMovs = movements.filter((m) => m.siteId === s.id && (m.product || "gasoil") === "gasoil" && m.date === d);
      const reception = sumQty(dayMovs, ["reception"]);
      const ventes = sumQty(dayMovs, ["sortie"]); // camion : flux brut du compteur (retour cuve inclus, cf. théorique)
      const chargementLaitiers = s.isMobile ? 0 : sumQty(dayMovs, ["sortie_camion"]);
      const retourCamions = s.isMobile ? 0 : sumQty(dayMovs, ["retour_camion"]);
      const theorique = stockDebut + reception + retourCamions - ventes - chargementLaitiers;
      const inv = pickLatestInv(inventaires.filter((i) => i.siteId === s.id && (i.product || "gasoil") === "gasoil" && i.date === d));
      if (inv) {
        const dayEcart = inv.stockPhysique - theorique;
        ecartCumule += dayEcart; daysWithJauge++;
        // Alerte perte > 500 L en une seule journée (site ou camion).
        if (dayEcart <= -500) bigLosses.push({ site: s, date: d, ecart: dayEcart });
      }
      cur.setDate(cur.getDate() + 1);
    }
    return { site: s, ecartCumule, daysWithJauge };
  }).sort((a, b) => a.ecartCumule - b.ecartCumule);
  const ecartReseauTotal = ecartRows.filter((r) => !r.site.isMobile).reduce((a, r) => a + r.ecartCumule, 0);
  bigLosses.sort((a, b) => a.ecart - b.ecart);

  // Alerte saisie manquante : à partir de 6h00, signale les sites (fixes et camions) sans
  // Stock fin saisi pour la veille — sauf un camion dont le dernier stock connu était à zéro
  // (pas en service, pas d'alerte à lui envoyer).
  const now = new Date();
  const yesterday = new Date(now); yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayStr = `${yesterday.getFullYear()}-${pad2(yesterday.getMonth() + 1)}-${pad2(yesterday.getDate())}`;
  const missingSites = now.getHours() >= 6
    ? sites.filter((s) => {
        const hasYesterday = inventaires.some((i) => i.siteId === s.id && (i.product || "gasoil") === "gasoil" && i.date === yesterdayStr);
        if (hasYesterday) return false;
        if (s.isMobile) {
          const lastInv = pickLatestInv(inventaires.filter((i) => i.siteId === s.id && (i.product || "gasoil") === "gasoil"));
          // Pas d'alerte si le dernier stock connu était à zéro, ou si ce camion n'a jamais eu
          // de jauge enregistrée du tout (même situation : camion hors service).
          if (!lastInv || Number(lastInv.stockPhysique) === 0) return false;
        }
        return true;
      })
    : [];

  return (
    <div className="somip-fade">
      <p style={{ marginTop: -8, marginBottom: 14, fontSize: 13, color: C.sub }}>
        Vos données sont sauvegardées automatiquement et restent disponibles après fermeture ou actualisation de la page.
        Les capacités et stocks initiaux des sites restent des valeurs à vérifier/ajuster depuis la page Sites.
      </p>

      {missingSites.length > 0 && (
        <div className="somip-panel" style={{ padding: "12px 16px", marginBottom: 18, borderLeft: `3px solid ${C.danger}` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
            <AlertTriangle size={16} color={C.danger} />
            <strong style={{ fontSize: 13, color: C.ink }}>Saisie du {yesterdayStr} manquante</strong>
          </div>
          <p style={{ margin: 0, fontSize: 12.5, color: C.sub }}>
            Stock fin non saisi pour : <strong style={{ color: C.ink }}>{missingSites.map((s) => s.name).join(", ")}</strong>.
          </p>
        </div>
      )}

      {demoCount > 0 && canManage && (
        <div className="somip-panel" style={{ padding: "12px 16px", marginBottom: 18, display: "flex", alignItems: "center", justifyContent: "space-between", borderLeft: `3px solid ${C.warning}`, flexWrap: "wrap", gap: 10 }}>
          <span style={{ fontSize: 12.5, color: C.sub }}>
            <strong style={{ color: C.ink }}>{demoCount} écriture(s)</strong> marquée(s) <DemoBadge /> sont encore présentes (jeu d'exemple de la V1).
          </span>
          <ConfirmTextButton onConfirm={purgeDemoMovements} label="Supprimer les écritures d'exemple" confirmLabel="Confirmer la suppression ?" className="somip-btn somip-btn-ghost" />
        </div>
      )}

      <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 22 }}>
        <StatCard label="Stock total réseau" value={fmt(totalStock)} unit="L" accent={C.blue} icon={Fuel} />
        <StatCard label="Capacité totale" value={fmt(totalCapacity)} unit="L" accent={C.sub} icon={Factory} />
        <StatCard label="Réceptions (mois)" value={fmt(receptionsMonth)} unit="L" accent={C.success} icon={ArrowDownCircle} />
        <StatCard label="Sorties (mois)" value={fmt(sortiesMonth)} unit="L" accent={C.orange} icon={ArrowUpCircle} />
        <StatCard label="Sites en alerte" value={alerts.length} unit={`/ ${rows.length}`} accent={C.danger} icon={AlertTriangle} />
        <StatCard label="Sites hors objectif freinte" value={horsObjectif} unit={`/ ${rows.length}`} accent={C.warning} icon={ClipboardList} />
        <StatCard label="Gain/Perte réseau (mois)" value={`${ecartReseauTotal >= 0 ? "+" : ""}${fmt(ecartReseauTotal)}`} unit="L" accent={ecartReseauTotal < 0 ? C.danger : ecartReseauTotal > 0 ? C.success : C.sub} icon={TrendingDown} />
      </div>

      {bigLosses.length > 0 && (
        <div className="somip-panel" style={{ marginBottom: 18, padding: 18, borderLeft: `4px solid ${C.danger}` }}>
          <h3 style={{ margin: "0 0 4px", fontSize: 14, color: C.danger, display: "flex", alignItems: "center", gap: 8 }}>
            <AlertTriangle size={17} /> Perte de plus de 500 L en une journée
          </h3>
          <p style={{ margin: "0 0 12px", fontSize: 12, color: C.sub }}>Mois en cours — à vérifier en priorité.</p>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {bigLosses.map((b, i) => (
              <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: C.bg, borderRadius: 6, padding: "8px 12px", fontSize: 13 }}>
                <span><strong>{b.site.name}</strong> {!b.site.isMobile && <span style={{ color: C.sub }}>({b.site.code})</span>} — {b.date}</span>
                <span className="somip-mono" style={{ fontWeight: 700, color: C.danger }}>{fmt(b.ecart)} L</span>
              </div>
            ))}
          </div>
        </div>
      )}

      <div style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
        <div className="somip-panel" style={{ flex: "1 1 380px", padding: 18 }}>
          <h3 style={{ margin: "0 0 14px", fontSize: 14 }}>Niveau de stock par site</h3>
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            {rows.map((r) => (
              <div key={r.id}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5, fontSize: 12.5 }}>
                  <span style={{ fontWeight: 600 }}>{r.name} <span style={{ color: C.sub, fontWeight: 500 }}>({r.code})</span></span>
                  <span className="somip-mono" style={{ color: C.sub }}>{fmt(r.stock)} / {fmt(r.capacity)} L</span>
                </div>
                <GaugeBar pct={r.pct} color={statusColor[r.status]} />
              </div>
            ))}
          </div>
        </div>

        <div className="somip-panel" style={{ flex: "1 1 380px", padding: 18, minHeight: 320 }}>
          <h3 style={{ margin: "0 0 14px", fontSize: 14 }}>Stock actuel par site (L)</h3>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={rows} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F3" vertical={false} />
              <XAxis dataKey="code" tick={{ fontSize: 11, fill: C.sub }} axisLine={{ stroke: C.border }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: C.sub }} axisLine={false} tickLine={false} />
              <Tooltip formatter={(v) => [`${fmt(v)} L`, "Stock actuel"]} labelFormatter={(l, p) => (p && p[0] ? p[0].payload.name : l)} contentStyle={{ fontSize: 12.5, borderRadius: 8, border: `1px solid ${C.border}` }} />
              <Bar dataKey="stock" radius={[5, 5, 0, 0]}>
                {rows.map((r) => <Cell key={r.id} fill={statusColor[r.status]} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="somip-panel" style={{ marginTop: 18, padding: 18 }}>
        <h3 style={{ margin: "0 0 12px", fontSize: 14 }}>Derniers mouvements</h3>
        <table className="somip-table">
          <thead><tr><th>Date</th><th>Site</th><th>Type</th><th style={{ textAlign: "right" }}>Quantité</th></tr></thead>
          <tbody>
            {recent.length === 0 && <EmptyRow colSpan={4} text="Aucun mouvement enregistré." />}
            {recent.map((m) => {
              const site = sites.find((s) => s.id === m.siteId);
              const meta = TYPE_META[m.type];
              return (
                <tr key={m.id}>
                  <td className="somip-mono">{m.date}</td>
                  <td>{site?.name || "—"}</td>
                  <td><Badge color={meta.color}>{meta.label}</Badge>{m.isDemo && <DemoBadge />}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: meta.color, fontWeight: 600 }}>{meta.sign} {fmt(m.quantity)} L</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Sites                                                                 */
/* ------------------------------------------------------------------ */
function SitesView({ sites, movements, stockOf, addSite, editSite, removeSite, toggleSiteActive, productStocks, saveProductStock, truckAssignments, assignTruck, siteMeters, addSiteMeter, removeSiteMeter, siteTanks, addSiteTank, removeSiteTank, siteDepotageMeters, addSiteDepotageMeter, removeSiteDepotageMeter }) {
  const [form, setForm] = useState({ name: "", code: "", capacity: "", stockInitial: "", isMobile: false });
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [lubSiteId, setLubSiteId] = useState(LUBRICANT_SITE_IDS[0]);
  const [lubProduct, setLubProduct] = useState(LUBRICANTS[0].id);
  const [lubForm, setLubForm] = useState({ capacity: "", stockInitial: "" });
  const trucks = sites.filter((s) => s.isMobile);
  const stations = sites.filter((s) => !s.isMobile);
  const [assignForm, setAssignForm] = useState({ truckId: trucks[0]?.id || "", stationId: stations[0]?.id || "", startDate: todayStr() });
  const [meterSiteId, setMeterSiteId] = useState(stations[0]?.id || "");
  const [newMeterName, setNewMeterName] = useState("");
  const [tankSiteId, setTankSiteId] = useState(stations[0]?.id || "");
  const [newTankName, setNewTankName] = useState("");
  const [depotageSiteId, setDepotageSiteId] = useState(stations[0]?.id || "");
  const [newDepotageName, setNewDepotageName] = useState("");

  const submitAdd = () => {
    if (!form.name.trim() || !form.code.trim() || !form.capacity) return;
    addSite(form);
    setForm({ name: "", code: "", capacity: "", stockInitial: "", isMobile: false });
  };
  const startEdit = (s) => { setEditingId(s.id); setEditForm({ ...s }); };
  const saveEdit = () => { editSite(editingId, editForm); setEditingId(null); };

  const currentLubStock = productStocks.find((p) => p.siteId === lubSiteId && p.product === lubProduct);
  const loadLubForEdit = (siteId, product) => {
    setLubSiteId(siteId); setLubProduct(product);
    const ps = productStocks.find((p) => p.siteId === siteId && p.product === product);
    setLubForm({ capacity: ps ? String(ps.capacity) : "", stockInitial: ps ? String(ps.stockInitial) : "" });
  };
  const submitLub = () => {
    if (!lubForm.capacity) return;
    saveProductStock({ siteId: lubSiteId, product: lubProduct, capacity: lubForm.capacity, stockInitial: lubForm.stockInitial });
  };

  const submitAssign = () => {
    if (!assignForm.truckId || !assignForm.stationId || !assignForm.startDate) return;
    assignTruck(assignForm);
  };
  const assignmentsSorted = [...truckAssignments].sort((a, b) => (a.startDate < b.startDate ? 1 : -1));

  const currentMeters = siteMeters.filter((m) => m.siteId === meterSiteId);
  const displayedMeters = currentMeters.length ? currentMeters : metersForSite(sites.find((s) => s.id === meterSiteId)).map((name, i) => ({ id: `default-${i}`, name, isDefault: true }));
  const submitMeter = () => {
    if (!newMeterName.trim() || !meterSiteId) return;
    addSiteMeter({ siteId: meterSiteId, name: newMeterName });
    setNewMeterName("");
  };
  const currentTanks = siteTanks.filter((t) => t.siteId === tankSiteId);
  const submitTank = () => {
    if (!newTankName.trim() || !tankSiteId) return;
    addSiteTank({ siteId: tankSiteId, name: newTankName });
    setNewTankName("");
  };
  const currentDepotageMeters = siteDepotageMeters.filter((m) => m.siteId === depotageSiteId);
  const submitDepotageMeter = () => {
    if (!newDepotageName.trim() || !depotageSiteId) return;
    addSiteDepotageMeter({ siteId: depotageSiteId, name: newDepotageName });
    setNewDepotageName("");
  };

  return (
    <div className="somip-fade" style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      <div className="somip-panel" style={{ flex: "2 1 520px", padding: 18 }}>
        <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Sites externalisés ({sites.length})</h3>
        <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.sub }}>Capacités et stocks initiaux : à vérifier et ajuster selon vos valeurs réelles.</p>
        <table className="somip-table">
          <thead><tr><th>Code</th><th>Site</th><th style={{ textAlign: "right" }}>Capacité (L)</th><th style={{ textAlign: "right" }}>Stock actuel (L)</th><th>Statut</th><th></th></tr></thead>
          <tbody>
            {sites.map((s) => {
              const stock = stockOf(s.id);
              const isEditing = editingId === s.id;
              return (
                <tr key={s.id}>
                  {isEditing ? (
                    <>
                      <td><input className="somip-input" style={{ width: 70 }} value={editForm.code} onChange={(e) => setEditForm({ ...editForm, code: e.target.value })} /></td>
                      <td><input className="somip-input" value={editForm.name} onChange={(e) => setEditForm({ ...editForm, name: e.target.value })} /></td>
                      <td><input type="number" className="somip-input" style={{ textAlign: "right" }} value={editForm.capacity} onChange={(e) => setEditForm({ ...editForm, capacity: e.target.value })} /></td>
                      <td style={{ textAlign: "right", color: C.sub, fontSize: 12 }}>calculé</td>
                      <td></td>
                      <td style={{ whiteSpace: "nowrap" }}>
                        <button className="somip-btn somip-btn-primary" style={{ padding: "5px 10px", fontSize: 12 }} onClick={saveEdit}>OK</button>
                        <button onClick={() => setEditingId(null)} style={{ border: "none", background: "none", cursor: "pointer", marginLeft: 4 }}><X size={16} color={C.sub} /></button>
                      </td>
                    </>
                  ) : (
                    <>
                      <td style={{ fontWeight: 700, color: C.blue }}>{s.code}</td>
                      <td>{s.name}{s.isMobile && <span style={{ marginLeft: 6 }}><Badge color={C.orange}>Camion</Badge></span>}</td>
                      <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(s.capacity)}</td>
                      <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600 }}>{fmt(stock)}</td>
                      <td>{s.active === false ? <Badge color={C.danger}>Désactivé</Badge> : <Badge color={C.success}>Actif</Badge>}</td>
                      <td style={{ whiteSpace: "nowrap", textAlign: "right" }}>
                        <button onClick={() => startEdit(s)} style={{ border: "none", background: "none", cursor: "pointer", padding: 5 }}><Pencil size={14} color={C.sub} /></button>
                        <button onClick={() => toggleSiteActive(s.id, s.active === false)} title={s.active === false ? "Réactiver" : "Désactiver"} style={{ border: "none", background: "none", cursor: "pointer", padding: 5 }}>
                          {s.active === false ? <CheckCircle2 size={14} color={C.success} /> : <CloudOff size={14} color={C.warning} />}
                        </button>
                        <ConfirmIconButton onConfirm={() => removeSite(s)} title="Supprimer le site" />
                      </td>
                    </>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div className="somip-panel" style={{ flex: "1 1 280px", padding: 18 }}>
        <h3 style={{ margin: "0 0 14px", fontSize: 14 }}>Ajouter un site</h3>
        <Field label="Nom du site"><input className="somip-input" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="Ex : Dépôt Moanda, ou FK253AA" /></Field>
        <Field label="Code (court)"><input className="somip-input" value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} placeholder="Ex : DPM, ou FK253AA" /></Field>
        <Field label="Capacité (L)"><input type="number" className="somip-input" value={form.capacity} onChange={(e) => setForm({ ...form, capacity: e.target.value })} placeholder="30000" /></Field>
        <Field label="Stock initial (L)"><input type="number" className="somip-input" value={form.stockInitial} onChange={(e) => setForm({ ...form, stockInitial: e.target.value })} placeholder="0" /></Field>
        <label style={{ display: "flex", alignItems: "center", gap: 8, margin: "4px 0 14px", fontSize: 12.5, cursor: "pointer" }}>
          <input type="checkbox" checked={form.isMobile} onChange={(e) => setForm({ ...form, isMobile: e.target.checked })} />
          Camion (station mobile) — ex : FK253AA
        </label>
        <button className="somip-btn somip-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={submitAdd} disabled={!form.name || !form.code || !form.capacity}>
          <Plus size={15} /> Ajouter le site
        </button>
      </div>

      <div className="somip-panel" style={{ flex: "1 1 280px", padding: 18 }}>
        <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Lubrifiants (Prehomo / Okouma)</h3>
        <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.sub }}>Capacité et stock initial par produit, en litres.</p>
        <Field label="Site">
          <select className="somip-select" value={lubSiteId} onChange={(e) => loadLubForEdit(e.target.value, lubProduct)}>
            {LUBRICANT_SITE_IDS.map((id) => <option key={id} value={id}>{sites.find((s) => s.id === id)?.name || id}</option>)}
          </select>
        </Field>
        <Field label="Produit">
          <select className="somip-select" value={lubProduct} onChange={(e) => loadLubForEdit(lubSiteId, e.target.value)}>
            {LUBRICANTS.map((l) => <option key={l.id} value={l.id}>{l.label} (densité {l.densite})</option>)}
          </select>
        </Field>
        <div style={{ display: "flex", gap: 8 }}>
          <div style={{ flex: 1 }}><Field label="Capacité (L)"><input type="number" className="somip-input" value={lubForm.capacity} onChange={(e) => setLubForm({ ...lubForm, capacity: e.target.value })} placeholder="Ex : 1000" /></Field></div>
          <div style={{ flex: 1 }}><Field label="Stock initial (L)"><input type="number" className="somip-input" value={lubForm.stockInitial} onChange={(e) => setLubForm({ ...lubForm, stockInitial: e.target.value })} placeholder="0" /></Field></div>
        </div>
        {currentLubStock && <p style={{ margin: "-6px 0 10px", fontSize: 11, color: C.sub }}>Déjà enregistré : capacité {fmt(currentLubStock.capacity)} L, stock initial {fmt(currentLubStock.stockInitial)} L.</p>}
        <button className="somip-btn somip-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={submitLub} disabled={!lubForm.capacity}>
          <Check size={15} /> Enregistrer
        </button>
      </div>

      {trucks.length > 0 && (
        <div className="somip-panel" style={{ flex: "1 1 320px", padding: 18 }}>
          <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Affectation des camions</h3>
          <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.sub }}>Quel camion travaille sur quelle station, avec l'historique des changements (panne, remplacement...).</p>
          <div style={{ display: "flex", gap: 8 }}>
            <div style={{ flex: 1 }}>
              <Field label="Camion">
                <select className="somip-select" value={assignForm.truckId} onChange={(e) => setAssignForm({ ...assignForm, truckId: e.target.value })}>
                  {trucks.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                </select>
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Station">
                <select className="somip-select" value={assignForm.stationId} onChange={(e) => setAssignForm({ ...assignForm, stationId: e.target.value })}>
                  {stations.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
              </Field>
            </div>
          </div>
          <Field label="Depuis le"><input type="date" className="somip-input" value={assignForm.startDate} onChange={(e) => setAssignForm({ ...assignForm, startDate: e.target.value })} /></Field>
          <button className="somip-btn somip-btn-primary" style={{ width: "100%", justifyContent: "center", marginBottom: 16 }} onClick={submitAssign}>
            <Check size={15} /> Affecter
          </button>
          <p style={{ margin: "0 0 8px", fontSize: 12, fontWeight: 700, color: C.ink }}>Historique</p>
          <table className="somip-table">
            <thead><tr><th>Camion</th><th>Station</th><th>Depuis</th><th>Jusqu'au</th></tr></thead>
            <tbody>
              {assignmentsSorted.length === 0 && <EmptyRow colSpan={4} text="Aucune affectation." />}
              {assignmentsSorted.map((a) => (
                <tr key={a.id}>
                  <td>{sites.find((s) => s.id === a.truckId)?.name || a.truckId}</td>
                  <td>{sites.find((s) => s.id === a.stationId)?.name || a.stationId}</td>
                  <td className="somip-mono">{a.startDate}</td>
                  <td className="somip-mono">{a.endDate || <Badge color={C.success}>en cours</Badge>}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <div className="somip-panel" style={{ flex: "1 1 300px", padding: 18 }}>
        <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Compteurs par site</h3>
        <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.sub }}>Ajoute ou retire un compteur si la configuration physique d'un site change. Sur un camion, ça sert aussi à définir des "postes" (ex. Poste 1 / Poste 2) quand la saisie se fait en deux temps dans la journée.</p>
        <Field label="Site">
          <select className="somip-select" value={meterSiteId} onChange={(e) => setMeterSiteId(e.target.value)}>
            <optgroup label="Sites fixes">
              {stations.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </optgroup>
            <optgroup label="Camions">
              {trucks.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
            </optgroup>
          </select>
        </Field>
        <table className="somip-table" style={{ marginBottom: 12 }}>
          <thead><tr><th>Compteur</th><th></th></tr></thead>
          <tbody>
            {displayedMeters.length === 0 && <EmptyRow colSpan={2} text="Aucun compteur." />}
            {displayedMeters.map((m) => (
              <tr key={m.id}>
                <td>{m.name}{m.isDefault && <span style={{ color: C.sub, fontSize: 11 }}> (par défaut)</span>}</td>
                <td style={{ textAlign: "right" }}>{!m.isDefault && <ConfirmIconButton onConfirm={() => removeSiteMeter(m)} />}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{ display: "flex", gap: 8 }}>
          <input className="somip-input" style={{ flex: 1 }} value={newMeterName} onChange={(e) => setNewMeterName(e.target.value)} placeholder="Ex : Compteur 3" />
          <button className="somip-btn somip-btn-primary" onClick={submitMeter} disabled={!newMeterName.trim()}><Plus size={15} /></button>
        </div>
        <p style={{ marginTop: 10, fontSize: 11, color: C.sub }}>
          Dès que tu ajoutes un premier compteur pour un site, la liste par défaut est remplacée par celle-ci — pense à recréer les compteurs existants si besoin.
        </p>
      </div>

      <div className="somip-panel" style={{ flex: "1 1 300px", padding: 18 }}>
        <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Cuves par site</h3>
        <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.sub }}>Pour les inventaires officiels : un site peut avoir plusieurs cuves physiques, chacune relevée séparément.</p>
        <Field label="Site">
          <select className="somip-select" value={tankSiteId} onChange={(e) => setTankSiteId(e.target.value)}>
            {stations.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </Field>
        <table className="somip-table" style={{ marginBottom: 12 }}>
          <thead><tr><th>Cuve</th><th></th></tr></thead>
          <tbody>
            {currentTanks.length === 0 && <EmptyRow colSpan={2} text="Aucune cuve configurée." />}
            {currentTanks.map((t) => (
              <tr key={t.id}>
                <td>{t.name}</td>
                <td style={{ textAlign: "right" }}><ConfirmIconButton onConfirm={() => removeSiteTank(t)} /></td>
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{ display: "flex", gap: 8 }}>
          <input className="somip-input" style={{ flex: 1 }} value={newTankName} onChange={(e) => setNewTankName(e.target.value)} placeholder="Ex : Cuve 1" />
          <button className="somip-btn somip-btn-primary" onClick={submitTank} disabled={!newTankName.trim()}><Plus size={15} /></button>
        </div>
      </div>

      <div className="somip-panel" style={{ flex: "1 1 300px", padding: 18 }}>
        <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Compteurs de dépotage par site</h3>
        <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.sub }}>Pour les inventaires officiels : un site peut avoir plusieurs compteurs de dépotage (livraison), chacun relevé séparément.</p>
        <Field label="Site">
          <select className="somip-select" value={depotageSiteId} onChange={(e) => setDepotageSiteId(e.target.value)}>
            {stations.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </Field>
        <table className="somip-table" style={{ marginBottom: 12 }}>
          <thead><tr><th>Compteur de dépotage</th><th></th></tr></thead>
          <tbody>
            {currentDepotageMeters.length === 0 && <EmptyRow colSpan={2} text="Aucun compteur de dépotage configuré." />}
            {currentDepotageMeters.map((m) => (
              <tr key={m.id}>
                <td>{m.name}</td>
                <td style={{ textAlign: "right" }}><ConfirmIconButton onConfirm={() => removeSiteDepotageMeter(m)} /></td>
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{ display: "flex", gap: 8 }}>
          <input className="somip-input" style={{ flex: 1 }} value={newDepotageName} onChange={(e) => setNewDepotageName(e.target.value)} placeholder="Ex : Dépotage 1" />
          <button className="somip-btn somip-btn-primary" onClick={submitDepotageMeter} disabled={!newDepotageName.trim()}><Plus size={15} /></button>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Réceptions                                                            */
/* ------------------------------------------------------------------ */
/* ------------------------------------------------------------------ */
/* Saisie journalière (écran unique : réception, sortie/camion, retour) */
/* ------------------------------------------------------------------ */
function DailyEntryView({ sites, movements, inventaires, productStocks, siteMeters, addMovement, addInventaire, deleteMovement, deleteInventaire, settings, canWrite, canManage, assignedSiteIds, truckAssignments }) {
  const hasSiteRestriction = assignedSiteIds && assignedSiteIds.length > 0;
  const activeSites = sites.filter((s) => s.active !== false);
  const availableSites = hasSiteRestriction ? activeSites.filter((s) => assignedSiteIds.includes(s.id)) : activeSites;
  const [siteId, setSiteId] = useState((hasSiteRestriction ? assignedSiteIds[0] : sites[0]?.id) || "");
  const [product, setProduct] = useState("gasoil");
  const [date, setDate] = useState(todayStr());
  const [receptions, setReceptions] = useState([{ quantite: "", ref: "" }]);
  const [indexAvant, setIndexAvant] = useState("");
  const [indexApres, setIndexApres] = useState("");
  const [compteur, setCompteur] = useState("");
  const [compteurReadings, setCompteurReadings] = useState([{ compteur: "", indexAvant: "", indexApres: "" }]);
  const [indexBloque, setIndexBloque] = useState(false);
  const [sortieDirecte, setSortieDirecte] = useState("");
  const [chargements, setChargements] = useState([{ camion: "", quantite: "" }]);
  const [retoursCuve, setRetoursCuve] = useState([{ camion: "", quantite: "" }]);
  const [destinataire, setDestinataire] = useState("");
  const [retourQty, setRetourQty] = useState("");
  const [retourNote, setRetourNote] = useState("");
  const [retourCamionTruckId, setRetourCamionTruckId] = useState("");
  const [retourCuveTruckQty, setRetourCuveTruckQty] = useState("");
  const [retourCuveTruckNote, setRetourCuveTruckNote] = useState("");
  const [tempC, setTempC] = useState("");
  const [densite, setDensite] = useState("");
  const [stockFinMesure, setStockFinMesure] = useState("");
  const [commentaireInv, setCommentaireInv] = useState("");
  const [stockFinPhotos, setStockFinPhotos] = useState([]);
  const [stockDebutConfirm, setStockDebutConfirm] = useState("");

  const truckSites = sites.filter((s) => s.isMobile);
  const stationSites = sites.filter((s) => !s.isMobile);
  const isLubSite = LUBRICANT_SITE_IDS.includes(siteId);
  const isLub = product !== "gasoil";
  const lubDensite = LUBRICANTS.find((l) => l.id === product)?.densite || 0;
  const isMobileSite = sites.find((s) => s.id === siteId)?.isMobile || false;
  const existingInv = inventaires.find((i) => i.siteId === siteId && (i.product || "gasoil") === product && i.date === date);
  const skipVcf = isLub;

  useEffect(() => { if (!isLubSite) setProduct("gasoil"); }, [siteId, isLubSite]);

  const site = sites.find((s) => s.id === siteId);
  const meters = metersForSite(site, siteMeters);
  useEffect(() => { setCompteur(meters[0] || "Compteur"); }, [siteId]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { setCompteurReadings([{ compteur: meters[0] || "Compteur", indexAvant: "", indexApres: "" }]); }, [siteId]); // eslint-disable-line react-hooks/exhaustive-deps
  const lastIndexForMeter = (meterName) => movements
    .filter((m) => m.siteId === siteId && (m.product || "gasoil") === "gasoil" && (m.compteur || meters[0]) === meterName && (m.type === "sortie" || m.type === "sortie_camion" || m.type === "retour_camion") && m.indexApres !== undefined)
    .sort((a, b) => (a.date + (a.createdAt || "")).localeCompare(b.date + (b.createdAt || "")))
    .slice(-1)[0]?.indexApres;
  const productStockEntry = productStocks.find((p) => p.siteId === siteId && p.product === product);
  const stockDebut = isLub
    ? stockBeforeDateProduct(productStockEntry?.stockInitial || 0, movements, siteId, product, date, inventaires)
    : (site ? stockBeforeDate(site, movements, date, inventaires) : 0);
  const isFirstOfMonth = date.slice(-2) === "01";

  useEffect(() => {
    if (isFirstOfMonth) setStockDebutConfirm(String(Math.round(stockDebut)));
    else setStockDebutConfirm("");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [siteId, product, date]);
  const stockDebutEffective = isFirstOfMonth && stockDebutConfirm !== "" ? Number(stockDebutConfirm) : stockDebut;
  const receptionN = receptions.reduce((a, r) => a + (Number(r.quantite) || 0), 0);
  const retourN = (isLub || isMobileSite) ? 0 : (isLubSite ? retoursCuve.reduce((a, r) => a + (Number(r.quantite) || 0), 0) : (Number(retourQty) || 0));
  const retourCuveTruckN = isMobileSite ? (Number(retourCuveTruckQty) || 0) : 0;
  const isMultiCompteurEntry = meters.length > 1 && !isLub;
  const readingFlows = compteurReadings.map((r) => ({
    ...r,
    flow: r.indexAvant !== "" && r.indexApres !== "" ? Number(r.indexApres) - Number(r.indexAvant) : 0,
    valid: r.indexAvant === "" && r.indexApres === "" ? true : (r.indexAvant !== "" && r.indexApres !== "" && Number(r.indexApres) > Number(r.indexAvant)),
  }));
  const sortieQtySingle = indexAvant !== "" && indexApres !== "" ? Number(indexApres) - Number(indexAvant) : 0;
  const sortieValidSingle = indexAvant === "" && indexApres === "" ? true : (indexAvant !== "" && indexApres !== "" && sortieQtySingle > 0);
  // Panne de compteur (Superviseur uniquement) : la vente du jour est saisie directement, sans
  // passer par l'index — utile quand le compteur est bloqué et ne peut pas être relevé.
  const sortieQty = indexBloque ? (Number(sortieDirecte) || 0) : (isMultiCompteurEntry ? readingFlows.reduce((a, r) => a + r.flow, 0) : sortieQtySingle);
  const sortieValid = indexBloque ? true : (isMultiCompteurEntry ? readingFlows.every((r) => r.valid) : sortieValidSingle);
  const totalChargements = isLubSite && !isLub ? chargements.reduce((a, c) => a + (Number(c.quantite) || 0), 0) : 0;
  const chargementsValid = totalChargements <= sortieQty;
  const venteStation = isLubSite && !isLub ? Math.max(0, sortieQty - totalChargements) : sortieQty;
  // Sur un camion, le retour cuve passe par le même compteur que les ventes : l'index mesure le
  // flux total (vente terrain + retour cuve confondus), le retour cuve doit donc être déduit de
  // la vente réelle plutôt que soustrait une seconde fois du stock.
  // Le retour cuve peut être saisi séparément de l'index (un compteur à la fois) : pas de
  // comparaison avec sortieQty de cette session précise, qui pourrait être vide ce jour-là.
  const venteTruck = isMobileSite ? Math.max(0, sortieQty - retourCuveTruckN) : sortieQty;
  const stockTheoriqueAmbiant = stockDebutEffective + receptionN + retourN - sortieQty;
  const stockTheorique15 = !skipVcf && site ? stockBeforeDate15(site, movements, date, inventaires) : 0;

  // Chaque produit (gasoil, chaque lubrifiant) et chaque camion a son propre compteur de sortie.
  const lastIndexForSite = movements
    .filter((m) => m.siteId === siteId && (m.product || "gasoil") === product && (meters.length > 1 ? (m.compteur || meters[0]) === compteur : true) && ((isLub || isMobileSite) ? m.type === "sortie" : (m.type === "sortie" || m.type === "sortie_camion" || m.type === "retour_camion")) && m.indexApres !== undefined)
    .sort((a, b) => (a.date + (a.createdAt || "")).localeCompare(b.date + (b.createdAt || "")))
    .slice(-1)[0]?.indexApres;
  const indexMismatch = lastIndexForSite !== undefined && indexAvant !== "" && Number(indexAvant) !== lastIndexForSite;

  const vcfFor = (qty) => (skipVcf ? null : correctVolumeTo15({ volumeAmbiant: qty, tempC: tempC === "" ? NaN : Number(tempC), densiteObservee: Number(densite) || 0 }));
  const vcfExtra = (qty) => {
    const r = vcfFor(qty);
    return r ? { temperatureC: Number(tempC), densiteObservee: Number(densite), densite15: r.densite15, vcf: r.vcf, volumeCorrige15: r.volume15 } : {};
  };
  const vcfPreview = skipVcf ? null : vcfFor(receptionN || sortieQty || retourN || retourCuveTruckN || 1);

  const stockFinN = Number(stockFinMesure) || 0;
  const vcfFin = skipVcf ? null : correctVolumeTo15({ volumeAmbiant: stockFinN, tempC: tempC === "" ? NaN : Number(tempC), densiteObservee: Number(densite) || 0 });
  // Le Gain/Perte "officiel" est toujours en base ambiante (voir addInventaire) : le 15°C est
  // indicatif (page dédiée), il ne doit jamais se mélanger au résultat de l'équation de stock.
  const has15 = !!vcfFin;
  const theoriqueUsed = stockTheoriqueAmbiant;
  const physiqueUsed = stockFinN;
  const ecart = stockFinMesure === "" ? null : physiqueUsed - theoriqueUsed;
  const preview = ecart === null ? null : classifyEcart(ecart, theoriqueUsed, settings.objectifFreinte);

  const hasSomethingToSave = receptionN > 0 || sortieQty > 0 || retourN > 0 || retourCuveTruckN > 0 || stockFinMesure !== "";
  const stockFinConflict = stockFinMesure !== "" && !!existingInv;
  const photosWithoutStockFin = stockFinPhotos.length > 0 && stockFinMesure === "";
  const canSubmit = sortieValid && chargementsValid && hasSomethingToSave && !stockFinConflict && !photosWithoutStockFin && (!isFirstOfMonth || !hasSomethingToSave || stockDebutConfirm !== "");

  const resetDayFields = () => {
    setReceptions([{ quantite: "", ref: "" }]);
    setIndexAvant(""); setIndexApres(""); setDestinataire(""); setChargements([{ camion: "", quantite: "" }]);
    setIndexBloque(false); setSortieDirecte("");
    setCompteurReadings([{ compteur: meters[0] || "Compteur", indexAvant: "", indexApres: "" }]);
    setRetourQty(""); setRetourNote(""); setRetourCamionTruckId(""); setRetoursCuve([{ camion: "", quantite: "" }]); setRetourCuveTruckQty(""); setRetourCuveTruckNote(""); setTempC(""); setDensite("");
    setStockFinMesure(""); setCommentaireInv(""); setStockFinPhotos([]);
  };

  const [submitting, setSubmitting] = useState(false);

  const submit = async () => {
    if (!siteId || !date || !canSubmit || submitting) return;
    setSubmitting(true);
    try {
      if (isFirstOfMonth) {
        const diff = Number(stockDebutConfirm) - stockDebut;
        if (Math.abs(diff) >= 1) {
          const prev = new Date(date);
          prev.setDate(prev.getDate() - 1);
          const adjDate = `${prev.getFullYear()}-${pad2(prev.getMonth() + 1)}-${pad2(prev.getDate())}`;
          const ok = await addMovement({ siteId, product, type: "ajustement", date: adjDate, quantity: Math.abs(diff), delta: diff, commentaire: "Confirmation du stock début de mois" });
          if (!ok) return;
        }
      }
      if (receptionN > 0) {
        for (const r of receptions) {
          const qty = Number(r.quantite) || 0;
          if (qty > 0) {
            const ok = await addMovement({ siteId, product, type: "reception", date, quantity: qty, delta: qty, ref: r.ref, ...vcfExtra(qty) });
            if (!ok) return;
          }
        }
      }
      if (sortieQty > 0) {
        const compteurField = meters.length > 1 ? compteur : undefined;
        if (indexBloque) {
          // Panne de compteur (Superviseur) : la quantité saisie est le flux total (comme le
          // ferait le compteur) — les chargements laitiers doivent en être déduits pour obtenir
          // la vraie vente, exactement comme en fonctionnement normal (venteStation).
          if (venteStation > 0) {
            const ok = await addMovement({ siteId, product, type: "sortie", date, quantity: venteStation, delta: -venteStation, commentaire: "Compteur en panne — saisie directe sans index", destinataire, compteur: compteurField, ...vcfExtra(venteStation) });
            if (!ok) return;
          }
          if (isLubSite && !isLub) {
            for (const c of chargements) {
              const qty = Number(c.quantite) || 0;
              if (qty > 0 && c.camion) {
                const ok2 = await addMovement({ siteId, product, type: "sortie_camion", date, quantity: qty, delta: -qty, camion: c.camion, ...vcfExtra(qty) });
                if (!ok2) return;
                // Transfert (camion normalement affecté ailleurs) : la réception se crée
                // automatiquement côté camion, sans double saisie.
                const label = transferLabel(sites, truckAssignments || [], c.camion, siteId, date);
                if (label) {
                  const ok3 = await addMovement({ siteId: c.camion, product, type: "reception", date, quantity: qty, delta: qty, commentaire: label, ...vcfExtra(qty) });
                  if (!ok3) return;
                }
              }
            }
          }
        } else if (isMultiCompteurEntry) {
          // Le(s) compteur(s) mesurent le flux total (vente + chargements camions confondus) :
          // on répartit les chargements sur les compteurs saisis (dans l'ordre), puis on
          // enregistre le reliquat "vente" par compteur avec son propre index.
          let remaining = totalChargements;
          for (const r of readingFlows) {
            if (r.flow <= 0) continue;
            const used = Math.min(r.flow, remaining);
            remaining -= used;
            const vente = r.flow - used;
            if (vente > 0) {
              const ok = await addMovement({ siteId, product, type: "sortie", date, quantity: vente, delta: -vente, indexAvant: Number(r.indexAvant), indexApres: Number(r.indexApres), compteur: r.compteur || undefined, destinataire, ...vcfExtra(vente) });
              if (!ok) return;
            }
          }
          for (const c of chargements) {
            const qty = Number(c.quantite) || 0;
            if (qty > 0 && c.camion) {
              const ok = await addMovement({ siteId, product, type: "sortie_camion", date, quantity: qty, delta: -qty, camion: c.camion, ...vcfExtra(qty) });
              if (!ok) return;
              // Transfert (camion normalement affecté ailleurs) : la réception se crée
              // automatiquement côté camion, sans double saisie.
              const label = transferLabel(sites, truckAssignments || [], c.camion, siteId, date);
              if (label) {
                const ok3 = await addMovement({ siteId: c.camion, product, type: "reception", date, quantity: qty, delta: qty, commentaire: label, ...vcfExtra(qty) });
                if (!ok3) return;
              }
            }
          }
        } else if (isLub || isMobileSite) {
          // La "Sortie" enregistre toujours le flux COMPLET du compteur (jamais réduit par le
          // retour cuve) : Sortie et Retour cuve peuvent être saisis à des moments différents,
          // donc on ne peut pas se fier à ce qui est rempli "en même temps". Le retour cuve est
          // neutre pour le stock (delta 0, voir plus bas) — c'est lui qui s'ajuste, jamais la Sortie.
          const qty = sortieQty;
          if (qty > 0) {
            const ok = await addMovement({ siteId, product, type: "sortie", date, quantity: qty, delta: -qty, indexAvant: Number(indexAvant), indexApres: Number(indexApres), destinataire, compteur: compteurField });
            if (!ok) return;
          }
        } else {
          const ok = await addMovement({ siteId, product, type: "sortie", date, quantity: sortieQty, delta: -sortieQty, indexAvant: Number(indexAvant), indexApres: Number(indexApres), compteur: compteurField, destinataire, ...vcfExtra(sortieQty) });
          if (!ok) return;
        }
      }
      if (retourN > 0) {
        if (isLubSite) {
          for (const r of retoursCuve) {
            const qty = Number(r.quantite) || 0;
            if (qty > 0) {
              const ok = await addMovement({ siteId, product, type: "retour_camion", date, quantity: qty, delta: qty, camion: r.camion || undefined, ...vcfExtra(qty) });
              if (!ok) return;
            }
          }
        } else {
          const ok = await addMovement({ siteId, product, type: "retour_camion", date, quantity: retourN, delta: retourN, camion: retourCamionTruckId || undefined, destination: retourNote, ...vcfExtra(retourN) });
          if (!ok) return;
        }
      }
      if (retourCuveTruckN > 0) {
        // Delta = 0 volontairement : le retour cuve passe par le même compteur que la Sortie,
        // qui a déjà enregistré la totalité du flux (retour inclus). Ce mouvement sert seulement
        // à distinguer, pour l'affichage et les rapports, quelle part de cette Sortie était un
        // retour cuve plutôt qu'une vente terrain — sans jamais soustraire le stock une 2e fois.
        const ok = await addMovement({ siteId, product, type: "retour_cuve_camion", date, quantity: retourCuveTruckN, delta: 0, destination: retourCuveTruckNote, ...vcfExtra(retourCuveTruckN) });
        if (!ok) return;
      }
      if (stockFinMesure !== "") {
        const invExtra = vcfFin
          ? { temperatureC: Number(tempC), densiteObservee: Number(densite), densite15: vcfFin.densite15, vcf: vcfFin.vcf, stockPhysique15: vcfFin.volume15 }
          : {};
        const okInv = await addInventaire({ siteId, product, date, stockPhysique: stockFinN, commentaire: commentaireInv, photoFiles: stockFinPhotos, ...invExtra });
        if (!okInv) return;
      }
      resetDayFields();
    } finally {
      setSubmitting(false);
    }
  };

  const dayMovs = movements.filter((m) => m.siteId === siteId && (m.product || "gasoil") === product && m.date === date).sort((a, b) => (a.createdAt || "").localeCompare(b.createdAt || ""));

  return (
    <div className="somip-fade" style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      {canWrite && (
        <div className="somip-panel" style={{ flex: "1 1 340px", padding: 18 }}>
          <h3 style={{ margin: "0 0 14px", fontSize: 14 }}>Saisie du jour</h3>
          <div style={{ display: "flex", gap: 8 }}>
            <div style={{ flex: 1 }}>
              <Field label="Site">
                {hasSiteRestriction && availableSites.length === 1 ? (
                  <div className="somip-input" style={{ background: C.bg, color: C.ink, fontWeight: 600, display: "flex", alignItems: "center" }}>
                    {availableSites[0].name}
                  </div>
                ) : (
                  <select className="somip-select" value={siteId} onChange={(e) => setSiteId(e.target.value)}>
                    {availableSites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                )}
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Date"><input type="date" className="somip-input" value={date} onChange={(e) => setDate(e.target.value)} /></Field>
            </div>
          </div>

          {isLubSite && (
            <Field label="Produit">
              <select className="somip-select" value={product} onChange={(e) => setProduct(e.target.value)}>
                {PRODUCTS.map((p) => <option key={p.id} value={p.id}>{p.label}</option>)}
              </select>
            </Field>
          )}

          {isLub && !productStockEntry && (
            <p style={{ margin: "-6px 0 12px", fontSize: 11.5, color: C.warning }}>
              Aucune capacité/stock initial défini pour {LUBRICANTS.find((l) => l.id === product)?.label} sur ce site — configure-le sur la page "Sites" (Superviseur).
            </p>
          )}

          {existingInv && canManage && (
            <div style={{ margin: "-6px 0 12px", padding: "9px 12px", background: "#FCEAEA", borderRadius: 8, fontSize: 11.5, color: C.danger }}>
              Cette journée est déjà enregistrée pour ce site/produit (Stock fin : {fmt(existingInv.stockPhysique)} L). Pour éviter un doublon, l'enregistrement est bloqué. Si tu dois corriger cette journée, supprime d'abord la ligne "Stock fin" dans le tableau à droite, puis ressaisis.
            </div>
          )}
          {existingInv && !canManage && (
            <div style={{ margin: "-6px 0 12px", padding: "9px 12px", background: "#FCEAEA", borderRadius: 8, fontSize: 11.5, color: C.danger }}>
              Cette journée est déjà enregistrée pour ce site/produit (Stock fin : {fmt(existingInv.stockPhysique)} L). Pour éviter un doublon, l'enregistrement est bloqué — demande au Superviseur de corriger si besoin.
            </div>
          )}

          <div style={{ background: C.bg, borderRadius: 8, padding: "9px 12px", marginBottom: 14, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ fontSize: 12.5, color: C.sub, fontWeight: 600 }}>Stock début (calculé)</span>
            <span className="somip-mono" style={{ fontWeight: 700 }}>{fmt(stockDebut)} L{isLub && ` (≈ ${fmt(stockDebut * lubDensite)} kg)`}</span>
          </div>

          {isFirstOfMonth && (
            <>
              <p style={{ margin: "0 0 6px", fontSize: 12, fontWeight: 700, color: C.warning }}>1er du mois — Stock début du mois <span style={{ color: C.danger }}>*</span></p>
              <Field label="Confirme ou corrige le stock début du mois (L, obligatoire)">
                <input type="number" className="somip-input" value={stockDebutConfirm} onChange={(e) => setStockDebutConfirm(e.target.value)} placeholder="Ex : 20000" />
              </Field>
              <p style={{ margin: "-6px 0 12px", fontSize: 11, color: C.sub }}>
                Pré-rempli avec le stock calculé automatiquement ({fmt(stockDebut)} L). Corrige cette valeur si le relevé physique de début de mois est différent — l'écart sera enregistré comme un ajustement.
              </p>
            </>
          )}

          <p style={{ margin: "0 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>{isMobileSite ? "Chargement" : "Réception"}</p>
          {receptions.map((r, idx) => (
            <div key={idx} style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "flex-end" }}>
              <div style={{ flex: 1 }}>
                <Field label="Quantité reçue (L)">
                  <input type="number" className="somip-input" value={r.quantite} onChange={(e) => setReceptions((prev) => prev.map((row, i) => (i === idx ? { ...row, quantite: e.target.value } : row)))} placeholder="0" />
                </Field>
              </div>
              {!isMobileSite && (
                <div style={{ flex: 1 }}>
                  <Field label="N° Bon de livraison">
                    <input className="somip-input" value={r.ref} onChange={(e) => setReceptions((prev) => prev.map((row, i) => (i === idx ? { ...row, ref: e.target.value } : row)))} placeholder="BL-XXXX" />
                  </Field>
                </div>
              )}
              {receptions.length > 1 && (
                <button onClick={() => setReceptions((prev) => prev.filter((_, i) => i !== idx))} style={{ border: "none", background: "none", cursor: "pointer", padding: "9px 4px" }}>
                  <X size={16} color={C.danger} />
                </button>
              )}
            </div>
          ))}
          <button className="somip-btn somip-btn-secondary" style={{ fontSize: 12, padding: "6px 12px", marginBottom: 10 }} onClick={() => setReceptions((prev) => [...prev, { quantite: "", ref: "" }])}>
            <Plus size={13} /> Ajouter une réception
          </button>
          {receptions.filter((r) => Number(r.quantite) > 0).length > 1 && (
            <p style={{ margin: "-4px 0 10px", fontSize: 11, color: C.sub }}>
              Total reçu : {fmt(receptionN)} L, sur {receptions.filter((r) => Number(r.quantite) > 0).length} réceptions.
            </p>
          )}
          {isMobileSite && (
            <p style={{ margin: "-6px 0 10px", fontSize: 11, color: C.warning }}>
              Pense à saisir aussi ce chargement côté Prehomo/Okouma (Chargement laitiers) — les deux côtés sont indépendants. Pas de N° de bon pour un chargement interne.
            </p>
          )}
          {isLub && receptionN > 0 && <p style={{ margin: "-6px 0 10px", fontSize: 11, color: C.sub }}>≈ {fmt(receptionN * lubDensite)} kg</p>}

          {canWrite && (
            <label style={{ display: "flex", alignItems: "center", gap: 7, margin: "10px 0 8px", cursor: "pointer", fontSize: 12, color: C.warning, fontWeight: 600 }}>
              <input type="checkbox" checked={indexBloque} onChange={(e) => setIndexBloque(e.target.checked)} />
              Compteur en panne — saisir la sortie directement, sans index
            </label>
          )}
          {indexBloque ? (
            <>
              <p style={{ margin: "0 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Sortie (saisie directe — compteur en panne)</p>
              <Field label="Quantité sortie (L)"><input type="number" className="somip-input" value={sortieDirecte} onChange={(e) => setSortieDirecte(e.target.value)} placeholder="0" /></Field>
              <p style={{ margin: "-4px 0 10px", fontSize: 11, color: C.sub }}>Le compteur ne sera pas mis à jour — pense à noter l'index réel dès qu'il redevient lisible.</p>
              {isLubSite && !isLub && (
                <>
                  <p style={{ margin: "10px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Chargement laitiers (prélevé sur cette sortie)</p>
                  {chargements.map((c, idx) => (
                    <div key={idx} style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "flex-end" }}>
                      <div style={{ flex: 1 }}>
                        <Field label="Camion">
                          <select className="somip-select" value={c.camion} onChange={(e) => setChargements((prev) => prev.map((r, i) => (i === idx ? { ...r, camion: e.target.value } : r)))}>
                            <option value="">— choisir —</option>
                            {truckSites.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                          </select>
                        </Field>
                      </div>
                      <div style={{ flex: 1 }}>
                        <Field label="Quantité chargée (L)">
                          <input type="number" className="somip-input" value={c.quantite} onChange={(e) => setChargements((prev) => prev.map((r, i) => (i === idx ? { ...r, quantite: e.target.value } : r)))} placeholder="0" />
                        </Field>
                      </div>
                      {chargements.length > 1 && (
                        <button onClick={() => setChargements((prev) => prev.filter((_, i) => i !== idx))} style={{ border: "none", background: "none", cursor: "pointer", padding: "9px 4px" }}>
                          <X size={16} color={C.danger} />
                        </button>
                      )}
                    </div>
                  ))}
                  <button className="somip-btn somip-btn-secondary" style={{ fontSize: 12, padding: "6px 12px", marginBottom: 10 }} onClick={() => setChargements((prev) => [...prev, { camion: "", quantite: "" }])}>
                    <Plus size={13} /> Ajouter un camion
                  </button>
                  {!chargementsValid && <p style={{ margin: "-4px 0 10px", fontSize: 11.5, color: C.danger }}>Le total chargé ({fmt(totalChargements)} L) dépasse la quantité saisie ({fmt(sortieQty)} L).</p>}
                  {sortieQty > 0 && (
                    <div style={{ background: C.bg, borderRadius: 8, padding: "9px 12px", marginBottom: 12, fontSize: 12.5 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: totalChargements > 0 ? 4 : 0 }}>
                        <span style={{ color: C.sub, fontWeight: 600 }}>Vente réelle (déduite des chargements)</span>
                        <span className="somip-mono" style={{ fontWeight: 700 }}>{fmt(venteStation)} L</span>
                      </div>
                      {totalChargements > 0 && (
                        <div style={{ display: "flex", justifyContent: "space-between" }}>
                          <span style={{ color: C.sub, fontWeight: 600 }}>Chargements laitiers (total)</span>
                          <span className="somip-mono" style={{ fontWeight: 700, color: C.orange }}>{fmt(totalChargements)} L</span>
                        </div>
                      )}
                    </div>
                  )}
                </>
              )}
            </>
          ) : isMultiCompteurEntry ? (
            <>
              <p style={{ margin: "10px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>{isLubSite ? "Sortie (compteurs — flux total : vente + chargements camions)" : isMobileSite ? "Sortie Fiche Terrain (postes)" : "Sortie (compteurs)"}</p>
              {compteurReadings.map((r, idx) => {
                const lastIdx = r.compteur ? lastIndexForMeter(r.compteur) : undefined;
                const mismatch = lastIdx !== undefined && r.indexAvant !== "" && Number(r.indexAvant) !== lastIdx;
                return (
                  <div key={idx} style={{ border: `1px solid ${C.border}`, borderRadius: 8, padding: 10, marginBottom: 8 }}>
                    <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
                      {meters.length > 1 && (
                        <div style={{ flex: 1 }}>
                          <Field label={isMobileSite ? "Poste" : "Compteur"}>
                            <select className="somip-select" value={r.compteur} onChange={(e) => setCompteurReadings((prev) => prev.map((row, i) => (i === idx ? { ...row, compteur: e.target.value } : row)))}>
                              {meters.map((m) => <option key={m} value={m}>{m}</option>)}
                            </select>
                          </Field>
                        </div>
                      )}
                      <div style={{ flex: 1 }}><Field label="Index avant"><input type="number" className="somip-input" value={r.indexAvant} onChange={(e) => setCompteurReadings((prev) => prev.map((row, i) => (i === idx ? { ...row, indexAvant: e.target.value } : row)))} placeholder="Ex : 45210" /></Field></div>
                      <div style={{ flex: 1 }}><Field label="Index après"><input type="number" className="somip-input" value={r.indexApres} onChange={(e) => setCompteurReadings((prev) => prev.map((row, i) => (i === idx ? { ...row, indexApres: e.target.value } : row)))} placeholder="Ex : 47210" /></Field></div>
                      {compteurReadings.length > 1 && (
                        <button onClick={() => setCompteurReadings((prev) => prev.filter((_, i) => i !== idx))} style={{ border: "none", background: "none", cursor: "pointer", padding: "9px 4px" }}>
                          <X size={16} color={C.danger} />
                        </button>
                      )}
                    </div>
                    {lastIdx !== undefined && (
                      <p style={{ margin: "4px 0 0", fontSize: 11, color: mismatch ? C.warning : C.sub }}>
                        Dernier index enregistré ({r.compteur}) : {fmt(lastIdx)}{mismatch && " — vérifie ton index avant."}
                      </p>
                    )}
                  </div>
                );
              })}
              {compteurReadings.length < meters.length && (
                <button className="somip-btn somip-btn-secondary" style={{ fontSize: 12, padding: "6px 12px", marginBottom: 10 }} onClick={() => {
                  const used = compteurReadings.map((r) => r.compteur);
                  const next = meters.find((m) => !used.includes(m)) || meters[0];
                  setCompteurReadings((prev) => [...prev, { compteur: next, indexAvant: "", indexApres: "" }]);
                }}>
                  <Plus size={13} /> Ajouter un 2e compteur (même journée)
                </button>
              )}
              {!sortieValid && <p style={{ margin: "-6px 0 10px", fontSize: 11.5, color: C.danger }}>L'index après doit être supérieur à l'index avant, pour chaque compteur.</p>}
              {sortieQty > 0 && compteurReadings.filter((r) => r.indexAvant !== "" && r.indexApres !== "").length > 1 && (
                <div style={{ background: C.bg, borderRadius: 8, padding: "9px 12px", marginBottom: 12, fontSize: 12.5, display: "flex", justifyContent: "space-between" }}>
                  <span style={{ color: C.sub, fontWeight: 600 }}>Total sortie (tous compteurs)</span>
                  <span className="somip-mono" style={{ fontWeight: 700 }}>{fmt(sortieQty)} L</span>
                </div>
              )}

              {isLubSite && (
                <>
                  <p style={{ margin: "12px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Chargement laitiers (prélevé sur ce flux)</p>
                  <p style={{ margin: "0 0 8px", fontSize: 11, color: C.warning }}>Pense à saisir aussi ce chargement côté camion (Chargement) — sauf pour un transfert (camion affecté à un autre site), ajouté automatiquement.</p>
                  {chargements.map((c, idx) => (
                    <div key={idx} style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "flex-end" }}>
                      <div style={{ flex: 1 }}>
                        <Field label="Camion">
                          <select className="somip-select" value={c.camion} onChange={(e) => setChargements((prev) => prev.map((r, i) => (i === idx ? { ...r, camion: e.target.value } : r)))}>
                            <option value="">— choisir —</option>
                            {truckSites.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                          </select>
                        </Field>
                      </div>
                      <div style={{ flex: 1 }}>
                        <Field label="Quantité chargée (L)">
                          <input type="number" className="somip-input" value={c.quantite} onChange={(e) => setChargements((prev) => prev.map((r, i) => (i === idx ? { ...r, quantite: e.target.value } : r)))} placeholder="0" />
                        </Field>
                      </div>
                      {chargements.length > 1 && (
                        <button onClick={() => setChargements((prev) => prev.filter((_, i) => i !== idx))} style={{ border: "none", background: "none", cursor: "pointer", padding: "9px 4px" }}>
                          <X size={16} color={C.danger} />
                        </button>
                      )}
                    </div>
                  ))}
                  {chargements.map((c, idx2) => {
                    const label = c.camion ? transferLabel(sites, truckAssignments || [], c.camion, siteId, date) : "";
                    return label ? (
                      <p key={idx2} style={{ margin: "-4px 0 8px", fontSize: 11, color: C.success, fontWeight: 600 }}>
                        {label} — la réception sera ajoutée automatiquement côté camion, rien à saisir en plus.
                      </p>
                    ) : null;
                  })}
                  <button className="somip-btn somip-btn-secondary" style={{ fontSize: 12, padding: "6px 12px", marginBottom: 10 }} onClick={() => setChargements((prev) => [...prev, { camion: "", quantite: "" }])}>
                    <Plus size={13} /> Ajouter un camion
                  </button>
                  {!chargementsValid && <p style={{ margin: "-4px 0 10px", fontSize: 11.5, color: C.danger }}>Le total chargé ({fmt(totalChargements)} L) dépasse le flux total des compteurs ({fmt(sortieQty)} L).</p>}
                  {sortieQty > 0 && (
                    <div style={{ background: C.bg, borderRadius: 8, padding: "9px 12px", marginBottom: 12, fontSize: 12.5 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: totalChargements > 0 ? 4 : 0 }}>
                        <span style={{ color: C.sub, fontWeight: 600 }}>Vente station (calculée)</span>
                        <span className="somip-mono" style={{ fontWeight: 700 }}>{fmt(venteStation)} L</span>
                      </div>
                      {totalChargements > 0 && (
                        <div style={{ display: "flex", justifyContent: "space-between" }}>
                          <span style={{ color: C.sub, fontWeight: 600 }}>Chargements laitiers (total)</span>
                          <span className="somip-mono" style={{ fontWeight: 700, color: C.orange }}>{fmt(totalChargements)} L</span>
                        </div>
                      )}
                    </div>
                  )}
                </>
              )}
            </>
          ) : (isLub || isMobileSite) ? (
            <>
              <p style={{ margin: "10px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>{isMobileSite ? "Sortie Fiche Terrain (compteur)" : "Sortie (compteur de livraison)"}</p>
              <div style={{ display: "flex", gap: 8 }}>
                <div style={{ flex: 1 }}><Field label="Index avant"><input type="number" className="somip-input" value={indexAvant} onChange={(e) => setIndexAvant(e.target.value)} placeholder="Ex : 120" /></Field></div>
                <div style={{ flex: 1 }}><Field label="Index après"><input type="number" className="somip-input" value={indexApres} onChange={(e) => setIndexApres(e.target.value)} placeholder="Ex : 145" /></Field></div>
              </div>
              {lastIndexForSite !== undefined && (
                <p style={{ margin: "-6px 0 8px", fontSize: 11, color: indexMismatch ? C.warning : C.sub }}>
                  Dernier index enregistré pour {isMobileSite ? "ce camion" : "ce produit"} : {fmt(lastIndexForSite)}{indexMismatch && (isMobileSite ? " — écart avec l'index d'avant (normal après un secours en carrière), pas bloquant." : " — vérifie ton index avant.")}
                </p>
              )}
              {!sortieValid && <p style={{ margin: "-6px 0 10px", fontSize: 11.5, color: C.danger }}>L'index après doit être supérieur à l'index avant.</p>}
              {isLub && sortieQty > 0 && <p style={{ margin: "-6px 0 10px", fontSize: 11, color: C.sub }}>≈ {fmt(sortieQty * lubDensite)} kg</p>}
            </>
          ) : (
            <>
              <p style={{ margin: "10px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Sortie (compteur)</p>
              {meters.length > 1 && (
                <Field label="Compteur">
                  <select className="somip-select" value={compteur} onChange={(e) => setCompteur(e.target.value)}>
                    {meters.map((m) => <option key={m} value={m}>{m}</option>)}
                  </select>
                </Field>
              )}
              <div style={{ display: "flex", gap: 8 }}>
                <div style={{ flex: 1 }}><Field label="Index avant"><input type="number" className="somip-input" value={indexAvant} onChange={(e) => setIndexAvant(e.target.value)} placeholder="Ex : 45210" /></Field></div>
                <div style={{ flex: 1 }}><Field label="Index après"><input type="number" className="somip-input" value={indexApres} onChange={(e) => setIndexApres(e.target.value)} placeholder="Ex : 47210" /></Field></div>
              </div>
              {lastIndexForSite !== undefined && (
                <p style={{ margin: "-6px 0 8px", fontSize: 11, color: indexMismatch ? C.warning : C.sub }}>
                  Dernier index enregistré sur ce site{meters.length > 1 ? ` (${compteur})` : ""} : {fmt(lastIndexForSite)}{indexMismatch && " — vérifie ton index avant."}
                </p>
              )}
              {!sortieValid && <p style={{ margin: "-6px 0 10px", fontSize: 11.5, color: C.danger }}>L'index après doit être supérieur à l'index avant.</p>}
            </>
          )}

          {isLubSite && !isLub && !isMobileSite && (
            <>
              <p style={{ margin: "10px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Retour cuve (camion)</p>
              {retoursCuve.map((r, idx) => (
                <div key={idx} style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "flex-end" }}>
                  <div style={{ flex: 1 }}>
                    <Field label="Camion">
                      <select className="somip-select" value={r.camion} onChange={(e) => setRetoursCuve((prev) => prev.map((row, i) => (i === idx ? { ...row, camion: e.target.value } : row)))}>
                        <option value="">— choisir —</option>
                        {truckSites.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                      </select>
                    </Field>
                  </div>
                  <div style={{ flex: 1 }}>
                    <Field label="Quantité retournée (L)">
                      <input type="number" className="somip-input" value={r.quantite} onChange={(e) => setRetoursCuve((prev) => prev.map((row, i) => (i === idx ? { ...row, quantite: e.target.value } : row)))} placeholder="0" />
                    </Field>
                  </div>
                  {retoursCuve.length > 1 && (
                    <button onClick={() => setRetoursCuve((prev) => prev.filter((_, i) => i !== idx))} style={{ border: "none", background: "none", cursor: "pointer", padding: "9px 4px" }}>
                      <X size={16} color={C.danger} />
                    </button>
                  )}
                </div>
              ))}
              <button className="somip-btn somip-btn-secondary" style={{ fontSize: 12, padding: "6px 12px", marginBottom: 10 }} onClick={() => setRetoursCuve((prev) => [...prev, { camion: "", quantite: "" }])}>
                <Plus size={13} /> Ajouter un camion
              </button>
              {retourN > 0 && (
                <p style={{ margin: "-4px 0 10px", fontSize: 11, color: C.sub }}>
                  Total retourné : {fmt(retourN)} L, sur {retoursCuve.filter((r) => Number(r.quantite) > 0).length} camion(s).
                </p>
              )}
              <p style={{ margin: "-4px 0 12px", fontSize: 11, color: C.warning }}>
                Pense à saisir aussi ce retour côté camion (Retour cuve) — les deux côtés sont indépendants.
              </p>
            </>
          )}

          {isMobileSite && (
            <>
              <p style={{ margin: "10px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Retour Cuve (vers un site)</p>
              <div style={{ display: "flex", gap: 8 }}>
                <div style={{ flex: 1 }}><Field label="Quantité rendue (L)"><input type="number" className="somip-input" value={retourCuveTruckQty} onChange={(e) => setRetourCuveTruckQty(e.target.value)} placeholder="0" /></Field></div>
                <div style={{ flex: 1 }}>
                  <Field label="Site destinataire">
                    <select className="somip-select" value={retourCuveTruckNote} onChange={(e) => setRetourCuveTruckNote(e.target.value)}>
                      <option value="">— choisir —</option>
                      {stationSites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                    </select>
                  </Field>
                </div>
              </div>
              {retourCuveTruckQty && retourCuveTruckNote && (
                <p style={{ margin: "-6px 0 8px", fontSize: 11, color: C.sub }}>
                  Ce retour ({fmt(retourCuveTruckN)} L) concerne {stationSites.find((s) => s.id === retourCuveTruckNote)?.name}.
                </p>
              )}
              {sortieQty > 0 && retourCuveTruckN > 0 && (
                <div style={{ background: C.bg, borderRadius: 8, padding: "9px 12px", marginBottom: 12, fontSize: 12.5 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                    <span style={{ color: C.sub, fontWeight: 600 }}>Vente terrain nette (indicatif — rapports)</span>
                    <span className="somip-mono" style={{ fontWeight: 700 }}>{fmt(venteTruck)} L</span>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <span style={{ color: C.sub, fontWeight: 600 }}>dont retour cuve (inclus dans le compteur)</span>
                    <span className="somip-mono" style={{ fontWeight: 700, color: C.orange }}>{fmt(retourCuveTruckN)} L</span>
                  </div>
                </div>
              )}
            </>
          )}

          {!isLub && (
            <>
              <p style={{ margin: "10px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Température &amp; densité (du jour)</p>
              <VcfMiniPanel tempC={tempC} densite={densite} onTempC={setTempC} onDensite={setDensite} result={vcfPreview} compact />
            </>
          )}

          <div style={{ background: C.bg, borderRadius: 8, padding: "9px 12px", margin: "12px 0", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ fontSize: 12.5, color: C.sub, fontWeight: 600 }}>Stock théorique (calculé)</span>
            <span className="somip-mono" style={{ fontWeight: 700 }}>{fmt(stockTheoriqueAmbiant)} L{isLub && ` (≈ ${fmt(stockTheoriqueAmbiant * lubDensite)} kg)`}</span>
          </div>

          <p style={{ margin: "10px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Stock fin — jauge mesurée <span style={{ color: C.danger, fontWeight: 700 }}>*</span></p>
          <Field label={`Stock fin mesuré (L, obligatoire)`}><input type="number" className="somip-input" value={stockFinMesure} onChange={(e) => setStockFinMesure(e.target.value)} placeholder="Lecture directe de la jauge" /></Field>
          {isLub && stockFinMesure !== "" && <p style={{ margin: "-6px 0 10px", fontSize: 11, color: C.sub }}>≈ {fmt(stockFinN * lubDensite)} kg</p>}
          <Field label="Commentaire inventaire (optionnel)"><textarea className="somip-textarea" rows={2} value={commentaireInv} onChange={(e) => setCommentaireInv(e.target.value)} /></Field>
          <Field label="Photos justificatives (perte, déversement, éclatement de filtre...)">
            <PhotoPicker files={stockFinPhotos} setFiles={setStockFinPhotos} />
          </Field>
          {stockFinPhotos.length > 0 && stockFinMesure === "" && (
            <p style={{ margin: "-6px 0 10px", fontSize: 11.5, color: C.danger, fontWeight: 600 }}>
              ⚠️ Remplis le Stock fin ci-dessus pour que {stockFinPhotos.length > 1 ? "ces photos soient" : "cette photo soit"} enregistrée(s) — sans Stock fin, {stockFinPhotos.length > 1 ? "elles ne seront pas" : "elle ne sera pas"} sauvegardée(s).
            </p>
          )}

          {preview && (
            <div style={{ background: C.bg, borderRadius: 8, padding: 12, margin: "4px 0 14px", fontSize: 12.5 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                <span style={{ color: C.sub }}>Base retenue</span>
                <Badge color={C.sub}>Ambiant</Badge>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                <span style={{ color: C.sub }}>Écart (Gain/Perte)</span>
                <span className="somip-mono" style={{ fontWeight: 700, color: NATURE_META[preview.nature].color }}>{ecart >= 0 ? "+" : ""}{fmt(ecart)} L{isLub && ` (${ecart >= 0 ? "+" : ""}${fmt(ecart * lubDensite)} kg)`}</span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between" }}>
                <span style={{ color: C.sub }}>Nature</span>
                <Badge color={NATURE_META[preview.nature].color}>{NATURE_META[preview.nature].label}</Badge>
              </div>
              {has15 && <p style={{ margin: "8px 0 0", fontSize: 11, color: C.sub }}>Le Gain/Perte officiel reste toujours en base ambiante. La version corrigée à 15°C est disponible séparément dans Rapports → État des stocks — 15°C.</p>}
            </div>
          )}

          <button className="somip-btn somip-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={submit} disabled={!canSubmit || submitting}>
            {submitting ? <Loader2 size={15} style={{ animation: "somipSpin .8s linear infinite" }} /> : <Plus size={15} />} {submitting ? "Enregistrement..." : "Enregistrer"}
          </button>
          <p style={{ margin: "8px 0 0", fontSize: 11, color: C.sub }}>
            Tu peux enregistrer un compteur à la fois (reviens plus tard pour les autres) — seul le Stock fin est à saisir une seule fois, en fin de journée.
          </p>
          {!canSubmit && !hasSomethingToSave && <p style={{ margin: "8px 0 0", fontSize: 11.5, color: C.sub }}>Remplis au moins un champ (Réception, Sortie, Retour ou Stock fin) pour enregistrer.</p>}
          {!canSubmit && stockFinConflict && <p style={{ margin: "8px 0 0", fontSize: 11.5, color: C.danger }}>Un Stock fin existe déjà pour ce jour — supprime-le d'abord si tu veux le corriger.</p>}
          {!canSubmit && hasSomethingToSave && !stockFinConflict && isFirstOfMonth && stockDebutConfirm === "" && <p style={{ margin: "8px 0 0", fontSize: 11.5, color: C.sub }}>Le Stock début du mois est obligatoire pour enregistrer.</p>}
        </div>
      )}

      <div className="somip-panel" style={{ flex: "1 1 380px", padding: 18 }}>
        <h3 style={{ margin: "0 0 14px", fontSize: 14 }}>Mouvements du {date} — {site?.name}{isLub && ` — ${LUBRICANTS.find((l) => l.id === product)?.label}`}</h3>
        <table className="somip-table">
          <thead><tr><th>Type</th><th>Détail</th><th style={{ textAlign: "right" }}>Quantité</th>{canManage && <th></th>}</tr></thead>
          <tbody>
            {dayMovs.length === 0 && <EmptyRow colSpan={canManage ? 4 : 3} text="Aucune écriture pour ce jour." />}
            {dayMovs.map((m) => {
              const meta = TYPE_META[m.type];
              const label = m.type === "reception" ? (isMobileSite ? "Chargement" : "Réception") : m.type === "sortie" ? (isMobileSite ? "Sortie terrain" : "Vente") : m.type === "sortie_camion" ? "Camion" : m.type === "retour_camion" ? "Retour" : m.type === "retour_cuve_camion" ? "Retour cuve" : "Ajustement";
              const nameOf = (id) => sites.find((s) => s.id === id)?.name || id;
              const detail = m.type === "reception" ? (m.ref || "—") : m.type === "sortie" ? (m.destinataire || "—") : m.type === "retour_cuve_camion" ? (m.destination ? nameOf(m.destination) : "—") : m.camion ? `${nameOf(m.camion)} ${m.type === "retour_camion" ? "←" : "→"} ${m.destination || "—"}` : "—";
              const transfer = m.type === "sortie_camion" && m.camion ? transferLabel(sites, truckAssignments || [], m.camion, siteId, m.date) : "";
              return (
                <tr key={m.id}>
                  <td><Badge color={meta.color}>{label}</Badge></td>
                  <td style={{ color: C.sub }}>
                    {detail}
                    {transfer && <div style={{ marginTop: 2, fontSize: 11, color: C.orange, fontWeight: 600 }}>{transfer}</div>}
                  </td>
                  <td className="somip-mono" style={{ textAlign: "right", color: meta.color, fontWeight: 600 }}>{meta.sign} {fmt(m.quantity)} L</td>
                  {canManage && <td style={{ textAlign: "right" }}><ConfirmIconButton onConfirm={() => deleteMovement(m.id)} /></td>}
                </tr>
              );
            })}
            {existingInv && (
              <tr>
                <td><Badge color={C.blue}>Stock fin</Badge></td>
                <td style={{ color: C.sub }}>
                  Jauge mesurée{existingInv.commentaire ? ` — ${existingInv.commentaire}` : ""}
                  {existingInv.photoUrls?.length > 0 && (
                    <div style={{ display: "flex", gap: 4, marginTop: 4 }}>
                      {existingInv.photoUrls.map((url, i) => (
                        <a key={i} href={url} target="_blank" rel="noopener noreferrer"><img src={url} alt="" style={{ width: 32, height: 32, objectFit: "cover", borderRadius: 4, border: `1px solid ${C.border}` }} /></a>
                      ))}
                    </div>
                  )}
                </td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600 }}>{fmt(existingInv.stockPhysique)} L</td>
                {canManage && <td style={{ textAlign: "right" }}><ConfirmIconButton onConfirm={() => deleteInventaire(existingInv)} /></td>}
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}


function ReceptionsView({ sites, movements, addMovement, deleteMovement, canWrite, canManage }) {
  const [form, setForm] = useState({ siteId: sites[0]?.id || "", date: todayStr(), quantity: "", ref: "", commentaire: "" });
  const [filterSite, setFilterSite] = useState("all");
  const [tempC, setTempC] = useState("");
  const [densite, setDensite] = useState("");

  const vcfResult = correctVolumeTo15({
    volumeAmbiant: Number(form.quantity) || 0,
    tempC: tempC === "" ? NaN : Number(tempC),
    densiteObservee: Number(densite) || 0,
  });

  const submit = () => {
    if (!form.siteId || !form.date || !form.quantity || Number(form.quantity) <= 0) return;
    const extra = vcfResult
      ? { temperatureC: Number(tempC), densiteObservee: Number(densite), densite15: vcfResult.densite15, vcf: vcfResult.vcf, volumeCorrige15: vcfResult.volume15 }
      : {};
    addMovement({ siteId: form.siteId, type: "reception", date: form.date, quantity: Number(form.quantity), delta: Number(form.quantity), ref: form.ref, commentaire: form.commentaire, ...extra });
    setForm({ ...form, quantity: "", ref: "", commentaire: "" });
    setTempC(""); setDensite("");
  };

  const list = movements.filter((m) => m.type === "reception").filter((m) => filterSite === "all" || m.siteId === filterSite).sort((a, b) => (a.date < b.date ? 1 : -1));

  return (
    <div className="somip-fade" style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      {canWrite && (
        <div className="somip-panel" style={{ flex: "1 1 280px", padding: 18 }}>
          <h3 style={{ margin: "0 0 14px", fontSize: 14 }}>Nouvelle réception</h3>
          <Field label="Site">
            <select className="somip-select" value={form.siteId} onChange={(e) => setForm({ ...form, siteId: e.target.value })}>
              {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
          </Field>
          <Field label="Date"><input type="date" className="somip-input" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></Field>
          <Field label="Quantité reçue (L, ambiant)"><input type="number" className="somip-input" value={form.quantity} onChange={(e) => setForm({ ...form, quantity: e.target.value })} placeholder="10000" /></Field>
          <Field label="Référence bon de livraison"><input className="somip-input" value={form.ref} onChange={(e) => setForm({ ...form, ref: e.target.value })} placeholder="BL-XXXX" /></Field>
          <VcfMiniPanel tempC={tempC} densite={densite} onTempC={setTempC} onDensite={setDensite} result={vcfResult} compact />
          <Field label="Commentaire (optionnel)"><textarea className="somip-textarea" rows={2} value={form.commentaire} onChange={(e) => setForm({ ...form, commentaire: e.target.value })} /></Field>
          <button className="somip-btn somip-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={submit} disabled={!form.quantity || Number(form.quantity) <= 0}>
            <Plus size={15} /> Enregistrer la réception
          </button>
        </div>
      )}

      <div className="somip-panel" style={{ flex: "2 1 480px", padding: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <h3 style={{ margin: 0, fontSize: 14 }}>Historique des réceptions</h3>
          <select className="somip-select" style={{ width: 200 }} value={filterSite} onChange={(e) => setFilterSite(e.target.value)}>
            <option value="all">Tous les sites</option>
            {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </div>
        <table className="somip-table">
          <thead><tr><th>Date</th><th>Site</th><th>Référence</th><th style={{ textAlign: "right" }}>Quantité (ambiant)</th><th style={{ textAlign: "right" }}>Volume 15°C</th>{canManage && <th></th>}</tr></thead>
          <tbody>
            {list.length === 0 && <EmptyRow colSpan={canManage ? 6 : 5} text="Aucune réception enregistrée." />}
            {list.map((m) => {
              const site = sites.find((s) => s.id === m.siteId);
              return (
                <tr key={m.id}>
                  <td className="somip-mono">{m.date}</td>
                  <td>{site?.name}{m.isDemo && <DemoBadge />}</td>
                  <td style={{ color: C.sub }}>{m.ref || "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success, fontWeight: 600 }}>+ {fmt(m.quantity)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{m.volumeCorrige15 ? `${fmt(m.volumeCorrige15)} L` : "—"}</td>
                  {canManage && <td style={{ textAlign: "right" }}><ConfirmIconButton onConfirm={() => deleteMovement(m.id)} /></td>}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Sorties                                                               */
/* ------------------------------------------------------------------ */
function SortiesView({ sites, movements, addMovement, deleteMovement, canWrite, canManage }) {
  const [tab, setTab] = useState("sortie");
  const [filterSite, setFilterSite] = useState("all");
  const [form, setForm] = useState({ siteId: sites[0]?.id || "", date: todayStr(), destinataire: "", camion: TRUCKS[0], destination: "", indexAvant: "", indexApres: "", commentaire: "" });
  const [tempC, setTempC] = useState("");
  const [densite, setDensite] = useState("");
  const isReturn = tab === "retour_camion";

  // La quantité n'est pas saisie : elle est calculée depuis le compteur (Index après − Index avant).
  const quantity = form.indexAvant !== "" && form.indexApres !== "" ? Number(form.indexApres) - Number(form.indexAvant) : 0;
  const indexValid = form.indexAvant !== "" && form.indexApres !== "" && quantity > 0;

  // Ventes, chargements et retours camions partagent le même compteur sur certains sites (Prehomo, Okouma...) :
  // on retrouve le dernier index enregistré (tous types confondus) pour repérer une rupture de séquence.
  const lastIndexForSite = movements
    .filter((m) => m.siteId === form.siteId && (m.type === "sortie" || m.type === "sortie_camion" || m.type === "retour_camion") && m.indexApres !== undefined)
    .sort((a, b) => (a.date + (a.createdAt || "")).localeCompare(b.date + (b.createdAt || "")))
    .slice(-1)[0]?.indexApres;
  const indexMismatch = lastIndexForSite !== undefined && form.indexAvant !== "" && Number(form.indexAvant) !== lastIndexForSite;

  const vcfResult = correctVolumeTo15({
    volumeAmbiant: quantity,
    tempC: tempC === "" ? NaN : Number(tempC),
    densiteObservee: Number(densite) || 0,
  });

  const submit = () => {
    if (!form.siteId || !form.date || !indexValid) return;
    const extra = vcfResult
      ? { temperatureC: Number(tempC), densiteObservee: Number(densite), densite15: vcfResult.densite15, vcf: vcfResult.vcf, volumeCorrige15: vcfResult.volume15 }
      : {};
    const base = {
      siteId: form.siteId, type: tab, date: form.date, quantity, delta: isReturn ? quantity : -quantity, commentaire: form.commentaire,
      indexAvant: Number(form.indexAvant), indexApres: Number(form.indexApres), ...extra,
    };
    const payload = tab === "sortie" ? { ...base, destinataire: form.destinataire } : { ...base, camion: form.camion, destination: form.destination };
    addMovement(payload);
    setForm({ ...form, destinataire: "", destination: "", indexAvant: "", indexApres: "", commentaire: "" });
    setTempC(""); setDensite("");
  };

  const list = movements.filter((m) => m.type === "sortie" || m.type === "sortie_camion" || m.type === "retour_camion").filter((m) => filterSite === "all" || m.siteId === filterSite).sort((a, b) => (a.date < b.date ? 1 : -1));

  return (
    <div className="somip-fade" style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      {canWrite && (
        <div className="somip-panel" style={{ flex: "1 1 300px", padding: 18 }}>
          <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
            <button className={`somip-tab ${tab === "sortie" ? "active" : ""}`} onClick={() => setTab("sortie")}>Sortie standard</button>
            <button className={`somip-tab ${tab === "sortie_camion" ? "active" : ""}`} onClick={() => setTab("sortie_camion")}>Vers camion laitier</button>
            <button className={`somip-tab ${tab === "retour_camion" ? "active" : ""}`} onClick={() => setTab("retour_camion")}>Retour camion (cuve)</button>
          </div>
          <Field label="Site">
            <select className="somip-select" value={form.siteId} onChange={(e) => setForm({ ...form, siteId: e.target.value })}>
              {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
          </Field>
          <Field label="Date"><input type="date" className="somip-input" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></Field>
          <div style={{ display: "flex", gap: 8 }}>
            <div style={{ flex: 1 }}><Field label="Index avant (compteur)"><input type="number" className="somip-input" value={form.indexAvant} onChange={(e) => setForm({ ...form, indexAvant: e.target.value })} placeholder="Ex : 45210" /></Field></div>
            <div style={{ flex: 1 }}><Field label="Index après (compteur)"><input type="number" className="somip-input" value={form.indexApres} onChange={(e) => setForm({ ...form, indexApres: e.target.value })} placeholder="Ex : 47210" /></Field></div>
          </div>
          {lastIndexForSite !== undefined && (
            <p style={{ margin: "-8px 0 8px", fontSize: 11, color: indexMismatch ? C.warning : C.sub }}>
              Dernier index enregistré sur ce site (ventes + camions) : {fmt(lastIndexForSite)}
              {indexMismatch && " — vérifie ton index avant, il ne correspond pas."}
            </p>
          )}
          <div style={{ background: C.bg, borderRadius: 8, padding: "9px 12px", marginBottom: 12, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ fontSize: 12.5, color: C.sub, fontWeight: 600 }}>{isReturn ? "Quantité retournée (calculée)" : "Quantité sortie (calculée)"}</span>
            <span className="somip-mono" style={{ fontWeight: 700, color: form.indexAvant !== "" && form.indexApres !== "" && quantity <= 0 ? C.danger : (isReturn ? C.success : C.orange) }}>{fmt(quantity)} L</span>
          </div>
          {form.indexAvant !== "" && form.indexApres !== "" && quantity <= 0 && (
            <p style={{ margin: "-8px 0 12px", fontSize: 11.5, color: C.danger }}>L'index après doit être supérieur à l'index avant.</p>
          )}
          {tab === "sortie" ? (
            <Field label="Destinataire / motif"><input className="somip-input" value={form.destinataire} onChange={(e) => setForm({ ...form, destinataire: e.target.value })} placeholder="Ex : Atelier, Engin X..." /></Field>
          ) : (
            <>
              <Field label="Camion laitier">
                <select className="somip-select" value={form.camion} onChange={(e) => setForm({ ...form, camion: e.target.value })}>
                  {TRUCKS.map((t) => <option key={t} value={t}>{t}</option>)}
                </select>
              </Field>
              <Field label={isReturn ? "Provenance / motif du retour" : "Destination (carrière / engin)"}>
                <input className="somip-input" value={form.destination} onChange={(e) => setForm({ ...form, destination: e.target.value })} placeholder={isReturn ? "Ex : Reliquat Carrière Nord" : "Ex : Carrière Nord"} />
              </Field>
            </>
          )}
          <Field label="Commentaire (optionnel)"><textarea className="somip-textarea" rows={2} value={form.commentaire} onChange={(e) => setForm({ ...form, commentaire: e.target.value })} /></Field>
          <VcfMiniPanel tempC={tempC} densite={densite} onTempC={setTempC} onDensite={setDensite} result={vcfResult} compact />
          <button className={`somip-btn ${isReturn ? "somip-btn-primary" : "somip-btn-secondary"}`} style={{ width: "100%", justifyContent: "center" }} onClick={submit} disabled={!indexValid}>
            <Plus size={15} /> {isReturn ? "Enregistrer le retour" : "Enregistrer la sortie"}
          </button>
        </div>
      )}

      <div className="somip-panel" style={{ flex: "2 1 480px", padding: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <h3 style={{ margin: 0, fontSize: 14 }}>Historique des sorties &amp; retours</h3>
          <select className="somip-select" style={{ width: 200 }} value={filterSite} onChange={(e) => setFilterSite(e.target.value)}>
            <option value="all">Tous les sites</option>
            {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead><tr><th>Date</th><th>Site</th><th>Type</th><th>Détail</th><th style={{ textAlign: "right" }}>Index avant→après</th><th style={{ textAlign: "right" }}>Quantité (ambiant)</th><th style={{ textAlign: "right" }}>Volume 15°C</th>{canManage && <th></th>}</tr></thead>
            <tbody>
              {list.length === 0 && <EmptyRow colSpan={canManage ? 8 : 7} text="Aucune sortie enregistrée." />}
              {list.map((m) => {
                const site = sites.find((s) => s.id === m.siteId);
                const meta = TYPE_META[m.type];
                const typeLabel = m.type === "sortie" ? "Standard" : m.type === "sortie_camion" ? "Camion" : "Retour";
                const detail = m.type === "sortie" ? (m.destinataire || "—") : `${m.camion} ${m.type === "retour_camion" ? "←" : "→"} ${m.destination || "—"}`;
                return (
                  <tr key={m.id}>
                    <td className="somip-mono">{m.date}</td>
                    <td>{site?.name}{m.isDemo && <DemoBadge />}</td>
                    <td><Badge color={meta.color}>{typeLabel}</Badge></td>
                    <td style={{ color: C.sub }}>{detail}</td>
                    <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{m.indexAvant !== undefined && m.indexApres !== undefined ? `${fmt(m.indexAvant)} → ${fmt(m.indexApres)}` : "—"}</td>
                    <td className="somip-mono" style={{ textAlign: "right", color: meta.color, fontWeight: 600 }}>{meta.sign} {fmt(m.quantity)} L</td>
                    <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{m.volumeCorrige15 ? `${fmt(m.volumeCorrige15)} L` : "—"}</td>
                    {canManage && <td style={{ textAlign: "right" }}><ConfirmIconButton onConfirm={() => deleteMovement(m.id)} /></td>}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Inventaires                                                           */
/* ------------------------------------------------------------------ */
function InventairesView({ sites, inventaires, stockOf, stockOf15, addInventaire, deleteInventaire, settings, updateSettings, canWrite, canManage, canInventaireOfficiel, inventairesOfficiels, addInventaireOfficiel, deleteInventaireOfficiel, siteTanks, siteDepotageMeters, siteMeters, signInventaireOfficiel, canSignSomip, canSignOperateur, canSignTotal, isTotalEnergiesOnly }) {
  const [mainTab, setMainTab] = useState(isTotalEnergiesOnly ? "officiel" : "rapide");
  const [siteId, setSiteId] = useState(sites[0]?.id || "");
  const [date, setDate] = useState(todayStr());
  const [stockPhysique, setStockPhysique] = useState("");
  const [commentaire, setCommentaire] = useState("");
  const [filterSite, setFilterSite] = useState("all");
  const [objectifDraft, setObjectifDraft] = useState(settings.objectifFreinte);
  const [tempC, setTempC] = useState("");
  const [densite, setDensite] = useState("");

  useEffect(() => { setObjectifDraft(settings.objectifFreinte); }, [settings.objectifFreinte]);

  const theoriqueAmbiant = stockOf(siteId);
  const theorique15 = stockOf15(siteId);
  const physiqueNum = Number(stockPhysique) || 0;
  const vcfResult = correctVolumeTo15({
    volumeAmbiant: physiqueNum,
    tempC: tempC === "" ? NaN : Number(tempC),
    densiteObservee: Number(densite) || 0,
  });
  const has15 = !!vcfResult;
  const theoriqueUsed = has15 ? theorique15 : theoriqueAmbiant;
  const physiqueUsed = has15 ? vcfResult.volume15 : physiqueNum;
  const ecart = stockPhysique === "" ? null : physiqueUsed - theoriqueUsed;
  const preview = ecart === null ? null : classifyEcart(ecart, theoriqueUsed, settings.objectifFreinte);

  const submit = () => {
    if (!siteId || !date || stockPhysique === "") return;
    const extra = vcfResult
      ? { temperatureC: Number(tempC), densiteObservee: Number(densite), densite15: vcfResult.densite15, vcf: vcfResult.vcf, stockPhysique15: vcfResult.volume15 }
      : {};
    addInventaire({ siteId, date, stockPhysique: physiqueNum, commentaire, ...extra });
    setStockPhysique(""); setCommentaire(""); setTempC(""); setDensite("");
  };

  const list = inventaires.filter((i) => filterSite === "all" || i.siteId === filterSite).sort((a, b) => (a.date < b.date ? 1 : -1));

  return (
    <div className="somip-fade">
      <div className="somip-no-print" style={{ display: "flex", gap: 8, marginBottom: 18 }}>
        {!isTotalEnergiesOnly && <button className={`somip-tab ${mainTab === "rapide" ? "active" : ""}`} onClick={() => setMainTab("rapide")}>Inventaire rapide</button>}
        <button className={`somip-tab ${mainTab === "officiel" ? "active" : ""}`} onClick={() => setMainTab("officiel")}>Inventaire officiel (inopiné / mensuel)</button>
      </div>

      {mainTab === "rapide" && !isTotalEnergiesOnly && (
      <div style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      <div style={{ display: "flex", flexDirection: "column", gap: 18, flex: "1 1 300px" }}>
        {canWrite && (
        <div className="somip-panel" style={{ padding: 18 }}>
          <h3 style={{ margin: "0 0 14px", fontSize: 14 }}>Nouvel inventaire</h3>
          <Field label="Site">
            <select className="somip-select" value={siteId} onChange={(e) => setSiteId(e.target.value)}>
              {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
          </Field>
          <Field label="Date"><input type="date" className="somip-input" value={date} onChange={(e) => setDate(e.target.value)} /></Field>
          <Field label="Stock physique mesuré (L, ambiant)"><input type="number" className="somip-input" value={stockPhysique} onChange={(e) => setStockPhysique(e.target.value)} placeholder="Ex : 12450" /></Field>
          <VcfMiniPanel tempC={tempC} densite={densite} onTempC={setTempC} onDensite={setDensite} result={vcfResult} compact />
          <Field label="Commentaire (optionnel)"><textarea className="somip-textarea" rows={2} value={commentaire} onChange={(e) => setCommentaire(e.target.value)} /></Field>

          <div style={{ background: C.bg, borderRadius: 8, padding: 12, marginBottom: 10, fontSize: 12.5 }}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
              <span style={{ color: C.sub, fontWeight: 600 }}>Base de calcul de l'écart</span>
              <Badge color={has15 ? C.blue : C.sub}>{has15 ? "Volume à 15°C" : "Ambiant (température/densité non fournies)"}</Badge>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
              <span style={{ color: C.sub }}>Théorique ambiant</span>
              <span className="somip-mono" style={{ color: C.sub }}>{fmt(theoriqueAmbiant)} L</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ color: C.sub }}>Théorique à 15°C</span>
              <span className="somip-mono" style={{ color: C.sub }}>{fmt(theorique15)} L</span>
            </div>
            <div style={{ height: 1, background: C.border, margin: "8px 0" }} />
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ color: C.sub }}>Théorique retenu ({has15 ? "15°C" : "ambiant"})</span>
              <span className="somip-mono" style={{ fontWeight: 600 }}>{fmt(theoriqueUsed)} L</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ color: C.sub }}>Physique retenu ({has15 ? "15°C" : "ambiant"})</span>
              <span className="somip-mono" style={{ fontWeight: 600 }}>{stockPhysique === "" ? "—" : `${fmt(physiqueUsed)} L`}</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ color: C.sub }}>Nature de l'écart</span>
              {preview ? <Badge color={NATURE_META[preview.nature].color}>{NATURE_META[preview.nature].label}</Badge> : <span style={{ color: C.sub }}>—</span>}
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ color: C.sub }}>Écart en litres</span>
              <span className="somip-mono" style={{ fontWeight: 600, color: preview ? NATURE_META[preview.nature].color : C.ink }}>{preview ? `${preview.ecartL >= 0 ? "+" : ""}${fmt(preview.ecartL)} L` : "—"}</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ color: C.sub }}>Écart en ‰</span>
              <span className="somip-mono" style={{ fontWeight: 600, color: preview ? NATURE_META[preview.nature].color : C.ink }}>{preview ? `${preview.ecartPermille >= 0 ? "+" : ""}${preview.ecartPermille.toFixed(2)} ‰` : "—"}</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ color: C.sub }}>Taux de freinte</span>
              <span className="somip-mono" style={{ fontWeight: 600 }}>{preview && preview.nature === "perte" ? `${preview.tauxFreinte.toFixed(2)} ‰` : "—"}</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ color: C.sub }}>Objectif SOMIP</span>
              <span className="somip-mono" style={{ fontWeight: 600 }}>{settings.objectifFreinte.toFixed(1)} ‰</span>
            </div>
            {preview && (
              <div style={{ marginTop: 8 }}>
                <Badge color={CONFORMITE_META[preview.conformite].color}>{CONFORMITE_META[preview.conformite].label}</Badge>
              </div>
            )}
          </div>

          <p style={{ margin: "0 0 12px", fontSize: 11, color: C.sub }}>
            Dès que Température et Densité sont renseignées, la perte/gain, l'écart et le taux de freinte sont calculés sur le volume corrigé à 15°C (théorique 15°C vs physique 15°C). Sans ces mesures, le calcul reste en litres ambiants.
          </p>

          <button className="somip-btn somip-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={submit} disabled={stockPhysique === ""}>
            <Plus size={15} /> Valider l'inventaire
          </button>
        </div>
        )}

        {canManage && (
        <div className="somip-panel" style={{ padding: 18 }}>
          <h3 style={{ margin: "0 0 6px", fontSize: 14 }}>Objectif de freinte</h3>
          <p style={{ margin: "0 0 12px", fontSize: 12, color: C.sub }}>
            Seuil de tolérance appliqué aux pertes (taux de freinte). Standard SOMIP : 3 ‰ (3/1000). Modifiable si nécessaire.
          </p>
          <div style={{ display: "flex", gap: 8 }}>
            <input type="number" step="0.1" className="somip-input" value={objectifDraft} onChange={(e) => setObjectifDraft(e.target.value)} />
            <button className="somip-btn somip-btn-ghost" style={{ whiteSpace: "nowrap" }} onClick={() => updateSettings({ objectifFreinte: Number(objectifDraft) || 0 })} disabled={Number(objectifDraft) === settings.objectifFreinte}>
              Enregistrer
            </button>
          </div>
        </div>
        )}
      </div>

      <div className="somip-panel" style={{ flex: "2 1 560px", padding: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <h3 style={{ margin: 0, fontSize: 14 }}>Historique des inventaires</h3>
          <select className="somip-select" style={{ width: 200 }} value={filterSite} onChange={(e) => setFilterSite(e.target.value)}>
            <option value="all">Tous les sites</option>
            {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead>
              <tr>
                <th>Date</th><th>Site</th><th>Base</th><th>Nature</th>
                <th style={{ textAlign: "right" }}>Théorique</th><th style={{ textAlign: "right" }}>Physique</th>
                <th style={{ textAlign: "right" }}>Écart (L)</th><th style={{ textAlign: "right" }}>Écart (‰)</th>
                <th style={{ textAlign: "right" }}>Taux de freinte</th><th>Statut (objectif)</th>{canManage && <th></th>}
              </tr>
            </thead>
            <tbody>
              {list.length === 0 && <EmptyRow colSpan={canManage ? 11 : 10} text="Aucun inventaire enregistré." />}
              {list.map((i) => {
                const site = sites.find((s) => s.id === i.siteId);
                const nature = i.nature || (i.ecart === 0 ? "neutre" : i.ecart < 0 ? "perte" : "gain");
                const conformite = i.conformite || "conforme";
                const basis15 = i.basisEcart === "15c";
                const theoriqueAff = i.stockTheorique !== undefined ? i.stockTheorique : i.stockTheoriqueAmbiant;
                const physiqueAff = i.stockPhysiqueUsed !== undefined ? i.stockPhysiqueUsed : i.stockPhysique;
                return (
                  <tr key={i.id}>
                    <td className="somip-mono">{i.date}</td>
                    <td>{site?.name}</td>
                    <td><Badge color={basis15 ? C.blue : C.sub}>{basis15 ? "15°C" : "Ambiant"}</Badge></td>
                    <td><Badge color={NATURE_META[nature].color}>{NATURE_META[nature].label}</Badge></td>
                    <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{fmt(theoriqueAff)} L</td>
                    <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{fmt(physiqueAff)} L</td>
                    <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600, color: NATURE_META[nature].color }}>{i.ecart >= 0 ? "+" : ""}{fmt(i.ecart)} L</td>
                    <td className="somip-mono" style={{ textAlign: "right" }}>{i.ecartPermille >= 0 ? "+" : ""}{i.ecartPermille.toFixed(2)} ‰</td>
                    <td className="somip-mono" style={{ textAlign: "right" }}>{nature === "perte" ? `${i.tauxFreinte.toFixed(2)} ‰` : "—"}</td>
                    <td><Badge color={CONFORMITE_META[conformite].color}>{CONFORMITE_META[conformite].label}</Badge></td>
                    {canManage && <td style={{ textAlign: "right" }}><ConfirmIconButton onConfirm={() => deleteInventaire(i)} /></td>}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
      </div>
      )}

      {mainTab === "officiel" && (
        <InventaireOfficielTab sites={sites} siteTanks={siteTanks} siteDepotageMeters={siteDepotageMeters} siteMeters={siteMeters} inventairesOfficiels={inventairesOfficiels} addInventaireOfficiel={addInventaireOfficiel} deleteInventaireOfficiel={deleteInventaireOfficiel} canWrite={canInventaireOfficiel} canManage={canManage} signInventaireOfficiel={signInventaireOfficiel} canSignSomip={canSignSomip} canSignOperateur={canSignOperateur} canSignTotal={canSignTotal} isTotalEnergiesOnly={isTotalEnergiesOnly} />
      )}
    </div>
  );
}

/* ---- Inventaire officiel (inopiné / mensuel) — cuve par cuve, avec PDF signé ---- */
function InventaireOfficielTab({ sites, siteTanks, siteDepotageMeters, siteMeters, inventairesOfficiels, addInventaireOfficiel, deleteInventaireOfficiel, canWrite, canManage, signInventaireOfficiel, canSignSomip, canSignOperateur, canSignTotal, isTotalEnergiesOnly }) {
  const fixedSites = sites.filter((s) => !s.isMobile);
  const [siteId, setSiteId] = useState(fixedSites[0]?.id || "");
  const [date, setDate] = useState(todayStr());
  const [type, setType] = useState("mensuel");
  const [produit, setProduit] = useState("gasoil");
  const [inventoriste, setInventoriste] = useState("");
  const [operateur, setOperateur] = useState("");
  const [commentaire, setCommentaire] = useState("");
  const [filterSite, setFilterSite] = useState("all");
  const [signingId, setSigningId] = useState(null);
  const [signingRole, setSigningRole] = useState(null);
  const emptyCuve = (name) => ({ cuve: name, hauteur: "", densite: "", temperatureC: "", eau: false, stockAmbiant: "" });
  const tanksForSite = siteTanks.filter((t) => t.siteId === siteId);
  const [cuveReadings, setCuveReadings] = useState([]);
  useEffect(() => {
    setCuveReadings(tanksForSite.length ? tanksForSite.map((t) => emptyCuve(t.name)) : [emptyCuve("Cuve 1")]);
  }, [siteId]); // eslint-disable-line react-hooks/exhaustive-deps

  // Volume à 15°C calculé CUVE PAR CUVE (densité/température propres à chaque cuve).
  const cuvesComputed = cuveReadings.map((c) => {
    const amb = Number(c.stockAmbiant) || 0;
    let volume15 = null;
    if (c.temperatureC !== "" && c.densite !== "") {
      const corr = correctVolumeTo15({ volumeAmbiant: amb, tempC: Number(c.temperatureC), densiteObservee: Number(c.densite) });
      if (corr) volume15 = corr.volume15;
    }
    return { ...c, stockAmbiantNum: amb, volume15 };
  });
  const stockAmbiantTotal = cuvesComputed.reduce((a, c) => a + c.stockAmbiantNum, 0);
  const allHave15 = cuvesComputed.length > 0 && cuvesComputed.every((c) => c.volume15 !== null);
  const stock15Total = allHave15 ? cuvesComputed.reduce((a, c) => a + c.volume15, 0) : null;

  const updateCuve = (idx, field, value) => setCuveReadings((prev) => prev.map((c, i) => (i === idx ? { ...c, [field]: value } : c)));
  const addCuveRow = () => setCuveReadings((prev) => [...prev, emptyCuve(`Cuve ${prev.length + 1}`)]);
  const removeCuveRow = (idx) => setCuveReadings((prev) => prev.filter((_, i) => i !== idx));

  const emptyDepotage = (name) => ({ compteur: name, indexFin: "" });
  const depotageMetersForSite = siteDepotageMeters.filter((m) => m.siteId === siteId);
  const [depotageReadings, setDepotageReadings] = useState([]);
  useEffect(() => {
    setDepotageReadings(depotageMetersForSite.length ? depotageMetersForSite.map((m) => emptyDepotage(m.name)) : []);
  }, [siteId]); // eslint-disable-line react-hooks/exhaustive-deps
  const updateDepotage = (idx, field, value) => setDepotageReadings((prev) => prev.map((d, i) => (i === idx ? { ...d, [field]: value } : d)));
  const addDepotageRow = () => setDepotageReadings((prev) => [...prev, emptyDepotage(`Dépotage ${prev.length + 1}`)]);
  const removeDepotageRow = (idx) => setDepotageReadings((prev) => prev.filter((_, i) => i !== idx));

  // Index (relevé compteur) — liste indépendante des cuves, basée sur les compteurs de vente
  // habituels du site (les mêmes que pour la Saisie journalière).
  const emptyIndex = (name) => ({ compteur: name, indexFin: "" });
  const currentSite = sites.find((s) => s.id === siteId);
  const salesMetersForSite = metersForSite(currentSite, siteMeters);
  const [indexReadings, setIndexReadings] = useState([]);
  useEffect(() => {
    setIndexReadings(salesMetersForSite.map((name) => emptyIndex(name)));
  }, [siteId]); // eslint-disable-line react-hooks/exhaustive-deps
  const updateIndex = (idx, field, value) => setIndexReadings((prev) => prev.map((d, i) => (i === idx ? { ...d, [field]: value } : d)));
  const addIndexRow = () => setIndexReadings((prev) => [...prev, emptyIndex(`Compteur ${prev.length + 1}`)]);
  const removeIndexRow = (idx) => setIndexReadings((prev) => prev.filter((_, i) => i !== idx));

  const canSubmit = siteId && date && cuveReadings.length > 0 && cuveReadings.every((c) => c.stockAmbiant !== "");

  const submit = () => {
    if (!canSubmit) return;
    addInventaireOfficiel({
      siteId, date, type, produit, inventoriste, operateur,
      cuves: cuveReadings.map((c) => ({
        cuve: c.cuve, hauteur: c.hauteur === "" ? null : Number(c.hauteur), densite: c.densite === "" ? null : Number(c.densite),
        temperatureC: c.temperatureC === "" ? null : Number(c.temperatureC), eau: !!c.eau, stockAmbiant: Number(c.stockAmbiant) || 0,
      })),
      depotage: depotageReadings.map((d) => ({ compteur: d.compteur, indexFin: d.indexFin === "" ? null : Number(d.indexFin) })),
      indexCompteurs: indexReadings.map((d) => ({ compteur: d.compteur, indexFin: d.indexFin === "" ? null : Number(d.indexFin) })),
      commentaire,
    });
    setCuveReadings(tanksForSite.length ? tanksForSite.map((t) => emptyCuve(t.name)) : [emptyCuve("Cuve 1")]);
    setDepotageReadings(depotageMetersForSite.length ? depotageMetersForSite.map((m) => emptyDepotage(m.name)) : []);
    setIndexReadings(salesMetersForSite.map((name) => emptyIndex(name)));
    setInventoriste(""); setOperateur(""); setCommentaire("");
  };

  const list = inventairesOfficiels.filter((i) => filterSite === "all" || i.siteId === filterSite).sort((a, b) => (a.date < b.date ? 1 : -1));

  return (
    <div style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      {canWrite && (
        <div className="somip-panel" style={{ padding: 18, flex: "1 1 460px" }}>
          <h3 style={{ margin: "0 0 14px", fontSize: 14 }}>Nouvel inventaire officiel</h3>
          <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
            <button className={`somip-tab ${type === "mensuel" ? "active" : ""}`} style={{ flex: 1, textAlign: "center", fontSize: 12.5 }} onClick={() => setType("mensuel")}>Mensuel</button>
            <button className={`somip-tab ${type === "inopine" ? "active" : ""}`} style={{ flex: 1, textAlign: "center", fontSize: 12.5 }} onClick={() => setType("inopine")}>Inopiné</button>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <div style={{ flex: 1 }}>
              <Field label="Site">
                <select className="somip-select" value={siteId} onChange={(e) => setSiteId(e.target.value)}>
                  {fixedSites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
              </Field>
            </div>
            <div style={{ flex: 1 }}><Field label="Date"><input type="date" className="somip-input" value={date} onChange={(e) => setDate(e.target.value)} /></Field></div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <div style={{ flex: 1 }}>
              <Field label="Produit">
                <select className="somip-select" value={produit} onChange={(e) => setProduit(e.target.value)}>
                  <option value="gasoil">Gasoil</option>
                  <option value="lubrifiant_vrac">Lubrifiant vrac</option>
                </select>
              </Field>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <div style={{ flex: 1 }}><Field label="Inventoriste"><input className="somip-input" value={inventoriste} onChange={(e) => setInventoriste(e.target.value)} placeholder="Nom" /></Field></div>
            <div style={{ flex: 1 }}><Field label="Opérateur"><input className="somip-input" value={operateur} onChange={(e) => setOperateur(e.target.value)} placeholder="Nom" /></Field></div>
          </div>

          <p style={{ margin: "12px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Relevé cuve par cuve</p>
          {tanksForSite.length === 0 && (
            <p style={{ margin: "0 0 10px", fontSize: 11.5, color: C.warning }}>Aucune cuve configurée pour ce site — ajoute-les depuis la page Sites, ou saisis-les directement ci-dessous.</p>
          )}
          {cuvesComputed.map((c, idx) => (
            <div key={idx} style={{ border: `1px solid ${C.border}`, borderRadius: 8, padding: 10, marginBottom: 8 }}>
              <div style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "flex-end" }}>
                <div style={{ flex: 1.3 }}>
                  <Field label="Cuve"><input className="somip-input" value={c.cuve} onChange={(e) => updateCuve(idx, "cuve", e.target.value)} placeholder="Cuve 1" /></Field>
                </div>
                <div style={{ flex: 1 }}>
                  <Field label="Hauteur (mm)"><input type="number" className="somip-input" value={c.hauteur} onChange={(e) => updateCuve(idx, "hauteur", e.target.value)} placeholder="0" /></Field>
                </div>
                <div style={{ flex: 1 }}>
                  <Field label="Densité"><input type="number" step="0.001" className="somip-input" value={c.densite} onChange={(e) => updateCuve(idx, "densite", e.target.value)} placeholder="0.840" /></Field>
                </div>
                <div style={{ flex: 1 }}>
                  <Field label="Température (°C)"><input type="number" className="somip-input" value={c.temperatureC} onChange={(e) => updateCuve(idx, "temperatureC", e.target.value)} placeholder="28" /></Field>
                </div>
                {cuveReadings.length > 1 && (
                  <button onClick={() => removeCuveRow(idx)} style={{ border: "none", background: "none", cursor: "pointer", padding: "9px 4px" }}>
                    <X size={16} color={C.danger} />
                  </button>
                )}
              </div>
              <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
                <div style={{ flex: 1 }}>
                  <Field label="Stock (L, ambiant)"><input type="number" className="somip-input" value={c.stockAmbiant} onChange={(e) => updateCuve(idx, "stockAmbiant", e.target.value)} placeholder="0" /></Field>
                </div>
                <div style={{ flex: 1, paddingBottom: 9 }}>
                  <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, cursor: "pointer" }}>
                    <input type="checkbox" checked={c.eau} onChange={(e) => updateCuve(idx, "eau", e.target.checked)} />
                    Présence d'eau
                  </label>
                </div>
                <div style={{ flex: 1, textAlign: "right", paddingBottom: 9 }}>
                  <span style={{ fontSize: 11, color: C.sub, display: "block" }}>Volume à 15°C</span>
                  <span className="somip-mono" style={{ fontWeight: 700, color: c.volume15 !== null ? C.blue : C.sub }}>{c.volume15 !== null ? `${fmt(c.volume15)} L` : "—"}</span>
                </div>
              </div>
            </div>
          ))}
          <button className="somip-btn somip-btn-secondary" style={{ fontSize: 12, padding: "6px 12px", marginBottom: 12 }} onClick={addCuveRow}>
            <Plus size={13} /> Ajouter une cuve
          </button>

          <p style={{ margin: "16px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Index — Relevé compteur</p>
          <p style={{ margin: "0 0 8px", fontSize: 11, color: C.sub }}>Indépendant des cuves — les compteurs habituels du site (mêmes que pour la Saisie journalière).</p>
          {indexReadings.map((d, idx) => (
            <div key={idx} style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "flex-end" }}>
              <div style={{ flex: 1.2 }}>
                <Field label="Compteur"><input className="somip-input" value={d.compteur} onChange={(e) => updateIndex(idx, "compteur", e.target.value)} placeholder="Poste 1" /></Field>
              </div>
              <div style={{ flex: 1 }}>
                <Field label="Index fin"><input type="number" className="somip-input" value={d.indexFin} onChange={(e) => updateIndex(idx, "indexFin", e.target.value)} placeholder="0" /></Field>
              </div>
              <button onClick={() => removeIndexRow(idx)} style={{ border: "none", background: "none", cursor: "pointer", padding: "9px 4px" }}>
                <X size={16} color={C.danger} />
              </button>
            </div>
          ))}
          <button className="somip-btn somip-btn-secondary" style={{ fontSize: 12, padding: "6px 12px", marginBottom: 12 }} onClick={addIndexRow}>
            <Plus size={13} /> Ajouter un compteur
          </button>

          <p style={{ margin: "16px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Index des compteurs de dépotage</p>
          {depotageMetersForSite.length === 0 && depotageReadings.length === 0 && (
            <p style={{ margin: "0 0 10px", fontSize: 11.5, color: C.warning }}>Aucun compteur de dépotage configuré pour ce site — ajoute-les depuis la page Sites, ou saisis-les directement ci-dessous.</p>
          )}
          {depotageReadings.map((d, idx) => (
            <div key={idx} style={{ display: "flex", gap: 8, marginBottom: 6, alignItems: "flex-end" }}>
              <div style={{ flex: 1.2 }}>
                <Field label="Compteur de dépotage"><input className="somip-input" value={d.compteur} onChange={(e) => updateDepotage(idx, "compteur", e.target.value)} placeholder="Dépotage 1" /></Field>
              </div>
              <div style={{ flex: 1 }}>
                <Field label="Index fin"><input type="number" className="somip-input" value={d.indexFin} onChange={(e) => updateDepotage(idx, "indexFin", e.target.value)} placeholder="0" /></Field>
              </div>
              <button onClick={() => removeDepotageRow(idx)} style={{ border: "none", background: "none", cursor: "pointer", padding: "9px 4px" }}>
                <X size={16} color={C.danger} />
              </button>
            </div>
          ))}
          <button className="somip-btn somip-btn-secondary" style={{ fontSize: 12, padding: "6px 12px", marginBottom: 12 }} onClick={addDepotageRow}>
            <Plus size={13} /> Ajouter un compteur de dépotage
          </button>

          <Field label="Commentaire (optionnel)"><textarea className="somip-textarea" rows={2} value={commentaire} onChange={(e) => setCommentaire(e.target.value)} /></Field>

          <div style={{ background: C.bg, borderRadius: 8, padding: 12, margin: "4px 0 14px", fontSize: 12.5 }}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
              <span style={{ color: C.sub, fontWeight: 600 }}>Stock total ambiant</span>
              <span className="somip-mono" style={{ fontWeight: 700 }}>{fmt(stockAmbiantTotal)} L</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <span style={{ color: C.sub, fontWeight: 600 }}>Stock total à 15°C</span>
              <span className="somip-mono" style={{ fontWeight: 700, color: stock15Total !== null ? C.blue : C.sub }}>{stock15Total !== null ? `${fmt(stock15Total)} L` : "— (densité/température manquantes sur au moins une cuve)"}</span>
            </div>
          </div>

          <button className="somip-btn somip-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={submit} disabled={!canSubmit}>
            <Plus size={15} /> Enregistrer l'inventaire officiel
          </button>
        </div>
      )}

      <div className="somip-panel" style={{ flex: "2 1 560px", padding: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
          <h3 style={{ margin: 0, fontSize: 14 }}>Historique — Inventaires officiels</h3>
          <select className="somip-select" style={{ width: 200 }} value={filterSite} onChange={(e) => setFilterSite(e.target.value)}>
            <option value="all">Tous les sites</option>
            {fixedSites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead>
              <tr>
                <th>Date</th><th>Site</th><th>Produit</th><th>Type</th><th>Inventoriste</th><th>Opérateur</th>
                <th style={{ textAlign: "right" }}>Stock ambiant</th><th style={{ textAlign: "right" }}>Stock 15°C</th><th>Signatures</th><th></th>{canManage && <th></th>}
              </tr>
            </thead>
            <tbody>
              {list.length === 0 && <EmptyRow colSpan={canManage ? 11 : 10} text="Aucun inventaire officiel enregistré." />}
              {list.map((inv) => (
                <tr key={inv.id}>
                  <td className="somip-mono">{inv.date}</td>
                  <td>{sites.find((s) => s.id === inv.siteId)?.name}</td>
                  <td style={{ color: C.sub }}>{PRODUIT_INVENTAIRE_LABELS[inv.produit] || inv.produit}</td>
                  <td><Badge color={inv.type === "inopine" ? C.orange : C.blue}>{TYPE_INVENTAIRE_LABELS[inv.type]}</Badge></td>
                  <td style={{ color: C.sub }}>{inv.inventoriste || "—"}</td>
                  <td style={{ color: C.sub }}>{inv.operateur || "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(inv.stockAmbiant)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: inv.stock15 !== undefined ? C.blue : C.sub }}>{inv.stock15 !== undefined ? `${fmt(inv.stock15)} L` : "—"}</td>
                  <td>
                    <button className="somip-btn somip-btn-ghost" style={{ fontSize: 11.5, padding: "4px 8px" }} onClick={() => setSigningId(signingId === inv.id ? null : inv.id)}>
                      <Pencil size={12} /> {[inv.signatureSomipUrl, inv.signatureOperateurUrl, inv.signatureTotalUrl].filter(Boolean).length}/3
                    </button>
                  </td>
                  <td>
                    <button className="somip-btn somip-btn-ghost" style={{ fontSize: 11.5, padding: "4px 8px" }} onClick={() => exportInventaireOfficielToPdf(inv, sites.find((s) => s.id === inv.siteId))}>
                      <Download size={12} /> PDF
                    </button>
                  </td>
                  {canManage && <td style={{ textAlign: "right" }}><ConfirmIconButton onConfirm={() => deleteInventaireOfficiel(inv)} /></td>}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {signingId && (() => {
          const inv = list.find((i) => i.id === signingId);
          if (!inv) return null;
          const slots = [
            { role: "somip", label: "SOMIP", url: inv.signatureSomipUrl, by: inv.signatureSomipBy, at: inv.signatureSomipAt, can: canSignSomip },
            { role: "operateur", label: "Opérateur", url: inv.signatureOperateurUrl, by: inv.signatureOperateurBy, at: inv.signatureOperateurAt, can: canSignOperateur },
            { role: "total", label: "TotalEnergies", url: inv.signatureTotalUrl, by: inv.signatureTotalBy, at: inv.signatureTotalAt, can: canSignTotal },
          ];
          return (
            <div className="somip-panel" style={{ marginTop: 14, padding: 14, background: C.bg }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                <p style={{ margin: 0, fontSize: 12.5, fontWeight: 700, color: C.ink }}>
                  Signatures — {sites.find((s) => s.id === inv.siteId)?.name} — {inv.date}
                </p>
                <button onClick={() => { setSigningId(null); setSigningRole(null); }} style={{ border: "none", background: "none", cursor: "pointer" }}><X size={16} color={C.sub} /></button>
              </div>
              <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
                {slots.map((slot) => (
                  <div key={slot.role} style={{ flex: "1 1 220px", border: `1px solid ${C.border}`, borderRadius: 8, padding: 10, background: "#fff" }}>
                    <p style={{ margin: "0 0 8px", fontSize: 12, fontWeight: 700, color: C.ink }}>{slot.label}</p>
                    {slot.url ? (
                      <>
                        <img src={slot.url} alt="" style={{ maxWidth: "100%", height: 60, objectFit: "contain", border: `1px solid ${C.border}`, borderRadius: 6, background: "#fff" }} />
                        <p style={{ margin: "6px 0 0", fontSize: 11, color: C.sub }}>Signé par {slot.by || "—"}{slot.at ? ` le ${new Date(slot.at).toLocaleString("fr-FR")}` : ""}</p>
                      </>
                    ) : slot.can && signingRole === slot.role ? (
                      <SignaturePad
                        onSave={(dataUrl) => { signInventaireOfficiel(inv, slot.role, dataUrl); setSigningRole(null); }}
                        onCancel={() => setSigningRole(null)}
                      />
                    ) : slot.can ? (
                      <button className="somip-btn somip-btn-primary" style={{ fontSize: 12 }} onClick={() => setSigningRole(slot.role)}>Signer</button>
                    ) : (
                      <p style={{ margin: 0, fontSize: 11, color: C.sub }}>Non signé</p>
                    )}
                  </div>
                ))}
              </div>
            </div>
          );
        })()}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Correction 15°C — calculateur autonome                               */
/* ------------------------------------------------------------------ */
function VcfView() {
  const [volumeAmbiant, setVolumeAmbiant] = useState("");
  const [tempC, setTempC] = useState("");
  const [densite, setDensite] = useState("");

  const result = correctVolumeTo15({
    volumeAmbiant: Number(volumeAmbiant) || 0,
    tempC: tempC === "" ? NaN : Number(tempC),
    densiteObservee: Number(densite) || 0,
  });
  const ecartVolume = result ? result.volume15 - Number(volumeAmbiant) : null;

  return (
    <div className="somip-fade" style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      <div className="somip-panel" style={{ flex: "1 1 320px", padding: 18 }}>
        <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Volume ambiant → Température → Densité → Volume à 15°C</h3>
        <p style={{ margin: "0 0 16px", fontSize: 12.5, color: C.sub }}>
          Formule ASTM D1250 / API MPMS Chapitre 11.1 (Tables 53B/54B, produits pétroliers généralisés) — plage valide {VCF_MIN_DENSITY} à {VCF_MAX_DENSITY} kg/m³.
        </p>
        <Field label="Volume à température ambiante (L)">
          <input type="number" className="somip-input" value={volumeAmbiant} onChange={(e) => setVolumeAmbiant(e.target.value)} placeholder="Ex : 10000" />
        </Field>
        <Field label="Température observée (°C)">
          <input type="number" step="0.1" className="somip-input" value={tempC} onChange={(e) => setTempC(e.target.value)} placeholder="Ex : 29.4" />
        </Field>
        <Field label="Densité observée à cette température (kg/m³)">
          <input type="number" step="0.1" className="somip-input" value={densite} onChange={(e) => setDensite(e.target.value)} placeholder="Ex : 845" />
        </Field>

        {volumeAmbiant !== "" && tempC !== "" && densite !== "" && !result && (
          <p style={{ fontSize: 12.5, color: C.warning, margin: "4px 0 0" }}>
            Vérifiez les valeurs saisies (densité hors plage {VCF_MIN_DENSITY}–{VCF_MAX_DENSITY} kg/m³, ou champ manquant).
          </p>
        )}
      </div>

      <div className="somip-panel" style={{ flex: "1 1 320px", padding: 18 }}>
        <h3 style={{ margin: "0 0 14px", fontSize: 14 }}>Résultat</h3>
        {!result ? (
          <p style={{ fontSize: 13, color: C.sub }}>Renseignez le volume, la température et la densité pour lancer le calcul.</p>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
              <span style={{ color: C.sub }}>Densité ramenée à 15°C</span>
              <span className="somip-mono" style={{ fontWeight: 600 }}>{result.densite15.toFixed(2)} kg/m³</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
              <span style={{ color: C.sub }}>Coefficient ALPHA</span>
              <span className="somip-mono" style={{ fontWeight: 600 }}>{result.alpha.toFixed(7)}</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
              <span style={{ color: C.sub }}>Facteur de correction (VCF)</span>
              <span className="somip-mono" style={{ fontWeight: 600 }}>{result.vcf.toFixed(5)}</span>
            </div>
            <div style={{ height: 1, background: C.border }} />
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 15 }}>
              <span style={{ fontWeight: 600 }}>Volume corrigé à 15°C</span>
              <span className="somip-mono" style={{ fontWeight: 700, color: C.blue, fontSize: 17 }}>{fmt(result.volume15)} L</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5 }}>
              <span style={{ color: C.sub }}>Écart vs volume ambiant</span>
              <span className="somip-mono" style={{ fontWeight: 600, color: ecartVolume >= 0 ? C.success : C.danger }}>{ecartVolume >= 0 ? "+" : ""}{fmt(ecartVolume)} L</span>
            </div>
            <p style={{ margin: "6px 0 0", fontSize: 11, color: C.sub }}>
              Ce calculateur est indépendant du registre des mouvements. Pour rattacher une correction 15°C à une réception ou un inventaire précis, renseigne les champs "Correction à 15°C" disponibles dans ces formulaires.
            </p>
          </div>
        )}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Rapports                                                              */
/* ------------------------------------------------------------------ */
function ReportsView({ sites, movements, inventaires, productStocks, truckAssignments, settings, stockOf, bilans, saveBilan, deleteBilan, canManage }) {
  const [tab, setTab] = useState("synthese_mensuelle_site");
  const TABS = [
    { id: "synthese_mensuelle_site", label: "Synthèse journalière du mois" },
    { id: "synthese_mensuelle_site_15", label: "Synthèse journalière du mois — 15°C" },
    { id: "synthese_mensuelle_lub", label: "Synthèse journalière du mois — Lubrifiants" },
    { id: "synthese_station_jour", label: "Synthèse journalière — Station (site + camion)" },
    { id: "exposition", label: "Exposition", superviseurOnly: true },
    { id: "exposition_comilog", label: "Suivi Stocks Comilog", superviseurOnly: true },
    { id: "bons", label: "Bons de livraison", superviseurOnly: true },
    { id: "bilan", label: "Bilan Matières", superviseurOnly: true },
    { id: "ecart_mensuel", label: "Gain/Perte du mois", superviseurOnly: true },
    { id: "transferts", label: "Transferts entre sites", superviseurOnly: true },
  ].filter((t) => !t.superviseurOnly || canManage);
  useEffect(() => { if (!TABS.some((t) => t.id === tab)) setTab(TABS[0]?.id || "synthese_mensuelle_site"); }, [canManage]); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <div className="somip-fade">
      <div className="somip-no-print" style={{ display: "flex", gap: 8, marginBottom: 18, flexWrap: "wrap" }}>
        {TABS.map((t) => (
          <button key={t.id} className={`somip-tab ${tab === t.id ? "active" : ""}`} onClick={() => setTab(t.id)}>{t.label}</button>
        ))}
      </div>
      {tab === "synthese_mensuelle_site" && <MonthlySiteLedgerReport sites={sites} movements={movements} inventaires={inventaires} />}
      {tab === "synthese_mensuelle_site_15" && <MonthlySiteLedgerReport15 sites={sites} movements={movements} inventaires={inventaires} />}
      {tab === "synthese_mensuelle_lub" && <LubricantMonthlyLedgerReport sites={sites} movements={movements} inventaires={inventaires} productStocks={productStocks} />}
      {tab === "synthese_station_jour" && <StationDailyLedgerReport sites={sites} movements={movements} inventaires={inventaires} truckAssignments={truckAssignments} />}
      {tab === "exposition" && canManage && <ExposureReport sites={sites} movements={movements} inventaires={inventaires} truckAssignments={truckAssignments} productStocks={productStocks} />}
      {tab === "exposition_comilog" && canManage && <ExpositionComilogReport sites={sites} movements={movements} inventaires={inventaires} truckAssignments={truckAssignments} productStocks={productStocks} />}
      {tab === "bons" && canManage && <DeliveryNotesReport sites={sites} movements={movements} />}
      {tab === "bilan" && canManage && <BilanMatieresView sites={sites} bilans={bilans} saveBilan={saveBilan} deleteBilan={deleteBilan} canManage={canManage} />}
      {tab === "ecart_mensuel" && canManage && <EcartMensuelReport sites={sites} movements={movements} inventaires={inventaires} />}
      {tab === "transferts" && canManage && <TransfersReport sites={sites} movements={movements} truckAssignments={truckAssignments} />}
    </div>
  );
}

/* ---- Synthèse journalière du mois, par site (esprit Excel : une ligne par jour) ---- */
function MonthlySiteLedgerReport({ sites, movements, inventaires }) {
  const [siteId, setSiteId] = useState(sites[0]?.id || "");
  const [month, setMonth] = useState(currentMonth());
  const site = sites.find((s) => s.id === siteId);
  const isTruck = !!site?.isMobile;
  const isLubSite = site ? LUBRICANT_SITE_IDS.includes(site.id) : false;
  const bounds = monthBounds(month);

  const days = [];
  if (site) {
    let cur = new Date(bounds.start);
    const end = new Date(bounds.end);
    while (cur <= end) {
      const d = `${cur.getFullYear()}-${pad2(cur.getMonth() + 1)}-${pad2(cur.getDate())}`;
      const stockDebut = stockBeforeDate(site, movements, d, inventaires);
      const dayMovs = movements.filter((m) => m.siteId === site.id && (m.product || "gasoil") === "gasoil" && m.date === d);
      const reception = sumQty(dayMovs, ["reception"]);
      const ventesRaw = sumQty(dayMovs, ["sortie"]); // flux total mesuré au compteur (camion : vente + retour cuve confondus)
      const chargementLaitiers = isTruck ? 0 : sumQty(dayMovs, ["sortie_camion"]);
      const retourCamions = isTruck ? 0 : sumQty(dayMovs, ["retour_camion"]);
      const retourCuve = isTruck ? sumQty(dayMovs, ["retour_cuve_camion"]) : 0;
      const ventes = isTruck ? Math.max(0, ventesRaw - retourCuve) : ventesRaw; // affichage : vente terrain nette
      const stockTheorique = stockDebut + reception + retourCamions - ventesRaw - chargementLaitiers;
      const sortWithIndex = dayMovs.filter((m) => (m.type === "sortie" || m.type === "sortie_camion") && m.indexAvant !== undefined && m.indexApres !== undefined).sort((a, b) => (a.createdAt || "").localeCompare(b.createdAt || ""));
      const indexAvant = sortWithIndex.length ? sortWithIndex[0].indexAvant : null;
      const indexApres = sortWithIndex.length ? sortWithIndex[sortWithIndex.length - 1].indexApres : null;
      const inv = pickLatestInv(inventaires.filter((i) => i.siteId === site.id && (i.product || "gasoil") === "gasoil" && i.date === d));
      const stockJauge = inv ? inv.stockPhysique : null;
      days.push({ date: d, stockDebut, reception, ventes, chargementLaitiers, retourCuve, indexAvant, indexApres, stockTheorique, stockJauge, ecart: stockJauge !== null ? stockJauge - stockTheorique : null });
      cur.setDate(cur.getDate() + 1);
    }
  }

  const totalReception = days.reduce((a, d) => a + d.reception, 0);
  const totalVentes = days.reduce((a, d) => a + d.ventes, 0);
  const totalChargementLaitiers = days.reduce((a, d) => a + d.chargementLaitiers, 0);
  const totalRetourCuve = days.reduce((a, d) => a + d.retourCuve, 0);
  const daysWithJauge = days.filter((d) => d.stockJauge !== null);
  const lastDayWithJauge = daysWithJauge.length ? daysWithJauge[daysWithJauge.length - 1] : null;
  const firstDay = days[0] || null;
  const lastDay = days[days.length - 1] || null;
  const ecartCumule = daysWithJauge.reduce((a, d) => a + (d.ecart || 0), 0);
  const joursEnPerte = days.filter((d) => d.ecart !== null && d.ecart < 0).length;
  const joursEnGain = days.filter((d) => d.ecart !== null && d.ecart > 0).length;

  const receptionLabel = isTruck ? "Chargement" : "Réception";
  const ventesLabel = isTruck ? "Sortie Fiche Terrain" : "Ventes";

  const doExcel = () => exportToExcel(`SOMIP_Synthese_${site?.code || ""}_${month}.xlsx`, [{
    name: "Synthèse", rows: days.map((d) => ({
      Date: d.date, "Stock début (L)": Math.round(d.stockDebut), [`${receptionLabel} (L)`]: Math.round(d.reception), [`${ventesLabel} (L)`]: Math.round(d.ventes),
      ...(isLubSite ? { "Chargement laitiers (L)": Math.round(d.chargementLaitiers) } : {}),
      ...(isTruck ? { "Retour Cuve (L)": Math.round(d.retourCuve) } : {}),
      "Index avant": d.indexAvant ?? "", "Index après": d.indexApres ?? "",
      "Stock théorique (L)": Math.round(d.stockTheorique), "Stock jauge (L)": d.stockJauge !== null ? Math.round(d.stockJauge) : "",
      "Gain/Perte (L)": d.ecart !== null ? Math.round(d.ecart) : "",
    })),
  }]);

  const doPdf = () => exportToPdf({
    filename: `SOMIP_Synthese_${site?.code || ""}_${month}.pdf`,
    title: `Synthèse journalière — ${site?.name || ""}`,
    period: `Mois de ${bounds.start} au ${bounds.end}`,
    columns: [
      "Date", "Stock début", receptionLabel, ventesLabel,
      ...(isLubSite ? ["Chargement laitiers"] : []), ...(isTruck ? ["Retour Cuve"] : []),
      "Index avant", "Index après", "Stock théorique", "Stock jauge", "Gain/Perte",
    ],
    rows: days.map((d) => [
      d.date, `${fmt(d.stockDebut)} L`, d.reception ? `+${fmt(d.reception)} L` : "—", d.ventes ? `${fmt(d.ventes)} L` : "—",
      ...(isLubSite ? [d.chargementLaitiers ? `${fmt(d.chargementLaitiers)} L` : "—"] : []),
      ...(isTruck ? [d.retourCuve ? `−${fmt(d.retourCuve)} L` : "—"] : []),
      d.indexAvant !== null ? fmt(d.indexAvant) : "—", d.indexApres !== null ? fmt(d.indexApres) : "—",
      `${fmt(d.stockTheorique)} L`, d.stockJauge !== null ? `${fmt(d.stockJauge)} L` : "—",
      d.ecart !== null ? `${d.ecart >= 0 ? "+" : ""}${fmt(d.ecart)} L` : "—",
    ]),
  });

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14, display: "flex", gap: 12, flexWrap: "wrap" }}>
        <Field label="Site">
          <select className="somip-select" style={{ maxWidth: 260 }} value={siteId} onChange={(e) => setSiteId(e.target.value)}>
            {sites.map((s) => <option key={s.id} value={s.id}>{s.name}{s.isMobile ? " (camion)" : ""}</option>)}
          </select>
        </Field>
        <Field label="Mois"><input type="month" className="somip-input" style={{ maxWidth: 200 }} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
      </div>

      {site && (
        <div className="somip-panel" style={{ padding: 18, marginBottom: 16 }}>
          <h4 style={{ margin: "0 0 12px", fontSize: 13 }}>Cumul du mois — {site.name}</h4>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <MiniStat label="Stock début (1er jour)" value={firstDay ? `${fmt(firstDay.stockDebut)} L` : "—"} />
            <MiniStat label={`Total ${receptionLabel}`} value={`+${fmt(totalReception)} L`} color={C.success} />
            <MiniStat label={`Total ${ventesLabel}`} value={`${fmt(totalVentes)} L`} />
            {isLubSite && <MiniStat label="Total Chargement laitiers" value={`${fmt(totalChargementLaitiers)} L`} color={C.orange} />}
            {isTruck && <MiniStat label="Total Retour Cuve" value={`${fmt(totalRetourCuve)} L`} />}
            <MiniStat label="Stock théorique (dernier jour)" value={lastDay ? `${fmt(lastDay.stockTheorique)} L` : "—"} bold />
            <MiniStat label="Stock jauge (dernière mesure)" value={lastDayWithJauge ? `${fmt(lastDayWithJauge.stockJauge)} L (${lastDayWithJauge.date})` : "—"} bold />
            <MiniStat label="Gain/Perte cumulé" value={daysWithJauge.length ? `${ecartCumule >= 0 ? "+" : ""}${fmt(ecartCumule)} L` : "—"} color={ecartCumule < 0 ? C.danger : ecartCumule > 0 ? C.success : undefined} />
            <MiniStat label="Jours en perte / en gain" value={`${joursEnPerte} / ${joursEnGain}`} />
          </div>
        </div>
      )}

      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title={`Synthèse journalière — ${site?.name || ""}`} period={`Mois de ${bounds.start} au ${bounds.end}`} />
        <ReportToolbar onExcel={doExcel} onPdf={doPdf} onPrint={() => window.print()} />
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead>
              <tr>
                <th>Date</th><th style={{ textAlign: "right" }}>Stock début</th>
                <th style={{ textAlign: "right" }}>{receptionLabel}</th><th style={{ textAlign: "right" }}>{ventesLabel}</th>
                {isLubSite && <th style={{ textAlign: "right" }}>Chargement laitiers</th>}
                {isTruck && <th style={{ textAlign: "right" }}>Retour Cuve</th>}
                <th style={{ textAlign: "right" }}>Index avant</th><th style={{ textAlign: "right" }}>Index après</th>
                <th style={{ textAlign: "right" }}>Stock théorique</th><th style={{ textAlign: "right" }}>Stock jauge</th>
                <th style={{ textAlign: "right" }}>Gain/Perte</th>
              </tr>
            </thead>
            <tbody>
              {days.length === 0 && <EmptyRow colSpan={isLubSite || isTruck ? 10 : 9} text="Sélectionne un site." />}
              {days.map((d) => (
                <tr key={d.date}>
                  <td className="somip-mono" style={{ fontWeight: 600 }}>{d.date}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(d.stockDebut)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: d.reception ? C.success : C.sub }}>{d.reception ? `+${fmt(d.reception)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: d.ventes ? C.ink : C.sub, fontWeight: d.ventes ? 600 : 400 }}>{d.ventes ? `${fmt(d.ventes)} L` : "—"}</td>
                  {isLubSite && <td className="somip-mono" style={{ textAlign: "right", color: d.chargementLaitiers ? C.orange : C.sub, fontWeight: d.chargementLaitiers ? 600 : 400 }}>{d.chargementLaitiers ? `${fmt(d.chargementLaitiers)} L` : "—"}</td>}
                  {isTruck && <td className="somip-mono" style={{ textAlign: "right", color: d.retourCuve ? C.danger : C.sub }}>{d.retourCuve ? `−${fmt(d.retourCuve)} L` : "—"}</td>}
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{d.indexAvant !== null ? fmt(d.indexAvant) : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{d.indexApres !== null ? fmt(d.indexApres) : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(d.stockTheorique)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{d.stockJauge !== null ? `${fmt(d.stockJauge)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600, color: d.ecart === null ? C.sub : d.ecart < 0 ? C.danger : d.ecart > 0 ? C.success : C.sub }}>
                    {d.ecart !== null ? `${d.ecart >= 0 ? "+" : ""}${fmt(d.ecart)} L` : "—"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          {isTruck
            ? "Stock théorique = Stock début + Chargement − Sortie Fiche Terrain − Retour Cuve (calcul pur). Stock jauge = dernière mesure physique saisie ce jour-là. Gain/Perte = Stock jauge − Stock théorique."
            : isLubSite
            ? "Stock théorique = Stock début + Réception − Ventes − Chargement laitiers (calcul pur). Le Chargement laitiers crée automatiquement le Chargement correspondant côté camion. Stock jauge = dernière mesure physique saisie ce jour-là. Gain/Perte = Stock jauge − Stock théorique."
            : "Stock théorique = Stock début + Réception − Ventes (calcul pur). Stock jauge = dernière mesure physique saisie ce jour-là. Gain/Perte = Stock jauge − Stock théorique."}
        </p>
      </div>
    </div>
  );
}

/* ---- Synthèse journalière du mois, par site — Base 15°C ---- */
function MonthlySiteLedgerReport15({ sites, movements, inventaires }) {
  const [siteId, setSiteId] = useState(sites[0]?.id || "");
  const [month, setMonth] = useState(currentMonth());
  const site = sites.find((s) => s.id === siteId);
  const isTruck = !!site?.isMobile;
  const isLubSite = site ? LUBRICANT_SITE_IDS.includes(site.id) : false;
  const bounds = monthBounds(month);

  const days = [];
  if (site) {
    let cur = new Date(bounds.start);
    const end = new Date(bounds.end);
    while (cur <= end) {
      const d = `${cur.getFullYear()}-${pad2(cur.getMonth() + 1)}-${pad2(cur.getDate())}`;
      const stockDebut = stockBeforeDate15(site, movements, d, inventaires);
      const dayMovs = movements.filter((m) => m.siteId === site.id && (m.product || "gasoil") === "gasoil" && m.date === d);
      const reception = sumQty15(dayMovs, ["reception"]);
      const ventesRaw = sumQty15(dayMovs, ["sortie"]);
      const chargementLaitiers = isTruck ? 0 : sumQty15(dayMovs, ["sortie_camion"]);
      const retourCamions = isTruck ? 0 : sumQty15(dayMovs, ["retour_camion"]);
      const retourCuve = isTruck ? sumQty15(dayMovs, ["retour_cuve_camion"]) : 0;
      const ventes = isTruck ? Math.max(0, ventesRaw - retourCuve) : ventesRaw;
      const stockTheorique = stockDebut + reception + retourCamions - ventesRaw - chargementLaitiers;
      const sortWithIndex = dayMovs.filter((m) => (m.type === "sortie" || m.type === "sortie_camion") && m.indexAvant !== undefined && m.indexApres !== undefined).sort((a, b) => (a.createdAt || "").localeCompare(b.createdAt || ""));
      const indexAvant = sortWithIndex.length ? sortWithIndex[0].indexAvant : null;
      const indexApres = sortWithIndex.length ? sortWithIndex[sortWithIndex.length - 1].indexApres : null;
      const inv = pickLatestInv(inventaires.filter((i) => i.siteId === site.id && (i.product || "gasoil") === "gasoil" && i.date === d && i.stockPhysique15 !== undefined));
      const stockJauge = inv ? inv.stockPhysique15 : null;
      days.push({ date: d, stockDebut, reception, ventes, chargementLaitiers, retourCuve, indexAvant, indexApres, stockTheorique, stockJauge, hasTemp: !!inv, ecart: stockJauge !== null ? stockJauge - stockTheorique : null });
      cur.setDate(cur.getDate() + 1);
    }
  }

  const totalReception = days.reduce((a, d) => a + d.reception, 0);
  const totalVentes = days.reduce((a, d) => a + d.ventes, 0);
  const totalChargementLaitiers = days.reduce((a, d) => a + d.chargementLaitiers, 0);
  const totalRetourCuve = days.reduce((a, d) => a + d.retourCuve, 0);
  const daysWithJauge = days.filter((d) => d.stockJauge !== null);
  const lastDayWithJauge = daysWithJauge.length ? daysWithJauge[daysWithJauge.length - 1] : null;
  const firstDay = days[0] || null;
  const lastDay = days[days.length - 1] || null;
  const ecartCumule = daysWithJauge.reduce((a, d) => a + (d.ecart || 0), 0);

  const receptionLabel = isTruck ? "Chargement" : "Réception";
  const ventesLabel = isTruck ? "Sortie Fiche Terrain" : "Ventes";

  const doExcel = () => exportToExcel(`SOMIP_Synthese15_${site?.code || ""}_${month}.xlsx`, [{
    name: "Synthèse 15°C", rows: days.map((d) => ({
      Date: d.date, "Stock début 15°C (L)": Math.round(d.stockDebut), [`${receptionLabel} 15°C (L)`]: Math.round(d.reception), [`${ventesLabel} 15°C (L)`]: Math.round(d.ventes),
      ...(isLubSite ? { "Chargement laitiers 15°C (L)": Math.round(d.chargementLaitiers) } : {}),
      ...(isTruck ? { "Retour Cuve 15°C (L)": Math.round(d.retourCuve) } : {}),
      "Index avant": d.indexAvant ?? "", "Index après": d.indexApres ?? "",
      "Stock théorique 15°C (L)": Math.round(d.stockTheorique), "Stock jauge 15°C (L)": d.stockJauge !== null ? Math.round(d.stockJauge) : "",
      "Gain/Perte 15°C (L)": d.ecart !== null ? Math.round(d.ecart) : "",
    })),
  }]);

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14, display: "flex", gap: 12, flexWrap: "wrap" }}>
        <Field label="Site">
          <select className="somip-select" style={{ maxWidth: 260 }} value={siteId} onChange={(e) => setSiteId(e.target.value)}>
            {sites.map((s) => <option key={s.id} value={s.id}>{s.name}{s.isMobile ? " (camion)" : ""}</option>)}
          </select>
        </Field>
        <Field label="Mois"><input type="month" className="somip-input" style={{ maxWidth: 200 }} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
      </div>

      {site && (
        <div className="somip-panel" style={{ padding: 18, marginBottom: 16 }}>
          <h4 style={{ margin: "0 0 12px", fontSize: 13 }}>Cumul du mois — {site.name} (15°C)</h4>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <MiniStat label="Stock début (1er jour)" value={firstDay ? `${fmt(firstDay.stockDebut)} L` : "—"} />
            <MiniStat label={`Total ${receptionLabel}`} value={`+${fmt(totalReception)} L`} color={C.success} />
            <MiniStat label={`Total ${ventesLabel}`} value={`${fmt(totalVentes)} L`} />
            {isLubSite && <MiniStat label="Total Chargement laitiers" value={`${fmt(totalChargementLaitiers)} L`} color={C.orange} />}
            {isTruck && <MiniStat label="Total Retour Cuve" value={`${fmt(totalRetourCuve)} L`} />}
            <MiniStat label="Stock théorique (dernier jour)" value={lastDay ? `${fmt(lastDay.stockTheorique)} L` : "—"} bold />
            <MiniStat label="Stock jauge (dernière mesure 15°C)" value={lastDayWithJauge ? `${fmt(lastDayWithJauge.stockJauge)} L (${lastDayWithJauge.date})` : "—"} bold />
            <MiniStat label="Gain/Perte cumulé" value={daysWithJauge.length ? `${ecartCumule >= 0 ? "+" : ""}${fmt(ecartCumule)} L` : "—"} color={ecartCumule < 0 ? C.danger : ecartCumule > 0 ? C.success : undefined} />
          </div>
        </div>
      )}

      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title={`Synthèse journalière — ${site?.name || ""} — Base 15°C`} period={`Mois de ${bounds.start} au ${bounds.end}`} />
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead>
              <tr>
                <th>Date</th><th style={{ textAlign: "right" }}>Stock début</th>
                <th style={{ textAlign: "right" }}>{receptionLabel}</th><th style={{ textAlign: "right" }}>{ventesLabel}</th>
                {isLubSite && <th style={{ textAlign: "right" }}>Chargement laitiers</th>}
                {isTruck && <th style={{ textAlign: "right" }}>Retour Cuve</th>}
                <th style={{ textAlign: "right" }}>Index avant</th><th style={{ textAlign: "right" }}>Index après</th>
                <th style={{ textAlign: "right" }}>Stock théorique</th><th style={{ textAlign: "right" }}>Stock jauge</th>
                <th style={{ textAlign: "right" }}>Gain/Perte</th>
              </tr>
            </thead>
            <tbody>
              {days.length === 0 && <EmptyRow colSpan={isLubSite || isTruck ? 10 : 9} text="Sélectionne un site." />}
              {days.map((d) => (
                <tr key={d.date}>
                  <td className="somip-mono" style={{ fontWeight: 600 }}>{d.date}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(d.stockDebut)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: d.reception ? C.success : C.sub }}>{d.reception ? `+${fmt(d.reception)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: d.ventes ? C.ink : C.sub, fontWeight: d.ventes ? 600 : 400 }}>{d.ventes ? `${fmt(d.ventes)} L` : "—"}</td>
                  {isLubSite && <td className="somip-mono" style={{ textAlign: "right", color: d.chargementLaitiers ? C.orange : C.sub, fontWeight: d.chargementLaitiers ? 600 : 400 }}>{d.chargementLaitiers ? `${fmt(d.chargementLaitiers)} L` : "—"}</td>}
                  {isTruck && <td className="somip-mono" style={{ textAlign: "right", color: d.retourCuve ? C.danger : C.sub }}>{d.retourCuve ? `−${fmt(d.retourCuve)} L` : "—"}</td>}
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{d.indexAvant !== null ? fmt(d.indexAvant) : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{d.indexApres !== null ? fmt(d.indexApres) : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(d.stockTheorique)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{d.stockJauge !== null ? `${fmt(d.stockJauge)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600, color: d.ecart === null ? C.sub : d.ecart < 0 ? C.danger : d.ecart > 0 ? C.success : C.sub }}>
                    {d.ecart !== null ? `${d.ecart >= 0 ? "+" : ""}${fmt(d.ecart)} L` : "—"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          Toutes les valeurs sont corrigées à 15°C. Stock jauge = dernière mesure avec température/densité renseignées ce mois-ci. Les jours sans température/densité saisie n'affichent pas de Stock jauge ici (voir la version Ambiant).
        </p>
      </div>
    </div>
  );
}

/* ---- Rapport journalier ---- */
function DailyReport({ sites, movements, inventaires }) {
  const [date, setDate] = useState(todayStr());
  const rows = sites.map((s) => {
    const stockDebut = stockBeforeDate(s, movements, date, inventaires);
    const dayMovs = movements.filter((m) => m.siteId === s.id && (m.product || "gasoil") === "gasoil" && m.date === date);
    const receptions = sumQty(dayMovs, ["reception"]);
    const retours = sumQty(dayMovs, ["retour_camion"]);
    const sorties = sumQty(dayMovs, ["sortie", "sortie_camion"]);
    const ajustement = dayMovs.filter((m) => m.type === "ajustement").reduce((a, m) => a + m.delta, 0);
    const stockFin = stockDebut + receptions + retours - sorties + ajustement;
    const inv = pickLatestInv(inventaires.filter((i) => i.siteId === s.id && (i.product || "gasoil") === "gasoil" && i.date === date));
    return { site: s, stockDebut, receptions, retours, sorties, ajustement, stockFin, inv };
  });
  const totals = rows.reduce((a, r) => ({ stockDebut: a.stockDebut + r.stockDebut, receptions: a.receptions + r.receptions, retours: a.retours + r.retours, sorties: a.sorties + r.sorties, stockFin: a.stockFin + r.stockFin }), { stockDebut: 0, receptions: 0, retours: 0, sorties: 0, stockFin: 0 });

  const doExcel = () => exportToExcel(`SOMIP_Rapport_Journalier_${date}.xlsx`, [{
    name: "Journalier", rows: rows.map((r) => ({
      Site: r.site.name, "Stock début (L)": Math.round(r.stockDebut), "Réceptions (L)": Math.round(r.receptions),
      "Retours camions (L)": Math.round(r.retours), "Sorties (L)": Math.round(r.sorties), "Ajustement inventaire (L)": Math.round(r.ajustement), "Stock fin (L)": Math.round(r.stockFin),
      "Inventaire du jour": r.inv ? `${NATURE_META[r.inv.nature].label} ${r.inv.ecart >= 0 ? "+" : ""}${Math.round(r.inv.ecart)} L` : "",
    })),
  }]);

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14 }}>
        <Field label="Date du rapport"><input type="date" className="somip-input" style={{ maxWidth: 220 }} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="Rapport journalier de stock" period={`Journée du ${date}`} />
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead><tr><th>Site</th><th style={{ textAlign: "right" }}>Stock début</th><th style={{ textAlign: "right" }}>Réceptions</th><th style={{ textAlign: "right" }}>Retours camions</th><th style={{ textAlign: "right" }}>Sorties</th><th style={{ textAlign: "right" }}>Stock fin</th><th>Inventaire du jour</th></tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(r.stockDebut)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>{r.receptions ? `+${fmt(r.receptions)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>{r.retours ? `+${fmt(r.retours)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.sorties ? C.ink : C.sub, fontWeight: r.sorties ? 600 : 400 }}>{r.sorties ? `${fmt(r.sorties)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stockFin)} L</td>
                  <td>{r.inv ? <Badge color={NATURE_META[r.inv.nature].color}>{NATURE_META[r.inv.nature].label} {r.inv.ecart >= 0 ? "+" : ""}{fmt(r.inv.ecart)} L</Badge> : <span style={{ color: C.sub }}>—</span>}</td>
                </tr>
              ))}
              <tr>
                <td style={{ fontWeight: 700 }}>Total réseau</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(totals.stockDebut)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.success }}>+{fmt(totals.receptions)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.success }}>+{fmt(totals.retours)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(totals.sorties)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(totals.stockFin)} L</td>
                <td></td>
              </tr>
            </tbody>
          </table>
        </div>
        <div style={{ marginTop: 20, height: 240 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={rows.map((r) => ({ code: r.site.code, Réceptions: r.receptions, "Retours": r.retours, Sorties: r.sorties }))} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F3" vertical={false} />
              <XAxis dataKey="code" tick={{ fontSize: 11, fill: C.sub }} axisLine={{ stroke: C.border }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: C.sub }} axisLine={false} tickLine={false} />
              <Tooltip formatter={(v) => `${fmt(v)} L`} contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${C.border}` }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="Réceptions" fill={C.success} radius={[4, 4, 0, 0]} />
              <Bar dataKey="Retours" fill={C.blue} radius={[4, 4, 0, 0]} />
              <Bar dataKey="Sorties" fill={C.danger} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

/* ---- Rapport décadaire ---- */
function DecadeReport({ sites, movements, inventaires }) {
  const [date, setDate] = useState(todayStr());
  const bounds = decadeBounds(date);
  const rows = sites.map((s) => {
    const stockDebut = stockBeforeDate(s, movements, bounds.start, inventaires);
    const rangeMovs = movementsInRange(movements, s.id, bounds.start, bounds.end);
    const receptions = sumQty(rangeMovs, ["reception"]);
    const retours = sumQty(rangeMovs, ["retour_camion"]);
    const sorties = sumQty(rangeMovs, ["sortie", "sortie_camion"]);
    const ajustement = rangeMovs.filter((m) => m.type === "ajustement").reduce((a, m) => a + m.delta, 0);
    const stockFin = stockDebut + receptions + retours - sorties + ajustement;
    const invCount = inventaires.filter((i) => i.siteId === s.id && i.date >= bounds.start && i.date <= bounds.end).length;
    return { site: s, stockDebut, receptions, retours, sorties, stockFin, invCount };
  });

  const doExcel = () => exportToExcel(`SOMIP_Rapport_Decadaire_${bounds.start}_${bounds.end}.xlsx`, [{
    name: "Décadaire", rows: rows.map((r) => ({
      Site: r.site.name, "Stock début (L)": Math.round(r.stockDebut), "Réceptions (L)": Math.round(r.receptions),
      "Retours camions (L)": Math.round(r.retours), "Sorties (L)": Math.round(r.sorties), "Stock fin (L)": Math.round(r.stockFin), "Inventaires réalisés": r.invCount,
    })),
  }]);

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14 }}>
        <Field label="Date de référence (détermine la décade)"><input type="date" className="somip-input" style={{ maxWidth: 220 }} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="Rapport décadaire de stock" period={`${bounds.label} de ${bounds.monthLabel} — du ${bounds.start} au ${bounds.end}`} />
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead><tr><th>Site</th><th style={{ textAlign: "right" }}>Stock début décade</th><th style={{ textAlign: "right" }}>Réceptions</th><th style={{ textAlign: "right" }}>Retours camions</th><th style={{ textAlign: "right" }}>Sorties</th><th style={{ textAlign: "right" }}>Stock fin décade</th><th style={{ textAlign: "right" }}>Inventaires</th></tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(r.stockDebut)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>+{fmt(r.receptions)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>+{fmt(r.retours)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.danger }}>−{fmt(r.sorties)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stockFin)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{r.invCount}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{ marginTop: 20, height: 240 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={rows.map((r) => ({ code: r.site.code, Réceptions: r.receptions, Retours: r.retours, Sorties: r.sorties }))} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F3" vertical={false} />
              <XAxis dataKey="code" tick={{ fontSize: 11, fill: C.sub }} axisLine={{ stroke: C.border }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: C.sub }} axisLine={false} tickLine={false} />
              <Tooltip formatter={(v) => `${fmt(v)} L`} contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${C.border}` }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="Réceptions" fill={C.success} radius={[4, 4, 0, 0]} />
              <Bar dataKey="Retours" fill={C.blue} radius={[4, 4, 0, 0]} />
              <Bar dataKey="Sorties" fill={C.danger} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

/* ---- Rapport mensuel ---- */
function MonthlyReport({ sites, movements, inventaires, settings }) {
  const [month, setMonth] = useState(currentMonth());
  const bounds = monthBounds(month);
  const rows = sites.map((s) => {
    const stockDebut = stockBeforeDate(s, movements, bounds.start, inventaires);
    const rangeMovs = movementsInRange(movements, s.id, bounds.start, bounds.end);
    const receptions = sumQty(rangeMovs, ["reception"]);
    const retours = sumQty(rangeMovs, ["retour_camion"]);
    const sorties = sumQty(rangeMovs, ["sortie", "sortie_camion"]);
    const stockFin = stockThroughDate(s, movements, bounds.end, inventaires);
    const monthInv = inventaires.filter((i) => i.siteId === s.id && i.date >= bounds.start && i.date <= bounds.end);
    const ecartCumule = monthInv.reduce((a, i) => a + i.ecart, 0);
    const pertes = monthInv.filter((i) => i.nature === "perte");
    const tauxMoyen = pertes.length ? pertes.reduce((a, i) => a + i.tauxFreinte, 0) / pertes.length : null;
    return { site: s, stockDebut, receptions, retours, sorties, stockFin, nbInv: monthInv.length, ecartCumule, tauxMoyen };
  });

  const doExcel = () => exportToExcel(`SOMIP_Rapport_Mensuel_${month}.xlsx`, [{
    name: "Mensuel", rows: rows.map((r) => ({
      Site: r.site.name, "Stock début mois (L)": Math.round(r.stockDebut), "Réceptions (L)": Math.round(r.receptions),
      "Retours camions (L)": Math.round(r.retours), "Sorties (L)": Math.round(r.sorties), "Stock fin mois (L)": Math.round(r.stockFin), "Nb inventaires": r.nbInv,
      "Écart cumulé (L)": Math.round(r.ecartCumule), "Taux de freinte moyen (‰)": r.tauxMoyen !== null ? r.tauxMoyen.toFixed(2) : "",
    })),
  }]);

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14 }}>
        <Field label="Mois du rapport"><input type="month" className="somip-input" style={{ maxWidth: 220 }} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="Rapport mensuel de stock" period={`Mois de ${bounds.start} au ${bounds.end}`} />
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead><tr><th>Site</th><th style={{ textAlign: "right" }}>Stock début</th><th style={{ textAlign: "right" }}>Réceptions</th><th style={{ textAlign: "right" }}>Retours camions</th><th style={{ textAlign: "right" }}>Sorties</th><th style={{ textAlign: "right" }}>Stock fin</th><th style={{ textAlign: "right" }}>Nb inv.</th><th style={{ textAlign: "right" }}>Écart cumulé</th><th style={{ textAlign: "right" }}>Freinte moy.</th></tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(r.stockDebut)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>+{fmt(r.receptions)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>+{fmt(r.retours)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.danger }}>−{fmt(r.sorties)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stockFin)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{r.nbInv}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.ecartCumule < 0 ? C.danger : r.ecartCumule > 0 ? C.success : C.sub }}>{r.ecartCumule ? `${r.ecartCumule >= 0 ? "+" : ""}${fmt(r.ecartCumule)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.tauxMoyen !== null && r.tauxMoyen > settings.objectifFreinte ? C.danger : C.sub }}>{r.tauxMoyen !== null ? `${r.tauxMoyen.toFixed(2)} ‰` : "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{ marginTop: 20, height: 240 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={rows.map((r) => ({ code: r.site.code, freinte: r.tauxMoyen || 0 }))} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F3" vertical={false} />
              <XAxis dataKey="code" tick={{ fontSize: 11, fill: C.sub }} axisLine={{ stroke: C.border }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: C.sub }} axisLine={false} tickLine={false} />
              <Tooltip formatter={(v) => `${Number(v).toFixed(2)} ‰`} contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${C.border}` }} />
              <ReferenceLine y={settings.objectifFreinte} stroke={C.warning} strokeDasharray="4 4" label={{ value: `Objectif ${settings.objectifFreinte}‰`, fontSize: 11, fill: C.warning, position: "insideTopRight" }} />
              <Bar dataKey="freinte" radius={[4, 4, 0, 0]}>
                {rows.map((r) => <Cell key={r.site.id} fill={r.tauxMoyen !== null && r.tauxMoyen > settings.objectifFreinte ? C.danger : C.blue} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

/* ---- État journalier des stocks (photographie à une date) ---- */
function StockStatementAmbiant({ sites, movements, inventaires }) {
  const [date, setDate] = useState(todayStr());

  const rows = sites.map((s) => {
    const stockDebut = stockBeforeDate(s, movements, date, inventaires);
    const dayMovs = movements.filter((m) => m.siteId === s.id && (m.product || "gasoil") === "gasoil" && m.date === date);
    const daySorties = dayMovs.filter((m) => m.type === "sortie" || m.type === "sortie_camion" || m.type === "retour_camion").sort((a, b) => (a.createdAt || "").localeCompare(b.createdAt || ""));
    const reception = sumQty(dayMovs, ["reception"]);
    const ventes = sumQty(dayMovs, ["sortie"]);
    const chargementsCamions = sumQty(dayMovs, ["sortie_camion"]);
    const retourCamions = sumQty(dayMovs, ["retour_camion"]);
    const sorties = ventes + chargementsCamions;
    const stockTheorique = stockDebut + reception + retourCamions - sorties;
    const sortWithIndex = daySorties.filter((m) => m.indexAvant !== undefined && m.indexApres !== undefined);
    const sortIndexAvant = sortWithIndex.length ? sortWithIndex[0].indexAvant : null;
    const sortIndexApres = sortWithIndex.length ? sortWithIndex[sortWithIndex.length - 1].indexApres : null;
    const inv = pickLatestInv(inventaires.filter((i) => i.siteId === s.id && (i.product || "gasoil") === "gasoil" && i.date === date));
    const stockFin = inv ? inv.stockPhysique : null;
    const ecart = stockFin !== null ? stockFin - stockTheorique : null;
    return { site: s, stockDebut, reception, ventes, chargementsCamions, retourCamions, sortIndexAvant, sortIndexApres, stockTheorique, stockFin, ecart };
  });

  const doExcel = () => exportToExcel(`SOMIP_Etat_Journalier_Ambiant_${date}.xlsx`, [{
    name: "Etat ambiant", rows: rows.map((r) => ({
      Site: r.site.name, "Stock début (L)": Math.round(r.stockDebut), "Réception (L)": Math.round(r.reception),
      "Ventes (L)": Math.round(r.ventes), "Chargement camions (L)": Math.round(r.chargementsCamions), "Retour camions (L)": Math.round(r.retourCamions),
      "Index avant": r.sortIndexAvant ?? "", "Index après": r.sortIndexApres ?? "",
      "Stock théorique (L)": Math.round(r.stockTheorique),
      "Stock fin mesuré - jauge (L)": r.stockFin !== null ? Math.round(r.stockFin) : "",
      "Gain/Perte (L)": r.ecart !== null ? Math.round(r.ecart) : "",
    })),
  }]);

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14 }}>
        <Field label="Date de l'état journalier"><input type="date" className="somip-input" style={{ maxWidth: 220 }} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="État journalier des stocks — Base ambiante" period={`Journée du ${date}`} />
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead>
              <tr>
                <th>Site</th><th style={{ textAlign: "right" }}>Stock début</th>
                <th style={{ textAlign: "right" }}>Réception</th>
                <th style={{ textAlign: "right" }}>Ventes</th><th style={{ textAlign: "right" }}>Chargement camions</th>
                <th style={{ textAlign: "right" }}>Retour camions</th>
                <th style={{ textAlign: "right" }}>Index avant</th><th style={{ textAlign: "right" }}>Index après</th>
                <th style={{ textAlign: "right" }}>Stock théorique</th><th style={{ textAlign: "right" }}>Stock fin (jauge)</th>
                <th style={{ textAlign: "right" }}>Gain/Perte</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name} <span style={{ color: C.sub, fontWeight: 500 }}>({r.site.code})</span></td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(r.stockDebut)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.reception ? C.success : C.sub }}>{r.reception ? `+${fmt(r.reception)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.ventes ? C.ink : C.sub, fontWeight: r.ventes ? 600 : 400 }}>{r.ventes ? `${fmt(r.ventes)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.chargementsCamions ? C.orange : C.sub }}>{r.chargementsCamions ? `−${fmt(r.chargementsCamions)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.retourCamions ? C.success : C.sub }}>{r.retourCamions ? `+${fmt(r.retourCamions)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{r.sortIndexAvant !== null ? fmt(r.sortIndexAvant) : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{r.sortIndexApres !== null ? fmt(r.sortIndexApres) : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stockTheorique)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{r.stockFin !== null ? `${fmt(r.stockFin)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600, color: r.ecart === null ? C.sub : r.ecart < 0 ? C.danger : r.ecart > 0 ? C.success : C.sub }}>
                    {r.ecart !== null ? `${r.ecart >= 0 ? "+" : ""}${fmt(r.ecart)} L` : "—"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          Base entièrement en volumes ambiants (aucune correction de température). Stock fin = valeur mesurée à la jauge, saisie lors de la Saisie journalière. Gain/Perte = Stock fin − Stock théorique. N'apparaît que pour les sites où un Stock fin a été saisi ce jour-là.
        </p>
      </div>
    </div>
  );
}

function StockStatement15({ sites, movements, inventaires }) {
  const [date, setDate] = useState(todayStr());

  const rows = sites.map((s) => {
    const stockDebut = stockBeforeDate15(s, movements, date, inventaires);
    const dayMovs = movements.filter((m) => m.siteId === s.id && (m.product || "gasoil") === "gasoil" && m.date === date);
    const daySorties = dayMovs.filter((m) => m.type === "sortie" || m.type === "sortie_camion" || m.type === "retour_camion").sort((a, b) => (a.createdAt || "").localeCompare(b.createdAt || ""));
    const reception = sumQty15(dayMovs, ["reception"]);
    const ventes = sumQty15(dayMovs, ["sortie"]);
    const chargementsCamions = sumQty15(dayMovs, ["sortie_camion"]);
    const retourCamions = sumQty15(dayMovs, ["retour_camion"]);
    const sorties = ventes + chargementsCamions;
    const stockTheorique = stockDebut + reception + retourCamions - sorties;
    const sortWithIndex = daySorties.filter((m) => m.indexAvant !== undefined && m.indexApres !== undefined);
    const sortIndexAvant = sortWithIndex.length ? sortWithIndex[0].indexAvant : null;
    const sortIndexApres = sortWithIndex.length ? sortWithIndex[sortWithIndex.length - 1].indexApres : null;
    const inv = pickLatestInv(inventaires.filter((i) => i.siteId === s.id && (i.product || "gasoil") === "gasoil" && i.date === date));
    const stockFin = inv && inv.stockPhysique15 !== undefined ? inv.stockPhysique15 : null;
    const ecart = stockFin !== null ? stockFin - stockTheorique : null;
    return { site: s, stockDebut, reception, ventes, chargementsCamions, retourCamions, sortIndexAvant, sortIndexApres, stockTheorique, stockFin, ecart, hasTemp: inv ? inv.stockPhysique15 !== undefined : false };
  });

  const doExcel = () => exportToExcel(`SOMIP_Etat_Journalier_15C_${date}.xlsx`, [{
    name: "Etat 15C", rows: rows.map((r) => ({
      Site: r.site.name, "Stock début 15°C (L)": Math.round(r.stockDebut), "Réception 15°C (L)": Math.round(r.reception),
      "Ventes 15°C (L)": Math.round(r.ventes), "Chargement camions 15°C (L)": Math.round(r.chargementsCamions), "Retour camions 15°C (L)": Math.round(r.retourCamions),
      "Index avant": r.sortIndexAvant ?? "", "Index après": r.sortIndexApres ?? "",
      "Stock théorique 15°C (L)": Math.round(r.stockTheorique),
      "Stock fin 15°C - jauge (L)": r.stockFin !== null ? Math.round(r.stockFin) : "",
      "Gain/Perte 15°C (L)": r.ecart !== null ? Math.round(r.ecart) : "",
    })),
  }]);

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14 }}>
        <Field label="Date de l'état journalier"><input type="date" className="somip-input" style={{ maxWidth: 220 }} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="État journalier des stocks — Base à 15°C" period={`Journée du ${date}`} />
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead>
              <tr>
                <th>Site</th><th style={{ textAlign: "right" }}>Stock début</th>
                <th style={{ textAlign: "right" }}>Réception</th>
                <th style={{ textAlign: "right" }}>Ventes</th><th style={{ textAlign: "right" }}>Chargement camions</th>
                <th style={{ textAlign: "right" }}>Retour camions</th>
                <th style={{ textAlign: "right" }}>Index avant</th><th style={{ textAlign: "right" }}>Index après</th>
                <th style={{ textAlign: "right" }}>Stock théorique</th><th style={{ textAlign: "right" }}>Stock fin (jauge)</th>
                <th style={{ textAlign: "right" }}>Gain/Perte</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name} <span style={{ color: C.sub, fontWeight: 500 }}>({r.site.code})</span></td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(r.stockDebut)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.reception ? C.success : C.sub }}>{r.reception ? `+${fmt(r.reception)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.ventes ? C.ink : C.sub, fontWeight: r.ventes ? 600 : 400 }}>{r.ventes ? `${fmt(r.ventes)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.chargementsCamions ? C.orange : C.sub }}>{r.chargementsCamions ? `−${fmt(r.chargementsCamions)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.retourCamions ? C.success : C.sub }}>{r.retourCamions ? `+${fmt(r.retourCamions)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{r.sortIndexAvant !== null ? fmt(r.sortIndexAvant) : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{r.sortIndexApres !== null ? fmt(r.sortIndexApres) : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stockTheorique)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{r.stockFin !== null ? `${fmt(r.stockFin)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600, color: r.ecart === null ? C.sub : r.ecart < 0 ? C.danger : r.ecart > 0 ? C.success : C.sub }}>
                    {r.ecart !== null ? `${r.ecart >= 0 ? "+" : ""}${fmt(r.ecart)} L` : "—"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          Toutes les valeurs sont corrigées à 15°C (ASTM D1250 / API MPMS 11.1). Stock fin = valeur mesurée à la jauge, corrigée avec la température/densité saisies au moment de la mesure. N'apparaît que pour les sites où une température et une densité ont été renseignées avec le Stock fin ce jour-là.
        </p>
      </div>
    </div>
  );
}

/* ---- Exposition (synthèse consolidée du réseau) ---- */
function decadeBoundsExplicit(monthStr, decadeNum) {
  const [y, m] = monthStr.split("-").map(Number);
  if (decadeNum === 1) return { start: `${monthStr}-01`, end: `${monthStr}-10`, label: "1ère décade (du 1 au 10)" };
  if (decadeNum === 2) return { start: `${monthStr}-11`, end: `${monthStr}-20`, label: "2e décade (du 11 au 20)" };
  const lastDay = new Date(y, m, 0).getDate();
  return { start: `${monthStr}-21`, end: `${monthStr}-${pad2(lastDay)}`, label: `3e décade (du 21 au ${lastDay})` };
}

function ExposureReport({ sites, movements, inventaires, truckAssignments, productStocks }) {
  const [month, setMonth] = useState(currentMonth());
  const [decadeNum, setDecadeNum] = useState(1);
  const bounds = decadeBoundsExplicit(month, decadeNum);
  const fixedSites = sites.filter((s) => !s.isMobile);
  const [selectedSiteIds, setSelectedSiteIds] = useState(fixedSites.map((s) => s.id));
  const toggleSite = (id) => setSelectedSiteIds((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  const activeSites = fixedSites.filter((s) => selectedSiteIds.includes(s.id));

  const rows = activeSites.map((s) => {
    // Ventes = ventes propres du site (hors chargements laitiers, qui sont un transfert interne,
    // pas une vente) + les ventes des camions qui lui sont rattachés chaque jour de la décade
    // (Sortie Fiche Terrain), en tenant compte des changements d'affectation en cours de période.
    let ventesSite = 0, ventesTrucks = 0;
    let cur = new Date(bounds.start);
    const end = new Date(bounds.end);
    while (cur <= end) {
      const d = `${cur.getFullYear()}-${pad2(cur.getMonth() + 1)}-${pad2(cur.getDate())}`;
      const dayMovsSite = movements.filter((m) => m.siteId === s.id && (m.product || "gasoil") === "gasoil" && m.date === d);
      ventesSite += sumQty(dayMovsSite, ["sortie"]);
      const truckIds = trucksAssignedAt(truckAssignments, s.id, d);
      for (const truckId of truckIds) {
        const dayMovsTruck = movements.filter((m) => m.siteId === truckId && (m.product || "gasoil") === "gasoil" && m.date === d);
        ventesTrucks += sumQty(dayMovsTruck, ["sortie"]) - sumQty(dayMovsTruck, ["retour_cuve_camion"]);
      }
      cur.setDate(cur.getDate() + 1);
    }
    const ventesCumulees = ventesSite + ventesTrucks;
    // Stock en consignation (affichage) : pour Prehomo/Okouma, site + camion(s) rattaché(s) combinés.
    const stockConsignation = combinedStockAtDate(s, sites, movements, inventaires, truckAssignments, bounds.end);
    // La demande d'approvisionnement ne tient compte QUE de la capacité et du dernier stock du
    // site lui-même — jamais des camions (ni leur stock, ni leur capacité).
    const siteInv = pickLatestInv(inventaires.filter((i) => i.siteId === s.id && (i.product || "gasoil") === "gasoil" && i.date <= bounds.end));
    const siteStockOnly = siteInv ? siteInv.stockPhysique : stockThroughDate(s, movements, bounds.end, inventaires);
    const demandeAppro = roundDown5000(s.capacity - siteStockOnly);
    return { site: s, ventesCumulees, stockConsignation, demandeAppro };
  });
  const totalVentes = rows.reduce((a, r) => a + r.ventesCumulees, 0);
  const totalStock = rows.reduce((a, r) => a + r.stockConsignation, 0);
  const totalDemande = rows.reduce((a, r) => a + (r.demandeAppro || 0), 0);

  // Huiles — Prehomo et Okouma uniquement.
  const huilesRows = [];
  for (const siteId of LUBRICANT_SITE_IDS) {
    const s = sites.find((x) => x.id === siteId);
    if (!s || !selectedSiteIds.includes(siteId)) continue;
    for (const lub of LUBRICANTS) {
      const ps = productStocks.find((p) => p.siteId === siteId && p.product === lub.id);
      if (!ps) continue;
      const rangeMovs = movementsInRange(movements, siteId, bounds.start, bounds.end, lub.id);
      const ventes = sumQty(rangeMovs, ["sortie"]);
      const inv = pickLatestInv(inventaires.filter((i) => i.siteId === siteId && (i.product || "gasoil") === lub.id && i.date <= bounds.end));
      const stockConsignation = inv ? inv.stockPhysique : anchoredStock(siteId, ps.stockInitial, movements, inventaires, lub.id, bounds.end, true);
      const demandeAppro = roundDown5000(ps.capacity - stockConsignation);
      huilesRows.push({ site: s, lub, ventes, stockConsignation, demandeAppro });
    }
  }
  const totalHuilesVentes = huilesRows.reduce((a, r) => a + r.ventes, 0);
  const totalHuilesStock = huilesRows.reduce((a, r) => a + r.stockConsignation, 0);
  const totalHuilesDemande = huilesRows.reduce((a, r) => a + r.demandeAppro, 0);

  const titre = `Exposition au ${formatDateLong(todayStr())}`;
  const venteLabel = "Vente";
  const decadeLabel = `Décade ${decadeNum}`;

  const doExcel = () => exportToExcel(`SOMIP_Exposition_${month}_D${decadeNum}.xlsx`, [
    { name: "Exposition", rows: rows.map((r) => ({
      Site: r.site.name, [`${venteLabel} (L)`]: Math.round(r.ventesCumulees),
      "Stock en consignation (L)": Math.round(r.stockConsignation), "Demande d'approvisionnement (L)": Math.round(r.demandeAppro),
    })) },
    ...(huilesRows.length ? [{ name: "Lubrifiants", rows: huilesRows.map((r) => ({
      Site: r.site.name, Produit: r.lub.label, [`${venteLabel} (L)`]: Math.round(r.ventes),
      "Stock en consignation (L)": Math.round(r.stockConsignation), "Demande d'approvisionnement (L)": Math.round(r.demandeAppro),
    })) }] : []),
  ]);

  const doPdf = () => {
    const GASOIL_ORDER = [
      { code: "OKM", label: "OKOUMA" }, { code: "CIM", label: "CIM" }, { code: "GTR", label: "GARE" }, { code: "PRH", label: "PREHOMO" },
      { code: "CMM", label: "CMM" }, { code: "FCV", label: "SETRAG FRANCEVILLE" }, { code: "GSB", label: "GSEZ BENGUIA" }, { code: "LPK", label: "AMD LIPAKA" },
    ];
    const gasoilOrdered = GASOIL_ORDER.map((o) => rows.find((r) => r.site.code === o.code)).filter(Boolean)
      .map((r, i) => ({ ...r, label: GASOIL_ORDER[i].label }));
    const PRODUCT_ORDER = ["rubia_tir7400", "ac50", "ac30", "sw10"];
    const lubFor = (siteId, prodId) => huilesRows.find((r) => r.site.id === siteId && r.lub.id === prodId);
    exportExpositionModelPdf({
      dateStr: todayStr(), decadeNum, monthLabel: `${FRENCH_MONTHS[Number(month.slice(5, 7)) - 1].toUpperCase()} ${month.slice(0, 4)}`,
      gasoilRows: gasoilOrdered, totalVentes, totalStock, totalDemande,
      productOrder: PRODUCT_ORDER, lubFor,
      filename: `SOMIP_Exposition_${month}_D${decadeNum}.pdf`,
    });
  };
  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14, display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
        <Field label="Mois"><input type="month" className="somip-input" style={{ maxWidth: 200 }} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
        <Field label="Décade">
          <select className="somip-select" style={{ maxWidth: 260 }} value={decadeNum} onChange={(e) => setDecadeNum(Number(e.target.value))}>
            <option value={1}>1ère décade (1 au 10)</option>
            <option value={2}>2e décade (11 au 20)</option>
            <option value={3}>3e décade (21 à la fin)</option>
          </select>
        </Field>
      </div>
      <div className="somip-no-print somip-panel" style={{ padding: 14, marginBottom: 14 }}>
        <p style={{ margin: "0 0 8px", fontSize: 12, fontWeight: 700, color: C.ink }}>Sites à inclure</p>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          {fixedSites.map((s) => (
            <label key={s.id} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, cursor: "pointer" }}>
              <input type="checkbox" checked={selectedSiteIds.includes(s.id)} onChange={() => toggleSite(s.id)} />
              {s.name}
            </label>
          ))}
        </div>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title={titre} period={decadeLabel} showEditedDate={false} />
        <ReportToolbar onExcel={doExcel} onPdf={doPdf} onPrint={() => window.print()} />

        <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 20 }}>
          <StatCard label="Ventes cumulées réseau (décade)" value={fmt(totalVentes)} unit="L" accent={C.blue} icon={ArrowUpCircle} />
          <StatCard label="Stock en consignation réseau" value={fmt(totalStock)} unit="L" accent={C.navy} icon={Fuel} />
          <StatCard label="Demande d'approvisionnement réseau" value={fmt(totalDemande)} unit="L" accent={C.orange} icon={Truck} />
        </div>

        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead>
              <tr>
                <th>Site</th><th style={{ textAlign: "right" }}>{venteLabel}</th>
                <th style={{ textAlign: "right" }}>Stock en consignation</th><th style={{ textAlign: "right" }}>Demande d'approvisionnement</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name} <span style={{ color: C.sub, fontWeight: 500 }}>({r.site.code})</span></td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600 }}>{fmt(r.ventesCumulees)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stockConsignation)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.orange }}>{fmt(r.demandeAppro)} L</td>
                </tr>
              ))}
              <tr>
                <td style={{ fontWeight: 700 }}>Total réseau</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(totalVentes)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(totalStock)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.orange }}>{fmt(totalDemande)} L</td>
              </tr>
            </tbody>
          </table>
        </div>

        {huilesRows.length > 0 && (
          <>
            <h4 style={{ margin: "22px 0 10px", fontSize: 13, color: C.ink }}>Lubrifiants — Prehomo &amp; Okouma</h4>
            <div style={{ overflowX: "auto" }}>
              <table className="somip-table">
                <thead>
                  <tr>
                    <th>Site</th><th>Produit</th><th style={{ textAlign: "right" }}>{venteLabel}</th>
                    <th style={{ textAlign: "right" }}>Stock en consignation</th><th style={{ textAlign: "right" }}>Demande d'approvisionnement</th>
                  </tr>
                </thead>
                <tbody>
                  {huilesRows.map((r, i) => (
                    <tr key={i}>
                      <td style={{ fontWeight: 600 }}>{r.site.name}</td>
                      <td>{r.lub.label}</td>
                      <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600 }}>{fmt(r.ventes)} L</td>
                      <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stockConsignation)} L</td>
                      <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.orange }}>{fmt(r.demandeAppro)} L</td>
                    </tr>
                  ))}
                  <tr>
                    <td style={{ fontWeight: 700 }} colSpan={2}>Total lubrifiants</td>
                    <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(totalHuilesVentes)} L</td>
                    <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(totalHuilesStock)} L</td>
                    <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.orange }}>{fmt(totalHuilesDemande)} L</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </>
        )}

        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          {venteLabel} = ventes propres du site (hors chargements laitiers) + ventes du camion rattaché chaque jour de la décade (sur Prehomo/Okouma). Stock en consignation = jauge mesurée à la fin de la décade (site + camion rattaché, pour Prehomo/Okouma). Demande d'approvisionnement = Capacité du site − stock en consignation du site seul (jamais le camion), arrondie à l'inférieur au multiple de 5000 L.
        </p>
      </div>
    </div>
  );
}

const COMILOG_SITE_CODES = ["PRH", "OKM", "CIM", "CMM", "GTR"];

/* ---- Suivi Stocks Comilog — envoi quotidien : ventes & réception de la veille, creux à date ---- */
function ExpositionComilogReport({ sites, movements, inventaires, truckAssignments, productStocks }) {
  const [mvtDate, setMvtDate] = useState(todayStr());
  const [stockDate, setStockDate] = useState(todayStr());
  const allFixedSites = sites.filter((s) => !s.isMobile);
  const [selectedSiteIds, setSelectedSiteIds] = useState(allFixedSites.map((s) => s.id));
  const toggleSite = (id) => setSelectedSiteIds((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  const activeSites = allFixedSites.filter((s) => selectedSiteIds.includes(s.id));

  const rows = activeSites.map((s) => {
    // Ventes = ventes propres du site (hors chargements laitiers, transfert interne, pas une
    // vente) + les ventes du/des camion(s) rattaché(s) ce jour-là (Sortie Fiche Terrain).
    const dayMovs = movements.filter((m) => m.siteId === s.id && (m.product || "gasoil") === "gasoil" && m.date === mvtDate);
    const ventesSite = sumQty(dayMovs, ["sortie"]);
    const reception = sumQty(dayMovs, ["reception"]);
    let ventesTrucks = 0;
    for (const truckId of trucksAssignedAt(truckAssignments, s.id, mvtDate)) {
      const dayMovsTruck = movements.filter((m) => m.siteId === truckId && (m.product || "gasoil") === "gasoil" && m.date === mvtDate);
      ventesTrucks += sumQty(dayMovsTruck, ["sortie"]) - sumQty(dayMovsTruck, ["retour_cuve_camion"]);
    }
    const ventes = ventesSite + ventesTrucks;
    // Stock en consignation (affichage) : pour Prehomo/Okouma, site + camion(s) rattaché(s) combinés.
    const stockConsignation = combinedStockAtDate(s, sites, movements, inventaires, truckAssignments, stockDate);
    // La demande d'approvisionnement ne tient compte QUE de la capacité et du dernier stock du
    // site lui-même — jamais des camions (ni leur stock, ni leur capacité).
    const siteInv = pickLatestInv(inventaires.filter((i) => i.siteId === s.id && (i.product || "gasoil") === "gasoil" && i.date <= stockDate));
    const siteStockOnly = siteInv ? siteInv.stockPhysique : stockThroughDate(s, movements, stockDate, inventaires);
    const demandeAppro = roundDown5000(s.capacity - siteStockOnly);
    return { site: s, ventes, reception, stockConsignation, demandeAppro };
  });
  const totalVentes = rows.reduce((a, r) => a + r.ventes, 0);
  const totalReception = rows.reduce((a, r) => a + r.reception, 0);
  const totalStock = rows.reduce((a, r) => a + r.stockConsignation, 0);
  const totalDemande = rows.reduce((a, r) => a + r.demandeAppro, 0);

  const titre = `Suivi Stocks Comilog au ${formatDateLong(todayStr())}`;

  const doExcel = () => exportToExcel(`SOMIP_Suivi_Stocks_Comilog_${stockDate}.xlsx`, [
    { name: "Suivi Stocks Comilog", rows: rows.map((r) => ({
      Site: r.site.name, [`Ventes du ${mvtDate} (L)`]: Math.round(r.ventes), [`Réception du ${mvtDate} (L)`]: Math.round(r.reception),
      [`Suivi jauges Comilog au ${stockDate} (L)`]: Math.round(r.stockConsignation), "Demande d'approvisionnement (L)": Math.round(r.demandeAppro),
    })) },
  ]);

  const doPdf = () => exportToPdf({
    filename: `SOMIP_Suivi_Stocks_Comilog_${stockDate}.pdf`,
    title: titre,
    period: `Ventes et Réception du ${formatDateShort(mvtDate)} — Stock et Creux du ${formatDateShort(stockDate)}`,
    columns: ["Site", "Ventes", "Réception", `Suivi jauges Comilog au ${formatDateShort(stockDate)}`, "Demande d'approvisionnement"],
    rows: rows.map((r) => [r.site.name, `${fmt(r.ventes)} L`, `${fmt(r.reception)} L`, `${fmt(r.stockConsignation)} L`, `${fmt(r.demandeAppro)} L`]),
    totalsRow: ["Total", `${fmt(totalVentes)} L`, `${fmt(totalReception)} L`, `${fmt(totalStock)} L`, `${fmt(totalDemande)} L`],
  });

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14, display: "flex", gap: 12, flexWrap: "wrap" }}>
        <Field label="Date des ventes"><input type="date" className="somip-input" style={{ maxWidth: 200 }} value={mvtDate} onChange={(e) => setMvtDate(e.target.value)} /></Field>
        <Field label="Date du stock"><input type="date" className="somip-input" style={{ maxWidth: 200 }} value={stockDate} onChange={(e) => setStockDate(e.target.value)} /></Field>
      </div>
      <div className="somip-no-print somip-panel" style={{ padding: 14, marginBottom: 14 }}>
        <p style={{ margin: "0 0 8px", fontSize: 12, fontWeight: 700, color: C.ink }}>Sites à inclure</p>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          {allFixedSites.map((s) => (
            <label key={s.id} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, cursor: "pointer" }}>
              <input type="checkbox" checked={selectedSiteIds.includes(s.id)} onChange={() => toggleSite(s.id)} />
              {s.name}
            </label>
          ))}
        </div>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title={titre} period={`Ventes et Réception du ${formatDateShort(mvtDate)} — Stock et Creux du ${formatDateShort(stockDate)}`} showEditedDate={false} />
        <ReportToolbar onExcel={doExcel} onPdf={doPdf} onPrint={() => window.print()} />

        <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 20 }}>
          <StatCard label={`Ventes cumulées (${formatDateShort(mvtDate)})`} value={fmt(totalVentes)} unit="L" accent={C.blue} icon={ArrowUpCircle} />
          <StatCard label={`Réceptions cumulées (${formatDateShort(mvtDate)})`} value={fmt(totalReception)} unit="L" accent={C.success} icon={ArrowDownCircle} />
          <StatCard label={`Suivi jauges Comilog au ${formatDateShort(stockDate)}`} value={fmt(totalStock)} unit="L" accent={C.navy} icon={Fuel} />
          <StatCard label={`Demande d'approvisionnement (${formatDateShort(stockDate)})`} value={fmt(totalDemande)} unit="L" accent={C.orange} icon={Truck} />
        </div>

        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead>
              <tr>
                <th>Site</th><th style={{ textAlign: "right" }}>Ventes ({formatDateShort(mvtDate)})</th><th style={{ textAlign: "right" }}>Réception ({formatDateShort(mvtDate)})</th>
                <th style={{ textAlign: "right" }}>Suivi jauges Comilog au {formatDateShort(stockDate)}</th><th style={{ textAlign: "right" }}>Demande d'approvisionnement</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name} <span style={{ color: C.sub, fontWeight: 500 }}>({r.site.code})</span></td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600 }}>{fmt(r.ventes)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>{fmt(r.reception)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stockConsignation)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.orange }}>{fmt(r.demandeAppro)} L</td>
                </tr>
              ))}
              <tr>
                <td style={{ fontWeight: 700 }}>Total</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(totalVentes)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.success }}>{fmt(totalReception)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(totalStock)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.orange }}>{fmt(totalDemande)} L</td>
              </tr>
            </tbody>
          </table>
        </div>

        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          Ventes = celles du jour choisi (indépendant de la date du stock) ; sur Prehomo/Okouma, incluent celles du camion rattaché ce jour-là. Stock en consignation = jauge mesurée à la date choisie (site + camion rattaché, pour Prehomo/Okouma). Demande d'approvisionnement = Capacité du site − stock du site seul (jamais le camion), arrondie à l'inférieur au multiple de 5000 L.
        </p>
      </div>
    </div>
  );
}

/* ---- Synthèse Station (site fixe + camions rattachés, ex : Prehomo/Okouma) ---- */
function trucksAssignedAt(assignments, stationId, dateStr) {
  return assignments.filter((a) => a.stationId === stationId && a.startDate <= dateStr && (!a.endDate || a.endDate >= dateStr)).map((a) => a.truckId);
}
// Recherche inverse : le site auquel un camion est rattaché à une date donnée (ou null si aucun).
function stationAssignedTo(assignments, truckId, dateStr) {
  const a = assignments.find((x) => x.truckId === truckId && x.startDate <= dateStr && (!x.endDate || x.endDate >= dateStr));
  return a ? a.stationId : null;
}
// Libellé "Transfert [Site chargé] → [Site normalement affecté]" quand un camion est chargé
// sur un site différent de celui auquel il est habituellement rattaché ce jour-là (secours,
// panne d'un site...). Retourne "" si le camion est chargé sur son site habituel.
function transferLabel(sites, truckAssignments, camionId, chargeSiteId, dateStr) {
  const assignedId = stationAssignedTo(truckAssignments, camionId, dateStr);
  if (!assignedId || assignedId === chargeSiteId) return "";
  const chargeSite = sites.find((s) => s.id === chargeSiteId);
  const assignedSite = sites.find((s) => s.id === assignedId);
  if (!chargeSite || !assignedSite) return "";
  return `Transfert ${chargeSite.name} → ${assignedSite.name}`;
}
// Dernier stock d'un site à une date donnée. Pour Prehomo/Okouma, additionne la jauge du site
// à celle du (des) camion(s) qui lui sont rattachés à cette date (camion = sous-site de la station).
function combinedStockAtDate(s, sites, movements, inventaires, truckAssignments, dateStr) {
  const siteInv = pickLatestInv(inventaires.filter((i) => i.siteId === s.id && (i.product || "gasoil") === "gasoil" && i.date <= dateStr));
  let stock = siteInv ? siteInv.stockPhysique : stockThroughDate(s, movements, dateStr, inventaires);
  if (LUBRICANT_SITE_IDS.includes(s.id)) {
    for (const truckId of trucksAssignedAt(truckAssignments, s.id, dateStr)) {
      const truck = sites.find((x) => x.id === truckId);
      if (!truck) continue;
      const tInv = pickLatestInv(inventaires.filter((i) => i.siteId === truckId && (i.product || "gasoil") === "gasoil" && i.date <= dateStr));
      stock += tInv ? tInv.stockPhysique : stockThroughDate(truck, movements, dateStr, inventaires);
    }
  }
  return stock;
}
function truckIntervalsForStation(assignments, stationId, boundsStart, boundsEnd) {
  return assignments
    .filter((a) => a.stationId === stationId && a.startDate <= boundsEnd && (!a.endDate || a.endDate >= boundsStart))
    .map((a) => ({
      truckId: a.truckId,
      start: a.startDate > boundsStart ? a.startDate : boundsStart,
      end: a.endDate && a.endDate < boundsEnd ? a.endDate : boundsEnd,
    }));
}

/* ---- Synthèse journalière — Station (site fixe + camion(s) rattaché(s), une ligne par jour) ---- */
function StationDailyLedgerReport({ sites, movements, inventaires, truckAssignments }) {
  const stations = LUBRICANT_SITE_IDS.map((id) => sites.find((s) => s.id === id)).filter(Boolean);
  const [stationId, setStationId] = useState(stations[0]?.id || "");
  const [month, setMonth] = useState(currentMonth());
  const station = sites.find((s) => s.id === stationId);
  const bounds = monthBounds(month);

  const days = [];
  const truckEcartByTruck = {}; // { truckId: { name, sum, count } } — cumul propre à chaque camion, sur ses seuls jours jaugés
  if (station) {
    let cur = new Date(bounds.start);
    const end = new Date(bounds.end);
    while (cur <= end) {
      const d = `${cur.getFullYear()}-${pad2(cur.getMonth() + 1)}-${pad2(cur.getDate())}`;
      const siteStockDebut = stockBeforeDate(station, movements, d, inventaires);
      const dayMovsSite = movements.filter((m) => m.siteId === station.id && (m.product || "gasoil") === "gasoil" && m.date === d);
      const reception = sumQty(dayMovsSite, ["reception"]);
      const ventesDirectes = sumQty(dayMovsSite, ["sortie"]);
      const chargementLaitiers = sumQty(dayMovsSite, ["sortie_camion"]);
      // Transferts : chargement d'un camion normalement affecté à un AUTRE site (secours,
      // panne...) — signalé à part pour ne pas passer pour une perte non expliquée.
      const chargementTransfers = dayMovsSite.filter((m) => m.type === "sortie_camion" && m.camion)
        .map((m) => transferLabel(sites, truckAssignments, m.camion, station.id, d)).filter(Boolean);
      const retourCamions = sumQty(dayMovsSite, ["retour_camion"]);
      const siteTheorique = siteStockDebut + reception + retourCamions - ventesDirectes - chargementLaitiers;
      const siteInv = pickLatestInv(inventaires.filter((i) => i.siteId === station.id && (i.product || "gasoil") === "gasoil" && i.date === d));
      const siteEcart = siteInv ? siteInv.stockPhysique - siteTheorique : null;

      const truckIdsToday = trucksAssignedAt(truckAssignments, station.id, d);
      let trucksStockDebut = 0, trucksTheorique = 0, trucksVentesTerrain = 0, trucksJaugeOuTheorique = 0;
      const truckDetails = [];
      for (const truckId of truckIdsToday) {
        const truck = sites.find((s) => s.id === truckId);
        if (!truck) continue;
        const tStockDebut = stockBeforeDate(truck, movements, d, inventaires);
        const dayMovsTruck = movements.filter((m) => m.siteId === truckId && (m.product || "gasoil") === "gasoil" && m.date === d);
        const tChargement = sumQty(dayMovsTruck, ["reception"]);
        const tSortieTerrainRaw = sumQty(dayMovsTruck, ["sortie"]); // flux total mesuré (vente + retour cuve confondus)
        const tRetourCuve = sumQty(dayMovsTruck, ["retour_cuve_camion"]);
        const tSortieTerrain = Math.max(0, tSortieTerrainRaw - tRetourCuve); // affichage : vente terrain nette
        const tTheorique = tStockDebut + tChargement - tSortieTerrainRaw;
        const tInv = pickLatestInv(inventaires.filter((i) => i.siteId === truckId && (i.product || "gasoil") === "gasoil" && i.date === d));
        trucksStockDebut += tStockDebut;
        trucksTheorique += tTheorique;
        trucksVentesTerrain += tSortieTerrain;
        // Un camion n'est pas forcément jaugé chaque jour : on utilise sa jauge du jour si elle
        // existe, sinon son théorique (déjà ancré sur sa dernière jauge connue), pour la ligne du jour.
        trucksJaugeOuTheorique += tInv ? tInv.stockPhysique : tTheorique;
        if (tInv) {
          const tEcart = tInv.stockPhysique - tTheorique;
          if (!truckEcartByTruck[truckId]) truckEcartByTruck[truckId] = { name: truck.name, sum: 0 };
          truckEcartByTruck[truckId].sum += tEcart;
        }
        truckDetails.push({ truck, tSortieTerrain, tChargement, tRetourCuve, tTheorique, tJauge: tInv ? tInv.stockPhysique : null });
      }

      const stockDebutCombine = siteStockDebut + trucksStockDebut;
      const ventesCombinees = ventesDirectes + trucksVentesTerrain;
      const stockTheoriqueCombine = siteTheorique + trucksTheorique;
      // Le site est jaugé chaque jour par construction : c'est la seule condition requise pour
      // afficher un Stock jauge combiné (les camions non jaugés ce jour-là utilisent leur théorique).
      const stockJaugeCombine = siteInv !== null ? siteInv.stockPhysique + trucksJaugeOuTheorique : null;
      const ecart = stockJaugeCombine !== null ? stockJaugeCombine - stockTheoriqueCombine : null;

      days.push({ date: d, stockDebutCombine, reception, ventesCombinees, chargementLaitiers, chargementTransfers, stockTheoriqueCombine, stockJaugeCombine, ecart, siteEcart, truckDetails, nbTrucks: truckIdsToday.length });
      cur.setDate(cur.getDate() + 1);
    }
  }

  const totalReception = days.reduce((a, d) => a + d.reception, 0);
  const totalVentes = days.reduce((a, d) => a + d.ventesCombinees, 0);
  const daysWithJauge = days.filter((d) => d.stockJaugeCombine !== null);
  const lastDayWithJauge = daysWithJauge.length ? daysWithJauge[daysWithJauge.length - 1] : null;
  const firstDay = days[0] || null;
  const lastDay = days[days.length - 1] || null;
  // Le cumul "officiel" est la somme des écarts propres du site + de chaque camion, chacun sur
  // ses seuls jours réellement jaugés — exactement comme leurs rapports individuels respectifs.
  // (Le cumul jour par jour ci-dessus sert à l'affichage détaillé, mais sous-estime le cumul réel
  // les jours où un camion n'a pas de jauge, puisqu'il est alors compté comme sans écart ce jour-là.)
  const siteEcartCumule = days.reduce((a, d) => a + (d.siteEcart || 0), 0);
  const truckEcarts = Object.values(truckEcartByTruck);
  const trucksEcartCumule = truckEcarts.reduce((a, t) => a + t.sum, 0);
  const ecartCumule = siteEcartCumule + trucksEcartCumule;

  const doExcel = () => exportToExcel(`SOMIP_Synthese_Station_${station?.code || ""}_${month}.xlsx`, [{
    name: "Synthèse Station", rows: days.map((d) => ({
      Date: d.date, "Stock début combiné (L)": Math.round(d.stockDebutCombine), "Réception (L)": Math.round(d.reception),
      "Ventes combinées (L)": Math.round(d.ventesCombinees), "Chargement laitiers (L)": Math.round(d.chargementLaitiers),
      "Camions rattachés": d.nbTrucks, "Stock théorique combiné (L)": Math.round(d.stockTheoriqueCombine),
      "Stock jauge combiné (L)": d.stockJaugeCombine !== null ? Math.round(d.stockJaugeCombine) : "",
      "Gain/Perte (L)": d.ecart !== null ? Math.round(d.ecart) : "",
    })),
  }]);

  const doPdf = () => exportToPdf({
    filename: `SOMIP_Synthese_Station_${station?.code || ""}_${month}.pdf`,
    title: `Synthèse journalière — ${station?.name || ""} (site + camions)`,
    period: `Mois de ${bounds.start} au ${bounds.end}`,
    columns: ["Date", "Stock début combiné", "Réception", "Ventes combinées", "Chargement laitiers", "Camions", "Stock théorique combiné", "Stock jauge combiné", "Gain/Perte"],
    rows: days.map((d) => [
      d.date, `${fmt(d.stockDebutCombine)} L`, d.reception ? `+${fmt(d.reception)} L` : "—", d.ventesCombinees ? `${fmt(d.ventesCombinees)} L` : "—",
      d.chargementLaitiers ? `${fmt(d.chargementLaitiers)} L` : "—", String(d.nbTrucks || 0),
      `${fmt(d.stockTheoriqueCombine)} L`, d.stockJaugeCombine !== null ? `${fmt(d.stockJaugeCombine)} L` : "—",
      d.ecart !== null ? `${d.ecart >= 0 ? "+" : ""}${fmt(d.ecart)} L` : "—",
    ]),
  });

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14, display: "flex", gap: 12, flexWrap: "wrap" }}>
        <Field label="Station">
          <select className="somip-select" style={{ maxWidth: 240 }} value={stationId} onChange={(e) => setStationId(e.target.value)}>
            {stations.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </Field>
        <Field label="Mois"><input type="month" className="somip-input" style={{ maxWidth: 200 }} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
      </div>

      {station && (
        <div className="somip-panel" style={{ padding: 18, marginBottom: 16 }}>
          <h4 style={{ margin: "0 0 12px", fontSize: 13 }}>Cumul du mois — {station.name} (site + camions rattachés)</h4>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <MiniStat label="Stock début combiné (1er jour)" value={firstDay ? `${fmt(firstDay.stockDebutCombine)} L` : "—"} />
            <MiniStat label="Total Réceptions (site)" value={`+${fmt(totalReception)} L`} color={C.success} />
            <MiniStat label="Total Ventes combinées" value={`${fmt(totalVentes)} L`} />
            <MiniStat label="Stock théorique combiné (dernier jour)" value={lastDay ? `${fmt(lastDay.stockTheoriqueCombine)} L` : "—"} bold />
            <MiniStat label="Stock jauge combiné (dernière mesure complète)" value={lastDayWithJauge ? `${fmt(lastDayWithJauge.stockJaugeCombine)} L (${lastDayWithJauge.date})` : "—"} bold />
            <MiniStat label="Gain/Perte cumulé (site + camions)" value={`${ecartCumule >= 0 ? "+" : ""}${fmt(ecartCumule)} L`} color={ecartCumule < 0 ? C.danger : ecartCumule > 0 ? C.success : undefined} />
          </div>
          <p style={{ margin: "12px 0 0", fontSize: 12, color: C.sub }}>
            Détail du cumul : <strong style={{ color: C.ink }}>{station.name}</strong> {siteEcartCumule >= 0 ? "+" : ""}{fmt(siteEcartCumule)} L
            {truckEcarts.map((t) => <span key={t.name}> · <strong style={{ color: C.ink }}>{t.name}</strong> {t.sum >= 0 ? "+" : ""}{fmt(t.sum)} L</span>)}
            {truckEcarts.length === 0 && " (aucun camion jaugé ce mois-ci)"}
          </p>
        </div>
      )}

      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title={`Synthèse journalière — ${station?.name || ""} (site + camions)`} period={`Mois de ${bounds.start} au ${bounds.end}`} />
        <ReportToolbar onExcel={doExcel} onPdf={doPdf} onPrint={() => window.print()} />
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead>
              <tr>
                <th>Date</th><th style={{ textAlign: "right" }}>Stock début combiné</th>
                <th style={{ textAlign: "right" }}>Réception</th><th style={{ textAlign: "right" }}>Ventes combinées</th>
                <th style={{ textAlign: "right" }}>Chargement laitiers</th><th style={{ textAlign: "right" }}>Camions</th>
                <th style={{ textAlign: "right" }}>Stock théorique combiné</th><th style={{ textAlign: "right" }}>Stock jauge combiné</th>
                <th style={{ textAlign: "right" }}>Gain/Perte</th>
              </tr>
            </thead>
            <tbody>
              {days.length === 0 && <EmptyRow colSpan={9} text="Sélectionne une station." />}
              {days.map((d) => (
                <tr key={d.date}>
                  <td className="somip-mono" style={{ fontWeight: 600 }}>{d.date}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(d.stockDebutCombine)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: d.reception ? C.success : C.sub }}>{d.reception ? `+${fmt(d.reception)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: d.ventesCombinees ? C.ink : C.sub, fontWeight: d.ventesCombinees ? 600 : 400 }}>{d.ventesCombinees ? `${fmt(d.ventesCombinees)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: d.chargementLaitiers ? C.orange : C.sub }}>
                    {d.chargementLaitiers ? `${fmt(d.chargementLaitiers)} L` : "—"}
                    {d.chargementTransfers?.length > 0 && d.chargementTransfers.map((t, i) => (
                      <div key={i} style={{ fontSize: 10, fontWeight: 600, color: C.warning }}>{t}</div>
                    ))}
                  </td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{d.nbTrucks || "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(d.stockTheoriqueCombine)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{d.stockJaugeCombine !== null ? `${fmt(d.stockJaugeCombine)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600, color: d.ecart === null ? C.sub : d.ecart < 0 ? C.danger : d.ecart > 0 ? C.success : C.sub }}>
                    {d.ecart !== null ? `${d.ecart >= 0 ? "+" : ""}${fmt(d.ecart)} L` : "—"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          Équation basée sur les mouvements de la station ET du (des) camion(s) qui lui sont rattachés ce jour-là (page Sites → Affectation des camions) : Stock théorique combiné = équation du site + équation de chaque camion rattaché (le chargement/transfert interne s'annule automatiquement dans la somme). Stock jauge combiné apparaît dès que le site est jaugé ce jour-là ; un camion non jaugé ce jour précis utilise son théorique (ancré sur sa dernière jauge connue) à la place.
        </p>
      </div>
    </div>
  );
}

function MiniStat({ label, value, color, bold }) {
  return (
    <div style={{ background: C.bg, borderRadius: 8, padding: "10px 14px", minWidth: 150 }}>
      <div style={{ fontSize: 11, color: C.sub, fontWeight: 600, marginBottom: 4 }}>{label}</div>
      <div className="somip-mono" style={{ fontSize: 15, fontWeight: bold ? 700 : 600, color: color || C.ink }}>{value}</div>
    </div>
  );
}

/* ---- Registre des bons de livraison ---- */
/* ---- Bilan Matières global (mensuel ou trimestriel), saisie manuelle + historique + diagramme ---- */
function BilanMatieresView({ sites, bilans, saveBilan, deleteBilan, canManage }) {
  const fixedSites = sites.filter((s) => !s.isMobile);
  const [siteId, setSiteId] = useState(fixedSites[0]?.id || "");
  const [periodType, setPeriodType] = useState("mensuel");
  const [monthKey, setMonthKey] = useState(currentMonth());
  const [year, setYear] = useState(new Date().getFullYear());
  const [quarter, setQuarter] = useState(Math.floor(new Date().getMonth() / 3) + 1);
  const [decadeMonth, setDecadeMonth] = useState(currentMonth());
  const [decadeNum, setDecadeNum] = useState(1);
  const periodKey = periodType === "mensuel" ? monthKey : periodType === "trimestriel" ? `${year}-Q${quarter}` : `${decadeMonth}-D${decadeNum}`;
  const site = sites.find((s) => s.id === siteId);

  const emptyForm = { reception15: "", ventes15: "", transferts15: "", stockFin15: "", commentaire: "" };
  const [form, setForm] = useState(emptyForm);
  const [photoFiles, setPhotoFiles] = useState([]);
  const [existingPhotoUrls, setExistingPhotoUrls] = useState([]);

  const sorted = [...bilans].filter((b) => b.siteId === siteId && b.periodType === periodType).sort((a, b) => (a.periodKey < b.periodKey ? -1 : 1));
  const existing = sorted.find((b) => b.periodKey === periodKey);
  const idx = sorted.findIndex((b) => b.periodKey === periodKey);
  const previous = idx > 0 ? sorted[idx - 1] : (idx === -1 ? sorted[sorted.length - 1] : null);
  const stockDebut = previous ? previous.stockFin15 : 0;

  const receptionN = Number(form.reception15) || 0;
  const ventesN = Number(form.ventes15) || 0;
  const transfertsN = Number(form.transferts15) || 0;
  const stockFinN = Number(form.stockFin15) || 0;
  const stockTheorique = stockDebut + receptionN - ventesN + transfertsN;
  const ecart = form.stockFin15 !== "" ? stockFinN - stockTheorique : null;

  const resetForm = () => { setForm(emptyForm); setPhotoFiles([]); setExistingPhotoUrls([]); };

  const submit = () => {
    if (form.stockFin15 === "" || !siteId) return;
    saveBilan({ siteId, periodType, periodKey, ...form, photoFiles, existingPhotoUrls });
    resetForm();
  };

  const loadForEdit = (b) => {
    setForm({ reception15: String(b.reception15), ventes15: String(b.ventes15), transferts15: String(b.transferts15), stockFin15: String(b.stockFin15), commentaire: b.commentaire || "" });
    setExistingPhotoUrls(b.photoUrls || []);
    setPhotoFiles([]);
  };

  // Historique enrichi pour le tableau + le diagramme (ce site uniquement).
  const history = sorted.map((b, i) => {
    const prev = i > 0 ? sorted[i - 1] : null;
    const debut = prev ? prev.stockFin15 : 0;
    const theorique = debut + b.reception15 - b.ventes15 + b.transferts15;
    const ec = b.stockFin15 - theorique;
    return { ...b, stockDebut: debut, stockTheorique: theorique, ecart: ec };
  });

  const doPptx = () => exportBilanToPptx(history, periodType, site?.name);

  // ---- Vue combinée "Tous les sites", pour une période choisie séparément ----
  const [allPeriodType, setAllPeriodType] = useState("mensuel");
  const [allMonthKey, setAllMonthKey] = useState(currentMonth());
  const [allYear, setAllYear] = useState(new Date().getFullYear());
  const [allQuarter, setAllQuarter] = useState(Math.floor(new Date().getMonth() / 3) + 1);
  const [allDecadeMonth, setAllDecadeMonth] = useState(currentMonth());
  const [allDecadeNum, setAllDecadeNum] = useState(1);
  const allPeriodKey = allPeriodType === "mensuel" ? allMonthKey : allPeriodType === "trimestriel" ? `${allYear}-Q${allQuarter}` : `${allDecadeMonth}-D${allDecadeNum}`;

  const allSitesRows = fixedSites.map((s) => {
    const siteSorted = [...bilans].filter((b) => b.siteId === s.id && b.periodType === allPeriodType).sort((a, b) => (a.periodKey < b.periodKey ? -1 : 1));
    const entry = siteSorted.find((b) => b.periodKey === allPeriodKey);
    if (!entry) return null;
    const i = siteSorted.findIndex((b) => b.periodKey === allPeriodKey);
    const prev = i > 0 ? siteSorted[i - 1] : null;
    const debut = prev ? prev.stockFin15 : 0;
    const theorique = debut + entry.reception15 - entry.ventes15 + entry.transferts15;
    return { site: s, ...entry, stockDebut: debut, stockTheorique: theorique, ecart: entry.stockFin15 - theorique, photoUrls: entry.photoUrls };
  }).filter(Boolean);

  const doAllSitesPdf = () => exportToPdf({
    filename: `SOMIP_Bilan_Matieres_TousSites_${allPeriodType}_${allPeriodKey}.pdf`,
    title: "Bilan Matières — Tous les sites",
    period: `${PERIOD_TYPE_LABELS[allPeriodType]} — ${allPeriodKey}`,
    columns: ["Site", "Stock début", "Réception", "Ventes", "Transferts", "Théorique", "Stock fin", "Gain/Perte"],
    rows: allSitesRows.map((r) => [
      r.site.name, `${fmt(r.stockDebut)} L`, `+${fmt(r.reception15)} L`, `${fmt(r.ventes15)} L`,
      `${r.transferts15 >= 0 ? "+" : ""}${fmt(r.transferts15)} L`, `${fmt(r.stockTheorique)} L`, `${fmt(r.stockFin15)} L`,
      `${r.ecart >= 0 ? "+" : ""}${fmt(r.ecart)} L`,
    ]),
    totalsRow: ["TOTAL",
      `${fmt(allSitesRows.reduce((a, r) => a + r.stockDebut, 0))} L`, `+${fmt(allSitesRows.reduce((a, r) => a + r.reception15, 0))} L`,
      `${fmt(allSitesRows.reduce((a, r) => a + r.ventes15, 0))} L`, `${fmt(allSitesRows.reduce((a, r) => a + r.transferts15, 0))} L`,
      `${fmt(allSitesRows.reduce((a, r) => a + r.stockTheorique, 0))} L`, `${fmt(allSitesRows.reduce((a, r) => a + r.stockFin15, 0))} L`,
      `${fmt(allSitesRows.reduce((a, r) => a + r.ecart, 0))} L`,
    ],
  });
  const doAllSitesPptx = () => exportBilanAllSitesToPptx(allSitesRows, allPeriodType, allPeriodKey);

  return (
    <div className="somip-fade">
      <div style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap", marginBottom: 24 }}>
        {canManage && (
          <div className="somip-panel" style={{ flex: "1 1 320px", padding: 18 }}>
            <h3 style={{ margin: "0 0 14px", fontSize: 14 }}>Saisie du Bilan Matières</h3>
            <Field label="Site">
              <select className="somip-select" value={siteId} onChange={(e) => { setSiteId(e.target.value); resetForm(); }}>
                {fixedSites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </Field>
            <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
              <button className={`somip-tab ${periodType === "mensuel" ? "active" : ""}`} style={{ flex: 1, textAlign: "center", fontSize: 12.5 }} onClick={() => setPeriodType("mensuel")}>Mensuel</button>
              <button className={`somip-tab ${periodType === "decadaire" ? "active" : ""}`} style={{ flex: 1, textAlign: "center", fontSize: 12.5 }} onClick={() => setPeriodType("decadaire")}>Décadaire</button>
              <button className={`somip-tab ${periodType === "trimestriel" ? "active" : ""}`} style={{ flex: 1, textAlign: "center", fontSize: 12.5 }} onClick={() => setPeriodType("trimestriel")}>Trimestriel</button>
            </div>
            {periodType === "mensuel" && (
              <Field label="Mois"><input type="month" className="somip-input" value={monthKey} onChange={(e) => setMonthKey(e.target.value)} /></Field>
            )}
            {periodType === "decadaire" && (
              <div style={{ display: "flex", gap: 8 }}>
                <div style={{ flex: 1 }}><Field label="Mois"><input type="month" className="somip-input" value={decadeMonth} onChange={(e) => setDecadeMonth(e.target.value)} /></Field></div>
                <div style={{ flex: 1 }}>
                  <Field label="Décade">
                    <select className="somip-select" value={decadeNum} onChange={(e) => setDecadeNum(Number(e.target.value))}>
                      <option value={1}>1ère (1 au 10)</option><option value={2}>2e (11 au 20)</option><option value={3}>3e (21 à la fin)</option>
                    </select>
                  </Field>
                </div>
              </div>
            )}
            {periodType === "trimestriel" && (
              <div style={{ display: "flex", gap: 8 }}>
                <div style={{ flex: 1 }}><Field label="Année"><input type="number" className="somip-input" value={year} onChange={(e) => setYear(Number(e.target.value))} /></Field></div>
                <div style={{ flex: 1 }}>
                  <Field label="Trimestre">
                    <select className="somip-select" value={quarter} onChange={(e) => setQuarter(Number(e.target.value))}>
                      <option value={1}>T1 (Jan-Mars)</option><option value={2}>T2 (Avr-Juin)</option>
                      <option value={3}>T3 (Juil-Sept)</option><option value={4}>T4 (Oct-Déc)</option>
                    </select>
                  </Field>
                </div>
              </div>
            )}

            {existing && (
              <p style={{ margin: "-4px 0 10px", fontSize: 11.5, color: C.warning }}>
                Une saisie existe déjà pour cette période — enregistrer à nouveau la remplace. <button onClick={() => loadForEdit(existing)} style={{ border: "none", background: "none", color: C.blue, cursor: "pointer", textDecoration: "underline", padding: 0 }}>Charger pour modifier</button>
              </p>
            )}

            <div style={{ background: C.bg, borderRadius: 8, padding: "9px 12px", marginBottom: 12, display: "flex", justifyContent: "space-between" }}>
              <span style={{ fontSize: 12.5, color: C.sub, fontWeight: 600 }}>Stock début (période précédente — {site?.name})</span>
              <span className="somip-mono" style={{ fontWeight: 700 }}>{fmt(stockDebut)} L</span>
            </div>

            <Field label="Réception à 15°C (L)"><input type="number" className="somip-input" value={form.reception15} onChange={(e) => setForm({ ...form, reception15: e.target.value })} placeholder="0" /></Field>
            <Field label="Ventes à 15°C (L)"><input type="number" className="somip-input" value={form.ventes15} onChange={(e) => setForm({ ...form, ventes15: e.target.value })} placeholder="0" /></Field>
            <Field label="Transferts entre sites à 15°C (L, net)"><input type="number" className="somip-input" value={form.transferts15} onChange={(e) => setForm({ ...form, transferts15: e.target.value })} placeholder="0 (+ reçu, − envoyé)" /></Field>

            <div style={{ background: C.bg, borderRadius: 8, padding: "9px 12px", margin: "4px 0 12px", display: "flex", justifyContent: "space-between" }}>
              <span style={{ fontSize: 12.5, color: C.sub, fontWeight: 600 }}>Stock théorique (calculé)</span>
              <span className="somip-mono" style={{ fontWeight: 700 }}>{fmt(stockTheorique)} L</span>
            </div>

            <Field label="Stock fin mesuré à 15°C (L, obligatoire)"><input type="number" className="somip-input" value={form.stockFin15} onChange={(e) => setForm({ ...form, stockFin15: e.target.value })} placeholder="Jauge du site" /></Field>
            <Field label="Commentaire (optionnel)"><textarea className="somip-textarea" rows={2} value={form.commentaire} onChange={(e) => setForm({ ...form, commentaire: e.target.value })} /></Field>
            <Field label="Photos justificatives (optionnel)">
              <PhotoPicker files={photoFiles} setFiles={setPhotoFiles} existingUrls={existingPhotoUrls} onRemoveExisting={(i) => setExistingPhotoUrls((prev) => prev.filter((_, idx) => idx !== i))} />
            </Field>

            {ecart !== null && (
              <div style={{ background: C.bg, borderRadius: 8, padding: 12, margin: "4px 0 14px", fontSize: 12.5, display: "flex", justifyContent: "space-between" }}>
                <span style={{ color: C.sub, fontWeight: 600 }}>Gain/Perte</span>
                <span className="somip-mono" style={{ fontWeight: 700, color: ecart < 0 ? C.danger : ecart > 0 ? C.success : C.ink }}>{ecart >= 0 ? "+" : ""}{fmt(ecart)} L</span>
              </div>
            )}

            <div style={{ display: "flex", gap: 8 }}>
              <button className="somip-btn somip-btn-primary" style={{ flex: 1, justifyContent: "center" }} onClick={submit} disabled={form.stockFin15 === ""}>
                <Plus size={15} /> Enregistrer
              </button>
              <button className="somip-btn somip-btn-secondary" onClick={resetForm}><X size={15} /> Annuler</button>
            </div>
          </div>
        )}

        <div className="somip-panel" style={{ flex: "2 1 560px", padding: 18 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
            <h3 style={{ margin: 0, fontSize: 14 }}>Historique — {site?.name} — {PERIOD_TYPE_LABELS[periodType]}</h3>
            {history.length > 0 && (
              <button className="somip-btn somip-btn-primary" onClick={doPptx}><Download size={14} /> Export PowerPoint (ce site)</button>
            )}
          </div>
          <div style={{ overflowX: "auto", marginBottom: 20 }}>
            <table className="somip-table">
              <thead><tr><th>Période</th><th style={{ textAlign: "right" }}>Stock début</th><th style={{ textAlign: "right" }}>Réception</th><th style={{ textAlign: "right" }}>Ventes</th><th style={{ textAlign: "right" }}>Transferts</th><th style={{ textAlign: "right" }}>Stock théorique</th><th style={{ textAlign: "right" }}>Stock fin</th><th style={{ textAlign: "right" }}>Gain/Perte</th><th>Photos</th>{canManage && <th></th>}</tr></thead>
              <tbody>
                {history.length === 0 && <EmptyRow colSpan={canManage ? 10 : 9} text="Aucun Bilan Matières enregistré pour ce site." />}
                {history.map((b) => (
                  <tr key={b.id}>
                    <td style={{ fontWeight: 700 }}>{b.periodKey}</td>
                    <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(b.stockDebut)} L</td>
                    <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>+{fmt(b.reception15)} L</td>
                    <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(b.ventes15)} L</td>
                    <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{b.transferts15 >= 0 ? "+" : ""}{fmt(b.transferts15)} L</td>
                    <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(b.stockTheorique)} L</td>
                    <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(b.stockFin15)} L</td>
                    <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600, color: b.ecart < 0 ? C.danger : b.ecart > 0 ? C.success : C.sub }}>{b.ecart >= 0 ? "+" : ""}{fmt(b.ecart)} L</td>
                    <td>
                      {b.photoUrls?.length > 0 ? (
                        <div style={{ display: "flex", gap: 3 }}>
                          {b.photoUrls.slice(0, 3).map((url, i) => (
                            <a key={i} href={url} target="_blank" rel="noopener noreferrer"><img src={url} alt="" style={{ width: 24, height: 24, objectFit: "cover", borderRadius: 4, border: `1px solid ${C.border}` }} /></a>
                          ))}
                          {b.photoUrls.length > 3 && <span style={{ fontSize: 11, color: C.sub }}>+{b.photoUrls.length - 3}</span>}
                        </div>
                      ) : <span style={{ color: C.sub, fontSize: 11 }}>—</span>}
                    </td>
                    {canManage && <td style={{ textAlign: "right" }}><ConfirmIconButton onConfirm={() => deleteBilan(b)} /></td>}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {history.length > 0 && (
            <div style={{ height: 260 }}>
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={history.map((b) => ({ periode: b.periodKey, Réception: b.reception15, Ventes: b.ventes15, "Stock fin": b.stockFin15 }))} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F3" vertical={false} />
                  <XAxis dataKey="periode" tick={{ fontSize: 11, fill: C.sub }} axisLine={{ stroke: C.border }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: C.sub }} axisLine={false} tickLine={false} />
                  <Tooltip formatter={(v) => `${fmt(v)} L`} contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${C.border}` }} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Bar dataKey="Réception" fill={C.success} radius={[4, 4, 0, 0]} />
                  <Bar dataKey="Ventes" fill={C.orange} radius={[4, 4, 0, 0]} />
                  <Bar dataKey="Stock fin" fill={C.blue} radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
        </div>
      </div>

      <div className="somip-panel" style={{ padding: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
          <h3 style={{ margin: 0, fontSize: 14 }}>Tous les sites — export combiné</h3>
          {allSitesRows.length > 0 && (
            <div style={{ display: "flex", gap: 8 }}>
              <button className="somip-btn somip-btn-primary" onClick={doAllSitesPdf}><Download size={14} /> Export PDF</button>
              <button className="somip-btn somip-btn-primary" onClick={doAllSitesPptx}><Download size={14} /> Export PowerPoint</button>
            </div>
          )}
        </div>
        <div className="somip-no-print" style={{ marginBottom: 14, display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
          <div style={{ display: "flex", gap: 8 }}>
            <button className={`somip-tab ${allPeriodType === "mensuel" ? "active" : ""}`} style={{ fontSize: 12.5 }} onClick={() => setAllPeriodType("mensuel")}>Mensuel</button>
            <button className={`somip-tab ${allPeriodType === "decadaire" ? "active" : ""}`} style={{ fontSize: 12.5 }} onClick={() => setAllPeriodType("decadaire")}>Décadaire</button>
            <button className={`somip-tab ${allPeriodType === "trimestriel" ? "active" : ""}`} style={{ fontSize: 12.5 }} onClick={() => setAllPeriodType("trimestriel")}>Trimestriel</button>
          </div>
          {allPeriodType === "mensuel" && <Field label="Mois"><input type="month" className="somip-input" value={allMonthKey} onChange={(e) => setAllMonthKey(e.target.value)} /></Field>}
          {allPeriodType === "decadaire" && (
            <>
              <Field label="Mois"><input type="month" className="somip-input" value={allDecadeMonth} onChange={(e) => setAllDecadeMonth(e.target.value)} /></Field>
              <Field label="Décade">
                <select className="somip-select" value={allDecadeNum} onChange={(e) => setAllDecadeNum(Number(e.target.value))}>
                  <option value={1}>1ère (1 au 10)</option><option value={2}>2e (11 au 20)</option><option value={3}>3e (21 à la fin)</option>
                </select>
              </Field>
            </>
          )}
          {allPeriodType === "trimestriel" && (
            <>
              <Field label="Année"><input type="number" className="somip-input" style={{ maxWidth: 110 }} value={allYear} onChange={(e) => setAllYear(Number(e.target.value))} /></Field>
              <Field label="Trimestre">
                <select className="somip-select" value={allQuarter} onChange={(e) => setAllQuarter(Number(e.target.value))}>
                  <option value={1}>T1</option><option value={2}>T2</option><option value={3}>T3</option><option value={4}>T4</option>
                </select>
              </Field>
            </>
          )}
        </div>
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead><tr><th>Site</th><th style={{ textAlign: "right" }}>Stock début</th><th style={{ textAlign: "right" }}>Réception</th><th style={{ textAlign: "right" }}>Ventes</th><th style={{ textAlign: "right" }}>Transferts</th><th style={{ textAlign: "right" }}>Théorique</th><th style={{ textAlign: "right" }}>Stock fin</th><th style={{ textAlign: "right" }}>Gain/Perte</th></tr></thead>
            <tbody>
              {allSitesRows.length === 0 && <EmptyRow colSpan={8} text="Aucun site n'a de Bilan Matières saisi pour cette période." />}
              {allSitesRows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(r.stockDebut)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>+{fmt(r.reception15)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(r.ventes15)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{r.transferts15 >= 0 ? "+" : ""}{fmt(r.transferts15)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stockTheorique)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stockFin15)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600, color: r.ecart < 0 ? C.danger : r.ecart > 0 ? C.success : C.sub }}>{r.ecart >= 0 ? "+" : ""}{fmt(r.ecart)} L</td>
                </tr>
              ))}
              {allSitesRows.length > 0 && (
                <tr>
                  <td style={{ fontWeight: 700 }}>TOTAL</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(allSitesRows.reduce((a, r) => a + r.stockDebut, 0))} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.success }}>+{fmt(allSitesRows.reduce((a, r) => a + r.reception15, 0))} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(allSitesRows.reduce((a, r) => a + r.ventes15, 0))} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(allSitesRows.reduce((a, r) => a + r.transferts15, 0))} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(allSitesRows.reduce((a, r) => a + r.stockTheorique, 0))} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(allSitesRows.reduce((a, r) => a + r.stockFin15, 0))} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(allSitesRows.reduce((a, r) => a + r.ecart, 0))} L</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          Rassemble la saisie de chaque site pour la période choisie ci-dessus (indépendante de celle du panneau de saisie). Un site sans saisie pour cette période n'apparaît pas dans le tableau. Stock théorique = Stock début (= Stock fin de la période précédente, pour ce site) + Réception − Ventes ± Transferts.
        </p>
      </div>
    </div>
  );
}

/* ---- Gain/Perte du mois — sites et camions (ex-Tableau de bord) ---- */
function EcartMensuelReport({ sites, movements, inventaires }) {
  const [month, setMonth] = useState(currentMonth());
  const monthStartD = `${month}-01`;
  const [endDate, setEndDate] = useState(todayStr());

  const ecartRows = sites.map((s) => {
    let cur = new Date(monthStartD);
    const end = new Date(endDate);
    let ecartCumule = 0, daysWithJauge = 0;
    while (cur <= end) {
      const d = `${cur.getFullYear()}-${pad2(cur.getMonth() + 1)}-${pad2(cur.getDate())}`;
      const stockDebut = stockBeforeDate(s, movements, d, inventaires);
      const dayMovs = movements.filter((m) => m.siteId === s.id && (m.product || "gasoil") === "gasoil" && m.date === d);
      const reception = sumQty(dayMovs, ["reception"]);
      const ventes = sumQty(dayMovs, ["sortie"]);
      const chargementLaitiers = s.isMobile ? 0 : sumQty(dayMovs, ["sortie_camion"]);
      const retourCamions = s.isMobile ? 0 : sumQty(dayMovs, ["retour_camion"]);
      const theorique = stockDebut + reception + retourCamions - ventes - chargementLaitiers;
      const inv = pickLatestInv(inventaires.filter((i) => i.siteId === s.id && (i.product || "gasoil") === "gasoil" && i.date === d));
      if (inv) { ecartCumule += inv.stockPhysique - theorique; daysWithJauge++; }
      cur.setDate(cur.getDate() + 1);
    }
    return { site: s, ecartCumule, daysWithJauge };
  }).sort((a, b) => a.ecartCumule - b.ecartCumule);
  const ecartReseauTotal = ecartRows.filter((r) => !r.site.isMobile).reduce((a, r) => a + r.ecartCumule, 0);

  const doExcel = () => exportToExcel(`SOMIP_Gain_Perte_${month}.xlsx`, [{
    name: "Gain-Perte", rows: ecartRows.map((r) => ({
      Type: r.site.isMobile ? "Camion" : "Site", Nom: r.site.name,
      "Gain/Perte cumulé (L)": r.daysWithJauge > 0 ? Math.round(r.ecartCumule) : "", "Jours jaugés": r.daysWithJauge,
    })),
  }]);
  const doPdf = () => exportToPdf({
    filename: `SOMIP_Gain_Perte_${month}.pdf`,
    title: "Gain/Perte du mois — sites et camions",
    period: `Du ${monthStartD} au ${endDate}`,
    columns: ["Type", "Nom", "Gain/Perte cumulé", "Jours jaugés"],
    rows: ecartRows.map((r) => [r.site.isMobile ? "Camion" : "Site", r.site.name, r.daysWithJauge > 0 ? `${r.ecartCumule >= 0 ? "+" : ""}${fmt(r.ecartCumule)} L` : "—", String(r.daysWithJauge)]),
    totalsRow: ["", "Total réseau (sites)", `${ecartReseauTotal >= 0 ? "+" : ""}${fmt(ecartReseauTotal)} L`, ""],
  });

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14, display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
        <Field label="Mois"><input type="month" className="somip-input" style={{ maxWidth: 200 }} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
        <Field label="Jusqu'au"><input type="date" className="somip-input" style={{ maxWidth: 200 }} value={endDate} onChange={(e) => setEndDate(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="Gain/Perte du mois — sites et camions" period={`Du ${monthStartD} au ${endDate}`} />
        <ReportToolbar onExcel={doExcel} onPdf={doPdf} onPrint={() => window.print()} />
        <StatCard label="Gain/Perte réseau (sites)" value={`${ecartReseauTotal >= 0 ? "+" : ""}${fmt(ecartReseauTotal)}`} unit="L" accent={ecartReseauTotal < 0 ? C.danger : ecartReseauTotal > 0 ? C.success : C.sub} icon={TrendingDown} />
        <div style={{ overflowX: "auto", marginTop: 16 }}>
          <table className="somip-table">
            <thead><tr><th>Type</th><th>Nom</th><th style={{ textAlign: "right" }}>Gain/Perte cumulé</th><th style={{ textAlign: "right" }}>Jours jaugés</th></tr></thead>
            <tbody>
              {ecartRows.length === 0 && <EmptyRow colSpan={4} text="Aucune donnée." />}
              {ecartRows.map((r) => (
                <tr key={r.site.id}>
                  <td><Badge color={r.site.isMobile ? C.orange : C.blue}>{r.site.isMobile ? "Camion" : "Site"}</Badge></td>
                  <td style={{ fontWeight: 600 }}>{r.site.name} {!r.site.isMobile && <span style={{ color: C.sub, fontWeight: 500 }}>({r.site.code})</span>}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: r.ecartCumule < 0 ? C.danger : r.ecartCumule > 0 ? C.success : C.sub }}>
                    {r.daysWithJauge > 0 ? `${r.ecartCumule >= 0 ? "+" : ""}${fmt(r.ecartCumule)} L` : "—"}
                  </td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{r.daysWithJauge}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          Cumul du 1er du mois choisi jusqu'à la date choisie, calcul propre à chaque site et chaque camion (indépendants l'un de l'autre).
        </p>
      </div>
    </div>
  );
}

/* ---- Transferts entre sites (ex-Tableau de bord) ---- */
function TransfersReport({ sites, movements, truckAssignments }) {
  const [month, setMonth] = useState(currentMonth());
  const monthStartD = `${month}-01`;
  const monthEndD = (() => { const d = new Date(month + "-01"); d.setMonth(d.getMonth() + 1); d.setDate(d.getDate() - 1); return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; })();

  const transfers = movements.filter((m) => m.type === "sortie_camion" && m.camion && m.date >= monthStartD && m.date <= monthEndD)
    .map((m) => {
      const label = transferLabel(sites, truckAssignments || [], m.camion, m.siteId, m.date);
      if (!label) return null;
      return { date: m.date, camion: sites.find((s) => s.id === m.camion)?.name || m.camion, label, quantity: m.quantity };
    }).filter(Boolean).sort((a, b) => (a.date < b.date ? 1 : -1));
  const totalTransfers = transfers.reduce((a, t) => a + t.quantity, 0);
  const transfersByRoute = Object.values(
    transfers.reduce((acc, t) => {
      if (!acc[t.label]) acc[t.label] = { label: t.label, quantity: 0, count: 0 };
      acc[t.label].quantity += t.quantity;
      acc[t.label].count += 1;
      return acc;
    }, {})
  ).sort((a, b) => b.quantity - a.quantity);

  const doExcel = () => exportToExcel(`SOMIP_Transferts_${month}.xlsx`, [{
    name: "Transferts", rows: transfers.map((t) => ({ Date: t.date, Camion: t.camion, Détail: t.label, "Quantité (L)": Math.round(t.quantity) })),
  }]);
  const doPdf = () => exportToPdf({
    filename: `SOMIP_Transferts_${month}.pdf`,
    title: "Transferts entre sites",
    period: `${PERIOD_TYPE_LABELS.mensuel} — ${month}`,
    columns: ["Date", "Camion", "Détail", "Quantité"],
    rows: transfers.map((t) => [t.date, t.camion, t.label, `${fmt(t.quantity)} L`]),
    totalsRow: ["", "", "Total", `${fmt(totalTransfers)} L`],
  });

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14 }}>
        <Field label="Mois"><input type="month" className="somip-input" style={{ maxWidth: 200 }} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="Transferts entre sites" period={`${PERIOD_TYPE_LABELS.mensuel} — ${month}`} />
        <ReportToolbar onExcel={doExcel} onPdf={doPdf} onPrint={() => window.print()} />
        <StatCard label="Total transféré" value={fmt(totalTransfers)} unit="L" accent={C.orange} icon={Truck} />
        <p style={{ margin: "14px 0 0", fontSize: 12, color: C.sub }}>Camion chargé sur un site différent de celui où il est normalement affecté (secours, panne...). La réception a été ajoutée automatiquement côté camion.</p>

        {transfersByRoute.length > 1 && (
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", margin: "14px 0" }}>
            {transfersByRoute.map((r, i) => (
              <div key={i} style={{ background: C.bg, borderRadius: 8, padding: "8px 12px", fontSize: 12.5 }}>
                <span style={{ fontWeight: 600, color: C.ink }}>{r.label}</span>
                <span style={{ color: C.sub }}> — {r.count} transfert{r.count > 1 ? "s" : ""} — </span>
                <span className="somip-mono" style={{ fontWeight: 700, color: C.orange }}>{fmt(r.quantity)} L</span>
              </div>
            ))}
          </div>
        )}

        <div style={{ overflowX: "auto", marginTop: 16 }}>
          <table className="somip-table">
            <thead><tr><th>Date</th><th>Camion</th><th>Détail</th><th style={{ textAlign: "right" }}>Quantité</th></tr></thead>
            <tbody>
              {transfers.length === 0 && <EmptyRow colSpan={4} text="Aucun transfert ce mois-ci." />}
              {transfers.map((t, i) => (
                <tr key={i}>
                  <td className="somip-mono">{t.date}</td>
                  <td style={{ fontWeight: 600 }}>{t.camion}</td>
                  <td style={{ color: C.orange, fontWeight: 600 }}>{t.label}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(t.quantity)} L</td>
                </tr>
              ))}
              {transfers.length > 0 && (
                <tr>
                  <td colSpan={3} style={{ fontWeight: 700 }}>Total</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.orange }}>{fmt(totalTransfers)} L</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function DeliveryNotesReport({ sites, movements }) {
  const [filterSite, setFilterSite] = useState("all");
  const [start, setStart] = useState("");
  const [end, setEnd] = useState("");
  const [search, setSearch] = useState("");

  const rows = movements
    .filter((m) => m.type === "reception")
    .filter((m) => filterSite === "all" || m.siteId === filterSite)
    .filter((m) => (!start || m.date >= start) && (!end || m.date <= end))
    .filter((m) => !search.trim() || (m.ref || "").toLowerCase().includes(search.trim().toLowerCase()))
    .sort((a, b) => (a.date < b.date ? 1 : (a.date > b.date ? -1 : (a.createdAt || "").localeCompare(b.createdAt || ""))));

  const productLabel = (m) => PRODUCTS.find((p) => p.id === (m.product || "gasoil"))?.label || m.product;

  const doExcel = () => exportToExcel(`SOMIP_Bons_Livraison.xlsx`, [{
    name: "Bons de livraison", rows: rows.map((m) => ({
      Date: m.date, Site: sites.find((s) => s.id === m.siteId)?.name || m.siteId, Produit: productLabel(m),
      "N° Bon": m.ref || "", "Quantité (L)": Math.round(m.quantity), "Quantité 15°C (L)": m.volumeCorrige15 ? Math.round(m.volumeCorrige15) : "",
    })),
  }]);

  const doPdf = () => exportToPdf({
    filename: "SOMIP_Bons_Livraison.pdf",
    title: "Registre des bons de livraison",
    period: start || end ? `Du ${start || "…"} au ${end || "…"}` : "Toutes les réceptions",
    columns: ["Date", "Site", "Produit", "N° Bon", "Quantité", "Quantité 15°C"],
    rows: rows.map((m) => [
      m.date, sites.find((s) => s.id === m.siteId)?.name || m.siteId, productLabel(m),
      m.ref || "—", `+${fmt(m.quantity)} L`, m.volumeCorrige15 ? `${fmt(m.volumeCorrige15)} L` : "—",
    ]),
  });

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14, display: "flex", gap: 12, flexWrap: "wrap" }}>
        <Field label="Rechercher un N° de bon"><input className="somip-input" style={{ maxWidth: 220 }} value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Ex : BL-2451" /></Field>
        <Field label="Site">
          <select className="somip-select" style={{ maxWidth: 240 }} value={filterSite} onChange={(e) => setFilterSite(e.target.value)}>
            <option value="all">Tous les sites</option>
            {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </Field>
        <Field label="Du (optionnel)"><input type="date" className="somip-input" style={{ maxWidth: 180 }} value={start} onChange={(e) => setStart(e.target.value)} /></Field>
        <Field label="Au (optionnel)"><input type="date" className="somip-input" style={{ maxWidth: 180 }} value={end} onChange={(e) => setEnd(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="Registre des bons de livraison" period={start || end ? `Du ${start || "…"} au ${end || "…"}` : "Toutes les réceptions"} />
        <ReportToolbar onExcel={doExcel} onPdf={doPdf} onPrint={() => window.print()} />
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead><tr><th>Date</th><th>Site</th><th>Produit</th><th>N° Bon</th><th style={{ textAlign: "right" }}>Quantité</th><th style={{ textAlign: "right" }}>Quantité 15°C</th></tr></thead>
            <tbody>
              {rows.length === 0 && <EmptyRow colSpan={6} text="Aucune réception trouvée." />}
              {rows.map((m) => (
                <tr key={m.id}>
                  <td className="somip-mono">{m.date}</td>
                  <td style={{ fontWeight: 600 }}>{sites.find((s) => s.id === m.siteId)?.name || m.siteId}</td>
                  <td>{productLabel(m)}</td>
                  <td className="somip-mono" style={{ fontWeight: 700, color: C.blue }}>{m.ref || "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success, fontWeight: 600 }}>+{fmt(m.quantity)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{m.volumeCorrige15 ? `${fmt(m.volumeCorrige15)} L` : "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          {rows.length} bon(s) listé(s). Les "Chargements" créés automatiquement côté camion (miroir d'une sortie vers camion) apparaissent aussi ici — repérables par leur mention "Chargement automatique depuis...".
        </p>
      </div>
    </div>
  );
}

/* ---- Synthèse journalière du mois — Lubrifiants (une ligne par jour, esprit Excel) ---- */
function LubricantMonthlyLedgerReport({ sites, movements, inventaires, productStocks }) {
  const [siteId, setSiteId] = useState(LUBRICANT_SITE_IDS[0]);
  const [productId, setProductId] = useState(LUBRICANTS[0].id);
  const [month, setMonth] = useState(currentMonth());
  const site = sites.find((s) => s.id === siteId);
  const lub = LUBRICANTS.find((l) => l.id === productId);
  const productStockEntry = productStocks.find((p) => p.siteId === siteId && p.product === productId);
  const bounds = monthBounds(month);

  const days = [];
  if (site && lub) {
    let cur = new Date(bounds.start);
    const end = new Date(bounds.end);
    while (cur <= end) {
      const d = `${cur.getFullYear()}-${pad2(cur.getMonth() + 1)}-${pad2(cur.getDate())}`;
      const stockDebut = stockBeforeDateProduct(productStockEntry?.stockInitial || 0, movements, siteId, productId, d, inventaires);
      const dayMovs = movements.filter((m) => m.siteId === siteId && (m.product || "gasoil") === productId && m.date === d);
      const reception = sumQty(dayMovs, ["reception"]);
      const ventes = sumQty(dayMovs, ["sortie"]);
      const stockTheorique = stockDebut + reception - ventes;
      const sortWithIndex = dayMovs.filter((m) => m.type === "sortie" && m.indexAvant !== undefined && m.indexApres !== undefined).sort((a, b) => (a.createdAt || "").localeCompare(b.createdAt || ""));
      const indexAvant = sortWithIndex.length ? sortWithIndex[0].indexAvant : null;
      const indexApres = sortWithIndex.length ? sortWithIndex[sortWithIndex.length - 1].indexApres : null;
      const inv = pickLatestInv(inventaires.filter((i) => i.siteId === siteId && (i.product || "gasoil") === productId && i.date === d));
      const stockJauge = inv ? inv.stockPhysique : null;
      days.push({ date: d, stockDebut, reception, ventes, indexAvant, indexApres, stockTheorique, stockJauge, ecart: stockJauge !== null ? stockJauge - stockTheorique : null });
      cur.setDate(cur.getDate() + 1);
    }
  }

  const totalReception = days.reduce((a, d) => a + d.reception, 0);
  const totalVentes = days.reduce((a, d) => a + d.ventes, 0);
  const daysWithJauge = days.filter((d) => d.stockJauge !== null);
  const lastDayWithJauge = daysWithJauge.length ? daysWithJauge[daysWithJauge.length - 1] : null;
  const firstDay = days[0] || null;
  const lastDay = days[days.length - 1] || null;
  const ecartCumule = daysWithJauge.reduce((a, d) => a + (d.ecart || 0), 0);

  const doExcel = () => exportToExcel(`SOMIP_Synthese_${lub?.label || ""}_${site?.code || ""}_${month}.xlsx`, [{
    name: "Synthèse", rows: days.map((d) => ({
      Date: d.date, "Stock début (L)": Math.round(d.stockDebut), "Réception (L)": Math.round(d.reception), "Ventes (L)": Math.round(d.ventes),
      "Index avant": d.indexAvant ?? "", "Index après": d.indexApres ?? "",
      "Stock théorique (L)": Math.round(d.stockTheorique), "Stock jauge (L)": d.stockJauge !== null ? Math.round(d.stockJauge) : "",
      "Gain/Perte (L)": d.ecart !== null ? Math.round(d.ecart) : "",
    })),
  }]);

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14, display: "flex", gap: 12, flexWrap: "wrap" }}>
        <Field label="Site">
          <select className="somip-select" style={{ maxWidth: 200 }} value={siteId} onChange={(e) => setSiteId(e.target.value)}>
            {LUBRICANT_SITE_IDS.map((id) => <option key={id} value={id}>{sites.find((s) => s.id === id)?.name || id}</option>)}
          </select>
        </Field>
        <Field label="Produit">
          <select className="somip-select" style={{ maxWidth: 200 }} value={productId} onChange={(e) => setProductId(e.target.value)}>
            {LUBRICANTS.map((l) => <option key={l.id} value={l.id}>{l.label}</option>)}
          </select>
        </Field>
        <Field label="Mois"><input type="month" className="somip-input" style={{ maxWidth: 200 }} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
      </div>

      {site && lub && (
        <div className="somip-panel" style={{ padding: 18, marginBottom: 16 }}>
          <h4 style={{ margin: "0 0 12px", fontSize: 13 }}>Cumul du mois — {site.name} — {lub.label}</h4>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <MiniStat label="Stock début (1er jour)" value={firstDay ? `${fmt(firstDay.stockDebut)} L` : "—"} />
            <MiniStat label="Total Réceptions" value={`+${fmt(totalReception)} L`} color={C.success} />
            <MiniStat label="Total Ventes" value={`${fmt(totalVentes)} L`} />
            <MiniStat label="Stock théorique (dernier jour)" value={lastDay ? `${fmt(lastDay.stockTheorique)} L (≈ ${fmt(lastDay.stockTheorique * lub.densite)} kg)` : "—"} bold />
            <MiniStat label="Stock jauge (dernière mesure)" value={lastDayWithJauge ? `${fmt(lastDayWithJauge.stockJauge)} L (${lastDayWithJauge.date})` : "—"} bold />
            <MiniStat label="Gain/Perte cumulé" value={daysWithJauge.length ? `${ecartCumule >= 0 ? "+" : ""}${fmt(ecartCumule)} L` : "—"} color={ecartCumule < 0 ? C.danger : ecartCumule > 0 ? C.success : undefined} />
          </div>
          {!productStockEntry && <p style={{ margin: "10px 0 0", fontSize: 11.5, color: C.warning }}>Capacité/stock initial non configuré pour ce produit — définis-le sur la page Sites.</p>}
        </div>
      )}

      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title={`Synthèse journalière — ${site?.name || ""} — ${lub?.label || ""}`} period={`Mois de ${bounds.start} au ${bounds.end}`} />
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead>
              <tr>
                <th>Date</th><th style={{ textAlign: "right" }}>Stock début</th>
                <th style={{ textAlign: "right" }}>Réception</th><th style={{ textAlign: "right" }}>Ventes</th>
                <th style={{ textAlign: "right" }}>Index avant</th><th style={{ textAlign: "right" }}>Index après</th>
                <th style={{ textAlign: "right" }}>Stock théorique</th><th style={{ textAlign: "right" }}>Stock jauge</th>
                <th style={{ textAlign: "right" }}>Gain/Perte</th>
              </tr>
            </thead>
            <tbody>
              {days.length === 0 && <EmptyRow colSpan={9} text="Sélectionne un site et un produit." />}
              {days.map((d) => (
                <tr key={d.date}>
                  <td className="somip-mono" style={{ fontWeight: 600 }}>{d.date}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(d.stockDebut)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: d.reception ? C.success : C.sub }}>{d.reception ? `+${fmt(d.reception)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: d.ventes ? C.ink : C.sub, fontWeight: d.ventes ? 600 : 400 }}>{d.ventes ? `${fmt(d.ventes)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{d.indexAvant !== null ? fmt(d.indexAvant) : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{d.indexApres !== null ? fmt(d.indexApres) : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(d.stockTheorique)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{d.stockJauge !== null ? `${fmt(d.stockJauge)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600, color: d.ecart === null ? C.sub : d.ecart < 0 ? C.danger : d.ecart > 0 ? C.success : C.sub }}>
                    {d.ecart !== null ? `${d.ecart >= 0 ? "+" : ""}${fmt(d.ecart)} L` : "—"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          Stock théorique = Stock début + Réception − Ventes (calcul pur). Stock jauge = dernière mesure physique saisie ce jour-là. Gain/Perte = Stock jauge − Stock théorique.
        </p>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Utilisateurs                                                          */
/* ------------------------------------------------------------------ */
/* ------------------------------------------------------------------ */
/* Personnalisation — logo & couleurs (Superviseur uniquement)         */
/* ------------------------------------------------------------------ */
function BrandingView({ settings, updateTheme }) {
  const [logoFile, setLogoFile] = useState(null);
  const [logoTotalFile, setLogoTotalFile] = useState(null);
  const [colorPrimary, setColorPrimary] = useState(settings.colorPrimary || "#0071BD");
  const [colorAccent, setColorAccent] = useState(settings.colorAccent || "#F16B16");
  const [saving, setSaving] = useState(false);

  const previewLogo = logoFile ? URL.createObjectURL(logoFile) : settings.logoUrl;
  const previewLogoTotal = logoTotalFile ? URL.createObjectURL(logoTotalFile) : settings.logoTotalUrl;
  const dirty = !!logoFile || !!logoTotalFile || colorPrimary !== (settings.colorPrimary || "#0071BD") || colorAccent !== (settings.colorAccent || "#F16B16");

  const submit = async () => {
    setSaving(true);
    await updateTheme({ logoFile, logoTotalFile, colorPrimary, colorAccent });
    setSaving(false);
    setLogoFile(null); setLogoTotalFile(null);
    window.location.reload();
  };

  return (
    <div className="somip-fade" style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      <div className="somip-panel" style={{ flex: "1 1 380px", padding: 18 }}>
        <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Personnalisation</h3>
        <p style={{ margin: "0 0 18px", fontSize: 12.5, color: C.sub }}>Logo et couleurs principales, appliqués à toute l'application et aux rapports (PDF/PowerPoint).</p>

        <Field label="Logo SOMIP">
          <label className="somip-btn somip-btn-secondary" style={{ fontSize: 12, padding: "6px 12px", cursor: "pointer", display: "inline-flex" }}>
            <ImagePlus size={14} /> {settings.logoUrl || logoFile ? "Changer le logo" : "Ajouter un logo"}
            <input type="file" accept="image/*" style={{ display: "none" }} onChange={(e) => setLogoFile(e.target.files?.[0] || null)} />
          </label>
        </Field>
        {previewLogo && (
          <div style={{ margin: "8px 0 16px", padding: 14, background: C.bg, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <img src={previewLogo} alt="Logo" style={{ maxHeight: 70, maxWidth: "100%" }} />
          </div>
        )}

        <Field label="Logo TotalEnergies (pages Inventaires)">
          <label className="somip-btn somip-btn-secondary" style={{ fontSize: 12, padding: "6px 12px", cursor: "pointer", display: "inline-flex" }}>
            <ImagePlus size={14} /> {settings.logoTotalUrl || logoTotalFile ? "Changer le logo" : "Ajouter un logo"}
            <input type="file" accept="image/*" style={{ display: "none" }} onChange={(e) => setLogoTotalFile(e.target.files?.[0] || null)} />
          </label>
        </Field>
        {previewLogoTotal && (
          <div style={{ margin: "8px 0 16px", padding: 14, background: C.bg, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <img src={previewLogoTotal} alt="Logo TotalEnergies" style={{ maxHeight: 70, maxWidth: "100%" }} />
          </div>
        )}

        <Field label="Couleur primaire">
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <input type="color" value={colorPrimary} onChange={(e) => setColorPrimary(e.target.value)} style={{ width: 44, height: 36, border: `1px solid ${C.border}`, borderRadius: 6, padding: 2, cursor: "pointer" }} />
            <input className="somip-input" value={colorPrimary} onChange={(e) => setColorPrimary(e.target.value)} style={{ flex: 1 }} />
          </div>
        </Field>
        <Field label="Couleur d'accent">
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <input type="color" value={colorAccent} onChange={(e) => setColorAccent(e.target.value)} style={{ width: 44, height: 36, border: `1px solid ${C.border}`, borderRadius: 6, padding: 2, cursor: "pointer" }} />
            <input className="somip-input" value={colorAccent} onChange={(e) => setColorAccent(e.target.value)} style={{ flex: 1 }} />
          </div>
        </Field>

        <button className="somip-btn somip-btn-primary" style={{ width: "100%", justifyContent: "center", marginTop: 8 }} onClick={submit} disabled={!dirty || saving}>
          <Check size={15} /> {saving ? "Enregistrement..." : "Enregistrer"}
        </button>
      </div>

      <div className="somip-panel" style={{ flex: "1 1 280px", padding: 18 }}>
        <h4 style={{ margin: "0 0 12px", fontSize: 13 }}>Aperçu</h4>
        <div style={{ height: 5, background: `linear-gradient(90deg, ${colorPrimary} 0%, ${colorPrimary} 60%, ${colorAccent} 60%, ${colorAccent} 100%)`, borderRadius: 3, marginBottom: 14 }} />
        <button className="somip-btn" style={{ background: colorPrimary, color: "#fff", border: "none", marginBottom: 10 }}>Bouton principal</button>
        <br />
        <span style={{ display: "inline-block", padding: "4px 10px", borderRadius: 6, background: colorAccent, color: "#fff", fontSize: 12.5, fontWeight: 600 }}>Badge d'accent</span>
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>La page se recharge automatiquement après l'enregistrement pour appliquer les couleurs partout.</p>
      </div>
    </div>
  );
}


function UsersView({ profiles, updateUserRole, updateUserSites, toggleUserActive, sites, session }) {
  const [editingId, setEditingId] = useState(null);
  const [roleDraft, setRoleDraft] = useState("");
  const [siteDraft, setSiteDraft] = useState([]);
  const [deletingErr, setDeletingErr] = useState(null);
  const [resetForId, setResetForId] = useState(null);
  const [resetPassword, setResetPassword] = useState("");
  const [resetErr, setResetErr] = useState(null);
  const [resetMsg, setResetMsg] = useState(null);
  const [form, setForm] = useState({ fullName: "", username: "", email: "", password: "", role: "lecture", assignedSiteIds: [] });
  const [creating, setCreating] = useState(false);
  const [createErr, setCreateErr] = useState(null);
  const [createMsg, setCreateMsg] = useState(null);

  const deleteAccount = async (userId) => {
    setDeletingErr(null);
    try {
      const res = await fetch("/api/delete-user", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${session?.access_token || ""}` },
        body: JSON.stringify({ userId }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Erreur lors de la suppression du compte.");
    } catch (e) {
      setDeletingErr(e.message || "Erreur lors de la suppression du compte.");
    }
  };

  const submitReset = async () => {
    setResetErr(null); setResetMsg(null);
    if (!resetPassword || resetPassword.length < 6) { setResetErr("Le mot de passe doit contenir au moins 6 caractères."); return; }
    try {
      const res = await fetch("/api/reset-password", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${session?.access_token || ""}` },
        body: JSON.stringify({ userId: resetForId, newPassword: resetPassword }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Erreur lors de la réinitialisation.");
      setResetMsg("Mot de passe réinitialisé — communique-le à la personne.");
      setResetPassword("");
    } catch (e) {
      setResetErr(e.message || "Erreur lors de la réinitialisation.");
    }
  };

  const startEdit = (u) => { setEditingId(u.id); setRoleDraft(u.role); setSiteDraft(u.assignedSiteIds && u.assignedSiteIds.length ? u.assignedSiteIds : (u.assignedSiteId ? [u.assignedSiteId] : [])); };
  const saveEdit = () => {
    updateUserRole(editingId, roleDraft);
    updateUserSites(editingId, siteDraft);
    setEditingId(null);
  };
  const toggleSite = (list, setList, siteId) => {
    setList(list.includes(siteId) ? list.filter((id) => id !== siteId) : [...list, siteId]);
  };

  const isOnline = (u) => {
    if (!u.lastSeenAt) return false;
    return Date.now() - new Date(u.lastSeenAt).getTime() < 90 * 1000;
  };
  const lastSeenLabel = (u) => {
    if (!u.lastSeenAt) return "Jamais connecté";
    const diffMin = Math.round((Date.now() - new Date(u.lastSeenAt).getTime()) / 60000);
    if (diffMin < 2) return "À l'instant";
    if (diffMin < 60) return `Il y a ${diffMin} min`;
    const diffH = Math.round(diffMin / 60);
    if (diffH < 24) return `Il y a ${diffH} h`;
    return `Il y a ${Math.round(diffH / 24)} j`;
  };
  const sitesLabel = (u) => {
    const ids = u.assignedSiteIds && u.assignedSiteIds.length ? u.assignedSiteIds : (u.assignedSiteId ? [u.assignedSiteId] : []);
    if (ids.length === 0) return "Tous les sites";
    return ids.map((id) => sites.find((s) => s.id === id)?.name || id).join(", ");
  };

  const createAccount = async () => {
    setCreateErr(null); setCreateMsg(null);
    if (!form.fullName.trim() || (!form.username.trim() && !form.email.trim()) || !form.password) { setCreateErr("Le nom, le mot de passe, et soit l'e-mail soit le nom d'utilisateur, sont requis."); return; }
    if (form.email.trim() && !form.email.includes("@")) { setCreateErr("L'adresse e-mail n'a pas l'air valide."); return; }
    if (form.password.length < 6) { setCreateErr("Le mot de passe doit contenir au moins 6 caractères."); return; }
    setCreating(true);
    try {
      const res = await fetch("/api/create-user", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${session?.access_token || ""}` },
        body: JSON.stringify({ username: form.username.trim(), email: form.email.trim(), password: form.password, fullName: form.fullName.trim(), role: form.role, assignedSiteIds: form.assignedSiteIds }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Erreur lors de la création du compte.");
      setCreateMsg(`Compte créé pour ${form.fullName.trim()} (${ROLE_LABELS[form.role]}). Identifiant de connexion : "${data.loginEmail}" — communique-le avec le mot de passe.`);
      setForm({ fullName: "", username: "", email: "", password: "", role: "lecture", assignedSiteIds: [] });
    } catch (e) {
      setCreateErr(e.message || "Erreur lors de la création du compte.");
    } finally {
      setCreating(false);
    }
  };

  return (
    <div className="somip-fade" style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      <div className="somip-panel" style={{ flex: "1 1 560px", padding: 18 }}>
        <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Comptes ({profiles.length})</h3>
        <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.sub }}>
          Créés par toi ci-contre, ou par auto-inscription (rôle "Lecture" par défaut dans ce cas) — modifie le rôle et le site assigné ici à tout moment.
        </p>
        {deletingErr && <p style={{ color: C.danger, fontSize: 12.5, margin: "0 0 10px" }}>{deletingErr}</p>}
        <table className="somip-table">
          <thead><tr><th>Nom</th><th>Rôle</th><th>Site assigné</th><th>Présence</th><th>Statut</th><th></th></tr></thead>
          <tbody>
            {profiles.length === 0 && <EmptyRow colSpan={6} text="Aucun compte pour le moment." />}
            {profiles.map((u) => {
              const isEditing = editingId === u.id;
              return (
                <tr key={u.id}>
                  {isEditing ? (
                    <>
                      <td style={{ fontWeight: 600 }}>{u.name}</td>
                      <td>
                        <select className="somip-select" value={roleDraft} onChange={(e) => setRoleDraft(e.target.value)}>
                          {ROLE_VALUES.map((r) => <option key={r} value={r}>{ROLE_LABELS[r]}</option>)}
                        </select>
                      </td>
                      <td>
                        <div style={{ maxHeight: 110, overflowY: "auto", border: `1px solid ${C.border}`, borderRadius: 6, padding: 6, minWidth: 160 }}>
                          {sites.map((s) => (
                            <label key={s.id} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, padding: "2px 0", cursor: "pointer" }}>
                              <input type="checkbox" checked={siteDraft.includes(s.id)} onChange={() => toggleSite(siteDraft, setSiteDraft, s.id)} />
                              {s.name}
                            </label>
                          ))}
                        </div>
                        {siteDraft.length === 0 && <p style={{ margin: "4px 0 0", fontSize: 11, color: C.sub }}>Aucun coché = tous les sites</p>}
                      </td>
                      <td></td>
                      <td></td>
                      <td style={{ whiteSpace: "nowrap" }}>
                        <button className="somip-btn somip-btn-primary" style={{ padding: "5px 10px", fontSize: 12 }} onClick={saveEdit}>OK</button>
                        <button onClick={() => setEditingId(null)} style={{ border: "none", background: "none", cursor: "pointer", marginLeft: 4 }}><X size={16} color={C.sub} /></button>
                      </td>
                    </>
                  ) : (
                    <>
                      <td style={{ fontWeight: 600 }}>{u.name}</td>
                      <td><Badge color={C.blue}>{ROLE_LABELS[u.role] || u.role}</Badge></td>
                      <td style={{ color: C.sub, fontSize: 12.5, maxWidth: 220 }}>{sitesLabel(u)}</td>
                      <td>
                        {isOnline(u) ? (
                          <span style={{ display: "inline-flex", alignItems: "center", gap: 5, fontSize: 12, color: C.success, fontWeight: 600 }}>
                            <span style={{ width: 7, height: 7, borderRadius: "50%", background: C.success, display: "inline-block" }} /> En ligne
                          </span>
                        ) : (
                          <span style={{ fontSize: 12, color: C.sub }}>{lastSeenLabel(u)}</span>
                        )}
                      </td>
                      <td>
                        {u.active === false ? <Badge color={C.danger}>Désactivé</Badge> : <Badge color={C.success}>Actif</Badge>}
                      </td>
                      <td style={{ whiteSpace: "nowrap", textAlign: "right" }}>
                        <button onClick={() => startEdit(u)} style={{ border: "none", background: "none", cursor: "pointer", padding: 5 }}><Pencil size={14} color={C.sub} /></button>
                        {u.id !== session?.user?.id && (
                          <button onClick={() => toggleUserActive(u.id, u.active === false)} title={u.active === false ? "Réactiver" : "Désactiver"} style={{ border: "none", background: "none", cursor: "pointer", padding: 5 }}>
                            {u.active === false ? <CheckCircle2 size={14} color={C.success} /> : <CloudOff size={14} color={C.warning} />}
                          </button>
                        )}
                        <button onClick={() => { setResetForId(u.id); setResetPassword(""); setResetErr(null); setResetMsg(null); }} title="Réinitialiser le mot de passe" style={{ border: "none", background: "none", cursor: "pointer", padding: 5 }}>
                          <Lock size={14} color={C.sub} />
                        </button>
                        {u.id !== session?.user?.id && <ConfirmIconButton onConfirm={() => deleteAccount(u.id)} />}
                      </td>
                    </>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
        {resetForId && (
          <div className="somip-panel" style={{ marginTop: 14, padding: 14, background: C.bg }}>
            <p style={{ margin: "0 0 8px", fontSize: 12.5, fontWeight: 700, color: C.ink }}>
              Réinitialiser le mot de passe — {profiles.find((u) => u.id === resetForId)?.name}
            </p>
            <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
              <div style={{ flex: 1 }}>
                <Field label="Nouveau mot de passe"><input type="text" className="somip-input" value={resetPassword} onChange={(e) => setResetPassword(e.target.value)} placeholder="Au moins 6 caractères" /></Field>
              </div>
              <button className="somip-btn somip-btn-primary" style={{ padding: "9px 14px" }} onClick={submitReset}>Valider</button>
              <button onClick={() => setResetForId(null)} style={{ border: "none", background: "none", cursor: "pointer", padding: 9 }}><X size={16} color={C.sub} /></button>
            </div>
            {resetErr && <p style={{ color: C.danger, fontSize: 12, margin: "8px 0 0" }}>{resetErr}</p>}
            {resetMsg && <p style={{ color: C.success, fontSize: 12, margin: "8px 0 0" }}>{resetMsg}</p>}
          </div>
        )}
      </div>

      <div className="somip-panel" style={{ flex: "1 1 280px", padding: 18 }}>
        <h3 style={{ margin: "0 0 10px", fontSize: 14 }}>Créer un compte</h3>
        <Field label="Nom complet"><input className="somip-input" value={form.fullName} onChange={(e) => setForm({ ...form, fullName: e.target.value })} placeholder="Ex : Jean Mabiala" /></Field>
        <Field label="Adresse e-mail professionnelle (si la personne en a une)"><input type="email" className="somip-input" value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} placeholder="Ex : jean.mabiala@somip-sarl.ga" /></Field>
        {!form.email.trim() && (
          <Field label="Sinon, nom d'utilisateur (un identifiant sera généré)"><input className="somip-input" value={form.username} onChange={(e) => setForm({ ...form, username: e.target.value })} placeholder="Ex : jean.mabiala" /></Field>
        )}
        <Field label="Mot de passe provisoire"><input type="text" className="somip-input" value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} placeholder="Au moins 6 caractères" /></Field>
        <Field label="Rôle">
          <select className="somip-select" value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })}>
            {ROLE_VALUES.map((r) => <option key={r} value={r}>{ROLE_LABELS[r]}</option>)}
          </select>
        </Field>
        <Field label="Sites assignés (optionnel)">
          <div style={{ maxHeight: 130, overflowY: "auto", border: `1px solid ${C.border}`, borderRadius: 6, padding: 8 }}>
            {sites.map((s) => (
              <label key={s.id} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, padding: "3px 0", cursor: "pointer" }}>
                <input type="checkbox" checked={form.assignedSiteIds.includes(s.id)} onChange={() => toggleSite(form.assignedSiteIds, (v) => setForm({ ...form, assignedSiteIds: v }), s.id)} />
                {s.name}
              </label>
            ))}
          </div>
          <p style={{ margin: "4px 0 0", fontSize: 11, color: C.sub }}>Aucun coché = accès à tous les sites</p>
        </Field>
        {createErr && <p style={{ color: C.danger, fontSize: 12.5, margin: "0 0 10px" }}>{createErr}</p>}
        {createMsg && <p style={{ color: C.success, fontSize: 12.5, margin: "0 0 10px" }}>{createMsg}</p>}
        <button className="somip-btn somip-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={createAccount} disabled={creating}>
          <Plus size={15} /> {creating ? "Création..." : "Créer le compte"}
        </button>
        <p style={{ marginTop: 10, fontSize: 11, color: C.sub }}>
          Aucun e-mail requis : un identifiant de connexion est généré à partir du nom d'utilisateur. Communique cet identifiant et le mot de passe à la personne — elle les saisit à la place de l'e-mail pour se connecter.
        </p>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Historique                                                            */
/* ------------------------------------------------------------------ */
function HistoryView({ audit }) {
  return (
    <div className="somip-fade somip-panel" style={{ padding: 18 }}>
      <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Historique des modifications</h3>
      <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.sub }}>Les 300 dernières actions, les plus récentes en premier.</p>
      <table className="somip-table">
        <thead><tr><th>Date / heure</th><th>Utilisateur</th><th>Action</th><th>Détail</th></tr></thead>
        <tbody>
          {audit.length === 0 && <EmptyRow colSpan={4} text="Aucune action enregistrée pour le moment." />}
          {audit.map((a) => (
            <tr key={a.id}>
              <td className="somip-mono" style={{ whiteSpace: "nowrap" }}>{new Date(a.ts).toLocaleString("fr-FR")}</td>
              <td>{a.user}</td>
              <td><Badge color={C.blue}>{a.action}</Badge></td>
              <td style={{ color: C.sub }}>{a.detail}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

