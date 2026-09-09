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
  PRH: ["Compteur 1", "Compteur 2"],
  OKM: ["Compteur 1", "Compteur 2"],
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

const SETTINGS_SEED = { objectifFreinte: 3, logoUrl: null, colorPrimary: "#0071BD", colorAccent: "#F16B16" };
const rowToSettings = (r) => r ? {
  objectifFreinte: Number(r.objectif_freinte), logoUrl: r.logo_url || null,
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
async function exportToPdf({ filename, title, period, columns, rows, totalsRow }) {
  const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const [pR, pG, pB] = hexToRgb(C.blue), [aR, aG, aB] = hexToRgb(C.orange);
  // Bandeau bicolore SOMIP.
  doc.setFillColor(pR, pG, pB);
  doc.rect(0, 0, pageWidth * 0.6, 6, "F");
  doc.setFillColor(aR, aG, aB);
  doc.rect(pageWidth * 0.6, 0, pageWidth * 0.4, 6, "F");
  let textX = 30;
  if (CURRENT_LOGO_URL) {
    try {
      const dataUrl = await loadImageDataUrl(CURRENT_LOGO_URL);
      const fmt = dataUrl.includes("image/png") ? "PNG" : "JPEG";
      doc.addImage(dataUrl, fmt, 30, 14, 26, 26);
      textX = 64;
    } catch (e) { /* logo indisponible : on continue sans */ }
  }
  doc.setFontSize(15);
  doc.setTextColor(pR, pG, pB);
  doc.setFont(undefined, "bold");
  doc.text("SOMIP — Stock Gasoil", textX, 28);
  doc.setFontSize(9);
  doc.setTextColor(90, 100, 110);
  doc.setFont(undefined, "normal");
  doc.text("Zone Sud-Est · Gabon", textX, 42);
  doc.setFontSize(9);
  doc.text(`Édité le ${new Date().toLocaleDateString("fr-FR")} à ${new Date().toLocaleTimeString("fr-FR")}`, pageWidth - 30, 28, { align: "right" });
  doc.setDrawColor(pR, pG, pB);
  doc.setLineWidth(1);
  doc.line(30, 50, pageWidth - 30, 50);
  doc.setFontSize(13);
  doc.setTextColor(20, 30, 40);
  doc.setFont(undefined, "bold");
  doc.text(title, 30, 68);
  if (period) {
    doc.setFontSize(10);
    doc.setTextColor(aR, aG, aB);
    doc.setFont(undefined, "bold");
    doc.text(period, 30, 82);
  }
  autoTable(doc, {
    startY: 94,
    head: [columns],
    body: rows,
    foot: totalsRow ? [totalsRow] : undefined,
    theme: "grid",
    headStyles: { fillColor: [pR, pG, pB], textColor: 255, fontStyle: "bold" },
    footStyles: { fillColor: [aR, aG, aB, 0.15], textColor: [20, 30, 40], fontStyle: "bold" },
    styles: { fontSize: 9, cellPadding: 5 },
    margin: { left: 30, right: 30 },
  });
  doc.save(filename);
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

function ReportHeader({ title, period }) {
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
        <div style={{ textAlign: "right", fontSize: 11, color: C.sub }}>
          Édité le {new Date().toLocaleDateString("fr-FR")} à {new Date().toLocaleTimeString("fr-FR")}
        </div>
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
const ROLE_VALUES = ["superviseur", "operateur", "chauffeur", "lecture"];
const ROLE_LABELS = { superviseur: "Superviseur", operateur: "Opérateur", chauffeur: "Chauffeur", lecture: "Lecture" };
const PERIOD_TYPE_LABELS = { mensuel: "Mensuel", trimestriel: "Trimestriel", decadaire: "Décadaire" };
// canManage : sites, utilisateurs, réglages, modification/suppression, historique.
// canWrite  : peut ajouter des réceptions/sorties/inventaires (saisie).
function permsFor(role) {
  return {
    canManage: role === "superviseur",
    canWrite: role === "superviseur" || role === "operateur" || role === "chauffeur",
  };
}

/* ------------------------------------------------------------------ */
/* Base de données partagée (Supabase) — conversion lignes <-> objets  */
/* ------------------------------------------------------------------ */
const numOrUndef = (v) => (v === null || v === undefined ? undefined : Number(v));

const rowToSite = (r) => ({ id: r.id, code: r.code, name: r.name, capacity: Number(r.capacity), stockInitial: Number(r.stock_initial), isMobile: !!r.is_mobile });
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

const rowToAssignment = (r) => ({ id: r.id, truckId: r.truck_id, stationId: r.station_id, startDate: r.start_date, endDate: r.end_date || null });
const assignmentToRow = (a) => ({ truck_id: a.truckId, station_id: a.stationId, start_date: a.startDate, end_date: a.endDate ?? null });


const rowToAudit = (r) => ({ id: r.id, ts: r.ts, user: r.user_name, action: r.action, detail: r.detail });
const rowToProfile = (r) => ({ id: r.id, name: r.full_name, role: r.role, lastSeenAt: r.last_seen_at || null, assignedSiteId: r.assigned_site_id || null });

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
  const [inventaires, setInventaires] = useState([]);
  const [productStocks, setProductStocks] = useState([]);
  const [siteMeters, setSiteMeters] = useState([]);
  const [bilans, setBilans] = useState([]);
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
          const [sitesData, movementsData, inventairesData, profilesData, auditData, productStocksData, assignmentsData, siteMetersData, bilansData] = await Promise.all([
            fetchTable("sites", rowToSite),
            fetchTable("movements", rowToMovement, "date"),
            fetchTable("inventaires", rowToInventaire, "date"),
            fetchTable("profiles", rowToProfile),
            fetchTable("audit", rowToAudit, "ts", false),
            fetchTable("product_stocks", rowToProductStock),
            fetchTable("truck_assignments", rowToAssignment, "start_date"),
            fetchTable("site_meters", rowToSiteMeter, "name"),
            fetchTable("bilan_matieres", rowToBilan, "period_key"),
          ]);
          let settingsRow = null;
          try {
            const res = await supabase.from("settings").select("*").eq("id", 1).maybeSingle();
            settingsRow = res.data;
          } catch (e) { /* réglages optionnels : on garde la valeur par défaut si ça échoue */ }
          return { sitesData, movementsData, inventairesData, profilesData, auditData, productStocksData, assignmentsData, siteMetersData, bilansData, settingsRow };
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
      const [s, m, i, p, a, ps, ta, sm, bl] = await Promise.all([
        fetchTable("sites", rowToSite),
        fetchTable("movements", rowToMovement, "date"),
        fetchTable("inventaires", rowToInventaire, "date"),
        fetchTable("profiles", rowToProfile),
        fetchTable("audit", rowToAudit, "ts", false),
        fetchTable("product_stocks", rowToProductStock),
        fetchTable("truck_assignments", rowToAssignment, "start_date"),
        fetchTable("site_meters", rowToSiteMeter, "name"),
        fetchTable("bilan_matieres", rowToBilan, "period_key"),
      ]);
      setSites(s); setMovements(m); setInventaires(i); setProfiles(p); setAudit(a); setProductStocks(ps); setTruckAssignments(ta); setSiteMeters(sm); setBilans(bl);
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
  const saveBilan = ({ periodType, periodKey, reception15, ventes15, transferts15, stockFin15, commentaire, photoFiles = [], existingPhotoUrls = [] }) => withSync(async () => {
    const newUrls = photoFiles.length ? await uploadPhotos(photoFiles, `bilans/global/${periodType}`) : [];
    const photoUrls = [...existingPhotoUrls, ...newUrls];
    const row = bilanToRow({ siteId: null, periodType, periodKey, reception15: Number(reception15) || 0, ventes15: Number(ventes15) || 0, transferts15: Number(transferts15) || 0, stockFin15: Number(stockFin15) || 0, commentaire, photoUrls, createdBy: currentUserName });
    const { data, error } = await supabase.from("bilan_matieres").upsert(row, { onConflict: "period_type,period_key" }).select().maybeSingle();
    if (error) throw error;
    if (!data) throw new Error("Le Bilan Matières n'a pas pu être confirmé par le serveur — réessaie.");
    const saved = rowToBilan(data);
    setBilans((prev) => {
      const exists = prev.some((b) => !b.siteId && b.periodType === periodType && b.periodKey === periodKey);
      return exists ? prev.map((b) => (!b.siteId && b.periodType === periodType && b.periodKey === periodKey ? saved : b)) : [...prev, saved];
    });
    appendAudit("Bilan Matières", `${PERIOD_TYPE_LABELS[periodType]} ${periodKey}`);
    flash("Bilan Matières enregistré.");
  });
  const deleteBilan = (bilan) => withSync(async () => {
    const { data, error } = await supabase.from("bilan_matieres").delete().eq("id", bilan.id).select();
    if (error) throw error;
    if (!data || data.length === 0) throw new Error("Suppression refusée par la base de données.");
    setBilans((prev) => prev.filter((b) => b.id !== bilan.id));
    flash("Bilan supprimé.");
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
    const { data, error } = await supabase.from("movements").insert(movementToRow(record)).select().maybeSingle();
    if (error) throw error;
    const saved = data ? rowToMovement(data) : record;
    setMovements((prev) => [...prev, saved]);
    appendAudit(TYPE_META[payload.type].label, `${fmt(payload.quantity)} L — ${sites.find((s) => s.id === payload.siteId)?.name || ""}`);
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
    const photoUrls = photoFiles.length ? await uploadPhotos(photoFiles, `inventaires/${siteId}`) : [];
    const invDraft = {
      siteId, product, date, stockPhysique, commentaire, basisEcart,
      stockTheoriqueAmbiant: theoriqueAmbiant, stockTheorique15: theorique15,
      stockTheorique: theoriqueUsed, stockPhysiqueUsed: physiqueUsed,
      ecart: cls.ecartL, ecartPermille: cls.ecartPermille, nature: cls.nature,
      tauxFreinte: cls.tauxFreinte, objectifUtilise: cls.objectif, conformite: cls.conformite,
      adjustmentId: null, photoUrls, createdBy: currentUserName, createdAt: new Date().toISOString(), ...vcfFields,
    };
    const { data: dataI, error: e2 } = await supabase.from("inventaires").insert(inventaireToRow(invDraft)).select().maybeSingle();
    if (e2) throw e2;
    if (!dataI) throw new Error("L'inventaire n'a pas pu être confirmé par le serveur — réessaie.");
    const invRecord = rowToInventaire(dataI);
    setInventaires((prev) => [...prev, invRecord]);
    appendAudit("Inventaire", `${sites.find((s) => s.id === siteId)?.name || ""} — base ${has15 ? "15°C" : "ambiante"} — ${NATURE_META[cls.nature].label} ${ecart >= 0 ? "+" : ""}${fmt(ecart)} L (${cls.ecartPermille >= 0 ? "+" : ""}${cls.ecartPermille.toFixed(2)} ‰)`);
    flash("Inventaire enregistré.");
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
  const updateTheme = ({ logoFile, colorPrimary, colorAccent }) => withSync(async () => {
    let logoUrl = settings.logoUrl;
    if (logoFile) {
      const urls = await uploadPhotos([logoFile], "branding");
      logoUrl = urls[0];
    }
    const next = { ...settings, logoUrl, colorPrimary: colorPrimary || settings.colorPrimary, colorAccent: colorAccent || settings.colorAccent };
    const { error } = await supabase.from("settings").update({
      logo_url: next.logoUrl, color_primary: next.colorPrimary, color_accent: next.colorAccent,
    }).eq("id", 1);
    if (error) throw error;
    setSettings(next);
    appendAudit("Personnalisation", "Logo et/ou couleurs mis à jour");
    flash("Personnalisation enregistrée.");
  });
  useEffect(() => { applyTheme(settings.colorPrimary, settings.colorAccent); setCurrentLogoUrl(settings.logoUrl); }, [settings.colorPrimary, settings.colorAccent, settings.logoUrl]);

  /* ---- mutations : rôle d'un utilisateur (Superviseur uniquement) ---- */
  const updateUserRole = (userId, role) => withSync(async () => {
    const { error } = await supabase.from("profiles").update({ role }).eq("id", userId);
    if (error) throw error;
    const target = profiles.find((u) => u.id === userId);
    setProfiles((prev) => prev.map((u) => (u.id === userId ? { ...u, role } : u)));
    appendAudit("Modification rôle utilisateur", `${target?.name || ""} → ${ROLE_LABELS[role]}`);
    flash("Rôle mis à jour.");
  });
  const updateUserSite = (userId, assignedSiteId) => withSync(async () => {
    const { error } = await supabase.from("profiles").update({ assigned_site_id: assignedSiteId || null }).eq("id", userId);
    if (error) throw error;
    const target = profiles.find((u) => u.id === userId);
    setProfiles((prev) => prev.map((u) => (u.id === userId ? { ...u, assignedSiteId: assignedSiteId || null } : u)));
    appendAudit("Modification site assigné", `${target?.name || ""} → ${sites.find((s) => s.id === assignedSiteId)?.name || "Tous les sites"}`);
    flash("Site assigné mis à jour.");
  });

  const NAV = [
    { id: "dashboard", label: "Tableau de bord", icon: LayoutDashboard, show: true },
    { id: "sites", label: "Sites", icon: Factory, show: perms.canManage },
    { id: "saisie", label: "Saisie journalière", icon: ClipboardList, show: true },
    { id: "inventaires", label: "Inventaires", icon: ClipboardList, show: true },
    { id: "vcf", label: "Correction 15°C", icon: Thermometer, show: true },
    { id: "rapports", label: "Rapports", icon: FileBarChart, show: true },
    { id: "utilisateurs", label: "Utilisateurs", icon: Users, show: perms.canManage },
    { id: "personnalisation", label: "Personnalisation", icon: Palette, show: perms.canManage },
    { id: "historique", label: "Historique", icon: History, show: perms.canManage },
  ].filter((n) => n.show);
  const viewTitle = NAV.find((n) => n.id === view)?.label || "";

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
              <SyncIndicator status={syncStatus} lastSync={lastSync} />
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
          {view === "dashboard" && <Dashboard sites={sites} movements={movements} inventaires={inventaires} stockOf={stockOf} purgeDemoMovements={purgeDemoMovements} canManage={perms.canManage} />}
          {view === "sites" && perms.canManage && <SitesView sites={sites} movements={movements} stockOf={stockOf} addSite={addSite} editSite={editSite} removeSite={removeSite} productStocks={productStocks} saveProductStock={saveProductStock} truckAssignments={truckAssignments} assignTruck={assignTruck} siteMeters={siteMeters} addSiteMeter={addSiteMeter} removeSiteMeter={removeSiteMeter} />}
          {view === "saisie" && <DailyEntryView sites={sites} movements={movements} inventaires={inventaires} productStocks={productStocks} siteMeters={siteMeters} saveProductStock={saveProductStock} addMovement={addMovement} addInventaire={addInventaire} deleteMovement={deleteMovement} deleteInventaire={deleteInventaire} settings={settings} canWrite={perms.canWrite} canManage={perms.canManage} assignedSiteId={profile?.assignedSiteId} />}
          {view === "inventaires" && <InventairesView sites={sites} inventaires={inventaires} stockOf={stockOf} stockOf15={stockOf15} addInventaire={addInventaire} deleteInventaire={deleteInventaire} settings={settings} updateSettings={updateSettings} canWrite={perms.canWrite} canManage={perms.canManage} />}
          {view === "vcf" && <VcfView />}
          {view === "rapports" && <ReportsView sites={sites} movements={movements} inventaires={inventaires} productStocks={productStocks} truckAssignments={truckAssignments} settings={settings} stockOf={stockOf} bilans={bilans} saveBilan={saveBilan} deleteBilan={deleteBilan} canManage={perms.canManage} />}
          {view === "utilisateurs" && perms.canManage && <UsersView profiles={profiles} updateUserRole={updateUserRole} updateUserSite={updateUserSite} sites={sites} session={session} />}
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
function Dashboard({ sites, movements, inventaires, stockOf, purgeDemoMovements, canManage }) {
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

  return (
    <div className="somip-fade">
      <p style={{ marginTop: -8, marginBottom: 14, fontSize: 13, color: C.sub }}>
        Vos données sont sauvegardées automatiquement et restent disponibles après fermeture ou actualisation de la page.
        Les capacités et stocks initiaux des sites restent des valeurs à vérifier/ajuster depuis la page Sites.
      </p>

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
      </div>

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
function SitesView({ sites, movements, stockOf, addSite, editSite, removeSite, productStocks, saveProductStock, truckAssignments, assignTruck, siteMeters, addSiteMeter, removeSiteMeter }) {
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
  const displayedMeters = currentMeters.length ? currentMeters : metersForSite(stations.find((s) => s.id === meterSiteId)).map((name, i) => ({ id: `default-${i}`, name, isDefault: true }));
  const submitMeter = () => {
    if (!newMeterName.trim() || !meterSiteId) return;
    addSiteMeter({ siteId: meterSiteId, name: newMeterName });
    setNewMeterName("");
  };

  return (
    <div className="somip-fade" style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      <div className="somip-panel" style={{ flex: "2 1 520px", padding: 18 }}>
        <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Sites externalisés ({sites.length})</h3>
        <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.sub }}>Capacités et stocks initiaux : à vérifier et ajuster selon vos valeurs réelles.</p>
        <table className="somip-table">
          <thead><tr><th>Code</th><th>Site</th><th style={{ textAlign: "right" }}>Capacité (L)</th><th style={{ textAlign: "right" }}>Stock actuel (L)</th><th></th></tr></thead>
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
                      <td style={{ whiteSpace: "nowrap", textAlign: "right" }}>
                        <button onClick={() => startEdit(s)} style={{ border: "none", background: "none", cursor: "pointer", padding: 5 }}><Pencil size={14} color={C.sub} /></button>
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
        <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.sub }}>Ajoute ou retire un compteur si la configuration physique d'un site change.</p>
        <Field label="Site">
          <select className="somip-select" value={meterSiteId} onChange={(e) => setMeterSiteId(e.target.value)}>
            {stations.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
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
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Réceptions                                                            */
/* ------------------------------------------------------------------ */
/* ------------------------------------------------------------------ */
/* Saisie journalière (écran unique : réception, sortie/camion, retour) */
/* ------------------------------------------------------------------ */
function DailyEntryView({ sites, movements, inventaires, productStocks, siteMeters, addMovement, addInventaire, deleteMovement, deleteInventaire, settings, canWrite, canManage, assignedSiteId }) {
  const availableSites = assignedSiteId ? sites.filter((s) => s.id === assignedSiteId) : sites;
  const [siteId, setSiteId] = useState(assignedSiteId || sites[0]?.id || "");
  const [product, setProduct] = useState("gasoil");
  const [date, setDate] = useState(todayStr());
  const [receptionQty, setReceptionQty] = useState("");
  const [receptionRef, setReceptionRef] = useState("");
  const [indexAvant, setIndexAvant] = useState("");
  const [indexApres, setIndexApres] = useState("");
  const [compteur, setCompteur] = useState("");
  const [compteurReadings, setCompteurReadings] = useState([{ compteur: "", indexAvant: "", indexApres: "" }]);
  const [chargements, setChargements] = useState([{ camion: "", quantite: "" }]);
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
  const receptionN = Number(receptionQty) || 0;
  const retourN = (isLub || isMobileSite) ? 0 : (Number(retourQty) || 0);
  const retourCuveTruckN = isMobileSite ? (Number(retourCuveTruckQty) || 0) : 0;
  const isMultiCompteurEntry = isLubSite && !isLub && !isMobileSite;
  const readingFlows = compteurReadings.map((r) => ({
    ...r,
    flow: r.indexAvant !== "" && r.indexApres !== "" ? Number(r.indexApres) - Number(r.indexAvant) : 0,
    valid: r.indexAvant === "" && r.indexApres === "" ? true : (r.indexAvant !== "" && r.indexApres !== "" && Number(r.indexApres) > Number(r.indexAvant)),
  }));
  const sortieQtySingle = indexAvant !== "" && indexApres !== "" ? Number(indexApres) - Number(indexAvant) : 0;
  const sortieValidSingle = indexAvant === "" && indexApres === "" ? true : (indexAvant !== "" && indexApres !== "" && sortieQtySingle > 0);
  const sortieQty = isMultiCompteurEntry ? readingFlows.reduce((a, r) => a + r.flow, 0) : sortieQtySingle;
  const sortieValid = isMultiCompteurEntry ? readingFlows.every((r) => r.valid) : sortieValidSingle;
  const totalChargements = isLubSite && !isLub ? chargements.reduce((a, c) => a + (Number(c.quantite) || 0), 0) : 0;
  const chargementsValid = totalChargements <= sortieQty;
  const venteStation = isLubSite && !isLub ? Math.max(0, sortieQty - totalChargements) : sortieQty;
  const stockTheoriqueAmbiant = stockDebutEffective + receptionN + retourN - sortieQty - retourCuveTruckN;
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
  const canSubmit = sortieValid && chargementsValid && hasSomethingToSave && !stockFinConflict && (!isFirstOfMonth || !hasSomethingToSave || stockDebutConfirm !== "");

  const resetDayFields = () => {
    setReceptionQty(""); setReceptionRef("");
    setIndexAvant(""); setIndexApres(""); setDestinataire(""); setChargements([{ camion: "", quantite: "" }]);
    setCompteurReadings([{ compteur: meters[0] || "Compteur", indexAvant: "", indexApres: "" }]);
    setRetourQty(""); setRetourNote(""); setRetourCamionTruckId(""); setRetourCuveTruckQty(""); setRetourCuveTruckNote(""); setTempC(""); setDensite("");
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
        const ok = await addMovement({ siteId, product, type: "reception", date, quantity: receptionN, delta: receptionN, ref: receptionRef, ...vcfExtra(receptionN) });
        if (!ok) return;
      }
      if (sortieQty > 0) {
        const compteurField = meters.length > 1 ? compteur : undefined;
        if (isLub || isMobileSite) {
          const ok = await addMovement({ siteId, product, type: "sortie", date, quantity: sortieQty, delta: -sortieQty, indexAvant: Number(indexAvant), indexApres: Number(indexApres), destinataire, compteur: compteurField });
          if (!ok) return;
        } else if (isMultiCompteurEntry) {
          // Le(s) compteur(s) mesurent le flux total (vente + chargements camions confondus) :
          // on répartit les chargements sur les compteurs saisis (dans l'ordre), puis on
          // enregistre le reliquat "vente" par compteur avec son propre index. Chaque
          // chargement camion crée automatiquement la réception correspondante côté camion.
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
            }
          }
        } else {
          const ok = await addMovement({ siteId, product, type: "sortie", date, quantity: sortieQty, delta: -sortieQty, indexAvant: Number(indexAvant), indexApres: Number(indexApres), compteur: compteurField, destinataire, ...vcfExtra(sortieQty) });
          if (!ok) return;
        }
      }
      if (retourN > 0) {
        const ok = await addMovement({ siteId, product, type: "retour_camion", date, quantity: retourN, delta: retourN, camion: retourCamionTruckId || undefined, destination: retourNote, ...vcfExtra(retourN) });
        if (!ok) return;
      }
      if (retourCuveTruckN > 0) {
        const ok = await addMovement({ siteId, product, type: "retour_cuve_camion", date, quantity: retourCuveTruckN, delta: -retourCuveTruckN, destination: retourCuveTruckNote, ...vcfExtra(retourCuveTruckN) });
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
                {assignedSiteId ? (
                  <div className="somip-input" style={{ background: C.bg, color: C.ink, fontWeight: 600, display: "flex", alignItems: "center" }}>
                    {sites.find((s) => s.id === assignedSiteId)?.name || "Site assigné"}
                  </div>
                ) : (
                  <select className="somip-select" value={siteId} onChange={(e) => setSiteId(e.target.value)}>
                    {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
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
          <div style={{ display: "flex", gap: 8 }}>
            <div style={{ flex: 1 }}><Field label="Quantité reçue (L)"><input type="number" className="somip-input" value={receptionQty} onChange={(e) => setReceptionQty(e.target.value)} placeholder="0" /></Field></div>
            {!isMobileSite && <div style={{ flex: 1 }}><Field label="N° Bon de livraison"><input className="somip-input" value={receptionRef} onChange={(e) => setReceptionRef(e.target.value)} placeholder="BL-XXXX" /></Field></div>}
          </div>
          {isMobileSite && (
            <p style={{ margin: "-6px 0 10px", fontSize: 11, color: C.warning }}>
              Pense à saisir aussi ce chargement côté Prehomo/Okouma (Chargement laitiers) — les deux côtés sont indépendants. Pas de N° de bon pour un chargement interne.
            </p>
          )}
          {isLub && receptionN > 0 && <p style={{ margin: "-6px 0 10px", fontSize: 11, color: C.sub }}>≈ {fmt(receptionN * lubDensite)} kg</p>}

          {(isLub || isMobileSite) ? (
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
          ) : isLubSite ? (
            <>
              <p style={{ margin: "10px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Sortie (compteurs — flux total : vente + chargements camions)</p>
              {compteurReadings.map((r, idx) => {
                const lastIdx = r.compteur ? lastIndexForMeter(r.compteur) : undefined;
                const mismatch = lastIdx !== undefined && r.indexAvant !== "" && Number(r.indexAvant) !== lastIdx;
                return (
                  <div key={idx} style={{ border: `1px solid ${C.border}`, borderRadius: 8, padding: 10, marginBottom: 8 }}>
                    <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
                      {meters.length > 1 && (
                        <div style={{ flex: 1 }}>
                          <Field label="Compteur">
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

              <p style={{ margin: "12px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Chargement laitiers (prélevé sur ce flux)</p>
              <p style={{ margin: "0 0 8px", fontSize: 11, color: C.warning }}>Pense à saisir aussi ce chargement côté camion (Chargement) — les deux côtés sont indépendants.</p>
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
              <div style={{ display: "flex", gap: 8 }}>
                <div style={{ flex: 1 }}><Field label="Quantité retournée (L)"><input type="number" className="somip-input" value={retourQty} onChange={(e) => setRetourQty(e.target.value)} placeholder="0" /></Field></div>
                <div style={{ flex: 1 }}>
                  <Field label="Camion">
                    <select className="somip-select" value={retourCamionTruckId} onChange={(e) => setRetourCamionTruckId(e.target.value)}>
                      <option value="">— non précisé —</option>
                      {truckSites.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                    </select>
                  </Field>
                </div>
              </div>
              {retourCamionTruckId && (
                <p style={{ margin: "-6px 0 10px", fontSize: 11, color: C.sub }}>
                  Ce retour ({fmt(Number(retourQty) || 0)} L) concerne {truckSites.find((t) => t.id === retourCamionTruckId)?.name}.
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
                <p style={{ margin: "-6px 0 12px", fontSize: 11, color: C.sub }}>
                  Ce retour ({fmt(retourCuveTruckN)} L) concerne {stationSites.find((s) => s.id === retourCuveTruckNote)?.name}.
                </p>
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
              return (
                <tr key={m.id}>
                  <td><Badge color={meta.color}>{label}</Badge></td>
                  <td style={{ color: C.sub }}>{detail}</td>
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
function InventairesView({ sites, inventaires, stockOf, stockOf15, addInventaire, deleteInventaire, settings, updateSettings, canWrite, canManage }) {
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
    <div className="somip-fade" style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
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
    { id: "exposition_comilog", label: "Exposition Comilog", superviseurOnly: true },
    { id: "bons", label: "Bons de livraison", superviseurOnly: true },
    { id: "bilan", label: "Bilan Matières", superviseurOnly: true },
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
      {tab === "exposition" && canManage && <ExposureReport sites={sites} movements={movements} inventaires={inventaires} truckAssignments={truckAssignments} />}
      {tab === "exposition_comilog" && canManage && <ExpositionComilogReport sites={sites} movements={movements} inventaires={inventaires} truckAssignments={truckAssignments} />}
      {tab === "bons" && canManage && <DeliveryNotesReport sites={sites} movements={movements} />}
      {tab === "bilan" && canManage && <BilanMatieresView sites={sites} bilans={bilans} saveBilan={saveBilan} deleteBilan={deleteBilan} canManage={canManage} />}
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
      const ventes = sumQty(dayMovs, ["sortie"]);
      const chargementLaitiers = isTruck ? 0 : sumQty(dayMovs, ["sortie_camion"]);
      const retourCamions = isTruck ? 0 : sumQty(dayMovs, ["retour_camion"]);
      const retourCuve = isTruck ? sumQty(dayMovs, ["retour_cuve_camion"]) : 0;
      const stockTheorique = stockDebut + reception + retourCamions - ventes - chargementLaitiers - retourCuve;
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
      const ventes = sumQty15(dayMovs, ["sortie"]);
      const chargementLaitiers = isTruck ? 0 : sumQty15(dayMovs, ["sortie_camion"]);
      const retourCamions = isTruck ? 0 : sumQty15(dayMovs, ["retour_camion"]);
      const retourCuve = isTruck ? sumQty15(dayMovs, ["retour_cuve_camion"]) : 0;
      const stockTheorique = stockDebut + reception + retourCamions - ventes - chargementLaitiers - retourCuve;
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

function ExposureReport({ sites, movements, inventaires, truckAssignments }) {
  const [month, setMonth] = useState(currentMonth());
  const [decadeNum, setDecadeNum] = useState(1);
  const bounds = decadeBoundsExplicit(month, decadeNum);
  const fixedSites = sites.filter((s) => !s.isMobile);

  const rows = fixedSites.map((s) => {
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
        ventesTrucks += sumQty(dayMovsTruck, ["sortie"]);
      }
      cur.setDate(cur.getDate() + 1);
    }
    const ventesCumulees = ventesSite + ventesTrucks;
    const jaugeInv = pickLatestInv(inventaires.filter((i) => i.siteId === s.id && (i.product || "gasoil") === "gasoil" && i.date <= bounds.end));
    const jaugeADate = jaugeInv ? jaugeInv.stockPhysique : null;
    // Les camions ne livrent que par multiples de 5000 L (5000/15000/20000/35000) :
    // la demande d'approvisionnement est toujours arrondie à l'inférieur, au multiple de 5000 le plus proche.
    const demandeAppro = jaugeADate !== null ? Math.floor(Math.max(0, s.capacity - jaugeADate) / 5000) * 5000 : null;
    return { site: s, ventesCumulees, jaugeADate, demandeAppro };
  });
  const totalVentes = rows.reduce((a, r) => a + r.ventesCumulees, 0);
  const totalDemande = rows.reduce((a, r) => a + (r.demandeAppro || 0), 0);

  const doExcel = () => exportToExcel(`SOMIP_Exposition_${month}_D${decadeNum}.xlsx`, [{
    name: "Exposition", rows: rows.map((r) => ({
      Site: r.site.name, "Ventes cumulées décade (L)": Math.round(r.ventesCumulees),
      "Jauge à date (L)": r.jaugeADate !== null ? Math.round(r.jaugeADate) : "",
      "Demande d'approvisionnement (L)": r.demandeAppro !== null ? Math.round(r.demandeAppro) : "",
    })),
  }]);

  const doPdf = () => exportToPdf({
    filename: `SOMIP_Exposition_${month}_D${decadeNum}.pdf`,
    title: "Exposition — Ventes cumulées & demande d'approvisionnement",
    period: `${bounds.label} — ${bounds.start} au ${bounds.end}`,
    columns: ["Site", "Ventes cumulées (décade)", "Jauge à date", "Demande d'approvisionnement"],
    rows: rows.map((r) => [
      `${r.site.name} (${r.site.code})`,
      `${fmt(r.ventesCumulees)} L`,
      r.jaugeADate !== null ? `${fmt(r.jaugeADate)} L` : "—",
      r.demandeAppro !== null ? `${fmt(r.demandeAppro)} L` : "—",
    ]),
    totalsRow: ["Total réseau", `${fmt(totalVentes)} L`, "", `${fmt(totalDemande)} L`],
  });

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14, display: "flex", gap: 12, flexWrap: "wrap" }}>
        <Field label="Mois"><input type="month" className="somip-input" style={{ maxWidth: 200 }} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
        <Field label="Décade">
          <select className="somip-select" style={{ maxWidth: 260 }} value={decadeNum} onChange={(e) => setDecadeNum(Number(e.target.value))}>
            <option value={1}>1ère décade (1 au 10)</option>
            <option value={2}>2e décade (11 au 20)</option>
            <option value={3}>3e décade (21 à la fin)</option>
          </select>
        </Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="Exposition — Ventes cumulées &amp; demande d'approvisionnement" period={`${bounds.label} — ${bounds.start} au ${bounds.end}`} />
        <ReportToolbar onExcel={doExcel} onPdf={doPdf} onPrint={() => window.print()} />

        <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 20 }}>
          <StatCard label="Ventes cumulées réseau (décade)" value={fmt(totalVentes)} unit="L" accent={C.blue} icon={ArrowUpCircle} />
          <StatCard label="Demande d'approvisionnement réseau" value={fmt(totalDemande)} unit="L" accent={C.orange} icon={Truck} />
        </div>

        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead>
              <tr>
                <th>Site</th><th style={{ textAlign: "right" }}>Ventes cumulées (décade)</th>
                <th style={{ textAlign: "right" }}>Jauge à date</th><th style={{ textAlign: "right" }}>Demande d'approvisionnement</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name} <span style={{ color: C.sub, fontWeight: 500 }}>({r.site.code})</span></td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600 }}>{fmt(r.ventesCumulees)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{r.jaugeADate !== null ? `${fmt(r.jaugeADate)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.orange }}>{r.demandeAppro !== null ? `${fmt(r.demandeAppro)} L` : "—"}</td>
                </tr>
              ))}
              <tr>
                <td style={{ fontWeight: 700 }}>Total réseau</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(totalVentes)} L</td>
                <td></td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.orange }}>{fmt(totalDemande)} L</td>
              </tr>
            </tbody>
          </table>
        </div>
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          Ventes cumulées = somme des ventes (et sorties vers camion) de la décade sélectionnée. Jauge à date = dernière mesure physique connue à la fin de la décade. Demande d'approvisionnement = Capacité − Jauge à date, arrondie à l'inférieur au multiple de 5000 L (livraisons par camions de 5000/15000/20000/35000 L).
        </p>
      </div>
    </div>
  );
}

const COMILOG_SITE_CODES = ["PRH", "OKM", "CIM", "CMM", "GTR"];

/* ---- Exposition Comilog — envoi quotidien : ventes & réception de la veille, creux à date ---- */
function ExpositionComilogReport({ sites, movements, inventaires, truckAssignments }) {
  const [mvtDate, setMvtDate] = useState(todayStr());
  const [stockDate, setStockDate] = useState(todayStr());
  const comilogSites = sites.filter((s) => COMILOG_SITE_CODES.includes(s.code));

  const rows = comilogSites.map((s) => {
    // Ventes = ventes propres du site (hors chargements laitiers, transfert interne, pas une
    // vente) + les ventes du/des camion(s) rattaché(s) ce jour-là (Sortie Fiche Terrain).
    const dayMovs = movements.filter((m) => m.siteId === s.id && (m.product || "gasoil") === "gasoil" && m.date === mvtDate);
    const ventesSite = sumQty(dayMovs, ["sortie"]);
    const reception = sumQty(dayMovs, ["reception"]);
    let ventesTrucks = 0;
    for (const truckId of trucksAssignedAt(truckAssignments, s.id, mvtDate)) {
      const dayMovsTruck = movements.filter((m) => m.siteId === truckId && (m.product || "gasoil") === "gasoil" && m.date === mvtDate);
      ventesTrucks += sumQty(dayMovsTruck, ["sortie"]);
    }
    const ventes = ventesSite + ventesTrucks;
    const jaugeInv = pickLatestInv(inventaires.filter((i) => i.siteId === s.id && (i.product || "gasoil") === "gasoil" && i.date <= stockDate));
    const jaugeADate = jaugeInv ? jaugeInv.stockPhysique : stockThroughDate(s, movements, stockDate, inventaires);
    const creux = Math.max(0, s.capacity - jaugeADate);
    return { site: s, ventes, reception, jaugeADate, creux };
  });
  const totalVentes = rows.reduce((a, r) => a + r.ventes, 0);
  const totalReception = rows.reduce((a, r) => a + r.reception, 0);
  const totalCreux = rows.reduce((a, r) => a + r.creux, 0);

  const doExcel = () => exportToExcel(`SOMIP_Exposition_Comilog_${stockDate}.xlsx`, [{
    name: "Exposition Comilog", rows: rows.map((r) => ({
      Site: r.site.name, [`Ventes du ${mvtDate} (L)`]: Math.round(r.ventes), [`Réception du ${mvtDate} (L)`]: Math.round(r.reception),
      [`Creux au ${stockDate} (L)`]: Math.round(r.creux),
    })),
  }]);

  const doPdf = () => exportToPdf({
    filename: `SOMIP_Exposition_Comilog_${stockDate}.pdf`,
    title: "Exposition Comilog",
    period: `Ventes & réception du ${mvtDate} — Creux au ${stockDate}`,
    columns: ["Site", "Ventes", "Réception", "Creux"],
    rows: rows.map((r) => [r.site.name, `${fmt(r.ventes)} L`, `${fmt(r.reception)} L`, `${fmt(r.creux)} L`]),
    totalsRow: ["Total", `${fmt(totalVentes)} L`, `${fmt(totalReception)} L`, `${fmt(totalCreux)} L`],
  });

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14, display: "flex", gap: 12, flexWrap: "wrap" }}>
        <Field label="Date des ventes / réception"><input type="date" className="somip-input" style={{ maxWidth: 200 }} value={mvtDate} onChange={(e) => setMvtDate(e.target.value)} /></Field>
        <Field label="Date du stock / creux"><input type="date" className="somip-input" style={{ maxWidth: 200 }} value={stockDate} onChange={(e) => setStockDate(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="Exposition Comilog" period={`Ventes & réception du ${mvtDate} — Creux au ${stockDate}`} />
        <ReportToolbar onExcel={doExcel} onPdf={doPdf} onPrint={() => window.print()} />

        <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 20 }}>
          <StatCard label={`Ventes cumulées (${mvtDate})`} value={fmt(totalVentes)} unit="L" accent={C.blue} icon={ArrowUpCircle} />
          <StatCard label={`Réception cumulée (${mvtDate})`} value={fmt(totalReception)} unit="L" accent={C.success} icon={ArrowDownCircle} />
          <StatCard label={`Creux cumulé (${stockDate})`} value={fmt(totalCreux)} unit="L" accent={C.orange} icon={Truck} />
        </div>

        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead>
              <tr>
                <th>Site</th><th style={{ textAlign: "right" }}>Ventes ({mvtDate})</th>
                <th style={{ textAlign: "right" }}>Réception ({mvtDate})</th><th style={{ textAlign: "right" }}>Creux ({stockDate})</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name} <span style={{ color: C.sub, fontWeight: 500 }}>({r.site.code})</span></td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600 }}>{fmt(r.ventes)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>{fmt(r.reception)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.orange }}>{fmt(r.creux)} L</td>
                </tr>
              ))}
              <tr>
                <td style={{ fontWeight: 700 }}>Total</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(totalVentes)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.success }}>{fmt(totalReception)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.orange }}>{fmt(totalCreux)} L</td>
              </tr>
            </tbody>
          </table>
        </div>
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          Sites Comilog : Prehomo, Okouma, CIM, CMM, Gare Traction. Ventes/Réception = celles du jour choisi ci-dessus (indépendant de la date du stock). Creux = Capacité − Stock à la date choisie (jauge mesurée si disponible, sinon théorique).
        </p>
      </div>
    </div>
  );
}

/* ---- Synthèse Gasoil (les 8 sites) ---- */
function GasoilSynthesisReport({ sites, movements, inventaires }) {
  const [month, setMonth] = useState(currentMonth());
  const bounds = monthBounds(month);

  const rows = sites.filter((s) => !s.isMobile).map((site) => {
    const stockDebut = stockBeforeDate(site, movements, bounds.start, inventaires);
    const rangeMovs = movementsInRange(movements, site.id, bounds.start, bounds.end);
    const reception = sumQty(rangeMovs, ["reception"]);
    const retourCamions = sumQty(rangeMovs, ["retour_camion"]);
    const sorties = sumQty(rangeMovs, ["sortie", "sortie_camion"]);
    const ajustement = rangeMovs.filter((m) => m.type === "ajustement").reduce((a, m) => a + m.delta, 0);
    const stockTheoriqueFin = stockDebut + reception + retourCamions - sorties + ajustement;
    const monthInv = inventaires.filter((i) => i.siteId === site.id && (i.product || "gasoil") === "gasoil" && i.date >= bounds.start && i.date <= bounds.end).sort((a, b) => (a.date < b.date ? 1 : -1));
    const dernierInv = pickLatestInv(monthInv);
    const stockFin = dernierInv ? dernierInv.stockPhysique : null;
    const ecart = stockFin !== null ? stockFin - stockTheoriqueFin : null;
    return { site, stockDebut, reception, retourCamions, sorties, stockTheoriqueFin, stockFin, ecart };
  });

  const doExcel = () => exportToExcel(`SOMIP_Synthese_Gasoil_${month}.xlsx`, [{
    name: "Synthèse Gasoil", rows: rows.map((r) => ({
      Site: r.site.name, "Stock début mois (L)": Math.round(r.stockDebut), "Réception (L)": Math.round(r.reception),
      "Retour camions (L)": Math.round(r.retourCamions), "Sorties (L)": Math.round(r.sorties),
      "Stock théorique fin (L)": Math.round(r.stockTheoriqueFin),
      "Stock fin mesuré (L)": r.stockFin !== null ? Math.round(r.stockFin) : "",
      "Écart (L)": r.ecart !== null ? Math.round(r.ecart) : "",
    })),
  }]);

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14 }}>
        <Field label="Mois de la synthèse"><input type="month" className="somip-input" style={{ maxWidth: 220 }} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="Synthèse Gasoil — Zone Sud-Est" period={`Mois de ${bounds.start} au ${bounds.end}`} />
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead>
              <tr>
                <th>Site</th>
                <th style={{ textAlign: "right" }}>Stock début mois</th>
                <th style={{ textAlign: "right" }}>Réception</th><th style={{ textAlign: "right" }}>Retour camions</th><th style={{ textAlign: "right" }}>Sorties</th>
                <th style={{ textAlign: "right" }}>Stock théorique fin</th><th style={{ textAlign: "right" }}>Stock fin mesuré</th>
                <th style={{ textAlign: "right" }}>Écart</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name} <span style={{ color: C.sub, fontWeight: 500 }}>({r.site.code})</span></td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(r.stockDebut)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.reception ? C.success : C.sub }}>{r.reception ? `+${fmt(r.reception)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.retourCamions ? C.success : C.sub }}>{r.retourCamions ? `+${fmt(r.retourCamions)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.sorties ? C.ink : C.sub, fontWeight: r.sorties ? 600 : 400 }}>{r.sorties ? `${fmt(r.sorties)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stockTheoriqueFin)} L</td>
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
          Stock fin mesuré = dernier inventaire gasoil enregistré pour ce site dans le mois.
        </p>
      </div>
    </div>
  );
}

/* ---- Synthèse Camions (stations mobiles) ---- */
function TruckSynthesisReport({ sites, movements, inventaires }) {
  const [month, setMonth] = useState(currentMonth());
  const bounds = monthBounds(month);
  const trucks = sites.filter((s) => s.isMobile);

  const rows = trucks.map((truck) => {
    const stockDebut = stockBeforeDate(truck, movements, bounds.start, inventaires);
    const rangeMovs = movementsInRange(movements, truck.id, bounds.start, bounds.end);
    const chargement = sumQty(rangeMovs, ["reception"]);
    const sortieTerrain = sumQty(rangeMovs, ["sortie"]);
    const retourCuve = sumQty(rangeMovs, ["retour_cuve_camion"]);
    const ajustement = rangeMovs.filter((m) => m.type === "ajustement").reduce((a, m) => a + m.delta, 0);
    const stockTheoriqueFin = stockDebut + chargement - sortieTerrain - retourCuve + ajustement;
    const monthInv = inventaires.filter((i) => i.siteId === truck.id && (i.product || "gasoil") === "gasoil" && i.date >= bounds.start && i.date <= bounds.end).sort((a, b) => (a.date < b.date ? 1 : -1));
    const dernierInv = pickLatestInv(monthInv);
    const stockFin = dernierInv ? dernierInv.stockPhysique : null;
    const ecart = stockFin !== null ? stockFin - stockTheoriqueFin : null;
    return { truck, stockDebut, chargement, sortieTerrain, retourCuve, stockTheoriqueFin, stockFin, ecart };
  });

  const doExcel = () => exportToExcel(`SOMIP_Synthese_Camions_${month}.xlsx`, [{
    name: "Synthèse Camions", rows: rows.map((r) => ({
      Camion: r.truck.name, "Stock début mois (L)": Math.round(r.stockDebut), "Chargement (L)": Math.round(r.chargement),
      "Sortie Fiche Terrain (L)": Math.round(r.sortieTerrain), "Retour Cuve (L)": Math.round(r.retourCuve),
      "Stock théorique fin (L)": Math.round(r.stockTheoriqueFin),
      "Stock fin mesuré (L)": r.stockFin !== null ? Math.round(r.stockFin) : "",
      "Écart (L)": r.ecart !== null ? Math.round(r.ecart) : "",
    })),
  }]);

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14 }}>
        <Field label="Mois de la synthèse"><input type="month" className="somip-input" style={{ maxWidth: 220 }} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="Synthèse Camions laitiers" period={`Mois de ${bounds.start} au ${bounds.end}`} />
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />
        {trucks.length === 0 ? (
          <p style={{ fontSize: 12.5, color: C.sub }}>Aucun camion enregistré. Ajoute-les sur la page "Sites" en cochant "Camion (station mobile)".</p>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table className="somip-table">
              <thead>
                <tr>
                  <th>Camion</th>
                  <th style={{ textAlign: "right" }}>Stock début mois</th>
                  <th style={{ textAlign: "right" }}>Chargement</th><th style={{ textAlign: "right" }}>Sortie Fiche Terrain</th><th style={{ textAlign: "right" }}>Retour Cuve</th>
                  <th style={{ textAlign: "right" }}>Stock théorique fin</th><th style={{ textAlign: "right" }}>Stock fin mesuré</th>
                  <th style={{ textAlign: "right" }}>Écart</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r) => (
                  <tr key={r.truck.id}>
                    <td style={{ fontWeight: 700, color: C.blue }}>{r.truck.name}</td>
                    <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(r.stockDebut)} L</td>
                    <td className="somip-mono" style={{ textAlign: "right", color: r.chargement ? C.success : C.sub }}>{r.chargement ? `+${fmt(r.chargement)} L` : "—"}</td>
                    <td className="somip-mono" style={{ textAlign: "right", color: r.sortieTerrain ? C.danger : C.sub }}>{r.sortieTerrain ? `−${fmt(r.sortieTerrain)} L` : "—"}</td>
                    <td className="somip-mono" style={{ textAlign: "right", color: r.retourCuve ? C.danger : C.sub }}>{r.retourCuve ? `−${fmt(r.retourCuve)} L` : "—"}</td>
                    <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stockTheoriqueFin)} L</td>
                    <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{r.stockFin !== null ? `${fmt(r.stockFin)} L` : "—"}</td>
                    <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600, color: r.ecart === null ? C.sub : r.ecart < 0 ? C.danger : r.ecart > 0 ? C.success : C.sub }}>
                      {r.ecart !== null ? `${r.ecart >= 0 ? "+" : ""}${fmt(r.ecart)} L` : "—"}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          Équation : Stock début + Chargement − Sortie Fiche Terrain − Retour Cuve = Stock théorique fin. Stock fin mesuré = dernier inventaire du mois pour ce camion.
        </p>
      </div>
    </div>
  );
}

/* ---- Synthèse Station (site fixe + camions rattachés, ex : Prehomo/Okouma) ---- */
function trucksAssignedAt(assignments, stationId, dateStr) {
  return assignments.filter((a) => a.stationId === stationId && a.startDate <= dateStr && (!a.endDate || a.endDate >= dateStr)).map((a) => a.truckId);
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
        const tSortieTerrain = sumQty(dayMovsTruck, ["sortie"]);
        const tRetourCuve = sumQty(dayMovsTruck, ["retour_cuve_camion"]);
        const tTheorique = tStockDebut + tChargement - tSortieTerrain - tRetourCuve;
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

      days.push({ date: d, stockDebutCombine, reception, ventesCombinees, chargementLaitiers, stockTheoriqueCombine, stockJaugeCombine, ecart, siteEcart, truckDetails, nbTrucks: truckIdsToday.length });
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
                  <td className="somip-mono" style={{ textAlign: "right", color: d.chargementLaitiers ? C.orange : C.sub }}>{d.chargementLaitiers ? `${fmt(d.chargementLaitiers)} L` : "—"}</td>
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

function StationSynthesisReport({ sites, movements, inventaires, truckAssignments }) {
  const stations = LUBRICANT_SITE_IDS.map((id) => sites.find((s) => s.id === id)).filter(Boolean);
  const [stationId, setStationId] = useState(stations[0]?.id || "");
  const [month, setMonth] = useState(currentMonth());
  const bounds = monthBounds(month);
  const station = sites.find((s) => s.id === stationId);

  const intervals = station ? truckIntervalsForStation(truckAssignments, stationId, bounds.start, bounds.end) : [];
  const truckIdsAtEnd = station ? trucksAssignedAt(truckAssignments, stationId, bounds.end) : [];

  const siteStockDebut = station ? stockBeforeDate(station, movements, bounds.start, inventaires) : 0;
  const siteRangeMovs = station ? movementsInRange(movements, station.id, bounds.start, bounds.end) : [];
  const siteReception = sumQty(siteRangeMovs, ["reception"]);
  const siteRetourCamions = sumQty(siteRangeMovs, ["retour_camion"]);
  const siteVentesDirectes = sumQty(siteRangeMovs, ["sortie"]);
  const siteSortieCamion = sumQty(siteRangeMovs, ["sortie_camion"]);
  const siteStockTheoriqueFin = station ? stockThroughDate(station, movements, bounds.end, inventaires) : 0;

  const truckRows = intervals.map((iv) => {
    const truck = sites.find((s) => s.id === iv.truckId);
    const intervalMovs = movements.filter((m) => m.siteId === iv.truckId && (m.product || "gasoil") === "gasoil" && m.date >= iv.start && m.date <= iv.end);
    const ventesTerrain = sumQty(intervalMovs, ["sortie"]);
    const chargement = sumQty(intervalMovs, ["reception"]);
    const retourCuve = sumQty(intervalMovs, ["retour_cuve_camion"]);
    const stillAssigned = truckIdsAtEnd.includes(iv.truckId);
    const stockFinTruck = stillAssigned && truck ? stockThroughDate(truck, movements, bounds.end, inventaires) : null;
    return { truck, iv, ventesTerrain, chargement, retourCuve, stillAssigned, stockFinTruck };
  });

  const ventesGlobales = siteVentesDirectes + truckRows.reduce((a, r) => a + r.ventesTerrain, 0);
  const stockTheoriqueCombine = siteStockTheoriqueFin + truckRows.filter((r) => r.stillAssigned).reduce((a, r) => a + (r.stockFinTruck || 0), 0);

  const monthInv = station ? inventaires.filter((i) => i.siteId === station.id && (i.product || "gasoil") === "gasoil" && i.date >= bounds.start && i.date <= bounds.end).sort((a, b) => (a.date < b.date ? 1 : -1)) : [];
  const dernierInvSite = pickLatestInv(monthInv);
  const stockFinMesureSite = dernierInvSite ? dernierInvSite.stockPhysique : null;

  const doExcel = () => exportToExcel(`SOMIP_Synthese_Station_${station?.code || stationId}_${month}.xlsx`, [
    {
      name: "Global", rows: [{
        Station: station?.name || "", "Stock début mois (L)": Math.round(siteStockDebut),
        "Réception site (L)": Math.round(siteReception), "Retour camions site (L)": Math.round(siteRetourCamions),
        "Ventes globales (site + camions terrain) (L)": Math.round(ventesGlobales),
        "Stock théorique combiné fin (L)": Math.round(stockTheoriqueCombine),
      }],
    },
    {
      name: "Détail camions", rows: truckRows.map((r) => ({
        Camion: r.truck?.name || r.iv.truckId, Du: r.iv.start, Au: r.iv.end,
        "Chargement (L)": Math.round(r.chargement), "Ventes terrain (L)": Math.round(r.ventesTerrain), "Retour cuve (L)": Math.round(r.retourCuve),
        "Toujours affecté fin de mois": r.stillAssigned ? "Oui" : "Non",
        "Stock camion fin de mois (L)": r.stockFinTruck !== null ? Math.round(r.stockFinTruck) : "",
      })),
    },
  ]);

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14, display: "flex", gap: 12, flexWrap: "wrap" }}>
        <Field label="Station">
          <select className="somip-select" style={{ maxWidth: 240 }} value={stationId} onChange={(e) => setStationId(e.target.value)}>
            {stations.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </Field>
        <Field label="Mois"><input type="month" className="somip-input" style={{ maxWidth: 220 }} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title={`Synthèse Station — ${station?.name || ""}`} period={`Mois de ${bounds.start} au ${bounds.end} (site + camions rattachés)`} />
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />

        <h4 style={{ margin: "0 0 10px", fontSize: 13 }}>Vue globale (consolidée)</h4>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 20 }}>
          <MiniStat label="Stock début (site)" value={`${fmt(siteStockDebut)} L`} />
          <MiniStat label="Réception (site)" value={`+${fmt(siteReception)} L`} color={C.success} />
          <MiniStat label="Ventes globales (site + terrain camions)" value={`${fmt(ventesGlobales)} L`} />
          <MiniStat label="Stock théorique combiné fin" value={`${fmt(stockTheoriqueCombine)} L`} bold />
          <MiniStat label="Stock fin mesuré (site, dernier inv.)" value={stockFinMesureSite !== null ? `${fmt(stockFinMesureSite)} L` : "—"} />
        </div>

        <h4 style={{ margin: "0 0 10px", fontSize: 13 }}>Détail par camion rattaché sur la période</h4>
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead><tr><th>Camion</th><th>Du</th><th>Au</th><th style={{ textAlign: "right" }}>Chargement</th><th style={{ textAlign: "right" }}>Ventes terrain</th><th style={{ textAlign: "right" }}>Retour cuve</th><th>Affecté fin de mois</th><th style={{ textAlign: "right" }}>Stock camion fin</th></tr></thead>
            <tbody>
              {truckRows.length === 0 && <EmptyRow colSpan={8} text="Aucun camion rattaché à cette station sur cette période." />}
              {truckRows.map((r, idx) => (
                <tr key={idx}>
                  <td style={{ fontWeight: 700, color: C.blue }}>{r.truck?.name || r.iv.truckId}</td>
                  <td className="somip-mono">{r.iv.start}</td>
                  <td className="somip-mono">{r.iv.end}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>{r.chargement ? `+${fmt(r.chargement)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.ventesTerrain ? C.ink : C.sub, fontWeight: r.ventesTerrain ? 600 : 400 }}>{r.ventesTerrain ? `${fmt(r.ventesTerrain)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.danger }}>{r.retourCuve ? `−${fmt(r.retourCuve)} L` : "—"}</td>
                  <td>{r.stillAssigned ? <Badge color={C.success}>Oui</Badge> : <Badge color={C.sub}>Non — réaffecté</Badge>}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600 }}>{r.stockFinTruck !== null ? `${fmt(r.stockFinTruck)} L` : "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          "Ventes globales" additionne les ventes directes du site et les sorties fiche terrain de chaque camion, uniquement pour ses jours de rattachement à cette station (les transferts internes site↔camion s'annulent automatiquement). "Stock théorique combiné fin" n'inclut que les camions encore rattachés à cette station à la fin du mois — un camion réaffecté ailleurs en cours de mois apparaît "Non" et son stock de fin appartient désormais à sa nouvelle station.
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
  const [periodType, setPeriodType] = useState("mensuel");
  const [monthKey, setMonthKey] = useState(currentMonth());
  const [year, setYear] = useState(new Date().getFullYear());
  const [quarter, setQuarter] = useState(Math.floor(new Date().getMonth() / 3) + 1);
  const [decadeMonth, setDecadeMonth] = useState(currentMonth());
  const [decadeNum, setDecadeNum] = useState(1);
  const periodKey = periodType === "mensuel" ? monthKey : periodType === "trimestriel" ? `${year}-Q${quarter}` : `${decadeMonth}-D${decadeNum}`;

  const emptyForm = { reception15: "", ventes15: "", transferts15: "", stockFin15: "", commentaire: "" };
  const [form, setForm] = useState(emptyForm);
  const [photoFiles, setPhotoFiles] = useState([]);
  const [existingPhotoUrls, setExistingPhotoUrls] = useState([]);

  // Saisie globale (tous sites confondus) : on ne retient que les entrées sans site_id.
  const sorted = [...bilans].filter((b) => !b.siteId && b.periodType === periodType).sort((a, b) => (a.periodKey < b.periodKey ? -1 : 1));
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
    if (form.stockFin15 === "") return;
    saveBilan({ periodType, periodKey, ...form, photoFiles, existingPhotoUrls });
    resetForm();
  };

  const loadForEdit = (b) => {
    setForm({ reception15: String(b.reception15), ventes15: String(b.ventes15), transferts15: String(b.transferts15), stockFin15: String(b.stockFin15), commentaire: b.commentaire || "" });
    setExistingPhotoUrls(b.photoUrls || []);
    setPhotoFiles([]);
  };

  // Historique enrichi pour le tableau + le diagramme.
  const history = sorted.map((b, i) => {
    const prev = i > 0 ? sorted[i - 1] : null;
    const debut = prev ? prev.stockFin15 : 0;
    const theorique = debut + b.reception15 - b.ventes15 + b.transferts15;
    const ec = b.stockFin15 - theorique;
    return { ...b, stockDebut: debut, stockTheorique: theorique, ecart: ec };
  });

  const doPptx = () => exportBilanToPptx(history, periodType, null);
  const doPdf = () => exportToPdf({
    filename: `SOMIP_Bilan_Matieres_${periodType}_${new Date().toISOString().slice(0, 10)}.pdf`,
    title: "Bilan Matières — Tous sites",
    period: PERIOD_TYPE_LABELS[periodType],
    columns: ["Période", "Stock début", "Réception", "Ventes", "Transferts", "Stock théorique", "Stock fin", "Gain/Perte"],
    rows: history.map((b) => [
      b.periodKey, `${fmt(b.stockDebut)} L`, `${fmt(b.reception15)} L`, `${fmt(b.ventes15)} L`,
      `${b.transferts15 >= 0 ? "+" : ""}${fmt(b.transferts15)} L`, `${fmt(b.stockTheorique)} L`, `${fmt(b.stockFin15)} L`,
      `${b.ecart >= 0 ? "+" : ""}${fmt(b.ecart)} L`,
    ]),
  });

  return (
    <div className="somip-fade" style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      {canManage && (
        <div className="somip-panel" style={{ flex: "1 1 320px", padding: 18 }}>
          <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Saisie du Bilan Matières</h3>
          <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.sub }}>Saisie globale — tous sites confondus.</p>
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
            <span style={{ fontSize: 12.5, color: C.sub, fontWeight: 600 }}>Stock début (période précédente)</span>
            <span className="somip-mono" style={{ fontWeight: 700 }}>{fmt(stockDebut)} L</span>
          </div>

          <Field label="Réception globale à 15°C (L)"><input type="number" className="somip-input" value={form.reception15} onChange={(e) => setForm({ ...form, reception15: e.target.value })} placeholder="0" /></Field>
          <Field label="Ventes globales à 15°C (L)"><input type="number" className="somip-input" value={form.ventes15} onChange={(e) => setForm({ ...form, ventes15: e.target.value })} placeholder="0" /></Field>
          <Field label="Transferts entre sites à 15°C (L, net)"><input type="number" className="somip-input" value={form.transferts15} onChange={(e) => setForm({ ...form, transferts15: e.target.value })} placeholder="0 (+ reçu, − envoyé)" /></Field>

          <div style={{ background: C.bg, borderRadius: 8, padding: "9px 12px", margin: "4px 0 12px", display: "flex", justifyContent: "space-between" }}>
            <span style={{ fontSize: 12.5, color: C.sub, fontWeight: 600 }}>Stock théorique (calculé)</span>
            <span className="somip-mono" style={{ fontWeight: 700 }}>{fmt(stockTheorique)} L</span>
          </div>

          <Field label="Stock fin mesuré à 15°C (L, obligatoire)"><input type="number" className="somip-input" value={form.stockFin15} onChange={(e) => setForm({ ...form, stockFin15: e.target.value })} placeholder="Jauge globale" /></Field>
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
          <h3 style={{ margin: 0, fontSize: 14 }}>Historique — Tous sites — {PERIOD_TYPE_LABELS[periodType]}</h3>
          {history.length > 0 && (
            <div style={{ display: "flex", gap: 8 }}>
              <button className="somip-btn somip-btn-primary" onClick={doPdf}><Download size={14} /> Export PDF</button>
              <button className="somip-btn somip-btn-primary" onClick={doPptx}><Download size={14} /> Export PowerPoint</button>
            </div>
          )}
        </div>
        <div style={{ overflowX: "auto", marginBottom: 20 }}>
          <table className="somip-table">
            <thead><tr><th>Période</th><th style={{ textAlign: "right" }}>Stock début</th><th style={{ textAlign: "right" }}>Réception</th><th style={{ textAlign: "right" }}>Ventes</th><th style={{ textAlign: "right" }}>Transferts</th><th style={{ textAlign: "right" }}>Stock théorique</th><th style={{ textAlign: "right" }}>Stock fin</th><th style={{ textAlign: "right" }}>Gain/Perte</th><th>Photos</th>{canManage && <th></th>}</tr></thead>
            <tbody>
              {history.length === 0 && <EmptyRow colSpan={canManage ? 10 : 9} text="Aucun Bilan Matières enregistré." />}
              {history.map((b) => (
                <tr key={b.id}>
                  <td style={{ fontWeight: 700 }}>{b.periodKey}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(b.stockDebut)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>{fmt(b.reception15)} L</td>
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
          <div style={{ height: 280 }}>
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
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          Stock théorique = Stock début (= Stock fin de la période précédente) + Réception − Ventes ± Transferts. Toutes les valeurs sont à saisir déjà corrigées à 15°C.
        </p>
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

function LossGainReport({ sites, inventaires }) {
  const [start, setStart] = useState("");
  const [end, setEnd] = useState("");
  const filtered = inventaires.filter((i) => (!start || i.date >= start) && (!end || i.date <= end));
  const rows = sites.map((s) => {
    const list = filtered.filter((i) => i.siteId === s.id);
    const pertes = list.filter((i) => i.nature === "perte");
    const gains = list.filter((i) => i.nature === "gain");
    const totalPertes = pertes.reduce((a, i) => a + Math.abs(i.ecart), 0);
    const totalGains = gains.reduce((a, i) => a + i.ecart, 0);
    const tauxMoyen = pertes.length ? pertes.reduce((a, i) => a + i.tauxFreinte, 0) / pertes.length : null;
    const horsObjectif = list.filter((i) => i.conformite === "non_conforme").length;
    return { site: s, nb: list.length, totalPertes, totalGains, tauxMoyen, horsObjectif };
  });

  const doExcel = () => exportToExcel(`SOMIP_PertesGains_${start || "debut"}_${end || "fin"}.xlsx`, [{
    name: "Pertes-Gains", rows: rows.map((r) => ({
      Site: r.site.name, "Nb inventaires": r.nb, "Total pertes (L)": Math.round(r.totalPertes), "Total gains (L)": Math.round(r.totalGains),
      "Freinte moyenne (‰)": r.tauxMoyen !== null ? r.tauxMoyen.toFixed(2) : "", "Inventaires hors objectif": r.horsObjectif,
    })),
  }]);

  return (
    <div>
      <div className="somip-no-print" style={{ display: "flex", gap: 12, marginBottom: 14, flexWrap: "wrap" }}>
        <Field label="Depuis (optionnel)"><input type="date" className="somip-input" value={start} onChange={(e) => setStart(e.target.value)} /></Field>
        <Field label="Jusqu'à (optionnel)"><input type="date" className="somip-input" value={end} onChange={(e) => setEnd(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="Pertes / gains par site" period={start || end ? `Période du ${start || "début"} au ${end || "aujourd'hui"}` : "Toutes les données disponibles"} />
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead><tr><th>Site</th><th style={{ textAlign: "right" }}>Nb inventaires</th><th style={{ textAlign: "right" }}>Total pertes</th><th style={{ textAlign: "right" }}>Total gains</th><th style={{ textAlign: "right" }}>Freinte moyenne</th><th style={{ textAlign: "right" }}>Hors objectif</th></tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{r.nb}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.danger, fontWeight: 600 }}>{r.totalPertes ? `−${fmt(r.totalPertes)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success, fontWeight: 600 }}>{r.totalGains ? `+${fmt(r.totalGains)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{r.tauxMoyen !== null ? `${r.tauxMoyen.toFixed(2)} ‰` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.horsObjectif ? C.danger : C.sub }}>{r.horsObjectif}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{ marginTop: 20, height: 240 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={rows.map((r) => ({ code: r.site.code, Pertes: r.totalPertes, Gains: r.totalGains }))} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F3" vertical={false} />
              <XAxis dataKey="code" tick={{ fontSize: 11, fill: C.sub }} axisLine={{ stroke: C.border }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: C.sub }} axisLine={false} tickLine={false} />
              <Tooltip formatter={(v) => `${fmt(v)} L`} contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${C.border}` }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="Pertes" fill={C.danger} radius={[4, 4, 0, 0]} />
              <Bar dataKey="Gains" fill={C.success} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
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
  const [colorPrimary, setColorPrimary] = useState(settings.colorPrimary || "#0071BD");
  const [colorAccent, setColorAccent] = useState(settings.colorAccent || "#F16B16");
  const [saving, setSaving] = useState(false);

  const previewLogo = logoFile ? URL.createObjectURL(logoFile) : settings.logoUrl;
  const dirty = !!logoFile || colorPrimary !== (settings.colorPrimary || "#0071BD") || colorAccent !== (settings.colorAccent || "#F16B16");

  const submit = async () => {
    setSaving(true);
    await updateTheme({ logoFile, colorPrimary, colorAccent });
    setSaving(false);
    setLogoFile(null);
    window.location.reload();
  };

  return (
    <div className="somip-fade" style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      <div className="somip-panel" style={{ flex: "1 1 380px", padding: 18 }}>
        <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Personnalisation</h3>
        <p style={{ margin: "0 0 18px", fontSize: 12.5, color: C.sub }}>Logo et couleurs principales, appliqués à toute l'application et aux rapports (PDF/PowerPoint).</p>

        <Field label="Logo">
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


function UsersView({ profiles, updateUserRole, updateUserSite, sites, session }) {
  const [editingId, setEditingId] = useState(null);
  const [roleDraft, setRoleDraft] = useState("");
  const [siteDraft, setSiteDraft] = useState("");
  const [deletingErr, setDeletingErr] = useState(null);
  const [form, setForm] = useState({ fullName: "", username: "", password: "", role: "lecture", assignedSiteId: "" });
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

  const startEdit = (u) => { setEditingId(u.id); setRoleDraft(u.role); setSiteDraft(u.assignedSiteId || ""); };
  const saveEdit = () => {
    updateUserRole(editingId, roleDraft);
    updateUserSite(editingId, siteDraft || null);
    setEditingId(null);
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

  const createAccount = async () => {
    setCreateErr(null); setCreateMsg(null);
    if (!form.fullName.trim() || !form.username.trim() || !form.password) { setCreateErr("Tous les champs sont requis."); return; }
    if (form.password.length < 6) { setCreateErr("Le mot de passe doit contenir au moins 6 caractères."); return; }
    setCreating(true);
    try {
      const res = await fetch("/api/create-user", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${session?.access_token || ""}` },
        body: JSON.stringify({ username: form.username.trim(), password: form.password, fullName: form.fullName.trim(), role: form.role, assignedSiteId: form.assignedSiteId || null }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Erreur lors de la création du compte.");
      setCreateMsg(`Compte créé pour ${form.fullName.trim()} (${ROLE_LABELS[form.role]}). Identifiant de connexion : "${data.loginEmail}" — communique-le avec le mot de passe.`);
      setForm({ fullName: "", username: "", password: "", role: "lecture", assignedSiteId: "" });
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
          <thead><tr><th>Nom</th><th>Rôle</th><th>Site assigné</th><th>Présence</th><th></th></tr></thead>
          <tbody>
            {profiles.length === 0 && <EmptyRow colSpan={5} text="Aucun compte pour le moment." />}
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
                        <select className="somip-select" value={siteDraft} onChange={(e) => setSiteDraft(e.target.value)}>
                          <option value="">Tous les sites</option>
                          {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                        </select>
                      </td>
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
                      <td style={{ color: C.sub, fontSize: 12.5 }}>{u.assignedSiteId ? (sites.find((s) => s.id === u.assignedSiteId)?.name || u.assignedSiteId) : "Tous les sites"}</td>
                      <td>
                        {isOnline(u) ? (
                          <span style={{ display: "inline-flex", alignItems: "center", gap: 5, fontSize: 12, color: C.success, fontWeight: 600 }}>
                            <span style={{ width: 7, height: 7, borderRadius: "50%", background: C.success, display: "inline-block" }} /> En ligne
                          </span>
                        ) : (
                          <span style={{ fontSize: 12, color: C.sub }}>{lastSeenLabel(u)}</span>
                        )}
                      </td>
                      <td style={{ whiteSpace: "nowrap", textAlign: "right" }}>
                        <button onClick={() => startEdit(u)} style={{ border: "none", background: "none", cursor: "pointer", padding: 5 }}><Pencil size={14} color={C.sub} /></button>
                        {u.id !== session?.user?.id && <ConfirmIconButton onConfirm={() => deleteAccount(u.id)} />}
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
        <h3 style={{ margin: "0 0 10px", fontSize: 14 }}>Créer un compte</h3>
        <Field label="Nom complet"><input className="somip-input" value={form.fullName} onChange={(e) => setForm({ ...form, fullName: e.target.value })} placeholder="Ex : Jean Mabiala" /></Field>
        <Field label="Nom d'utilisateur (identifiant de connexion)"><input className="somip-input" value={form.username} onChange={(e) => setForm({ ...form, username: e.target.value })} placeholder="Ex : jean.mabiala" /></Field>
        <Field label="Mot de passe provisoire"><input type="text" className="somip-input" value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} placeholder="Au moins 6 caractères" /></Field>
        <Field label="Rôle">
          <select className="somip-select" value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })}>
            {ROLE_VALUES.map((r) => <option key={r} value={r}>{ROLE_LABELS[r]}</option>)}
          </select>
        </Field>
        <Field label="Site assigné (optionnel)">
          <select className="somip-select" value={form.assignedSiteId} onChange={(e) => setForm({ ...form, assignedSiteId: e.target.value })}>
            <option value="">Tous les sites</option>
            {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
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

