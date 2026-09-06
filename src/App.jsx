import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  LayoutDashboard, Factory, ArrowDownCircle, ArrowUpCircle, ClipboardList,
  Truck, AlertTriangle, Plus, X, Trash2, Pencil, Fuel, RotateCcw, Check,
  Users, History, Loader2, CheckCircle2, AlertCircle, CloudOff, Thermometer,
  FileBarChart, Download, Printer, TrendingDown, TrendingUp, LogOut, Lock, Mail, Menu,
} from "lucide-react";
import * as XLSX from "xlsx";
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
const PRODUCTS = [{ id: "gasoil", label: "Gasoil" }, ...LUBRICANTS];

const MOVEMENTS_SEED = [
  { id: "m1", siteId: "okouma", type: "reception", date: "2026-08-28", quantity: 15000, delta: 15000, ref: "BL-2891", commentaire: "Livraison TotalEnergies", isDemo: true },
  { id: "m2", siteId: "okouma", type: "sortie_camion", date: "2026-08-30", quantity: 4200, delta: -4200, camion: "Camion Laitier 2", destination: "Carrière Nord — Engins", commentaire: "", isDemo: true },
  { id: "m3", siteId: "prehomo", type: "sortie", date: "2026-09-01", quantity: 2600, delta: -2600, destinataire: "Atelier mécanique", commentaire: "", isDemo: true },
  { id: "m4", siteId: "lipaka", type: "reception", date: "2026-09-02", quantity: 8000, delta: 8000, ref: "BL-2903", commentaire: "", isDemo: true },
  { id: "m5", siteId: "traction", type: "sortie", date: "2026-09-02", quantity: 1800, delta: -1800, destinataire: "Locomotive 12", commentaire: "", isDemo: true },
];

const SETTINGS_SEED = { objectifFreinte: 3 };

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
const fmt = (n) => Math.round(n).toLocaleString("fr-FR");
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
function stockBeforeDate(site, movements, dateExclusive) {
  return movements.filter((m) => m.siteId === site.id && (m.product || "gasoil") === "gasoil" && m.date < dateExclusive).reduce((a, m) => a + m.delta, site.stockInitial);
}
function stockThroughDate(site, movements, dateInclusive) {
  return movements.filter((m) => m.siteId === site.id && (m.product || "gasoil") === "gasoil" && m.date <= dateInclusive).reduce((a, m) => a + m.delta, site.stockInitial);
}
function stockBeforeDate15(site, movements, dateExclusive) {
  return movements.filter((m) => m.siteId === site.id && (m.product || "gasoil") === "gasoil" && m.date < dateExclusive).reduce((a, m) => a + Math.sign(m.delta) * movementQty15(m), site.stockInitial);
}
function stockBeforeDateProduct(stockInitial, movements, siteId, product, dateExclusive) {
  return movements.filter((m) => m.siteId === siteId && (m.product || "gasoil") === product && m.date < dateExclusive).reduce((a, m) => a + m.delta, stockInitial);
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

function ReportHeader({ title, period }) {
  return (
    <div className="somip-print-only" style={{ marginBottom: 16 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", borderBottom: `2px solid ${C.blue}`, paddingBottom: 10, marginBottom: 10 }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 15, color: C.navy }}>SOMIP — Stock Gasoil</div>
          <div style={{ fontSize: 11, color: C.sub }}>Zone Sud-Est · Gabon</div>
        </div>
        <div style={{ textAlign: "right", fontSize: 11, color: C.sub }}>
          Édité le {new Date().toLocaleDateString("fr-FR")} à {new Date().toLocaleTimeString("fr-FR")}
        </div>
      </div>
      <h2 style={{ margin: "0 0 2px", fontSize: 16 }}>{title}</h2>
      {period && <div style={{ fontSize: 12.5, color: C.sub }}>{period}</div>}
    </div>
  );
}

function ReportToolbar({ onExcel, onPrint }) {
  return (
    <div className="somip-no-print" style={{ display: "flex", gap: 8, marginBottom: 14 }}>
      <button className="somip-btn somip-btn-ghost" onClick={onExcel}><Download size={14} /> Export Excel</button>
      <button className="somip-btn somip-btn-ghost" onClick={onPrint}><Printer size={14} /> Export PDF (impression)</button>
    </div>
  );
}



/* ------------------------------------------------------------------ */
/* Rôles & permissions                                                  */
/* ------------------------------------------------------------------ */
const ROLE_VALUES = ["superviseur", "operateur", "chauffeur", "lecture"];
const ROLE_LABELS = { superviseur: "Superviseur", operateur: "Opérateur", chauffeur: "Chauffeur", lecture: "Lecture" };
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
  camion: r.camion || undefined, destination: r.destination || undefined, isDemo: !!r.is_demo,
  temperatureC: numOrUndef(r.temperature_c), densiteObservee: numOrUndef(r.densite_observee),
  densite15: numOrUndef(r.densite15), vcf: numOrUndef(r.vcf), volumeCorrige15: numOrUndef(r.volume_corrige15),
  indexAvant: numOrUndef(r.index_avant), indexApres: numOrUndef(r.index_apres),
  createdBy: r.created_by, createdAt: r.created_at,
});
const movementToRow = (m) => ({
  site_id: m.siteId, type: m.type, date: m.date, quantity: m.quantity, delta: m.delta,
  product: m.product || "gasoil",
  ref: m.ref ?? null, commentaire: m.commentaire ?? null, destinataire: m.destinataire ?? null,
  camion: m.camion ?? null, destination: m.destination ?? null, is_demo: !!m.isDemo,
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
  vcf: numOrUndef(r.vcf), stockPhysique15: numOrUndef(r.stock_physique15), createdBy: r.created_by, createdAt: r.created_at,
});
const inventaireToRow = (i) => ({
  site_id: i.siteId, date: i.date, stock_physique: i.stockPhysique, commentaire: i.commentaire ?? null,
  product: i.product || "gasoil",
  basis_ecart: i.basisEcart, stock_theorique_ambiant: i.stockTheoriqueAmbiant ?? null, stock_theorique15: i.stockTheorique15 ?? null,
  stock_theorique: i.stockTheorique ?? null, stock_physique_used: i.stockPhysiqueUsed ?? null, ecart: i.ecart, ecart_permille: i.ecartPermille,
  nature: i.nature, taux_freinte: i.tauxFreinte, objectif_utilise: i.objectifUtilise ?? null, conformite: i.conformite,
  adjustment_id: i.adjustmentId ?? null, temperature_c: i.temperatureC ?? null, densite_observee: i.densiteObservee ?? null,
  densite15: i.densite15 ?? null, vcf: i.vcf ?? null, stock_physique15: i.stockPhysique15 ?? null, created_by: i.createdBy ?? null,
});

const rowToProductStock = (r) => ({ id: r.id, siteId: r.site_id, product: r.product, capacity: Number(r.capacity), stockInitial: Number(r.stock_initial) });
const productStockToRow = (p) => ({ site_id: p.siteId, product: p.product, capacity: p.capacity, stock_initial: p.stockInitial });

const rowToAssignment = (r) => ({ id: r.id, truckId: r.truck_id, stationId: r.station_id, startDate: r.start_date, endDate: r.end_date || null });
const assignmentToRow = (a) => ({ truck_id: a.truckId, station_id: a.stationId, start_date: a.startDate, end_date: a.endDate ?? null });


const rowToAudit = (r) => ({ id: r.id, ts: r.ts, user: r.user_name, action: r.action, detail: r.detail });
const rowToProfile = (r) => ({ id: r.id, name: r.full_name, role: r.role });

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
          <div style={{ width: 34, height: 34, borderRadius: 8, background: C.blue, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Fuel size={18} color="#fff" />
          </div>
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
        <Field label="E-mail">
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
          const [sitesData, movementsData, inventairesData, profilesData, auditData, productStocksData, assignmentsData] = await Promise.all([
            fetchTable("sites", rowToSite),
            fetchTable("movements", rowToMovement, "date"),
            fetchTable("inventaires", rowToInventaire, "date"),
            fetchTable("profiles", rowToProfile),
            fetchTable("audit", rowToAudit, "ts", false),
            fetchTable("product_stocks", rowToProductStock),
            fetchTable("truck_assignments", rowToAssignment, "start_date"),
          ]);
          let settingsRow = null;
          try {
            const res = await supabase.from("settings").select("*").eq("id", 1).maybeSingle();
            settingsRow = res.data;
          } catch (e) { /* réglages optionnels : on garde la valeur par défaut si ça échoue */ }
          return { sitesData, movementsData, inventairesData, profilesData, auditData, productStocksData, assignmentsData, settingsRow };
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
        setSettings(result.settingsRow ? { objectifFreinte: Number(result.settingsRow.objectif_freinte) } : SETTINGS_SEED);
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
  useEffect(() => {
    if (loading || !session || !profile) return;
    const interval = setInterval(async () => {
      const [s, m, i, p, a, ps, ta] = await Promise.all([
        fetchTable("sites", rowToSite),
        fetchTable("movements", rowToMovement, "date"),
        fetchTable("inventaires", rowToInventaire, "date"),
        fetchTable("profiles", rowToProfile),
        fetchTable("audit", rowToAudit, "ts", false),
        fetchTable("product_stocks", rowToProductStock),
        fetchTable("truck_assignments", rowToAssignment, "start_date"),
      ]);
      setSites(s); setMovements(m); setInventaires(i); setProfiles(p); setAudit(a); setProductStocks(ps); setTruckAssignments(ta);
      const { data: se } = await supabase.from("settings").select("*").eq("id", 1).maybeSingle();
      if (se) setSettings({ objectifFreinte: Number(se.objectif_freinte) });
      setLastSync(new Date());
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

    // Miroir automatique site <-> camion, pour que les deux équations restent toujours synchronisées.
    const isTruckId = (id) => sites.some((s) => s.id === id && s.isMobile);
    if (payload.type === "sortie_camion" && isTruckId(payload.camion)) {
      const mirror = {
        id: uid(), siteId: payload.camion, product: "gasoil", type: "reception", date: payload.date,
        quantity: payload.quantity, delta: payload.quantity, ref: `Chargement automatique depuis ${sites.find((s) => s.id === payload.siteId)?.name || ""}`,
        temperatureC: payload.temperatureC, densiteObservee: payload.densiteObservee, densite15: payload.densite15, vcf: payload.vcf, volumeCorrige15: payload.volumeCorrige15,
        createdBy: currentUserName, createdAt: new Date().toISOString(), isDemo: false,
      };
      const { data: dm, error: em } = await supabase.from("movements").insert(movementToRow(mirror)).select().maybeSingle();
      if (!em) setMovements((prev) => [...prev, dm ? rowToMovement(dm) : mirror]);
    }
    if (payload.type === "retour_camion" && isTruckId(payload.camion)) {
      const mirror = {
        id: uid(), siteId: payload.camion, product: "gasoil", type: "retour_cuve_camion", date: payload.date,
        quantity: payload.quantity, delta: -payload.quantity, destination: payload.siteId,
        temperatureC: payload.temperatureC, densiteObservee: payload.densiteObservee, densite15: payload.densite15, vcf: payload.vcf, volumeCorrige15: payload.volumeCorrige15,
        createdBy: currentUserName, createdAt: new Date().toISOString(), isDemo: false,
      };
      const { data: dm, error: em } = await supabase.from("movements").insert(movementToRow(mirror)).select().maybeSingle();
      if (!em) setMovements((prev) => [...prev, dm ? rowToMovement(dm) : mirror]);
    }
    if (payload.type === "retour_cuve_camion" && sites.some((s) => s.id === payload.destination)) {
      const mirror = {
        id: uid(), siteId: payload.destination, product: "gasoil", type: "retour_camion", date: payload.date,
        quantity: payload.quantity, delta: payload.quantity, camion: payload.siteId, destination: `Retour automatique depuis ${sites.find((s) => s.id === payload.siteId)?.name || ""}`,
        temperatureC: payload.temperatureC, densiteObservee: payload.densiteObservee, densite15: payload.densite15, vcf: payload.vcf, volumeCorrige15: payload.volumeCorrige15,
        createdBy: currentUserName, createdAt: new Date().toISOString(), isDemo: false,
      };
      const { data: dm, error: em } = await supabase.from("movements").insert(movementToRow(mirror)).select().maybeSingle();
      if (!em) setMovements((prev) => [...prev, dm ? rowToMovement(dm) : mirror]);
    }
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
  const addInventaire = ({ siteId, product = "gasoil", date, stockPhysique, commentaire, temperatureC, densiteObservee, densite15, vcf, stockPhysique15 }) => withSync(async () => {
    // Le Gain/Perte "officiel" est toujours le résultat de l'équation de stock en base ambiante
    // (Stock physique mesuré − Stock théorique). La base 15°C est fournie à titre indicatif
    // (page "État des stocks — 15°C"), mais ne doit jamais se substituer à ce résultat, pour
    // éviter de confondre un écart réel avec un simple effet de conversion ambiant/15°C.
    const theoriqueAmbiant = stockOf(siteId, product);
    const theorique15 = stockOf15(siteId, product);
    const basisEcart = "ambiant";
    const theoriqueUsed = theoriqueAmbiant;
    const physiqueUsed = stockPhysique;
    const ecart = physiqueUsed - theoriqueUsed;
    const cls = classifyEcart(ecart, theoriqueUsed, settings.objectifFreinte);
    const has15 = stockPhysique15 !== undefined;
    const vcfFields = has15 ? { temperatureC, densiteObservee, densite15, vcf, stockPhysique15 } : {};
    const adjMovementDraft = {
      siteId, product, type: "ajustement", date, quantity: Math.abs(ecart), delta: ecart, isDemo: false,
      commentaire: `Ajustement suite à l'inventaire du ${date} (base ambiante)`,
      createdBy: currentUserName, createdAt: new Date().toISOString(),
    };
    const { data: dataM, error: e1 } = await supabase.from("movements").insert(movementToRow(adjMovementDraft)).select().maybeSingle();
    if (e1) throw e1;
    if (!dataM) throw new Error("L'ajustement n'a pas pu être confirmé par le serveur — réessaie.");
    const adjMovement = rowToMovement(dataM);
    const invDraft = {
      siteId, product, date, stockPhysique, commentaire, basisEcart,
      stockTheoriqueAmbiant: theoriqueAmbiant, stockTheorique15: theorique15,
      stockTheorique: theoriqueUsed, stockPhysiqueUsed: physiqueUsed,
      ecart: cls.ecartL, ecartPermille: cls.ecartPermille, nature: cls.nature,
      tauxFreinte: cls.tauxFreinte, objectifUtilise: cls.objectif, conformite: cls.conformite,
      adjustmentId: adjMovement.id, createdBy: currentUserName, createdAt: new Date().toISOString(), ...vcfFields,
    };
    const { data: dataI, error: e2 } = await supabase.from("inventaires").insert(inventaireToRow(invDraft)).select().maybeSingle();
    if (e2) {
      // Annule l'ajustement déjà inséré pour ne pas laisser un mouvement orphelin sans inventaire associé.
      await supabase.from("movements").delete().eq("id", adjMovement.id);
      throw e2;
    }
    if (!dataI) {
      await supabase.from("movements").delete().eq("id", adjMovement.id);
      throw new Error("L'inventaire n'a pas pu être confirmé par le serveur — réessaie.");
    }
    const invRecord = rowToInventaire(dataI);
    setMovements((prev) => [...prev, adjMovement]);
    setInventaires((prev) => [...prev, invRecord]);
    appendAudit("Inventaire", `${sites.find((s) => s.id === siteId)?.name || ""} — base ${has15 ? "15°C" : "ambiante"} — ${NATURE_META[cls.nature].label} ${ecart >= 0 ? "+" : ""}${fmt(ecart)} L (${cls.ecartPermille >= 0 ? "+" : ""}${cls.ecartPermille.toFixed(2)} ‰)`);
    flash("Inventaire enregistré et stock ajusté.");
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
    const { error } = await supabase.from("settings").update({ objectif_freinte: next.objectifFreinte }).eq("id", 1);
    if (error) throw error;
    setSettings(next);
    appendAudit("Modification objectif de freinte", `Nouvel objectif : ${next.objectifFreinte} ‰`);
    flash("Objectif mis à jour.");
  });

  /* ---- mutations : rôle d'un utilisateur (Superviseur uniquement) ---- */
  const updateUserRole = (userId, role) => withSync(async () => {
    const { error } = await supabase.from("profiles").update({ role }).eq("id", userId);
    if (error) throw error;
    const target = profiles.find((u) => u.id === userId);
    setProfiles((prev) => prev.map((u) => (u.id === userId ? { ...u, role } : u)));
    appendAudit("Modification rôle utilisateur", `${target?.name || ""} → ${ROLE_LABELS[role]}`);
    flash("Rôle mis à jour.");
  });

  const NAV = [
    { id: "dashboard", label: "Tableau de bord", icon: LayoutDashboard, show: true },
    { id: "sites", label: "Sites", icon: Factory, show: perms.canManage },
    { id: "saisie", label: "Saisie journalière", icon: ClipboardList, show: true },
    { id: "inventaires", label: "Inventaires", icon: ClipboardList, show: true },
    { id: "vcf", label: "Correction 15°C", icon: Thermometer, show: true },
    { id: "rapports", label: "Rapports", icon: FileBarChart, show: true },
    { id: "utilisateurs", label: "Utilisateurs", icon: Users, show: perms.canManage },
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
          <div style={{ width: 32, height: 32, borderRadius: 8, background: C.blue, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Fuel size={17} color="#fff" />
          </div>
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
          {view === "sites" && perms.canManage && <SitesView sites={sites} movements={movements} stockOf={stockOf} addSite={addSite} editSite={editSite} removeSite={removeSite} productStocks={productStocks} saveProductStock={saveProductStock} truckAssignments={truckAssignments} assignTruck={assignTruck} />}
          {view === "saisie" && <DailyEntryView sites={sites} movements={movements} inventaires={inventaires} productStocks={productStocks} saveProductStock={saveProductStock} addMovement={addMovement} addInventaire={addInventaire} deleteMovement={deleteMovement} settings={settings} canWrite={perms.canWrite} canManage={perms.canManage} />}
          {view === "inventaires" && <InventairesView sites={sites} inventaires={inventaires} stockOf={stockOf} stockOf15={stockOf15} addInventaire={addInventaire} deleteInventaire={deleteInventaire} settings={settings} updateSettings={updateSettings} canWrite={perms.canWrite} canManage={perms.canManage} />}
          {view === "vcf" && <VcfView />}
          {view === "rapports" && <ReportsView sites={sites} movements={movements} inventaires={inventaires} productStocks={productStocks} truckAssignments={truckAssignments} settings={settings} stockOf={stockOf} />}
          {view === "utilisateurs" && perms.canManage && <UsersView profiles={profiles} updateUserRole={updateUserRole} session={session} />}
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
function SitesView({ sites, movements, stockOf, addSite, editSite, removeSite, productStocks, saveProductStock, truckAssignments, assignTruck }) {
  const [form, setForm] = useState({ name: "", code: "", capacity: "", stockInitial: "", isMobile: false });
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [lubSiteId, setLubSiteId] = useState(LUBRICANT_SITE_IDS[0]);
  const [lubProduct, setLubProduct] = useState(LUBRICANTS[0].id);
  const [lubForm, setLubForm] = useState({ capacity: "", stockInitial: "" });
  const trucks = sites.filter((s) => s.isMobile);
  const stations = sites.filter((s) => !s.isMobile);
  const [assignForm, setAssignForm] = useState({ truckId: trucks[0]?.id || "", stationId: stations[0]?.id || "", startDate: todayStr() });

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
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Réceptions                                                            */
/* ------------------------------------------------------------------ */
/* ------------------------------------------------------------------ */
/* Saisie journalière (écran unique : réception, sortie/camion, retour) */
/* ------------------------------------------------------------------ */
function DailyEntryView({ sites, movements, inventaires, productStocks, addMovement, addInventaire, deleteMovement, settings, canWrite, canManage }) {
  const [siteId, setSiteId] = useState(sites[0]?.id || "");
  const [product, setProduct] = useState("gasoil");
  const [date, setDate] = useState(todayStr());
  const [receptionQty, setReceptionQty] = useState("");
  const [receptionRef, setReceptionRef] = useState("");
  const [indexAvant, setIndexAvant] = useState("");
  const [indexApres, setIndexApres] = useState("");
  const [sortieMode, setSortieMode] = useState("vente"); // "vente" | "camion"
  const [destinataire, setDestinataire] = useState("");
  const [camion, setCamion] = useState("");
  const [destination, setDestination] = useState("");
  const [retourQty, setRetourQty] = useState("");
  const [retourNote, setRetourNote] = useState("");
  const [retourCamionTruckId, setRetourCamionTruckId] = useState("");
  const [retourCuveTruckQty, setRetourCuveTruckQty] = useState("");
  const [retourCuveTruckNote, setRetourCuveTruckNote] = useState("");
  const [tempC, setTempC] = useState("");
  const [densite, setDensite] = useState("");
  const [stockFinMesure, setStockFinMesure] = useState("");
  const [commentaireInv, setCommentaireInv] = useState("");
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
  useEffect(() => { if (!isLubSite) setSortieMode("vente"); }, [siteId, isLubSite]);
  useEffect(() => { if (truckSites[0] && !camion) setCamion(truckSites[0].id); }, [truckSites, camion]);

  const site = sites.find((s) => s.id === siteId);
  const productStockEntry = productStocks.find((p) => p.siteId === siteId && p.product === product);
  const stockDebut = isLub
    ? stockBeforeDateProduct(productStockEntry?.stockInitial || 0, movements, siteId, product, date)
    : (site ? stockBeforeDate(site, movements, date) : 0);
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
  const sortieQty = indexAvant !== "" && indexApres !== "" ? Number(indexApres) - Number(indexAvant) : 0;
  const sortieValid = indexAvant === "" && indexApres === "" ? true : (indexAvant !== "" && indexApres !== "" && sortieQty > 0);
  const stockTheoriqueAmbiant = stockDebutEffective + receptionN + retourN - sortieQty - retourCuveTruckN;
  const stockTheorique15 = !skipVcf && site ? stockBeforeDate15(site, movements, date) : 0;

  // Chaque produit (gasoil, chaque lubrifiant) et chaque camion a son propre compteur de sortie.
  const lastIndexForSite = movements
    .filter((m) => m.siteId === siteId && (m.product || "gasoil") === product && ((isLub || isMobileSite) ? m.type === "sortie" : (m.type === "sortie" || m.type === "sortie_camion" || m.type === "retour_camion")) && m.indexApres !== undefined)
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

  const canSubmit = sortieValid && stockFinMesure !== "" && (!isFirstOfMonth || stockDebutConfirm !== "") && !existingInv;

  const resetDayFields = () => {
    setReceptionQty(""); setReceptionRef("");
    setIndexAvant(""); setIndexApres(""); setDestinataire(""); setDestination("");
    setRetourQty(""); setRetourNote(""); setRetourCamionTruckId(""); setRetourCuveTruckQty(""); setRetourCuveTruckNote(""); setTempC(""); setDensite("");
    setStockFinMesure(""); setCommentaireInv("");
  };

  const [submitting, setSubmitting] = useState(false);

  const submit = async () => {
    if (!siteId || !date || !canSubmit || submitting) return;
    setSubmitting(true);
    try {
      if (isFirstOfMonth) {
        const diff = Number(stockDebutConfirm) - stockDebut;
        if (diff !== 0) {
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
        let ok;
        if (isLub || isMobileSite) {
          ok = await addMovement({ siteId, product, type: "sortie", date, quantity: sortieQty, delta: -sortieQty, indexAvant: Number(indexAvant), indexApres: Number(indexApres), destinataire });
        } else {
          const base = { siteId, product, type: sortieMode === "camion" ? "sortie_camion" : "sortie", date, quantity: sortieQty, delta: -sortieQty, indexAvant: Number(indexAvant), indexApres: Number(indexApres), ...vcfExtra(sortieQty) };
          ok = await addMovement(sortieMode === "camion" ? { ...base, camion, destination } : { ...base, destinataire });
        }
        if (!ok) return;
      }
      if (retourN > 0) {
        const ok = await addMovement({ siteId, product, type: "retour_camion", date, quantity: retourN, delta: retourN, camion: retourCamionTruckId || undefined, destination: retourNote, ...vcfExtra(retourN) });
        if (!ok) return;
      }
      if (retourCuveTruckN > 0) {
        const ok = await addMovement({ siteId, product, type: "retour_cuve_camion", date, quantity: retourCuveTruckN, delta: -retourCuveTruckN, destination: retourCuveTruckNote, ...vcfExtra(retourCuveTruckN) });
        if (!ok) return;
      }
      const invExtra = vcfFin
        ? { temperatureC: Number(tempC), densiteObservee: Number(densite), densite15: vcfFin.densite15, vcf: vcfFin.vcf, stockPhysique15: vcfFin.volume15 }
        : {};
      const okInv = await addInventaire({ siteId, product, date, stockPhysique: stockFinN, commentaire: commentaireInv, ...invExtra });
      if (!okInv) return;
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
                <select className="somip-select" value={siteId} onChange={(e) => setSiteId(e.target.value)}>
                  {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
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
            <div style={{ flex: 1 }}><Field label="N° Bon de livraison"><input className="somip-input" value={receptionRef} onChange={(e) => setReceptionRef(e.target.value)} placeholder="BL-XXXX" /></Field></div>
          </div>
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
                  Dernier index enregistré pour {isMobileSite ? "ce camion" : "ce produit"} : {fmt(lastIndexForSite)}{indexMismatch && " — vérifie ton index avant."}
                </p>
              )}
              {!sortieValid && <p style={{ margin: "-6px 0 10px", fontSize: 11.5, color: C.danger }}>L'index après doit être supérieur à l'index avant.</p>}
              {isLub && sortieQty > 0 && <p style={{ margin: "-6px 0 10px", fontSize: 11, color: C.sub }}>≈ {fmt(sortieQty * lubDensite)} kg</p>}
            </>
          ) : isLubSite ? (
            <>
              <p style={{ margin: "10px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Sortie (compteur)</p>
              <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                <button className={`somip-tab ${sortieMode === "vente" ? "active" : ""}`} style={{ flex: 1, textAlign: "center" }} onClick={() => setSortieMode("vente")}>Vente</button>
                <button className={`somip-tab ${sortieMode === "camion" ? "active" : ""}`} style={{ flex: 1, textAlign: "center" }} onClick={() => setSortieMode("camion")}>Vers camion</button>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <div style={{ flex: 1 }}><Field label="Index avant"><input type="number" className="somip-input" value={indexAvant} onChange={(e) => setIndexAvant(e.target.value)} placeholder="Ex : 45210" /></Field></div>
                <div style={{ flex: 1 }}><Field label="Index après"><input type="number" className="somip-input" value={indexApres} onChange={(e) => setIndexApres(e.target.value)} placeholder="Ex : 47210" /></Field></div>
              </div>
              {lastIndexForSite !== undefined && (
                <p style={{ margin: "-6px 0 8px", fontSize: 11, color: indexMismatch ? C.warning : C.sub }}>
                  Dernier index enregistré sur ce site : {fmt(lastIndexForSite)}{indexMismatch && " — vérifie ton index avant."}
                </p>
              )}
              {sortieMode === "camion" && (
                <div style={{ display: "flex", gap: 8 }}>
                  <div style={{ flex: 1 }}>
                    <Field label="Camion">
                      <select className="somip-select" value={camion} onChange={(e) => setCamion(e.target.value)}>
                        {truckSites.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                      </select>
                    </Field>
                  </div>
                </div>
              )}
              {!sortieValid && <p style={{ margin: "-6px 0 10px", fontSize: 11.5, color: C.danger }}>L'index après doit être supérieur à l'index avant.</p>}
            </>
          ) : (
            <>
              <p style={{ margin: "10px 0 6px", fontSize: 12, fontWeight: 700, color: C.ink }}>Sortie (compteur)</p>
              <div style={{ display: "flex", gap: 8 }}>
                <div style={{ flex: 1 }}><Field label="Index avant"><input type="number" className="somip-input" value={indexAvant} onChange={(e) => setIndexAvant(e.target.value)} placeholder="Ex : 45210" /></Field></div>
                <div style={{ flex: 1 }}><Field label="Index après"><input type="number" className="somip-input" value={indexApres} onChange={(e) => setIndexApres(e.target.value)} placeholder="Ex : 47210" /></Field></div>
              </div>
              {lastIndexForSite !== undefined && (
                <p style={{ margin: "-6px 0 8px", fontSize: 11, color: indexMismatch ? C.warning : C.sub }}>
                  Dernier index enregistré sur ce site : {fmt(lastIndexForSite)}{indexMismatch && " — vérifie ton index avant."}
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
                  Créera automatiquement un "Retour Cuve" de {fmt(Number(retourQty) || 0)} L côté {truckSites.find((t) => t.id === retourCamionTruckId)?.name}.
                </p>
              )}
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
                  Créera automatiquement un "Retour camion (cuve)" de {fmt(retourCuveTruckN)} L côté {stationSites.find((s) => s.id === retourCuveTruckNote)?.name}.
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
            {submitting ? <Loader2 size={15} style={{ animation: "somipSpin .8s linear infinite" }} /> : <Plus size={15} />} {submitting ? "Enregistrement..." : "Enregistrer la journée"}
          </button>
          {!canSubmit && stockFinMesure === "" && <p style={{ margin: "8px 0 0", fontSize: 11.5, color: C.sub }}>Le Stock fin (jauge) est obligatoire pour enregistrer.</p>}
          {!canSubmit && stockFinMesure !== "" && isFirstOfMonth && stockDebutConfirm === "" && <p style={{ margin: "8px 0 0", fontSize: 11.5, color: C.sub }}>Le Stock début du mois est obligatoire pour enregistrer.</p>}
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
                <td style={{ color: C.sub }}>Jauge mesurée{existingInv.commentaire ? ` — ${existingInv.commentaire}` : ""}</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600 }}>{fmt(existingInv.stockPhysique)} L</td>
                {canManage && <td></td>}
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
function ReportsView({ sites, movements, inventaires, productStocks, truckAssignments, settings, stockOf }) {
  const [tab, setTab] = useState("journalier");
  const TABS = [
    { id: "journalier", label: "Journalier" },
    { id: "decadaire", label: "Décadaire" },
    { id: "mensuel", label: "Mensuel" },
    { id: "synthese_gasoil", label: "Synthèse Gasoil" },
    { id: "synthese_camions", label: "Synthèse Camions" },
    { id: "synthese_station", label: "Synthèse Station (Prehomo/Okouma)" },
    { id: "etat_ambiant", label: "État des stocks — Ambiant" },
    { id: "etat_15", label: "État des stocks — 15°C" },
    { id: "exposition", label: "Exposition" },
    { id: "pertesgains", label: "Pertes/Gains par site" },
    { id: "lubrifiants", label: "Synthèse Lubrifiants" },
    { id: "bons", label: "Bons de livraison" },
  ];
  return (
    <div className="somip-fade">
      <div className="somip-no-print" style={{ display: "flex", gap: 8, marginBottom: 18, flexWrap: "wrap" }}>
        {TABS.map((t) => (
          <button key={t.id} className={`somip-tab ${tab === t.id ? "active" : ""}`} onClick={() => setTab(t.id)}>{t.label}</button>
        ))}
      </div>
      {tab === "journalier" && <DailyReport sites={sites} movements={movements} inventaires={inventaires} />}
      {tab === "decadaire" && <DecadeReport sites={sites} movements={movements} inventaires={inventaires} />}
      {tab === "mensuel" && <MonthlyReport sites={sites} movements={movements} inventaires={inventaires} settings={settings} />}
      {tab === "synthese_gasoil" && <GasoilSynthesisReport sites={sites} movements={movements} inventaires={inventaires} />}
      {tab === "synthese_camions" && <TruckSynthesisReport sites={sites} movements={movements} inventaires={inventaires} />}
      {tab === "synthese_station" && <StationSynthesisReport sites={sites} movements={movements} inventaires={inventaires} truckAssignments={truckAssignments} />}
      {tab === "etat_ambiant" && <StockStatementAmbiant sites={sites} movements={movements} inventaires={inventaires} />}
      {tab === "etat_15" && <StockStatement15 sites={sites} movements={movements} inventaires={inventaires} />}
      {tab === "exposition" && <ExposureReport sites={sites} movements={movements} inventaires={inventaires} stockOf={stockOf} />}
      {tab === "pertesgains" && <LossGainReport sites={sites} inventaires={inventaires} />}
      {tab === "synthese_gasoil" && <GasoilSynthesisReport sites={sites} movements={movements} inventaires={inventaires} />}
      {tab === "lubrifiants" && <LubricantSynthesisReport sites={sites} movements={movements} inventaires={inventaires} productStocks={productStocks} />}
      {tab === "bons" && <DeliveryNotesReport sites={sites} movements={movements} />}
    </div>
  );
}

/* ---- Rapport journalier ---- */
function DailyReport({ sites, movements, inventaires }) {
  const [date, setDate] = useState(todayStr());
  const rows = sites.map((s) => {
    const stockDebut = stockBeforeDate(s, movements, date);
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
                  <td className="somip-mono" style={{ textAlign: "right", color: C.danger }}>{r.sorties ? `−${fmt(r.sorties)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stockFin)} L</td>
                  <td>{r.inv ? <Badge color={NATURE_META[r.inv.nature].color}>{NATURE_META[r.inv.nature].label} {r.inv.ecart >= 0 ? "+" : ""}{fmt(r.inv.ecart)} L</Badge> : <span style={{ color: C.sub }}>—</span>}</td>
                </tr>
              ))}
              <tr>
                <td style={{ fontWeight: 700 }}>Total réseau</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(totals.stockDebut)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.success }}>+{fmt(totals.receptions)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.success }}>+{fmt(totals.retours)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.danger }}>−{fmt(totals.sorties)} L</td>
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
    const stockDebut = stockBeforeDate(s, movements, bounds.start);
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
    const stockDebut = stockBeforeDate(s, movements, bounds.start);
    const rangeMovs = movementsInRange(movements, s.id, bounds.start, bounds.end);
    const receptions = sumQty(rangeMovs, ["reception"]);
    const retours = sumQty(rangeMovs, ["retour_camion"]);
    const sorties = sumQty(rangeMovs, ["sortie", "sortie_camion"]);
    const stockFin = stockThroughDate(s, movements, bounds.end);
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
    const stockDebut = stockBeforeDate(s, movements, date);
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
    const stockDebut = stockBeforeDate15(s, movements, date);
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
function ExposureReport({ sites, movements, inventaires, stockOf }) {
  const month = currentMonth();
  const rows = sites.map((s) => {
    const stock = stockOf(s.id);
    const pct = s.capacity ? (stock / s.capacity) * 100 : 0;
    const status = pct < 20 ? "danger" : pct < 35 ? "warning" : "ok";
    const lastInv = [...inventaires].filter((i) => i.siteId === s.id).sort((a, b) => (a.date < b.date ? 1 : -1))[0];
    return { site: s, stock, pct, status, lastInv };
  });
  const totalStock = rows.reduce((a, r) => a + r.stock, 0);
  const totalCap = rows.reduce((a, r) => a + r.site.capacity, 0);
  const alertSites = rows.filter((r) => r.status !== "ok");
  const nonConformes = rows.filter((r) => r.lastInv?.conformite === "non_conforme");
  const receptionsMonth = movements.filter((m) => m.type === "reception" && (m.product || "gasoil") === "gasoil" && m.date.startsWith(month)).reduce((a, m) => a + m.quantity, 0);
  const sortiesMonth = movements.filter((m) => (m.type === "sortie" || m.type === "sortie_camion") && (m.product || "gasoil") === "gasoil" && m.date.startsWith(month)).reduce((a, m) => a + m.quantity, 0);
  const statusColor = { ok: C.blue, warning: C.warning, danger: C.danger };

  const doExcel = () => exportToExcel(`SOMIP_Exposition_${todayStr()}.xlsx`, [{
    name: "Exposition", rows: rows.map((r) => ({
      Site: r.site.name, "Stock (L)": Math.round(r.stock), "Capacité (L)": r.site.capacity, "Remplissage (%)": r.pct.toFixed(1),
      "Dernier inventaire": r.lastInv ? r.lastInv.date : "", "Statut freinte": r.lastInv ? CONFORMITE_META[r.lastInv.conformite].label : "",
    })),
  }]);

  return (
    <div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="Exposition — situation consolidée du réseau" period={`Situation instantanée au ${todayStr()}`} />
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />

        <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 20 }}>
          <StatCard label="Stock total réseau" value={fmt(totalStock)} unit="L" accent={C.blue} icon={Fuel} />
          <StatCard label="Taux de remplissage réseau" value={totalCap ? ((totalStock / totalCap) * 100).toFixed(1) : 0} unit="%" accent={C.sub} icon={Factory} />
          <StatCard label="Sites en alerte stock" value={alertSites.length} unit={`/ ${rows.length}`} accent={C.danger} icon={AlertTriangle} />
          <StatCard label="Sites hors objectif freinte" value={nonConformes.length} unit={`/ ${rows.length}`} accent={C.warning} icon={TrendingDown} />
          <StatCard label="Réceptions (mois en cours)" value={fmt(receptionsMonth)} unit="L" accent={C.success} icon={ArrowDownCircle} />
          <StatCard label="Sorties (mois en cours)" value={fmt(sortiesMonth)} unit="L" accent={C.orange} icon={ArrowUpCircle} />
        </div>

        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead><tr><th>Site</th><th style={{ textAlign: "right" }}>Stock</th><th style={{ textAlign: "right" }}>Remplissage</th><th>Statut stock</th><th>Dernier inventaire</th><th>Statut freinte</th></tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(r.stock)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{r.pct.toFixed(1)} %</td>
                  <td><Badge color={statusColor[r.status]}>{r.status === "ok" ? "Normal" : r.status === "warning" ? "Vigilance" : "Critique"}</Badge></td>
                  <td style={{ color: C.sub }}>{r.lastInv ? r.lastInv.date : "Aucun"}</td>
                  <td>{r.lastInv ? <Badge color={CONFORMITE_META[r.lastInv.conformite].color}>{CONFORMITE_META[r.lastInv.conformite].label}</Badge> : <span style={{ color: C.sub }}>—</span>}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ---- Pertes / Gains par site ---- */
/* ---- Synthèse Gasoil (les 8 sites) ---- */
function GasoilSynthesisReport({ sites, movements, inventaires }) {
  const [month, setMonth] = useState(currentMonth());
  const bounds = monthBounds(month);

  const rows = sites.filter((s) => !s.isMobile).map((site) => {
    const stockDebut = stockBeforeDate(site, movements, bounds.start);
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
                  <td className="somip-mono" style={{ textAlign: "right", color: r.sorties ? C.danger : C.sub }}>{r.sorties ? `−${fmt(r.sorties)} L` : "—"}</td>
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
    const stockDebut = stockBeforeDate(truck, movements, bounds.start);
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

function StationSynthesisReport({ sites, movements, inventaires, truckAssignments }) {
  const stations = LUBRICANT_SITE_IDS.map((id) => sites.find((s) => s.id === id)).filter(Boolean);
  const [stationId, setStationId] = useState(stations[0]?.id || "");
  const [month, setMonth] = useState(currentMonth());
  const bounds = monthBounds(month);
  const station = sites.find((s) => s.id === stationId);

  const intervals = station ? truckIntervalsForStation(truckAssignments, stationId, bounds.start, bounds.end) : [];
  const truckIdsAtEnd = station ? trucksAssignedAt(truckAssignments, stationId, bounds.end) : [];

  const siteStockDebut = station ? stockBeforeDate(station, movements, bounds.start) : 0;
  const siteRangeMovs = station ? movementsInRange(movements, station.id, bounds.start, bounds.end) : [];
  const siteReception = sumQty(siteRangeMovs, ["reception"]);
  const siteRetourCamions = sumQty(siteRangeMovs, ["retour_camion"]);
  const siteVentesDirectes = sumQty(siteRangeMovs, ["sortie"]);
  const siteSortieCamion = sumQty(siteRangeMovs, ["sortie_camion"]);
  const siteStockTheoriqueFin = station ? stockThroughDate(station, movements, bounds.end) : 0;

  const truckRows = intervals.map((iv) => {
    const truck = sites.find((s) => s.id === iv.truckId);
    const intervalMovs = movements.filter((m) => m.siteId === iv.truckId && (m.product || "gasoil") === "gasoil" && m.date >= iv.start && m.date <= iv.end);
    const ventesTerrain = sumQty(intervalMovs, ["sortie"]);
    const chargement = sumQty(intervalMovs, ["reception"]);
    const retourCuve = sumQty(intervalMovs, ["retour_cuve_camion"]);
    const stillAssigned = truckIdsAtEnd.includes(iv.truckId);
    const stockFinTruck = stillAssigned && truck ? stockThroughDate(truck, movements, bounds.end) : null;
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
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />
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

/* ---- Synthèse Lubrifiants (Prehomo, Okouma) ---- */
function LubricantSynthesisReport({ sites, movements, inventaires, productStocks }) {
  const [month, setMonth] = useState(currentMonth());
  const bounds = monthBounds(month);

  const rows = [];
  LUBRICANT_SITE_IDS.forEach((siteId) => {
    const site = sites.find((s) => s.id === siteId);
    LUBRICANTS.forEach((lub) => {
      const ps = productStocks.find((p) => p.siteId === siteId && p.product === lub.id);
      const stockInitial = ps ? ps.stockInitial : 0;
      const stockDebut = stockBeforeDateProduct(stockInitial, movements, siteId, lub.id, bounds.start);
      const rangeMovs = movements.filter((m) => m.siteId === siteId && (m.product || "gasoil") === lub.id && m.date >= bounds.start && m.date <= bounds.end);
      const reception = sumQty(rangeMovs, ["reception"]);
      const sorties = sumQty(rangeMovs, ["sortie"]);
      const ajustement = rangeMovs.filter((m) => m.type === "ajustement").reduce((a, m) => a + m.delta, 0);
      const stockTheoriqueFin = stockDebut + reception - sorties + ajustement;
      const monthInv = inventaires.filter((i) => i.siteId === siteId && (i.product || "gasoil") === lub.id && i.date >= bounds.start && i.date <= bounds.end).sort((a, b) => (a.date < b.date ? 1 : -1));
      const dernierInv = pickLatestInv(monthInv);
      const stockFin = dernierInv ? dernierInv.stockPhysique : null;
      const ecart = stockFin !== null ? stockFin - stockTheoriqueFin : null;
      rows.push({
        site, siteId, lub, hasConfig: !!ps, stockDebut, reception, sorties, stockTheoriqueFin, stockFin, ecart,
      });
    });
  });

  const doExcel = () => exportToExcel(`SOMIP_Synthese_Lubrifiants_${month}.xlsx`, [{
    name: "Synthèse Lubrifiants", rows: rows.map((r) => ({
      Site: r.site?.name || r.siteId, Produit: r.lub.label, "Densité (kg/L)": r.lub.densite,
      "Stock début mois (L)": Math.round(r.stockDebut), "Stock début mois (kg)": Math.round(r.stockDebut * r.lub.densite),
      "Réception (L)": Math.round(r.reception), "Sorties (L)": Math.round(r.sorties),
      "Stock théorique fin (L)": Math.round(r.stockTheoriqueFin), "Stock théorique fin (kg)": Math.round(r.stockTheoriqueFin * r.lub.densite),
      "Stock fin mesuré (L)": r.stockFin !== null ? Math.round(r.stockFin) : "",
      "Stock fin mesuré (kg)": r.stockFin !== null ? Math.round(r.stockFin * r.lub.densite) : "",
      "Écart (L)": r.ecart !== null ? Math.round(r.ecart) : "", "Écart (kg)": r.ecart !== null ? Math.round(r.ecart * r.lub.densite) : "",
    })),
  }]);

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14 }}>
        <Field label="Mois de la synthèse"><input type="month" className="somip-input" style={{ maxWidth: 220 }} value={month} onChange={(e) => setMonth(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="Synthèse Lubrifiants — Prehomo &amp; Okouma" period={`Mois de ${bounds.start} au ${bounds.end}`} />
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead>
              <tr>
                <th>Site</th><th>Produit</th>
                <th style={{ textAlign: "right" }}>Stock début mois</th>
                <th style={{ textAlign: "right" }}>Réception</th><th style={{ textAlign: "right" }}>Sorties</th>
                <th style={{ textAlign: "right" }}>Stock théorique fin</th><th style={{ textAlign: "right" }}>Stock fin mesuré</th>
                <th style={{ textAlign: "right" }}>Écart</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={`${r.siteId}-${r.lub.id}`}>
                  <td style={{ fontWeight: 600 }}>{r.site?.name || r.siteId}</td>
                  <td>{r.lub.label} <span style={{ color: C.sub, fontSize: 11 }}>(d={r.lub.densite})</span></td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>
                    {r.hasConfig ? <>{fmt(r.stockDebut)} L<div style={{ color: C.sub, fontSize: 10.5 }}>{fmt(r.stockDebut * r.lub.densite)} kg</div></> : <span style={{ color: C.warning, fontSize: 11 }}>non configuré</span>}
                  </td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.reception ? C.success : C.sub }}>{r.reception ? `+${fmt(r.reception)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: r.sorties ? C.danger : C.sub }}>{r.sorties ? `−${fmt(r.sorties)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>
                    {fmt(r.stockTheoriqueFin)} L<div style={{ color: C.sub, fontSize: 10.5, fontWeight: 500 }}>{fmt(r.stockTheoriqueFin * r.lub.densite)} kg</div>
                  </td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>
                    {r.stockFin !== null ? <>{fmt(r.stockFin)} L<div style={{ color: C.sub, fontSize: 10.5, fontWeight: 500 }}>{fmt(r.stockFin * r.lub.densite)} kg</div></> : "—"}
                  </td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 600, color: r.ecart === null ? C.sub : r.ecart < 0 ? C.danger : r.ecart > 0 ? C.success : C.sub }}>
                    {r.ecart !== null ? `${r.ecart >= 0 ? "+" : ""}${fmt(r.ecart)} L` : "—"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p style={{ marginTop: 14, fontSize: 11, color: C.sub }}>
          Stock fin mesuré = dernier inventaire enregistré pour ce produit dans le mois. "Non configuré" = capacité/stock initial pas encore défini sur la page Sites.
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
function UsersView({ profiles, updateUserRole, session }) {
  const [editingId, setEditingId] = useState(null);
  const [roleDraft, setRoleDraft] = useState("");
  const [form, setForm] = useState({ fullName: "", email: "", password: "", role: "lecture" });
  const [creating, setCreating] = useState(false);
  const [createErr, setCreateErr] = useState(null);
  const [createMsg, setCreateMsg] = useState(null);

  const startEdit = (u) => { setEditingId(u.id); setRoleDraft(u.role); };
  const saveEdit = () => { updateUserRole(editingId, roleDraft); setEditingId(null); };

  const createAccount = async () => {
    setCreateErr(null); setCreateMsg(null);
    if (!form.fullName.trim() || !form.email.trim() || !form.password) { setCreateErr("Tous les champs sont requis."); return; }
    if (form.password.length < 6) { setCreateErr("Le mot de passe doit contenir au moins 6 caractères."); return; }
    setCreating(true);
    try {
      const res = await fetch("/api/create-user", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${session?.access_token || ""}` },
        body: JSON.stringify({ email: form.email.trim(), password: form.password, fullName: form.fullName.trim(), role: form.role }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Erreur lors de la création du compte.");
      setCreateMsg(`Compte créé pour ${form.fullName.trim()} (${ROLE_LABELS[form.role]}). Communique-lui l'e-mail et le mot de passe.`);
      setForm({ fullName: "", email: "", password: "", role: "lecture" });
    } catch (e) {
      setCreateErr(e.message || "Erreur lors de la création du compte.");
    } finally {
      setCreating(false);
    }
  };

  return (
    <div className="somip-fade" style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      <div className="somip-panel" style={{ flex: "1 1 520px", padding: 18 }}>
        <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Comptes ({profiles.length})</h3>
        <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.sub }}>
          Créés par toi ci-contre, ou par auto-inscription (rôle "Lecture" par défaut dans ce cas) — modifie le rôle ici à tout moment.
        </p>
        <table className="somip-table">
          <thead><tr><th>Nom</th><th>Rôle</th><th></th></tr></thead>
          <tbody>
            {profiles.length === 0 && <EmptyRow colSpan={3} text="Aucun compte pour le moment." />}
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
                      <td style={{ whiteSpace: "nowrap" }}>
                        <button className="somip-btn somip-btn-primary" style={{ padding: "5px 10px", fontSize: 12 }} onClick={saveEdit}>OK</button>
                        <button onClick={() => setEditingId(null)} style={{ border: "none", background: "none", cursor: "pointer", marginLeft: 4 }}><X size={16} color={C.sub} /></button>
                      </td>
                    </>
                  ) : (
                    <>
                      <td style={{ fontWeight: 600 }}>{u.name}</td>
                      <td><Badge color={C.blue}>{ROLE_LABELS[u.role] || u.role}</Badge></td>
                      <td style={{ whiteSpace: "nowrap", textAlign: "right" }}>
                        <button onClick={() => startEdit(u)} style={{ border: "none", background: "none", cursor: "pointer", padding: 5 }}><Pencil size={14} color={C.sub} /></button>
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
        <Field label="E-mail"><input type="email" className="somip-input" value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} placeholder="prenom.nom@somip-sarl.ga" /></Field>
        <Field label="Mot de passe provisoire"><input type="text" className="somip-input" value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} placeholder="Au moins 6 caractères" /></Field>
        <Field label="Rôle">
          <select className="somip-select" value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })}>
            {ROLE_VALUES.map((r) => <option key={r} value={r}>{ROLE_LABELS[r]}</option>)}
          </select>
        </Field>
        {createErr && <p style={{ color: C.danger, fontSize: 12.5, margin: "0 0 10px" }}>{createErr}</p>}
        {createMsg && <p style={{ color: C.success, fontSize: 12.5, margin: "0 0 10px" }}>{createMsg}</p>}
        <button className="somip-btn somip-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={createAccount} disabled={creating}>
          <Plus size={15} /> {creating ? "Création..." : "Créer le compte"}
        </button>
        <p style={{ marginTop: 10, fontSize: 11, color: C.sub }}>
          La personne peut se connecter immédiatement avec cet e-mail et ce mot de passe. Communique-les-lui directement.
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
