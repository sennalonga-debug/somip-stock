import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  LayoutDashboard, Factory, ArrowDownCircle, ArrowUpCircle, ClipboardList,
  Truck, AlertTriangle, Plus, X, Trash2, Pencil, Fuel, RotateCcw, Check,
  Users, History, Loader2, CheckCircle2, AlertCircle, CloudOff, Thermometer,
  FileBarChart, Download, Printer, TrendingDown, TrendingUp,
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell, ReferenceLine, Legend,
} from "recharts";

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

const TRUCKS = Array.from({ length: 6 }, (_, i) => `Camion Laitier ${i + 1}`);

const MOVEMENTS_SEED = [
  { id: "m1", siteId: "okouma", type: "reception", date: "2026-08-28", quantity: 15000, delta: 15000, ref: "BL-2891", commentaire: "Livraison TotalEnergies", isDemo: true },
  { id: "m2", siteId: "okouma", type: "sortie_camion", date: "2026-08-30", quantity: 4200, delta: -4200, camion: "Camion Laitier 2", destination: "Carrière Nord — Engins", commentaire: "", isDemo: true },
  { id: "m3", siteId: "prehomo", type: "sortie", date: "2026-09-01", quantity: 2600, delta: -2600, destinataire: "Atelier mécanique", commentaire: "", isDemo: true },
  { id: "m4", siteId: "lipaka", type: "reception", date: "2026-09-02", quantity: 8000, delta: 8000, ref: "BL-2903", commentaire: "", isDemo: true },
  { id: "m5", siteId: "traction", type: "sortie", date: "2026-09-02", quantity: 1800, delta: -1800, destinataire: "Locomotive 12", commentaire: "", isDemo: true },
];

const ROLES = ["Superviseur", "Opérateur", "Chauffeur", "Lecture"];
const USERS_SEED = [
  { id: "u_albain", name: "Albain Longa", role: "Superviseur" },
  { id: "u_demo", name: "Utilisateur Démo", role: "Lecture" },
];

const SETTINGS_SEED = { objectifFreinte: 3 };

const TYPE_META = {
  reception: { label: "Réception", color: C.success, sign: "+" },
  sortie: { label: "Sortie", color: C.danger, sign: "−" },
  sortie_camion: { label: "Sortie vers camion laitier", color: C.orange, sign: "−" },
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
  return movements.filter((m) => m.siteId === site.id && m.date < dateExclusive).reduce((a, m) => a + m.delta, site.stockInitial);
}
function stockThroughDate(site, movements, dateInclusive) {
  return movements.filter((m) => m.siteId === site.id && m.date <= dateInclusive).reduce((a, m) => a + m.delta, site.stockInitial);
}
function movementsInRange(movements, siteId, startInclusive, endInclusive) {
  return movements.filter((m) => m.siteId === siteId && m.date >= startInclusive && m.date <= endInclusive);
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
/* Persistent storage layer                                             */
/* Two modes:                                                           */
/*  - "claude"     : window.storage (shared key-value store, only       */
/*                    available when this component runs inside a       */
/*                    Claude.ai artifact).                              */
/*  - "local"      : browser localStorage (used automatically once      */
/*                    this app is hosted on its own — e.g. Vercel/      */
/*                    Netlify/VPS). Persistence is then per-browser,    */
/*                    NOT shared between different users/devices,       */
/*                    since there is no backend server here. Wiring a   */
/*                    real shared database (e.g. via a small API) is    */
/*                    the natural next step for true multi-device       */
/*                    multi-user persistence in production.             */
/* ------------------------------------------------------------------ */
const STORAGE_MODE = typeof window !== "undefined" && window.storage ? "claude" : "local";
const STORAGE_AVAILABLE = typeof window !== "undefined" && (!!window.storage || !!window.localStorage);
const KEY = (name) => `somip_${name}`;

async function loadRaw(name) {
  if (!STORAGE_AVAILABLE) return null;
  try {
    if (STORAGE_MODE === "claude") {
      const res = await window.storage.get(KEY(name), true);
      return res && res.value ? JSON.parse(res.value) : null;
    }
    const raw = window.localStorage.getItem(KEY(name));
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null; // key not found yet, or transient error — caller decides the default
  }
}

async function saveRaw(name, value) {
  if (!STORAGE_AVAILABLE) return false;
  try {
    if (STORAGE_MODE === "claude") {
      const res = await window.storage.set(KEY(name), JSON.stringify(value), true);
      return !!res;
    }
    window.localStorage.setItem(KEY(name), JSON.stringify(value));
    return true;
  } catch (e) {
    return false;
  }
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
/* App                                                                   */
/* ------------------------------------------------------------------ */
export default function App() {
  const [loading, setLoading] = useState(true);
  const [sites, setSites] = useState([]);
  const [movements, setMovements] = useState([]);
  const [inventaires, setInventaires] = useState([]);
  const [users, setUsers] = useState([]);
  const [audit, setAudit] = useState([]);
  const [settings, setSettings] = useState(SETTINGS_SEED);
  const [currentUserId, setCurrentUserId] = useState(null);
  const [view, setView] = useState("dashboard");
  const [notice, setNotice] = useState(null);
  const [syncStatus, setSyncStatus] = useState(STORAGE_AVAILABLE ? "ok" : "unavailable");
  const [lastSync, setLastSync] = useState(null);
  const noticeTimer = useRef(null);

  const flash = (msg) => {
    clearTimeout(noticeTimer.current);
    setNotice(msg);
    noticeTimer.current = setTimeout(() => setNotice(null), 3000);
  };

  const currentUserName = users.find((u) => u.id === currentUserId)?.name || "Utilisateur inconnu";

  /* ---- initial load ---- */
  useEffect(() => {
    let cancelled = false;
    (async () => {
      const [sitesData, movementsData, inventairesData, usersData, auditData, settingsData] = await Promise.all([
        loadRaw("sites"), loadRaw("movements"), loadRaw("inventaires"), loadRaw("users"), loadRaw("audit"), loadRaw("settings"),
      ]);
      if (cancelled) return;
      const finalSites = sitesData ?? SITES_SEED;
      const finalMovements = movementsData ?? MOVEMENTS_SEED;
      const finalInventaires = inventairesData ?? [];
      const finalUsers = usersData ?? USERS_SEED;
      const finalAudit = auditData ?? [];
      const finalSettings = settingsData ?? SETTINGS_SEED;
      setSites(finalSites);
      setMovements(finalMovements);
      setInventaires(finalInventaires);
      setUsers(finalUsers);
      setAudit(finalAudit);
      setSettings(finalSettings);
      setCurrentUserId(finalUsers[0]?.id || null);
      if (STORAGE_AVAILABLE) {
        if (sitesData === null) saveRaw("sites", finalSites);
        if (movementsData === null) saveRaw("movements", finalMovements);
        if (inventairesData === null) saveRaw("inventaires", finalInventaires);
        if (usersData === null) saveRaw("users", finalUsers);
        if (auditData === null) saveRaw("audit", finalAudit);
        if (settingsData === null) saveRaw("settings", finalSettings);
        setLastSync(new Date());
      }
      setLoading(false);
    })();
    return () => { cancelled = true; };
  }, []);

  /* ---- background sync: pick up changes made by other users ---- */
  useEffect(() => {
    if (loading || !STORAGE_AVAILABLE) return;
    const interval = setInterval(async () => {
      const [s, m, i, u, a, se] = await Promise.all([
        loadRaw("sites"), loadRaw("movements"), loadRaw("inventaires"), loadRaw("users"), loadRaw("audit"), loadRaw("settings"),
      ]);
      if (s !== null) setSites(s);
      if (m !== null) setMovements(m);
      if (i !== null) setInventaires(i);
      if (u !== null) setUsers(u);
      if (a !== null) setAudit(a);
      if (se !== null) setSettings(se);
      setLastSync(new Date());
    }, 7000);
    return () => clearInterval(interval);
  }, [loading]);

  /* ---- persistence helper ---- */
  const persist = async (name, value) => {
    if (!STORAGE_AVAILABLE) { setSyncStatus("unavailable"); return; }
    setSyncStatus("saving");
    const ok = await saveRaw(name, value);
    setSyncStatus(ok ? "ok" : "error");
    if (ok) setLastSync(new Date());
  };

  const appendAudit = (action, detail) => {
    const entry = { id: uid(), ts: new Date().toISOString(), user: currentUserName, action, detail };
    const next = [entry, ...audit].slice(0, 300);
    setAudit(next);
    persist("audit", next);
  };

  /* ---- derived ---- */
  // Stock "ambiant" (jauge physique) — sert au tableau de bord, à la capacité des cuves.
  const stockOf = (siteId) => {
    const site = sites.find((s) => s.id === siteId);
    if (!site) return 0;
    return movements.filter((m) => m.siteId === siteId).reduce((acc, m) => acc + m.delta, site.stockInitial);
  };
  // Stock théorique "à 15°C" — utilise le volume corrigé de chaque mouvement quand il est
  // disponible (température + densité renseignées), sinon retombe sur le volume ambiant de
  // ce mouvement. C'est cette valeur qui sert de référence pour l'écart d'inventaire dès
  // qu'une mesure de température/densité est fournie.
  const stockOf15 = (siteId) => {
    const site = sites.find((s) => s.id === siteId);
    if (!site) return 0;
    return movements.filter((m) => m.siteId === siteId).reduce((acc, m) => acc + Math.sign(m.delta) * movementQty15(m), site.stockInitial);
  };

  /* ---- mutations: sites ---- */
  const addSite = (form) => {
    const site = { id: uid(), name: form.name.trim(), code: form.code.trim().toUpperCase(), capacity: Number(form.capacity), stockInitial: Number(form.stockInitial) || 0 };
    const next = [...sites, site];
    setSites(next); persist("sites", next);
    appendAudit("Ajout site", `${site.name} (${site.code})`);
    flash("Site ajouté.");
  };
  const editSite = (id, patch) => {
    const next = sites.map((s) => (s.id === id ? { ...s, ...patch, capacity: Number(patch.capacity), stockInitial: Number(patch.stockInitial) } : s));
    setSites(next); persist("sites", next);
    appendAudit("Modification site", `${patch.name} (${patch.code})`);
    flash("Site mis à jour.");
  };
  const removeSite = (site) => {
    if (movements.some((m) => m.siteId === site.id)) { flash(`Impossible : "${site.name}" a des mouvements enregistrés.`); return; }
    const next = sites.filter((s) => s.id !== site.id);
    setSites(next); persist("sites", next);
    appendAudit("Suppression site", site.name);
    flash("Site supprimé.");
  };

  /* ---- mutations: movements ---- */
  const addMovement = (payload) => {
    const record = { id: uid(), createdBy: currentUserName, createdAt: new Date().toISOString(), isDemo: false, ...payload };
    const next = [...movements, record];
    setMovements(next); persist("movements", next);
    appendAudit(TYPE_META[payload.type].label, `${fmt(payload.quantity)} L — ${sites.find((s) => s.id === payload.siteId)?.name || ""}`);
    return record;
  };
  const deleteMovement = (id) => {
    const m = movements.find((mm) => mm.id === id);
    const next = movements.filter((mm) => mm.id !== id);
    setMovements(next); persist("movements", next);
    if (m) appendAudit("Suppression mouvement", `${TYPE_META[m.type]?.label || m.type} — ${fmt(m.quantity)} L`);
  };
  const purgeDemoMovements = () => {
    const count = movements.filter((m) => m.isDemo).length;
    const next = movements.filter((m) => !m.isDemo);
    setMovements(next); persist("movements", next);
    appendAudit("Purge données démo", `${count} écriture(s) d'exemple supprimée(s)`);
    flash("Écritures de démonstration supprimées.");
  };

  /* ---- mutations: inventaires ---- */
  const addInventaire = ({ siteId, date, stockPhysique, commentaire, temperatureC, densiteObservee, densite15, vcf, stockPhysique15 }) => {
    const theoriqueAmbiant = stockOf(siteId);
    const theorique15 = stockOf15(siteId);
    const has15 = stockPhysique15 !== undefined;
    // Dès que la température et la densité sont renseignées, l'écart officiel (perte/gain,
    // taux de freinte, conformité) est calculé sur la base du volume corrigé à 15°C.
    // Sans mesure de température/densité, on retombe sur le volume ambiant (comportement V1/V2).
    const basisEcart = has15 ? "15c" : "ambiant";
    const theoriqueUsed = has15 ? theorique15 : theoriqueAmbiant;
    const physiqueUsed = has15 ? stockPhysique15 : stockPhysique;
    const ecart = physiqueUsed - theoriqueUsed;
    const cls = classifyEcart(ecart, theoriqueUsed, settings.objectifFreinte);
    const adjId = uid();
    const vcfFields = has15 ? { temperatureC, densiteObservee, densite15, vcf, stockPhysique15 } : {};
    const nextMovements = [...movements, {
      id: adjId, siteId, type: "ajustement", date, quantity: Math.abs(ecart), delta: ecart, isDemo: false,
      commentaire: `Ajustement suite à l'inventaire du ${date} (base ${has15 ? "15°C" : "ambiante"})`,
      createdBy: currentUserName, createdAt: new Date().toISOString(),
    }];
    const nextInventaires = [...inventaires, {
      id: uid(), siteId, date, stockPhysique, commentaire, basisEcart,
      stockTheoriqueAmbiant: theoriqueAmbiant, stockTheorique15: theorique15,
      stockTheorique: theoriqueUsed, stockPhysiqueUsed: physiqueUsed,
      ecart: cls.ecartL, ecartPermille: cls.ecartPermille, nature: cls.nature,
      tauxFreinte: cls.tauxFreinte, objectifUtilise: cls.objectif, conformite: cls.conformite,
      adjustmentId: adjId, createdBy: currentUserName, createdAt: new Date().toISOString(), ...vcfFields,
    }];
    setMovements(nextMovements); persist("movements", nextMovements);
    setInventaires(nextInventaires); persist("inventaires", nextInventaires);
    appendAudit("Inventaire", `${sites.find((s) => s.id === siteId)?.name || ""} — base ${has15 ? "15°C" : "ambiante"} — ${NATURE_META[cls.nature].label} ${ecart >= 0 ? "+" : ""}${fmt(ecart)} L (${cls.ecartPermille >= 0 ? "+" : ""}${cls.ecartPermille.toFixed(2)} ‰)`);
    flash("Inventaire enregistré et stock ajusté.");
  };

  const updateSettings = (patch) => {
    const next = { ...settings, ...patch };
    setSettings(next); persist("settings", next);
    appendAudit("Modification objectif de freinte", `Nouvel objectif : ${next.objectifFreinte} ‰`);
    flash("Objectif mis à jour.");
  };
  const deleteInventaire = (inv) => {
    const nextInv = inventaires.filter((i) => i.id !== inv.id);
    const nextMov = movements.filter((m) => m.id !== inv.adjustmentId);
    setInventaires(nextInv); persist("inventaires", nextInv);
    setMovements(nextMov); persist("movements", nextMov);
    appendAudit("Suppression inventaire", `${sites.find((s) => s.id === inv.siteId)?.name || ""} — ${inv.date}`);
  };

  /* ---- mutations: users ---- */
  const addUser = (name, role) => {
    const u = { id: uid(), name: name.trim(), role };
    const next = [...users, u];
    setUsers(next); persist("users", next);
    appendAudit("Ajout utilisateur", `${u.name} (${u.role})`);
    flash("Utilisateur ajouté.");
  };
  const editUser = (id, patch) => {
    const next = users.map((u) => (u.id === id ? { ...u, ...patch } : u));
    setUsers(next); persist("users", next);
    appendAudit("Modification utilisateur", `${patch.name} (${patch.role})`);
    flash("Utilisateur mis à jour.");
  };
  const removeUser = (u) => {
    if (users.length <= 1) { flash("Impossible : au moins un utilisateur doit rester."); return; }
    const next = users.filter((x) => x.id !== u.id);
    setUsers(next); persist("users", next);
    if (currentUserId === u.id) setCurrentUserId(next[0]?.id || null);
    appendAudit("Suppression utilisateur", u.name);
    flash("Utilisateur supprimé.");
  };

  /* ---- dangerous: full factory reset (kept for testing, double-armed) ---- */
  const factoryReset = () => {
    setSites(SITES_SEED); persist("sites", SITES_SEED);
    setMovements(MOVEMENTS_SEED); persist("movements", MOVEMENTS_SEED);
    setInventaires([]); persist("inventaires", []);
    setUsers(USERS_SEED); persist("users", USERS_SEED);
    setCurrentUserId(USERS_SEED[0].id);
    appendAudit("Réinitialisation complète", "Toutes les données remplacées par le jeu de démonstration");
    flash("Toutes les données ont été réinitialisées.");
  };

  const NAV = [
    { id: "dashboard", label: "Tableau de bord", icon: LayoutDashboard },
    { id: "sites", label: "Sites", icon: Factory },
    { id: "receptions", label: "Réceptions", icon: ArrowDownCircle },
    { id: "sorties", label: "Sorties", icon: ArrowUpCircle },
    { id: "inventaires", label: "Inventaires", icon: ClipboardList },
    { id: "vcf", label: "Correction 15°C", icon: Thermometer },
    { id: "rapports", label: "Rapports", icon: FileBarChart },
    { id: "utilisateurs", label: "Utilisateurs", icon: Users },
    { id: "historique", label: "Historique", icon: History },
  ];
  const viewTitle = NAV.find((n) => n.id === view)?.label || "";

  if (loading) {
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
        @media print {
          .somip-no-print, .somip-sidebar, .somip-header { display: none !important; }
          .somip-print-only { display: block !important; }
          .somip-scroll { overflow: visible !important; height: auto !important; padding: 0 !important; }
          body, .somip-app { background: #fff !important; }
          .somip-panel { border: none !important; }
        }
      `}</style>

      {/* Sidebar */}
      <aside className="somip-sidebar" style={{ width: 226, background: `linear-gradient(180deg, ${C.navy}, ${C.navyLight})`, display: "flex", flexDirection: "column", padding: "20px 14px", flexShrink: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 9, padding: "0 6px 22px" }}>
          <div style={{ width: 32, height: 32, borderRadius: 8, background: C.blue, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Fuel size={17} color="#fff" />
          </div>
          <div>
            <div style={{ color: "#fff", fontWeight: 700, fontSize: 14.5, letterSpacing: 0.2 }}>SOMIP</div>
            <div style={{ color: "#8CA0B4", fontSize: 10.5, fontWeight: 500 }}>Stock Gasoil</div>
          </div>
        </div>
        <nav style={{ display: "flex", flexDirection: "column", gap: 3 }}>
          {NAV.map((n) => (
            <button key={n.id} className={`somip-nav-item ${view === n.id ? "active" : ""}`} onClick={() => setView(n.id)}>
              <n.icon size={16} />{n.label}
            </button>
          ))}
        </nav>
        <div style={{ flex: 1 }} />
        <ConfirmTextButton
          onConfirm={factoryReset}
          label="Réinitialiser (usine)"
          confirmLabel="Effacer & recharger la démo ?"
          className="somip-nav-item"
        />
        <div style={{ color: "#5C7288", fontSize: 10.5, padding: "10px 6px 0" }}>Zone Sud-Est · Gabon</div>
      </aside>

      {/* Main */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", minWidth: 0 }}>
        <header className="somip-header" style={{ padding: "16px 28px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "space-between", background: "#fff", gap: 16, flexWrap: "wrap" }}>
          <div>
            <h1 style={{ margin: 0, fontSize: 18, fontWeight: 700 }}>{viewTitle}</h1>
            <SyncIndicator status={syncStatus} lastSync={lastSync} />
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <div style={{ fontSize: 12, color: C.sub }}>
              <div style={{ fontWeight: 600, color: C.sub, marginBottom: 2 }}>Connecté en tant que</div>
              <select className="somip-select" style={{ padding: "6px 9px", fontSize: 12.5, minWidth: 170 }} value={currentUserId || ""} onChange={(e) => setCurrentUserId(e.target.value)}>
                {users.map((u) => <option key={u.id} value={u.id}>{u.name} — {u.role}</option>)}
              </select>
            </div>
            <div style={{ fontSize: 12.5, color: C.sub, textAlign: "right" }}>
              {new Date().toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long", year: "numeric" })}
              {notice && <div className="somip-fade" style={{ marginTop: 4, color: C.success, fontWeight: 600 }}>{notice}</div>}
            </div>
          </div>
        </header>

        <div className="somip-scroll" style={{ flex: 1, padding: "24px 28px" }}>
          {view === "dashboard" && <Dashboard sites={sites} movements={movements} inventaires={inventaires} stockOf={stockOf} purgeDemoMovements={purgeDemoMovements} />}
          {view === "sites" && <SitesView sites={sites} movements={movements} stockOf={stockOf} addSite={addSite} editSite={editSite} removeSite={removeSite} />}
          {view === "receptions" && <ReceptionsView sites={sites} movements={movements} addMovement={addMovement} deleteMovement={deleteMovement} />}
          {view === "sorties" && <SortiesView sites={sites} movements={movements} addMovement={addMovement} deleteMovement={deleteMovement} />}
          {view === "inventaires" && <InventairesView sites={sites} inventaires={inventaires} stockOf={stockOf} stockOf15={stockOf15} addInventaire={addInventaire} deleteInventaire={deleteInventaire} settings={settings} updateSettings={updateSettings} />}
          {view === "vcf" && <VcfView />}
          {view === "rapports" && <ReportsView sites={sites} movements={movements} inventaires={inventaires} settings={settings} stockOf={stockOf} />}
          {view === "utilisateurs" && <UsersView users={users} addUser={addUser} editUser={editUser} removeUser={removeUser} />}
          {view === "historique" && <HistoryView audit={audit} />}
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Dashboard                                                            */
/* ------------------------------------------------------------------ */
function Dashboard({ sites, movements, inventaires, stockOf, purgeDemoMovements }) {
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
  const receptionsMonth = movements.filter((m) => m.type === "reception" && m.date.startsWith(month)).reduce((a, m) => a + m.quantity, 0);
  const sortiesMonth = movements.filter((m) => (m.type === "sortie" || m.type === "sortie_camion") && m.date.startsWith(month)).reduce((a, m) => a + m.quantity, 0);
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

      {demoCount > 0 && (
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
function SitesView({ sites, movements, stockOf, addSite, editSite, removeSite }) {
  const [form, setForm] = useState({ name: "", code: "", capacity: "", stockInitial: "" });
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});

  const submitAdd = () => {
    if (!form.name.trim() || !form.code.trim() || !form.capacity) return;
    addSite(form);
    setForm({ name: "", code: "", capacity: "", stockInitial: "" });
  };
  const startEdit = (s) => { setEditingId(s.id); setEditForm({ ...s }); };
  const saveEdit = () => { editSite(editingId, editForm); setEditingId(null); };

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
                      <td>{s.name}</td>
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
        <Field label="Nom du site"><input className="somip-input" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="Ex : Dépôt Moanda" /></Field>
        <Field label="Code (court)"><input className="somip-input" value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} placeholder="Ex : DPM" /></Field>
        <Field label="Capacité (L)"><input type="number" className="somip-input" value={form.capacity} onChange={(e) => setForm({ ...form, capacity: e.target.value })} placeholder="30000" /></Field>
        <Field label="Stock initial (L)"><input type="number" className="somip-input" value={form.stockInitial} onChange={(e) => setForm({ ...form, stockInitial: e.target.value })} placeholder="0" /></Field>
        <button className="somip-btn somip-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={submitAdd} disabled={!form.name || !form.code || !form.capacity}>
          <Plus size={15} /> Ajouter le site
        </button>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Réceptions                                                            */
/* ------------------------------------------------------------------ */
function ReceptionsView({ sites, movements, addMovement, deleteMovement }) {
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

      <div className="somip-panel" style={{ flex: "2 1 480px", padding: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <h3 style={{ margin: 0, fontSize: 14 }}>Historique des réceptions</h3>
          <select className="somip-select" style={{ width: 200 }} value={filterSite} onChange={(e) => setFilterSite(e.target.value)}>
            <option value="all">Tous les sites</option>
            {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </div>
        <table className="somip-table">
          <thead><tr><th>Date</th><th>Site</th><th>Référence</th><th style={{ textAlign: "right" }}>Quantité (ambiant)</th><th style={{ textAlign: "right" }}>Volume 15°C</th><th></th></tr></thead>
          <tbody>
            {list.length === 0 && <EmptyRow colSpan={6} text="Aucune réception enregistrée." />}
            {list.map((m) => {
              const site = sites.find((s) => s.id === m.siteId);
              return (
                <tr key={m.id}>
                  <td className="somip-mono">{m.date}</td>
                  <td>{site?.name}{m.isDemo && <DemoBadge />}</td>
                  <td style={{ color: C.sub }}>{m.ref || "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success, fontWeight: 600 }}>+ {fmt(m.quantity)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{m.volumeCorrige15 ? `${fmt(m.volumeCorrige15)} L` : "—"}</td>
                  <td style={{ textAlign: "right" }}><ConfirmIconButton onConfirm={() => deleteMovement(m.id)} /></td>
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
function SortiesView({ sites, movements, addMovement, deleteMovement }) {
  const [tab, setTab] = useState("sortie");
  const [filterSite, setFilterSite] = useState("all");
  const [form, setForm] = useState({ siteId: sites[0]?.id || "", date: todayStr(), quantity: "", destinataire: "", camion: TRUCKS[0], destination: "", commentaire: "" });
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
    const base = { siteId: form.siteId, type: tab, date: form.date, quantity: Number(form.quantity), delta: -Number(form.quantity), commentaire: form.commentaire, ...extra };
    const payload = tab === "sortie" ? { ...base, destinataire: form.destinataire } : { ...base, camion: form.camion, destination: form.destination };
    addMovement(payload);
    setForm({ ...form, quantity: "", destinataire: "", destination: "", commentaire: "" });
    setTempC(""); setDensite("");
  };

  const list = movements.filter((m) => m.type === "sortie" || m.type === "sortie_camion").filter((m) => filterSite === "all" || m.siteId === filterSite).sort((a, b) => (a.date < b.date ? 1 : -1));

  return (
    <div className="somip-fade" style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      <div className="somip-panel" style={{ flex: "1 1 300px", padding: 18 }}>
        <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          <button className={`somip-tab ${tab === "sortie" ? "active" : ""}`} onClick={() => setTab("sortie")}>Sortie standard</button>
          <button className={`somip-tab ${tab === "sortie_camion" ? "active" : ""}`} onClick={() => setTab("sortie_camion")}>Vers camion laitier</button>
        </div>
        <Field label="Site source">
          <select className="somip-select" value={form.siteId} onChange={(e) => setForm({ ...form, siteId: e.target.value })}>
            {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </Field>
        <Field label="Date"><input type="date" className="somip-input" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></Field>
        <Field label="Quantité (L)"><input type="number" className="somip-input" value={form.quantity} onChange={(e) => setForm({ ...form, quantity: e.target.value })} placeholder="2000" /></Field>
        {tab === "sortie" ? (
          <Field label="Destinataire / motif"><input className="somip-input" value={form.destinataire} onChange={(e) => setForm({ ...form, destinataire: e.target.value })} placeholder="Ex : Atelier, Engin X..." /></Field>
        ) : (
          <>
            <Field label="Camion laitier">
              <select className="somip-select" value={form.camion} onChange={(e) => setForm({ ...form, camion: e.target.value })}>
                {TRUCKS.map((t) => <option key={t} value={t}>{t}</option>)}
              </select>
            </Field>
            <Field label="Destination (carrière / engin)"><input className="somip-input" value={form.destination} onChange={(e) => setForm({ ...form, destination: e.target.value })} placeholder="Ex : Carrière Nord" /></Field>
          </>
        )}
        <Field label="Commentaire (optionnel)"><textarea className="somip-textarea" rows={2} value={form.commentaire} onChange={(e) => setForm({ ...form, commentaire: e.target.value })} /></Field>
        <VcfMiniPanel tempC={tempC} densite={densite} onTempC={setTempC} onDensite={setDensite} result={vcfResult} compact />
        <button className="somip-btn somip-btn-secondary" style={{ width: "100%", justifyContent: "center" }} onClick={submit} disabled={!form.quantity || Number(form.quantity) <= 0}>
          <Plus size={15} /> Enregistrer la sortie
        </button>
      </div>

      <div className="somip-panel" style={{ flex: "2 1 480px", padding: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <h3 style={{ margin: 0, fontSize: 14 }}>Historique des sorties</h3>
          <select className="somip-select" style={{ width: 200 }} value={filterSite} onChange={(e) => setFilterSite(e.target.value)}>
            <option value="all">Tous les sites</option>
            {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </div>
        <table className="somip-table">
          <thead><tr><th>Date</th><th>Site</th><th>Type</th><th>Détail</th><th style={{ textAlign: "right" }}>Quantité (ambiant)</th><th style={{ textAlign: "right" }}>Volume 15°C</th><th></th></tr></thead>
          <tbody>
            {list.length === 0 && <EmptyRow colSpan={7} text="Aucune sortie enregistrée." />}
            {list.map((m) => {
              const site = sites.find((s) => s.id === m.siteId);
              const meta = TYPE_META[m.type];
              const detail = m.type === "sortie" ? (m.destinataire || "—") : `${m.camion} → ${m.destination || "—"}`;
              return (
                <tr key={m.id}>
                  <td className="somip-mono">{m.date}</td>
                  <td>{site?.name}{m.isDemo && <DemoBadge />}</td>
                  <td><Badge color={meta.color}>{m.type === "sortie" ? "Standard" : "Camion"}</Badge></td>
                  <td style={{ color: C.sub }}>{detail}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: meta.color, fontWeight: 600 }}>− {fmt(m.quantity)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{m.volumeCorrige15 ? `${fmt(m.volumeCorrige15)} L` : "—"}</td>
                  <td style={{ textAlign: "right" }}><ConfirmIconButton onConfirm={() => deleteMovement(m.id)} /></td>
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
/* Inventaires                                                           */
/* ------------------------------------------------------------------ */
function InventairesView({ sites, inventaires, stockOf, stockOf15, addInventaire, deleteInventaire, settings, updateSettings }) {
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
                <th style={{ textAlign: "right" }}>Taux de freinte</th><th>Statut (objectif)</th><th></th>
              </tr>
            </thead>
            <tbody>
              {list.length === 0 && <EmptyRow colSpan={11} text="Aucun inventaire enregistré." />}
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
                    <td style={{ textAlign: "right" }}><ConfirmIconButton onConfirm={() => deleteInventaire(i)} /></td>
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
function ReportsView({ sites, movements, inventaires, settings, stockOf }) {
  const [tab, setTab] = useState("journalier");
  const TABS = [
    { id: "journalier", label: "Journalier" },
    { id: "decadaire", label: "Décadaire" },
    { id: "mensuel", label: "Mensuel" },
    { id: "etat", label: "État des stocks" },
    { id: "exposition", label: "Exposition" },
    { id: "pertesgains", label: "Pertes/Gains par site" },
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
      {tab === "etat" && <StockStatementReport sites={sites} movements={movements} />}
      {tab === "exposition" && <ExposureReport sites={sites} movements={movements} inventaires={inventaires} stockOf={stockOf} />}
      {tab === "pertesgains" && <LossGainReport sites={sites} inventaires={inventaires} />}
    </div>
  );
}

/* ---- Rapport journalier ---- */
function DailyReport({ sites, movements, inventaires }) {
  const [date, setDate] = useState(todayStr());
  const rows = sites.map((s) => {
    const stockDebut = stockBeforeDate(s, movements, date);
    const dayMovs = movements.filter((m) => m.siteId === s.id && m.date === date);
    const receptions = sumQty(dayMovs, ["reception"]);
    const sorties = sumQty(dayMovs, ["sortie", "sortie_camion"]);
    const ajustement = dayMovs.filter((m) => m.type === "ajustement").reduce((a, m) => a + m.delta, 0);
    const stockFin = stockDebut + receptions - sorties + ajustement;
    const inv = inventaires.find((i) => i.siteId === s.id && i.date === date);
    return { site: s, stockDebut, receptions, sorties, ajustement, stockFin, inv };
  });
  const totals = rows.reduce((a, r) => ({ stockDebut: a.stockDebut + r.stockDebut, receptions: a.receptions + r.receptions, sorties: a.sorties + r.sorties, stockFin: a.stockFin + r.stockFin }), { stockDebut: 0, receptions: 0, sorties: 0, stockFin: 0 });

  const doExcel = () => exportToExcel(`SOMIP_Rapport_Journalier_${date}.xlsx`, [{
    name: "Journalier", rows: rows.map((r) => ({
      Site: r.site.name, "Stock début (L)": Math.round(r.stockDebut), "Réceptions (L)": Math.round(r.receptions),
      "Sorties (L)": Math.round(r.sorties), "Ajustement inventaire (L)": Math.round(r.ajustement), "Stock fin (L)": Math.round(r.stockFin),
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
            <thead><tr><th>Site</th><th style={{ textAlign: "right" }}>Stock début</th><th style={{ textAlign: "right" }}>Réceptions</th><th style={{ textAlign: "right" }}>Sorties</th><th style={{ textAlign: "right" }}>Stock fin</th><th>Inventaire du jour</th></tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(r.stockDebut)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>{r.receptions ? `+${fmt(r.receptions)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.danger }}>{r.sorties ? `−${fmt(r.sorties)} L` : "—"}</td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stockFin)} L</td>
                  <td>{r.inv ? <Badge color={NATURE_META[r.inv.nature].color}>{NATURE_META[r.inv.nature].label} {r.inv.ecart >= 0 ? "+" : ""}{fmt(r.inv.ecart)} L</Badge> : <span style={{ color: C.sub }}>—</span>}</td>
                </tr>
              ))}
              <tr>
                <td style={{ fontWeight: 700 }}>Total réseau</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(totals.stockDebut)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.success }}>+{fmt(totals.receptions)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.danger }}>−{fmt(totals.sorties)} L</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(totals.stockFin)} L</td>
                <td></td>
              </tr>
            </tbody>
          </table>
        </div>
        <div style={{ marginTop: 20, height: 240 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={rows.map((r) => ({ code: r.site.code, Réceptions: r.receptions, Sorties: r.sorties }))} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F3" vertical={false} />
              <XAxis dataKey="code" tick={{ fontSize: 11, fill: C.sub }} axisLine={{ stroke: C.border }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: C.sub }} axisLine={false} tickLine={false} />
              <Tooltip formatter={(v) => `${fmt(v)} L`} contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${C.border}` }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="Réceptions" fill={C.success} radius={[4, 4, 0, 0]} />
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
    const sorties = sumQty(rangeMovs, ["sortie", "sortie_camion"]);
    const ajustement = rangeMovs.filter((m) => m.type === "ajustement").reduce((a, m) => a + m.delta, 0);
    const stockFin = stockDebut + receptions - sorties + ajustement;
    const invCount = inventaires.filter((i) => i.siteId === s.id && i.date >= bounds.start && i.date <= bounds.end).length;
    return { site: s, stockDebut, receptions, sorties, stockFin, invCount };
  });

  const doExcel = () => exportToExcel(`SOMIP_Rapport_Decadaire_${bounds.start}_${bounds.end}.xlsx`, [{
    name: "Décadaire", rows: rows.map((r) => ({
      Site: r.site.name, "Stock début (L)": Math.round(r.stockDebut), "Réceptions (L)": Math.round(r.receptions),
      "Sorties (L)": Math.round(r.sorties), "Stock fin (L)": Math.round(r.stockFin), "Inventaires réalisés": r.invCount,
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
            <thead><tr><th>Site</th><th style={{ textAlign: "right" }}>Stock début décade</th><th style={{ textAlign: "right" }}>Réceptions</th><th style={{ textAlign: "right" }}>Sorties</th><th style={{ textAlign: "right" }}>Stock fin décade</th><th style={{ textAlign: "right" }}>Inventaires</th></tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(r.stockDebut)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>+{fmt(r.receptions)} L</td>
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
            <BarChart data={rows.map((r) => ({ code: r.site.code, Réceptions: r.receptions, Sorties: r.sorties }))} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F3" vertical={false} />
              <XAxis dataKey="code" tick={{ fontSize: 11, fill: C.sub }} axisLine={{ stroke: C.border }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: C.sub }} axisLine={false} tickLine={false} />
              <Tooltip formatter={(v) => `${fmt(v)} L`} contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${C.border}` }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="Réceptions" fill={C.success} radius={[4, 4, 0, 0]} />
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
    const sorties = sumQty(rangeMovs, ["sortie", "sortie_camion"]);
    const stockFin = stockThroughDate(s, movements, bounds.end);
    const monthInv = inventaires.filter((i) => i.siteId === s.id && i.date >= bounds.start && i.date <= bounds.end);
    const ecartCumule = monthInv.reduce((a, i) => a + i.ecart, 0);
    const pertes = monthInv.filter((i) => i.nature === "perte");
    const tauxMoyen = pertes.length ? pertes.reduce((a, i) => a + i.tauxFreinte, 0) / pertes.length : null;
    return { site: s, stockDebut, receptions, sorties, stockFin, nbInv: monthInv.length, ecartCumule, tauxMoyen };
  });

  const doExcel = () => exportToExcel(`SOMIP_Rapport_Mensuel_${month}.xlsx`, [{
    name: "Mensuel", rows: rows.map((r) => ({
      Site: r.site.name, "Stock début mois (L)": Math.round(r.stockDebut), "Réceptions (L)": Math.round(r.receptions),
      "Sorties (L)": Math.round(r.sorties), "Stock fin mois (L)": Math.round(r.stockFin), "Nb inventaires": r.nbInv,
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
            <thead><tr><th>Site</th><th style={{ textAlign: "right" }}>Stock début</th><th style={{ textAlign: "right" }}>Réceptions</th><th style={{ textAlign: "right" }}>Sorties</th><th style={{ textAlign: "right" }}>Stock fin</th><th style={{ textAlign: "right" }}>Nb inv.</th><th style={{ textAlign: "right" }}>Écart cumulé</th><th style={{ textAlign: "right" }}>Freinte moy.</th></tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{fmt(r.stockDebut)} L</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.success }}>+{fmt(r.receptions)} L</td>
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
function StockStatementReport({ sites, movements }) {
  const [date, setDate] = useState(todayStr());
  const rows = sites.map((s) => {
    const stock = stockThroughDate(s, movements, date);
    const pct = s.capacity ? (stock / s.capacity) * 100 : 0;
    const status = pct < 20 ? "danger" : pct < 35 ? "warning" : "ok";
    return { site: s, stock, pct, status };
  });
  const statusColor = { ok: C.blue, warning: C.warning, danger: C.danger };
  const statusLabel = { ok: "Normal", warning: "Vigilance", danger: "Critique" };
  const total = rows.reduce((a, r) => a + r.stock, 0);
  const totalCap = rows.reduce((a, r) => a + r.site.capacity, 0);

  const doExcel = () => exportToExcel(`SOMIP_Etat_Stocks_${date}.xlsx`, [{
    name: "Etat des stocks", rows: rows.map((r) => ({
      Site: r.site.name, "Stock (L)": Math.round(r.stock), "Capacité (L)": r.site.capacity,
      "Taux de remplissage (%)": r.pct.toFixed(1), Statut: statusLabel[r.status],
    })),
  }]);

  return (
    <div>
      <div className="somip-no-print" style={{ marginBottom: 14 }}>
        <Field label="Date de l'état des stocks"><input type="date" className="somip-input" style={{ maxWidth: 220 }} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
      </div>
      <div className="somip-print-area somip-panel" style={{ padding: 18 }}>
        <ReportHeader title="État journalier des stocks" period={`Situation au ${date}`} />
        <ReportToolbar onExcel={doExcel} onPrint={() => window.print()} />
        <div style={{ overflowX: "auto" }}>
          <table className="somip-table">
            <thead><tr><th>Site</th><th style={{ textAlign: "right" }}>Stock (L)</th><th style={{ textAlign: "right" }}>Capacité (L)</th><th style={{ textAlign: "right" }}>Remplissage</th><th>Statut</th></tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.site.id}>
                  <td style={{ fontWeight: 600 }}>{r.site.name} <span style={{ color: C.sub, fontWeight: 500 }}>({r.site.code})</span></td>
                  <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(r.stock)}</td>
                  <td className="somip-mono" style={{ textAlign: "right", color: C.sub }}>{fmt(r.site.capacity)}</td>
                  <td className="somip-mono" style={{ textAlign: "right" }}>{r.pct.toFixed(1)} %</td>
                  <td><Badge color={statusColor[r.status]}>{statusLabel[r.status]}</Badge></td>
                </tr>
              ))}
              <tr>
                <td style={{ fontWeight: 700 }}>Total réseau</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{fmt(total)}</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700, color: C.sub }}>{fmt(totalCap)}</td>
                <td className="somip-mono" style={{ textAlign: "right", fontWeight: 700 }}>{totalCap ? ((total / totalCap) * 100).toFixed(1) : 0} %</td>
                <td></td>
              </tr>
            </tbody>
          </table>
        </div>
        <div style={{ marginTop: 20, height: 240 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={rows.map((r) => ({ code: r.site.code, Stock: r.stock, Capacité: r.site.capacity }))} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F3" vertical={false} />
              <XAxis dataKey="code" tick={{ fontSize: 11, fill: C.sub }} axisLine={{ stroke: C.border }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: C.sub }} axisLine={false} tickLine={false} />
              <Tooltip formatter={(v) => `${fmt(v)} L`} contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${C.border}` }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="Capacité" fill="#DCE3E9" radius={[4, 4, 0, 0]} />
              <Bar dataKey="Stock" fill={C.blue} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
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
  const receptionsMonth = movements.filter((m) => m.type === "reception" && m.date.startsWith(month)).reduce((a, m) => a + m.quantity, 0);
  const sortiesMonth = movements.filter((m) => (m.type === "sortie" || m.type === "sortie_camion") && m.date.startsWith(month)).reduce((a, m) => a + m.quantity, 0);
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
function UsersView({ users, addUser, editUser, removeUser }) {
  const [name, setName] = useState("");
  const [role, setRole] = useState(ROLES[0]);
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});

  const submit = () => { if (!name.trim()) return; addUser(name, role); setName(""); setRole(ROLES[0]); };
  const startEdit = (u) => { setEditingId(u.id); setEditForm({ ...u }); };
  const saveEdit = () => { editUser(editingId, editForm); setEditingId(null); };

  return (
    <div className="somip-fade" style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
      <div className="somip-panel" style={{ flex: "2 1 480px", padding: 18 }}>
        <h3 style={{ margin: "0 0 4px", fontSize: 14 }}>Utilisateurs ({users.length})</h3>
        <p style={{ margin: "0 0 14px", fontSize: 12.5, color: C.sub }}>
          Identification simple pour attribuer les actions dans l'historique — ce n'est pas encore une authentification sécurisée (pas de mot de passe).
        </p>
        <table className="somip-table">
          <thead><tr><th>Nom</th><th>Rôle</th><th></th></tr></thead>
          <tbody>
            {users.map((u) => {
              const isEditing = editingId === u.id;
              return (
                <tr key={u.id}>
                  {isEditing ? (
                    <>
                      <td><input className="somip-input" value={editForm.name} onChange={(e) => setEditForm({ ...editForm, name: e.target.value })} /></td>
                      <td>
                        <select className="somip-select" value={editForm.role} onChange={(e) => setEditForm({ ...editForm, role: e.target.value })}>
                          {ROLES.map((r) => <option key={r} value={r}>{r}</option>)}
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
                      <td><Badge color={C.blue}>{u.role}</Badge></td>
                      <td style={{ whiteSpace: "nowrap", textAlign: "right" }}>
                        <button onClick={() => startEdit(u)} style={{ border: "none", background: "none", cursor: "pointer", padding: 5 }}><Pencil size={14} color={C.sub} /></button>
                        <ConfirmIconButton onConfirm={() => removeUser(u)} title="Supprimer l'utilisateur" />
                      </td>
                    </>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div className="somip-panel" style={{ flex: "1 1 260px", padding: 18 }}>
        <h3 style={{ margin: "0 0 14px", fontSize: 14 }}>Ajouter un utilisateur</h3>
        <Field label="Nom complet"><input className="somip-input" value={name} onChange={(e) => setName(e.target.value)} placeholder="Ex : Jean Mabiala" /></Field>
        <Field label="Rôle">
          <select className="somip-select" value={role} onChange={(e) => setRole(e.target.value)}>
            {ROLES.map((r) => <option key={r} value={r}>{r}</option>)}
          </select>
        </Field>
        <button className="somip-btn somip-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={submit} disabled={!name.trim()}>
          <Plus size={15} /> Ajouter
        </button>
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
