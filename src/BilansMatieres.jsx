// =====================================================================
// SOMIP Stock — Bilans matières (refonte style SAP) — ÉTAPE 2
//   • écran de sélection (Site / Mois / Année / Exécuter), comme une transaction SAP
//   • en-tête du bilan, onglets, barre d'outils, barre d'état
//   • onglet Synthèse (compte de stock) et onglet Réceptions (reprises du registre)
// Les onglets « Ventes journalières » et « Stocks début / fin » arrivent aux étapes suivantes.
//
// Économie Supabase : seules 3 petites requêtes partent à l'ouverture d'un bilan (le bilan du
// mois, le bilan du mois précédent, la densité de vente). Les réceptions et transferts viennent
// des mouvements déjà chargés par l'application : aucune requête supplémentaire.
// =====================================================================
import React, { useState, useMemo, useRef, useEffect, forwardRef, useImperativeHandle } from "react";
import { supabase } from "./supabaseClient.js";
import { receptionsDuMois, calculerSynthese, lignePresentation, construireDocumentBilan, calculerVente, calculerJourVentes, totauxParLigne, totauxVentes, joursDuMois, convertir15, lireNombre, decadesDuMois, calculerDecade, totauxDecades, transfertsParDecade, messagesVentes, messagesTransferts, messagesGeneraux, trierMessages, compterMessages, TEMP_MIN_PLAUSIBLE, TEMP_MAX_PLAUSIBLE, DENSITE_VENTE_DEFAUT } from "./bilanCalc.js";

const MOIS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];
const pad2 = (n) => String(n).padStart(2, "0");
const COLS = "id, site_id, period_type, period_key, statut, reception15, ventes15, transferts15, stock_fin15, stock_debut_ambiant, stock_debut_temp, stock_debut_densite, stock_debut15, stock_fin_ambiant, stock_fin_temp, stock_fin_densite, densite_vente, ventes_ambiant, commentaire, created_by, created_at, modified_by, modified_at, valide_par, valide_le";
const COLS_PREC = "id, statut, period_key, stock_fin15, stock_fin_ambiant, stock_fin_temp, stock_fin_densite";
const COLS_PRES = COLS + ", photo_urls";
const COLS_STOCKS = "stock_debut_ambiant, stock_debut_temp, stock_debut_densite, stock_debut15, stock_fin_ambiant, stock_fin_temp, stock_fin_densite, stock_fin15, modified_at, modified_by";
const STATUTS = {
  brouillon: { label: "Brouillon", cls: "sap-st-draft" },
  valide: { label: "Validé 🔒", cls: "sap-st-ok" },
  historique: { label: "Historique (ancien module)", cls: "sap-st-hist" },
};
const fmtL = (n) => (Number.isFinite(n) ? Math.round(n).toLocaleString("fr-FR") : "—");
const fmtSigned = (n) => (Number.isFinite(n) ? `${n > 0 ? "+" : ""}${Math.round(n).toLocaleString("fr-FR")}` : "—");
const fmtDateHeure = (iso) => (iso ? new Date(iso).toLocaleString("fr-FR", { dateStyle: "short", timeStyle: "short" }) : "—");
const fmtJour = (d) => (d ? `${d.slice(8, 10)}/${d.slice(5, 7)}/${d.slice(0, 4)}` : "—");
const num = (v) => (v === null || v === undefined || v === "" ? null : Number(v));
const clsSigne = (n) => (n > 0 ? "sap-gain" : n < 0 ? "sap-loss" : "");

// ---------------------------------------------------------------------
// Aspect (style Fiori) : petites icônes en SVG et surbrillance des jours fériés. Aucun calcul métier ici.
// ---------------------------------------------------------------------
const ICONES = {
  save: ["M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z", "M17 21v-8H7v8", "M7 3v5h8"],
  check: ["M20 6 9 17l-5-5"],
  shield: ["M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z", "m9 12 2 2 4-4"],
  back: ["m12 19-7-7 7-7", "M19 12H5"],
  sheet: ["M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z", "M14 2v6h6", "M8 13h8", "M8 17h8"],
  pdf: ["M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z", "M14 2v6h6", "M8 13h5", "M8 17h8"],
  print: ["M6 9V2h12v7", "M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2", "M6 14h12v8H6z"],
  search: ["M11 19a8 8 0 1 0 0-16 8 8 0 0 0 0 16z", "m21 21-4.35-4.35"],
  undo: ["M3 12a9 9 0 1 0 3-6.7L3 8", "M3 3v5h5"],
  play: ["m6 3 14 9-14 9z"],
  slides: ["M2 3h20", "M21 3v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V3", "m7 21 5-5 5 5"],
  db: ["M12 8c4.97 0 9-1.34 9-3s-4.03-3-9-3-9 1.34-9 3 4.03 3 9 3z", "M3 5v14c0 1.66 4.03 3 9 3s9-1.34 9-3V5", "M3 12c0 1.66 4.03 3 9 3s9-1.34 9-3"],
  down: ["M12 3v14", "m6 11 6 6 6-6", "M5 21h14"],
  up: ["M12 21V7", "m6 13 6-6 6 6", "M5 3h14"],
  swap: ["M7 16V4", "m3 8 4-4 4 4", "M17 8v12", "m21 16-4 4-4-4"],
  chart: ["M18 20V10", "M12 20V4", "M6 20v-6"],
  gain: ["M3 17l6-6 4 4 8-8", "M17 7h4v4"],
  perte: ["M3 7l6 6 4-4 8 8", "M17 17h4v-4"],
  egal: ["M5 9h14", "M5 15h14"],
  list: ["M8 6h13", "M8 12h13", "M8 18h13", "M3 6h.01", "M3 12h.01", "M3 18h.01"],
  edit: ["M12 20h9", "M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4z"],
};
const Ico = ({ n, s = 16 }) => (
  <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" focusable="false">
    {(ICONES[n] || []).map((d, i) => <path key={i} d={d} />)}
  </svg>
);
// Jours fériés du Gabon (surbrillance seulement) : dates fixes + lundi de Pâques, Ascension, lundi de Pentecôte.
// Les fêtes musulmanes (dates lunaires) ne sont pas incluses.
function joursFeries(annee) {
  const f = new Set([`${annee}-01-01`, `${annee}-05-01`, `${annee}-08-15`, `${annee}-08-16`, `${annee}-08-17`, `${annee}-11-01`, `${annee}-12-25`]);
  const a = annee % 19, b = Math.floor(annee / 100), cc = annee % 100, d = Math.floor(b / 4), e = b % 4, g0 = Math.floor((b + 8) / 25), g = Math.floor((b - g0 + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30, i = Math.floor(cc / 4), k = cc % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mois = Math.floor((h + l - 7 * m + 114) / 31), jour = ((h + l - 7 * m + 114) % 31) + 1;
  const paques = new Date(annee, mois - 1, jour, 12);
  const plus = (n) => { const x = new Date(paques); x.setDate(x.getDate() + n); return `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, "0")}-${String(x.getDate()).padStart(2, "0")}`; };
  f.add(plus(1)); f.add(plus(39)); f.add(plus(50));
  return f;
}

// Message lisible pour une erreur Supabase (script SQL de l'étape 1 non exécuté, droits, réseau…).
function erreurLisible(e) {
  const t = `${e?.message || ""} ${e?.code || ""}`;
  // Uniquement les vraies erreurs « colonne ou table absente » : le mot « statut » seul apparaît aussi dans le message
  // légitime de verrouillage (« Bilan verrouillé (statut : valide) … ») qui ne doit pas être pris pour une base obsolète.
  if (/does not exist|42703|42P01|PGRST205|schema cache/i.test(t)) {
    const script = /bilan_transferts/.test(t) ? "supabase_bilan_sap_transferts.sql" : /bilan_ventes_lignes|categorie/.test(t) ? "supabase_bilan_sap_lignes_vente.sql" : "supabase_bilan_sap_etape1.sql";
    return `La base n'est pas à jour : exécute le script ${script} dans Supabase.`;
  }
  if (/row-level security|permission denied|42501/i.test(t)) return "Droits insuffisants pour cette opération (site non attribué à ton compte ?).";
  if (/Failed to fetch|NetworkError|network/i.test(t)) return "Connexion impossible : vérifie ta connexion internet puis réessaie.";
  return e?.message || "Erreur inconnue.";
}

const STYLE = `
.sap-root { --navy: #0B2F52; --navy-2: #16466F; --blue: #0070B8; --red: #D43F0A; --red-logo: #F06810; --ink: #1B2A3B; --muted: #5B6B7E; --line: #DDE4EC; --soft: #F4F7FA; --ok: #0F7B38; --err: #C0262D; --radius: 12px; --shadow: 0 1px 2px rgba(11, 47, 82, .06), 0 6px 18px rgba(11, 47, 82, .08);
  font-family: Inter, "Segoe UI Variable Text", "Segoe UI", system-ui, -apple-system, "Helvetica Neue", Roboto, Arial, sans-serif; font-size: 14px; line-height: 1.45; color: var(--ink); background: var(--soft); border: 1px solid var(--line); border-radius: 16px; overflow: hidden; box-shadow: var(--shadow); }
.sap-root button, .sap-root input, .sap-root select, .sap-root textarea { font-family: inherit; }
/* Bandeau de titre */
.sap-shell { background: linear-gradient(135deg, var(--navy) 0%, var(--navy-2) 100%); color: #fff; padding: 20px 28px 18px; display: flex; justify-content: space-between; align-items: center; gap: 16px; flex-wrap: wrap; border-bottom: 4px solid var(--red-logo); }
.sap-shell-id { display: flex; flex-direction: column; gap: 3px; min-width: 0; }
.sap-kicker { font-size: 11px; letter-spacing: .14em; text-transform: uppercase; color: #A9C4DE; font-weight: 600; }
.sap-title { margin: 0; font-size: clamp(21px, 3.4vw, 30px); font-weight: 700; letter-spacing: -.01em; line-height: 1.15; }
.sap-title small { font-weight: 500; font-size: .6em; color: #CFE0F0; margin-left: 12px; letter-spacing: 0; }
.sap-shell > span { font-size: 15px; font-weight: 600; }
/* Barre d'outils et boutons */
.sap-toolbar { display: flex; gap: 8px; flex-wrap: wrap; padding: 14px 28px; background: #fff; border-bottom: 1px solid var(--line); align-items: center; }
.sap-btn { display: inline-flex; align-items: center; gap: 7px; background: #fff; border: 1px solid #C5D0DC; border-radius: 9px; padding: 8px 14px; font-size: 13px; font-weight: 600; color: var(--navy); cursor: pointer; line-height: 1.2; transition: background .15s, border-color .15s, box-shadow .15s, transform .05s; }
.sap-btn svg { flex: none; }
.sap-btn:hover:not(:disabled) { background: #EEF3F8; border-color: #9FB2C6; }
.sap-btn:active:not(:disabled) { transform: translateY(1px); }
.sap-btn:focus-visible { outline: 2px solid var(--red-logo); outline-offset: 2px; }
.sap-btn:disabled { color: #9AA8B8; background: #F6F8FA; border-color: #E3E8EE; cursor: not-allowed; }
.sap-btn-ghost { border-color: transparent; background: transparent; }
.sap-btn-main { background: var(--red); border-color: var(--red); color: #fff; box-shadow: 0 2px 6px rgba(212, 63, 10, .28); }
.sap-btn-main:hover:not(:disabled) { background: #B93508; border-color: #B93508; }
.sap-btn-main:disabled { background: #EBC1B0; border-color: #EBC1B0; color: #fff; box-shadow: none; }
.sap-btn-ok { background: var(--ok); border-color: var(--ok); color: #fff; box-shadow: 0 2px 6px rgba(15, 123, 56, .25); }
.sap-btn-ok:hover:not(:disabled) { background: #0B6A2F; border-color: #0B6A2F; }
.sap-btn-ok:disabled { background: #AAD3B9; border-color: #AAD3B9; color: #fff; box-shadow: none; }
/* Formulaires (écran de sélection) */
.sap-form { max-width: 600px; margin: 26px auto; padding: 24px 28px; background: #fff; border: 1px solid var(--line); border-radius: 14px; box-shadow: var(--shadow); }
.sap-row { display: grid; grid-template-columns: 160px 1fr; align-items: center; gap: 12px; margin-bottom: 14px; }
.sap-label { color: var(--muted); font-size: 12.5px; font-weight: 600; }
.sap-input { border: 1px solid #C5D0DC; border-radius: 9px; padding: 9px 12px; font-size: 14px; background: #fff; width: 100%; box-sizing: border-box; color: var(--ink); }
.sap-input:focus { outline: none; border-color: var(--red-logo); box-shadow: 0 0 0 3px rgba(240, 104, 16, .22); }
.sap-input:disabled { background: #F1F4F8; color: #6C7B8D; }
/* Informations du bilan */
.sap-header { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 10px 28px; padding: 14px 28px; background: var(--soft); border-bottom: 1px solid var(--line); }
.sap-hrow { display: flex; flex-direction: column; gap: 2px; font-size: 13px; }
.sap-hrow b { font-size: 12px; font-variant-caps: all-small-caps; letter-spacing: .07em; color: var(--muted); font-weight: 700; }
/* Badges de statut */
.sap-badge { display: inline-flex; align-items: center; gap: 7px; padding: 5px 13px; border-radius: 999px; font-size: 12.5px; font-weight: 700; width: fit-content; }
.sap-badge::before { content: ""; width: 8px; height: 8px; border-radius: 50%; background: currentColor; opacity: .85; }
.sap-st-draft { background: #FFE9D0; color: #8A3F00; }
.sap-st-ok { background: #DCF3E3; color: #0B5A2A; }
.sap-st-reopen { background: #DDEBFB; color: #0B4A86; }
.sap-st-hist { background: #E7EBF0; color: #4A5868; }
.sap-banner { padding: 10px 28px; background: #FFF4D6; color: #6B4A00; border-bottom: 1px solid #F0DC9A; font-size: 13px; }
.sap-banner-ok { padding: 10px 28px; background: #E3F5E9; color: #0B5A2A; border-bottom: 1px solid #BFE3CB; font-size: 13px; }
/* Progression des ventes saisies */
.sap-progress { display: flex; align-items: center; gap: 14px; flex-wrap: wrap; padding: 11px 28px 13px; background: #fff; border-bottom: 1px solid var(--line); }
.sap-progress-txt { font-size: 12.5px; font-weight: 700; color: var(--navy); min-width: 150px; }
.sap-progress-bar { flex: 1; min-width: 140px; height: 8px; border-radius: 99px; background: #E4EAF1; overflow: hidden; }
.sap-progress-fill { height: 100%; border-radius: 99px; background: linear-gradient(90deg, var(--red-logo), var(--red)); transition: width .45s ease; }
.sap-progress.done .sap-progress-fill { background: linear-gradient(90deg, #22A455, var(--ok)); }
.sap-progress-pct { font-size: 12px; color: var(--muted); font-weight: 600; min-width: 36px; text-align: right; }
/* Onglets */
.sap-tabs { display: flex; gap: 2px; padding: 0 28px; background: #fff; border-bottom: 1px solid var(--line); overflow-x: auto; flex-wrap: nowrap; }
.sap-tab { display: inline-flex; align-items: center; gap: 8px; border: none; border-bottom: 3px solid transparent; background: transparent; padding: 14px 14px 12px; margin-bottom: -1px; cursor: pointer; font-size: 13.5px; font-weight: 600; color: var(--muted); white-space: nowrap; transition: color .15s, background .15s; }
.sap-tab:hover { color: var(--navy); background: #F4F7FA; }
.sap-tab.on { color: var(--navy); border-bottom-color: var(--red-logo); }
.sap-tab:focus-visible { outline: 2px solid var(--red-logo); outline-offset: -2px; }
.sap-body { padding: 24px 28px 28px; background: var(--soft); min-height: 260px; overflow-x: auto; }
/* Cartes d'indicateurs */
.sap-kpis { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 18px; margin-bottom: 24px; }
.sap-kpi { position: relative; display: flex; gap: 14px; align-items: flex-start; background: #fff; border: 1px solid var(--line); border-radius: var(--radius); padding: 18px 18px 16px 22px; box-shadow: var(--shadow); overflow: hidden; }
.sap-kpi::before { content: ""; position: absolute; inset: 0 auto 0 0; width: 5px; background: var(--kc, var(--blue)); }
.sap-kpi-ic { flex: none; width: 44px; height: 44px; border-radius: 13px; display: grid; place-items: center; background: var(--kbg, #E1EEF9); color: var(--kc, var(--blue)); }
.sap-kpi-label { font-size: 11.5px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: .06em; }
.sap-kpi-val { font-size: 27px; font-weight: 700; color: var(--navy); font-variant-numeric: tabular-nums; line-height: 1.15; margin-top: 3px; }
.sap-kpi-val small { font-size: 13px; font-weight: 600; color: var(--muted); margin-left: 3px; }
.sap-kpi-sub { font-size: 12px; color: var(--muted); margin-top: 3px; }
.sap-kpi-debut { --kc: #0B5E9E; --kbg: #E1EEF9; }
.sap-kpi-rec { --kc: #0E7490; --kbg: #DDF1F6; }
.sap-kpi-ven { --kc: #C2410C; --kbg: #FDE6D8; }
.sap-kpi-neutre { --kc: #64748B; --kbg: #E8EDF2; }
.sap-kpi-gain { --kc: var(--ok); --kbg: #DCF3E3; } .sap-kpi-gain .sap-kpi-val { color: var(--ok); }
.sap-kpi-perte { --kc: var(--err); --kbg: #FBE1E2; } .sap-kpi-perte .sap-kpi-val { color: var(--err); }
/* Cartes */
.sap-card { background: #fff; border: 1px solid var(--line); border-radius: var(--radius); box-shadow: var(--shadow); overflow: hidden; margin-bottom: 24px; }
.sap-card-h { display: flex; align-items: center; gap: 9px; padding: 14px 20px; font-weight: 700; color: var(--navy); border-bottom: 1px solid var(--line); background: #FBFCFE; }
.sap-card-b { overflow-x: auto; }
.sap-card-pad { padding: 18px 20px; }
.sap-card-pad .sap-label { display: block; margin-bottom: 8px; font-size: 13.5px; font-weight: 700; color: var(--navy); }
/* Tableaux */
.sap-table { border-collapse: separate; border-spacing: 0; width: 100%; font-size: 13px; }
.sap-table th { background: var(--navy); color: #fff; text-align: left; font-weight: 600; font-size: 12px; letter-spacing: .03em; padding: 11px 14px; white-space: nowrap; }
.sap-table td { padding: 10px 14px; border-bottom: 1px solid #EAEFF4; }
.sap-table tbody tr:nth-child(even) td { background: #F9FBFD; }
.sap-table tbody tr:hover td { background: #EEF4FB; }
.sap-num { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
.sap-table th.sap-num { text-align: right; }
.sap-total td { font-weight: 700; background: #DCE6F1 !important; color: var(--navy); border-top: 2px solid var(--navy); }
.sap-compte th { background: #EAF0F6; color: var(--navy); }
.sap-compte td:first-child { font-weight: 500; }
.sap-compte .sap-final td { font-size: 16px; padding-top: 14px; padding-bottom: 14px; background: #D3E1F0 !important; border-top: 3px solid var(--navy); }
.sap-compte .sap-final .sap-num { font-size: 20px; }
.sap-gain { color: var(--ok); font-weight: 700; }
.sap-loss { color: var(--err); font-weight: 700; }
.sap-muted { color: var(--muted); }
/* Grille de saisie */
.sap-gridwrap { max-height: 64vh; overflow: auto; border: 1px solid var(--line); border-radius: var(--radius); background: #fff; box-shadow: var(--shadow); -webkit-overflow-scrolling: touch; }
.sap-grid { margin: 0; }
.sap-grid thead th { position: sticky; top: 0; z-index: 3; }
.sap-grid tfoot td { position: sticky; bottom: 0; z-index: 3; }
.sap-grid tbody tr { --rb: #fff; }
.sap-grid tbody tr:nth-child(even) { --rb: #F8FAFC; }
.sap-grid tbody tr.sap-off { --rb: #EAF3FC; }
.sap-grid tbody tr:hover { --rb: #E2EDF9; }
.sap-grid tbody td { background: var(--rb); padding: 3px 6px; }
.sap-grid tbody tr:nth-child(even) td, .sap-grid tbody tr:hover td { background: var(--rb); }
.sap-grid td.sap-d { padding: 6px 14px; white-space: nowrap; position: sticky; left: 0; z-index: 2; background: var(--rb); font-weight: 500; }
.sap-grid thead th:first-child { position: sticky; left: 0; z-index: 4; }
.sap-grid tfoot td:first-child { left: 0; z-index: 4; }
.sap-grid tfoot td { background: #DCE6F1 !important; }
.sap-grid tfoot td.sap-d { font-weight: 700; }
.sap-cell { width: 100%; min-width: 88px; box-sizing: border-box; border: 1px solid #D3DCE6; border-radius: 7px; background: #fff; padding: 6px 10px; font-size: 13px; text-align: right; color: var(--ink); font-variant-numeric: tabular-nums; margin: 2px 0; transition: border-color .12s, box-shadow .12s; }
.sap-cell:hover:not(:disabled) { border-color: #9FB2C6; }
.sap-cell:focus { outline: none; border-color: var(--red); box-shadow: 0 0 0 2px rgba(212, 63, 10, .38); background: #fff; }
.sap-cell:disabled { background: #F1F4F8; color: #4A596B; border-color: #E3E8EE; }
.sap-cell-err { border-color: var(--err); background: #FDECEC; }
.sap-cell-warn { background: #FFF6DB; border-color: #E8C35A; }
.sap-cell-prop { font-style: italic; color: var(--muted); }
.sap-ro { background: #F1F4F8 !important; color: #3E4C5E; font-style: italic; font-weight: 400; padding: 6px 14px !important; }
.sap-mark { width: 38px; text-align: center; padding: 0 !important; }
.sap-mark[data-etat="vide"]::before { content: ""; display: inline-block; width: 9px; height: 9px; border-radius: 50%; background: #F59E0B; vertical-align: middle; }
.sap-mark[data-etat="avenir"]::before { content: ""; display: inline-block; width: 9px; height: 9px; border-radius: 50%; background: #CBD5E1; vertical-align: middle; }
.sap-mark[data-etat="ok"]::before { content: "\\2713"; color: #16A34A; font-weight: 700; opacity: .8; }
.sap-row-dirty td.sap-d { box-shadow: inset 4px 0 0 var(--blue); }
.sap-today td.sap-d::after { content: ""; display: inline-block; width: 8px; height: 8px; margin-left: 8px; border-radius: 50%; background: var(--red-logo); box-shadow: 0 0 0 3px rgba(240, 104, 16, .22); vertical-align: middle; }
/* Messages */
.sap-status { display: flex; justify-content: space-between; align-items: center; gap: 12px; flex-wrap: wrap; padding: 10px 28px; background: #fff; border-top: 1px solid var(--line); font-size: 12.5px; color: var(--muted); }
.sap-msg { font-weight: 600; }
.sap-msg.sap-msg-ok, .sap-msg.sap-msg-warn, .sap-msg.sap-msg-err { display: inline-block; padding: 5px 14px 5px 7px; border-radius: 999px; animation: sap-in .3s ease-out; }
.sap-msg-ok { color: #0B5A2A; } .sap-msg-warn { color: #8A4200; } .sap-msg-err { color: #9F1D24; }
.sap-msg.sap-msg-ok { background: #E3F5E9; } .sap-msg.sap-msg-warn { background: #FFF0D9; } .sap-msg.sap-msg-err { background: #FCE4E4; }
.sap-msg-ic { display: inline-grid; place-items: center; width: 20px; height: 20px; border-radius: 50%; margin-right: 8px; font-size: 11px; font-weight: 800; color: #fff; vertical-align: middle; }
.sap-msg-ok .sap-msg-ic { background: #16A34A; } .sap-msg-warn .sap-msg-ic { background: #F59E0B; } .sap-msg-err .sap-msg-ic { background: #DC2626; }
.sap-msg-link { cursor: pointer; text-decoration: underline dotted; }
@keyframes sap-in { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: none; } }
.sap-msgpanel { border-top: 1px solid var(--line); background: #FDFEFF; max-height: 190px; overflow: auto; padding: 12px 28px 14px; font-size: 13px; }
.sap-msghead { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; font-weight: 700; color: var(--navy); }
.sap-msgitem { list-style: none; padding: 6px 10px; margin-bottom: 4px; border-radius: 8px; cursor: pointer; border-left: 4px solid currentColor; background: #fff; }
.sap-msgitem.sap-msg-err { background: #FEF3F3; } .sap-msgitem.sap-msg-warn { background: #FFF8EA; }
.sap-msgitem:hover, .sap-msgitem:focus { filter: brightness(.97); outline: none; }
@media (prefers-reduced-motion: reduce) { .sap-msg, .sap-progress-fill, .sap-btn { animation: none !important; transition: none !important; } }
/* Écrans étroits (portable réduit, tablette, téléphone) */
@media (max-width: 980px) { .sap-kpis { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
@media (max-width: 720px) {
  .sap-shell { padding: 16px 16px 14px; } .sap-toolbar { padding: 10px 14px; gap: 6px; } .sap-btn { padding: 8px 11px; font-size: 12.5px; }
  .sap-header { grid-template-columns: 1fr 1fr; gap: 10px 14px; }
  .sap-header, .sap-progress, .sap-tabs, .sap-body, .sap-status, .sap-banner, .sap-banner-ok, .sap-msgpanel { padding-left: 14px; padding-right: 14px; }
  .sap-body { padding-top: 16px; padding-bottom: 20px; } .sap-form { margin: 14px 10px; padding: 18px 16px; } .sap-row { grid-template-columns: 1fr; gap: 6px; }
  .sap-kpis { gap: 10px; } .sap-kpi { padding: 14px 12px 12px 18px; gap: 10px; } .sap-kpi-ic { width: 38px; height: 38px; } .sap-kpi-val { font-size: 21px; }
  .sap-gridwrap { max-height: 70vh; } .sap-table th, .sap-table td { padding-left: 10px; padding-right: 10px; }
}
@media (max-width: 460px) { .sap-progress-pct { display: none; } .sap-kpis { grid-template-columns: 1fr; } .sap-title small { display: block; margin: 2px 0 0; } }
/* Impression */
.sap-printonly { display: none; }
.sap-printonly h2 { margin: 0 0 2px; font-size: 18px; } .sap-printonly h3 { margin: 12px 0 4px; font-size: 13px; background: #1f3a5f; color: #fff; padding: 3px 8px; }
.sap-printonly table { border-collapse: collapse; width: 100%; font-size: 10.5px; margin-bottom: 6px; }
.sap-printonly th, .sap-printonly td { border: 1px solid #b9c0c7; padding: 2px 6px; text-align: right; } .sap-printonly th:first-child, .sap-printonly td:first-child { text-align: left; }
.sap-printonly thead th { background: #eef1f4; } .sap-printonly tfoot td { font-weight: 700; background: #f6f8f9; }
@media print { .sap-root > *:not(.sap-printonly):not(style) { display: none !important; } .sap-printonly { display: block !important; } .sap-root { border: none !important; box-shadow: none !important; } }
`;


// ---------------------------------------------------------------------
// Grille des ventes journalières (une ligne par jour du mois, éditable comme une grille ALV)
//   Saisie : volume (L ambiants) + température du jour. La densité de vente est fixe (grisée).
//   Calculé : densité à 15 °C, coefficient, volume à 15 °C (même fonction que « Correction 15 °C »).
//   Clavier : Tab / Entrée = cellule suivante ; flèches haut/bas = ligne précédente / suivante.
//   Collage : Ctrl+V d'une plage copiée dans Excel (1 ou 2 colonnes : volume, température).
//   Enregistrement : seules les lignes modifiées sont envoyées ; une ligne vidée est supprimée.
// ---------------------------------------------------------------------
const libLigne = (l) => (l === "" ? "Ventes" : l);
const VentesGrille = forwardRef(function VentesGrille({ bilan, editable, userName, siteName, setMsg, onDirty, onStats, onSaved, onParLigne, appendAudit, flash }, ref) {
  const jours = useMemo(() => joursDuMois(bilan.period_key), [bilan.period_key]);
  const densite = num(bilan.densite_vente) || DENSITE_VENTE_DEFAUT;
  // Lignes de vente du site : [""] = une seule ligne « Ventes » ; sinon les libellés configurés (Comilog, GSEZ…), plus les
  // éventuels libellés déjà présents dans les ventes enregistrées mais qui ne sont plus configurés (ils restent visibles et modifiables).
  const [cols, setCols] = useState([""]);
  const [legacy, setLegacy] = useState(false);   // base pas encore mise à jour (colonne « categorie » absente) : une seule ligne par jour
  const [vals, setVals] = useState({});          // jour -> { temp, v: { [ligne]: texte saisi } }
  const [saved, setSaved] = useState({});        // jour -> { temp, v: { [ligne]: nombre enregistré en base } }
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const refs = useRef({});
  const multi = cols.length > 1;
  // Affichage seulement : dimanches / jours fériés en bleu pâle, repère « aujourd'hui », état de chaque ligne (pastille / coche).
  const aujourdhui = useMemo(() => { const dj = new Date(); return `${dj.getFullYear()}-${String(dj.getMonth() + 1).padStart(2, "0")}-${String(dj.getDate()).padStart(2, "0")}`; }, []);
  const feries = useMemo(() => joursFeries(Number(String(bilan.period_key).slice(0, 4))), [bilan.period_key]);
  const estOff = (j) => new Date(`${j}T12:00:00`).getDay() === 0 || feries.has(j);
  const nCol = cols.length + 1;                  // colonnes saisissables : un volume par ligne, puis la température (commune à la journée)

  useEffect(() => {
    let off = false;
    (async () => {
      const cfg = await supabase.from("bilan_ventes_lignes").select("lignes").eq("site_id", bilan.site_id).maybeSingle();
      const configurees = !cfg.error && cfg.data && Array.isArray(cfg.data.lignes) ? cfg.data.lignes.filter(Boolean) : [];
      let r = await supabase.from("bilan_ventes").select("jour, categorie, volume_ambiant, temperature_c").eq("bilan_id", bilan.id).order("jour");
      let ancien = false;
      if (r.error && (r.error.code === "42703" || /categorie/.test(String(r.error.message || "")))) {
        r = await supabase.from("bilan_ventes").select("jour, volume_ambiant, temperature_c").eq("bilan_id", bilan.id).order("jour"); ancien = true;
      }
      if (off) return;
      if (r.error) { setMsg({ type: "err", text: erreurLisible(r.error) }); setLoading(false); return; }
      const ligneDe = (l) => (ancien ? "" : l.categorie || "");
      const base = !ancien && configurees.length >= 2 ? configurees : [""];
      const orphelines = [...new Set((r.data || []).map(ligneDe))].filter((x) => !base.includes(x));
      const v = {}, s = {};
      (r.data || []).forEach((l) => {
        const lab = ligneDe(l);
        if (!v[l.jour]) {
          v[l.jour] = { temp: l.temperature_c === null ? "" : String(l.temperature_c), v: {} };
          s[l.jour] = { temp: l.temperature_c === null ? null : Number(l.temperature_c), v: {} };
        }
        v[l.jour].v[lab] = String(l.volume_ambiant); s[l.jour].v[lab] = Number(l.volume_ambiant);
      });
      setCols([...base, ...orphelines]); setLegacy(ancien); setVals(v); setSaved(s); setLoading(false);
    })();
    return () => { off = true; };
  }, [bilan.id]); // eslint-disable-line react-hooks/exhaustive-deps

  const calc = useMemo(() => Object.fromEntries(jours.map((j) => [j, calculerJourVentes({ temp: vals[j]?.temp, volumes: cols.map((l) => vals[j]?.v?.[l] ?? "") }, densite)])), [jours, vals, cols, densite]);
  const totaux = useMemo(() => totauxVentes(jours.map((j) => calc[j])), [jours, calc]);
  const parLigne = useMemo(() => totauxParLigne(jours.map((j) => calc[j]), cols.length), [jours, calc, cols]);
  const modifiee = (j) => {
    const c = calc[j], s = saved[j];
    if (!c.saisi) return !!s;
    if (!c.ok) return true;
    if (!s || s.temp !== c.temp) return true;
    return cols.some((l, k) => (s.v[l] ?? 0) !== c.volumes[k]);
  };
  const nbModifiees = jours.filter(modifiee).length;
  useEffect(() => { if (onDirty) onDirty(nbModifiees); }, [nbModifiees]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { if (!loading && onStats) onStats(totaux.joursSaisis); }, [loading, totaux.joursSaisis]); // eslint-disable-line react-hooks/exhaustive-deps
  // Détail par ligne pour la Synthèse (valeurs enregistrées) : transmis à l'ouverture de la grille et après chaque enregistrement.
  const emettreParLigne = () => { if (onParLigne && multi) onParLigne({ labels: cols.map(libLigne), ambiant: parLigne.ambiant, volume15: parLigne.volume15 }); };
  useEffect(() => { if (!loading) emettreParLigne(); }, [loading]); // eslint-disable-line react-hooks/exhaustive-deps

  const setTemp = (j, valeur) => setVals((p) => ({ ...p, [j]: { v: {}, ...(p[j] || {}), temp: valeur } }));
  const setVol = (j, lab, valeur) => setVals((p) => ({ ...p, [j]: { temp: "", ...(p[j] || {}), v: { ...(((p[j] || {}).v) || {}), [lab]: valeur } } }));
  const aller = (i, col) => { const el = refs.current[`${i}-${col}`]; if (el) { el.focus(); } };
  const onKeyDown = (e, i, col) => {
    if (e.key === "Enter") { e.preventDefault(); if (col < nCol - 1) aller(i, col + 1); else aller(i + 1, 0); }
    else if (e.key === "ArrowDown") { e.preventDefault(); aller(i + 1, col); }
    else if (e.key === "ArrowUp") { e.preventDefault(); aller(i - 1, col); }
  };
  const onPaste = (e, i, col) => {
    const brut = (e.clipboardData && e.clipboardData.getData("text")) || "";
    const texte = brut.replace(/\r/g, "").replace(/\n+$/, "");
    if (!/[\n\t]/.test(texte)) return; // valeur simple : collage normal
    e.preventDefault();
    const lignes = texte.split("\n").map((l) => l.split("\t"));
    const suivant = { ...vals }; let ignorees = 0, remplies = 0;
    lignes.forEach((cases, k) => {
      const idx = i + k;
      if (idx >= jours.length) { ignorees += 1; return; }
      const cur = suivant[jours[idx]] || { temp: "", v: {} };
      const nv = { temp: cur.temp, v: { ...(cur.v || {}) } };
      cases.forEach((cell, d) => { const c2 = col + d; if (c2 > nCol - 1) return; if (c2 === nCol - 1) nv.temp = cell.trim(); else nv.v[cols[c2]] = cell.trim(); });
      suivant[jours[idx]] = nv;
      remplies += 1;
    });
    setVals(suivant);
    setMsg(ignorees ? { type: "warn", text: `${remplies} jour(s) collé(s) ; ${ignorees} ligne(s) ignorée(s) (au-delà de la fin du mois).` } : { type: "ok", text: `${remplies} jour(s) collé(s) depuis le presse-papiers. Pense à enregistrer.` });
  };

  // Contrôle sans écriture : renvoie le message d'erreur à afficher, ou null.
  const verifier = () => {
    if (!editable) return null;
    const erreurs = jours.filter((j) => calc[j].saisi && !calc[j].ok);
    return erreurs.length ? `${erreurs.length} ligne(s) en erreur (${erreurs.slice(0, 3).map(fmtJour).join(", ")}${erreurs.length > 3 ? "…" : ""}) : enregistrement bloqué.` : null;
  };
  // Enregistrement des seules journées modifiées (une ligne en base par journée ET par ligne de vente).
  // Renvoie { statut: "rien" | "ok" | "erreur", texte }.
  const save = async () => {
    if (!editable || saving) return { statut: "rien" };
    const aSauver = jours.filter((j) => calc[j].saisi && calc[j].ok && modifiee(j));
    const aSupprimer = jours.filter((j) => !calc[j].saisi && saved[j]);
    if (!aSauver.length && !aSupprimer.length) return { statut: "rien" };
    setSaving(true);
    try {
      if (aSauver.length) {
        const payload = aSauver.flatMap((j) => cols.map((lab, k) => ({
          bilan_id: bilan.id, jour: j, ...(legacy ? {} : { categorie: lab }), volume_ambiant: calc[j].volumes[k], temperature_c: calc[j].temp, densite_obs: calc[j].densiteObs,
          densite15: calc[j].densite15, vcf: calc[j].vcf, volume15: calc[j].volumes15[k], updated_by: userName || null,
        })));
        const r = await supabase.from("bilan_ventes").upsert(payload, { onConflict: legacy ? "bilan_id,jour" : "bilan_id,jour,categorie" });
        if (r.error) throw r.error;
      }
      if (aSupprimer.length) {
        const r = await supabase.from("bilan_ventes").delete().eq("bilan_id", bilan.id).in("jour", aSupprimer);
        if (r.error) throw r.error;
      }
      setSaved((p) => { const n = { ...p }; aSauver.forEach((j) => { n[j] = { temp: calc[j].temp, v: Object.fromEntries(cols.map((lab, k) => [lab, calc[j].volumes[k]])) }; }); aSupprimer.forEach((j) => { delete n[j]; }); return n; });
      // Totaux : recalculés par la base ; on relit uniquement les 4 champs du bilan (en-tête et synthèse à jour).
      const h = await supabase.from("bilan_matieres").update({ modified_by: userName || null }).eq("id", bilan.id).select("ventes15, ventes_ambiant, modified_at, modified_by").single();
      if (!h.error && h.data && onSaved) onSaved(h.data);
      emettreParLigne();
      if (flash) flash("Bilan enregistré.");
      if (appendAudit) appendAudit("Ventes du bilan matières enregistrées", `${siteName} — ${bilan.period_key} : ${aSauver.length} ligne(s), ${aSupprimer.length} suppression(s)`);
      return { statut: "ok", texte: `${aSauver.length} ligne(s) enregistrée(s)${aSupprimer.length ? `, ${aSupprimer.length} supprimée(s)` : ""}` };
    } catch (e) {
      return { statut: "erreur", texte: erreurLisible(e) };
    } finally { setSaving(false); }
  };
  useImperativeHandle(ref, () => ({ save, verifier, pret: () => !loading, messages: (aujourdhui) => messagesVentes(jours.map((j) => ({ jour: j, calc: calc[j] })), aujourdhui),
    donnees: () => {
      const lignesD = jours.map((j) => (calc[j].saisi && calc[j].ok ? { jour: j, saisi: true, volume: calc[j].volume, temp: calc[j].temp, densite15: calc[j].densite15, vcf: calc[j].vcf, volume15: calc[j].volume15, ...(multi ? { volumes: calc[j].volumes, volumes15: calc[j].volumes15 } : {}) } : { jour: j, saisi: false }));
      return multi ? { densite, cols: cols.map(libLigne), lignes: lignesD, totaux: { ...totaux, parLigne } } : { densite, lignes: lignesD, totaux };
    } }));

  if (loading) return <p className="sap-muted" data-testid="grille-chargement">Chargement des ventes du mois…</p>;
  return (
    <div>
      {!editable && <p className="sap-muted" style={{ margin: "0 0 8px" }} data-testid="grille-lecture">Bilan verrouillé : la grille est en lecture seule.</p>}
      <div className="sap-gridwrap" data-testid="sap-grille">
        <table className="sap-table sap-grid">
          <thead><tr>
            <th>Date</th>
            {multi ? cols.map((l) => <th key={l} className="sap-num">{libLigne(l)} (L)</th>) : <th className="sap-num">Volume vendu (L) *</th>}
            <th className="sap-num">Température (°C) *</th>
            {multi && <th className="sap-num">Total (L)</th>}
            <th className="sap-num">Densité</th><th className="sap-num">Densité à 15 °C</th><th className="sap-num">Coefficient</th><th className="sap-num">Volume 15 °C (L)</th><th className="sap-mark"></th>
          </tr></thead>
          <tbody>
            {jours.map((j, i) => {
              const c = calc[j]; const err = c.saisi && !c.ok; const warn = c.ok && c.avertissement;
              const jourSem = new Date(`${j}T12:00:00`).toLocaleDateString("fr-FR", { weekday: "short" });
              const cellule = (col, champ, valeur, onChange, testid, lib) => (
                <td key={testid}>
                  <input ref={(el) => { refs.current[`${i}-${col}`] = el; }} className={`sap-cell ${err ? "sap-cell-err" : warn && champ === "temp" ? "sap-cell-warn" : ""}`}
                    value={valeur} disabled={!editable} inputMode="decimal" aria-label={`${lib} ${j}`} data-testid={testid} title={err ? c.erreur : warn && champ === "temp" ? c.avertissement : undefined}
                    onChange={(e) => onChange(e.target.value)} onKeyDown={(e) => onKeyDown(e, i, col)} onPaste={(e) => onPaste(e, i, col)} onFocus={(e) => e.target.select()} />
                </td>
              );
              return (
                <tr key={j} data-jour={j} className={`${modifiee(j) ? "sap-row-dirty" : ""} ${estOff(j) ? "sap-off" : ""} ${j === aujourdhui ? "sap-today" : ""}`.trim()}>
                  <td className="sap-d" title={j === aujourdhui ? "Aujourd'hui" : estOff(j) ? "Dimanche ou jour férié" : undefined}>{fmtJour(j)} <span className="sap-muted">{jourSem}</span></td>
                  {cols.map((lab, k) => cellule(k, "volume", vals[j]?.v?.[lab] ?? "", (x) => setVol(j, lab, x), k === 0 ? `volume-${j}` : `volume${k}-${j}`, multi ? `Volume ${libLigne(lab)}` : "Volume"))}
                  {cellule(nCol - 1, "temp", vals[j]?.temp ?? "", (x) => setTemp(j, x), `temp-${j}`, "Température")}
                  {multi && <td className="sap-num sap-ro" data-testid={`tot-${j}`}>{c.saisi && c.ok ? fmtL(c.volume) : ""}</td>}
                  <td className="sap-num sap-ro" data-testid={`dens-${j}`}>{fmtL(densite)}</td>
                  <td className="sap-num sap-ro" data-testid={`d15-${j}`}>{c.saisi && c.ok && c.densite15 !== null && c.densite15 !== undefined ? c.densite15.toLocaleString("fr-FR", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) : ""}</td>
                  <td className="sap-num sap-ro" data-testid={`vcf-${j}`}>{c.saisi && c.ok && c.vcf ? c.vcf.toFixed(4).replace(".", ",") : ""}</td>
                  <td className="sap-num sap-ro" data-testid={`v15-${j}`}>{c.saisi && c.ok ? fmtL(c.volume15) : ""}</td>
                  <td className="sap-mark" data-etat={err ? "err" : warn ? "warn" : c.saisi && c.ok ? "ok" : j > aujourdhui ? "avenir" : "vide"} title={err ? c.erreur : warn ? c.avertissement : c.saisi ? "Journée saisie" : j > aujourdhui ? "" : "Journée non saisie"} data-testid={`mark-${j}`}>{err ? <span className="sap-msg-err">✖</span> : warn ? <span className="sap-msg-warn">⚠</span> : ""}</td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr className="sap-total">
              <td className="sap-d" data-testid="grille-jours">Total ({totaux.joursSaisis} / {jours.length} jours)</td>
              {multi ? cols.map((l, k) => <td key={l} className="sap-num" data-testid={`grille-total-ligne${k}`}>{fmtL(parLigne.ambiant[k])}</td>) : <td className="sap-num" data-testid="grille-total-ambiant">{fmtL(totaux.ambiant)}</td>}
              <td></td>
              {multi && <td className="sap-num" data-testid="grille-total-ambiant">{fmtL(totaux.ambiant)}</td>}
              <td></td><td></td><td></td>
              <td className="sap-num" data-testid="grille-total15">{fmtL(totaux.volume15)}</td><td></td>
            </tr>
          </tfoot>
        </table>
      </div>
      <p className="sap-muted" style={{ fontSize: 11.5, margin: "8px 0 0" }}>Tab / Entrée : cellule suivante · ↑ ↓ : changer de ligne · Ctrl+V : coller une plage depuis Excel ({multi ? "un volume par ligne ⇥ température" : "volume ⇥ température"}) · {multi ? "une seule température par jour, commune à toutes les lignes ; un volume vide vaut 0 · " : ""}densité de vente fixe : {fmtL(densite)} kg/m³.</p>
    </div>
  );
});


// ---------------------------------------------------------------------
// Onglet « Transferts » : transferts ENTRE SITES, saisis PAR DÉCADE, à l'AMBIANT, uniquement pour équilibrer les stocks.
//   Décade 1 = du 01 au 10 · décade 2 = du 11 au 20 · décade 3 = du 21 à la fin du mois.
//   « Envoyés » (à un autre site) sortent du stock ; « Reçus » (d'un autre site) y entrent. Le site concerné n'est pas saisi.
//   Aucune température ni conversion à 15 °C : seules les ventes sont converties (onglet Ventes journalières).
//   Les deux champs sont facultatifs ; une décade vide est ignorée. Enregistrement : seules les décades modifiées.
//   Pour information (jamais compté) : ce que l'application a déjà enregistré (transferts entre sites, chargements de camions).
// ---------------------------------------------------------------------
const TransfertsDecades = forwardRef(function TransfertsDecades({ bilan, editable, userName, siteName, movements, setMsg, onDirty, onStats, onSaved, appendAudit, flash }, ref) {
  const decades = useMemo(() => decadesDuMois(bilan.period_key), [bilan.period_key]);
  const info = useMemo(() => transfertsParDecade(movements, bilan.site_id, bilan.period_key), [movements, bilan.site_id, bilan.period_key]);
  const [vals, setVals] = useState({});   // n -> { transferts, retours } (texte saisi)
  const [saved, setSaved] = useState({}); // n -> { transferts, retours } (nombres enregistrés en base)
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const refs = useRef({});

  useEffect(() => {
    let off = false;
    (async () => {
      const r = await supabase.from("bilan_transferts").select("decade, transferts_ambiant, retours_ambiant").eq("bilan_id", bilan.id).order("decade");
      if (off) return;
      if (r.error) { setMsg({ type: "err", text: erreurLisible(r.error) }); setLoading(false); return; }
      const v = {}, s = {};
      (r.data || []).forEach((l) => {
        v[l.decade] = { envoyes: String(l.transferts_ambiant), recus: String(l.retours_ambiant) };
        s[l.decade] = { envoyes: Number(l.transferts_ambiant), recus: Number(l.retours_ambiant) };
      });
      setVals(v); setSaved(s); setLoading(false);
    })();
    return () => { off = true; };
  }, [bilan.id]); // eslint-disable-line react-hooks/exhaustive-deps

  const calc = useMemo(() => Object.fromEntries(decades.map((d) => [d.n, calculerDecade(vals[d.n] || {})])), [decades, vals]);
  const tot = useMemo(() => totauxDecades(decades.map((d) => calc[d.n])), [decades, calc]);
  const modifiee = (n) => {
    const c = calc[n], s = saved[n];
    if (!c.saisi) return !!s;
    if (!c.ok) return true;
    return !s || s.envoyes !== c.envoyes || s.recus !== c.recus;
  };
  const nbModifiees = decades.filter((d) => modifiee(d.n)).length;
  useEffect(() => { if (onDirty) onDirty(nbModifiees); }, [nbModifiees]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => { if (onStats) onStats(tot.nb); }, [tot.nb]); // eslint-disable-line react-hooks/exhaustive-deps

  const champs = ["envoyes", "recus"];
  const setCell = (n, champ, valeur) => setVals((p) => ({ ...p, [n]: { ...(p[n] || {}), [champ]: valeur } }));
  const aller = (i, col) => { const el = refs.current[`${i}-${col}`]; if (el) el.focus(); };
  const onKeyDown = (e, i, col) => {
    if (e.key === "Enter") { e.preventDefault(); if (col === 0) aller(i, 1); else aller(i + 1, 0); }
    else if (e.key === "ArrowDown") { e.preventDefault(); aller(i + 1, col); }
    else if (e.key === "ArrowUp") { e.preventDefault(); aller(i - 1, col); }
  };

  const verifier = () => {
    if (!editable) return null;
    const erreurs = decades.filter((d) => calc[d.n].saisi && !calc[d.n].ok);
    return erreurs.length ? `Transferts : ${erreurs.length} décade(s) en erreur (${erreurs.map((d) => `décade ${d.n}`).join(", ")}) : enregistrement bloqué.` : null;
  };
  const save = async () => {
    if (!editable || saving) return { statut: "rien" };
    const aSauver = decades.filter((d) => calc[d.n].saisi && calc[d.n].ok && (!saved[d.n] || saved[d.n].envoyes !== calc[d.n].envoyes || saved[d.n].recus !== calc[d.n].recus)).map((d) => d.n);
    const aSupprimer = decades.filter((d) => !calc[d.n].saisi && saved[d.n]).map((d) => d.n);
    if (!aSauver.length && !aSupprimer.length) return { statut: "rien" };
    setSaving(true);
    try {
      if (aSauver.length) {
        const payload = aSauver.map((n) => ({ bilan_id: bilan.id, decade: n, transferts_ambiant: calc[n].envoyes, retours_ambiant: calc[n].recus, updated_by: userName || null }));
        const r = await supabase.from("bilan_transferts").upsert(payload, { onConflict: "bilan_id,decade" });
        if (r.error) throw r.error;
      }
      if (aSupprimer.length) {
        const r = await supabase.from("bilan_transferts").delete().eq("bilan_id", bilan.id).in("decade", aSupprimer);
        if (r.error) throw r.error;
      }
      setSaved((p) => { const n = { ...p }; aSauver.forEach((k) => { n[k] = { envoyes: calc[k].envoyes, recus: calc[k].recus }; }); aSupprimer.forEach((k) => { delete n[k]; }); return n; });
      // Total net : recalculé par la base ; on relit uniquement les champs du bilan concernés.
      const h = await supabase.from("bilan_matieres").update({ modified_by: userName || null }).eq("id", bilan.id).select("transferts15, transferts_ambiant, modified_at, modified_by").single();
      if (!h.error && h.data && onSaved) onSaved(h.data);
      if (flash) flash("Bilan enregistré.");
      if (appendAudit) appendAudit("Transferts du bilan matières enregistrés", `${siteName} — ${bilan.period_key} : ${aSauver.length} décade(s), ${aSupprimer.length} suppression(s)`);
      return { statut: "ok", texte: `${aSauver.length} décade(s) enregistrée(s)${aSupprimer.length ? `, ${aSupprimer.length} supprimée(s)` : ""}` };
    } catch (e) {
      return { statut: "erreur", texte: erreurLisible(e) };
    } finally { setSaving(false); }
  };
  useImperativeHandle(ref, () => ({ save, verifier, pret: () => !loading, messages: () => messagesTransferts(decades.map((d) => ({ n: d.n, calc: calc[d.n] }))),
    donnees: () => ({ decades: decades.map((d) => ({ n: d.n, debut: d.debut, fin: d.fin, saisi: calc[d.n].saisi && calc[d.n].ok, envoyes: calc[d.n].envoyes, recus: calc[d.n].recus, net: calc[d.n].net })), tot }) }));

  if (loading) return <p className="sap-muted" data-testid="tr-chargement">Chargement des transferts du mois…</p>;
  const fmtInfo = (x) => {
    const a = x.nbEnvoyes || x.nbRecus ? `Envoyés : ${fmtL(x.envoyes)} L (${x.nbEnvoyes}) · Reçus : ${fmtL(x.recus)} L (${x.nbRecus})` : null;
    const b = x.nbChargements ? `Chargements camions : ${fmtL(x.chargements)} L (${x.nbChargements})` : null;
    return [a, b].filter(Boolean).join(" — ") || "—";
  };
  return (
    <div>
      {!editable && <p className="sap-muted" style={{ margin: "0 0 8px" }} data-testid="tr-lecture">Bilan verrouillé : les transferts sont en lecture seule.</p>}
      <div className="sap-gridwrap" style={{ maxHeight: "none" }} data-testid="sap-transferts">
        <table className="sap-table sap-grid">
          <thead><tr><th>Décade</th><th className="sap-num">Envoyés à d'autres sites (L)</th><th className="sap-num">Reçus d'autres sites (L)</th><th className="sap-num">Effet sur le stock (L)</th><th className="sap-mark"></th><th>Pour information : déjà enregistré dans l'application (non compté)</th></tr></thead>
          <tbody>
            {decades.map((d, i) => {
              const c = calc[d.n]; const err = c.saisi && !c.ok;
              return (
                <tr key={d.n} data-decade={d.n} className={modifiee(d.n) ? "sap-row-dirty" : ""}>
                  <td className="sap-d">Décade {d.n} <span className="sap-muted">du {fmtJour(d.debut).slice(0, 5)} au {fmtJour(d.fin).slice(0, 5)}</span></td>
                  {champs.map((champ, col) => (
                    <td key={champ}>
                      <input ref={(el) => { refs.current[`${i}-${col}`] = el; }} className={`sap-cell ${err ? "sap-cell-err" : ""}`} value={vals[d.n]?.[champ] ?? ""} disabled={!editable} inputMode="decimal"
                        aria-label={`${champ === "envoyes" ? "Transferts envoyés" : "Transferts reçus"} décade ${d.n}`} data-testid={`tr-${d.n}-${champ}`} title={err ? c.erreur : undefined}
                        onChange={(e) => setCell(d.n, champ, e.target.value)} onKeyDown={(e) => onKeyDown(e, i, col)} onFocus={(e) => e.target.select()} />
                    </td>
                  ))}
                  <td className="sap-num sap-ro" data-testid={`tr-${d.n}-net`}>{c.saisi && c.ok ? fmtSigned(c.net) : ""}</td>
                  <td className="sap-mark" title={err ? c.erreur : ""} data-testid={`tr-${d.n}-mark`}>{err ? <span className="sap-msg-err">✖</span> : ""}</td>
                  <td className="sap-d sap-muted" style={{ whiteSpace: "normal" }} data-testid={`tr-${d.n}-info`}>{fmtInfo(info[i])}</td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr className="sap-total">
              <td className="sap-d">Total du mois</td>
              <td className="sap-num" data-testid="tr-total-envoyes">{fmtL(tot.envoyes)}</td>
              <td className="sap-num" data-testid="tr-total-recus">{fmtL(tot.recus)}</td>
              <td className="sap-num" data-testid="tr-total-net">{fmtSigned(tot.net)}</td><td></td>
              <td className="sap-d sap-muted" style={{ whiteSpace: "normal" }} data-testid="tr-total-info">{fmtInfo(Object.fromEntries(["envoyes", "nbEnvoyes", "recus", "nbRecus", "chargements", "nbChargements"].map((k) => [k, info.reduce((a, x) => a + x[k], 0)])))}</td>
            </tr>
          </tfoot>
        </table>
      </div>
      <p className="sap-muted" style={{ fontSize: 11.5, margin: "8px 0 0" }}>Saisie à l'ambiant, sans température ni conversion à 15 °C : ces volumes servent uniquement à équilibrer les stocks (envoyés à un autre site = sorties du stock, reçus d'un autre site = entrées ; le site concerné n'est pas saisi). Les ventes restent saisies par jour, avec la température, dans l'onglet « Ventes journalières ». Les chiffres « pour information » ne sont jamais comptés.</p>
    </div>
  );
});

// ---------------------------------------------------------------------
// Onglet « Stocks début / fin » : volume ambiant, température et densité observée, ramenés à 15 °C
// (même fonction que « Correction 15 °C » : densité à 15 °C d'abord, puis coefficient).
//   • Stock début : proposé = stock fin du mois précédent ; modifiable par le Superviseur seul.
//   • Stock fin   : jaugeage du dernier jour, saisi par le Superviseur ou l'opérateur du site.
// ---------------------------------------------------------------------
const CHAMPS_STOCK = ["ambiant", "temp", "densite"];
const tousVides = (f) => CHAMPS_STOCK.every((k) => String(f[k] ?? "").trim() === "");
const enTexte = (o) => ({ ambiant: String(o.ambiant), temp: String(o.temp), densite: String(o.densite) });
const VIDE = { ambiant: "", temp: "", densite: "" };
function evalBloc(f) {
  if (tousVides(f)) return { vide: true };
  const a = lireNombre(f.ambiant), t = lireNombre(f.temp), d = lireNombre(f.densite);
  const premier = (cond) => CHAMPS_STOCK.find((k, i) => cond([a, t, d][i])); // premier champ fautif, pour y amener le curseur
  if ([a, t, d].some((x) => Number.isNaN(x))) return { ok: false, erreur: "Valeur non numérique", champ: premier((x) => Number.isNaN(x)) };
  if (a === null || t === null || d === null) return { ok: false, erreur: "Volume, température et densité sont obligatoires", champ: premier((x) => x === null) };
  const cv = convertir15({ ambiant: a, temp: t, densite: d });
  if (!cv.ok) return { ok: false, erreur: cv.erreur, champ: /Densité/.test(cv.erreur) ? "densite" : /Température/.test(cv.erreur) ? "temp" : "ambiant" };
  return { ok: true, a, t, d, conv: cv, avertissement: t < TEMP_MIN_PLAUSIBLE || t > TEMP_MAX_PLAUSIBLE ? `Température hors plage plausible (${TEMP_MIN_PLAUSIBLE} à ${TEMP_MAX_PLAUSIBLE} °C)` : undefined };
}
const memeBloc = (x, y) => {
  if (!x && !y) return true; if (!x || !y) return false;
  return ["ambiant", "temp", "densite", "v15"].every((k) => (x[k] === null && y[k] === null) || (x[k] !== null && y[k] !== null && Math.abs(x[k] - y[k]) < 1e-6));
};

const StocksPanel = forwardRef(function StocksPanel({ bilan, prec, editable, canManage, userName, siteName, onDirty, onSaved, appendAudit, flash }, ref) {
  const editableDeb = editable && canManage;
  const prop = prec && (prec.statut === "historique" || (prec.stock_fin_ambiant !== null && prec.stock_fin_ambiant !== undefined))
    ? { ambiant: num(prec.stock_fin_ambiant), temp: num(prec.stock_fin_temp), densite: num(prec.stock_fin_densite), v15: num(prec.stock_fin15) } : null;
  const propComplet = !!prop && prop.ambiant !== null && prop.temp !== null && prop.densite !== null;
  const prevV15 = prop ? prop.v15 : null;
  const savedDeb = bilan.stock_debut15 !== null && bilan.stock_debut15 !== undefined ? { ambiant: num(bilan.stock_debut_ambiant), temp: num(bilan.stock_debut_temp), densite: num(bilan.stock_debut_densite), v15: num(bilan.stock_debut15) } : null;
  const savedFin = bilan.stock_fin_ambiant !== null && bilan.stock_fin_ambiant !== undefined ? { ambiant: num(bilan.stock_fin_ambiant), temp: num(bilan.stock_fin_temp), densite: num(bilan.stock_fin_densite), v15: num(bilan.stock_fin15) } : null;

  const [deb, setDeb] = useState(() => (savedDeb && savedDeb.ambiant !== null ? enTexte(savedDeb) : !savedDeb && propComplet ? enTexte(prop) : VIDE));
  const [debSource, setDebSource] = useState(() => (!savedDeb ? "propose" : savedDeb.ambiant !== null ? "saisi" : "report"));
  const [debReport, setDebReport] = useState(() => (savedDeb && savedDeb.ambiant === null ? savedDeb.v15 : null));
  const [fin, setFin] = useState(() => (savedFin ? enTexte(savedFin) : VIDE));
  const [saving, setSaving] = useState(false);
  const refs = useRef({});

  const debEval = evalBloc(deb), finEval = evalBloc(fin);
  // Cible d'écriture : undefined = rien à écrire ; null = effacer ; "erreur" ; ou le bloc à enregistrer.
  let debTarget;
  if (debSource === "saisi") debTarget = debEval.vide ? (savedDeb ? null : undefined) : debEval.ok ? { ambiant: debEval.a, temp: debEval.t, densite: debEval.d, v15: debEval.conv.volume15 } : "erreur";
  else if (debSource === "report") debTarget = { ambiant: null, temp: null, densite: null, v15: debReport };
  const finTarget = finEval.vide ? (savedFin ? null : undefined) : finEval.ok ? { ambiant: finEval.a, temp: finEval.t, densite: finEval.d, v15: finEval.conv.volume15 } : "erreur";
  const dirty = (t, s) => t === "erreur" || (t === null && !!s) || (!!t && t !== "erreur" && !memeBloc(t, s));
  const dirtyDeb = editableDeb && dirty(debTarget, savedDeb), dirtyFin = editable && dirty(finTarget, savedFin);
  const nbDirty = (dirtyDeb ? 1 : 0) + (dirtyFin ? 1 : 0);
  useEffect(() => { if (onDirty) onDirty(nbDirty); }, [nbDirty]); // eslint-disable-line react-hooks/exhaustive-deps

  const ordre = ["deb-ambiant", "deb-temp", "deb-densite", "fin-ambiant", "fin-temp", "fin-densite"];
  const aller = (cle, sens) => { const dispo = ordre.filter((k) => refs.current[k] && !refs.current[k].disabled); const i = dispo.indexOf(cle); const nxt = dispo[i + sens]; if (nxt) refs.current[nxt].focus(); };
  const onKeyDown = (e, cle) => { if (e.key === "Enter") { e.preventDefault(); aller(cle, 1); } };
  const setChamp = (bloc, k, v) => { if (bloc === "deb") { setDeb((p) => ({ ...p, [k]: v })); setDebSource("saisi"); } else setFin((p) => ({ ...p, [k]: v })); };
  const reprendre = () => {
    if (propComplet) { setDeb(enTexte(prop)); setDebSource("saisi"); setDebReport(null); }
    else { setDeb(VIDE); setDebSource("report"); setDebReport(prevV15); }
  };

  const verifier = () => {
    if (dirtyDeb && debTarget === "erreur") return `Stock début : ${debEval.erreur}`;
    if (dirtyFin && finTarget === "erreur") return `Stock fin : ${finEval.erreur}`;
    return null;
  };
  // figerDebut (validation) : un stock début encore « proposé » est enregistré dans le bilan, pour qu'il ne change plus.
  const save = async ({ figerDebut = false } = {}) => {
    const forceDeb = figerDebut && editableDeb && !savedDeb && debSource === "propose" && (propComplet ? debEval.ok : prevV15 !== null);
    const targetDeb = forceDeb ? (propComplet ? { ambiant: debEval.a, temp: debEval.t, densite: debEval.d, v15: debEval.conv.volume15 } : { ambiant: null, temp: null, densite: null, v15: prevV15 }) : debTarget;
    const ecrireDeb = dirtyDeb || forceDeb;
    if (saving || !(ecrireDeb || dirtyFin)) return { statut: "rien" };
    const payload = { modified_by: userName || null };
    if (ecrireDeb) { const t = targetDeb; payload.stock_debut_ambiant = t ? t.ambiant : null; payload.stock_debut_temp = t ? t.temp : null; payload.stock_debut_densite = t ? t.densite : null; payload.stock_debut15 = t ? t.v15 : null; }
    if (dirtyFin) { const t = finTarget; payload.stock_fin_ambiant = t ? t.ambiant : null; payload.stock_fin_temp = t ? t.temp : null; payload.stock_fin_densite = t ? t.densite : null; payload.stock_fin15 = t ? t.v15 : 0; }
    setSaving(true);
    try {
      const r = await supabase.from("bilan_matieres").update(payload).eq("id", bilan.id).select(COLS_STOCKS).single();
      if (r.error) throw r.error;
      if (onSaved) onSaved(r.data);
      if (ecrireDeb && targetDeb === null) { setDeb(propComplet ? enTexte(prop) : VIDE); setDebSource("propose"); setDebReport(null); }
      if (forceDeb) { setDebSource(propComplet ? "saisi" : "report"); if (!propComplet) setDebReport(prevV15); }
      const parts = [ecrireDeb ? "stock début" : null, dirtyFin ? "stock fin" : null].filter(Boolean).join(" et ");
      if (flash) flash("Bilan enregistré.");
      if (appendAudit) appendAudit("Stocks du bilan matières enregistrés", `${siteName} — ${bilan.period_key} : ${parts}`);
      return { statut: "ok", texte: parts };
    } catch (e) {
      return { statut: "erreur", texte: erreurLisible(e) };
    } finally { setSaving(false); }
  };


  const debV15 = !debEval.vide && debEval.ok ? debEval.conv.volume15 : debSource === "report" ? debReport : debSource === "propose" ? prevV15 : null;
  const debRemarque = debSource === "saisi" ? (savedDeb && !dirtyDeb ? "Saisi (enregistré)" : debEval.vide ? "Vide : retour à la proposition du mois précédent" : "Saisi — à enregistrer")
    : debSource === "report" ? `Report du stock fin à 15 °C du mois précédent (sans détail ambiant)${dirtyDeb ? " — à enregistrer" : ""}`
    : prevV15 !== null ? `Proposé : stock fin du mois précédent (${fmtL(prevV15)} L à 15 °C)` : "Non disponible : pas de stock fin au mois précédent — saisie manuelle par le Superviseur";
  const finRemarque = savedFin && !dirtyFin ? "Jaugeage enregistré" : finEval.vide ? "Non saisi" : dirtyFin ? "À enregistrer" : "";
  // Messages de vérification : erreurs de saisie, stocks obligatoires pour valider, températures hors plage.
  const messages = () => {
    const out = [];
    const bloc = (nom, ev, v15, cle, manquant) => {
      if (!ev.vide && ev.ok === false) out.push({ niveau: "erreur", texte: `${nom} : ${ev.erreur}`, cible: { onglet: "stocks", champ: `stk-${cle}-${ev.champ || "ambiant"}` } });
      else if (v15 === null || v15 === undefined) out.push({ niveau: "erreur", texte: manquant, cible: { onglet: "stocks", champ: `stk-${cle}-ambiant` } });
      else if (ev.ok && ev.avertissement) out.push({ niveau: "avert", texte: `${nom} : ${ev.avertissement}`, cible: { onglet: "stocks", champ: `stk-${cle}-temp` } });
    };
    bloc("Stock début", debEval, debV15, "deb", "Stock début de mois non disponible (obligatoire pour valider)");
    bloc("Stock fin", finEval, !finEval.vide && finEval.ok ? finEval.conv.volume15 : null, "fin", "Stock fin de mois non saisi (obligatoire pour valider)");
    return out;
  };
  const donnees = () => {
    const bloc = (ev, v15, remarque) => (ev.ok ? { ambiant: ev.a, temp: ev.t, densite: ev.d, d15: ev.conv.densite15, vcf: ev.conv.vcf, v15, remarque } : { ambiant: null, temp: null, densite: null, d15: null, vcf: null, v15, remarque });
    return { deb: bloc(debEval, debV15, debRemarque), fin: bloc(finEval, !finEval.vide && finEval.ok ? finEval.conv.volume15 : null, finRemarque) };
  };
  useImperativeHandle(ref, () => ({ save, verifier, pret: () => true, messages, donnees }));

  const ligne = (cle, titre, f, ev, enabled, proposeStyle, v15, remarque) => {
    const err = ev.ok === false, warn = ev.ok && ev.avertissement;
    return (
      <tr key={cle} data-testid={`stk-${cle}`}>
        <td className="sap-d" style={{ fontWeight: 600 }}>{titre}</td>
        {CHAMPS_STOCK.map((k) => (
          <td key={k}>
            <input ref={(el) => { refs.current[`${cle}-${k}`] = el; }} className={`sap-cell ${err ? "sap-cell-err" : warn && k === "temp" ? "sap-cell-warn" : ""} ${proposeStyle ? "sap-cell-prop" : ""}`}
              value={f[k]} disabled={!enabled} inputMode="decimal" aria-label={`${titre} — ${k}`} data-testid={`stk-${cle}-${k}`} title={err ? ev.erreur : warn && k === "temp" ? ev.avertissement : undefined}
              onChange={(e) => setChamp(cle, k, e.target.value)} onKeyDown={(e) => onKeyDown(e, `${cle}-${k}`)} onFocus={(e) => e.target.select()} />
          </td>
        ))}
        <td className="sap-num sap-ro" data-testid={`stk-${cle}-d15`}>{ev.ok && ev.conv.densite15 !== null && ev.conv.densite15 !== undefined ? ev.conv.densite15.toLocaleString("fr-FR", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) : ""}</td>
        <td className="sap-num sap-ro" data-testid={`stk-${cle}-vcf`}>{ev.ok && ev.conv.vcf ? ev.conv.vcf.toFixed(4).replace(".", ",") : ""}</td>
        <td className="sap-num sap-ro" data-testid={`stk-${cle}-v15`} style={{ fontWeight: 600 }}>{v15 !== null && v15 !== undefined ? fmtL(v15) : ""}</td>
        <td className="sap-mark" title={err ? ev.erreur : warn ? ev.avertissement : ""} data-testid={`stk-${cle}-mark`}>{err ? <span className="sap-msg-err">✖</span> : warn ? <span className="sap-msg-warn">⚠</span> : ""}</td>
        <td className="sap-d sap-muted" style={{ whiteSpace: "normal", minWidth: 240 }} data-testid={`stk-${cle}-remarque`}>{remarque}</td>
      </tr>
    );
  };
  return (
    <div>
      {!editable && <p className="sap-muted" style={{ margin: "0 0 8px" }} data-testid="stk-lecture">Bilan verrouillé : les stocks sont en lecture seule.</p>}
      <div className="sap-gridwrap" style={{ maxHeight: "none" }}>
        <table className="sap-table sap-grid" data-testid="sap-stocks">
          <thead><tr><th>Poste</th><th className="sap-num">Volume ambiant (L) *</th><th className="sap-num">Température (°C) *</th><th className="sap-num">Densité observée *</th><th className="sap-num">Densité à 15 °C</th><th className="sap-num">Coefficient</th><th className="sap-num">Volume 15 °C (L)</th><th className="sap-mark"></th><th>Remarque</th></tr></thead>
          <tbody>
            {ligne("deb", "Stock début de mois", deb, debEval, editableDeb, debSource === "propose" && propComplet, debV15, debRemarque)}
            {ligne("fin", "Stock fin de mois (jaugeage)", fin, finEval, editable, false, !finEval.vide && finEval.ok ? finEval.conv.volume15 : null, finRemarque)}
          </tbody>
        </table>
      </div>
      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginTop: 10 }}>
        {editableDeb && prevV15 !== null && <button className="sap-btn" onClick={reprendre} data-testid="stk-reprendre">Reprendre le stock fin du mois précédent</button>}
        <span className="sap-muted" style={{ fontSize: 12 }}>Le stock début est proposé automatiquement ; seul le Superviseur peut le modifier. La densité observée est celle mesurée à la température du produit : elle est d'abord ramenée à 15 °C, puis le coefficient est appliqué.</span>
      </div>
    </div>
  );
});

// Contenu réservé à l'impression : invisible à l'écran, seul visible sur la feuille imprimée (voir @media print).
function DocImpression({ doc }) {
  return (
    <div className="sap-printonly" data-testid="sap-print">
      <h2>{doc.title}</h2>
      <p style={{ margin: "0 0 6px" }}>{doc.subtitle} — {doc.period}</p>
      {doc.sections.map((sec) => (
        <div key={sec.heading}>
          <h3>{sec.heading}</h3>
          <table><thead><tr>{sec.columns.map((col, i) => <th key={i}>{col}</th>)}</tr></thead>
            <tbody>{sec.rows.map((r, i) => <tr key={i}>{r.map((x, j) => <td key={j}>{x}</td>)}</tr>)}</tbody>
            {sec.totalsRow && <tfoot><tr>{sec.totalsRow.map((x, j) => <td key={j}>{x}</td>)}</tr></tfoot>}
          </table>
        </div>
      ))}
    </div>
  );
}

// ---------------------------------------------------------------------
// Présentation PowerPoint de PLUSIEURS sites dans un même fichier (Superviseur).
//   On choisit le mois, on coche les sites à inclure ; les chiffres viennent des bilans du nouveau module
//   (même synthèse que l'écran). Les textes (autonomie, natures, justificatifs, conclusion, recommandations)
//   et la génération du .pptx sont ceux du panneau « Présentation » existant (passé en paramètre).
// ---------------------------------------------------------------------
function PresentationSites({ sites, movements, month, year, canManage, userName, PresentationPanel, bilanPresentations, bilanPresStatus, saveBilanPresentation, objectifFreinte, onRetour }) {
  const key = `${year}-${pad2(month)}`;
  const prevKey = month === 1 ? `${year - 1}-12` : `${year}-${pad2(month - 1)}`;
  const [etat, setEtat] = useState({ loading: true, erreur: null, lignes: [] });
  const [coche, setCoche] = useState({});
  useEffect(() => {
    let off = false;
    (async () => {
      try {
        const [cur, prev] = await Promise.all([
          supabase.from("bilan_matieres").select(COLS_PRES).eq("period_type", "mensuel").eq("period_key", key),
          supabase.from("bilan_matieres").select(COLS_PREC + ", site_id").eq("period_type", "mensuel").eq("period_key", prevKey),
        ]);
        if (cur.error) throw cur.error;
        if (off) return;
        const lignes = sites.map((s) => ({ site: s, ...lignePresentation({ site: s, bilan: (cur.data || []).find((b) => b.site_id === s.id) || null, prec: prev.error ? null : (prev.data || []).find((b) => b.site_id === s.id) || null, movements }) }));
        setEtat({ loading: false, erreur: null, lignes });
        setCoche(Object.fromEntries(lignes.filter((l) => l.etat === "ok").map((l) => [l.site.id, true])));
      } catch (e) { if (!off) setEtat({ loading: false, erreur: erreurLisible(e), lignes: [] }); }
    })();
    return () => { off = true; };
  }, [key]); // eslint-disable-line react-hooks/exhaustive-deps
  const periodeLabel = `${MOIS[month - 1]} ${year}`;
  const ok = etat.lignes.filter((l) => l.etat === "ok");
  const choisies = ok.filter((l) => coche[l.site.id]);
  const rows = choisies.map((l) => l.ligne);
  const saved = (bilanPresentations || []).find((p) => p.periodType === "mensuel" && p.periodKey === key);
  const ready = bilanPresStatus === "ready" || bilanPresStatus === "unavailable" || bilanPresStatus === "error";
  const brouillons = choisies.filter((l) => l.statut === "brouillon").length;
  return (
    <div className="sap-root" data-testid="sap-presentation">
      <style>{STYLE}</style>
      <div className="sap-shell"><div className="sap-shell-id"><span className="sap-kicker">BILMAT · Plusieurs sites dans un même fichier</span><h2 className="sap-title">Présentation PowerPoint — {periodeLabel}</h2></div></div>
      <div className="sap-toolbar"><button className="sap-btn sap-btn-ghost" onClick={onRetour}><Ico n="back" />Retour</button></div>
      <div style={{ padding: "22px 28px" }}>
        {etat.loading && <p className="sap-muted" data-testid="pres-chargement">Chargement des bilans de {periodeLabel}…</p>}
        {etat.erreur && <p className="sap-msg sap-msg-err" data-testid="pres-erreur">✖ {etat.erreur}</p>}
        {!etat.loading && !etat.erreur && (
          <>
            <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 6, flexWrap: "wrap" }}>
              <b>Sites à inclure ({choisies.length} sur {ok.length})</b>
              <button className="sap-btn" data-testid="pres-all" onClick={() => setCoche(Object.fromEntries(ok.map((l) => [l.site.id, true])))}>Tout sélectionner</button>
              <button className="sap-btn" data-testid="pres-none" onClick={() => setCoche({})}>Aucun</button>
            </div>
            <div className="sap-gridwrap" style={{ maxHeight: "none" }}>
              <table className="sap-table" data-testid="pres-table">
                <thead><tr><th></th><th>Site</th><th>Statut du bilan</th><th className="sap-num">Stock théorique (L)</th><th className="sap-num">Stock physique (L)</th><th className="sap-num">Écart (L)</th><th className="sap-num">Taux (‰)</th><th>Remarque</th></tr></thead>
                <tbody>
                  {etat.lignes.map((l) => {
                    const l1 = l.ligne; const stt = l.statut ? (STATUTS[l.statut] || STATUTS.historique) : null;
                    const taux = l1 && l1.ventes15 ? (l1.ecart / l1.ventes15) * 1000 : null;
                    return (
                      <tr key={l.site.id} data-testid={`pres-ligne-${l.site.id}`}>
                        <td><input type="checkbox" disabled={l.etat !== "ok"} checked={!!coche[l.site.id]} onChange={(e) => setCoche((p) => ({ ...p, [l.site.id]: e.target.checked }))} aria-label={`Inclure ${l.site.name}`} data-testid={`pres-check-${l.site.id}`} /></td>
                        <td className="sap-d" style={{ fontWeight: 600 }}>{l.site.name}</td>
                        <td className="sap-d">{stt ? <span className={`sap-badge ${stt.cls}`}>{stt.label}</span> : "—"}</td>
                        <td className="sap-num">{l1 ? fmtL(l1.stockTheorique) : "—"}</td>
                        <td className="sap-num">{l1 ? fmtL(l1.stockFin15) : "—"}</td>
                        <td className={`sap-num ${l1 ? clsSigne(l1.ecart) : ""}`} data-testid={`pres-ecart-${l.site.id}`}>{l1 ? fmtSigned(l1.ecart) : "—"}</td>
                        <td className="sap-num">{taux === null ? "—" : `${taux > 0 ? "+" : ""}${taux.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`}</td>
                        <td className="sap-d sap-muted">{l.etat === "absent" ? "Aucun bilan pour ce mois" : l.etat === "incomplet" ? "Stock début ou stock fin manquant : écart non calculable" : l.statut === "brouillon" ? "Bilan non validé" : ""}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            {brouillons > 0 && <p className="sap-msg sap-msg-warn" style={{ margin: "8px 0 0" }} data-testid="pres-brouillons">⚠ {brouillons} site(s) sélectionné(s) ont un bilan non validé : les chiffres peuvent encore changer.</p>}
            <div style={{ marginTop: 12 }} data-testid="pres-panel">
              {!PresentationPanel ? <p className="sap-muted">Génération indisponible.</p>
                : !canManage ? <p className="sap-muted">La présentation est réservée au Superviseur.</p>
                : rows.length === 0 ? <p className="sap-muted" data-testid="pres-aucun">Coche au moins un site pour préparer la présentation.</p>
                : !ready ? <p className="sap-muted">Chargement des textes de la présentation…</p>
                : <PresentationPanel key={`${key}|${rows.map((r) => r.site.id).join(",")}|${saved ? saved.updatedAt || "x" : "-"}`} rows={rows} periodType="mensuel" periodKey={key} saved={saved} status={bilanPresStatus} onSave={saveBilanPresentation} defaultPreparePar={userName ? `${userName} — Superviseur Sites Externalisés` : ""} seuil={objectifFreinte} />}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

export default function BilansMatieresSap({ sites = [], movements = [], canManage = false, canEdit = false, assignedSiteIds = [], userName = "", appendAudit, flash, exportToExcel, exportToPdf, PresentationPanel, bilanPresentations = [], bilanPresStatus = "idle", saveBilanPresentation, objectifFreinte = 3 }) {
  const selectableSites = useMemo(() => sites.filter((s) => !s.isMobile && (!assignedSiteIds?.length || assignedSiteIds.includes(s.id))), [sites, assignedSiteIds]);
  const d0 = new Date(); const dPrev = new Date(d0.getFullYear(), d0.getMonth() - 1, 1); // par défaut : le mois qui vient de se terminer
  const [sel, setSel] = useState({ siteId: "", month: dPrev.getMonth() + 1, year: dPrev.getFullYear() });
  const [screen, setScreen] = useState("select");
  const [tab, setTab] = useState("synthese");
  const [loading, setLoading] = useState(false);
  const [msg, setMsg] = useState(null); // { type: "ok" | "warn" | "err", text }
  const [ventesOuvert, setVentesOuvert] = useState(false); // la grille n'est chargée qu'à la première ouverture de l'onglet
  const [nbModifVentes, setNbModifVentes] = useState(0);
  const [nbModifStocks, setNbModifStocks] = useState(0);
  const [stocksOuvert, setStocksOuvert] = useState(false);
  const [transfertsOuvert, setTransfertsOuvert] = useState(false);
  const [nbModifTransferts, setNbModifTransferts] = useState(0);
  const [nbTransferts, setNbTransferts] = useState(0);
  const transfertsRef = useRef(null);
  const stocksRef = useRef(null);
  const [joursSaisis, setJoursSaisis] = useState(0);
  const [joursDb, setJoursDb] = useState(0);                 // affichage : jours de ventes déjà enregistrés (barre de progression dès l'ouverture)
  const [ventesCharge, setVentesCharge] = useState(false);   // affichage : la grille est chargée (son compte est alors le plus à jour)
  const [rouvertSession, setRouvertSession] = useState(false); // affichage : bilan rouvert pendant cette session (badge bleu « Rouvert »)
  const gridRef = useRef(null);
  const [paramVal, setParamVal] = useState("");
  const [ventesParLigne, setVentesParLigne] = useState(null); // détail des ventes par ligne (sites à plusieurs lignes) : { labels, ambiant, volume15 }
  const [lignesTxt, setLignesTxt] = useState("");              // réglage Superviseur : lignes de vente du site choisi (séparées par des virgules)
  const [verif, setVerif] = useState(null); // résultat de « Vérifier » : { messages }
  const [obs, setObs] = useState("");           // observation du mois (reprise dans la présentation PowerPoint)
  const [docImpression, setDocImpression] = useState(null);
  const [occupe, setOccupe] = useState(false);
  const [bilan, setBilan] = useState(null);
  const [prec, setPrec] = useState(null);
  const siteId = sel.siteId || selectableSites[0]?.id || "";
  const site = sites.find((s) => s.id === (bilan?.site_id || siteId));
  const periodKey = bilan ? bilan.period_key : `${sel.year}-${pad2(sel.month)}`;
  const [pY, pM] = periodKey.split("-").map(Number);
  const periodLabel = `${MOIS[(pM || 1) - 1]} ${pY}`;

  const rec = useMemo(() => (bilan ? receptionsDuMois(movements, bilan.site_id, bilan.period_key) : null), [movements, bilan]);

  const obsDirty = !!bilan && bilan.statut === "brouillon" && canEdit && obs !== (bilan.commentaire || "");
  const nbModifiees = nbModifVentes + nbModifStocks + nbModifTransferts + (obsDirty ? 1 : 0);
  const executer = async () => {
    if (!siteId || loading) return;
    setLoading(true); setMsg(null);
    const key = `${sel.year}-${pad2(sel.month)}`;
    const prevKey = sel.month === 1 ? `${sel.year - 1}-12` : `${sel.year}-${pad2(sel.month - 1)}`;
    const siteName = sites.find((s) => s.id === siteId)?.name || siteId;
    try {
      const [cur, prev, param, cfg] = await Promise.all([
        supabase.from("bilan_matieres").select(COLS).eq("site_id", siteId).eq("period_type", "mensuel").eq("period_key", key).maybeSingle(),
        supabase.from("bilan_matieres").select(COLS_PREC).eq("site_id", siteId).eq("period_type", "mensuel").eq("period_key", prevKey).maybeSingle(),
        supabase.from("bilan_parametres").select("valeur").eq("cle", "densite_vente").maybeSingle(),
        supabase.from("bilan_ventes_lignes").select("lignes").eq("site_id", siteId).maybeSingle(),
      ]);
      if (cur.error) throw cur.error;
      let row = cur.data;
      let cree = false;
      if (!row) {
        if (!canEdit) { setMsg({ type: "warn", text: `Aucun bilan pour ${siteName} — ${MOIS[sel.month - 1]} ${sel.year}. Ton compte ne peut pas en créer.` }); return; }
        const densite = Number(param.data?.valeur) || DENSITE_VENTE_DEFAUT;
        const ins = await supabase.from("bilan_matieres").insert({
          site_id: siteId, period_type: "mensuel", period_key: key, statut: "brouillon", densite_vente: densite,
          reception15: 0, ventes15: 0, transferts15: 0, stock_fin15: 0, created_by: userName || null, modified_by: userName || null,
        }).select(COLS).single();
        if (ins.error && ins.error.code === "23505") {
          // créé entre-temps par quelqu'un d'autre : on l'ouvre
          const again = await supabase.from("bilan_matieres").select(COLS).eq("site_id", siteId).eq("period_type", "mensuel").eq("period_key", key).maybeSingle();
          if (again.error || !again.data) throw ins.error;
          row = again.data;
        } else if (ins.error) throw ins.error; else { row = ins.data; cree = true; }
      }
      // Ventes déjà saisies : nombre de jours (barre de progression) et, pour les sites à plusieurs lignes, détail par ligne pour la
      // Synthèse (facultatif : une erreur ici ne bloque pas l'ouverture)
      let parLigne = null, nbJoursDb = 0;
      try {
        const lignesCfg = !cfg.error && cfg.data && Array.isArray(cfg.data.lignes) ? cfg.data.lignes.filter(Boolean) : [];
        let rv = await supabase.from("bilan_ventes").select("jour, categorie, volume_ambiant, volume15").eq("bilan_id", row.id);
        if (rv.error && (rv.error.code === "42703" || /categorie/.test(String(rv.error.message || "")))) rv = await supabase.from("bilan_ventes").select("jour, volume_ambiant, volume15").eq("bilan_id", row.id);
        if (!rv.error) {
          nbJoursDb = new Set((rv.data || []).map((l) => l.jour)).size;
          if (lignesCfg.length >= 2) {
            const amb = lignesCfg.map(() => 0), v15 = lignesCfg.map(() => 0);
            (rv.data || []).forEach((l) => { const k = lignesCfg.indexOf(l.categorie || ""); if (k >= 0) { amb[k] += Number(l.volume_ambiant) || 0; v15[k] += Number(l.volume15) || 0; } });
            parLigne = { labels: lignesCfg, ambiant: amb, volume15: v15 };
          }
        }
      } catch (e) { parLigne = null; }
      setVentesParLigne(parLigne); setJoursDb(nbJoursDb); setVentesCharge(false); setRouvertSession(false);
      setBilan(row); setObs(row.commentaire || ""); setPrec(prev.error ? null : prev.data || null); setTab("synthese"); setScreen("bilan"); setVentesOuvert(false); setNbModifVentes(0); setStocksOuvert(false); setNbModifStocks(0); setTransfertsOuvert(false); setNbModifTransferts(0); setNbTransferts(0); setJoursSaisis(0);
      if (cree) {
        if (appendAudit) appendAudit("Bilan matières créé", `${siteName} — ${MOIS[sel.month - 1]} ${sel.year}`);
        if (flash) flash("Bilan créé en brouillon.");
        setMsg({ type: "ok", text: "Bilan créé en brouillon." });
      } else if (prev.error) {
        setMsg({ type: "warn", text: "Bilan du mois précédent illisible : le stock début ne peut pas être proposé." });
      }
    } catch (e) {
      setMsg({ type: "err", text: erreurLisible(e) });
    } finally { setLoading(false); }
  };
  const retour = () => {
    if (nbModifiees > 0 && typeof window !== "undefined" && !window.confirm("Des modifications ne sont pas enregistrées. Quitter quand même ?")) return;
    setScreen("select"); setBilan(null); setPrec(null); setMsg(null); setVerif(null); setVentesOuvert(false); setNbModifVentes(0); setStocksOuvert(false); setNbModifStocks(0); setTransfertsOuvert(false); setNbModifTransferts(0);
  };
  // Paramètre « densité de vente » (Superviseur) : chargé à l'ouverture de l'écran de sélection.
  useEffect(() => {
    if (!canManage || screen !== "select" || paramVal !== "") return; // chargé une seule fois par session
    let off = false;
    supabase.from("bilan_parametres").select("valeur").eq("cle", "densite_vente").maybeSingle().then((r) => { if (!off && !r.error) setParamVal(r.data ? String(r.data.valeur) : String(DENSITE_VENTE_DEFAUT)); });
    return () => { off = true; };
  }, [canManage, screen]);
  const aujourdhuiStr = () => { const d = new Date(); return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; };
  // Les onglets ne se chargent qu'à leur première ouverture : Vérifier et Valider les ouvrent tous (quelques petites requêtes).
  const ouvrirTout = async () => {
    setVentesOuvert(true); setStocksOuvert(true); setTransfertsOuvert(true);
    for (let i = 0; i < 120; i++) {
      if ([gridRef.current, transfertsRef.current, stocksRef.current].every((r) => r && r.pret())) return true;
      await new Promise((r) => setTimeout(r, 50));
    }
    return false;
  };
  // Tous les messages de vérification : ventes, transferts, stocks, puis contrôles généraux (erreurs d'abord).
  const collecter = () => {
    const out = [];
    if (gridRef.current) out.push(...gridRef.current.messages(aujourdhuiStr()));
    if (transfertsRef.current) out.push(...transfertsRef.current.messages());
    if (stocksRef.current) out.push(...stocksRef.current.messages());
    out.push(...messagesGeneraux({ periodKey: bilan.period_key, aujourdhui: aujourdhuiStr(), nbReceptionsSansMesure: rec ? rec.nbSansMesure : 0 }));
    return trierMessages(out);
  };
  // Clic sur un message : on ouvre l'onglet concerné et on place le curseur sur la cellule.
  const allerA = (cible) => {
    if (!cible) return;
    setTab(cible.onglet);
    if (cible.champ) setTimeout(() => { const el = document.querySelector(`[data-testid="${cible.champ}"]`); if (el) { el.focus(); if (el.scrollIntoView) el.scrollIntoView({ block: "nearest" }); } }, 80);
  };
  // Enregistre ventes, transferts et stocks d'un coup : contrôle des trois, puis écriture des seules modifications.
  const sauver = async ({ figerDebut = false } = {}) => {
    const comps = [["ventes", gridRef.current], ["transferts", transfertsRef.current], ["stocks", stocksRef.current]].filter(([, r]) => r);
    for (const [, r] of comps) { const e = r.verifier(); if (e) return { statut: "erreur", texte: e, bloque: true }; }
    const parts = [];
    for (const [nom, r] of comps) {
      const res = await r.save({ figerDebut });
      if (res.statut === "erreur") return { statut: "erreur", texte: res.texte };
      if (res.statut === "ok") parts.push([nom, res.texte]);
    }
    if (obsDirty) {
      const u = await supabase.from("bilan_matieres").update({ commentaire: obs.trim() || null, modified_by: userName || null }).eq("id", bilan.id).select("commentaire, modified_at, modified_by").single();
      if (u.error) return { statut: "erreur", texte: erreurLisible(u.error) };
      setBilan((b) => ({ ...b, ...u.data }));
      parts.push(["observation", "texte enregistré"]);
    }
    return { statut: parts.length ? "ok" : "rien", parts };
  };
  const enregistrer = async () => {
    setMsg(null);
    const r = await sauver();
    if (r.statut === "erreur") { setMsg({ type: "err", text: r.texte, cible: r.bloque ? (collecter().find((m) => m.niveau === "erreur") || {}).cible : undefined }); return; }
    if (r.statut === "rien") setMsg({ type: "warn", text: "Aucune modification à enregistrer." });
    else setMsg({ type: "ok", text: r.parts.length === 1 ? `Bilan enregistré (${r.parts[0][1]}).` : `Bilan enregistré (${r.parts.map(([n, t]) => `${n} : ${t}`).join(" ; ")}).` });
  };
  // Excel / PDF / impression : même contenu que l'écran (onglets Synthèse + détail). Les données enregistrées sont exportées :
  // s'il reste des modifications non enregistrées, on propose de les enregistrer d'abord.
  const exporter = async (format) => {
    if (occupe || !bilan) return;
    if (nbModifiees > 0 && editable) {
      if (!window.confirm("Des modifications ne sont pas enregistrées. Enregistrer maintenant avant d'exporter ?")) return;
      const r = await sauver();
      if (r.statut === "erreur") { setMsg({ type: "err", text: r.texte }); return; }
    }
    setOccupe(true); setMsg(null);
    try {
      let det = {};
      if (!hist) {
        if (!(await ouvrirTout())) { setMsg({ type: "err", text: "Chargement des onglets impossible : réessaie." }); return; }
        det = { ventes: gridRef.current.donnees(), transferts: transfertsRef.current.donnees(), stocks: stocksRef.current.donnees() };
      }
      const doc = construireDocumentBilan({
        siteName: site?.name || bilan.site_id, periodKey: bilan.period_key, periodLabel, statutLabel: st.label.replace(" 🔒", ""),
        entete: { creePar: bilan.created_by, creeLe: bilan.created_at, modifiePar: bilan.modified_by || bilan.created_by, modifieLe: bilan.modified_at || bilan.created_at, validePar: bilan.statut === "valide" ? bilan.valide_par : null, valideLe: bilan.valide_le },
        synthese: syn, ventesParLigne: !hist && ventesParLigne && ventesParLigne.labels.length > 1 ? ventesParLigne : null, observation: bilan.commentaire || "", ventes: det.ventes || null, transferts: det.transferts || null, stocks: det.stocks || null,
        receptions: { lignes: rec.lignes.map((l) => ({ date: l.date, ref: l.ref, quantity: l.quantity, temperatureC: l.temperatureC, densiteObservee: l.densiteObservee, volume15: l.volume15, converti: l.converti })), total15: rec.total15, totalAmbiant: rec.lignes.reduce((a, l) => a + (l.quantity || 0), 0) },
      });
      if (format === "excel") { exportToExcel(`${doc.filename}.xlsx`, doc.sheets); setMsg({ type: "ok", text: `Export Excel généré (${doc.filename}.xlsx).` }); }
      else if (format === "pdf") { await exportToPdf({ filename: `${doc.filename}.pdf`, title: doc.title, subtitle: doc.subtitle, period: doc.period, sections: doc.sections }); setMsg({ type: "ok", text: `Export PDF généré (${doc.filename}.pdf).` }); }
      else { setDocImpression(doc); setTimeout(() => window.print(), 150); setMsg({ type: "ok", text: "Impression lancée." }); }
      if (appendAudit) appendAudit("Bilan matières exporté", `${site?.name || bilan.site_id} — ${bilan.period_key} : ${{ excel: "Excel", pdf: "PDF", print: "impression" }[format]}`);
    } catch (e) {
      setMsg({ type: "err", text: erreurLisible(e) });
    } finally { setOccupe(false); }
  };
  const verifierBilan = async () => {
    if (occupe || !bilan || bilan.statut === "historique") return;
    setOccupe(true); setMsg(null);
    try {
      if (!(await ouvrirTout())) { setMsg({ type: "err", text: "Chargement des onglets impossible : réessaie." }); return; }
      const msgs = collecter(); const n = compterMessages(msgs);
      setVerif({ messages: msgs });
      const premiere = msgs[0] ? msgs[0].cible : undefined;
      setMsg(n.erreurs ? { type: "err", text: `Vérification : ${n.erreurs} erreur(s), ${n.avertissements} avertissement(s).`, cible: premiere }
        : n.avertissements ? { type: "warn", text: `Vérification : aucune erreur, ${n.avertissements} avertissement(s).`, cible: premiere }
        : { type: "ok", text: "Vérification terminée : aucune anomalie." });
    } finally { setOccupe(false); }
  };
  // Valider (Superviseur) : vérification complète, enregistrement, puis verrouillage. Les avertissements demandent confirmation.
  const valider = async () => {
    if (occupe || !bilan || bilan.statut !== "brouillon" || !canManage) return;
    setOccupe(true); setMsg(null);
    try {
      if (!(await ouvrirTout())) { setMsg({ type: "err", text: "Chargement des onglets impossible : réessaie." }); return; }
      const msgs = collecter(); const n = compterMessages(msgs);
      setVerif({ messages: msgs });
      if (n.erreurs) { setMsg({ type: "err", text: `Validation impossible : ${n.erreurs} erreur(s) à corriger.`, cible: msgs[0].cible }); return; }
      if (n.avertissements && !window.confirm(`${n.avertissements} avertissement(s) restent ouverts. Valider quand même et verrouiller le bilan ?`)) { setMsg({ type: "warn", text: "Validation annulée.", cible: msgs[0].cible }); return; }
      const r = await sauver({ figerDebut: true });
      if (r.statut === "erreur") { setMsg({ type: "err", text: r.texte }); return; }
      const u = await supabase.from("bilan_matieres").update({ statut: "valide", valide_par: userName || null, modified_by: userName || null }).eq("id", bilan.id).select("statut, valide_par, valide_le, modified_at, modified_by").single();
      if (u.error) throw u.error;
      setBilan((b) => ({ ...b, ...u.data })); setVerif(null); setRouvertSession(false);
      if (appendAudit) appendAudit("Bilan matières validé", `${site?.name || bilan.site_id} — ${bilan.period_key}`);
      if (flash) flash("Bilan validé et verrouillé.");
      setMsg({ type: "ok", text: "Bilan validé et verrouillé." });
    } catch (e) {
      setMsg({ type: "err", text: erreurLisible(e) });
    } finally { setOccupe(false); }
  };
  // Rouvrir (Superviseur) : le bilan validé repasse en Brouillon (la base efface « validé par / le »).
  const rouvrir = async () => {
    if (occupe || !bilan || bilan.statut !== "valide" || !canManage) return;
    if (!window.confirm("Rouvrir ce bilan ? Il repassera en Brouillon et redeviendra modifiable.")) return;
    setOccupe(true); setMsg(null);
    try {
      const u = await supabase.from("bilan_matieres").update({ statut: "brouillon", modified_by: userName || null }).eq("id", bilan.id).select("statut, valide_par, valide_le, modified_at, modified_by").single();
      if (u.error) throw u.error;
      setBilan((b) => ({ ...b, ...u.data })); setRouvertSession(true);
      if (appendAudit) appendAudit("Bilan matières rouvert", `${site?.name || bilan.site_id} — ${bilan.period_key}`);
      if (flash) flash("Bilan rouvert.");
      setMsg({ type: "ok", text: "Bilan rouvert : il est de nouveau en Brouillon." });
    } catch (e) {
      setMsg({ type: "err", text: erreurLisible(e) });
    } finally { setOccupe(false); }
  };
  // Lignes de vente du site choisi (Superviseur) : chargées quand on change de site.
  useEffect(() => {
    if (!canManage || !siteId) return undefined;
    let off = false;
    (async () => {
      const r = await supabase.from("bilan_ventes_lignes").select("lignes").eq("site_id", siteId).maybeSingle();
      if (!off) setLignesTxt(!r.error && r.data && Array.isArray(r.data.lignes) ? r.data.lignes.join(", ") : "");
    })();
    return () => { off = true; };
  }, [siteId, canManage]); // eslint-disable-line react-hooks/exhaustive-deps
  const enregistrerLignes = async () => {
    const l = String(lignesTxt).split(",").map((x) => x.trim()).filter(Boolean);
    const nomSite = sites.find((s) => s.id === siteId)?.name || siteId;
    if (l.length === 1) { setMsg({ type: "err", text: "Indique au moins 2 lignes (séparées par une virgule), ou laisse vide pour une seule ligne « Ventes »." }); return; }
    if (l.length > 6) { setMsg({ type: "err", text: "6 lignes de vente au maximum." }); return; }
    if (new Set(l.map((x) => x.toLowerCase())).size !== l.length) { setMsg({ type: "err", text: "Deux lignes portent le même nom." }); return; }
    const r = l.length
      ? await supabase.from("bilan_ventes_lignes").upsert({ site_id: siteId, lignes: l, updated_by: userName || null }, { onConflict: "site_id" })
      : await supabase.from("bilan_ventes_lignes").delete().eq("site_id", siteId);
    if (r.error) { setMsg({ type: "err", text: erreurLisible(r.error) }); return; }
    setLignesTxt(l.join(", "));
    setMsg({ type: "ok", text: l.length ? `Lignes de vente de « ${nomSite} » : ${l.join(" / ")}. Elles s'appliquent à la grille des ventes (bilans en cours et à venir).` : `« ${nomSite} » : une seule ligne « Ventes ».` });
    if (appendAudit) appendAudit("Paramètre bilan matières modifié", `Lignes de vente de ${nomSite} : ${l.length ? l.join(" / ") : "une seule ligne"}`);
  };
  const enregistrerParam = async () => {
    const v = Number(String(paramVal).replace(",", "."));
    if (!Number.isFinite(v) || v < 653 || v > 1075) { setMsg({ type: "err", text: "Densité de vente invalide (entre 653 et 1075 kg/m³)." }); return; }
    const r = await supabase.from("bilan_parametres").upsert({ cle: "densite_vente", valeur: v, updated_by: userName || null }, { onConflict: "cle" });
    if (r.error) { setMsg({ type: "err", text: erreurLisible(r.error) }); return; }
    setMsg({ type: "ok", text: `Densité de vente enregistrée : ${v} kg/m³. Elle s'applique aux bilans créés à partir de maintenant.` });
    if (appendAudit) appendAudit("Paramètre bilan matières modifié", `Densité de vente : ${v} kg/m³`);
  };

  // ---------------- Calculs de la synthèse (tout à 15 °C) ----------------
  const hist = bilan?.statut === "historique";
  const syn = bilan ? calculerSynthese({ bilan, prec, rec }) : null; // même définition que les exports et la présentation PowerPoint
  const lignesSynthese = syn ? 8 : 0;
  const nbLignes = tab === "synthese" ? lignesSynthese : tab === "receptions" ? rec?.lignes.length || 0 : tab === "ventes" ? joursSaisis : tab === "stocks" ? 2 : tab === "transferts" ? nbTransferts : 0;
  const editable = bilan?.statut === "brouillon" && canEdit;
  const st = bilan ? STATUTS[bilan.statut] || STATUTS.historique : null;
  // Badge « Rouvert » (bleu) : affichage seulement, pendant la session où le bilan a été rouvert (la base le remet en Brouillon).
  const stAff = bilan && rouvertSession && bilan.statut === "brouillon" ? { label: "Rouvert", cls: "sap-st-reopen" } : st;
  const bientot = "Disponible aux étapes suivantes";

  // ---------------- Écran de sélection ----------------
  if (screen === "presentation") {
    return <PresentationSites sites={selectableSites} movements={movements} month={sel.month} year={sel.year} canManage={canManage} userName={userName} PresentationPanel={PresentationPanel}
      bilanPresentations={bilanPresentations} bilanPresStatus={bilanPresStatus} saveBilanPresentation={saveBilanPresentation} objectifFreinte={objectifFreinte} onRetour={() => setScreen("select")} />;
  }
  if (screen === "select") {
    return (
      <div className="sap-root" data-testid="sap-select" onKeyDown={(e) => { if (e.key === "F8") { e.preventDefault(); executer(); } }}>
        <style>{STYLE}</style>
        <div className="sap-shell"><div className="sap-shell-id"><span className="sap-kicker">Transaction BILMAT</span><h2 className="sap-title">Bilan matières<small>Sélection</small></h2></div></div>
        <div className="sap-toolbar">
          <button className="sap-btn sap-btn-main" onClick={executer} disabled={loading || !siteId} title="Exécuter (F8)"><Ico n="play" />{loading ? "Chargement…" : "Exécuter (F8)"}</button>
          {canManage && PresentationPanel && <button className="sap-btn" onClick={() => setScreen("presentation")} data-testid="sap-btn-presentation" title="Présentation PowerPoint de plusieurs sites, pour le mois choisi"><Ico n="slides" />Présentation PowerPoint (plusieurs sites)…</button>}
        </div>
        <div className="sap-form">
          <div className="sap-row"><label className="sap-label" htmlFor="sap-site">Site *</label>
            <select id="sap-site" className="sap-input" value={siteId} onChange={(e) => setSel({ ...sel, siteId: e.target.value })}>
              {selectableSites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select></div>
          <div className="sap-row"><label className="sap-label" htmlFor="sap-mois">Mois *</label>
            <select id="sap-mois" className="sap-input" value={sel.month} onChange={(e) => setSel({ ...sel, month: Number(e.target.value) })}>
              {MOIS.map((m, i) => <option key={m} value={i + 1}>{pad2(i + 1)} — {m}</option>)}
            </select></div>
          <div className="sap-row"><label className="sap-label" htmlFor="sap-annee">Année *</label>
            <select id="sap-annee" className="sap-input" value={sel.year} onChange={(e) => setSel({ ...sel, year: Number(e.target.value) })}>
              {[d0.getFullYear() - 2, d0.getFullYear() - 1, d0.getFullYear(), d0.getFullYear() + 1].map((y) => <option key={y} value={y}>{y}</option>)}
            </select></div>
          <p className="sap-muted" style={{ margin: "14px 0 0", fontSize: 12 }}>
            Si le bilan existe déjà, il s'ouvre ; sinon il est créé en statut Brouillon{canEdit ? "" : " (création réservée aux comptes autorisés)"}.
          </p>
        </div>
        {canManage && (
          <div className="sap-form" data-testid="sap-param">
            <div className="sap-kicker" style={{ color: "var(--muted)", marginBottom: 14 }}>Paramètres du module (Superviseur)</div>
            <div className="sap-row"><label className="sap-label" htmlFor="sap-densite">Densité de vente</label>
              <div style={{ display: "flex", gap: 6 }}>
                <input id="sap-densite" className="sap-input" inputMode="decimal" value={paramVal} onChange={(e) => setParamVal(e.target.value)} />
                <button className="sap-btn" onClick={enregistrerParam}>Modifier</button>
              </div></div>
            <p className="sap-muted" style={{ margin: 0, fontSize: 12 }}>kg/m³ à l'ambiance, identique pour toutes les ventes. Paramètre réservé au Superviseur ; la valeur est figée dans chaque bilan à sa création.</p>
            <div className="sap-row" style={{ marginTop: 8 }} data-testid="sap-lignes"><label className="sap-label" htmlFor="sap-lignes-input">Lignes de vente du site</label>
              <div style={{ display: "flex", gap: 6 }}>
                <input id="sap-lignes-input" className="sap-input" value={lignesTxt} placeholder="ex. Comilog, GSEZ" disabled={!siteId} onChange={(e) => setLignesTxt(e.target.value)} />
                <button className="sap-btn" onClick={enregistrerLignes} disabled={!siteId} data-testid="sap-lignes-btn">Enregistrer</button>
              </div></div>
            <p className="sap-muted" style={{ margin: 0, fontSize: 12 }}>Pour le site choisi ci-dessus : sépare les lignes par une virgule (2 à 6), ou laisse vide pour une seule ligne « Ventes ». La grille des ventes aura une colonne par ligne et une seule température par jour.</p>
          </div>
        )}
        <div className="sap-status">
          <span className={msg ? `sap-msg sap-msg-${msg.type}` : "sap-muted"} role="status" data-testid="sap-msg" key={msg ? `${msg.type}-${msg.text}` : "pret"}>
            {msg ? <><span className="sap-msg-ic" aria-hidden="true">{{ ok: "✔", warn: "⚠", err: "✖" }[msg.type]}</span> {msg.text}</> : "Prêt."}
          </span>
          <span>{selectableSites.length} site(s) disponible(s)</span>
        </div>
      </div>
    );
  }

  // ---------------- Écran du bilan ----------------
  const ligne = (key, label, val, remarque, opts = {}) => (
    <tr key={key} className={`${opts.total ? "sap-total" : ""} ${opts.final ? "sap-final" : ""}`.trim()}>
      <td style={{ fontWeight: opts.bold ? 600 : 400 }}>{label}</td>
      <td className={`sap-num ${opts.cls || ""}`} data-testid={`syn-${key}`}>{opts.pct ? (Number.isFinite(val) ? `${val > 0 ? "+" : ""}${val.toFixed(2).replace(".", ",")} %` : "—") : opts.signed ? fmtSigned(val) : fmtL(val)}</td>
      <td className="sap-muted">{remarque}</td>
    </tr>
  );
  const c = syn.calc;
  const cnt = verif ? compterMessages(verif.messages) : null;
  return (
    <div className="sap-root" data-testid="sap-bilan">
      <style>{STYLE}</style>
      <div className="sap-shell" data-testid="sap-shell">
        <div className="sap-shell-id"><span className="sap-kicker">BILMAT · Bilan matières</span><h2 className="sap-title" data-testid="sap-titre">{site?.name || bilan.site_id}{" "}<small>{periodLabel}</small></h2></div>
        <span className={`sap-badge ${stAff.cls}`} data-testid="sap-statut">{stAff.label}</span>
      </div>
      <div className="sap-toolbar">
        <button className="sap-btn sap-btn-ghost" onClick={retour} title="Retour à la sélection"><Ico n="back" />Retour</button>
        <button className="sap-btn sap-btn-main" onClick={enregistrer} disabled={!(editable && (ventesOuvert || stocksOuvert || transfertsOuvert))} title={editable ? (ventesOuvert || stocksOuvert || transfertsOuvert ? "Enregistrer les modifications" : "Ouvre un onglet de saisie (ventes, transferts, stocks)") : "Bilan non modifiable"}><Ico n="save" />Enregistrer{nbModifiees > 0 ? ` (${nbModifiees})` : ""}</button>
        <button className="sap-btn" onClick={verifierBilan} disabled={hist || occupe} title={hist ? "Bilan historique : rien à vérifier" : "Contrôler la cohérence sans enregistrer"}><Ico n="search" />Vérifier</button>
        {bilan.statut === "valide" && canManage
          ? <button className="sap-btn" onClick={rouvrir} disabled={occupe} title="Rouvrir le bilan (Superviseur)"><Ico n="undo" />Rouvrir</button>
          : <button className="sap-btn sap-btn-ok" onClick={valider} disabled={!(canManage && bilan.statut === "brouillon") || occupe} title={canManage ? (bilan.statut === "brouillon" ? "Vérification complète puis verrouillage du bilan" : "Bilan non modifiable") : "Réservé au Superviseur"}><Ico n="shield" />Valider</button>}
        <button className="sap-btn" onClick={() => exporter("excel")} disabled={!exportToExcel || occupe} title={exportToExcel ? "Exporter la synthèse et le détail en Excel" : "Export indisponible"}><Ico n="sheet" />Exporter Excel</button>
        <button className="sap-btn" onClick={() => exporter("pdf")} disabled={!exportToPdf || occupe} title={exportToPdf ? "Exporter la synthèse et le détail en PDF" : "Export indisponible"}><Ico n="pdf" />Exporter PDF</button>
        <button className="sap-btn" onClick={() => exporter("print")} disabled={occupe} title="Imprimer la synthèse et le détail"><Ico n="print" />Imprimer</button>
      </div>
      {!hist && (() => {
        const total = joursDuMois(bilan.period_key).length; const n = ventesCharge ? joursSaisis : joursDb; const pct = total ? Math.min(100, Math.round((n / total) * 100)) : 0;
        return (
          <div className={`sap-progress${total > 0 && n >= total ? " done" : ""}`} data-testid="sap-progress">
            <span className="sap-progress-txt" data-testid="sap-progress-txt">{n} / {total} jours saisis</span>
            <div className="sap-progress-bar" role="progressbar" aria-valuemin={0} aria-valuemax={total} aria-valuenow={n} aria-label="Jours de ventes saisis"><div className="sap-progress-fill" style={{ width: `${pct}%` }} /></div>
            <span className="sap-progress-pct">{pct} %</span>
          </div>
        );
      })()}
      <div className="sap-header" data-testid="sap-entete">
        <div className="sap-hrow"><b>Site</b><span>{site?.name || bilan.site_id}</span></div>
        <div className="sap-hrow"><b>Période</b><span>{periodLabel}</span></div>
        <div className="sap-hrow"><b>Statut</b><span className={`sap-badge ${stAff.cls}`}>{stAff.label}</span></div>
        <div className="sap-hrow"><b>Créé par</b><span>{bilan.created_by || "—"} <span className="sap-muted">le {fmtDateHeure(bilan.created_at)}</span></span></div>
        <div className="sap-hrow"><b>Modifié par</b><span>{bilan.modified_by || bilan.created_by || "—"}</span></div>
        <div className="sap-hrow"><b>Dernière modification</b><span>{fmtDateHeure(bilan.modified_at || bilan.created_at)}</span></div>
        {bilan.statut === "valide" && <div className="sap-hrow"><b>Validé par</b><span data-testid="sap-valide-par">{bilan.valide_par || "—"} <span className="sap-muted">le {fmtDateHeure(bilan.valide_le)}</span></span></div>}
      </div>
      {hist && <div className="sap-banner" data-testid="sap-historique">Bilan historique (saisi avec l'ancien module) : affiché en lecture seule, avec les valeurs enregistrées.</div>}
      {bilan.statut === "valide" && <div className="sap-banner-ok" data-testid="sap-valide">Bilan validé et verrouillé : seul le Superviseur peut le rouvrir{canManage ? " (bouton « Rouvrir »)" : ""}.</div>}
      <div className="sap-tabs" role="tablist">
        {[["synthese", "Synthèse", "chart"], ["ventes", "Ventes journalières", "up"], ["transferts", "Transferts", "swap"], ["receptions", "Réceptions", "down"], ["stocks", "Stocks début / fin", "db"]].map(([id, label, ic]) => (
          <button key={id} role="tab" aria-selected={tab === id} className={`sap-tab ${tab === id ? "on" : ""}`} onClick={() => { setTab(id); if (id === "ventes" && !hist) setVentesOuvert(true); if (id === "stocks" && !hist) setStocksOuvert(true); if (id === "transferts" && !hist) setTransfertsOuvert(true); }}><Ico n={ic} s={15} />{label}</button>
        ))}
      </div>
      <div className="sap-body">
        {tab === "synthese" && (
          <>
            <div className="sap-kpis" data-testid="sap-kpis">
              <div className="sap-kpi sap-kpi-debut"><span className="sap-kpi-ic"><Ico n="db" s={22} /></span><div>
                <div className="sap-kpi-label">Stock début 15 °C</div>
                <div className="sap-kpi-val" data-testid="kpi-debut">{fmtL(syn.stockDebut)}{" "}<small>L</small></div>
                <div className="sap-kpi-sub">{syn.stockDebut === null ? "Non disponible" : syn.stockDebutPropose ? "Proposé : fin du mois précédent" : "Saisi"}</div></div></div>
              <div className="sap-kpi sap-kpi-rec"><span className="sap-kpi-ic"><Ico n="down" s={22} /></span><div>
                <div className="sap-kpi-label">Réceptions 15 °C</div>
                <div className="sap-kpi-val" data-testid="kpi-rec">{fmtL(syn.receptions)}{" "}<small>L</small></div>
                <div className="sap-kpi-sub">{hist ? "Saisie manuelle (ancien module)" : `${rec.lignes.length} bon(s) de livraison`}</div></div></div>
              <div className="sap-kpi sap-kpi-ven"><span className="sap-kpi-ic"><Ico n="up" s={22} /></span><div>
                <div className="sap-kpi-label">Ventes 15 °C</div>
                <div className="sap-kpi-val" data-testid="kpi-ven">{fmtL(syn.ventes)}{" "}<small>L</small></div>
                <div className="sap-kpi-sub">{hist ? "Saisie manuelle (ancien module)" : "Total des ventes journalières"}</div></div></div>
              <div className={`sap-kpi ${c ? (c.ecart > 0 ? "sap-kpi-gain" : c.ecart < 0 ? "sap-kpi-perte" : "sap-kpi-neutre") : "sap-kpi-neutre"}`}><span className="sap-kpi-ic"><Ico n={c ? (c.ecart > 0 ? "gain" : c.ecart < 0 ? "perte" : "egal") : "egal"} s={22} /></span><div>
                <div className="sap-kpi-label">Écart</div>
                <div className="sap-kpi-val" data-testid="kpi-ecart">{c ? fmtSigned(c.ecart) : "—"}{" "}<small>L</small></div>
                <div className="sap-kpi-sub" data-testid="kpi-ecart-pct">{c ? `${Number.isFinite(c.ecartPct) ? `${c.ecartPct > 0 ? "+" : ""}${c.ecartPct.toFixed(2).replace(".", ",")} % des ventes · ` : ""}${c.sens === "gain" ? "Gain" : c.sens === "perte" ? "Perte" : "Aucun écart"}` : "Stock début ou fin manquant"}</div></div></div>
            </div>
            <div className="sap-card">
              <div className="sap-card-h"><Ico n="list" s={17} />Compte de stock</div>
              <div className="sap-card-b">
          <table className="sap-table sap-compte" data-testid="sap-synthese">
            <thead><tr><th>Poste</th><th className="sap-num">Volume à 15 °C (L)</th><th>Remarque</th></tr></thead>
            <tbody>
              {ligne("debut", "Stock début de mois", syn.stockDebut, syn.stockDebut === null ? "Non disponible : pas de bilan du mois précédent" : syn.stockDebutPropose ? "Proposé : stock fin du mois précédent" : "Saisi")}
              {ligne("rec", "+ Réceptions", syn.receptions, hist ? "Saisie manuelle (ancien module)" : `${rec.lignes.length} bon(s) de livraison du registre${rec.nbSansMesure ? ` — ⚠ ${rec.nbSansMesure} sans température/densité (compté en ambiant)` : ""}`)}
              {ligne("ven", "− Ventes", syn.ventes, hist ? "Saisie manuelle (ancien module)" : "Total des ventes journalières (mis à jour par la base)")}
              {!hist && ventesParLigne && ventesParLigne.labels.length > 1 && ventesParLigne.labels.map((lab, k) => ligne(`ven${k}`, <span style={{ paddingLeft: 18 }}>dont {lab}</span>, ventesParLigne.volume15[k], "Ventes de la ligne (à 15 °C)"))}
              {ligne("tr", "± Transferts entre sites", syn.transferts, hist ? "Saisie manuelle (ancien module)" : "Saisie par décade, à l'ambiant (équilibre des stocks) : reçus − envoyés", { signed: true })}
              {ligne("theo", "= Stock théorique fin", c ? c.theorique : syn.stockDebut !== null ? syn.stockDebut + syn.receptions - syn.ventes + syn.transferts : null, "Début + réceptions − ventes ± transferts", { total: true })}
              {ligne("phys", "Stock physique fin (jaugeage)", syn.stockFin, syn.stockFin === null ? "Non saisi" : "Mesuré à la jauge, ramené à 15 °C", { bold: true })}
              {ligne("ecart", "Écart (physique − théorique)", c ? c.ecart : null, c ? (c.sens === "gain" ? "Gain" : c.sens === "perte" ? "Perte" : "Aucun écart") : "Stock début ou stock fin manquant", { total: true, final: true, signed: true, cls: c ? clsSigne(c.ecart) : "" })}
              {ligne("pct", "Écart en % des ventes", c ? c.ecartPct : null, c && c.ecartPct === null ? "Aucune vente sur le mois" : "Écart ÷ ventes du mois", { pct: true, cls: c ? clsSigne(c.ecart) : "" })}
            </tbody>
          </table>
              </div>
            </div>
          </>
        )}
        {tab === "synthese" && (
          <div className="sap-card sap-card-pad" data-testid="sap-obs-bloc">
            <label className="sap-label" htmlFor="sap-obs" style={{ display: "block", marginBottom: 4 }}>Observation du mois</label>
            <textarea id="sap-obs" data-testid="sap-obs" className="sap-input" style={{ width: "100%", maxWidth: 640 }} rows={3} value={obs} disabled={!editable} placeholder={editable ? "Ex. : boni hors freinte, justificatif camion…" : ""} onChange={(e) => setObs(e.target.value)} />
            <p className="sap-muted" style={{ margin: "2px 0 0", fontSize: 11.5 }}>Reprise dans la présentation PowerPoint (diapositive « Observations »). Vide = « RAS. ».</p>
          </div>
        )}
        {tab === "receptions" && (
          <div>
            <div className="sap-card sap-card-b"><table className="sap-table" data-testid="sap-receptions">
              <thead><tr><th>Date</th><th>N° bon</th><th className="sap-num">Quantité ambiante (L)</th><th className="sap-num">Température (°C)</th><th className="sap-num">Densité</th><th className="sap-num">Quantité 15 °C (L)</th></tr></thead>
              <tbody>
                {rec.lignes.length === 0 && <tr><td colSpan={6} className="sap-muted">Aucun bon de livraison de gasoil enregistré pour ce site sur ce mois.</td></tr>}
                {rec.lignes.map((l) => (
                  <tr key={l.id}>
                    <td>{fmtJour(l.date)}</td><td>{l.ref || "—"}</td><td className="sap-num">{fmtL(l.quantity)}</td>
                    <td className="sap-num">{l.temperatureC !== undefined && l.temperatureC !== null ? Number(l.temperatureC).toLocaleString("fr-FR", { maximumFractionDigits: 1 }) : "—"}</td>
                    <td className="sap-num">{l.densiteObservee !== undefined && l.densiteObservee !== null ? Number(l.densiteObservee).toLocaleString("fr-FR", { maximumFractionDigits: 1 }) : "—"}</td>
                    <td className="sap-num">{fmtL(l.volume15)}{!l.converti && <span title="Température ou densité manquante : volume ambiant" className="sap-msg-warn"> ⚠</span>}</td>
                  </tr>
                ))}
                {rec.lignes.length > 0 && <tr className="sap-total"><td colSpan={2}>Total</td><td className="sap-num">{fmtL(rec.totalAmbiant)}</td><td></td><td></td><td className="sap-num" data-testid="rec-total">{fmtL(rec.total15)}</td></tr>}
              </tbody>
            </table></div>
            {hist && Math.abs((syn.receptions || 0) - rec.total15) > 1 && (
              <p className="sap-muted" style={{ fontSize: 12, marginTop: 10 }} data-testid="rec-ecart-hist">
                Pour information : ce bilan historique reprend une saisie manuelle de {fmtL(syn.receptions)} L ; le registre indique {fmtL(rec.total15)} L.
              </p>
            )}
          </div>
        )}
        {ventesOuvert && !hist && (
          <div style={{ display: tab === "ventes" ? "block" : "none" }}>
            <VentesGrille ref={gridRef} bilan={bilan} editable={editable} userName={userName} siteName={site?.name || bilan.site_id} setMsg={setMsg} onDirty={setNbModifVentes} onStats={(n) => { setJoursSaisis(n); setVentesCharge(true); }} onParLigne={setVentesParLigne}
              onSaved={(h) => setBilan((b) => ({ ...b, ...h }))} appendAudit={appendAudit} flash={flash} />
          </div>
        )}
        {transfertsOuvert && !hist && (
          <div style={{ display: tab === "transferts" ? "block" : "none" }}>
            <TransfertsDecades ref={transfertsRef} bilan={bilan} editable={editable} userName={userName} siteName={site?.name || bilan.site_id} movements={movements} setMsg={setMsg}
              onDirty={setNbModifTransferts} onStats={setNbTransferts} onSaved={(h) => setBilan((b) => ({ ...b, ...h }))} appendAudit={appendAudit} flash={flash} />
          </div>
        )}
        {tab === "transferts" && hist && <p className="sap-muted" data-testid="sap-transferts-hist">Bilan historique : les transferts ont été saisis en un total (« {fmtSigned(num(bilan.transferts15) || 0)} L ») avec l'ancien module, sans détail par décade.</p>}
        {tab === "ventes" && hist && <p className="sap-muted" data-testid="sap-ventes-hist">Bilan historique : les ventes ont été saisies en un total mensuel avec l'ancien module (pas de détail journalier).</p>}
        {stocksOuvert && !hist && (
          <div style={{ display: tab === "stocks" ? "block" : "none" }}>
            <StocksPanel ref={stocksRef} bilan={bilan} prec={prec} editable={editable} canManage={canManage} userName={userName} siteName={site?.name || bilan.site_id}
              onDirty={setNbModifStocks} onSaved={(h) => setBilan((b) => ({ ...b, ...h }))} appendAudit={appendAudit} flash={flash} />
          </div>
        )}
        {tab === "stocks" && hist && (
          <p className="sap-muted" data-testid="sap-stocks-hist">Bilan historique : stock fin enregistré directement à 15 °C ({fmtL(num(bilan.stock_fin15))} L), sans détail de volume ambiant, température ni densité.</p>
        )}
      </div>
      {verif && (
        <div className="sap-msgpanel" data-testid="sap-verif">
          <div className="sap-msghead"><span>Résultat de la vérification — {cnt.erreurs} erreur(s) · {cnt.avertissements} avertissement(s)</span><button className="sap-btn" style={{ padding: "2px 8px" }} onClick={() => setVerif(null)}>Fermer</button></div>
          {verif.messages.length === 0 ? <div className="sap-msg-ok" data-testid="verif-ok">✔ Aucune anomalie.</div> : (
            <ul style={{ margin: 0, padding: 0 }}>
              {verif.messages.map((m, k) => (
                <li key={k} role="button" tabIndex={0} className={`sap-msgitem ${m.niveau === "erreur" ? "sap-msg-err" : "sap-msg-warn"}`} data-testid={`verif-${k}`}
                  onClick={() => allerA(m.cible)} onKeyDown={(e) => { if (e.key === "Enter") allerA(m.cible); }} title={m.cible ? "Cliquer pour aller à la cellule concernée" : undefined}>
                  {m.niveau === "erreur" ? "✖" : "⚠"} {m.texte}
                </li>
              ))}
            </ul>
          )}
        </div>
      )}
      <div className="sap-status">
        <span className={msg ? `sap-msg sap-msg-${msg.type}${msg.cible ? " sap-msg-link" : ""}` : "sap-muted"} role="status" data-testid="sap-msg" key={msg ? `${msg.type}-${msg.text}` : "pret"} onClick={() => msg && msg.cible && allerA(msg.cible)} title={msg && msg.cible ? "Cliquer pour aller à la cellule concernée" : undefined}>{msg ? <><span className="sap-msg-ic" aria-hidden="true">{{ ok: "✔", warn: "⚠", err: "✖" }[msg.type]}</span> {msg.text}</> : "Prêt."}</span>
        <span data-testid="sap-barre">Statut : {stAff.label.replace(" 🔒", "")} | {nbLignes} ligne(s) | Dernière modification : {fmtDateHeure(bilan.modified_at || bilan.created_at)}</span>
      </div>
      {docImpression && <DocImpression doc={docImpression} />}
    </div>
  );
}

