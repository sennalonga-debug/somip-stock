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
import React, { useState, useMemo } from "react";
import { supabase } from "./supabaseClient.js";
import { receptionsDuMois, transfertsNetsDuMois, calculerBilan, DENSITE_VENTE_DEFAUT } from "./bilanCalc.js";

const MOIS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];
const pad2 = (n) => String(n).padStart(2, "0");
const COLS = "id, site_id, period_type, period_key, statut, reception15, ventes15, transferts15, stock_fin15, stock_debut_ambiant, stock_debut_temp, stock_debut_densite, stock_debut15, stock_fin_ambiant, stock_fin_temp, stock_fin_densite, densite_vente, ventes_ambiant, commentaire, created_by, created_at, modified_by, modified_at, valide_par, valide_le";
const COLS_PREC = "id, statut, period_key, stock_fin15, stock_fin_ambiant";
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

// Message lisible pour une erreur Supabase (script SQL de l'étape 1 non exécuté, droits, réseau…).
function erreurLisible(e) {
  const t = `${e?.message || ""} ${e?.code || ""}`;
  if (/column .* does not exist|statut|stock_debut|42703/i.test(t)) return "La base n'est pas à jour : exécute le script supabase_bilan_sap_etape1.sql dans Supabase.";
  if (/row-level security|permission denied|42501/i.test(t)) return "Droits insuffisants pour cette opération (site non attribué à ton compte ?).";
  if (/Failed to fetch|NetworkError|network/i.test(t)) return "Connexion impossible : vérifie ta connexion internet puis réessaie.";
  return e?.message || "Erreur inconnue.";
}

const STYLE = `
.sap-root { font-family: "72", "Segoe UI", Arial, sans-serif; font-size: 13px; color: #32363a; background: #fff; border: 1px solid #c9ced3; border-radius: 4px; overflow: hidden; }
.sap-shell { background: #354a5f; color: #fff; padding: 8px 14px; font-size: 14px; font-weight: 600; display: flex; justify-content: space-between; align-items: center; gap: 10px; }
.sap-toolbar { display: flex; gap: 4px; flex-wrap: wrap; padding: 6px 10px; background: #f2f4f5; border-bottom: 1px solid #d9dde0; }
.sap-btn { background: #fff; border: 1px solid #b7bec5; border-radius: 3px; padding: 5px 11px; font-size: 12.5px; color: #0a6ed1; cursor: pointer; font-family: inherit; }
.sap-btn:hover:not(:disabled) { background: #eaf2fb; }
.sap-btn:disabled { color: #9aa3ab; cursor: not-allowed; background: #f7f8f9; }
.sap-btn-main { background: #0a6ed1; border-color: #0a6ed1; color: #fff; font-weight: 600; }
.sap-btn-main:hover:not(:disabled) { background: #085caf; }
.sap-btn-main:disabled { background: #9ec5ec; border-color: #9ec5ec; color: #fff; }
.sap-form { max-width: 520px; margin: 22px auto; padding: 0 16px; }
.sap-row { display: grid; grid-template-columns: 130px 1fr; align-items: center; gap: 10px; margin-bottom: 10px; }
.sap-label { color: #6a6d70; }
.sap-input { border: 1px solid #89919a; border-radius: 3px; padding: 6px 8px; font-size: 13px; background: #fff; font-family: inherit; width: 100%; box-sizing: border-box; }
.sap-header { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 4px 28px; padding: 10px 14px; background: #fafbfc; border-bottom: 1px solid #d9dde0; }
.sap-hrow { display: flex; gap: 8px; }
.sap-hrow b { color: #6a6d70; font-weight: 400; min-width: 128px; }
.sap-badge { display: inline-block; padding: 1px 9px; border-radius: 10px; font-size: 12px; font-weight: 600; }
.sap-st-draft { background: #fff3b8; color: #8a5a00; }
.sap-st-ok { background: #e1f3e8; color: #107e3e; }
.sap-st-hist { background: #e5e7ea; color: #5b6168; }
.sap-banner { padding: 7px 14px; background: #fff3b8; color: #6b4a00; border-bottom: 1px solid #f0dc7a; font-size: 12.5px; }
.sap-tabs { display: flex; gap: 2px; padding: 8px 10px 0; border-bottom: 2px solid #0a6ed1; background: #fff; flex-wrap: wrap; }
.sap-tab { border: 1px solid #d9dde0; border-bottom: none; background: #f2f4f5; padding: 6px 14px; cursor: pointer; font-size: 12.5px; border-radius: 3px 3px 0 0; color: #32363a; font-family: inherit; }
.sap-tab.on { background: #0a6ed1; color: #fff; border-color: #0a6ed1; font-weight: 600; }
.sap-body { padding: 12px 14px 16px; min-height: 240px; overflow-x: auto; }
.sap-table { border-collapse: collapse; width: 100%; font-size: 12.5px; }
.sap-table th { background: #e5e8ea; border: 1px solid #cfd4d9; padding: 5px 9px; text-align: left; font-weight: 600; }
.sap-table td { border: 1px solid #e0e4e7; padding: 4px 9px; }
.sap-table tbody tr:nth-child(even) td { background: #f6f7f8; }
.sap-num { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
.sap-total td { font-weight: 700; background: #e5e8ea !important; border-top: 2px solid #8a929a; }
.sap-gain { color: #107e3e; font-weight: 700; }
.sap-loss { color: #bb0000; font-weight: 700; }
.sap-muted { color: #6a6d70; }
.sap-status { display: flex; justify-content: space-between; gap: 12px; flex-wrap: wrap; padding: 5px 12px; background: #f2f4f5; border-top: 1px solid #d9dde0; font-size: 12px; color: #4b5057; }
.sap-msg { font-weight: 600; }
.sap-msg-ok { color: #107e3e; } .sap-msg-warn { color: #8a5a00; } .sap-msg-err { color: #bb0000; }
@media print { .sap-toolbar { display: none; } }
`;

export default function BilansMatieresSap({ sites = [], movements = [], canManage = false, canEdit = false, assignedSiteIds = [], userName = "", appendAudit, flash }) {
  const selectableSites = useMemo(() => sites.filter((s) => !s.isMobile && (!assignedSiteIds?.length || assignedSiteIds.includes(s.id))), [sites, assignedSiteIds]);
  const d0 = new Date(); const dPrev = new Date(d0.getFullYear(), d0.getMonth() - 1, 1); // par défaut : le mois qui vient de se terminer
  const [sel, setSel] = useState({ siteId: "", month: dPrev.getMonth() + 1, year: dPrev.getFullYear() });
  const [screen, setScreen] = useState("select");
  const [tab, setTab] = useState("synthese");
  const [loading, setLoading] = useState(false);
  const [msg, setMsg] = useState(null); // { type: "ok" | "warn" | "err", text }
  const [bilan, setBilan] = useState(null);
  const [prec, setPrec] = useState(null);
  const siteId = sel.siteId || selectableSites[0]?.id || "";
  const site = sites.find((s) => s.id === (bilan?.site_id || siteId));
  const periodKey = bilan ? bilan.period_key : `${sel.year}-${pad2(sel.month)}`;
  const [pY, pM] = periodKey.split("-").map(Number);
  const periodLabel = `${MOIS[(pM || 1) - 1]} ${pY}`;

  const rec = useMemo(() => (bilan ? receptionsDuMois(movements, bilan.site_id, bilan.period_key) : null), [movements, bilan]);
  const tr = useMemo(() => (bilan ? transfertsNetsDuMois(movements, bilan.site_id, bilan.period_key) : null), [movements, bilan]);

  const executer = async () => {
    if (!siteId || loading) return;
    setLoading(true); setMsg(null);
    const key = `${sel.year}-${pad2(sel.month)}`;
    const prevKey = sel.month === 1 ? `${sel.year - 1}-12` : `${sel.year}-${pad2(sel.month - 1)}`;
    const siteName = sites.find((s) => s.id === siteId)?.name || siteId;
    try {
      const [cur, prev, param] = await Promise.all([
        supabase.from("bilan_matieres").select(COLS).eq("site_id", siteId).eq("period_type", "mensuel").eq("period_key", key).maybeSingle(),
        supabase.from("bilan_matieres").select(COLS_PREC).eq("site_id", siteId).eq("period_type", "mensuel").eq("period_key", prevKey).maybeSingle(),
        supabase.from("bilan_parametres").select("valeur").eq("cle", "densite_vente").maybeSingle(),
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
      setBilan(row); setPrec(prev.error ? null : prev.data || null); setTab("synthese"); setScreen("bilan");
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
  const retour = () => { setScreen("select"); setBilan(null); setPrec(null); setMsg(null); };

  // ---------------- Calculs de la synthèse (tout à 15 °C) ----------------
  const hist = bilan?.statut === "historique";
  let syn = null;
  if (bilan) {
    const stockDebutSaisi = num(bilan.stock_debut15);
    const precFin = prec && (prec.statut === "historique" || prec.stock_fin_ambiant !== null) ? num(prec.stock_fin15) : null;
    const stockDebut = stockDebutSaisi !== null ? stockDebutSaisi : precFin;
    const stockFin = hist ? num(bilan.stock_fin15) : (bilan.stock_fin_ambiant !== null && bilan.stock_fin_ambiant !== undefined ? num(bilan.stock_fin15) : null);
    const receptions = hist ? num(bilan.reception15) : rec.total15;
    const ventes = num(bilan.ventes15) || 0;
    const transferts = hist ? num(bilan.transferts15) || 0 : tr.net15;
    const complet = stockDebut !== null && stockFin !== null;
    const calc = complet ? calculerBilan({ stockDebut15: stockDebut, receptions15: receptions, ventes15: ventes, transferts15: transferts, stockFin15: stockFin }) : null;
    syn = { stockDebut, stockDebutPropose: stockDebutSaisi === null && precFin !== null, stockFin, receptions, ventes, transferts, calc, complet };
  }
  const lignesSynthese = syn ? 8 : 0;
  const nbLignes = tab === "synthese" ? lignesSynthese : tab === "receptions" ? rec?.lignes.length || 0 : 0;
  const st = bilan ? STATUTS[bilan.statut] || STATUTS.historique : null;
  const bientot = "Disponible aux étapes suivantes";

  // ---------------- Écran de sélection ----------------
  if (screen === "select") {
    return (
      <div className="sap-root" data-testid="sap-select" onKeyDown={(e) => { if (e.key === "F8") { e.preventDefault(); executer(); } }}>
        <style>{STYLE}</style>
        <div className="sap-shell"><span>Bilan matières — Sélection</span><span style={{ fontWeight: 400, fontSize: 12 }}>Transaction BILMAT</span></div>
        <div className="sap-toolbar">
          <button className="sap-btn sap-btn-main" onClick={executer} disabled={loading || !siteId} title="Exécuter (F8)">{loading ? "Chargement…" : "Exécuter (F8)"}</button>
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
        <div className="sap-status">
          <span className={msg ? `sap-msg sap-msg-${msg.type}` : "sap-muted"} role="status" data-testid="sap-msg">
            {msg ? `${{ ok: "✔", warn: "⚠", err: "✖" }[msg.type]} ${msg.text}` : "Prêt."}
          </span>
          <span>{selectableSites.length} site(s) disponible(s)</span>
        </div>
      </div>
    );
  }

  // ---------------- Écran du bilan ----------------
  const ligne = (key, label, val, remarque, opts = {}) => (
    <tr key={key} className={opts.total ? "sap-total" : ""}>
      <td style={{ fontWeight: opts.bold ? 600 : 400 }}>{label}</td>
      <td className={`sap-num ${opts.cls || ""}`} data-testid={`syn-${key}`}>{opts.pct ? (Number.isFinite(val) ? `${val > 0 ? "+" : ""}${val.toFixed(2).replace(".", ",")} %` : "—") : opts.signed ? fmtSigned(val) : fmtL(val)}</td>
      <td className="sap-muted">{remarque}</td>
    </tr>
  );
  const c = syn.calc;
  return (
    <div className="sap-root" data-testid="sap-bilan">
      <style>{STYLE}</style>
      <div className="sap-shell"><span>Bilan matières — {site?.name || bilan.site_id} — {periodLabel}</span><span className={`sap-badge ${st.cls}`} data-testid="sap-statut">{st.label}</span></div>
      <div className="sap-toolbar">
        <button className="sap-btn" onClick={retour} title="Retour à la sélection">◀ Retour</button>
        {["Enregistrer", "Vérifier", "Valider", "Exporter Excel", "Exporter PDF", "Imprimer"].map((b) => <button key={b} className="sap-btn" disabled title={bientot}>{b}</button>)}
      </div>
      <div className="sap-header" data-testid="sap-entete">
        <div className="sap-hrow"><b>Site</b><span>{site?.name || bilan.site_id}</span></div>
        <div className="sap-hrow"><b>Période</b><span>{periodLabel}</span></div>
        <div className="sap-hrow"><b>Statut</b><span className={`sap-badge ${st.cls}`}>{st.label}</span></div>
        <div className="sap-hrow"><b>Créé par</b><span>{bilan.created_by || "—"} <span className="sap-muted">le {fmtDateHeure(bilan.created_at)}</span></span></div>
        <div className="sap-hrow"><b>Modifié par</b><span>{bilan.modified_by || bilan.created_by || "—"}</span></div>
        <div className="sap-hrow"><b>Dernière modification</b><span>{fmtDateHeure(bilan.modified_at || bilan.created_at)}</span></div>
      </div>
      {hist && <div className="sap-banner" data-testid="sap-historique">Bilan historique (saisi avec l'ancien module) : affiché en lecture seule, avec les valeurs enregistrées.</div>}
      <div className="sap-tabs" role="tablist">
        {[["synthese", "Synthèse"], ["ventes", "Ventes journalières"], ["receptions", "Réceptions"], ["stocks", "Stocks début / fin"]].map(([id, label]) => (
          <button key={id} role="tab" aria-selected={tab === id} className={`sap-tab ${tab === id ? "on" : ""}`} onClick={() => setTab(id)}>{label}</button>
        ))}
      </div>
      <div className="sap-body">
        {tab === "synthese" && (
          <table className="sap-table" data-testid="sap-synthese">
            <thead><tr><th>Poste</th><th className="sap-num">Volume à 15 °C (L)</th><th>Remarque</th></tr></thead>
            <tbody>
              {ligne("debut", "Stock début de mois", syn.stockDebut, syn.stockDebut === null ? "Non disponible : pas de bilan du mois précédent" : syn.stockDebutPropose ? "Proposé : stock fin du mois précédent" : "Saisi")}
              {ligne("rec", "+ Réceptions", syn.receptions, hist ? "Saisie manuelle (ancien module)" : `${rec.lignes.length} bon(s) de livraison du registre${rec.nbSansMesure ? ` — ⚠ ${rec.nbSansMesure} sans température/densité (compté en ambiant)` : ""}`)}
              {ligne("ven", "− Ventes", syn.ventes, hist ? "Saisie manuelle (ancien module)" : "Total des ventes journalières (mis à jour par la base)")}
              {ligne("tr", "± Transferts nets", syn.transferts, hist ? "Saisie manuelle (ancien module)" : `${tr.lignes.length} transfert(s) entre sites`, { signed: true })}
              {ligne("theo", "= Stock théorique fin", c ? c.theorique : syn.stockDebut !== null ? syn.stockDebut + syn.receptions - syn.ventes + syn.transferts : null, "Début + réceptions − ventes ± transferts", { total: true })}
              {ligne("phys", "Stock physique fin (jaugeage)", syn.stockFin, syn.stockFin === null ? "Non saisi" : "Mesuré à la jauge, ramené à 15 °C", { bold: true })}
              {ligne("ecart", "Écart (physique − théorique)", c ? c.ecart : null, c ? (c.sens === "gain" ? "Gain" : c.sens === "perte" ? "Perte" : "Aucun écart") : "Stock début ou stock fin manquant", { total: true, signed: true, cls: c ? clsSigne(c.ecart) : "" })}
              {ligne("pct", "Écart en % des ventes", c ? c.ecartPct : null, c && c.ecartPct === null ? "Aucune vente sur le mois" : "Écart ÷ ventes du mois", { pct: true, cls: c ? clsSigne(c.ecart) : "" })}
            </tbody>
          </table>
        )}
        {tab === "receptions" && (
          <div>
            <table className="sap-table" data-testid="sap-receptions">
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
            </table>
            {hist && Math.abs((syn.receptions || 0) - rec.total15) > 1 && (
              <p className="sap-muted" style={{ fontSize: 12, marginTop: 10 }} data-testid="rec-ecart-hist">
                Pour information : ce bilan historique reprend une saisie manuelle de {fmtL(syn.receptions)} L ; le registre indique {fmtL(rec.total15)} L.
              </p>
            )}
          </div>
        )}
        {(tab === "ventes" || tab === "stocks") && (
          <p className="sap-muted" data-testid="sap-bientot">{tab === "ventes" ? "La grille des ventes journalières (une ligne par jour, saisie au clavier, collage depuis Excel)" : "La saisie des stocks de début et de fin de mois"} sera disponible à l'étape suivante du déploiement.</p>
        )}
      </div>
      <div className="sap-status">
        <span className={msg ? `sap-msg sap-msg-${msg.type}` : "sap-muted"} role="status" data-testid="sap-msg">{msg ? `${{ ok: "✔", warn: "⚠", err: "✖" }[msg.type]} ${msg.text}` : "Prêt."}</span>
        <span data-testid="sap-barre">Statut : {st.label.replace(" 🔒", "")} | {nbLignes} ligne(s) | Dernière modification : {fmtDateHeure(bilan.modified_at || bilan.created_at)}</span>
      </div>
    </div>
  );
}

