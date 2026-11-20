// =====================================================================
// SOMIP Stock — Bilan matières (refonte style SAP) : CALCULS
//
// Fonctions pures (aucun accès base, aucun écran) : elles sont utilisées par l'écran de saisie,
// les exports Excel/PDF et la présentation PowerPoint, pour que tous affichent les mêmes chiffres.
// Toute conversion à 15 °C passe par correctVolumeTo15 (vcf15.js), la même fonction que le module
// « Correction 15 °C » : la densité observée à la température du produit est d'abord ramenée à
// la densité à 15 °C, puis le coefficient de correction est appliqué.
// =====================================================================
import { correctVolumeTo15, VCF_MIN_DENSITY, VCF_MAX_DENSITY } from "./vcf15.js";

export const DENSITE_VENTE_DEFAUT = 820; // densité observée des ventes (paramètre modifiable par le Superviseur)
export const TEMP_MIN_PLAUSIBLE = 15;
export const TEMP_MAX_PLAUSIBLE = 45;

// « 2026-10 » -> ["2026-10-01", …, "2026-10-31"] (années bissextiles comprises).
export function joursDuMois(periodKey) {
  const [y, m] = String(periodKey).split("-").map(Number);
  if (!y || !m || m < 1 || m > 12) return [];
  const n = new Date(Date.UTC(y, m, 0)).getUTCDate();
  const mm = String(m).padStart(2, "0");
  return Array.from({ length: n }, (_, i) => `${y}-${mm}-${String(i + 1).padStart(2, "0")}`);
}

// Nombre saisi au clavier ou collé depuis Excel : « 1 234,5 », « 1234.5 », « 1.234,5 », « 1,234.5 ».
// Vide -> null (non saisi) ; texte non numérique -> NaN (erreur à signaler).
export function lireNombre(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : NaN;
  let s = String(v).replace(/[\s\u00a0\u202f]/g, "");
  if (s === "") return null;
  const lastComma = s.lastIndexOf(","), lastDot = s.lastIndexOf(".");
  if (lastComma >= 0 && lastDot >= 0) {
    // les deux séparateurs : le dernier est le séparateur décimal, l'autre sépare les milliers
    s = lastComma > lastDot ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
  } else if (lastComma >= 0) {
    s = s.replace(",", ".");
  }
  if (!/^[+-]?(\d+\.?\d*|\.\d+)$/.test(s)) return NaN;
  return Number(s);
}

// Conversion d'une mesure (volume ambiant, température, densité observée) à 15 °C.
export function convertir15({ ambiant, temp, densite }) {
  if (!Number.isFinite(ambiant) || ambiant < 0) return { ok: false, erreur: "Volume invalide" };
  if (!Number.isFinite(temp)) return { ok: false, erreur: "Température obligatoire" };
  if (!Number.isFinite(densite) || densite < VCF_MIN_DENSITY || densite > VCF_MAX_DENSITY) {
    return { ok: false, erreur: `Densité hors plage (${VCF_MIN_DENSITY} à ${VCF_MAX_DENSITY} kg/m³)` };
  }
  if (ambiant === 0) return { ok: true, volume15: 0, densite15: null, vcf: null, methode: null };
  const r = correctVolumeTo15({ volumeAmbiant: ambiant, tempC: temp, densiteObservee: densite });
  if (!r) return { ok: false, erreur: "Conversion à 15 °C impossible" };
  return { ok: true, volume15: r.volume15, densite15: r.densite15, vcf: r.vcf, methode: r.methode };
}

// Une ligne de ventes journalières : volume vendu (L ambiants) + température du jour ; densité de vente fixe.
// `saisi` = false si la ligne est entièrement vide (jour sans saisie : avertissement, pas une erreur).
export function calculerVente({ volume, temp }, densiteVente = DENSITE_VENTE_DEFAUT) {
  const v = lireNombre(volume), t = lireNombre(temp);
  if (v === null && t === null) return { saisi: false, ok: true };
  if (Number.isNaN(v) || Number.isNaN(t)) return { saisi: true, ok: false, erreur: "Valeur non numérique", champ: Number.isNaN(v) ? "volume" : "temp" };
  if (v !== null && v < 0) return { saisi: true, ok: false, erreur: "Volume négatif", champ: "volume" };
  if (v === null) return { saisi: true, ok: false, erreur: "Volume obligatoire", champ: "volume" };
  if (t === null) return { saisi: true, ok: false, erreur: "Température obligatoire", champ: "temp" };
  const c = convertir15({ ambiant: v, temp: t, densite: densiteVente });
  if (!c.ok) return { saisi: true, ok: false, erreur: c.erreur, champ: "volume" };
  const hors = t < TEMP_MIN_PLAUSIBLE || t > TEMP_MAX_PLAUSIBLE;
  return {
    saisi: true, ok: true, volume: v, temp: t, densiteObs: densiteVente, densite15: c.densite15, vcf: c.vcf, volume15: c.volume15, methode: c.methode,
    avertissement: hors ? `Température hors plage plausible (${TEMP_MIN_PLAUSIBLE} à ${TEMP_MAX_PLAUSIBLE} °C)` : undefined,
  };
}

// Totaux des ventes du mois à partir des lignes calculées (calculerVente). Les lignes en erreur ne comptent pas.
export function totauxVentes(lignesCalculees) {
  let ambiant = 0, volume15 = 0, joursSaisis = 0, erreurs = 0;
  for (const l of lignesCalculees) {
    if (!l || !l.saisi) continue;
    if (!l.ok) { erreurs += 1; continue; }
    joursSaisis += 1; ambiant += l.volume; volume15 += l.volume15;
  }
  return { ambiant, volume15, joursSaisis, erreurs };
}

// ---- Ventes en plusieurs LIGNES (ex. Gare Traction : Comilog / GSEZ ; Setrag FCV : RUR / Hors RUR) ----
// Une journée = UNE température commune + un volume par ligne. Un volume laissé vide vaut 0 tant qu'au moins un volume est saisi.
// Même densité et même température pour toute la journée : le coefficient est commun, chaque ligne est ramenée à 15 °C au même coefficient.
// Avec une seule ligne, le résultat est celui de calculerVente (plus les tableaux volumes / volumes15).
const champVolume = (k) => (k === 0 ? "volume" : `volume${k}`);
export function calculerJourVentes({ temp, volumes }, densiteVente = DENSITE_VENTE_DEFAUT) {
  const vols = volumes && volumes.length ? volumes : [""];
  if (vols.length === 1) {
    const r = calculerVente({ volume: vols[0], temp }, densiteVente);
    if (!r.saisi || !r.ok) return r;
    return { ...r, volumes: [r.volume], volumes15: [r.volume15] };
  }
  const t = lireNombre(temp), vs = vols.map(lireNombre);
  if (t === null && vs.every((v) => v === null)) return { saisi: false, ok: true };
  const iNaN = vs.findIndex((v) => Number.isNaN(v));
  if (Number.isNaN(t) || iNaN >= 0) return { saisi: true, ok: false, erreur: "Valeur non numérique", champ: iNaN >= 0 ? champVolume(iNaN) : "temp" };
  const iNeg = vs.findIndex((v) => v !== null && v < 0);
  if (iNeg >= 0) return { saisi: true, ok: false, erreur: "Volume négatif", champ: champVolume(iNeg) };
  if (vs.every((v) => v === null)) return { saisi: true, ok: false, erreur: "Volume obligatoire (au moins une ligne)", champ: "volume" };
  if (t === null) return { saisi: true, ok: false, erreur: "Température obligatoire", champ: "temp" };
  const volumesNum = vs.map((v) => (v === null ? 0 : v));
  const total = volumesNum.reduce((a, c) => a + c, 0);
  const base = calculerVente({ volume: total, temp: t }, densiteVente);
  if (!base.ok) return { saisi: true, ok: false, erreur: base.erreur, champ: "volume" };
  return { ...base, volume: total, volumes: volumesNum, volumes15: volumesNum.map((v) => v * base.vcf) };
}

// Totaux par ligne (ambiant et 15 °C) sur les journées calculées par calculerJourVentes.
export function totauxParLigne(joursCalcules, nbLignes) {
  const ambiant = Array(nbLignes).fill(0), volume15 = Array(nbLignes).fill(0);
  for (const l of joursCalcules) {
    if (!l || !l.saisi || !l.ok || !l.volumes) continue;
    l.volumes.forEach((v, k) => { ambiant[k] += v; volume15[k] += l.volumes15[k]; });
  }
  return { ambiant, volume15 };
}

// Le compte de stock du mois, tout à 15 °C.
//   Stock théorique fin = Stock début + Réceptions − Ventes (+ Transferts nets, 0 s'il n'y en a pas)
//   Écart               = Stock physique fin − Stock théorique fin   (> 0 : gain, < 0 : perte)
//   Écart en %          = Écart ÷ Ventes du mois × 100              (null si aucune vente)
export function calculerBilan({ stockDebut15, receptions15, ventes15, transferts15 = 0, stockFin15 }) {
  const n = (x) => (Number.isFinite(x) ? x : 0);
  const theorique = n(stockDebut15) + n(receptions15) - n(ventes15) + n(transferts15);
  const ecart = n(stockFin15) - theorique;
  return { theorique, ecart, ecartPct: n(ventes15) > 0 ? (ecart / n(ventes15)) * 100 : null, sens: ecart > 0 ? "gain" : ecart < 0 ? "perte" : "nul" };
}

// Volume « officiel » d'un mouvement à 15 °C : celui enregistré avec le bon (registre des bons de livraison) ;
// à défaut (température ou densité non saisies), le volume ambiant.
const volume15Mouvement = (m) => (m.volumeCorrige15 !== undefined && m.volumeCorrige15 !== null ? m.volumeCorrige15 : m.quantity);
const estGasoil = (m) => (m.product || "gasoil") === "gasoil";

// Réceptions du site et du mois, reprises du registre des bons de livraison (aucune ressaisie).
export function receptionsDuMois(movements, siteId, periodKey) {
  const lignes = movements
    .filter((m) => m.type === "reception" && m.siteId === siteId && estGasoil(m) && String(m.date).startsWith(`${periodKey}-`))
    .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : (a.createdAt || "").localeCompare(b.createdAt || "")))
    .map((m) => ({
      id: m.id, date: m.date, ref: m.ref || "", quantity: m.quantity, temperatureC: m.temperatureC, densiteObservee: m.densiteObservee,
      volume15: volume15Mouvement(m), converti: m.volumeCorrige15 !== undefined && m.volumeCorrige15 !== null,
    }));
  return { lignes, total15: lignes.reduce((a, l) => a + l.volume15, 0), totalAmbiant: lignes.reduce((a, l) => a + l.quantity, 0), nbSansMesure: lignes.filter((l) => !l.converti).length };
}

// Transferts nets du site et du mois (entrées − sorties), à 15 °C quand la mesure existe.
export function transfertsNetsDuMois(movements, siteId, periodKey) {
  const lignes = movements
    .filter((m) => m.type === "transfert" && m.siteId === siteId && estGasoil(m) && String(m.date).startsWith(`${periodKey}-`))
    .map((m) => ({ id: m.id, date: m.date, siteLie: m.relatedSiteId, sens: m.delta < 0 ? -1 : 1, volume15: volume15Mouvement(m), converti: m.volumeCorrige15 !== undefined && m.volumeCorrige15 !== null }));
  return { lignes, net15: lignes.reduce((a, l) => a + l.sens * l.volume15, 0), nbSansMesure: lignes.filter((l) => !l.converti).length };
}

// ---------------------------------------------------------------------
// Transferts ENTRE SITES, par DÉCADE (volumes AMBIANTS, uniquement pour équilibrer les stocks)
//   Décade 1 = du 01 au 10, décade 2 = du 11 au 20, décade 3 = du 21 à la fin du mois.
//   « Envoyés » = volumes envoyés à un autre site (sortent du stock) ;
//   « Reçus »   = volumes reçus d'un autre site (entrent dans le stock).
//   Aucune température ni conversion à 15 °C : seules les ventes sont converties.
// ---------------------------------------------------------------------
export function decadesDuMois(periodKey) {
  const jours = joursDuMois(periodKey);
  if (!jours.length) return [];
  return [{ n: 1, debut: jours[0], fin: jours[9] }, { n: 2, debut: jours[10], fin: jours[19] }, { n: 3, debut: jours[20], fin: jours[jours.length - 1] }];
}

// Une décade saisie : { envoyes, recus } (texte ou nombre). Les deux vides -> décade non saisie (ignorée).
// Un seul des deux rempli : l'autre vaut 0. Négatif ou non numérique -> erreur. Effet net sur le stock = reçus − envoyés.
export function calculerDecade({ envoyes, recus }) {
  const e = lireNombre(envoyes), r = lireNombre(recus);
  if (e === null && r === null) return { saisi: false, ok: true };
  if (Number.isNaN(e) || Number.isNaN(r)) return { saisi: true, ok: false, erreur: "Valeur non numérique", champ: Number.isNaN(e) ? "envoyes" : "recus" };
  if ((e !== null && e < 0) || (r !== null && r < 0)) return { saisi: true, ok: false, erreur: "Volume négatif", champ: e !== null && e < 0 ? "envoyes" : "recus" };
  const ee = e === null ? 0 : e, rr = r === null ? 0 : r;
  return { saisi: true, ok: true, envoyes: ee, recus: rr, net: rr - ee };
}

// Effet net sur le stock du mois (reçus − envoyés), à partir des décades calculées (calculerDecade).
export function totauxDecades(decadesCalculees) {
  let envoyes = 0, recus = 0, nb = 0, erreurs = 0;
  for (const d of decadesCalculees) {
    if (!d || !d.saisi) continue;
    if (!d.ok) { erreurs += 1; continue; }
    nb += 1; envoyes += d.envoyes; recus += d.recus;
  }
  return { envoyes, recus, net: recus - envoyes, nb, erreurs };
}

// POUR INFORMATION (jamais compté) : ce que l'application a déjà enregistré pour ce site, par décade —
// transferts directs entre sites (« transfert » : envoyé = delta négatif, reçu = delta positif) et chargements
// de camions laitiers (« sortie_camion »). Sert à comparer avec ce que l'on saisit.
export function transfertsParDecade(movements, siteId, periodKey) {
  return decadesDuMois(periodKey).map((d) => {
    const dans = (m) => m.siteId === siteId && estGasoil(m) && m.date >= d.debut && m.date <= d.fin;
    const tr = movements.filter((m) => m.type === "transfert" && dans(m));
    const env = tr.filter((m) => (m.delta || 0) < 0), rec = tr.filter((m) => (m.delta || 0) > 0);
    const ch = movements.filter((m) => m.type === "sortie_camion" && dans(m));
    const somme = (a) => a.reduce((s, m) => s + Math.abs(m.quantity || 0), 0);
    return { n: d.n, debut: d.debut, fin: d.fin, envoyes: somme(env), nbEnvoyes: env.length, recus: somme(rec), nbRecus: rec.length, chargements: somme(ch), nbChargements: ch.length };
  });
}

// ---------------------------------------------------------------------
// Vérification du bilan (bouton « Vérifier » et contrôle avant « Valider »)
//   Un message = { niveau: "erreur" | "avert", texte, cible } ; cible = { onglet, champ } désigne la cellule à atteindre
//   en cliquant sur le message (champ = identifiant de la cellule), ou null.
//   Erreur (rouge) : valeur non numérique, volume négatif, ligne incomplète, stock obligatoire manquant.
//   Avertissement (jaune) : température hors plage plausible, jour sans saisie, mois non terminé, réception sans mesure.
// ---------------------------------------------------------------------
const fjMsg = (j) => `${j.slice(8, 10)}/${j.slice(5, 7)}/${j.slice(0, 4)}`;

// lignes = [{ jour, calc }] (calc = calculerVente). Les jours à venir (après `aujourdhui`) ne sont pas « sans saisie ».
export function messagesVentes(lignes, aujourdhui) {
  const out = [];
  for (const { jour, calc } of lignes) {
    if (calc.saisi && !calc.ok) out.push({ niveau: "erreur", texte: `Ventes ${fjMsg(jour)} : ${calc.erreur}`, cible: { onglet: "ventes", champ: `${calc.champ || "volume"}-${jour}` } });
    else if (calc.saisi && calc.avertissement) out.push({ niveau: "avert", texte: `Ventes ${fjMsg(jour)} : ${calc.avertissement}`, cible: { onglet: "ventes", champ: `temp-${jour}` } });
    else if (!calc.saisi && (!aujourdhui || jour <= aujourdhui)) out.push({ niveau: "avert", texte: `Ventes ${fjMsg(jour)} : jour sans saisie`, cible: { onglet: "ventes", champ: `volume-${jour}` } });
  }
  return out;
}

// decades = [{ n, calc }] (calc = calculerDecade)
export function messagesTransferts(decades) {
  return decades.filter((d) => d.calc.saisi && !d.calc.ok).map((d) => ({ niveau: "erreur", texte: `Transferts décade ${d.n} : ${d.calc.erreur}`, cible: { onglet: "transferts", champ: `tr-${d.n}-${d.calc.champ || "envoyes"}` } }));
}

export function messagesGeneraux({ periodKey, aujourdhui, nbReceptionsSansMesure }) {
  const out = [];
  const jours = joursDuMois(periodKey);
  if (jours.length && aujourdhui && aujourdhui < jours[jours.length - 1]) {
    out.push({ niveau: "avert", texte: `Le mois n'est terminé qu'au ${fjMsg(jours[jours.length - 1])} : le bilan est peut-être incomplet.`, cible: null });
  }
  if (nbReceptionsSansMesure > 0) out.push({ niveau: "avert", texte: `Réceptions : ${nbReceptionsSansMesure} bon(s) de livraison sans température/densité (comptés en volume ambiant).`, cible: { onglet: "receptions", champ: null } });
  return out;
}

// Erreurs d'abord, puis avertissements ; l'ordre d'arrivée est conservé dans chaque groupe.
export function trierMessages(messages) {
  return [...messages.filter((m) => m.niveau === "erreur"), ...messages.filter((m) => m.niveau !== "erreur")];
}
export const compterMessages = (messages) => ({ erreurs: messages.filter((m) => m.niveau === "erreur").length, avertissements: messages.filter((m) => m.niveau !== "erreur").length });

// ---------------------------------------------------------------------
// Synthèse du bilan (compte de stock à 15 °C) — UNE seule définition, partagée par l'écran, les exports
// Excel / PDF / impression et la présentation PowerPoint, pour qu'ils affichent toujours les mêmes chiffres.
//   bilan = ligne de bilan_matieres ; prec = bilan du mois précédent (stock fin) ; rec = receptionsDuMois(...)
// ---------------------------------------------------------------------
const nombreOuNull = (v) => (v === null || v === undefined || v === "" ? null : Number(v));
export function calculerSynthese({ bilan, prec, rec }) {
  const hist = bilan.statut === "historique";
  const debutSaisi = nombreOuNull(bilan.stock_debut15);
  const precFin = prec && (prec.statut === "historique" || (prec.stock_fin_ambiant !== null && prec.stock_fin_ambiant !== undefined)) ? nombreOuNull(prec.stock_fin15) : null;
  const stockDebut = debutSaisi !== null ? debutSaisi : precFin;
  const stockFin = hist ? nombreOuNull(bilan.stock_fin15) : (bilan.stock_fin_ambiant !== null && bilan.stock_fin_ambiant !== undefined ? nombreOuNull(bilan.stock_fin15) : null);
  const receptions = hist ? nombreOuNull(bilan.reception15) : (rec ? rec.total15 : 0);
  const ventes = nombreOuNull(bilan.ventes15) || 0;
  const transferts = nombreOuNull(bilan.transferts15) || 0;
  const complet = stockDebut !== null && stockFin !== null;
  const calc = complet ? calculerBilan({ stockDebut15: stockDebut, receptions15: receptions, ventes15: ventes, transferts15: transferts, stockFin15: stockFin }) : null;
  return { stockDebut, stockDebutPropose: debutSaisi === null && precFin !== null, stockFin, receptions, ventes, transferts, calc, complet };
}

// Ligne d'un site pour la présentation PowerPoint (même structure que l'ancien module).
//   etat : « absent » (aucun bilan), « incomplet » (stock début ou fin manquant : écart non calculable) ou « ok ».
export function lignePresentation({ site, bilan, prec, movements }) {
  if (!bilan) return { etat: "absent" };
  const rec = receptionsDuMois(movements, bilan.site_id, bilan.period_key);
  const synthese = calculerSynthese({ bilan, prec, rec });
  if (!synthese.complet) return { etat: "incomplet", statut: bilan.statut, synthese };
  return {
    etat: "ok", statut: bilan.statut, synthese,
    ligne: { site: { id: site.id, name: site.name }, reception15: synthese.receptions, ventes15: synthese.ventes, transferts15: synthese.transferts, stockTheorique: synthese.calc.theorique, stockFin15: synthese.stockFin, ecart: synthese.calc.ecart, commentaire: bilan.commentaire || "", photoUrls: bilan.photo_urls || [] },
  };
}

// ---------------------------------------------------------------------
// Document d'export d'un bilan (Excel, PDF et impression) : un seul contenu, deux présentations.
//   • sheets   : feuilles Excel, valeurs NUMÉRIQUES (pour pouvoir calculer dessus) ;
//   • sections : tableaux du PDF / de l'impression, valeurs mises en forme (espaces, signes).
// ---------------------------------------------------------------------
const espaces = (s) => String(s).replace(/[\u202f\u00a0]/g, " ").replace(/\u2212/g, "-");  // les polices PDF standard ne lisent ni l'espace fine ni le vrai signe moins
const fL = (n) => (Number.isFinite(n) ? espaces(Math.round(n).toLocaleString("fr-FR")) : "");
const fLs = (n) => (Number.isFinite(n) ? `${n > 0 ? "+" : ""}${fL(n)}` : "");
const fDec = (n, d) => (Number.isFinite(n) ? espaces(n.toLocaleString("fr-FR", { minimumFractionDigits: d, maximumFractionDigits: d })) : "");
const fJ = (j) => (j ? `${j.slice(8, 10)}/${j.slice(5, 7)}/${j.slice(0, 4)}` : "");
const fDH = (iso) => (iso ? espaces(new Date(iso).toLocaleString("fr-FR", { dateStyle: "short", timeStyle: "short" })) : "");
const slug = (s) => String(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "");

export function construireDocumentBilan({ siteName, periodKey, periodLabel, statutLabel, entete = {}, synthese, observation = "", ventes = null, transferts = null, receptions = { lignes: [], total15: 0, totalAmbiant: 0 }, stocks = null, ventesParLigne = null }) {
  const s = synthese, c = s.calc;
  const nbJours = ventes ? ventes.lignes.length : 0;
  const remDebut = s.stockDebut === null ? "Non disponible" : s.stockDebutPropose ? "Proposé : stock fin du mois précédent" : "Saisi";
  const poste = [
    ["Stock début de mois", s.stockDebut, remDebut],
    ["+ Réceptions", s.receptions, `${receptions.lignes.length} bon(s) de livraison`],
    ["− Ventes", s.ventes, "Total des ventes journalières"],
    ...(ventesParLigne && ventesParLigne.labels.length > 1 ? ventesParLigne.labels.map((lab, k) => [`   dont ${lab}`, ventesParLigne.volume15[k], "Ventes de la ligne"]) : []),
    ["± Transferts entre sites", s.transferts, "Saisie par décade, à l'ambiant : reçus − envoyés"],
    ["= Stock théorique fin", c ? c.theorique : (s.stockDebut !== null ? s.stockDebut + s.receptions - s.ventes + s.transferts : null), "Début + réceptions − ventes ± transferts"],
    ["Stock physique fin (jaugeage)", s.stockFin, s.stockFin === null ? "Non saisi" : "Mesuré à la jauge, ramené à 15 °C"],
    ["Écart (physique − théorique)", c ? c.ecart : null, c ? (c.sens === "gain" ? "Gain" : c.sens === "perte" ? "Perte" : "Aucun écart") : "Stock début ou stock fin manquant"],
  ];
  const pct = c && c.ecartPct !== null ? c.ecartPct : null;
  const signed = new Set(["± Transferts entre sites", "Écart (physique − théorique)"]);
  const sections = [{
    heading: "Synthèse", columns: ["Poste", "Volume à 15 °C (L)", "Remarque"],
    rows: [...poste.map(([p, v, r]) => [p, v === null ? "—" : signed.has(p) ? fLs(v) : fL(v), r]), ["Écart en % des ventes", pct === null ? "—" : `${pct > 0 ? "+" : ""}${fDec(pct, 2)} %`, pct === null ? (c ? "Aucune vente sur le mois" : "") : "Écart ÷ ventes du mois"]],
  }];
  if (observation && observation.trim()) sections.push({ heading: "Observation", columns: ["Observation du mois"], rows: [[observation.trim()]] });
  const sheets = [];
  const infos = [["Site", siteName], ["Période", periodLabel], ["Statut", statutLabel], ["Créé par", entete.creePar || ""], ["Créé le", fDH(entete.creeLe)], ["Modifié par", entete.modifiePar || ""], ["Dernière modification", fDH(entete.modifieLe)]];
  if (entete.validePar) { infos.push(["Validé par", entete.validePar], ["Validé le", fDH(entete.valideLe)]); }
  if (ventes && ventes.densite) infos.push(["Densité de vente (kg/m³)", ventes.densite]);
  if (observation && observation.trim()) infos.push(["Observation", observation.trim()]);
  sheets.push({ name: "Synthèse", rows: [...poste.map(([p, v, r]) => ({ Poste: p, "Volume à 15 °C (L)": v === null ? null : Math.round(v * 100) / 100, Remarque: r })), { Poste: "Écart en % des ventes", "Volume à 15 °C (L)": pct === null ? null : Math.round(pct * 10000) / 10000, Remarque: "Écart ÷ ventes du mois (en %)" }] });
  sheets.push({ name: "Informations", rows: infos.map(([Champ, Valeur]) => ({ Champ, Valeur })) });

  if (stocks) {
    const ligne = (nom, b) => [nom, b.ambiant === null || b.ambiant === undefined ? "" : fL(b.ambiant), b.temp === null || b.temp === undefined ? "" : fDec(b.temp, 1), b.densite === null || b.densite === undefined ? "" : fDec(b.densite, 1), b.d15 === null || b.d15 === undefined ? "" : fDec(b.d15, 1), b.vcf === null || b.vcf === undefined ? "" : fDec(b.vcf, 4), b.v15 === null || b.v15 === undefined ? "" : fL(b.v15)];
    sections.push({ heading: "Stocks début / fin", columns: ["Poste", "Volume ambiant (L)", "Température (°C)", "Densité observée", "Densité à 15 °C", "Coefficient", "Volume 15 °C (L)"], rows: [ligne("Stock début de mois", stocks.deb), ligne("Stock fin de mois (jaugeage)", stocks.fin)] });
    const num2 = (x) => (x === null || x === undefined ? null : x);
    sheets.push({ name: "Stocks début-fin", rows: [["Stock début de mois", stocks.deb], ["Stock fin de mois (jaugeage)", stocks.fin]].map(([Poste, b]) => ({ Poste, "Volume ambiant (L)": num2(b.ambiant), "Température (°C)": num2(b.temp), "Densité observée": num2(b.densite), "Densité à 15 °C": num2(b.d15), Coefficient: num2(b.vcf), "Volume 15 °C (L)": num2(b.v15) })) });
  }
  if (transferts) {
    sections.push({ heading: "Transferts entre sites (volumes ambiants)", columns: ["Décade", "Envoyés à d'autres sites (L)", "Reçus d'autres sites (L)", "Effet sur le stock (L)"],
      rows: transferts.decades.map((d) => [`Décade ${d.n} (du ${fJ(d.debut).slice(0, 5)} au ${fJ(d.fin).slice(0, 5)})`, d.saisi ? fL(d.envoyes) : "", d.saisi ? fL(d.recus) : "", d.saisi ? fLs(d.net) : ""]),
      totalsRow: ["Total du mois", fL(transferts.tot.envoyes), fL(transferts.tot.recus), fLs(transferts.tot.net)] });
    sheets.push({ name: "Transferts", rows: [...transferts.decades.map((d) => ({ "Décade": d.n, "Du": d.debut, "Au": d.fin, "Envoyés (L)": d.saisi ? d.envoyes : null, "Reçus (L)": d.saisi ? d.recus : null, "Effet sur le stock (L)": d.saisi ? d.net : null })), { "Décade": "Total", "Du": "", "Au": "", "Envoyés (L)": transferts.tot.envoyes, "Reçus (L)": transferts.tot.recus, "Effet sur le stock (L)": transferts.tot.net }] });
  }
  sections.push({ heading: "Réceptions (bons de livraison)", columns: ["Date", "N° bon", "Quantité ambiante (L)", "Température (°C)", "Densité", "Quantité 15 °C (L)"],
    rows: receptions.lignes.map((l) => [fJ(l.date), l.ref || "", fL(l.quantity), l.temperatureC === undefined || l.temperatureC === null ? "" : fDec(Number(l.temperatureC), 1), l.densiteObservee === undefined || l.densiteObservee === null ? "" : fDec(Number(l.densiteObservee), 1), fL(l.volume15)]),
    totalsRow: ["Total", "", fL(receptions.totalAmbiant), "", "", fL(receptions.total15)] });
  sheets.push({ name: "Réceptions", rows: [...receptions.lignes.map((l) => ({ Date: l.date, "N° bon": l.ref || "", "Quantité ambiante (L)": l.quantity, "Température (°C)": l.temperatureC === undefined ? null : l.temperatureC, Densité: l.densiteObservee === undefined ? null : l.densiteObservee, "Quantité 15 °C (L)": l.volume15 })), { Date: "Total", "N° bon": "", "Quantité ambiante (L)": receptions.totalAmbiant, "Température (°C)": null, Densité: null, "Quantité 15 °C (L)": receptions.total15 }] });
  if (ventes && ventes.cols && ventes.cols.length > 1) {
    // Plusieurs lignes de vente : une colonne par ligne, une seule température par jour, total du jour.
    const cols = ventes.cols, pl = ventes.totaux.parLigne;
    sections.push({ heading: "Ventes journalières", columns: ["Date", ...cols.map((c) => `${c} (L)`), "Température (°C)", "Total (L)", "Densité à 15 °C", "Coefficient", "Volume 15 °C (L)"],
      rows: ventes.lignes.map((l) => (l.saisi ? [fJ(l.jour), ...l.volumes.map((v) => fL(v)), fDec(l.temp, 1), fL(l.volume), fDec(l.densite15, 1), fDec(l.vcf, 4), fL(l.volume15)] : [fJ(l.jour), ...cols.map(() => ""), "", "", "", "", ""])),
      totalsRow: [`Total (${ventes.totaux.joursSaisis} / ${nbJours} jours)`, ...pl.ambiant.map((v) => fL(v)), "", fL(ventes.totaux.ambiant), "", "", fL(ventes.totaux.volume15)] });
    const r2 = (x) => Math.round(x * 100) / 100;
    sheets.push({ name: "Ventes journalières", rows: [
      ...ventes.lignes.map((l) => {
        const o = { Date: l.jour };
        cols.forEach((c, k) => { o[`${c} (L)`] = l.saisi ? l.volumes[k] : null; });
        o["Total vendu (L)"] = l.saisi ? l.volume : null; o["Température (°C)"] = l.saisi ? l.temp : null; o["Densité"] = ventes.densite;
        o["Densité à 15 °C"] = l.saisi ? Math.round(l.densite15 * 10) / 10 : null; o["Coefficient"] = l.saisi ? l.vcf : null;
        cols.forEach((c, k) => { o[`${c} à 15 °C (L)`] = l.saisi ? r2(l.volumes15[k]) : null; });
        o["Volume 15 °C (L)"] = l.saisi ? r2(l.volume15) : null; return o;
      }),
      (() => {
        const o = { Date: `Total (${ventes.totaux.joursSaisis} / ${nbJours} jours)` };
        cols.forEach((c, k) => { o[`${c} (L)`] = pl.ambiant[k]; });
        o["Total vendu (L)"] = ventes.totaux.ambiant; o["Température (°C)"] = null; o["Densité"] = null; o["Densité à 15 °C"] = null; o["Coefficient"] = null;
        cols.forEach((c, k) => { o[`${c} à 15 °C (L)`] = r2(pl.volume15[k]); });
        o["Volume 15 °C (L)"] = r2(ventes.totaux.volume15); return o;
      })(),
    ] });
  } else if (ventes) {
    sections.push({ heading: "Ventes journalières", columns: ["Date", "Volume vendu (L)", "Température (°C)", "Densité", "Densité à 15 °C", "Coefficient", "Volume 15 °C (L)"],
      rows: ventes.lignes.map((l) => (l.saisi ? [fJ(l.jour), fL(l.volume), fDec(l.temp, 1), fL(ventes.densite), fDec(l.densite15, 1), fDec(l.vcf, 4), fL(l.volume15)] : [fJ(l.jour), "", "", fL(ventes.densite), "", "", ""])),
      totalsRow: [`Total (${ventes.totaux.joursSaisis} / ${nbJours} jours)`, fL(ventes.totaux.ambiant), "", "", "", "", fL(ventes.totaux.volume15)] });
    sheets.push({ name: "Ventes journalières", rows: [...ventes.lignes.map((l) => ({ Date: l.jour, "Volume vendu (L)": l.saisi ? l.volume : null, "Température (°C)": l.saisi ? l.temp : null, Densité: ventes.densite, "Densité à 15 °C": l.saisi ? Math.round(l.densite15 * 10) / 10 : null, Coefficient: l.saisi ? l.vcf : null, "Volume 15 °C (L)": l.saisi ? Math.round(l.volume15 * 100) / 100 : null })), { Date: `Total (${ventes.totaux.joursSaisis} / ${nbJours} jours)`, "Volume vendu (L)": ventes.totaux.ambiant, "Température (°C)": null, Densité: null, "Densité à 15 °C": null, Coefficient: null, "Volume 15 °C (L)": Math.round(ventes.totaux.volume15 * 100) / 100 }] });
  }
  // PDF : les polices standard ne contiennent ni le vrai signe moins (−) ni l'espace fine : tout le texte est normalisé.
  const sectionsPdf = sections.map((x) => ({ heading: espaces(x.heading), columns: x.columns.map(espaces), rows: x.rows.map((r) => r.map(espaces)), totalsRow: x.totalsRow ? x.totalsRow.map(espaces) : undefined }));
  return { filename: `Bilan_matieres_${slug(siteName)}_${periodKey}`, title: espaces(`Bilan matières — ${siteName}`), subtitle: espaces(`Statut : ${statutLabel}`), period: espaces(periodLabel), sections: sectionsPdf, sheets };
}
