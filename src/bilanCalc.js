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
  if (Number.isNaN(v) || Number.isNaN(t)) return { saisi: true, ok: false, erreur: "Valeur non numérique" };
  if (v !== null && v < 0) return { saisi: true, ok: false, erreur: "Volume négatif" };
  if (v === null) return { saisi: true, ok: false, erreur: "Volume obligatoire" };
  if (t === null) return { saisi: true, ok: false, erreur: "Température obligatoire" };
  const c = convertir15({ ambiant: v, temp: t, densite: densiteVente });
  if (!c.ok) return { saisi: true, ok: false, erreur: c.erreur };
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
  if (Number.isNaN(e) || Number.isNaN(r)) return { saisi: true, ok: false, erreur: "Valeur non numérique" };
  if ((e !== null && e < 0) || (r !== null && r < 0)) return { saisi: true, ok: false, erreur: "Volume négatif" };
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
