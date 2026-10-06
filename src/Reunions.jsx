import { useState, useEffect, useRef } from "react";
import { supabase } from "./supabaseClient";
import {
  Video, Plus, Trash2, Pencil, ExternalLink, Calendar, Clock, X, Check, ArrowLeft,
} from "lucide-react";

/* ------------------------------------------------------------------ */
/* Module Réunions — visioconférence Jitsi Meet, affichée dans l'app.   */
/* Fichier entièrement séparé : ne modifie, n'importe ni ne réutilise   */
/* aucun état ou composant d'App.jsx (à l'exception du client Supabase */
/* partagé, déjà exporté tel quel par supabaseClient.js).               */
/*                                                                      */
/* Reçoit trois props d'App.jsx, toutes déjà calculées là-bas :         */
/*   canManage        -> true pour un Superviseur (créer/modifier/supprimer) */
/*   currentUserName  -> nom affiché, préremplit le nom dans la visio   */
/*   sites            -> liste des sites existants (id, name), pour     */
/*                        cocher les "sites concernés" d'une réunion    */
/* ------------------------------------------------------------------ */

// Petite palette locale, indépendante de celle d'App.jsx (évite toute
// dépendance croisée entre les deux fichiers).
const RC = {
  blue: "#0071BD",
  blueDark: "#00588F",
  orange: "#F16B16",
  ink: "#16212D",
  sub: "#5B6B7A",
  border: "#E4E8EC",
  bg: "#F3F5F8",
  success: "#1E8A5F",
  danger: "#C63C3C",
};

// Choix "rôle" sélectionnables au même titre qu'un site, dans "Sites concernés" — une réunion
// peut concerner les Superviseurs ou les comptes Lecture seule, pas seulement des sites précis.
// Stockés dans le même tableau site_ids, avec un préfixe qui les distingue d'un vrai id de site.
const ROLE_CHOICES = [
  { id: "role:superviseur", label: "Superviseurs" },
  { id: "role:lecture", label: "Lecture seule" },
];

function generateRoomName() {
  // Nom de salle unique, sans information sensible (pas le titre de la réunion) : n'importe
  // qui connaissant ce nom pourrait rejoindre la salle sur meet.jit.si, donc autant qu'il soit
  // imprévisible plutôt que de révéler l'objet de la réunion.
  const rand = (typeof crypto !== "undefined" && crypto.randomUUID)
    ? crypto.randomUUID().replace(/-/g, "").slice(0, 16)
    : Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2);
  return `SOMIP-${rand}`;
}

function combineDateTime(date, heure) {
  return new Date(`${date}T${(heure || "00:00").slice(0, 5)}:00`);
}

function fmtDateFr(dateStr) {
  const d = new Date(`${dateStr}T00:00:00`);
  return d.toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
}

/* ---- Petits composants de mise en forme, autonomes (pas de dépendance à App.jsx) ---- */
function RField({ label, children }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <label style={{ fontSize: 12, fontWeight: 600, color: RC.sub, marginBottom: 5, display: "block" }}>{label}</label>
      {children}
    </div>
  );
}
const rInputStyle = {
  width: "100%", padding: "9px 11px", borderRadius: 8, border: `1.5px solid ${RC.border}`,
  fontSize: 13.5, fontFamily: "inherit", color: RC.ink, background: "#fff", boxSizing: "border-box",
};
const rBtnPrimary = {
  display: "inline-flex", alignItems: "center", gap: 6, padding: "9px 16px", borderRadius: 9,
  fontSize: 13.5, fontWeight: 600, cursor: "pointer", border: "none", color: "#fff",
  background: `linear-gradient(135deg, ${RC.blue}, ${RC.blueDark})`,
};
const rBtnGhost = {
  display: "inline-flex", alignItems: "center", gap: 6, padding: "9px 16px", borderRadius: 9,
  fontSize: 13.5, fontWeight: 600, cursor: "pointer", border: `1px solid ${RC.border}`, color: RC.ink, background: "#fff",
};
const rPanel = { background: "#fff", border: `1px solid ${RC.border}`, borderRadius: 14, boxShadow: "0 1px 2px rgba(16,30,45,0.04), 0 4px 16px rgba(16,30,45,0.05)" };

/* ---- Salle de réunion : iframe Jitsi Meet plein écran dans la carte ---- */
function MeetingRoom({ reunion, currentUserName, onLeave }) {
  const displayName = encodeURIComponent(currentUserName || "Invité SOMIP");
  // Jitsi masque parfois lui-même le bouton de partage d'écran dans un cadre intégré (iframe),
  // par détection de fonctionnalité — même quand le navigateur l'autoriserait. On force sa
  // présence dans la barre d'outils via la configuration d'URL de Jitsi, en plus de l'autorisation
  // "display-capture" déjà donnée au cadre ci-dessous ; les deux sont nécessaires.
  const toolbarButtons = ["microphone", "camera", "desktop", "chat", "raisehand", "tileview", "fullscreen", "hangup"];
  const configParams = [
    "config.prejoinPageEnabled=true",
    `config.toolbarButtons=${encodeURIComponent(JSON.stringify(toolbarButtons))}`,
    "config.disableDeepLinking=true",
  ].join("&");
  const src = `https://meet.jit.si/${encodeURIComponent(reunion.room_name)}#userInfo.displayName=%22${displayName}%22&${configParams}`;

  return (
    <div className="somip-fade">
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
        <button onClick={onLeave} style={rBtnGhost}><ArrowLeft size={14} /> Retour aux réunions</button>
        <a href={src} target="_blank" rel="noreferrer" style={{ ...rBtnGhost, textDecoration: "none" }}>
          <ExternalLink size={14} /> Ouvrir dans un nouvel onglet
        </a>
      </div>
      <div style={{ ...rPanel, padding: 14 }}>
        <h3 style={{ margin: "0 0 4px", fontSize: 14, display: "flex", alignItems: "center", gap: 8 }}>
          <Video size={15} color={RC.blue} />{reunion.titre}
        </h3>
        <p style={{ margin: "0 0 12px", fontSize: 12, color: RC.sub }}>
          Ton nom (« {currentUserName || "Invité SOMIP"} ») est déjà prérempli. Les boutons caméra, micro et partage d'écran se trouvent dans la barre de la visioconférence ci-dessous.
        </p>
        <div style={{ borderRadius: 10, overflow: "hidden", border: `1px solid ${RC.border}` }}>
          <iframe
            key={reunion.id}
            title={`Réunion — ${reunion.titre}`}
            src={src}
            allow="camera; microphone; display-capture; fullscreen; autoplay; clipboard-write"
            allowFullScreen
            style={{ width: "100%", height: "min(70vh, 640px)", border: "none", display: "block" }}
          />
        </div>
        <p style={{ margin: "12px 0 0", fontSize: 11, color: RC.sub }}>
          Si la visioconférence ne se charge pas dans ce cadre (réseau d'entreprise restrictif, par exemple), utilise « Ouvrir dans un nouvel onglet » ci-dessus.
        </p>
      </div>
    </div>
  );
}

/* ---- Formulaire de création / modification (Superviseur uniquement) ---- */
function MeetingForm({ initial, sites, onCancel, onSubmit }) {
  const [titre, setTitre] = useState(initial?.titre || "");
  const [date, setDate] = useState(initial?.date || new Date().toISOString().slice(0, 10));
  const [heure, setHeure] = useState(initial?.heure ? initial.heure.slice(0, 5) : "09:00");
  const [siteIds, setSiteIds] = useState(initial?.site_ids || []);
  const [saving, setSaving] = useState(false);

  const toggleSite = (id) => setSiteIds((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  const valid = titre.trim() && date && heure;

  const submit = async () => {
    if (!valid || saving) return;
    setSaving(true);
    await onSubmit({ titre: titre.trim(), date, heure, siteIds });
    setSaving(false);
  };

  return (
    <div style={{ ...rPanel, padding: 18, marginBottom: 18 }}>
      <h3 style={{ margin: "0 0 14px", fontSize: 14, display: "flex", alignItems: "center", gap: 8 }}>
        <Plus size={15} color={RC.blue} />{initial ? "Modifier la réunion" : "Nouvelle réunion"}
      </h3>
      <RField label="Titre"><input style={rInputStyle} value={titre} onChange={(e) => setTitre(e.target.value)} placeholder="Ex : Point hebdomadaire Zone Sud-Est" /></RField>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
        <div style={{ flex: "1 1 160px" }}><RField label="Date"><input type="date" style={rInputStyle} value={date} onChange={(e) => setDate(e.target.value)} /></RField></div>
        <div style={{ flex: "1 1 120px" }}><RField label="Heure"><input type="time" style={rInputStyle} value={heure} onChange={(e) => setHeure(e.target.value)} /></RField></div>
      </div>
      <RField label="Sites / profils concernés (optionnel)">
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, padding: "4px 0" }}>
          {sites.map((s) => (
            <label key={s.id} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5, border: `1px solid ${siteIds.includes(s.id) ? RC.blue : RC.border}`, background: siteIds.includes(s.id) ? `${RC.blue}12` : "#fff", borderRadius: 20, padding: "5px 11px", cursor: "pointer" }}>
              <input type="checkbox" checked={siteIds.includes(s.id)} onChange={() => toggleSite(s.id)} style={{ margin: 0 }} />
              {s.name}
            </label>
          ))}
          {sites.length === 0 && <span style={{ fontSize: 12, color: RC.sub }}>Aucun site disponible.</span>}
        </div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, padding: "8px 0 0", marginTop: 6, borderTop: `1px dashed ${RC.border}` }}>
          {ROLE_CHOICES.map((r) => (
            <label key={r.id} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5, border: `1px solid ${siteIds.includes(r.id) ? RC.orange : RC.border}`, background: siteIds.includes(r.id) ? `${RC.orange}14` : "#fff", borderRadius: 20, padding: "5px 11px", cursor: "pointer" }}>
              <input type="checkbox" checked={siteIds.includes(r.id)} onChange={() => toggleSite(r.id)} style={{ margin: 0 }} />
              {r.label}
            </label>
          ))}
        </div>
      </RField>
      <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
        <button style={{ ...rBtnPrimary, opacity: valid && !saving ? 1 : 0.5 }} disabled={!valid || saving} onClick={submit}>
          <Check size={14} /> {saving ? "Enregistrement…" : "Enregistrer"}
        </button>
        <button style={rBtnGhost} onClick={onCancel}><X size={14} /> Annuler</button>
      </div>
    </div>
  );
}

/* ---- Vue principale, exportée : liste des réunions + salle active ---- */
export default function ReunionsView({ canManage, currentUserName, sites = [] }) {
  const [reunions, setReunions] = useState([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState(null);
  const [activeRoom, setActiveRoom] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [showPast, setShowPast] = useState(false);
  const mounted = useRef(true);

  const load = async () => {
    setLoadError(null);
    const { data, error } = await supabase.from("reunions").select("*").order("date", { ascending: true }).order("heure", { ascending: true });
    if (!mounted.current) return;
    if (error) setLoadError("Impossible de charger les réunions : " + error.message);
    else setReunions(data || []);
    setLoading(false);
  };
  useEffect(() => { mounted.current = true; load(); return () => { mounted.current = false; }; }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const now = new Date();
  const upcoming = reunions.filter((r) => combineDateTime(r.date, r.heure).getTime() >= now.getTime() - 2 * 60 * 60 * 1000)
    .sort((a, b) => combineDateTime(a.date, a.heure) - combineDateTime(b.date, b.heure));
  const past = reunions.filter((r) => !upcoming.includes(r)).sort((a, b) => combineDateTime(b.date, b.heure) - combineDateTime(a.date, a.heure));

  const createOrUpdate = async ({ titre, date, heure, siteIds }) => {
    if (editing) {
      const { error } = await supabase.from("reunions").update({ titre, date, heure, site_ids: siteIds }).eq("id", editing.id);
      if (error) { window.alert("Modification impossible : " + error.message); return; }
    } else {
      const { error } = await supabase.from("reunions").insert({ titre, date, heure, site_ids: siteIds, room_name: generateRoomName(), created_by: currentUserName || null });
      if (error) { window.alert("Création impossible : " + error.message); return; }
    }
    setShowForm(false); setEditing(null);
    load();
  };

  const remove = async (r) => {
    if (!window.confirm(`Supprimer la réunion « ${r.titre} » ?`)) return;
    const { error } = await supabase.from("reunions").delete().eq("id", r.id);
    if (error) { window.alert("Suppression impossible : " + error.message); return; }
    load();
  };

  if (activeRoom) {
    return <MeetingRoom reunion={activeRoom} currentUserName={currentUserName} onLeave={() => setActiveRoom(null)} />;
  }

  const siteNames = (ids) => (ids || []).map((id) => {
    const role = ROLE_CHOICES.find((r) => r.id === id);
    if (role) return role.label;
    return sites.find((s) => s.id === id)?.name;
  }).filter(Boolean);

  const MeetingRow = ({ r, isPast }) => (
    <div style={{ ...rPanel, padding: 16, marginBottom: 10, opacity: isPast ? 0.7 : 1 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 14.5, color: RC.ink, marginBottom: 4 }}>{r.titre}</div>
          <div style={{ display: "flex", gap: 14, flexWrap: "wrap", fontSize: 12.5, color: RC.sub }}>
            <span style={{ display: "flex", alignItems: "center", gap: 5 }}><Calendar size={13} />{fmtDateFr(r.date)}</span>
            <span style={{ display: "flex", alignItems: "center", gap: 5 }}><Clock size={13} />{(r.heure || "").slice(0, 5)}</span>
          </div>
          {siteNames(r.site_ids).length > 0 && (
            <div style={{ marginTop: 6, fontSize: 12, color: RC.sub }}>Concerne : {siteNames(r.site_ids).join(", ")}</div>
          )}
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          {!isPast && (
            <button style={rBtnPrimary} onClick={() => setActiveRoom(r)}>
              <Video size={14} /> Rejoindre
            </button>
          )}
          {canManage && (
            <>
              <button title="Modifier" onClick={() => { setEditing(r); setShowForm(true); }} style={{ border: "none", background: "none", cursor: "pointer", padding: 6, color: RC.sub }}><Pencil size={15} /></button>
              <button title="Supprimer" onClick={() => remove(r)} style={{ border: "none", background: "none", cursor: "pointer", padding: 6, color: RC.danger }}><Trash2 size={15} /></button>
            </>
          )}
        </div>
      </div>
    </div>
  );

  return (
    <div className="somip-fade">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 16 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 17, fontWeight: 800, display: "flex", alignItems: "center", gap: 9 }}>
            <Video size={17} color={RC.blue} /> Réunions
          </h2>
          <p style={{ margin: "4px 0 0", fontSize: 12.5, color: RC.sub }}>Visioconférence intégrée (Jitsi Meet) — caméra, micro et partage d'écran directement dans la salle.</p>
        </div>
        {canManage && !showForm && (
          <button style={rBtnPrimary} onClick={() => { setEditing(null); setShowForm(true); }}>
            <Plus size={15} /> Nouvelle réunion
          </button>
        )}
      </div>

      {showForm && (
        <MeetingForm
          initial={editing}
          sites={sites}
          onCancel={() => { setShowForm(false); setEditing(null); }}
          onSubmit={createOrUpdate}
        />
      )}

      {loadError && <p style={{ color: RC.danger, fontSize: 13, marginBottom: 14 }}>{loadError}</p>}
      {loading && <p style={{ color: RC.sub, fontSize: 13 }}>Chargement des réunions…</p>}

      {!loading && !loadError && (
        <>
          <h4 style={{ fontSize: 12.5, fontWeight: 700, color: RC.sub, textTransform: "uppercase", letterSpacing: ".03em", margin: "0 0 10px" }}>
            À venir {upcoming.length > 0 && `(${upcoming.length})`}
          </h4>
          {upcoming.length === 0 && (
            <div style={{ ...rPanel, padding: 20, textAlign: "center", color: RC.sub, fontSize: 13, marginBottom: 18 }}>
              Aucune réunion prévue pour l'instant.{canManage ? " Crée-en une avec le bouton ci-dessus." : ""}
            </div>
          )}
          {upcoming.map((r) => <MeetingRow key={r.id} r={r} isPast={false} />)}

          {past.length > 0 && (
            <div style={{ marginTop: 22 }}>
              <button onClick={() => setShowPast((v) => !v)} style={{ border: "none", background: "none", cursor: "pointer", fontSize: 12.5, fontWeight: 700, color: RC.sub, padding: 0, marginBottom: 10 }}>
                {showPast ? "▾" : "▸"} Réunions passées ({past.length})
              </button>
              {showPast && past.map((r) => <MeetingRow key={r.id} r={r} isPast={true} />)}
            </div>
          )}
        </>
      )}
    </div>
  );
}

