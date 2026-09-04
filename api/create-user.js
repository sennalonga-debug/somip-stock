// Fonction serveur Vercel — jamais exécutée dans le navigateur.
// Utilise la clé "secret" Supabase (SUPABASE_SERVICE_ROLE_KEY), qui reste
// uniquement sur le serveur et n'est jamais envoyée au client.
import { createClient } from "@supabase/supabase-js";

const supabaseUrl = process.env.VITE_SUPABASE_URL;
const serviceRoleKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
const VALID_ROLES = ["superviseur", "operateur", "chauffeur", "lecture"];

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Méthode non autorisée." });
    return;
  }
  if (!supabaseUrl || !serviceRoleKey) {
    res.status(500).json({ error: "Configuration serveur manquante : SUPABASE_SERVICE_ROLE_KEY n'est pas défini dans les variables d'environnement Vercel." });
    return;
  }

  const authHeader = req.headers.authorization || "";
  const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;
  if (!token) {
    res.status(401).json({ error: "Non authentifié." });
    return;
  }

  const admin = createClient(supabaseUrl, serviceRoleKey, { auth: { autoRefreshToken: false, persistSession: false } });

  // Vérifie que la personne qui appelle est bien connectée et Superviseur.
  const { data: userData, error: userErr } = await admin.auth.getUser(token);
  if (userErr || !userData?.user) {
    res.status(401).json({ error: "Session invalide, reconnecte-toi." });
    return;
  }
  const { data: requesterProfile, error: profErr } = await admin
    .from("profiles").select("role").eq("id", userData.user.id).maybeSingle();
  if (profErr || !requesterProfile || requesterProfile.role !== "superviseur") {
    res.status(403).json({ error: "Seul un Superviseur peut créer un compte." });
    return;
  }

  const { email, password, fullName, role } = req.body || {};
  if (!email || !password || !fullName || !role) {
    res.status(400).json({ error: "Tous les champs sont requis." });
    return;
  }
  if (!VALID_ROLES.includes(role)) {
    res.status(400).json({ error: "Rôle invalide." });
    return;
  }
  if (String(password).length < 6) {
    res.status(400).json({ error: "Le mot de passe doit contenir au moins 6 caractères." });
    return;
  }

  const { data: created, error: createErr } = await admin.auth.admin.createUser({
    email, password, email_confirm: true, user_metadata: { full_name: fullName },
  });
  if (createErr) {
    res.status(400).json({ error: createErr.message });
    return;
  }

  // Le compte est créé avec le rôle par défaut "lecture" (déclencheur automatique) : on l'ajuste.
  const { error: updateErr } = await admin
    .from("profiles").update({ role, full_name: fullName }).eq("id", created.user.id);
  if (updateErr) {
    res.status(200).json({ warning: "Compte créé, mais le rôle n'a pas pu être appliqué automatiquement.", userId: created.user.id });
    return;
  }

  res.status(200).json({ success: true, userId: created.user.id });
}
