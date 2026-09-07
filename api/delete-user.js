// Fonction serveur Vercel — jamais exécutée dans le navigateur.
// Utilise la clé "secret" Supabase (SUPABASE_SERVICE_ROLE_KEY), qui reste
// uniquement sur le serveur et n'est jamais envoyée au client.
import { createClient } from "@supabase/supabase-js";

const supabaseUrl = process.env.VITE_SUPABASE_URL;
const serviceRoleKey = process.env.SUPABASE_SERVICE_ROLE_KEY;

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

  const { data: userData, error: userErr } = await admin.auth.getUser(token);
  if (userErr || !userData?.user) {
    res.status(401).json({ error: "Session invalide, reconnecte-toi." });
    return;
  }
  const { data: requesterProfile, error: profErr } = await admin
    .from("profiles").select("role").eq("id", userData.user.id).maybeSingle();
  const detectedRole = requesterProfile?.role ? String(requesterProfile.role).trim().toLowerCase() : null;
  if (profErr) {
    res.status(500).json({ error: `Erreur de lecture du profil : ${profErr.message}` });
    return;
  }
  if (detectedRole !== "superviseur") {
    res.status(403).json({ error: "Seul un Superviseur peut supprimer un compte." });
    return;
  }

  const { userId } = req.body || {};
  if (!userId) {
    res.status(400).json({ error: "Identifiant de compte manquant." });
    return;
  }
  if (userId === userData.user.id) {
    res.status(400).json({ error: "Tu ne peux pas supprimer ton propre compte." });
    return;
  }

  const { error: deleteErr } = await admin.auth.admin.deleteUser(userId);
  if (deleteErr) {
    res.status(400).json({ error: deleteErr.message });
    return;
  }
  // Sécurité supplémentaire si la fiche profil n'a pas été retirée automatiquement.
  await admin.from("profiles").delete().eq("id", userId);

  res.status(200).json({ success: true });
}
