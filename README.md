# SOMIP — Stock Gasoil

Application de gestion de stock de gasoil pour les sites externalisés SOMIP (Zone Sud-Est, Gabon).

## ⚠️ À lire avant de déployer — persistance des données

Ce code vient d'un artifact Claude, où les données étaient sauvegardées dans un
stockage clé-valeur **partagé et permanent** fourni par Anthropic
(`window.storage`). Cette fonctionnalité **n'existe pas** en dehors de Claude.

Pour que l'application reste fonctionnelle une fois hébergée, elle **bascule
automatiquement** sur le `localStorage` du navigateur :

- ✅ Les données survivent à la fermeture/réouverture du navigateur.
- ❌ Elles sont **propres à chaque navigateur/ordinateur** : deux personnes qui
  ouvrent le site depuis deux appareils différents auront chacune leurs
  propres données, qui ne se synchronisent pas.
- ❌ Un changement de navigateur, de mode navigation privée, ou un
  "vider les données de navigation" efface tout.

**C'est donc adapté pour un usage mono-poste / test**, mais **pas** pour
plusieurs opérateurs travaillant ensemble sur les mêmes données en production.

### Pour une vraie persistance partagée en production

Il faut ajouter un backend avec une vraie base de données (PostgreSQL, MongoDB…)
et remplacer les fonctions `loadRaw` / `saveRaw` dans `src/App.jsx` par des
appels à une API. C'est exactement l'objet du projet **Somip-Stock** déjà en
cours de construction séparément (FastAPI + MongoDB, déploiement sur VPS avec
le sous-domaine `stok.somip-sarl.ga`). Ce dossier-ci est une alternative
légère, pas un remplacement de ce projet.

## Installation locale

Prérequis : [Node.js](https://nodejs.org) (version 18 ou plus récente).

```bash
npm install
npm run dev
```

Ouvre ensuite l'adresse affichée dans le terminal (en général
`http://localhost:5173`).

## Déploiement en ligne (lien public)

### Option la plus simple : Vercel ou Netlify (glisser-déposer)

1. Sur ton poste, lance :
   ```bash
   npm install
   npm run build
   ```
   Cela crée un dossier `dist/` contenant le site prêt à héberger.
2. Va sur [vercel.com](https://vercel.com) ou [netlify.com](https://netlify.com),
   crée un compte gratuit, puis glisse-dépose le dossier `dist/` dans
   l'interface ("Deploy" / "Add new site → Deploy manually").
3. Tu obtiens immédiatement un lien public en `https://...`.

### Option recommandée à terme : connecter un dépôt Git

1. Crée un dépôt (GitHub, GitLab...) et pousses-y ce dossier.
2. Sur Vercel/Netlify, choisis "Importer depuis Git" et sélectionne le dépôt.
3. Chaque mise à jour poussée sur le dépôt redéploie automatiquement le site.

## Structure du projet

```
somip-stock-web/
├── index.html          Page HTML d'entrée
├── package.json        Dépendances (React, recharts, lucide-react, xlsx)
├── vite.config.js       Configuration du bundler
└── src/
    ├── main.jsx          Point d'entrée React
    └── App.jsx           L'application complète (tout le code fonctionnel)
```

## Fonctionnalités incluses

Tableau de bord, gestion des sites, réceptions, sorties (standard et vers
camions laitiers), inventaires avec calcul perte/gain et taux de freinte,
correction de volume à 15°C (ASTM D1250 / API MPMS 11.1), rapports
(journalier, décadaire, mensuel, état des stocks, exposition, pertes/gains
par site) avec export Excel et export PDF (impression), gestion des
utilisateurs et historique des modifications.
