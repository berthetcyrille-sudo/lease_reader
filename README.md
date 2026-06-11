# Lease Reader

Outil d'extraction automatique de données de baux commerciaux via Claude (Anthropic).  
Stack : React + Vite · Supabase (Edge Function + PostgreSQL) · Vercel

---

## Fonctionnalités

- Drop zone PDF / DOCX
- Extraction de 33 champs structurés par Claude Opus
- Historique des extractions en base Supabase
- Export Excel double format (tableau 1 ligne + fiche clé/valeur)

---

## Installation

### 1. Cloner et installer

```bash
git clone https://github.com/VOTRE_COMPTE/lease-reader.git
cd lease-reader
npm install
```

### 2. Supabase — base de données

Dans **Supabase > SQL Editor**, exécuter le fichier :

```
supabase/migrations/001_extractions.sql
```

Cela crée la table `extractions` avec RLS activé.

### 3. Supabase — Edge Function

```bash
# Installer Supabase CLI si besoin
npm install -g supabase

# Se connecter
supabase login

# Lier au projet
supabase link --project-ref VOTRE_PROJECT_ID

# Déployer la fonction
supabase functions deploy extract-lease

# Ajouter la clé Anthropic comme secret
supabase secrets set ANTHROPIC_API_KEY=sk-ant-XXXXXXXX
```

### 4. Variables d'environnement (local)

```bash
cp .env.example .env.local
```

Renseigner dans `.env.local` :

```
VITE_SUPABASE_URL=https://VOTRE_PROJECT_ID.supabase.co
VITE_SUPABASE_ANON_KEY=VOTRE_ANON_KEY
```

Les clés sont disponibles dans Supabase > Settings > API.

### 5. Lancer en local

```bash
npm run dev
```

---

## Déploiement Vercel

1. Pousser le repo sur GitHub
2. Importer dans Vercel
3. Ajouter les variables d'environnement :
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
4. Deploy

---

## Structure du projet

```
lease-reader/
├── src/
│   ├── App.jsx          # Composant principal
│   ├── App.css          # Styles
│   ├── fields.js        # Champs et prompt d'extraction
│   ├── export.js        # Export Excel (xlsx)
│   ├── supabase.js      # Client Supabase
│   ├── main.jsx         # Entry point
│   └── index.css        # Reset / variables CSS
├── supabase/
│   ├── functions/
│   │   └── extract-lease/
│   │       └── index.ts # Edge Function (appel Claude API)
│   └── migrations/
│       └── 001_extractions.sql
├── .env.example
├── vite.config.js
└── package.json
```

---

## Ajouter / modifier des champs

Tout est centralisé dans `src/fields.js` :
- `SECTIONS` : organisation des champs par section
- `EXTRACTION_PROMPT` : prompt envoyé à Claude

Ajouter un champ = ajouter une entrée dans la section concernée et la clé correspondante dans le prompt.

---

## Évolutions prévues

- [ ] Auth Supabase (multi-utilisateurs)
- [ ] Comparaison de plusieurs baux
- [ ] Alertes sur dates clés (break options, congés)
- [ ] Import batch (plusieurs fichiers)
