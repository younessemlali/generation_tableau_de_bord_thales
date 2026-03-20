# 📊 Thales — Besoins & Candidatures Dashboard

Application Streamlit d'analyse des fichiers pivot PIXID (Randstad Intérim et/ou Expectra).

---

## 🚀 Déploiement sur Streamlit Cloud

1. **Forker** ce repository sur GitHub
2. Aller sur [share.streamlit.io](https://share.streamlit.io)
3. Connecter votre compte GitHub
4. Sélectionner ce repo → branch `main` → fichier `app.py`
5. Cliquer **Deploy**

---

## 📂 Structure du repo

```
votre-repo/
├── app.py                  ← Application principale
├── requirements.txt        ← Dépendances Python
├── README.md
├── .streamlit/
│   ├── config.toml         ← Thème et configuration
│   └── secrets.toml        ← (à créer localement, NE PAS committer)
└── data/                   ← Dépôt des fichiers PIXID (mode GitHub)
    ├── randstad_interims.xlsx
    └── expectra.xlsx
```

---

## 🔗 Mode GitHub (recommandé pour usage régulier)

Déposez vos fichiers dans le dossier `/data/` du repo.
L'application les charge automatiquement — aucun upload nécessaire.

### Configuration requise

Sur **Streamlit Cloud** → votre app → **Settings → Secrets**, ajoutez :

```toml
GITHUB_RAW_URL = "https://raw.githubusercontent.com/VOTRE_USER/VOTRE_REPO/main"
```

Puis déposez vos fichiers dans `/data/` :
- `data/randstad_interims.xlsx` → données Randstad Intérim
- `data/expectra.xlsx` → données Expectra

Le dashboard se met à jour automatiquement à chaque nouveau dépôt.

### Pour un repo privé

Ajoutez également votre token GitHub dans les secrets :
```toml
GITHUB_RAW_URL = "https://raw.githubusercontent.com/VOTRE_USER/VOTRE_REPO/main"
GITHUB_TOKEN = "ghp_votre_token_ici"
```

---

## ⬆️ Mode Upload (usage ponctuel)

Sélectionnez "Upload manuel" dans la sidebar et déposez vos fichiers directement.

**Formats acceptés** : `.xlsx` `.xls` `.csv`

---

## 📁 Fichiers PIXID attendus

| Fichier | Fournisseur | Colonnes clés |
|---|---|---|
| Randstad Intérim | RI | 20 colonnes (sans Statut EdB) |
| Expectra | EXP | 21 colonnes (avec Statut EdB) |

---

## 🎯 Fonctionnalités

- **2 modes** : Upload manuel ou chargement depuis GitHub
- **3 vues** : Randstad Intérim / Expectra / Consolidé RI+EXP
- **4 onglets** : Tableau de Bord, Recherche, Analyses & Statistiques, Actions Requises
- **Recherche** : site, SIRET (liste + saisie libre), qualification, division, semaine, situation
- **6 analyses** : taux couverture, délais, tendance, qualifications en tension, RI vs EXP, stats par site
- **Export Excel** téléchargeable

---

## ⚙️ Lancer en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Pour le mode GitHub en local, créez `.streamlit/secrets.toml` :
```toml
GITHUB_RAW_URL = "https://raw.githubusercontent.com/VOTRE_USER/VOTRE_REPO/main"
```
