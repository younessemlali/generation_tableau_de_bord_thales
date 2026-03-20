# 📊 Thales — Besoins & Candidatures Dashboard

Application Streamlit pour analyser les fichiers pivot Thales (Randstad Intérim et/ou Expectra).

## 🚀 Déploiement sur Streamlit Cloud

1. **Forker** ce repository sur GitHub
2. Aller sur [share.streamlit.io](https://share.streamlit.io)
3. Connecter votre compte GitHub
4. Sélectionner ce repo → branch `main` → fichier `app.py`
5. Cliquer **Deploy** — en ligne en ~2 minutes

## 📁 Fichiers acceptés

| Format | Extension | Notes |
|---|---|---|
| Excel moderne | `.xlsx` | Format recommandé |
| Excel ancien | `.xls` | Compatible |
| CSV | `.csv` | Séparateur `,` ou `;` détecté automatiquement, encodages UTF-8 / Latin-1 supportés |

| Fichier | Description |
|---|---|
| Randstad Intérim | Fichier pivot RI — 20 colonnes (sans Statut EdB) |
| Expectra | Fichier pivot EXP — 21 colonnes (avec Statut EdB) |

## 🎯 Fonctionnalités

- ✅ Upload de 1 ou 2 fichiers (RI seul, EXP seul, ou les deux)
- ✅ Vue par fournisseur : Randstad Intérim / Expectra / Consolidé RI+EXP
- ✅ Filtres dynamiques : site Thales, semaine de diffusion, statut EdB
- ✅ KPIs colorés par catégorie
- ✅ Graphiques interactifs : barres par semaine, camembert, top sites
- ✅ Comparaison RI vs EXP côte à côte
- ✅ Tableau des expressions critiques
- ✅ Export Excel téléchargeable

## 🏗️ Structure

```
thales_dashboard/
├── app.py              # Application principale
├── requirements.txt    # Dépendances Python
├── .streamlit/
│   └── config.toml     # Thème et configuration
└── README.md
```

## ⚙️ Lancer en local

```bash
pip install -r requirements.txt
streamlit run app.py
```
