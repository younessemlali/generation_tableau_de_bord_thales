# 📊 Thales — Besoins & Candidatures Dashboard

Application Streamlit pour analyser les fichiers pivot Thales (Randstad Intérim et/ou Expectra).

## 🚀 Déploiement sur Streamlit Cloud

1. **Forker** ce repository sur GitHub
2. Aller sur [share.streamlit.io](https://share.streamlit.io)
3. Connecter votre compte GitHub
4. Sélectionner ce repository, branch `main`, fichier `app.py`
5. Cliquer **Deploy**

## 📁 Fichiers attendus

| Fichier | Description | Colonnes clés |
|---|---|---|
| Randstad Intérim | Fichier pivot RI | 20 colonnes (sans Statut EdB) |
| Expectra | Fichier pivot EXP | 21 colonnes (avec Statut EdB) |

## 🎯 Fonctionnalités

- Upload d'un ou deux fichiers Excel
- Vue par fournisseur (RI / EXP) ou consolidée
- Filtres dynamiques : site, semaine, statut EdB
- KPIs interactifs
- Graphiques : barres par semaine, camembert, top sites
- Comparaison RI vs EXP côte à côte
- Tableau des expressions critiques
- Export Excel téléchargeable

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
