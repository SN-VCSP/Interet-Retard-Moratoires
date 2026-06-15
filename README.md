# 💶 Calculateur d'Intérêts de Retard — Eurovia / VINCI Construction

![Version](https://img.shields.io/badge/version-2.0.0-blue)
![Python](https://img.shields.io/badge/python-3.10+-green)
![Streamlit](https://img.shields.io/badge/streamlit-1.28+-red)
![License](https://img.shields.io/badge/license-proprietary-gray)

Application Streamlit professionnelle pour le calcul des intérêts moratoires et pénalités de retard de paiement dans le secteur BTP, conformément au droit français.

---

## 📋 Vue d'ensemble fonctionnelle

Cette application permet de calculer les **intérêts de retard** dus par un client (privé ou public) en cas de paiement au-delà du délai contractuel ou légal. Elle implémente les dispositions du **Code de commerce** et du **Code de la commande publique**.

### Workflow utilisateur

1. L'utilisateur saisit les informations de la facture (montant TTC, dates)
2. Il choisit le type de client (Privé / Public) et le mode de taux (Légal / Manuel)
3. L'application récupère automatiquement les taux BCE en vigueur
4. Le calcul est effectué par segmentation temporelle (par semestre ou année)
5. L'utilisateur visualise le détail et peut exporter le résultat (HTML imprimable A4 / Excel)

---

## ⚖️ Bases légales

| Mode | Référence | Taux appliqué | Période d'actualisation | Début des intérêts |
|------|-----------|---------------|-------------------------|-------------------|
| **Client Privé** | Art. L.441-10 C.Com | BCE (MRO) + 10 pts | Semestrielle (1er janv. / 1er juil.) | J+1 après échéance |
| **Client Public** | Art. R.2192-31 CCP | BCE (MRO) + 8 pts | Annuelle (1er janvier) | J+1 après échéance |
| **Manuel** | Clause contractuelle | Taux fixe défini | Aucune | Jour d'échéance |

**Indemnité forfaitaire** : 40 € systématiquement ajoutée (Art. D.441-5 C.Com), applicable à tout professionnel en retard de paiement.

---

## 🧮 Formule de calcul

### Formule de base (par segment temporel)

$$
I_{segment} = M \times \frac{T}{100} \times \frac{J}{365}
$$

Où :
- $M$ = Montant TTC de la facture impayée (€)
- $T$ = Taux d'intérêt applicable (%) = Taux BCE + majoration
- $J$ = Nombre de jours du segment
- Base de calcul : **exact/365** (base calendaire réelle)

### Total

$$
\text{Total dû} = \left(\sum_{i=1}^{n} I_i\right) + 40\text{ €}
$$

### Arrondi

Arrondi bancaire (`ROUND_HALF_UP`) à 2 décimales appliqué sur chaque segment individuellement, puis sur le total.

---

## 🔄 Logique de segmentation

Le calcul est **segmenté par période** car le taux BCE peut évoluer au cours du retard de paiement :

### Client Privé (semestriel)

```
Échéance: 15/03/2024 | Paiement: 20/09/2024

Segment 1: 16/03/2024 → 30/06/2024 (107 jours) → Taux BCE au 01/01/2024 + 10 pts
Segment 2: 01/07/2024 → 19/09/2024 (81 jours)  → Taux BCE au 01/07/2024 + 10 pts
```

**Bornes semestrielles** : 1er janvier et 1er juillet de chaque année.

### Client Public (annuel)

```
Échéance: 15/11/2023 | Paiement: 20/03/2024

Segment 1: 16/11/2023 → 31/12/2023 (46 jours) → Taux BCE au 01/01/2023 + 8 pts
Segment 2: 01/01/2024 → 19/03/2024 (79 jours) → Taux BCE au 01/01/2024 + 8 pts
```

**Bornes annuelles** : 1er janvier de chaque année.

### Mode Manuel

Pas de segmentation — un seul taux fixe contractuel appliqué sur toute la période.

---

## 📡 Récupération des taux BCE (MRO)

L'application télécharge automatiquement les taux de la **facilité de refinancement principal (MRO)** de la BCE :

### Source primaire
- **API BCE** : `https://data-api.ecb.europa.eu/service/data/FM/D.U2.EUR.4F.KR.MRR_FR.LEV`
- Format CSV avec détection automatique du séparateur

### Source fallback
- **FRED (Federal Reserve)** : `https://fred.stlouisfed.org/graph/fredgraph.csv?id=ECBMRRFR`
- Activée si l'API BCE échoue

### Mise en cache
- Cache Streamlit avec TTL de **1 heure** (`@st.cache_data(ttl=3600)`)
- Bouton de rafraîchissement manuel en sidebar
- Le schedule conserve uniquement les dates de **changement de taux** (pas chaque jour)

### Détermination du taux à une date

```python
# Recherche du dernier taux applicable avant ou à la date donnée
for entry in schedule:
    if entry['start'] <= date_cible:
        applicable_rate = entry['rate']
```

---

## 🏗️ Architecture de l'application

```
app.py                  # Application monolithique Streamlit
├── Configuration       # CSS custom (design VINCI / Apple)
├── Classes             # ClientType, TauxMode (Enum), SegmentInteret, ResultatCalcul (dataclass)
├── Utilitaires         # Arrondi, formatage, dates
├── Taux BCE            # Téléchargement + parsing + cache
├── Calcul              # 3 fonctions de calcul (privé, public, manuel)
├── Exports             # HTML, CSV, Excel (openpyxl)
└── Interface           # Header, sidebar, formulaire, résultats
assets/
├── logo.png            # Logo principal (en-tête)
└── mon_logo.png        # Logo sidebar
requirements.txt        # Dépendances Python
```

---

## 🚀 Installation

### Prérequis

- Python 3.10+
- pip

### Installation locale

```bash
# Cloner ou télécharger le projet
cd InteretVCSP-main

# Créer un environnement virtuel (recommandé)
python -m venv venv
source venv/bin/activate      # Linux/Mac
venv\Scripts\activate         # Windows

# Installer les dépendances
pip install -r requirements.txt

# Lancer l'application
streamlit run app.py
```

### Dépendances

| Package | Usage |
|---------|-------|
| `streamlit>=1.28.0` | Framework web de l'interface |
| `pandas>=2.0.0` | Tableaux de données (affichage résultats) |
| `certifi>=2023.0.0` | Certificats SSL pour requêtes HTTPS |
| `openpyxl>=3.1.0` | Génération des exports Excel (.xlsx) |

---

## 🔒 Sécurité

- **SSL/TLS** : Vérification des certificats via `certifi` pour les appels API BCE/FRED
- **XSS** : Échappement HTML des saisies utilisateur dans les rapports exportés
- **Validation des entrées** : Contrôle des types et bornes via les widgets Streamlit natifs
- **Pas de persistance** : Aucune base de données, aucune donnée stockée côté serveur (session uniquement)
- **Pas d'authentification** : L'application ne gère pas de données personnelles sensibles

---

## 📊 Détail des modes de calcul

### 1. Client Privé — Art. L.441-10 C.Com

> *"Le taux des pénalités de retard est égal au taux d'intérêt appliqué par la Banque centrale européenne à son opération de refinancement la plus récente majoré de **10 points de pourcentage**."*

- Le taux est actualisé semestriellement : taux au 1er janvier pour S1, taux au 1er juillet pour S2
- Les intérêts courent à compter du jour suivant la date d'échéance
- Exigibles de plein droit, sans rappel nécessaire

### 2. Client Public — Art. R.2192-31 CCP

> *"Le taux des intérêts moratoires est égal au taux d'intérêt de la principale facilité de refinancement appliquée par la BCE [...] majoré de **8 points de pourcentage**."*

- Le taux est réactualisé au 1er janvier de chaque année
- Les intérêts courent à compter du jour suivant l'expiration du délai de paiement
- Délais légaux : 30 jours (sauf exception à 50 ou 60 jours)

### 3. Mode Manuel

- Taux contractuel fixe défini par l'utilisateur
- Applicable quand une clause pénale spécifique est prévue au contrat
- L'indemnité forfaitaire de 40 € reste due

---

## 📤 Exports disponibles

| Format | Contenu |
|--------|---------|
| **HTML (impression A4)** | Rapport compact optimisé pour impression A4, stylisé avec référence légale |
| **Excel** (.xlsx) | Tableau structuré avec styles, adapté aux services comptables |

---

## ⚠️ Limitations connues

1. **Base 365 fixe** : Pas de gestion des années bissextiles (366 jours). L'usage de 365 est la convention standard pour ce type de calcul.
2. **Calculateur d'échéance publique** : La date calculée n'est pas automatiquement reportée dans le formulaire (limitation Streamlit).
3. **Historique non persistant** : L'historique est perdu à la fermeture de la session Streamlit.
4. **Taux public annuel** : L'implémentation segmente par année civile. Pour un retard couvrant plusieurs semestres, le taux pourrait théoriquement différer selon l'interprétation stricte du texte R.2192-31.

---

## 🧪 Exemples de calcul

### Exemple 1 : Client Privé

- Montant : 50 000 €
- Échéance : 01/03/2025
- Paiement : 15/09/2025
- Taux BCE au 01/01/2025 : 3,15 %
- Taux BCE au 01/07/2025 : 2,65 % (hypothétique)

| Période | Jours | Taux | Intérêts |
|---------|-------|------|----------|
| 02/03 → 30/06/2025 | 121 j | 13,15 % | 2 181,44 € |
| 01/07 → 14/09/2025 | 76 j | 12,65 % | 1 316,99 € |

**Total intérêts** : 3 498,43 € + **40 €** = **3 538,43 €**

### Exemple 2 : Client Public

- Montant : 100 000 €
- Échéance : 15/11/2024
- Paiement : 01/04/2025
- Taux BCE au 01/01/2024 : 4,50 %
- Taux BCE au 01/01/2025 : 3,15 %

| Période | Jours | Taux | Intérêts |
|---------|-------|------|----------|
| 16/11 → 31/12/2024 | 46 j | 12,50 % | 1 575,34 € |
| 01/01 → 31/03/2025 | 90 j | 11,15 % | 2 749,32 € |

**Total intérêts** : 4 324,66 € + **40 €** = **4 364,66 €**

---

## 📄 Licence

Application propriétaire — © 2026 Eurovia / VINCI Construction.
Usage interne exclusif.
