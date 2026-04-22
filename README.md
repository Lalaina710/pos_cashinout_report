# Rapport Cash In/Out PdV (Odoo 18)

Module Odoo 18 — Export Excel des mouvements caisse (Cash In / Cash Out) par Point de Vente.

Développé pour **SOPROMER**.

## Fonctionnalités

- Export Excel à 2 onglets : Détail des mouvements + Synthèse par PdV
- Filtres : période, point(s) de vente, type de mouvement (IN / OUT / tous)
- Catégorisation auto : Cash In / Cash Out / Écart clôture / Règlement session / Autre
- Options :
  - Exclure les encaissements session (`POS/XXXXX` sans `-in-`/`-out-`)
  - Exclure les écarts de comptage clôture
  - Synthèse par PdV (totaux IN / OUT / solde net)

## Source des données

Table `account.bank.statement.line` filtrée sur `pos_session_id IS NOT NULL`.

Les références sont analysées selon le pattern auto-généré par Odoo POS :
```
<session_name>-<in|out>-<motif utilisateur>
```

## Installation

1. Placer le module dans un `addons_path` Odoo :
   ```
   /opt/odoo18/custom_addons/dev/pos_cashinout_report/
   ```
2. Mettre à jour la liste des modules (Apps → Update Apps List)
3. Installer : **Rapport Cash In/Out PdV**

## Utilisation

Menu :
```
Point de Vente → Analyse → Rapport Cash In/Out
```

1. Choisir la période
2. Optionnel : filtrer par PdV, type, inclure/exclure encaissements et écarts
3. Cliquer **Exporter Excel** → téléchargement automatique

## Dépendances

- `point_of_sale`
- `account`
- Python : `xlsxwriter`

## Structure

```
pos_cashinout_report/
├── __init__.py
├── __manifest__.py
├── README.md
├── security/
│   └── ir.model.access.csv
└── wizard/
    ├── __init__.py
    ├── pos_cashinout_report_wizard.py
    └── pos_cashinout_report_wizard_views.xml
```

## Catégories détectées

| Catégorie | Critère ref | Type |
|-----------|-------------|------|
| Écart clôture | contient "écart" / "ecart" | IN ou OUT selon signe |
| Règlement session | commence par `pos/` sans `-in-`/`-out-` | IN ou OUT selon signe |
| Cash Out | contient `-out-` | OUT |
| Cash In | contient `-in-` | IN |
| Autre | aucun des ci-dessus | IN ou OUT selon signe |

## Licence

LGPL-3

## Auteur

SOPROMER (Madagascar) — Rakotoarisoa Hervé
