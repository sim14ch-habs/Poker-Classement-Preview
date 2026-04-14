# Workflow Preview (avant publication publique)

Ce mode publie les changements sur un site de previsualisation separé.

## URL Preview cible

`https://sim14ch-habs.github.io/Poker-Classement-Preview/`

## Prerequis (une seule fois)

1. Creer le repo GitHub: `sim14ch-habs/Poker-Classement-Preview`
2. Dans le repo preview, activer GitHub Pages sur la branche `main`.
3. Verifier que le dossier local existe: `Poker-Classement-preview`.

## Utilisation

- Validation + OCR + publication preview:
  - `lancer_validation_photo_ui_preview.bat`

- Export seul + publication preview:
  - `lancer_export_preview.bat`

## Publication publique + miroir preview

- Continuer d'utiliser:
  - `lancer_validation_photo_ui.bat`

Depuis ce lanceur public, le site complet est publié en premier, puis le preview est synchronisé automatiquement avec les mêmes données Excel. Le miroir preview ne renvoie pas de courriel.

Le mode preview n'impacte pas le site public.
