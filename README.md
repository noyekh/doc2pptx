# doc2pptx

**doc2pptx** est un outil Python qui génère et édite des présentations PowerPoint (.pptx) à partir de contenu structuré en Markdown.

## 🚀 Fonctionnalités

* **Génération de présentations** à partir de fichiers Markdown structurés
* **Templates PowerPoint** personnalisables avec sélection intelligente de layouts
* **Mapping automatique** du contenu aux layouts PowerPoint appropriés
* **Gestion du dépassement** de texte et de contenu
* **Interface en ligne de commande** simple et intuitive
* **Support pour le contenu riche** : texte, listes à puces, tableaux, images, diagrammes Mermaid

## 📋 Prérequis

* Python 3.12+
* [Conda](https://docs.conda.io/en/latest/) (optionnel, pour isoler l’environnement)

## ⚙️ Installation

1. Clonez le dépôt.
2. Créez et activez l’environnement Conda :

   ```bash
   conda env create -f env.yml
   conda activate doc2pptx
   ```
3. Installez le package en mode développement :

   ```bash
   pip install -e .
   ```

## 🔧 Utilisation

### 1. Préparez votre fichier Markdown

Votre document doit débuter par un frontmatter YAML (optionnel) puis suivre une structure hiérarchique :

```markdown
---
title: Stratégie Marketing Digital 2025
author: Marie Dupont
description: Présentation de la stratégie marketing digital pour l'année 2025
---

# Stratégie Marketing Digital 2025

## Introduction

Texte introductif...

### Principaux objectifs

* Objectif 1
* Objectif 2

## Analyse du marché actuel

| Tendance                  | Impact           | Position actuelle     |
|---------------------------|------------------|-----------------------|
| Intelligence artificielle | Très élevé       | En développement      |
| Marketing de contenu      | Élevé            | Bien établi           |
...

![Description de l'image](path/to/image.jpg)
```

### 2. Générez la présentation

```bash
doc2pptx generate \
  --template data/templates/my_template.pptx \
  --output   data/output/presentation.pptx \
  --ai-optimize        # (facultatif) pour IA
  --content-planning   # (facultatif) pour planification du contenu
  data/input/example.md
```

**Options CLI principales :**

| Option               | Alias | Description                                     |
| -------------------- | ----- | ----------------------------------------------- |
| `--template`         | `-t`  | Chemin vers le template PPTX                    |
| `--output`           | `-o`  | Chemin du fichier de sortie `.pptx`             |
| `--ai-optimize`      |       | Active l’optimisation IA de la mise en page     |
| `--content-planning` |       | Active la planification intelligente du contenu |
| `--verbose`          | `-v`  | Affiche les logs détaillés                      |

### 3. Exécution rapide via Python

```python
from doc2pptx.cli import main  # lève une exception si non configuré
import sys
sys.argv = [
    'doc2pptx', 'generate',
    '--template', 'data/templates/base_template.pptx',
    'data/input/example.md'
]
main()
```

## 📑 Format Markdown pris en charge

* **Frontmatter YAML** : `title`, `author`, `description`, `template_path`
* **Titres** : `# H1` → slide titre, `## H2` → section, `### H3` → slide
* **Texte** : paragraphes séparés par une ligne vide
* **Listes** : `*` ou `-` pour puces, `1.` pour numérotées
* **Tableaux** : syntaxe `| … | … | … |` (styles : `| … | style:accent1 |`)
* **Images** : `![texte alternatif](chemin/vers/image.jpg)`
* **Mermaid** : bloc ``mermaid` … ``
* **Graphiques** : bloc ``chart` … ``

## 🧪 Tests

```bash
pytest  # unitaires et E2E
pytest --cov=doc2pptx  # avec couverture
```

## 🛠️ Architecture du projet

```
src/doc2pptx/
├── core/              # Modèles Pydantic & réglages
├── ingest/            # Chargeur Markdown (et JSON déprécié)
├── layout/            # Règles de mapping vers layouts
├── ppt/               # Génération PPTX (builder, templates)
├── editor/            # Édition post-génération (JSON seed)
├── llm/               # Optimisation IA & planification
└── cli.py             # Interface Typer (generate / edit)
```

## 📄 Licence

Ce projet est sous licence [MIT](LICENSE).