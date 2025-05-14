# doc2pptx

**doc2pptx** est un outil Python permettant de générer et d'éditer des présentations PowerPoint (.pptx) à partir de contenu structuré en JSON ou Markdown.

## 🚀 Fonctionnalités

- **Génération de présentations** à partir de JSON/Markdown structuré
- **Templates PowerPoint** personnalisables avec sélection intelligente de layouts
- **Mapping automatique** du contenu aux layouts PowerPoint appropriés
- **Gestion du dépassement** de texte et de contenu
- **Interface en ligne de commande** simple et intuitive
- **Support pour le contenu riche** (texte, listes à puces, tableaux, images, diagrammes Mermaid)

## 📋 Prérequis

- Python 3.12+
- [Conda](https://docs.conda.io/en/latest/) pour la gestion de l'environnement

## ⚙️ Installation

1. Clonez le dépôt :

```bash
git clone https://github.com/username/doc2pptx.git
cd doc2pptx
```

2. Créez et activez l'environnement Conda :

```bash
conda env create -f env.yml
conda activate doc2pptx
```

3. Installez le package en mode développement :

```bash
pip install -e .
```

## 🔧 Utilisation

### Génération d'une présentation

```bash
doc2pptx generate --input input.json --template template.pptx --output output.pptx
```

### Options de génération

```
--input, -i          Fichier JSON d'entrée avec le contenu de la présentation
--template, -t       Fichier template PowerPoint (.pptx)
--output, -o         Chemin du fichier PowerPoint de sortie (.pptx)
--verbose, -v        Active les messages détaillés
```

### Format du fichier JSON d'entrée

```json
{
  "id": "ma-presentation",
  "title": "Titre de la présentation",
  "author": "Nom de l'auteur",
  "description": "Description de la présentation",
  "metadata": {
    "category": "Catégorie",
    "keywords": "mot-clé1, mot-clé2"
  },
  "sections": [
    {
      "id": "section-1",
      "title": "Titre de la section",
      "type": "title",
      "slides": [
        {
          "id": "slide-1",
          "title": "Titre de la diapositive",
          "layout_name": "Diapositive de titre",
          "blocks": [
            {
              "id": "block-1",
              "title": "Titre du bloc",
              "content": {
                "content_type": "text",
                "text": "Contenu texte du bloc"
              }
            }
          ],
          "notes": "Notes pour cette diapositive"
        }
      ]
    }
  ]
}
```

## 🧪 Tests

Exécutez les tests unitaires et end-to-end avec pytest :

```bash
pytest
```

Ou avec couverture de code :

```bash
pytest --cov=doc2pptx
```

## 🛠️ Architecture du projet

```
src/doc2pptx/
├── core/              # Modèles Pydantic de base
│   ├── models.py      # Section, SlideBlock, Presentation
│   └── settings.py    # BaseSettings (clés API, chemins)
├── ingest/            # Chargement des données
│   ├── json_loader.py # load_sections()
│   └── markdown_loader.py
├── layout/            # Mapping section → layout
│   ├── rules.yaml
│   └── selector.py
├── ppt/               # Génération PPTX
│   ├── template_loader.py
│   ├── builder.py     # PPTBuilder.build()
│   ├── overflow.py    # OverflowHandler
│   └── image.py       # Unsplash + mermaid
├── editor/            # Édition post-génération
│   ├── models.py      # MoveSlide, UpdateText…
│   ├── apply.py       # apply_commands()
│   └── utils.py
├── llm/               # Fonction NL → commandes
│   └── command_parser.py
└── cli.py             # Typer CLI (generate, edit, prompt)
```

## 📝 Exemple d'utilisation

1. Préparez un fichier JSON avec votre contenu structuré :

```bash
cat > example.json << EOF
{
  "id": "exemple-presentation",
  "title": "Présentation d'exemple",
  "author": "John Doe",
  "sections": [
    {
      "id": "intro",
      "title": "Introduction",
      "type": "title",
      "slides": [
        {
          "id": "slide-1",
          "title": "Titre de la présentation",
          "layout_name": "Diapositive de titre",
          "blocks": [
            {
              "id": "block-1",
              "content": {
                "content_type": "text",
                "text": "Présentation générée avec doc2pptx"
              }
            }
          ]
        }
      ]
    }
  ]
}
EOF
```

2. Générez la présentation :

```bash
doc2pptx generate --input example.json --template template.pptx --output presentation.pptx
```

3. Ouvrez le fichier PowerPoint généré :

```bash
# Sous Windows
start presentation.pptx

# Sous macOS
open presentation.pptx

# Sous Linux
xdg-open presentation.pptx
```

## 📈 Roadmap

- [x] Génération de base à partir de JSON
- [x] Sélection intelligente de layouts
- [x] Gestion du dépassement de texte
- [ ] Support complet des images (Unsplash + locales)
- [ ] Support des diagrammes Mermaid
- [ ] Édition de présentations existantes
- [ ] Commandes en langage naturel

## 🤝 Contribution

Les contributions sont les bienvenues ! N'hésitez pas à ouvrir une issue ou à soumettre une pull request.

## 📄 Licence

Ce projet est sous licence [MIT](LICENSE).