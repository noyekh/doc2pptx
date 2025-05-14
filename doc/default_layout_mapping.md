# Mapping des types de contenu aux layouts PowerPoint

Ce document décrit comment les différents types de contenu sont associés aux layouts PowerPoint disponibles dans le template. Cette association est définie dans le fichier `rules.yaml` et implémentée par la classe `LayoutSelector`.

## Layouts disponibles

Les layouts suivants sont disponibles dans le template par défaut (`base_template.pptx`) :

| Nom du layout | Description | Utilisé pour |
|---------------|-------------|-------------|
| **Diapositive de titre** | Contient un titre principal, une zone de sous-titre, une ligne de séparation et un espace pour logo | Pages de titre, couverture de présentation |
| **Introduction** | Titre avec image de fond, zone de texte principale et numéro de page | Pages d'introduction, contexte général |
| **Chapitre 1** | Grand titre avec fond d'image et ligne de séparation | Titres de sections, transitions entre parties |
| **Titre et texte 1 visuel gauche** | Titre avec grande image à gauche et zone de texte à droite | Contenu avec images explicatives, photos |
| **Titre et texte 1 histogramme** | Titre avec texte à gauche et graphique à droite | Données chiffrées, statistiques, tendances |
| **Titre et 3 colonnes** | Titre avec trois zones de texte distinctes | Comparaisons, listes parallèles, triptyques |
| **Titre et texte** | Simple titre et grande zone de texte | Contenu textuel, explications, listes |
| **Titre et tableau** | Titre avec espace dédié pour un tableau | Données structurées, comparatifs, grilles |

## Règles de mapping par défaut

Le système applique les règles suivantes pour déterminer le layout le plus approprié :

### 1. Mapping par type de section

| Type de section | Layout par défaut |
|-----------------|-------------------|
| title           | Diapositive de titre |
| introduction    | Introduction |
| content         | Titre et texte |
| conclusion      | Chapitre 1 |
| appendix        | Titre et texte |
| custom          | Titre et texte |
| agenda          | Titre et texte |

### 2. Mapping par type de contenu

| Type de contenu | Layout par défaut |
|-----------------|-------------------|
| text            | Titre et texte |
| bullet_points   | Titre et texte |
| table           | Titre et tableau |
| image           | Titre et texte 1 visuel gauche |
| chart           | Titre et texte 1 histogramme |
| mermaid         | Titre et texte 1 histogramme |
| code            | Titre et texte |

### 3. Mapping par motifs de contenu

Certains motifs spécifiques dans le contenu peuvent déclencher l'utilisation d'un layout particulier :

| Motif dans le contenu | Layout associé |
|-----------------------|----------------|
| "thank you"           | Chapitre 1 |
| "agenda"              | Titre et texte |
| "columns"             | Titre et 3 colonnes |
| "comparison"          | Titre et 3 colonnes |
| "key points"          | Titre et texte |
| "market share"        | Titre et texte 1 histogramme |
| "graph" ou "chart"    | Titre et texte 1 histogramme |
| "image" ou "picture"  | Titre et texte 1 visuel gauche |
| "table"               | Titre et tableau |
| "questions"           | Chapitre 1 |

### 4. Mapping par combinaisons de contenu

Pour les slides contenant plusieurs types de contenu, des règles spéciales s'appliquent :

| Combinaison | Layout associé |
|-------------|----------------|
| image + texte | Titre et texte 1 visuel gauche |
| texte + texte | Titre et 3 colonnes |
| graphique + texte | Titre et texte 1 histogramme |
| 3+ blocs de points | Titre et 3 colonnes |

## Structure des placeholders par layout

Chaque layout contient différents placeholders qui peuvent recevoir du contenu. Voici la structure des placeholders pour chaque layout :

### Diapositive de titre
- Placeholder idx=0, type=TITLE : Titre principal
- Placeholder idx=1, type=SUBTITLE : Sous-titre ou date/auteur

### Introduction
- Placeholder idx=0, type=TITLE : Titre
- Placeholder idx=1, type=BODY : Contenu principal
- Placeholder idx=12, type=SLIDE_NUMBER : Numéro de page

### Chapitre 1
- Placeholder idx=0, type=TITLE : Titre de section

### Titre et texte 1 visuel gauche
- Placeholder idx=0, type=TITLE : Titre
- Placeholder idx=1, type=BODY : Contenu textuel (droite)
- Placeholder idx=2, type=PICTURE : Image (gauche)
- Placeholder idx=12, type=SLIDE_NUMBER : Numéro de page

### Titre et texte 1 histogramme
- Placeholder idx=0, type=TITLE : Titre
- Placeholder idx=1, type=BODY : Contenu textuel (gauche)
- Placeholder idx=2, type=CHART : Graphique (droite)
- Placeholder idx=12, type=SLIDE_NUMBER : Numéro de page

### Titre et 3 colonnes
- Placeholder idx=0, type=TITLE : Titre
- Placeholder idx=1, type=BODY : Colonne 1
- Placeholder idx=2, type=BODY : Colonne 2
- Placeholder idx=3, type=BODY : Colonne 3
- Placeholder idx=12, type=SLIDE_NUMBER : Numéro de page

### Titre et texte
- Placeholder idx=0, type=TITLE : Titre
- Placeholder idx=1, type=BODY : Contenu principal
- Placeholder idx=12, type=SLIDE_NUMBER : Numéro de page

### Titre et tableau
- Placeholder idx=0, type=TITLE : Titre
- Placeholder idx=12, type=SLIDE_NUMBER : Numéro de page
- (Le tableau est inséré en tant que forme TABLE)

## Logique de sélection des layouts

L'algorithme de sélection des layouts applique les règles suivantes, dans cet ordre :

1. Si une slide a déjà un `layout_name` défini et qu'il n'est pas "auto", ce layout est utilisé
2. Sinon, tente de trouver un layout basé sur le type de section (title, introduction, etc.)
3. Si aucun layout n'est trouvé et qu'il n'y a pas de slide, utilise le layout basé sur le type de section ou le layout par défaut
4. Si une slide est fournie, examine son contenu :
   - Détermine le type de contenu principal du premier bloc (text, image, etc.)
   - Vérifie si le contenu correspond à des motifs spécifiques (regex)
   - Vérifie les combinaisons de types de contenu
   - Si plus de 3 blocs de contenu, utilise le layout "multi_block_layout"
   - Si exactement 2 blocs de contenu, utilise le layout "two_block_layout"
5. En dernier recours, utilise le layout par défaut "Titre et texte"

## Configuration et personnalisation

### Extension des règles

Pour personnaliser ou étendre les règles de mapping, vous pouvez :

1. Modifier le fichier `rules.yaml` dans le dossier `layout/`
2. Ajouter de nouveaux motifs de détection dans la section `content_patterns`
3. Définir de nouvelles combinaisons dans la section `content_combinations`
4. Ajuster les mappings de types de section et de contenu

### Création de nouveaux layouts

Si vous souhaitez ajouter de nouveaux layouts :

1. Ouvrez le template PowerPoint en mode "Master Slide"
2. Créez et nommez vos nouveaux layouts
3. Mettez à jour `rules.yaml` pour inclure les nouveaux layouts
4. Utilisez la méthode `TemplateLoader.analyze_template()` pour vérifier que vos layouts sont correctement détectés

## Considérations techniques

- Les noms des layouts dans `rules.yaml` doivent correspondre **exactement** aux noms des layouts dans le template PowerPoint
- Lorsque vous modifiez les règles, exécutez les tests unitaires pour vérifier la compatibilité
- Les indices des placeholders peuvent varier selon les versions de PowerPoint, vérifiez-les avec `TemplateLoader.get_placeholder_mapping()`
- Pour les contenus complexes, vous pouvez définir manuellement le layout dans votre JSON d'entrée :
  ```json
  "slides": [
    {
      "layout_name": "Titre et 3 colonnes",
      ...
    }
  ]
  ```

## Bonnes pratiques

1. **Cohérence** : Utilisez des layouts similaires pour des contenus similaires
2. **Simplicité** : Évitez de surcharger les slides avec trop d'éléments
3. **Progression logique** : Utilisez les layouts pour guider le flux de votre présentation
4. **Validation** : Testez votre présentation après génération pour vérifier que les layouts sont adaptés
5. **Default fallback** : Assurez-vous que le layout par défaut "Titre et texte" fonctionne bien pour les cas non gérés

## Exemples

### Exemple 1 : Section de titre
```json
{
  "type": "title",
  "title": "Rapport annuel 2025",
  "subtitle": "Résultats et perspectives"
}
```
→ Utilise le layout "Diapositive de titre"

### Exemple 2 : Contenu avec image
```json
{
  "type": "image_left",
  "title": "Expansion internationale",
  "content": "Notre stratégie d'expansion sur les marchés émergents...",
  "image": {
    "source": "unsplash",
    "query": "global business expansion",
    "alt_text": "Business meeting international"
  }
}
```
→ Utilise le layout "Titre et texte 1 visuel gauche"

### Exemple 3 : Tableau de données
```json
{
  "type": "table",
  "title": "Résultats financiers par trimestre",
  "content": [
    ["Trimestre", "Chiffre d'affaires", "Croissance"],
    ["Q1 2025", "€2.3M", "+4.5%"],
    ["Q2 2025", "€2.7M", "+5.2%"],
    ["Q3 2025", "€3.1M", "+6.7%"]
  ]
}
```
→ Utilise le layout "Titre et tableau"

### Exemple 4 : Contenu en colonnes multiples
```json
{
  "type": "two_column",
  "title": "Comparaison des offres",
  "content": {
    "left": "**Plan Standard**\n- Fonctionnalité A\n- Fonctionnalité B\n- Support basique",
    "right": "**Plan Premium**\n- Toutes les fonctionnalités standard\n- Fonctionnalités avancées\n- Support 24/7"
  }
}
```
→ Utilise le layout "Titre et 3 colonnes"

### Exemple 5 : Diagramme ou graphique
```json
{
  "type": "chart",
  "title": "Répartition des ventes par région",
  "content": "```mermaid\npie title Ventes 2025\n    \"Europe\" : 42\n    \"Amériques\" : 31\n    \"Asie\" : 21\n    \"Autres\" : 6\n```"
}
```
→ Utilise le layout "Titre et texte 1 histogramme"

## Résolution des problèmes courants

### Layout non trouvé
Si vous rencontrez l'erreur "Layout not found", vérifiez :
- Que le nom du layout existe exactement dans le template
- Que le fichier rules.yaml référence correctement ce layout
- Que le template a bien été chargé avant la création de la présentation

### Problèmes de mise en page
Si le contenu ne s'affiche pas correctement :
- Vérifiez que le type de contenu correspond aux placeholders disponibles
- Assurez-vous que la quantité de texte n'est pas trop importante pour l'espace
- Pour les contenus spéciaux (tableaux, images), vérifiez que le layout approprié est utilisé

### Personnalisation incorrecte
Si vos règles personnalisées ne fonctionnent pas :
- Vérifiez la syntaxe YAML dans rules.yaml
- Assurez-vous que les expressions régulières sont valides
- Testez vos règles avec des cas simples avant d'ajouter de la complexité

## Méthodes d'extension avancées

### Classes de contenu
Vous pouvez définir des classes de contenu personnalisées pour regrouper des règles similaires :

```yaml
content_classes:
  financial_data:
    patterns: ["revenue", "profit", "sales", "growth"]
    layout: "Titre et texte 1 histogramme"
  
  product_features:
    patterns: ["features", "benefits", "specifications"]
    layout: "Titre et 3 colonnes"
```

### Pondération des règles
Vous pouvez attribuer des poids différents aux règles pour affiner la sélection :

```yaml
rules_weights:
  section_type: 10     # Priorité la plus haute
  content_pattern: 8
  content_type: 6
  block_count: 4       # Priorité la plus basse
```

### Layouts conditionnels
Vous pouvez définir des conditions complexes pour la sélection de layouts :

```yaml
conditional_layouts:
  - condition:
      section_index: 0
      content_contains: "welcome"
    layout: "Diapositive de titre"
  
  - condition:
      section_index: -1
      content_contains: "thank|questions"
    layout: "Chapitre 1"
```

## Conclusion

Le système de mapping de layouts de doc2pptx offre une grande flexibilité pour associer différents types de contenu aux layouts PowerPoint appropriés. En comprenant la logique de sélection et en personnalisant les règles selon vos besoins, vous pouvez créer des présentations visuellement cohérentes et professionnelles à partir de données structurées.

Pour plus d'informations, consultez :
- La documentation de la classe `LayoutSelector`
- Le fichier `rules.yaml` pour les configurations par défaut
- Les exemples dans le dossier `examples/`