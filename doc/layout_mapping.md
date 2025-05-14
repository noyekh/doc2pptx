# Mapping des Layouts dans Doc2PPTX

Ce document explique comment **Doc2PPTX** sélectionne et utilise les layouts (dispositions) PowerPoint pour différents types de contenu.

---

## 1. Principes du mapping de layouts

Doc2PPTX associe automatiquement le contenu Markdown aux layouts appropriés dans votre template PowerPoint :

1. **Analyse** du contenu (texte, listes, tableaux, images)
2. **Détermination** du type de layout optimal
3. **Association** au layout correspondant dans votre template

---

## 2. Types de layouts standards

| Nom de layout              | Utilisation                    | Description                                             |
| -------------------------- | ------------------------------ | ------------------------------------------------------- |
| Diapositive de titre       | Slide de titre                 | Layout avec titre principal et sous-titre               |
| Slides d'introduction      | Slides d’introduction          | Layout avec titre et zone de contenu large              |
| Slides de contenu standard | Titre et texte                 | Layout avec titre et zone de texte                      |
| Slides avec tableaux       | Titre et tableau               | Layout avec titre et zone de texte + tableau            |
| Slides avec images         | Titre et texte 1 visuel gauche | Layout avec titre, image à gauche et texte à droite     |
| Slides avec graphiques     | Titre et texte 1 histogramme   | Layout avec titre, texte à gauche et graphique à droite |
| Slides multi-colonnes      | Titre et 3 colonnes            | Layout avec titre et trois zones de contenu             |
| Chapitre 1                 | En-têtes de section            | Layout avec titre de section uniquement                 |

---

## 3. Mapping des layouts par type de contenu

### 3.1 Slides de titre

* **Contenu** : Titre H1 (`#`) et éventuellement un bloc de texte
* **Layout utilisé** : `Diapositive de titre`
* **Placeholders** : Titre, Sous-titre (optionnel)

### 3.2 Slides de section

* **Contenu** : Titre H2 (`##`)
* **Layout utilisé** : `Chapitre 1` ou équivalent
* **Placeholders** : Titre

### 3.3 Slides de texte standard

* **Contenu** : Titre H3 (`###`) suivi de texte et/ou listes
* **Layout utilisé** : `Titre et texte`
* **Placeholders** : Titre, Corps de texte

### 3.4 Slides avec tableau

* **Contenu** : Titre H3 (`###`), éventuellement texte + tableau
* **Layout utilisé** : `Titre et tableau`
* **Placeholders** : Titre, Zone de texte (si disponible)
* **Remarque** : Si aucun texte n’accompagne le tableau et que l’option AI est activée, une description sera générée.

### 3.5 Slides avec image

* **Contenu** : Titre H3 (`###`), texte & référence d’image
* **Layout utilisé** : `Titre et texte 1 visuel gauche`
* **Placeholders** : Titre, Image, Corps de texte

### 3.6 Slides avec graphique

* **Contenu** : Titre H3 (`###`), texte + description de graphique
* **Layout utilisé** : `Titre et texte 1 histogramme`
* **Placeholders** : Titre, Corps de texte, Graphique

### 3.7 Slides multi-colonnes

* **Contenu** : Titre H3 (`###`) et plusieurs blocs de contenu
* **Layout utilisé** : `Titre et 3 colonnes`
* **Placeholders** : Titre, Colonnes de contenu

---

## 4. Ordre de recherche des layouts

1. Correspondance exacte par nom
2. Correspondance insensible à la casse
3. Correspondance basée sur des motifs de mots-clés
4. Sélection en fonction du type de contenu et capacités du layout

---

## 5. Mapping personnalisé des layouts

En cas de template PowerPoint personnalisé, Doc2PPTX mappe les layouts génériques aux layouts spécifiques du template par :

* Noms exacts
* Noms similaires
* Motifs de mots-clés
* Capacités détectées (titre, contenu, table, etc.)

### 5.1 Motifs de correspondance

| Layout générique      | Motifs de recherche                              |
| --------------------- | ------------------------------------------------ |
| Title Slide           | "titre", "title", "diapositive de titre"         |
| Title and Content     | "titre et texte", "content", "text"              |
| Title and Two Content | "titre et 3 colonnes", "columns", "two content"  |
| Title and Table       | "titre et tableau", "table"                      |
| Title Only            | "titre seul", "chapitre", "section"              |
| Title and Image       | "titre et texte 1 visuel", "image"               |
| Title and Chart       | "titre et texte 1 histogramme", "chart", "graph" |

### 5.2 Capacités des layouts

| Capacité           | Description                                       |
| ------------------ | ------------------------------------------------- |
| Supports Title     | Placeholder de titre présent                      |
| Supports Content   | Placeholder de contenu/corps de texte             |
| Supports Table     | Conçu pour accueillir des tableaux                |
| Supports Image     | Placeholder d’image présent                       |
| Supports Chart     | Placeholder de graphique présent                  |
| Max Content Blocks | Nombre maximal de blocs de contenu pris en charge |

---

## 6. Optimisation par IA

Lorsque l’option `use_ai=True` est activée, Doc2PPTX peut :

* Analyser avec précision les layouts du template
* Déterminer les cas d’usage optimaux pour chaque layout
* Suggérer des améliorations de distribution du contenu
* Générer des descriptions pour les tableaux si nécessaire

---

## 7. Planification intelligente du contenu

Doc2PPTX inclut un système de planification qui :

* Analyse votre contenu Markdown
* Identifie les relations thématiques
* Regroupe le contenu connexe sur les mêmes slides
* Équilibre la quantité de contenu par slide
* Conserve ensemble le texte introductif et ses listes associées

**Caractéristiques clés :**

* Association automatique texte-tableau
* Évitement de slides isolées sans contexte
* Maintien des groupes logiques
* Répartition équilibrée du contenu

---

## 8. Exemples de mapping

### Exemple 1 : Contenu texte standard

```markdown
### Objectifs du projet

Les objectifs principaux sont :

* Améliorer la satisfaction client
* Réduire les coûts opérationnels
* Augmenter l’efficacité des processus
```

**Layout sélectionné** : "Titre et texte"

### Exemple 2 : Tableau avec texte descriptif

```markdown
### Performance financière

Notre performance financière a dépassé les attentes.

| Trimestre | Revenus | Croissance |
|-----------|---------|------------|
| Q1 2025   | 2,4 M€  | +12 %      |
| Q2 2025   | 2,8 M€  | +17 %      |
```

**Layout sélectionné** : "Titre et tableau"

### Exemple 3 : Image avec description

```markdown
### Nouveau modèle de produit

Notre nouveau modèle présente un design innovant.

![Photo du nouveau modèle de produit](nouveau_produit.jpg)
```

**Layout sélectionné** : "Titre et texte 1 visuel gauche"

---

## 9. Dépannage des problèmes de layout

* **Vérifiez les noms de layouts** : Assurez-vous que votre template contient des noms reconnaissables.
* **Inspectez la structure du contenu** : Suivez la hiérarchie Markdown attendue.
* **Titres explicites** : Chaque slide devrait avoir un titre H3 clair.
* **Analyse du template** : Utilisez `template_loader.analyze_template()` pour vérifier les capacités.
* **Activer l’IA** : Passez `use_ai=True` pour une détection améliorée.