# Format d'entrée Markdown pour Doc2PPTX

Ce document décrit le format Markdown attendu par **Doc2PPTX** pour générer des présentations PowerPoint.

---

## 1. Structure générale

Le Markdown pour Doc2PPTX suit une structure hiérarchique qui permet de définir les sections et les slides de votre présentation :

1. **Métadonnées YAML** *(optionnel)* — Au début du fichier
2. **Titre de la présentation** — Titre principal H1 (`#`)
3. **Sections** — Titres H2 (`##`)
4. **Slides** — Titres H3 (`###`)
5. **Contenu** — Texte, listes, tableaux et images

### 1.1 Exemple de structure

```yaml
---
title: Présentation Exemple
author: Jean Dupont
description: Une présentation de démonstration
---
```

```markdown
# Présentation Exemple

## Introduction

### Principaux objectifs
* Premier objectif
* Deuxième objectif
* Troisième objectif

### Contexte
Texte explicatif sur le contexte...

## Analyse

### Données importantes
| Catégorie | Valeur | Tendance |
|-----------|--------|----------|
| A         | 42     | ↑        |
| B         | 18     | ↓        |
```

---

## 2. Métadonnées YAML

Le frontmatter YAML au début du document permet de définir des métadonnées :

```yaml
---
title: Titre de la présentation
author: Nom de l'auteur
description: Description brève
template_path: chemin/vers/template.pptx  # optionnel
---
```

**Clés prises en charge :**

* `title` : Titre de la présentation
* `author` : Auteur
* `description` : Description courte
* `template_path` : Chemin vers un template PowerPoint

---

## 3. Titres et hiérarchie

| Niveau          | Markdown | Usage                                     |
| --------------- | -------- | ----------------------------------------- |
| Slide titre     | `# H1`   | Titre de présentation (slide d'ouverture) |
| Section         | `## H2`  | Slide de section                          |
| Slide           | `### H3` | Titre d'une slide                         |
| Bloc de contenu | `####+`  | Titres de sous-sections ou blocs internes |

---

## 4. Types de contenu pris en charge

### 4.1 Texte

Le texte simple est converti en blocs de texte. Les paragraphes sont séparés par des lignes vides.

```markdown
Ceci est un paragraphe normal.

Ceci est un autre paragraphe.
```

### 4.2 Formatage de texte

**Markdown standard :**

* **Gras** : `**texte**`
* *Italique* : `*texte*`
* ~~Barré~~ : `~~texte~~`
* **Souligné** : `__texte__`

**Extensions Doc2PPTX :**

* `{color:red}Texte{/color}` pour la couleur
* `{highlight:yellow}Texte{/highlight}` pour le surlignage
* `{size:14pt}Texte{/size}` pour la taille

### 4.3 Listes

**Listes à puces :**

```markdown
* Élément 1
* Élément 2
  * Sous-élément A
  * Sous-élément B
```

**Listes numérotées :**

```markdown
1. Première étape
2. Deuxième étape
3. Troisième étape
```

### 4.4 Tableaux

Les tableaux Markdown sont convertis en tableaux PowerPoint :

```markdown
| Col 1      | Col 2      | Col 3      |
|------------|------------|------------|
| Valeur A1  | Valeur A2  | Valeur A3  |
| Valeur B1  | Valeur B2  | Valeur B3  |
```

**Styles disponibles :** `default`, `minimal`, `grid`, `accent1`, `accent2`, `accent3`.
Pour appliquer un style, ajoutez `| style:accent1 |` après les en-têtes.

### 4.5 Images

Syntaxe classique :

```markdown
![Texte alternatif](chemin/vers/image.jpg)
```

* Chemin local : `./images/graph.jpg`
* URL : `https://exemple.com/img.png`
* Requête de recherche : `query:graphique croissance`

---

## 5. Pratiques recommandées

* **Organisation logique** : Progression cohérente des slides.
* **Titres explicites** : Titres H3 clairs pour chaque slide.
* **Simplicité** : 3–5 points par liste.
* **Contexte** : Associer texte et tableaux.

---

## 6. Exemples d'utilisation

### 6.1 Slide texte + liste

```markdown
### Notre approche

* **Analyse rigoureuse** des données
* *Conception itérative*
* Mise en œuvre progressive
```

### 6.2 Slide tableau

```markdown
### Résultats financiers

| Indicateur            | T1 2024 | T2 2024 | Évolution |
|-----------------------|---------|---------|-----------|
| Chiffre d'affaires    | 2,3 M€  | 2,8 M€  | +21,7%    |
| Marge brute           | 1,1 M€  | 1,4 M€  | +27,3%    |
```

### 6.3 Slide image

```markdown
### Impact environnemental

![Réduction CO₂ sur 5 ans](emissions_co2.jpg)
```
