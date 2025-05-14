# Formatage avancé dans Doc2PPTX

Ce document détaille les options de formatage et de mise en forme disponibles dans **Doc2PPTX** pour personnaliser l'apparence de vos présentations.

---

## 1. Formatage du texte

### 1.1 Formats Markdown standard

| Format   | Syntaxe Markdown           | Exemple                      | Résultat                  |
| -------- | -------------------------- | ---------------------------- | ------------------------- |
| Gras     | `**texte**` ou `__texte__` | **Texte important**          | Texte en gras             |
| Italique | `*texte*` ou `_texte_`     | *Information complémentaire* | Texte en italique         |
| Barré    | `~~texte~~`                | ~~Information obsolète~~     | Texte barré               |
| Combiné  | `**_texte_**`              | ***Très important***         | Texte en gras et italique |

### 1.2 Formats étendus spécifiques à Doc2PPTX

| Format        | Syntaxe                               | Exemple                                  | Description                |
| ------------- | ------------------------------------- | ---------------------------------------- | -------------------------- |
| Souligné      | `__texte__`                           | **Élément clé**                          | Texte souligné             |
| Couleur       | `{color:valeur}texte{/color}`         | {color\:red}Attention{/color}            | Change la couleur du texte |
| Surligné      | `{highlight:valeur}texte{/highlight}` | {highlight\:yellow}À retenir{/highlight} | Surligne le texte          |
| Taille police | `{size:valeur}texte{/size}`           | {size:20pt}Grand titre{/size}            | Change la taille du texte  |

### 1.3 Valeurs supportées

* **Couleurs** : noms (`red`, `blue`, `green`, etc.), hexadécimal (`#FF0000`), RGB (`rgb(255,0,0)`)
* **Tailles** : points (`14pt`), pixels (`18px`), numérique (`16` → points par défaut)

---

## 2. Listes et puces

### 2.1 Types de listes

* **Listes à puces** :

  ```markdown
  * Premier élément
  * Deuxième élément
    * Sous-élément A
    * Sous-élément B
  * Troisième élément
  ```

* **Listes numérotées** :

  ```markdown
  1. Première étape
  2. Deuxième étape
     1. Sous-étape A
     2. Sous-étape B
  3. Troisième étape
  ```

### 2.2 Mise en forme et détection

* Doc2PPTX détecte automatiquement le type de liste selon la syntaxe.
* Possibilité de combiner formatage et listes :

  ```markdown
  * **Action prioritaire** : Compléter l'analyse des besoins
  * *Action secondaire* : Consulter les parties prenantes
  * ~~Action abandonnée~~ : Ne pas poursuivre cette piste
  ```

---

## 3. Tableaux

### 3.1 Structure de base

```markdown
| En-tête 1        | En-tête 2        | En-tête 3        |
|------------------|------------------|------------------|
| Ligne 1, Col 1   | Ligne 1, Col 2   | Ligne 1, Col 3   |
| Ligne 2, Col 1   | Ligne 2, Col 2   | Ligne 2, Col 3   |
```

### 3.2 Styles prédéfinis

| Style   | Description                             | Usage                      |
| ------- | --------------------------------------- | -------------------------- |
| default | En-tête bleu, texte noir                | Tableaux formels           |
| minimal | Lignes fines                            | Données simples            |
| grid    | Grille complète, alternance de couleurs | Tableaux de données denses |
| accent1 | Couleur Accent1 du thème                | Tableaux importants        |
| accent2 | Couleur Accent2 du thème                | Tableaux secondaires       |
| accent3 | Couleur Accent3 du thème                | Variante de style          |

**Appliquer un style** : ajouter `| style:accent1 |` à la fin de la ligne d’en-têtes.

### 3.3 Alignement automatique

* Numéros et montants : alignement à droite
* Texte normal : alignement à gauche
* En-têtes : centrés

---

## 4. Images

### 4.1 Syntaxe de base

```markdown
![Texte alternatif](chemin/vers/image.jpg)
```

### 4.2 Sources d’images

1. Chemins locaux : `![Alt](./images/graph.png)`
2. URLs : `![Logo](https://exemple.com/logo.png)`
3. Requêtes API : `![Ventes](query:sales growth chart business)`

### 4.3 Attributs spéciaux

Annotations dans la balise alt :

```markdown
![Diagramme{width=50%,position=right}](diagramme.png)
```

---

## 5. Diagrammes & graphiques

### 5.1 Diagrammes Mermaid

````markdown
```mermaid
graph TD
  A[Début] --> B{Décision}
  B -->|Oui| C[Action 1]
  B -->|Non| D[Action 2]
  C --> E[Fin]
  D --> E
````

````

### 5.2 Graphiques de données

```markdown
```chart
type: bar
data: {
  labels: ["Jan", "Fév", "Mar"],
  datasets: [{ label: "Ventes 2025", data: [65, 59, 80] }]
}
````

````

**Types supportés** : `bar`, `line`, `pie`, `radar`, `polarArea`

---

## 6. Blocs de code

### 6.1 Syntaxe de base

```markdown
```python
print("Hello, world!")
````

````

### 6.2 Options avancées

```markdown
```python{line_numbers=true,highlight=[2,3]}
def calculate():
    x = 10
    y = 20
    return x + y
````

````
- `line_numbers` : `true`/`false`
- `highlight` : liste de numéros de lignes
- `theme` : `default`, `dark`, `light`

---

## 7. Mise en page spéciale

### 7.1 Colonnes

```markdown
::: columns
### Colonne gauche
Contenu...
:::

::: columns
### Colonne droite
Contenu...
:::
````

### 7.2 Boîtes d’information

```markdown
::: info
**Information importante**
Ceci est une zone mise en évidence.
:::

::: warning
**Attention**
Avertissement.
:::

::: success
**Succès**
Opération réussie !
:::
```

---

## 8. Notes du présentateur

```markdown
::: notes
Notes pour le présentateur :
- Mentionner le contexte
- Insister sur le point 2
- Prévoir 2 min pour cette slide
:::
```

---

## 9. Contrôles avancés

* **Forcer un layout** : `{layout="Titre et 3 colonnes"}`
* **ID de slide** : `{#slide-intro}` puis `[conclusion](#slide-conclusion)`
* **Séparateurs explicites** : `---`

---

## 10. Bonnes pratiques

* **Cohérence** : styles homogènes
* **Simplicité** : éviter les excès
* **Hiérarchie visuelle** : différencier niveaux
* **Espacement** : aérer les contenus
* **Contraste** : assurer lisibilité
* **Couleurs modérées** : 2–3 couleurs clés

---

## 11. Exemples de formatage avancé

### Exemple 1 : Texte formaté + liste

```markdown
### Stratégie de croissance 2025

Notre {size:16pt}**stratégie de croissance**{/size} repose sur trois axes :

* **Expansion marché** : {highlight:yellow}3 nouveaux marchés{/highlight}
* *Développement produit* : lancement __Solutions Pro__
* **Partenariats** : alliances avec {color:blue}leaders tech{/color}
```

### Exemple 2 : Tableau stylé

```markdown
### Comparaison des performances

| Indicateur       | **2023** | **2024** | **2025 (proj.)** | *Évolution* | style:accent2 |
|------------------|----------|----------|------------------|-------------|---------------|
| CA (M€)          | 12,4     | 15,7     | 19,8             | +59,7 %     |               |
| Marge brute (%)  | 38       | 42       | 45               | +7 pts      |               |
| Clients actifs   | 1 240    | 1 580    | 2 100            | +69,4 %     |               |
```

### Exemple 3 : Diagramme + explication

````markdown
### Processus d'innovation

```mermaid
graph LR
  A[Idéation] --> B[Conception]
  B --> C[Prototype]
  C --> D[Test]
  D --> A
````

Chaque itération améliore la solution selon les besoins du marché.

```

---

Ce document rassemble l’ensemble des options de formatage avancé dans Doc2PPTX pour des présentations riches et personnalisées.

```