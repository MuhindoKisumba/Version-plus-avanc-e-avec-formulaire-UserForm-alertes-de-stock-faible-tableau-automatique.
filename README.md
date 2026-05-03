#  Gestion de Stock Excel (VBA)

##  Description
Ce projet est une application simple de **gestion de stock sous Excel en VBA**.  
Il permet de gérer des produits, suivre les quantités, calculer la valeur du stock et détecter les niveaux critiques.

---

##  Fonctionnalités

-  Initialisation automatique du tableau de stock
-  Ajout de produits via un formulaire
-  Modification des produits existants
-  Suppression de produits
-  Calcul automatique de la valeur du stock
-  Alerte visuelle et message pour les stocks faibles

---

##  Structure du tableau

| Colonne | Description |
|--------|------------|
| ID | Identifiant unique du produit |
| Produit | Nom du produit |
| Quantité | Quantité en stock |
| Prix | Prix unitaire |
| Seuil Alerte | Niveau minimum de stock |
| Valeur Stock | Quantité × Prix |

---

##  Fonctionnement

### 1. Initialisation
La macro `InitialiserTable` :
- Efface la feuille **Stock**
- Crée les en-têtes du tableau

### 2. Calcul du stock
La macro `CalculValeurStock` :
- Calcule automatiquement la valeur totale pour chaque produit

### 3. Alerte stock faible
La macro `AlerteStockFaible` :
- Compare la quantité avec le seuil
- Colore les lignes en rouge si stock faible
- Affiche une alerte avec les produits concernés

### 4. Formulaire utilisateur (UserForm)
Le formulaire `frmStock` permet :

####  Ajouter un produit
- Bouton : `btnAjouter`
- Ajoute une nouvelle ligne
- Met à jour les calculs et alertes

####  Modifier un produit
- Bouton : `btnModifier`
- Recherche par ID
- Met à jour les données

####  Supprimer un produit
- Bouton : `btnSupprimer`
- Supprime la ligne correspondante

---

##  Interface utilisateur

Le formulaire contient :
- `txtID` → ID du produit
- `txtNom` → Nom du produit
- `txtQte` → Quantité
- `txtPrix` → Prix
- `txtSeuil` → Seuil d’alerte

---

##  Utilisation

1. Ouvrir le fichier Excel
2. Lancer la macro :
   ```
   InitialiserTable
   ```
3. Ouvrir le formulaire :
   ```
   OuvrirFormulaire
   ```
4. Gérer les produits via l’interface

---

##  Prérequis

- Microsoft Excel avec macros activées
- Connaissances basiques en VBA (optionnel)

---

##  Améliorations possibles

-  Recherche avancée de produits
-  Tableau de bord avec graphiques
-  Sauvegarde automatique
-  Gestion des utilisateurs
-  Export en PDF ou CSV

---

##  Licence

Projet libre — tu peux le modifier et l’adapter selon tes besoins.
