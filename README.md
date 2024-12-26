# recup_mail_excel

Ce programme lit l'export de vos commandes HelloAsso (en fichier Excel) , extrait les données de la colonne spécifique des adresses mail (généralement F), supprime les doublons, et écrit les résultats dans un fichier texte.

## Prérequis

- Python 3.11 ou supérieur
- Les bibliothèques Python suivantes :
  - `pandas`
  - `openpyxl`

Vous pouvez installer les dépendances avec la commande suivante :

```bash
pip install pandas openpyxl
```

## Instructions d'Utilisation
1. Placez votre fichier Excel dans le même répertoire que le script.
2. Modifiez les variables fichier_excel et fichier_sortie dans le script :
- fichier_excel : Chemin vers votre fichier Excel.
- fichier_sortie : Nom ou chemin du fichier texte de sortie.
3. Exécutez le script avec la commande suivante :
```bash
python3 recup_mail_excel.py
```
4. Le fichier de sortie contiendra toutes les adresses e-mail uniques.

## Dépannage

- Erreur : "Missing optional dependency 'openpyxl'" :
Installez openpyxl avec la commande :
```bash
pip install openpyxl
```
- Erreur : "Fichier introuvable" :
Assurez-vous que le chemin du fichier Excel est correct.

- La colonne F n'existe pas :
Vérifiez que la colonne F existe dans votre fichier Excel et contient les données à extraire.
