import pandas as pd

def lire_colonne_f(fichier_excel, fichier_sortie):
    try:
        # Lecture du fichier Excel
        donnees = pd.read_excel(fichier_excel, engine='openpyxl')
        
        # Vérification si la colonne F existe
        if 'F' in donnees.columns or donnees.shape[1] >= 6:
            # Récupération de la colonne F (en index 5 si c'est la 6e colonne)
            colonne_f = donnees.iloc[:, 5]
            
            # Suppression des doublons
            emails_uniques = colonne_f.drop_duplicates()
            
            # Écriture des résultats dans un fichier texte
            with open(fichier_sortie, 'w') as fichier:
                for email in emails_uniques:
                    fichier.write(email + '\n')
            
            print(f"Les emails uniques ont été écrits dans {fichier_sortie}.")
        else:
            print("La colonne F n'existe pas dans le fichier fourni.")
    except FileNotFoundError:
        print("Le fichier spécifié est introuvable.")
    except Exception as e:
        print(f"Une erreur s'est produite : {e}")

# Exemple d'utilisation
fichier_excel = ".xlsx"  # Remplacez par le chemin de votre fichier Excel
fichier_sortie = "emails_uniques.txt"  # Nom du fichier de sortie
lire_colonne_f(fichier_excel, fichier_sortie)