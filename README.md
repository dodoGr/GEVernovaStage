# <u>Exercice Stage GEVernova</u>

## énoncé de l'exercice

 Il fallait réaliser un programme permettant de lire la page 'Summary' d'un fichier Excel et ajouter des pages.

Il fallait récupérer les valeurs présentes dans ce tableau : 
![Tableau de valeurs](TableauReference.png)

Les nouvelles pages doivent avoir pour nom les valeurs lues dans la colonne "SHEET". 
Ensuite, chaque page doit contenir un tableau de cette forme : 

![Exemple de tableau](exempleTableauExo.png)

## Explication de mes démarches 

J'ai décidé de coder ce programme en python. Pour cela, j'ai utilisé la bibliothèque 'pandas' et 'openpyxl'. 
Pour m'aider dans mes recherches, je me suis aidé de la documentations présente sur internet ainsi que de quelques video pour avoir les bases du lien entre python et excel.

[Doc Pandas](https://pandas.pydata.org/docs/user_guide/index.html) -- [Doc openpyxl](https://openpyxl.readthedocs.io/en/stable/) -- [Doc XlsxWritter](https://xlsxwriter.readthedocs.io/working_with_pandas.html) -- [Video Bases](https://www.youtube.com/watch?v=njpJoWE3WdI)

---
#### Récupération des données et création des feuilles
Pour commencé, j'ai commencé à lire le fichier et afficher ce qu'il contenait dans le terminal pour pourvoir savoir comment apparaissaient les données que j'allais devoir traiter.
J'aurais pu faire une simple boucle qui crée une page et un tableau avec des noms prédéfinis mais j'ai choisi d'automatiser cela et lire chaque case et stocker les valeurs pour les réutiliser en les affichant sur la bonne page et dans le bon tableau.
J'ai donc créé une fonction "table" qui prend pour paramètres la longeur de la colonne (le nombre de valeurs présentes sous le mot "SHEET"), les valeurs dans la colonne "SHEET", les valeurs dans la colonne "TITLE" et les valeurs dans la colonne "TESTS".

Voici le code de cette fonction :
```python
def table(valIndex, name_page, name_title, name_test):
    with pd.ExcelWriter('Exercice.xlsx', engine='openpyxl', mode='a') as writer:
        for i in range(0, valIndex):  # Boucle pour créer autant de page que de colonnes sous 'SHEET'
            new_sheet_name = f'{name_page[i]}'

            modele = pd.DataFrame({
                'Colonne 1': [name_title[i], name_test[i], np.nan],
                'Colonne 2': [np.nan, name_test[i] + ".1", name_test[i] + ".2"]
    
            #    'Colonne 1': ["Case " + str(i+1), "Test " + str(i+1), np.nan],
            #    'Colonne 2': [np.nan, "Test " + str(i+1) + ".1", "Test " + str(i+1) + ".2"]
                
            })

            if new_sheet_name not in writer.book.sheetnames:
                # Ajouter une nouvelle feuille
                modele.to_excel(writer, sheet_name=new_sheet_name, index=False)
                print(f"Nouvelle feuille '{new_sheet_name}' ajoutée avec succès.")
            else:
                print(f"La feuille '{new_sheet_name}' existe déjà.")

    # Charger le fichier Excel avec openpyxl
    wb = load_workbook(file_path)
    for i in range(0, valIndex):  # appliquer le formatage à toutes les nouvelles pages
        new_sheet_name = f'{name_page[i]}'
        if new_sheet_name in wb.sheetnames:
            nouvellePage = wb[new_sheet_name]  # accéder à la page
            formatage(nouvellePage, min_row=1, max_row=4, min_col=1, max_col=2)

    # Sauvegarder les modifications
    wb.save(file_path)
```

Pour récupérer la position des colonnes à lire, il fallait que je trouve le mot "SHEET" sur la feuille et que je retourne sa position.
Ensuite, j'ai fait une nouvelle fonction "lire_colonne" qui permet de lire autant de valeur possible, tant qu'on ne trouve pas de cellule vide et qui stocke chacune d'entre elle.
Je l'ai utilisé pour récupérer les valeurs de "SHEET", "TITLE" et "TESTS". 
```python
def lire_colonne(feuille, colonne): #fonction qui permet de lire les valeurs d'un colonne tant quon ne trouve pas de cellule vide

    ligne = 3  
    valeurs = []  # liste pour stocker les valeurs lues

    while True:
        # Lire la valeur de la cellule
        valeur = feuille.cell(row=ligne, column=colonne).value

        # Arrêter la boucle si la cellule est vide
        if valeur is None:
            break

        # Ajouter la valeur à la liste
        valeurs.append(valeur)

        # Passer à la ligne suivante
        ligne += 1

    return valeurs
```
---
#### Formatage de la feuille

Une fois que les pages étaient créées, il fallait formater le tableau. 
J'ai prédéfini un dataFrame : 
```
modele = pd.DataFrame({
    'Colonne 1': [name_title[i], name_test[i], np.nan],
    'Colonne 2': [np.nan, name_test[i] + ".1", name_test[i] + ".2"]
})
```
Mais il fallait aussi ajouter des bordures au tableau et assembler des cellules ensembles. 
J'ai donc créé une nouvelle fonction "formatage" qui me permet de faire en sorte que le tableau créé ressemble au modèle attendu.
```
def formatage(feuille, min_row, max_row, min_col, max_col):

    # Définir les styles de bordure
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Définir l'alignement au centre
    center_alignment = Alignment(horizontal='center', vertical='center')

    # Appliquer les bordures et l'alignement aux cellules
    for row in feuille.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = thin_border
            cell.alignment = center_alignment

    # Fusionner les cellules (exemple)
    feuille.merge_cells("A2:B2")
    feuille.merge_cells("A3:A4")
    print(f"Formatage appliqué à la feuille '{feuille.title}'.")
```
---
#### Problèmes rencontrés et amélioration possibles

<b>Problèmes</b>
- Quand on ne sauvegarde pas le fichier excel, on ne trouve pas  le mot "SHEET".
- Si une page est déjà existante, on ne peut pas la modifier via le code que j'ai fait.
- Si le fichier Excel n'est pas fermé lors de la compilation du code, une erreur apparait.

<b>Améliorations possibles</b>
- Faire un programme qui permet à l'utilisateur de pourvoir choisir quel fichier examiner et traiter.
- Améliorer la lecture des document en évitant de rechercher un mot ou nom particulier mais généraliser la recherhce.


---
#### Conclusion
Au final, malgrès quelques problèmes à régler et quelques améliorations possibles, j'arrive bien à faire de nouvelles pages contenant le tableau en fonction des valeurs trouvées tout en suivant le modèle demandé.
Durant cet exercice, je me suis familiarisé avec la bibliothèque pandas, j'ai découvert comment lier python et Excel et je suis content du résultat que j'ai obtenu.