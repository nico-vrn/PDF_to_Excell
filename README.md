# Traitement de vulnérabilités Checkmarx

<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/nico-vrn/PDF_to_Excell">
    <img src="images/logo.jpg" alt="Logo" width="100" height="80">
  </a>

  <h3 align="center">PDF to excell</h3>

  <p align="center">
    Traitement des fichiers d'analyses checkmarx
    <br />
   </p>
</div>

Ce programme écrit en Python récupère toutes les vulnérabilités d'un fichier d'analyse Checkmarx et les intégres dans un fichier Excel.

## Installation
1. Clonez le dépôt :

```sh
git clone https://github.com/nico-vrn/PDF_to_Excell.git
```

2. Installez les dépendances :

```sh
pip install -r requirements.txt
```

# Utilisation
1. Placez votre fichier d'analyse Checkmarx dans le répertoire du projet.

2. Exécutez le programme en utilisant la commande suivante :

```sh
python extract_pdf.py
```

3. Entrez le nom du fichier PDF que vous souhaitez analyser (sans l'extension .pdf).

4. Choisissez la langue du fichier PDF (fr ou en).

5. Le fichier Excel contenant les vulnérabilités et les classes sera généré dans le répertoire du projet.

# Fonctionnement du script
Le script suit les étapes suivantes :

1. Demander à l'utilisateur le nom du fichier PDF à récupérer, le nom du fichier Excel à créer et la langue du fichier PDF.
2. Trouver les pages à extraire en fonction de la langue choisie.
3. Extraire le texte de chaque page et le stocker dans une variable globale pour chaque page.
4. Extraire les données de toutes les pages et les stocker dans une variable.
5. Compter le nombre de lignes de la variable.
6. Écrire le résultat dans un fichier texte nommé "data.txt".
7. Supprimer les numéros de page dans les données extraites.

# Contributions
Les contributions sont les bienvenues ! Si vous souhaitez contribuer à ce projet, veuillez suivre les étapes suivantes :

1. Fork ce projet.

2. Créez une branche pour vos modifications :

```sh
git checkout -b ma-nouvelle-fonctionnalite
```

3. Faites vos modifications et commit :

```sh
git commit -am 'Ajout d'une nouvelle fonctionnalité'
``` 

4. Push les modifications sur votre branche :

```sh 
git push origin ma-nouvelle-fonctionnalite
```

5. Faites une pull request depuis votre branche vers la branche principale de ce projet.

# Licence
Ce projet est sous licence MIT. Veuillez consulter le fichier `LICENSE` pour plus d'informations.

# Auteurs
Lefranc Nicolas : Développeur principal
