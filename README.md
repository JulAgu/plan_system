# Plan_system
Systéme pour construire un plan global des essais expérimentals. Une application pour integrer et transformer de maniére dynamique des fichiers issus de R-expé.

# Setup
Je recommande d'utiliser un environnement contenant les dépendances suivantes : 

**python 3.11.9**
```
numpy                     2.1.0
openpyxl                  3.1.5
pandas                    2.2.2
pyinstaller               6.10.0
xlsxwriter                3.2.0
```

# Utilisation
Pour utiliser cet outil il suffit d'éxecuter main/app.py et de suivre les instructions de l'interface graphique.

Si vous voulez construire votre propre .exe de l'aplication, je recommande vivement utiliser pyinstaller. En partant de la racine du dépot :
```
cd main
pyinstaller --onedir app.py
```
