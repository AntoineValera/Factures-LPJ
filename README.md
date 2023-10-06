# Factures-LPJ

Une petite application pour renommer les fichier de factures plus facilement et de facon cohérente

Pour recompiler apres modificationde main .py (testé avec python 3.8.10) 

```
pip install pyinstaller
pyinstaller --distpath ./dist my_app.spec
```
my_app.spec peut etre facilement modifié.

Les pdf sont affichés avec pyMuPDF (pip install PyMuPDF)

N'oubliez pas d'inclure de logo dans le dossier de compilation (./dist)
