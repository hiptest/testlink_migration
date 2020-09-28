# testlink_migration

This script performs a migration from a Testlink project to CS project, in order to build the same scheme as the spreadsheet template (.xlxs) available in the app.

Install dependecencies
- Python 3
- lxml
- bs4
- openpyxl

Export your TestLink project as .XML format
Execute the python script

1)      Dans Testlink, exporter le cahier de test au format xml

2)      Exécuter le script python :

$ python <nom_du_fichier_script.py> –i <chemin_vers_fichier_export_testlink.xml>

3)      Le script a généré un fichier export.xlsx

4)      Dans Hiptest, créer un nouveau projet à partir du fichier export.xlsx
