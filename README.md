Convertisseur de Colonnes - Documentation
Ce script Python vous permet d'extraire et de convertir des colonnes spécifiques d'un fichier Excel dans différents formats de sortie tels que CSV, TXT, SQL, JSON, et XML.

Utilisation
Installer les dépendances

Assurez-vous d'avoir Python installé sur votre système. Vous pouvez installer les dépendances en exécutant la commande suivante :

pip install pandas sqlalchemy lxml
Exécuter le script

Placez le fichier Excel que vous souhaitez convertir dans le même répertoire que le script. Ensuite, changer le nom du fichier dans le code puis exécutez le script en utilisant la commande suivante :



python ou python3 script.py
Suivez les instructions

Le script vous demandera de choisir les colonnes que vous souhaitez extraire.
Sélectionnez le format de sortie parmi les options : CSV, TXT, SQL, JSON, ou XML.
Suivez les instructions supplémentaires pour certains formats, par exemple, spécifiez le nom de la table et les noms de colonnes pour le format SQL.
Résultats

Le script générera un fichier de sortie dans le même répertoire que le script. Le nom du fichier de sortie sera automatiquement généré avec le format spécifié.

Formats de Sortie Supportés
CSV: Fichier CSV avec les colonnes sélectionnées.
TXT: Fichier texte avec les colonnes sélectionnées, séparées par des tabulations.
SQL: Fichier SQL avec des instructions d'INSERT INTO pour une base de données SQLite temporaire.
JSON: Fichier JSON avec les données au format JSON.
XML: Fichier XML avec les données au format XML.
N'hésitez pas à explorer le code source pour plus de détails sur la logique de conversion pour chaque format.