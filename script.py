import os
import pandas as pd
from sqlalchemy import create_engine, text, VARCHAR

def exporter_colonnes_au_format(df, colonnes, format_sortie, chemin_sortie):
    colonnes_selectionnees = df[colonnes]

    if format_sortie == 'csv':
        colonnes_selectionnees.to_csv(chemin_sortie, index=False)
    elif format_sortie == 'txt':
        colonnes_selectionnees.to_csv(chemin_sortie, index=False, header=False, sep='\t')
    elif format_sortie == 'sql':
        # Convertir les colonnes en syntaxe SQL INSERT INTO
        engine = create_engine('sqlite:///:memory:')
        table_bdd = input("Entrez le nom de la table en BDD : ")
        colonnes_bdd = input("Entrez les noms que vous souhaitez donner aux colonnes en BDD (séparés par des virgules) : ").split(',')
        colonnes_selectionnees.columns = colonnes_bdd

        with engine.connect() as connection:
            colonnes_selectionnees.to_sql(table_bdd, connection, index=False, if_exists='replace', index_label=False, dtype={colonne: VARCHAR for colonne in colonnes_bdd})

            # Extraire le script d'insertion depuis la base de données SQLite temporaire
            query = text(f'SELECT * FROM {table_bdd};')
            result = connection.execute(query)
            rows = result.fetchall()

        with open(chemin_sortie, 'w') as fichier_sql:
            # Format correct pour le script SQL avec toutes les lignes dans une seule instruction
            fichier_sql.write(f"INSERT INTO {table_bdd} VALUES\n")
            for i, row in enumerate(rows):
                valeurs_formattees = ', '.join([f"'{v}'" for v in row])
                if i < len(rows) - 1:
                    fichier_sql.write(f"({valeurs_formattees}),\n")
                else:
                    fichier_sql.write(f"({valeurs_formattees});\n")
    elif format_sortie == 'json':
        colonnes_selectionnees.to_json(chemin_sortie, orient='records', lines=True)
    elif format_sortie == 'xml':
        colonnes_selectionnees.to_xml(chemin_sortie, index=False)
    else:
        print("Format de sortie non pris en charge.")

    print(f"Colonnes '{', '.join(colonnes)}' exportées avec succès en format {format_sortie} vers {chemin_sortie}.")

if __name__ == "__main__":
    # Obtenez le chemin absolu du script
    script_directory = os.path.dirname(os.path.abspath(__file__))

    # Spécifiez le chemin vers votre fichier Excel
    chemin_fichier_excel = os.path.join(script_directory, 'exemple.xlsx')

    # Lisez le fichier Excel
    df = pd.read_excel(chemin_fichier_excel)

    # Obtenez la liste des colonnes disponibles
    colonnes_disponibles = df.columns.tolist()
    print("Colonnes disponibles dans le fichier Excel:", colonnes_disponibles)

    # Demandez à l'utilisateur de choisir les colonnes
    colonnes_a_extraire = input("Entrez le nom de la ou des colonnes (séparées par des virgules) que vous souhaitez extraire : ").split(',')
    colonnes_a_extraire = [colonne.strip() for colonne in colonnes_a_extraire]

    # Demandez à l'utilisateur de choisir le format de sortie
    format_sortie = input("Choisissez le format de sortie (csv, txt, sql, json, xml) : ").lower()

    # Spécifiez le chemin de sortie pour le fichier converti à la racine du script
    chemin_sortie = os.path.join(script_directory, f'fichier_sortie.{format_sortie}')

    exporter_colonnes_au_format(df, colonnes_a_extraire, format_sortie, chemin_sortie)
