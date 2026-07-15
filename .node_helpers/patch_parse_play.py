# -*- coding: utf-8 -*-
"""
Patch de secours pour parse-play : evite un crash sur un champ (app_id)
que Google a visiblement retire/deplace de sa reponse interne.
Ce champ n'est pas utilise par notre script (on utilise notre propre
app_id de config), donc ce patch est sans risque.

USAGE : lance ce script DEPUIS le dossier .node_helpers :
    cd .node_helpers
    python patch_parse_play.py
"""

import os

FILE_PATH = os.path.join("node_modules", "parse-play", "dist", "index.js")

OLD = "app_id: payload[1][11][0][0] || data[77][0],"
NEW = "app_id: payload[1][11]?.[0]?.[0] || data[77]?.[0],"


def main():
    if not os.path.exists(FILE_PATH):
        print(f"ERREUR : fichier introuvable ({FILE_PATH}).")
        print("Assure-toi de lancer ce script DEPUIS le dossier .node_helpers.")
        return

    with open(FILE_PATH, "r", encoding="utf-8") as f:
        content = f.read()

    if OLD in content:
        content = content.replace(OLD, NEW)
        with open(FILE_PATH, "w", encoding="utf-8") as f:
            f.write(content)
        print("Patch appliqué avec succès.")
    elif NEW in content:
        print("Le patch était déjà appliqué (rien à faire).")
    else:
        print("Le texte à remplacer n'a pas été trouvé.")
        print("Le fichier a peut-être une version différente de parse-play.")


if __name__ == "__main__":
    main()
