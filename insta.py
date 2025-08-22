import os
import json
import glob
import pandas as pd

# === 1) Init fichier Excel ===
def initcsvfile(filename="subscriber_history.xlsx"):
    if not os.path.exists(filename):
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            pd.DataFrame(columns=["Pseudo"]).to_excel(writer, sheet_name="ABO", index=False)
            pd.DataFrame(columns=["Pseudo", "Profil"]).to_excel(writer, sheet_name="FAN", index=False)
            pd.DataFrame(columns=["Pseudo", "Profil"]).to_excel(writer, sheet_name="FDP", index=False)
        print(f"✅ Fichier {filename} créé (feuilles: ABO, FAN, FDP).")
    else:
        print(f"ℹ️ Fichier {filename} déjà existant.")

# === 2) Lecture followers ===
def load_followers(folder):
    followers = {}
    files = sorted(glob.glob(os.path.join(folder, "followers_*.json")))
    fallback = os.path.join(folder, "followers.json")
    if os.path.exists(fallback):
        files.insert(0, fallback)

    for path in files:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        for entry in data:
            try:
                username = entry["string_list_data"][0]["value"]
                href = entry["string_list_data"][0]["href"]
                followers[username.lower()] = (username, href)
            except (KeyError, IndexError, TypeError):
                continue
    return followers

# === 3) Lecture following ===
def load_following(folder):
    following = {}
    path = os.path.join(folder, "following.json")
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    for entry in data.get("relationships_following", []):
        try:
            username = entry["string_list_data"][0]["value"]
            href = entry["string_list_data"][0]["href"]
            following[username.lower()] = (username, href)
        except (KeyError, IndexError, TypeError):
            continue
    return following

# === 4) Traitement ===
def process_instagram_data(folder="", outfile="subscriber_history.xlsx"):
    initcsvfile(outfile)

    followers = load_followers(folder)   # dict {username.lower(): (username, href)}
    following = load_following(folder)

    followers_set = set(followers.keys())
    following_set = set(following.keys())

    # Catégories
    abo_set = followers_set & following_set   # réciproques
    fan_set = followers_set - following_set   # ils te suivent, tu ne les suis pas
    fdp_set = following_set - followers_set   # tu les suis, ils ne te suivent pas

    # Reconstruire les listes avec pseudo + lien quand dispo
    ABO = sorted([followers[u][0] for u in abo_set], key=str.lower)
    FAN = sorted([(followers[u][0], followers[u][1]) for u in fan_set], key=lambda x: x[0].lower())
    FDP = sorted([(following[u][0], following[u][1]) for u in fdp_set], key=lambda x: x[0].lower())

    # Sauvegarde Excel
    with pd.ExcelWriter(outfile, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
        pd.DataFrame(ABO, columns=["Pseudo"]).to_excel(writer, sheet_name="ABO", index=False)
        pd.DataFrame(FAN, columns=["Pseudo", "Profil"]).to_excel(writer, sheet_name="FAN", index=False)
        pd.DataFrame(FDP, columns=["Pseudo", "Profil"]).to_excel(writer, sheet_name="FDP", index=False)

    print("— RÉSUMÉ —")
    print(f"ABO (réciproques) : {len(ABO)}")
    print(f"FAN (ils te suivent, tu ne les suis pas) : {len(FAN)}")
    print(f"FDP (tu les suis, ils ne te suivent pas) : {len(FDP)}")
    print(f"(Totaux) Followers: {len(followers)} | Following: {len(following)}")
    print(f"✅ Données écrites dans: {outfile}")

# === 5) Lancer ===
if __name__ == "__main__":
    process_instagram_data(folder="", outfile="subscriber_history.xlsx")
