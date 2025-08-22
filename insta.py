import os
import json
import glob
import pandas as pd
from datetime import datetime
import shutil

# === 1) Init fichier Excel ===
def initcsvfile(filename="subscriber_history.xlsx"):
    if not os.path.exists(filename):
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            pd.DataFrame(columns=["Pseudo"]).to_excel(writer, sheet_name="ABO", index=False)
            pd.DataFrame(columns=["Pseudo", "Profil"]).to_excel(writer, sheet_name="FAN", index=False)
            pd.DataFrame(columns=["Pseudo", "Profil"]).to_excel(writer, sheet_name="FDP", index=False)
        print(f"✅ Fichier {filename} créé (feuilles: ABO, FAN, FDP).")

# === 2) Trouver automatiquement le dossier data Instagram ===
def find_instagram_data_folder(base="."):
    base = os.path.abspath(base)
    for entry in os.listdir(base):
        entry_path = os.path.join(base, entry)
        if os.path.isdir(entry_path) and "instagram" in entry.lower():
            candidate = os.path.join(entry_path, "connections", "followers_and_following")
            if os.path.isdir(candidate):
                return candidate
    return None

# === 3) Lecture followers ===
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

# === 4) Lecture following ===
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

# === 5) Comparer avec Excel existant ===
def load_previous_data(filename):
    if not os.path.exists(filename):
        return set(), set(), set()
    try:
        abo = set(pd.read_excel(filename, sheet_name="ABO")["Pseudo"].dropna().astype(str).str.lower())
        fan = set(pd.read_excel(filename, sheet_name="FAN")["Pseudo"].dropna().astype(str).str.lower())
        fdp = set(pd.read_excel(filename, sheet_name="FDP")["Pseudo"].dropna().astype(str).str.lower())
        return abo, fan, fdp
    except Exception:
        return set(), set(), set()

# === 6) Traitement principal ===
def process_instagram_data(outfile="subscriber_history.xlsx"):
    initcsvfile(outfile)

    folder = find_instagram_data_folder()
    if not folder:
        print("⚠️ Impossible de trouver un dossier Instagram avec connections/followers_and_following")
        return

    # Nouvelles données
    followers = load_followers(folder)
    following = load_following(folder)
    followers_set = set(followers.keys())
    following_set = set(following.keys())

    abo_set = followers_set & following_set
    fan_set = followers_set - following_set
    fdp_set = following_set - followers_set

    ABO = sorted([followers[u][0] for u in abo_set], key=str.lower)
    FAN = sorted([(followers[u][0], followers[u][1]) for u in fan_set], key=lambda x: x[0].lower())
    FDP = sorted([(following[u][0], following[u][1]) for u in fdp_set], key=lambda x: x[0].lower())

    # Anciennes données
    prev_abo, prev_fan, prev_fdp = load_previous_data(outfile)

    new_ABO = [f for f in ABO if f.lower() not in prev_abo]
    new_FAN = [f for f in [x[0] for x in FAN] if f.lower() not in prev_fan]
    new_FDP = [f for f in [x[0] for x in FDP] if f.lower() not in prev_fdp]

    # Sauvegarde Excel
    with pd.ExcelWriter(outfile, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
        pd.DataFrame(ABO, columns=["Pseudo"]).to_excel(writer, sheet_name="ABO", index=False)
        pd.DataFrame(FAN, columns=["Pseudo", "Profil"]).to_excel(writer, sheet_name="FAN", index=False)
        pd.DataFrame(FDP, columns=["Pseudo", "Profil"]).to_excel(writer, sheet_name="FDP", index=False)

        # Feuille NEW
        new_data = []
        for user in new_ABO:
            new_data.append({"Categorie": "ABO", "Pseudo": user})
        for user in new_FAN:
            href = followers[user.lower()][1] if user.lower() in followers else ""
            new_data.append({"Categorie": "FAN", "Pseudo": user, "Profil": href})
        for user in new_FDP:
            href = following[user.lower()][1] if user.lower() in following else ""
            new_data.append({"Categorie": "FDP", "Pseudo": user, "Profil": href})

        if new_data:
            pd.DataFrame(new_data).to_excel(writer, sheet_name="NEW", index=False)
        else:
            pd.DataFrame([{"Info": "Aucune nouveauté"}]).to_excel(writer, sheet_name="NEW", index=False)

    # Renommer le dossier parent (version sûre)
    parent_dir = os.path.dirname(os.path.dirname(folder))  # le dossier "instagram...."
    today_str = datetime.now().strftime("%d-%m-%Y")
    base_new_name = os.path.join(os.path.dirname(parent_dir), f"ABO_INSTA_{today_str}")
    new_name = base_new_name
    counter = 1
    while os.path.exists(new_name):
        new_name = f"{base_new_name}_{counter}"
        counter += 1

    os.rename(parent_dir, new_name)
    print(f"📂 Dossier renommé en : {new_name}")

    # Résumé
    print("— RÉSUMÉ —")
    print(f"ABO (réciproques) : {len(ABO)} (+{len(new_ABO)} nouveaux)")
    print(f"FAN (ils te suivent, tu ne les suis pas) : {len(FAN)} (+{len(new_FAN)} nouveaux)")
    print(f"FDP (tu les suis, ils ne te suivent pas) : {len(FDP)} (+{len(new_FDP)} nouveaux)")
    print(f"(Totaux) Followers: {len(followers)} | Following: {len(following)}")
    print(f"✅ Données écrites dans: {outfile}")

# === 7) Lancer ===
if __name__ == "__main__":
    process_instagram_data(outfile="subscriber_history.xlsx")
