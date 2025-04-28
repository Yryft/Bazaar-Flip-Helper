import requests, zipfile, io, shutil, re, json, os
from collections import defaultdict

API_URL = "https://api.hypixel.net/skyblock/bazaar"

repo = "NotEnoughUpdates/NotEnoughUpdates-REPO"
branch = "master"
folder_to_extract = "items"
# Path to the folder containing the JSON files
folder_path = "NotEnoughUpdates-REPO-master/items"

zip_url = f"https://github.com/{repo}/archive/refs/heads/{branch}.zip"
# Output dictionary
output = {}


def get_neu_data():
    # Download and extract only the folder
    print("Downloading NEU's REPO...")
    response = requests.get(zip_url)
    with zipfile.ZipFile(io.BytesIO(response.content)) as zip_ref:
        print("Extracting...")
        for member in zip_ref.namelist():
            if member.startswith(f"{repo.split('/')[-1]}-{branch}/{folder_to_extract}"):
                zip_ref.extract(member)
        print(".json files downloaded !")

# Regex to remove Minecraft color codes (like §9, §7, etc.)
def remove_color_codes(text):
    return re.sub(r"§.", "", text)

def jsons():
    # Iterate over each file in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".json"):
            file_path = os.path.join(folder_path, filename)
            with open(file_path, "r", encoding="utf-8") as file:
                try:
                    data = json.load(file)
                    item_id = data.get("internalname")
                    name = remove_color_codes(data.get("displayname", ""))
                    recipe = data.get("recipe", {})
                    item_id = data.get("internalname", "")
                    wiki_links = data.get("info", [])
                    wiki_url = wiki_links[-1] if wiki_links else ""

                    if item_id:
                        output[item_id] = {
                            "name": name,
                            "recipe": recipe,
                            "itemId": item_id,
                            "wiki": wiki_url
                        }
                except json.JSONDecodeError:
                    print(f"Error decoding JSON in file: {filename}")

    # Save to a single JSON file
    output_path = "items.json"
    with open(output_path, "w", encoding="utf-8") as outfile:
        json.dump(output, outfile, indent=4)

    print("Succès : items.json mis à jour.")
    shutil.rmtree('NotEnoughUpdates-REPO-master')
    print("Succès : NotEnoughUpdates-REPO-master supprimé.")

def parse_ingredients(recipe):
    totals = defaultdict(int)
    for value in recipe.values():
        if isinstance(value, str) and ":" in value:
            item_id, amount = value.split(":")
            totals[item_id] += int(amount)

    # This returns a dictionary with item_id as keys and their corresponding amounts as values
    return {k: v for k, v in totals.items()}

def load_craft_data(craft_filename="items.json"):
    try:
        with open(craft_filename, encoding="utf-8") as file:
            return json.load(file)

    except FileNotFoundError:
        print("Erreur", f"Le fichier {craft_filename} est introuvable.")
        return {}
    except json.JSONDecodeError:
        print("Erreur", f"Erreur lors de la lecture du fichier {craft_filename}.")
        return {}

def load_crafts(craft_data, bazaar_data):
    possible_crafts_bazaar = {
            item:      
            {
                "item_id": item,
                "ingredients": parse_ingredients(craft_data[item]["recipe"])
            }
            for item in craft_data
            if item in bazaar_data["products"] and craft_data[item]["recipe"]
    }
    with open(f"Ingredients.json", "w", encoding="utf-8") as file:
        json.dump(possible_crafts_bazaar, file, indent=4)
    print("Succès : Ingredients.json mis à jour.")


def get_bazaar_data():
    try:
        response = requests.get(API_URL)
        response.raise_for_status()  # Vérifie si la requête a réussi
        return response.json()
    except requests.exceptions.RequestException as e:
        print("Erreur", f"Erreur de connexion à l'API du bazaar : {e}")
        return None
    
get_neu_data()
jsons()
load_crafts(load_craft_data(), get_bazaar_data())
input("\n\nAppuie sur Entrée pour quitter...")