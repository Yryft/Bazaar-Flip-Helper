import json, requests
from collections import defaultdict

API_URL = "https://api.hypixel.net/skyblock/bazaar"

global times_crafted
times_crafted = 0

# Charger les données depuis le fichier crafts.json
def load_craft_data(craft_filename="crafts.json"):
    try:
        with open(craft_filename, encoding="utf-8") as file:
            return json.load(file)

    except FileNotFoundError:
        print("Erreur", f"Le fichier {craft_filename} est introuvable.")
        return {}
    except json.JSONDecodeError:
        print("Erreur", f"Erreur lors de la lecture du fichier {craft_filename}.")
        return {}
    
def get_bazaar_data():
    try:
        response = requests.get(API_URL)
        response.raise_for_status()  # Vérifie si la requête a réussi
        return response.json()
    except requests.exceptions.RequestException as e:
        print("Erreur", f"Erreur de connexion à l'API du bazaar : {e}")
        return None

def parse_ingredients(recipe):
    totals = defaultdict(int)
    for value in recipe.values():
        if isinstance(value, str) and ":" in value:
            item_id, amount = value.split(":")
            totals[item_id] += int(amount)

    # This returns a dictionary with item_id as keys and their corresponding amounts as values
    return {k: v for k, v in totals.items()}


def load_crafts(craft_data, bazaar_data):
    possible_crafts_bazaar = {
            item:      
            {
                "item_id": item,
                "ingredients": parse_ingredients(craft_data[item]["recipe"])
            }
            for item in craft_data
            if item in bazaar_data["products"] and "recipe" in craft_data[item]
    }

    with open(f"Ingredients.json", "w", encoding="utf-8") as file:
        json.dump(possible_crafts_bazaar, file, indent=4)
    return possible_crafts_bazaar

def calculate_profit(bazaar_data, craft_data, ingredients):
    profits = []
    craft = {}
    item_craft_profit = 0
    if not ingredients:
        print("Aucun ingredients")
        
    if not bazaar_data:
        print("Aucun bazaar")
        
    if not craft_data:
        print("Aucun craft")
        
    for item_id, product in bazaar_data["products"].items():
        if item_id in ingredients:
            item_craft_profit = 0
            for ingredient, quantity in ingredients[item_id]["ingredients"].items():
                if ingredient in bazaar_data["products"] and bazaar_data["products"][ingredient]["buy_summary"] and bazaar_data["products"][ingredient]["sell_summary"]:
                    buy_price = bazaar_data["products"][ingredient]["sell_summary"][0]["pricePerUnit"]
                    sell_price = bazaar_data["products"][ingredient]["buy_summary"][0]["pricePerUnit"]
                    if (sell_price / buy_price - 1)*100 > 80 and (sell_price - buy_price) > 100:
                        item_craft_profit = 0
                        break
                    profit = round(sell_price - buy_price, 1)
                    item_craft_profit += profit*quantity
            craft = {
                "craftable": True if item_craft_profit > 0 else False,
                "craft_profit": item_craft_profit,
                "materials": {
                    craft_data.get(ingredient, {}).get("name", ingredient): quantity
                    for ingredient, quantity in ingredients[item_id]["ingredients"].items()
                }
            }
            print(f"Item: {item_id}, Profit total du craft: {round(item_craft_profit):,}\n".replace(",", " ") if item_craft_profit > 0 else f"Profit trop grand pour le craft de {item_id}\n")
        else:
            craft = {
                "craftable": False,
                "craft_profit": 0,
                "materials": {}
            }
        # Vérifier si les données d'achat/vente sont disponibles
        if product["buy_summary"] and product["sell_summary"]:
            buy_price = product["sell_summary"][0]["pricePerUnit"]
            sell_price = product["buy_summary"][0]["pricePerUnit"]

            if buy_price == 0 or sell_price == 0:
                continue  # Si les prix sont à 0, on ignore cet item

            profit = round(sell_price - buy_price, 1)

            # Chercher le nom et l'ID wiki dans craft.json
            item_name = craft_data.get(item_id, {}).get("name", item_id)
            if not item_name == item_id:
                wiki_url = f"https://wiki.hypixel.net/{item_name}"  # Lien vers le wiki
            else:
                wiki_url = None
                
            # Ajouter les données au résultat
            profits.append({
                'item': item_id,
                'name': item_name,
                'buy_price': buy_price,
                'sell_price': sell_price,
                'profit': profit,
                'craft': craft,
                'sell_volume': product["sell_summary"][0]["amount"],
                'buy_volume': product["buy_summary"][0]["amount"],
                'wiki': wiki_url  # Lien vers le wiki
            })
    with open(f"Profits.json", "w", encoding="utf-8") as file:
        json.dump(profits, file, indent=4)
    return profits

calculate_profit(get_bazaar_data(), load_craft_data(), load_crafts(load_craft_data(), get_bazaar_data()))