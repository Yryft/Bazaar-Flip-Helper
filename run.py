import requests, time, json, os, platform, subprocess, threading
import xlsxwriter
from collections import defaultdict

# URL de l'API du bazaar
API_URL = "https://api.hypixel.net/skyblock/bazaar"


# Fonction mise à jour pour enregistrer ou modifier le fichier Excel
def save_to_excel(data, filename="bazaar_helper.xlsx"):
    # Remove existing file
    if os.path.exists(filename):
        os.remove(filename)

    # Create workbook with XlsxWriter (enable dynamic-array support)
    workbook = xlsxwriter.Workbook(filename)
    workbook.set_calc_mode('auto')

    # Formats
    header_fmt = workbook.add_format({"bold": True, "align": "center"})
    number_fmt = workbook.add_format({"num_format": "#,##0.00"})
    integer_fmt = workbook.add_format({"num_format": "#,##0"})
    bold_fmt = workbook.add_format({"bold": True})
    center_fmt = workbook.add_format({"align": "center"})
    yellow_fill = workbook.add_format({"bg_color": "#FFFF00", "align": "center", "bold": True})
    budget_fmt = workbook.add_format({"bg_color": "#FFFF00", "num_format": "#,##0", "align": "center", "bold": True})
    red_fill = workbook.add_format({"bg_color": "#FF0000"})

    # === Sheet1: Bazaar Analysis ===
    sheet = workbook.add_worksheet("Bazaar Analysis")
    headers = ["Item", "Buy Price", "Sell Price", "Profit", "Supply & Demand", "Sold last week", "Profit per coin", "Craftable", "Craft Profit"]
    sheet.write_row(0, 0, headers, header_fmt)

    # Write data rows
    for idx, values in enumerate(data, start=1):
        row = idx
        item_name = values["name"]
        wiki_url = values.get("wiki")
        # Item cell with hyperlink or centered name
        if wiki_url:
            sheet.write_url(row, 0, wiki_url,
                            cell_format=workbook.add_format({"font_color": "blue", "underline": 1, "align": "center"}),
                            string=item_name)
        else:
            sheet.write(row, 0, item_name, center_fmt)

        sheet.write(row, 1, values["buy_price"], number_fmt)
        sheet.write(row, 2, values["sell_price"], number_fmt)
        sheet.write(row, 3, values["profit"], number_fmt)
        sheet.write(row, 4, f"{values['buy_volume']}/{values['sell_volume']}")
        sheet.write(row, 5, values['buy_moving_week'], integer_fmt)
        profit_per_coin = values['profit'] / values['buy_price'] if values['buy_price'] else 0
        sheet.write(row, 6, profit_per_coin, number_fmt)
        sheet.write(row, 7, values['craft']['craftable'])
        sheet.write(row, 8, values['craft']['craft_profit'], number_fmt)

    # Add table on sheet1
    last_row = len(data)
    sheet.add_table(0, 0, last_row, len(headers)-1, {
        "name": "BazaarTable",
        "style": "Table Style Medium 9",
        "columns": [{"header": h} for h in headers]
    })
    # Adjust column widths
    for col in range(len(headers)):
        sheet.set_column(col, col, 15)

    # === Sheet2: Crafts ===
    craft_sheet = workbook.add_worksheet("Crafts")
    col = 0
    for values in data:
        if values['craft']['craftable']:
            materials = values['craft']['materials']
            craft_sheet.write(0, col, values['name'], header_fmt)
            craft_sheet.write(0, col+1, f"Q-{values['name']}", header_fmt)
            for r, (mat, qty) in enumerate(materials.items(), start=1):
                craft_sheet.write(r, col, mat)
                craft_sheet.write(r, col+1, qty)
            col += 2
    last_col = col - 1
    # Define headers list
    craft_headers = []
    for values in data:
        if values['craft']['craftable']:
            craft_headers.extend([values['name'], f"Q-{values['name']}"])
    # Add table on crafts sheet
    craft_sheet.add_table(0, 0, 15, last_col, {
        "name": "CraftTable",
        "style": "Table Style Medium 9",
        "columns": [{"header": h} for h in craft_headers]
    })
    for c in range(last_col+1):
        craft_sheet.set_column(c, c, 15)

    # === Sheet3: Flip Simulator ===
    sim = workbook.add_worksheet("Flip Simulator")
    sim.write("A1", "Simulateur de Flip", workbook.add_format({"bold": True, "font_size": 14}))
    labels = ["Sélectionne un item à flipper :", "Entrez votre budget en coins :", "Quantité possible à acheter :",
              "Profit total estimé :", "Prix unitaire d'achat :", "Prix unitaire de vente :", "Item craftable : ",
              "Profit du craft :", "Prix total du craft: ", "Quantité possible à craft: ", "Profit total estimé: ", "Vendus dans les 7 derniers jours: "]
    for i, txt in enumerate(labels, start=3):
        sim.write(f"A{i}", txt, bold_fmt)
    sim.write("B1", "UPDATE RED CELLS", workbook.add_format({"bg_color": "#FF0000", "align": "center", "bold": True}))
    sim.write("C1", "Ingrédients :", bold_fmt)
    sim.write("D1", "Quantité :", bold_fmt)
    sim.write("E1", "Quantité totale :", bold_fmt)

    # Preselect first item
    sim.write_formula("B3", "='Bazaar Analysis'!$A$2", yellow_fill)
    sim.write_number("B4", 100000000, budget_fmt)
    item_count = len(data)
    sim.data_validation("B3", {"validate": "list", "source": f"='Bazaar Analysis'!$A$2:$A${item_count+1}"})

    # Static formulas
    sim.write_formula("B5", '=IFERROR(ROUNDDOWN(B4 / VLOOKUP(B3, \'Bazaar Analysis\'!$A:$I, 2, FALSE), 0), 0)', integer_fmt)
    sim.write_formula("B6", '=IFERROR(B5 * VLOOKUP(B3, \'Bazaar Analysis\'!$A:$J, 4, FALSE), 0)', number_fmt)
    sim.write_formula("B7", '=IFERROR(VLOOKUP(B3, \'Bazaar Analysis\'!$A:$J, 2, FALSE), 0)', number_fmt)
    sim.write_formula("B8", '=IFERROR(VLOOKUP(B3, \'Bazaar Analysis\'!$A:$J, 3, FALSE), 0)', number_fmt)
    sim.write_formula("B9", '=IFERROR(VLOOKUP(B3, \'Bazaar Analysis\'!$A:$J, 8, FALSE), 0)')
    sim.write_formula("B10", '=IFERROR(VLOOKUP(B3, \'Bazaar Analysis\'!$A:$J, 9, FALSE), 0)', number_fmt)
    sim.write_formula("B12", '=IFERROR(INT(B4/B11), 0)', integer_fmt)
    sim.write_formula("B14", '=IFERROR(VLOOKUP(B3, \'Bazaar Analysis\'!$A:$J, 6, FALSE), 0)', integer_fmt)

    # Dynamic array spills
    sim.write_formula("C2", '''=IFERROR(LET(item, B3, col, MATCH(item, Crafts!$1:$1, 0), nbLignes, COUNTA(INDEX(Crafts!$2:$100,,col)), lignes, SEQUENCE(nbLignes), INDEX(Crafts!$2:$100, lignes, col)), "No craft")''', red_fill)
    sim.write_formula("D2", '''=IFERROR(LET(item, B3, col, MATCH(item, Crafts!$1:$1, 0), nbLignes, COUNTA(INDEX(Crafts!$2:$100,,col)), lignes, SEQUENCE(nbLignes), INDEX(Crafts!$2:$100, lignes, col+1)), "")''', red_fill)
    sim.write_formula("E2", '''=IFERROR(LET(item, B3, col, MATCH(item, Crafts!$1:$1, 0), rowCount, COUNTA(INDEX(Crafts!$2:$100,,col)), rows, SEQUENCE(rowCount), INDEX(Crafts!$2:$100, rows, col+1) * B12), "")''', red_fill)

    sim.write_dynamic_array_formula("B11:B11", '''=IFERROR(LET(item, B3, col, MATCH(item, Crafts!$1:$1, 0), nbLignes, COUNTA(INDEX(Crafts!$2:$100,,col)), lignes, SEQUENCE(nbLignes), ingredients, INDEX(Crafts!$2:$100, lignes, col), quantites, INDEX(Crafts!$2:$100, lignes, col+1), prix_unitaires, XLOOKUP(ingredients, 'Bazaar Analysis'!$A:$A, 'Bazaar Analysis'!$B:$B, "N/A"), SUM(prix_unitaires*quantites)), "Erreur")''', red_fill)

    sim.write_formula("B13", '=B10 * B12', integer_fmt)

    # Adjust columns width
    sim.set_column("A:A", 30)
    sim.set_column("B:B", 35)
    sim.set_column("C:E", 25)

    workbook.close()


# Charger les données depuis le fichier crafts.json
def load_craft_data(craft_filename="updater/items.json"):
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


def load_crafts():
    with open(f"updater/Ingredients.json", "r", encoding="utf-8") as file:
        possible_crafts_bazaar = json.load(file)
    return possible_crafts_bazaar

def calculate_profit(bazaar_data, craft_data, ingredients):
    profits = []
    craft = {}
        
    for item_id, product in bazaar_data["products"].items():
        print(f"{craft_data.get(item_id, {}).get("name", item_id)} : ")
        if item_id in ingredients:
            item_craft_profit = 0
            total_cost = 0
            for ingredient, quantity in ingredients[item_id]["ingredients"].items():
                if ingredient in bazaar_data["products"] and bazaar_data["products"][ingredient]["buy_summary"] and bazaar_data["products"][ingredient]["sell_summary"]:
                    buy_price = bazaar_data["products"][ingredient]["sell_summary"][0]["pricePerUnit"]
                    sell_price = bazaar_data["products"][ingredient]["buy_summary"][0]["pricePerUnit"]
                    total_cost += round(buy_price, 1) * quantity
            try:
                item_craft_profit = round(bazaar_data["products"][item_id]["sell_summary"][0]["pricePerUnit"] - total_cost) if total_cost > 0.05*bazaar_data["products"][item_id]["sell_summary"][0]["pricePerUnit"] else 0
                print(f"Total Craft Cost: {total_cost}/Direct Cost: {product['sell_summary'][0]['pricePerUnit']}\nProfit: {item_craft_profit}")
            except:
                item_craft_profit = 0
                print(f"Total Craft Cost: {total_cost}/Direct Cost: N/A\nProfit: N/A")
            craft = {
                "craftable": True if item_craft_profit > 0 else False,
                "craft_profit": item_craft_profit,
                "materials": {
                    craft_data.get(ingredient, {}).get("name", ingredient): quantity
                    for ingredient, quantity in ingredients[item_id]["ingredients"].items()
                }
            }
        else:
            craft = {
                "craftable": False,
                "craft_profit": 0,
                "materials": {}
            }
        # Vérifier si les données d'achat/vente sont disponibles
        if product["sell_summary"]:
            buy_price = product["sell_summary"][0]["pricePerUnit"]
            sell_volume = product["sell_summary"][0]["amount"]
        else:
            buy_price = 0
            sell_volume = 0
            
        if product["buy_summary"]:
            sell_price = product["buy_summary"][0]["pricePerUnit"]
            buy_volume = product["buy_summary"][0]["amount"]
        else:
            sell_price = 0
            buy_volume = 0

        profit = round(sell_price - buy_price, 1)

        # Chercher le nom et l'ID wiki dans craft.json
        item_name = craft_data.get(item_id, {}).get("name", item_id)
        wiki_url = craft_data.get(item_id, {}).get("wiki", "")
            
        print(f"Buy Price: {buy_price} | Sell Price: {sell_price} | Profit: {profit}\nWiki: {wiki_url}\n")
          
        # Ajouter les données au résultat
        profits.append({
            'item': item_id,
            'name': item_name,
            'buy_price': buy_price,
            'sell_price': sell_price,
            'profit': profit,
            'craft': craft,
            'sell_volume': sell_volume,
            'buy_volume': buy_volume,
            'buy_moving_week': product["quick_status"]["buyMovingWeek"],
            'wiki': wiki_url  # Lien vers le wiki
        })
    return profits

# Fonction pour exécuter l'analyse et afficher les résultats
def run_analysis():
    # Charger les données craft
    craft_data = load_craft_data()
    if not craft_data:
        return []

    # Récupérer les données du bazaar
    bazaar_data = get_bazaar_data()
    if not bazaar_data:
        return []

    # Calculer les profits
    results = calculate_profit(bazaar_data, craft_data, load_crafts())
    return results

def open_file(filename):
    if platform.system() == "Windows":
        os.startfile(filename)
    elif platform.system() == "Darwin":  # macOS
        subprocess.call(["open", filename])
    else:  # Linux
        subprocess.call(["xdg-open", filename])


# Lancer le programme
def start():
    try:
        subprocess.call(["taskkill", "/f", "/im", "excel.exe"])
    except:
        print("Excel est fermé.")
        
    filename = "bazaar_helper.xlsx"
    try:
        results = run_analysis()
        if not results:
            print("Aucun résultat. Aucune donnée à sauvegarder.")
            input("\n\nAppuie sur Entrée pour quitter...")
            return
    except Exception as e:
        print("Erreur : Une erreur s'est produite :", e)
        input("\n\nAppuie sur Entrée pour quitter...")
        return


    try:
        save_to_excel(results, filename)
        print("Succès : Données sauvegardées dans", filename)
        input("\n\nAppuie sur Entrée pour ouvrir le fichier...")
        open_file(filename)
        return
    except Exception as e:
        print("Erreur : Une erreur s'est produite :", e)
        input("\n\nAppuie sur Entrée pour quitter...")
        return

# Lancer le script
start()