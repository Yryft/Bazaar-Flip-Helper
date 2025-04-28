import json, requests

ah_data = []
auction_nb_total = 0
page = 0

# Load your custom items
with open("updater/items.json", "r") as f:
    items_data = json.load(f)

# Create mappings
wanted_item_names = {info["name"] for info in items_data.values()}  # set of names
name_to_id = {info["name"]: item_id for item_id, info in items_data.items()}  # name -> id

# Load the craft ingredients
with open("updater/ingredients.json", "r") as f:
    recipes_data = json.load(f)

# Create a set of item_ids that are used as ingredients
ingredient_item_ids = set()
for recipe in recipes_data.values():
    for ingredient_id in recipe["ingredients"].keys():
        ingredient_item_ids.add(ingredient_id)

while True:
    print(f"Processing page {page+1}")
    api = f"https://api.hypixel.net/v2/skyblock/auctions?page={page}"
    response = requests.get(api)
    data = response.json()
    auction_nb = 0
    
    for auction in data["auctions"]:
        if auction.get("bin") and auction["item_name"] in wanted_item_names:
            item_id = name_to_id[auction["item_name"]]
            # Only consider if this item_id is needed as an ingredient
            if item_id in ingredient_item_ids:
                temp_data = {
                    "item_id": item_id,
                    "price": auction["starting_bid"]
                }
                ah_data.append(temp_data)
                auction_nb += 1
                auction_nb_total += 1
    print(f"{auction_nb} auctions processed in page {page+1}.\n")
    page += 1
    if page >= data["totalPages"]:
        print(f"{page} pages processed.\n{auction_nb_total} total auctions processed.")
        break

# Now select only the cheapest auction for each item_id
lowest_price_auctions = {}

for auction in ah_data:
    item_id = auction["item_id"]
    price = auction["price"]
    
    if item_id not in lowest_price_auctions or price < lowest_price_auctions[item_id]["price"]:
        lowest_price_auctions[item_id] = auction

# Convert to list
lowest_ah_data = list(lowest_price_auctions.values())

# Save results
with open("lowest_auction_data.json", "w") as f:
    json.dump(lowest_ah_data, f, indent=4)

print(f"Saved {len(lowest_ah_data)} lowest price auctions.")
