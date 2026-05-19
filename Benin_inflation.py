import os
import hashlib
import re
from pathlib import Path
import pandas as pd
from docx import Document
from sqlalchemy import create_engine, text
from dotenv import load_dotenv  # Added for environment variable management

# 1. LOAD ENVIRONMENT CONFIGURATION Safely from your local hidden file
load_dotenv()

folder = Path("/Users/jpsossavi/Documents/Benin_inflation")
files = list(folder.glob("*.docx"))

def file_hash(path):
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()

hashes = {}
for f in files:
    h = file_hash(f)
    hashes.setdefault(h, []).append(f)

unique_files = [group[0] for group in hashes.values()]

def extract_week(file_name):
    name = file_name.replace(".docx", "")
    if "Prix_des_produits_du_" in name:
        name = name.split("Prix_des_produits_du_")[1]
    match = re.search(r"(.*?_\d{4})", name)
    if match:
        return match.group(1)
    return name

for f in unique_files[:10]:
    print("*" * 40)
    print("-" * 40)
    print(f.name)
    print(extract_week(f.name))
    print("-" * 40)

cities = ["Cotonou", "Porto-Novo", "Parakou", "Natitingou", "Bohicon", "Lokossa"]
all_data = []

for file_path in unique_files:
    try:
        doc = Document(file_path)
        table = doc.tables[1]

        rows = []
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            rows.append(cells)

        df = pd.DataFrame(rows)
        file_name = file_path.name
        week_text = extract_week(file_name)
        clean_rows = []

        for i in range(2, len(df), 2):
            if i + 1 >= len(df):
                continue

            price_row = df.iloc[i]
            variation_row = df.iloc[i + 1]
            product = price_row[0]

            for j, city in enumerate(cities):
                clean_rows.append({
                    "Semaine": week_text,
                    "Produit": product,
                    "Ville": city,
                    "Prix": price_row[j + 2],
                    "Variation": variation_row[j + 2],
                    "Source_File": file_name
                })

        all_data.append(pd.DataFrame(clean_rows))
        print("OK:", week_text, "|", file_name)

    except Exception as e:
        print("ERREUR:", file_path.name)
        print(e)

final_df = pd.concat(all_data, ignore_index=True)

final_df["Prix"] = (
    final_df["Prix"]
    .astype(str)
    .str.replace(" ", "", regex=False)
    .str.replace("\xa0", "", regex=False)
)
final_df["Prix"] = pd.to_numeric(final_df["Prix"], errors="coerce")

final_df["Variation"] = (
    final_df["Variation"]
    .astype(str)
    .str.replace(",", ".", regex=False)
    .str.replace(" ", "", regex=False)
    .str.replace("\xa0", "", regex=False)
)
final_df["Variation"] = pd.to_numeric(final_df["Variation"], errors="coerce")

month_map = {
    "Jan": "01", "jan": "01", "Fév": "02", "Fev": "02", "fév": "02",
    "Mar": "03", "mars": "03", "Avr": "04", "avril": "04", "Avril": "04",
    "mai": "05", "juin": "06", "juillet": "07", "août": "08", "aout": "08",
    "Sept": "09", "sept": "09", "Oct": "10", "oct": "10", "Nov": "11", "nov": "11",
    "Dec": "12", "Déc": "12", "dec": "12"
}

def parse_week_start(semaine):
    parts = semaine.split("_")
    year = parts[-1]
    start_day = parts[0]

    if parts[1] not in ["au"]:
        month = parts[1]
    else:
        month = parts[-2]

    month_num = month_map.get(month)
    if month_num is None:
        return pd.NaT

    return pd.to_datetime(f"{year}-{month_num}-{start_day}", errors="coerce")

final_df["week_start"] = final_df["Semaine"].apply(parse_week_start)
final_df["week_end"] = final_df["week_start"] + pd.Timedelta(days=6)

final_df = final_df[[
    "week_start", "week_end", "Produit", "Ville", "Prix", "Variation", "Semaine", "Source_File"
]]

output_path = folder / "benin_inflation_clean.csv"
final_df.to_csv(output_path, index=False, encoding="utf-8-sig")
print("Saved:", output_path)

# ==============================================================================
# SECURE DATABASE CONNECTION MANAGEMENT (No hardcoded credentials!)
# ==============================================================================
# Connect cleanly using your system environment URL variable string
engine = create_engine(os.getenv("DATABASE_URL"))

with engine.connect() as conn:
    conn.execute(text("SET FOREIGN_KEY_CHECKS = 0"))
    conn.execute(text("TRUNCATE TABLE food_prices"))
    conn.execute(text("TRUNCATE TABLE products"))
    conn.execute(text("TRUNCATE TABLE cities"))
    conn.execute(text("SET FOREIGN_KEY_CHECKS = 1"))
    conn.commit()
print("Old SQL data cleared safely!")

cities_df = final_df[["Ville"]].drop_duplicates().rename(columns={"Ville": "city_name"})
products_df = final_df[["Produit"]].drop_duplicates().rename(columns={"Produit": "product_name"})

cities_df.to_sql("cities", con=engine, if_exists="append", index=False)
products_df.to_sql("products", con=engine, if_exists="append", index=False)
print("Dimensions seeded successfully!")

cities_sql = pd.read_sql("SELECT * FROM cities", con=engine)
products_sql = pd.read_sql("SELECT * FROM products", con=engine)

final_df = final_df.merge(cities_sql, left_on="Ville", right_on="city_name", how="left")
final_df = final_df.merge(products_sql, left_on="Produit", right_on="product_name", how="left")

food_prices_df = final_df[[
    "week_start", "week_end", "Semaine", "product_id", "city_id", "Prix", "Variation", "Source_File"
]].rename(columns={
    "Semaine": "week_label", "Prix": "price", "Variation": "variation", "Source_File": "source_file"
})

# Production-grade adjustments to fit your optimized VARCHAR / DECIMAL SQL rules
food_prices_df["source_file"] = food_prices_df["source_file"].astype(str).str[:1024]
food_prices_df = food_prices_df.where(pd.notnull(food_prices_df), None)

food_prices_df.to_sql("food_prices", con=engine, if_exists="append", index=False)
print("Food prices imported successfully with DECIMAL types!")
