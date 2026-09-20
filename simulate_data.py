import json
import os
import random
from datetime import datetime, timedelta

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(DATA_DIR, exist_ok=True)

produits = [
    {"name": "Ordinateur Portable XPS 13", "reference": "REF-001", "quantity": 12, "price": 1299.99, "category": "Electronique"},
    {"name": "Souris Sans Fil Logitech", "reference": "REF-002", "quantity": 45, "price": 29.99, "category": "Electronique"},
    {"name": "Clavier Mécanique Corsair", "reference": "REF-003", "quantity": 3, "price": 149.50, "category": "Electronique"},
    {"name": "Chaise de Bureau Ergonomique", "reference": "REF-004", "quantity": 8, "price": 250.00, "category": "Mobilier"},
    {"name": "Bureau Assis-Debout", "reference": "REF-005", "quantity": 4, "price": 450.00, "category": "Mobilier"},
    {"name": "Câble HDMI 2m", "reference": "REF-006", "quantity": 120, "price": 9.99, "category": "Accessoires"},
    {"name": "Casque Audio Sony WH-1000XM4", "reference": "REF-007", "quantity": 15, "price": 349.99, "category": "Audio"},
    {"name": "Disque Dur Externe 2To", "reference": "REF-008", "quantity": 2, "price": 89.99, "category": "Stockage"}
]

with open(os.path.join(DATA_DIR, "produits.json"), "w", encoding="utf-8") as f:
    json.dump(produits, f, ensure_ascii=False, indent=2)

ventes = []
for i in range(25):
    p = random.choice(produits)
    qty = random.randint(1, 3)
    date_vente = (datetime.now() - timedelta(days=random.randint(0, 30))).strftime("%Y-%m-%d %H:%M:%S")
    ventes.append({
        "date": date_vente,
        "product_name": p["name"],
        "reference": p["reference"],
        "quantity": qty,
        "unit_price": p["price"],
        "total": round(p["price"] * qty, 2),
        "category": p["category"]
    })

with open(os.path.join(DATA_DIR, "ventes.json"), "w", encoding="utf-8") as f:
    json.dump(ventes, f, ensure_ascii=False, indent=2)

fournisseurs = [
    {"name": "Tech Data", "contact": "Jean Dupont", "email": "jean@techdata.fr", "phone": "0123456789", "products": ["Electronique", "Audio"]},
    {"name": "Office Supplies Corp", "contact": "Marie Martin", "email": "marie@officesupplies.com", "phone": "0987654321", "products": ["Mobilier", "Accessoires"]}
]

with open(os.path.join(DATA_DIR, "fournisseurs.json"), "w", encoding="utf-8") as f:
    json.dump(fournisseurs, f, ensure_ascii=False, indent=2)

print("Simulated data created successfully.")
