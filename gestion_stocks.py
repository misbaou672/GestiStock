"""
Application de Gestion des Stocks — GestiStock
Mini-projet Python 2026
Auteur : Misbaou DIALLO
Extensions : Excel import/export, alertes stock faible, gestion fournisseurs
"""

# Les bibliothèques dont j'ai besoin pour faire tourner l'appli
import json  # Pour sauvegarder et relire les données en JSON
import csv   # Pour l'export CSV
import os    # Pour gérer les chemins et dossiers
from datetime import datetime, date  # Pour dater les ventes
import tkinter as tk  # La base de l'interface graphique
from tkinter import ttk, messagebox, filedialog  # Les widgets en plus
from collections import defaultdict  # Pratique pour les stats
from PIL import Image, ImageTk  # Icônes vectorielles épurées

# J'essaie d'importer openpyxl pour Excel.
# Si c'est pas installé chez l'utilisateur, tant pis : les boutons Excel seront juste désactivés.
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    EXCEL_OK = True  # Nickel, Excel est dispo
except ImportError:
    EXCEL_OK = False  # Pas grave, le reste de l'appli tourne quand même


# ─────────────────────────────────────────────
#  COUCHE DONNÉES : tout ce qui touche au stockage
# ─────────────────────────────────────────────

# Les chemins vers mes fichiers de données
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")  # Dossier data/
PRODUCTS_FILE   = os.path.join(DATA_DIR, "produits.json")    # Mes produits
SALES_FILE      = os.path.join(DATA_DIR, "ventes.json")       # L'historique des ventes
SUPPLIERS_FILE  = os.path.join(DATA_DIR, "fournisseurs.json") # Mes fournisseurs
LOW_STOCK_THRESHOLD = 5  # En dessous de 5 pièces, je considère que c'est critique


def ensure_data_dir():
    """Je crée le dossier data/ s'il existe pas déjà"""
    os.makedirs(DATA_DIR, exist_ok=True)


def load_json(path):
    """Je charge un fichier JSON. S'il existe pas, je renvoie une liste vide."""
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return []  # Pas de fichier = pas de données, c'est tout


def save_json(path, data):
    """Je sauvegarde les données en JSON, bien indenté pour que ce soit lisible"""
    ensure_data_dir()  # On s'assure que le dossier existe
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)  # Indenté = plus joli à lire


def export_csv(path, data, fieldnames):
    """J'exporte en CSV, ça passe partout (Excel, Google Sheets, LibreOffice...)"""
    ensure_data_dir()
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()  # La ligne d'en-têtes
        writer.writerows(data)  # Et toutes les lignes


# ─────────────────────────────────────────────
#  FONCTIONS EXCEL : pour les exports/imports Excel un peu stylés
# ─────────────────────────────────────────────

def _xl_header_style(ws, row, cols, fill_hex="4C3F9F"):
    """Je mets un joli style sur les en-têtes de tableau Excel (fond coloré, texte blanc, gras)"""
    fill   = PatternFill("solid", fgColor=fill_hex)  # Fond de couleur
    font   = Font(bold=True, color="FFFFFF", size=11)  # Texte blanc et gras
    border = Border(bottom=Side(style="medium", color="FFFFFF"))  # Petite ligne blanche en dessous
    
    # J'applique le style à chaque colonne
    for col, title in enumerate(cols, 1):
        cell = ws.cell(row=row, column=col, value=title)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    ws.row_dimensions[row].height = 22  # Je laisse un peu d'air dans les en-têtes


def _xl_autofit(ws):
    """J'ajuste automatiquement la largeur des colonnes Excel selon leur contenu"""
    for col in ws.columns:
        # Je cherche la plus grande longueur dans la colonne
        max_len = max((len(str(c.value or "")) for c in col), default=8)
        # Largeur = max_len + 4, mais je plafonne à 40 pour pas avoir des colonnes énormes
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 40)


def export_excel(products, sales, suppliers, path):
    """Je crée un fichier Excel avec 4 feuilles : Produits, Ventes, Fournisseurs et Stock Faible"""
    wb = openpyxl.Workbook()  # Nouveau classeur vide

    # ── Feuille 1 : Produits ──
    ws_p = wb.active
    ws_p.title = "Produits"
    cols_p = ["Nom", "Reference", "Quantite", "Prix (EUR)", "Categorie"]
    _xl_header_style(ws_p, 1, cols_p)
    
    # Style spécial pour les produits en stock faible
    low_fill  = PatternFill("solid", fgColor="FFF3CD")  # Fond jaune pâle
    warn_font = Font(color="856404")  # Texte orangé
    
    # Je trie les produits par catégorie pour que ce soit plus clair
    for p in sorted(products, key=lambda x: x["category"]):
        row = ws_p.max_row + 1
        ws_p.append([p["name"], p["reference"], p["quantity"], p["price"], p["category"]])
        
        # Si le stock est faible, je surligne la ligne en jaune
        if p["quantity"] <= LOW_STOCK_THRESHOLD:
            for col in range(1, 6):
                ws_p.cell(row=row, column=col).fill = low_fill
                ws_p.cell(row=row, column=col).font = warn_font
    _xl_autofit(ws_p)

    # ── Feuille 2 : Ventes ──
    ws_s = wb.create_sheet("Ventes")
    cols_s = ["Date", "Produit", "Reference", "Quantite", "Prix unit. (EUR)", "Total (EUR)", "Categorie"]
    _xl_header_style(ws_s, 1, cols_s, fill_hex="1A7A4A")  # Vert pour les ventes
    
    # Je mets les ventes les plus récentes en haut
    for s in sorted(sales, key=lambda x: x["date"], reverse=True):
        ws_s.append([s["date"], s["product_name"], s["reference"],
                     s["quantity"], s["unit_price"], s["total"], s["category"]])
    
    # Une petite ligne de total à la fin si y'a des ventes
    if sales:
        total_row = ws_s.max_row + 1
        ws_s.cell(total_row, 1, "TOTAL").font = Font(bold=True)
        ws_s.cell(total_row, 6, sum(s["total"] for s in sales)).font = Font(bold=True, color="1A7A4A")
    _xl_autofit(ws_s)

    # ── Feuille 3 : Fournisseurs ──
    ws_f = wb.create_sheet("Fournisseurs")
    cols_f = ["Nom", "Contact", "Email", "Telephone", "Produits fournis"]
    _xl_header_style(ws_f, 1, cols_f, fill_hex="8B3A8B")  # Violet pour les fournisseurs
    
    for f in suppliers:
        ws_f.append([f["name"], f["contact"], f.get("email", ""),
                     f.get("phone", ""), ", ".join(f.get("products", []))])
    _xl_autofit(ws_f)

    # ── Feuille 4 : Alertes stock faible ──
    ws_a = wb.create_sheet("Stock Faible")
    cols_a = ["Nom", "Reference", "Quantite", "Categorie", "Statut"]
    _xl_header_style(ws_a, 1, cols_a, fill_hex="C0392B")  # Rouge pour les alertes
    red_fill = PatternFill("solid", fgColor="FDECEA")  # Fond rouge clair
    
    # Je ne mets que les produits en stock faible, triés du plus critique au moins critique
    for p in sorted(products, key=lambda x: x["quantity"]):
        if p["quantity"] <= LOW_STOCK_THRESHOLD:
            status = "CRITIQUE" if p["quantity"] == 0 else "FAIBLE"
            row = ws_a.max_row + 1
            ws_a.append([p["name"], p["reference"], p["quantity"], p["category"], status])
            # Tout en rouge pour bien voir
            for col in range(1, 6):
                ws_a.cell(row=row, column=col).fill = red_fill

    wb.save(path)  # Et voilà, on sauvegarde
    return path


def import_excel_products(path):
    """J'importe les produits depuis un fichier Excel (feuille 'Produits')"""
    wb = openpyxl.load_workbook(path, read_only=True)
    
    # Je cherche une feuille qui contient "produit" dans son nom
    sheet_name = None
    for name in wb.sheetnames:
        if "produit" in name.lower():
            sheet_name = name
            break
    
    if not sheet_name:
        raise ValueError("Aucune feuille 'Produits' trouvee dans ce fichier.")
    
    ws = wb[sheet_name]
    products = []
    headers = None
    
    # Je parcours toutes les lignes une par une
    for row in ws.iter_rows(values_only=True):
        if headers is None:
            # La première ligne, c'est les en-têtes
            headers = [str(h).strip() if h else "" for h in row]
            continue
        if not any(row):
            continue  # Ligne vide, je saute
            
        d = dict(zip(headers, row))  # J'associe en-tête → valeur
        
        # Je fais la correspondance français → anglais (avec et sans accents)
        mapping = {
            "Nom": "name", "Reference": "reference", "Référence": "reference",
            "Quantite": "quantity", "Quantité": "quantity",
            "Prix (EUR)": "price", "Prix (€)": "price",
            "Categorie": "category", "Catégorie": "category",
        }
        
        product = {}
        for fr_key, en_key in mapping.items():
            if fr_key in d and d[fr_key] is not None:
                product[en_key] = d[fr_key]
        
        # Il me faut au minimum un nom et une référence
        if "name" in product and "reference" in product:
            # Valeurs par défaut si jamais il manque des champs
            product.setdefault("quantity", 0)
            product.setdefault("price", 0.0)
            product.setdefault("category", "Autre")
            
            try:
                # Je convertis au bon type
                product["quantity"] = int(product["quantity"])
                product["price"]    = float(product["price"])
                products.append(product)
            except (ValueError, TypeError):
                pass  # Ligne bancale, je la zappe
    return products


# ─────────────────────────────────────────────
#  LOGIQUE MÉTIER : le cerveau de l'application
# ─────────────────────────────────────────────

class StockManager:
    """Le gestionnaire principal : il gère les produits, les ventes et les fournisseurs"""
    
    def __init__(self):
        """Au démarrage, je charge tout depuis les fichiers JSON"""
        self.products  = load_json(PRODUCTS_FILE)   # Les produits
        self.sales     = load_json(SALES_FILE)      # L'historique des ventes
        self.suppliers = load_json(SUPPLIERS_FILE)  # Les fournisseurs

    # ── GESTION DES PRODUITS ──────────────────────────────

    def _find_product(self, ref):
        """Je cherche un produit par sa référence (usage interne)"""
        for p in self.products:
            if p["reference"] == ref:
                return p
        return None  # Pas trouvé

    def add_product(self, name, reference, quantity, price, category):
        """J'ajoute un nouveau produit au stock"""
        # Je vérifie d'abord les champs obligatoires
        if not str(name).strip():
            raise ValueError("Le nom du produit est obligatoire.")
        if not str(reference).strip():
            raise ValueError("La reference est obligatoire.")
        if self._find_product(str(reference).strip()):
            raise ValueError(f"Reference '{reference}' deja utilisee.")
        
        # Je crée le produit en nettoyant bien les données
        product = {
            "name":      str(name).strip(),
            "reference": str(reference).strip(),
            "quantity":  int(quantity),
            "price":     float(price),
            "category":  str(category).strip() or "Autre",  # "Autre" si pas de catégorie
        }
        self.products.append(product)
        save_json(PRODUCTS_FILE, self.products)  # Je sauvegarde tout de suite
        return product

    def update_product(self, reference, **kwargs):
        """Je modifie un produit existant"""
        p = self._find_product(reference)
        if not p:
            raise ValueError(f"Produit '{reference}' introuvable.")
        
        # Je ne mets à jour que les champs autorisés
        for k, v in kwargs.items():
            if k in {"name", "quantity", "price", "category"}:
                p[k] = v
        save_json(PRODUCTS_FILE, self.products)
        return p

    def delete_product(self, reference):
        """Je supprime un produit du stock"""
        p = self._find_product(reference)
        if not p:
            raise ValueError(f"Produit '{reference}' introuvable.")
        self.products.remove(p)
        save_json(PRODUCTS_FILE, self.products)

    def get_products(self, sort_by="category"):
        """Je renvoie tous les produits, triés selon le critère demandé"""
        reverse = (sort_by == "quantity")  # Pour la quantité, je trie en décroissant
        default_value = 0 if reverse else ""  # Pour éviter les bugs de comparaison
        return sorted(self.products,
                      key=lambda p: p.get(sort_by, default_value),
                      reverse=reverse)

    def low_stock_products(self):
        """Je renvoie la liste des produits en stock faible (≤ 5 unités)"""
        return [p for p in self.products if p["quantity"] <= LOW_STOCK_THRESHOLD]

    def import_products_from_excel(self, path):
        """
        J'importe des produits depuis un fichier Excel et je mets à jour le stock.
        
        Args:
            path (str): Le chemin vers le fichier Excel
            
        Returns:
            tuple: (added, skipped) - combien ajoutés, combien ignorés
            
        À noter :
            - Les produits avec une référence déjà existante sont ignorés (pas d'écrasement)
            - Les nouveaux produits sont ajoutés à la fin
            - Le fichier produits.json est sauvegardé automatiquement
        """
        imported = import_excel_products(path)  # J'appelle la fonction d'import
        added, skipped = 0, 0  # Les compteurs
        
        for p in imported:
            if self._find_product(p["reference"]):
                skipped += 1  # Déjà existant → j'ignore
            else:
                self.products.append(p)  # Nouveau → j'ajoute
                added += 1
                
        save_json(PRODUCTS_FILE, self.products)  # Je sauvegarde
        return added, skipped

    # ── GESTION DES VENTES ─────────────────────────────────
    # Cette partie gère tout ce qui touche aux ventes :
    # - L'enregistrement d'une vente (avec mise à jour auto du stock)
    # - L'historique des ventes avec filtres par date
    # - Les calculs de totaux et de chiffre d'affaires

    def record_sale(self, reference, quantity):
        """
        J'enregistre une vente et je mets à jour le stock tout seul.
        
        C'est la fonction la plus importante : elle fait deux choses d'un coup :
        1. Elle déduit la quantité vendue du stock
        2. Elle crée un enregistrement dans l'historique des ventes
        
        Args:
            reference (str): La référence du produit vendu
            quantity (int): La quantité vendue (positive)
            
        Returns:
            dict: L'enregistrement de vente créé
            
        Erreurs possibles :
            ValueError: Si le produit existe pas, si la quantité est nulle
                        ou si y'a pas assez de stock
                        
        Exemple :
            >>> sale = manager.record_sale("ELEC-001", 2)
            >>> print(sale["total"])  # Affiche le total de la vente
        """
        # Étape 1 : je vérifie que le produit existe
        p = self._find_product(reference)
        if not p:
            raise ValueError(f"Produit '{reference}' introuvable.")
        
        # Étape 2 : je valide la quantité
        qty = int(quantity)
        if qty <= 0:
            raise ValueError("La quantite doit etre positive.")
        if p["quantity"] < qty:
            raise ValueError(f"Stock insuffisant. Disponible : {p['quantity']}.")
        
        # Étape 3 : je déduis du stock
        p["quantity"] -= qty
        
        # Étape 4 : je crée l'enregistrement de vente
        sale = {
            "reference":    reference,
            "product_name": p["name"],
            "quantity":     qty,
            "unit_price":   p["price"],  # Le prix au moment de la vente
            "total":        round(qty * p["price"], 2),
            "category":     p["category"],
            "date":         date.today().isoformat(),  # Date du jour au format ISO
        }
        self.sales.append(sale)
        
        # Étape 5 : je sauvegarde les deux fichiers (produits et ventes)
        save_json(PRODUCTS_FILE, self.products)
        save_json(SALES_FILE, self.sales)
        return sale

    def get_sales(self, from_date=None, to_date=None):
        """
        Je renvoie les ventes, avec un filtre optionnel par période.
        
        Utile pour :
        - Afficher l'historique complet
        - Générer des rapports sur une période précise
        - Calculer des stats temporelles
        
        Args:
            from_date (str, optional): Date de début (YYYY-MM-DD). None = pas de limite
            to_date (str, optional): Date de fin (YYYY-MM-DD). None = pas de limite
            
        Returns:
            list: La liste des ventes filtrées
            
        Exemples :
            >>> # Toutes les ventes
            >>> all_sales = manager.get_sales()
            >>> # Les ventes de janvier 2026
            >>> jan_sales = manager.get_sales("2026-01-01", "2026-01-31")
        """
        result = self.sales  # Je pars de toutes les ventes
        
        if from_date:
            result = [s for s in result if s["date"] >= from_date]
        if to_date:
            result = [s for s in result if s["date"] <= to_date]
        return result

    # ── RAPPORTS ET STATISTIQUES ───────────────────────────────
    # Ces fonctions transforment les données brutes en infos utiles :
    # - Top des produits vendus (pour savoir quoi recommander)
    # - CA par catégorie (pour voir ce qui rapporte le plus)
    # - Niveaux de stock (pour éviter les ruptures)
    # - Alertes stock faible (pour réapprovisionner à temps)

    def report_top_products(self, n=5, from_date=None, to_date=None):
        """
        Je calcule les n produits les plus vendus, avec leur chiffre d'affaires.
        
        Ça permet de voir :
        - Quels produits partent le plus (en quantité)
        - Quels produits rapportent le plus (en CA)
        
        Args:
            n (int): Le nombre de produits à retourner (5 par défaut)
            from_date (str, optional): Date de début
            to_date (str, optional): Date de fin
            
        Returns:
            list: Une liste de tuples (reference, {"qty": int, "revenue": float, "name": str})
                  triée par quantité décroissante
        """
        # Un defaultdict pour accumuler les stats par produit
        totals = defaultdict(lambda: {"qty": 0, "revenue": 0.0, "name": ""})
        
        for s in self.get_sales(from_date, to_date):
            ref = s["reference"]
            totals[ref]["qty"]     += s["quantity"]
            totals[ref]["revenue"] += s["total"]
            totals[ref]["name"]     = s["product_name"]
        
        sorted_products = sorted(totals.items(), key=lambda x: x[1]["qty"], reverse=True)
        return sorted_products[:n]

    def report_revenue_by_category(self, from_date=None, to_date=None):
        """
        Je calcule le chiffre d'affaires par catégorie de produits.
        
        Ça sert à :
        - Voir quelles catégories rapportent le plus
        - Décider quoi développer en priorité
        - Analyser la répartition du CA
        
        Returns:
            dict: Un dictionnaire {categorie: chiffre_affaires}
        """
        cats = defaultdict(float)
        for s in self.get_sales(from_date, to_date):
            cats[s["category"]] += s["total"]
        return dict(cats)

    def report_total_revenue(self, from_date=None, to_date=None):
        """
        Je calcule le chiffre d'affaires total sur une période.
        
        C'est un peu l'indicateur clé de performance. Ça sert à :
        - Suivre les objectifs de vente
        - Comparer les périodes entre elles
        - Calculer la rentabilité globale
        
        Returns:
            float: Le CA total en euros
        """
        return sum(s["total"] for s in self.get_sales(from_date, to_date))

    def report_stock_levels(self):
        """
        Je repère les produits avec les niveaux de stock les plus extrêmes.
        
        Ça aide à la gestion des stocks en montrant :
        - Les 5 produits avec le moins de stock (risque de rupture)
        - Les 5 produits avec le plus de stock (risque de surstock)
        
        Returns:
            tuple: (lowest_stock, highest_stock)
        """
        if not self.products:
            return [], []  # Pas de produits = listes vides
            
        sorted_products = sorted(self.products, key=lambda x: x["quantity"])
        lowest_stock = sorted_products[:5]              # Les 5 plus bas
        highest_stock = sorted_products[-5:][::-1]      # Les 5 plus hauts, dans l'ordre décroissant
        return lowest_stock, highest_stock

    # ── GESTION DES FOURNISSEURS (Extension 3) ───────────────
    # Les fournisseurs, c'est important pour les achats :
    # - Ajouter de nouveaux fournisseurs avec leurs coordonnées
    # - Leur associer des produits
    # - Garder leurs contacts pour les commandes
    # - Suivre les relations commerciales

    def _find_supplier(self, name):
        """
        Je cherche un fournisseur par son nom (insensible à la casse).
        
        Args:
            name (str): Le nom du fournisseur
            
        Returns:
            dict or None: Le fournisseur trouvé, ou None sinon
            
        À noter : "Apple" == "apple" == "APPLE" pour cette recherche
        """
        for s in self.suppliers:
            if s["name"].lower() == name.lower():
                return s
        return None

    def add_supplier(self, name, contact, email="", phone="", products=None):
        """
        J'ajoute un nouveau fournisseur.
        
        Args:
            name (str): Le nom du fournisseur (obligatoire)
            contact (str): Le nom du contact (obligatoire)
            email (str, optional): L'email
            phone (str, optional): Le téléphone
            products (list, optional): La liste des références fournies
            
        Returns:
            dict: Le fournisseur créé
        """
        # Je vérifie les champs obligatoires
        if not str(name).strip():
            raise ValueError("Le nom du fournisseur est obligatoire.")
        if not str(contact).strip():
            raise ValueError("Le nom du contact est obligatoire.")
        if self._find_supplier(name):
            raise ValueError(f"Fournisseur '{name}' existe deja.")
        
        # Je crée l'objet fournisseur
        sup = {
            "name":     str(name).strip(),
            "contact":  str(contact).strip(),
            "email":    str(email).strip(),
            "phone":    str(phone).strip(),
            "products": [p.strip() for p in (products or []) if p.strip()],
        }
        self.suppliers.append(sup)
        save_json(SUPPLIERS_FILE, self.suppliers)
        return sup

    def update_supplier(self, original_name, **kwargs):
        """
        Je mets à jour un fournisseur existant.
        
        Args:
            original_name (str): Le nom actuel du fournisseur
            **kwargs: Les champs à modifier (name, contact, email, phone, products)
            
        Returns:
            dict: Le fournisseur mis à jour
        """
        s = self._find_supplier(original_name)
        if not s:
            raise ValueError(f"Fournisseur '{original_name}' introuvable.")
            
        for k, v in kwargs.items():
            if k in {"name", "contact", "email", "phone", "products"}:
                s[k] = v
                
        save_json(SUPPLIERS_FILE, self.suppliers)
        return s

    def delete_supplier(self, name):
        """
        Je supprime un fournisseur.
        
        ⚠️  Attention : c'est définitif !
        Les produits associés ne sont PAS supprimés, juste la relation.
        """
        s = self._find_supplier(name)
        if not s:
            raise ValueError(f"Fournisseur '{name}' introuvable.")
            
        self.suppliers.remove(s)
        save_json(SUPPLIERS_FILE, self.suppliers)

    def get_suppliers(self):
        """Je renvoie tous les fournisseurs triés par ordre alphabétique"""
        return sorted(self.suppliers, key=lambda s: s["name"])

    # ── FONCTIONS D'EXPORT ────────────────────────────────
    # Pour exporter les données dans différents formats :
    # - CSV : compatible partout
    # - Excel : format natif avec mise en forme

    def export_products_csv(self):
        """J'exporte la liste des produits en CSV"""
        path = os.path.join(DATA_DIR, "produits.csv")
        columns = ["name", "reference", "quantity", "price", "category"]
        export_csv(path, self.products, columns)
        return path

    def export_sales_csv(self):
        """J'exporte l'historique des ventes en CSV"""
        path = os.path.join(DATA_DIR, "ventes.csv")
        columns = [
            "date", "reference", "product_name",
            "quantity", "unit_price", "total", "category"
        ]
        export_csv(path, self.sales, columns)
        return path

    def export_suppliers_csv(self):
        """J'exporte la liste des fournisseurs en CSV"""
        path = os.path.join(DATA_DIR, "fournisseurs.csv")
        
        # Je transforme la liste des produits en chaîne pour le CSV
        rows = []
        for s in self.suppliers:
            row = {**s, "products": ", ".join(s.get("products", []))}
            rows.append(row)
        
        columns = ["name", "contact", "email", "phone", "products"]
        export_csv(path, rows, columns)
        return path

    def export_all_excel(self):
        """
        J'exporte TOUT dans un seul fichier Excel multi-feuilles.
        
        C'est l'export le plus complet :
        - Feuille 1 : Produits (avec les stocks faibles mis en évidence)
        - Feuille 2 : Ventes (triées par date, avec total)
        - Feuille 3 : Fournisseurs (avec coordonnées)
        - Feuille 4 : Alertes stock faible
        """
        path = os.path.join(DATA_DIR, "gestion_stocks.xlsx")
        export_excel(self.products, self.sales, self.suppliers, path)
        return path


# ─────────────────────────────────────────────
#  INTERFACE GRAPHIQUE : le style et les couleurs
# ─────────────────────────────────────────────

# Palette de couleurs claire, sobre et professionnelle (Style SaaS / Enterprise Light Mode)
COLORS = {
    "bg":      "#f8fafc",  # Fond principal lumineux (Slate 50)
    "panel":   "#ffffff",  # Barre latérale blanc pur
    "card":    "#ffffff",  # Cartes blanches avec bordure
    "accent":  "#2563eb",  # Bleu professionnel (Royal Blue 600)
    "accent2": "#059669",  # Vert émeraude
    "accent3": "#4f46e5",  # Indigo
    "danger":  "#dc2626",  # Rouge alerte
    "warning": "#d97706",  # Ambre
    "text":    "#0f172a",  # Texte principal foncé (Slate 900)
    "subtext": "#64748b",  # Texte secondaire (Slate 500)
    "border":  "#e2e8f0",  # Bordures grises très claires (Slate 200)
    "entry":   "#f1f5f9",  # Champs de saisie (Slate 100)
}

# Mes polices d'écriture pour un rendu moderne
FONT_TITLE = ("Segoe UI", 22, "bold")  # Titres principaux
FONT_HEAD  = ("Segoe UI", 12, "bold")  # Sous-titres
FONT_LABEL = ("Segoe UI", 10)          # Textes normaux
FONT_SMALL = ("Segoe UI", 9)           # Petits textes
FONT_MONO  = ("Consolas", 10)          # Texte monospace


# ─────────────────────────────────────────────
#  COMPOSANTS RÉUTILISABLES : mes petites briques d'interface
# ─────────────────────────────────────────────

def make_card(parent, title="", **kwargs):
    """Je crée une carte stylisée avec une bordure et un titre optionnel"""
    outer = tk.Frame(parent, bg=COLORS["card"],
                     highlightbackground=COLORS["border"],
                     highlightthickness=1, **kwargs)
    if title:
        tk.Label(outer, text=title, font=FONT_HEAD,
                 bg=COLORS["card"], fg=COLORS["text"]).pack(anchor="w", padx=16, pady=(14, 6))
    return outer


def make_entry(parent, label, row, default=""):
    """Je crée un champ de saisie avec son label, déjà stylisé"""
    # Le label à gauche
    tk.Label(parent, text=label, font=FONT_LABEL,
             bg=COLORS["card"], fg=COLORS["subtext"]).grid(
        row=row, column=0, sticky="w", padx=(16, 8), pady=4)
    
    # Le champ de saisie à droite
    var = tk.StringVar(value=default)
    tk.Entry(parent, textvariable=var, font=FONT_LABEL,
             bg=COLORS["entry"], fg=COLORS["text"],
             insertbackground=COLORS["text"],
             relief="solid", bd=1).grid(
        row=row, column=1, sticky="ew", padx=(0, 16), pady=4, ipady=4)
    return var


def make_button(parent, text, command, color=None, **kwargs):
    """Je crée un bouton moderne avec un style cohérent"""
    color = color or COLORS["accent"]
    # Si le fond du bouton est clair, on utilise du texte sombre (#0f172a)
    fg_color = kwargs.pop("fg", None)
    if not fg_color:
        if color in (COLORS["card"], COLORS["entry"], COLORS["panel"], "#ffffff", "#f1f5f9", "#f8fafc", "#e2e8f0"):
            fg_color = COLORS["text"]
        else:
            fg_color = "#ffffff"
    return tk.Button(parent, text=text, command=command,
                     font=("Segoe UI", 10, "bold"),
                     bg=color, fg=fg_color, relief="flat", bd=0,
                     activebackground=COLORS["bg"], activeforeground=color,
                     cursor="hand2", padx=14, pady=7, **kwargs)


def style_tree(tree):
    """J'applique mon style lumineux aux tableaux Treeview"""
    s = ttk.Style()
    s.theme_use("clam")
    
    # Le style des cellules
    s.configure("Custom.Treeview",
                background="#ffffff", foreground="#0f172a",
                fieldbackground="#ffffff", rowheight=32,
                font=FONT_LABEL, borderwidth=0)
    
    # Le style des en-têtes
    s.configure("Custom.Treeview.Heading",
                background="#f1f5f9", foreground="#0f172a",
                font=("Segoe UI", 10, "bold"), relief="flat")
    
    # Quand on sélectionne une ligne
    s.map("Custom.Treeview",
          background=[("selected", COLORS["accent"])],
          foreground=[("selected", "white")])
    
    tree.configure(style="Custom.Treeview")


# ─────────────────────────────────────────────
#  FENÊTRE PRINCIPALE : le cœur de l'application
# ─────────────────────────────────────────────

class App(tk.Tk):
    """La fenêtre principale de GestiStock"""
    
    def __init__(self):
        super().__init__()
        self.title("GestiStock — Gestion des Stocks | Misbaou DIALLO")
        self.configure(bg=COLORS["bg"])
        self.geometry("1300x800")
        self.minsize(980, 640)  # En dessous, ça devient tout cassé
        
        self.manager = StockManager()  # Mon gestionnaire de données
        self._build_ui()  # Je construis l'interface
        self.after(300, self._check_low_stock_on_start)  # Alerte stock après 300ms

    def _build_ui(self):
        """Je construis toute l'interface"""
        # La barre latérale à gauche
        self.sidebar = tk.Frame(self, bg=COLORS["panel"], width=215)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)  # Largeur fixe

        # La zone principale à droite
        self.content = tk.Frame(self, bg=COLORS["bg"])
        self.content.pack(side="left", fill="both", expand=True)

        self._build_sidebar()  # Le menu de navigation

        # Je crée toutes les pages, mais une seule s'affiche à la fois
        self.frames = {}
        for F in (ProductsFrame, SalesFrame, SuppliersFrame, ReportsFrame):
            frame = F(self.content, self)
            self.frames[F.__name__] = frame
            frame.place(relx=0, rely=0, relwidth=1, relheight=1)

        self._show_frame("ProductsFrame")  # Page par défaut au démarrage

    def _build_sidebar(self):
        """Je construis la barre latérale avec le logo et la navigation"""
        # Charger les icônes PNG épurées
        icon_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")
        self.icons = {}
        for key in ["products", "sales", "suppliers", "reports", "csv", "excel"]:
            path = os.path.join(icon_dir, f"{key}.png")
            if os.path.exists(path):
                self.icons[key] = ImageTk.PhotoImage(Image.open(path))

        # Le logo et le nom de l'app
        tk.Label(self.sidebar, text="GESTISTOCK", font=("Segoe UI", 15, "bold"),
                 bg=COLORS["panel"], fg=COLORS["text"]).pack(pady=(24, 2))
        tk.Label(self.sidebar, text="GESTION DE STOCK & SUPPLIERS", font=("Segoe UI", 8),
                 bg=COLORS["panel"], fg=COLORS["subtext"]).pack(pady=(0, 16))

        # Une petite ligne de séparation
        ttk.Separator(self.sidebar, orient="horizontal").pack(fill="x", padx=16)

        # Mon menu de navigation avec icônes
        nav_items = [
            ("  Catalogue Produits",    "ProductsFrame", "products"),
            ("  Historique Ventes",     "SalesFrame", "sales"),
            ("  Annuaire Fournisseurs", "SuppliersFrame", "suppliers"),
            ("  Rapports & Bilans",     "ReportsFrame", "reports"),
        ]
        self.nav_btns = {}
        
        for label, fname, ikey in nav_items:
            img = self.icons.get(ikey)
            btn = tk.Button(
                self.sidebar, text=label, image=img, compound="left",
                font=("Segoe UI", 10, "bold"),
                anchor="w", padx=16, pady=10, bd=0, relief="flat",
                bg=COLORS["panel"], fg=COLORS["subtext"],
                activebackground=COLORS["accent"], activeforeground="white",
                cursor="hand2",
                command=lambda f=fname: self._show_frame(f),
            )
            btn.pack(fill="x", pady=2, padx=8)
            self.nav_btns[fname] = btn

        tk.Frame(self.sidebar, bg=COLORS["panel"]).pack(expand=True)
        ttk.Separator(self.sidebar, orient="horizontal").pack(fill="x", padx=16)

        btn_area = tk.Frame(self.sidebar, bg=COLORS["panel"])
        btn_area.pack(fill="x", pady=8, padx=8)

        make_button(btn_area, "  Export CSV (.csv)", self._export_csv,
                    image=self.icons.get("csv"), compound="left",
                    color=COLORS["card"]).pack(fill="x", pady=3)

        if EXCEL_OK:
            make_button(btn_area, "  Export Excel (.xlsx)", self._export_excel,
                        image=self.icons.get("excel"), compound="left",
                        color=COLORS["card"]).pack(fill="x", pady=3)
            make_button(btn_area, "  Import Excel (.xlsx)", self._import_excel,
                        image=self.icons.get("excel"), compound="left",
                        color=COLORS["card"]).pack(fill="x", pady=3)

    def _show_frame(self, name):
        for n, btn in self.nav_btns.items():
            is_active = (n == name)
            btn.configure(
                bg=COLORS["accent"] if is_active else COLORS["panel"],
                fg="white" if is_active else COLORS["subtext"]
            )
        self.frames[name].tkraise()
        self.frames[name].refresh()

    def _export_csv(self):
        p = self.manager.export_products_csv()
        s = self.manager.export_sales_csv()
        f = self.manager.export_suppliers_csv()
        messagebox.showinfo("Export CSV", f"3 fichiers crees dans data/ :\n"
                            f"- {os.path.basename(p)}\n"
                            f"- {os.path.basename(s)}\n"
                            f"- {os.path.basename(f)}")

    def _export_excel(self):
        path = self.manager.export_all_excel()
        messagebox.showinfo("Export Excel",
                            f"Fichier cree :\n{path}\n\n"
                            "4 feuilles : Produits, Ventes, Fournisseurs, Stock Faible")

    def _import_excel(self):
        path = filedialog.askopenfilename(
            title="Importer des produits depuis Excel",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls"), ("Tous", "*.*")]
        )
        if not path:
            return
        try:
            added, skipped = self.manager.import_products_from_excel(path)
            self.frames["ProductsFrame"].refresh()
            messagebox.showinfo("Import termine",
                                f"{added} produit(s) importe(s).\n"
                                f"{skipped} ignore(s) (reference deja existante).")
        except Exception as e:
            messagebox.showerror("Erreur d'import", str(e))

    def _check_low_stock_on_start(self):
        low = self.manager.low_stock_products()
        if low:
            lines = "\n".join(
                f"  {'[RUPTURE]' if p['quantity']==0 else '[FAIBLE]'} "
                f"{p['name']} — {p['quantity']} restant(s)"
                for p in sorted(low, key=lambda x: x["quantity"])
            )
            messagebox.showwarning(
                "Alerte Stock Faible",
                f"{len(low)} produit(s) sous le seuil d'alerte "
                f"({LOW_STOCK_THRESHOLD} unites) :\n\n{lines}\n\n"
                "Pensez a reapprovisionner !"
            )


# ─────────────────────────────────────────────
#  PAGE PRODUITS : la gestion complète du stock
# ─────────────────────────────────────────────
# C'est la page la plus utilisée :
# - Affichage de tous les produits avec leur stock
# - Recherche en temps réel
# - Tri par catégorie/nom/quantité
# - Ajout, modification, suppression de produits
# - Alertes visuelles pour les stocks faibles

class ProductsFrame(tk.Frame):
    """
    La page principale pour gérer les produits.
    
    Elle contient :
    - Un tableau avec tous les produits
    - Une barre de recherche pour filtrer
    - Un formulaire pour ajouter/modifier
    - Des boutons d'action (supprimer, modifier)
    - Des options de tri et des catégories rapides
    - Une alerte pour les stocks faibles
    """
    
    def __init__(self, parent, app):
        """
        J'initialise la page des produits.
        
        Args:
            parent: Le frame parent
            app: L'instance de l'app principale (pour le manager)
        """
        super().__init__(parent, bg=COLORS["bg"])
        self.app = app  # Référence à l'app
        self._build()

    def _build(self):
        """
        Je construis toute l'interface de la page produits.
        
        Organisation :
        1. En-tête avec le titre et la recherche
        2. Zone principale : tableau à gauche, formulaire à droite
        """
        # ── EN-TÊTE : titre et recherche ──
        hdr = tk.Frame(self, bg=COLORS["bg"])
        hdr.pack(fill="x", padx=24, pady=(20, 10))
        
        tk.Label(hdr, text="Gestion des Produits", font=FONT_TITLE,
                 bg=COLORS["bg"], fg=COLORS["text"]).pack(side="left")

        # La barre de recherche, stylisée
        sf = tk.Frame(hdr, bg=COLORS["card"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        sf.pack(side="right")
        
        tk.Label(sf, text="🔍", bg=COLORS["card"], fg=COLORS["subtext"]).pack(side="left", padx=8)
        
        self.search_var = tk.StringVar()
        # À chaque lettre tapée, je rafraîchis le tableau
        self.search_var.trace_add("write", lambda *_: self.refresh())
        tk.Entry(sf, textvariable=self.search_var, font=FONT_LABEL,
                 bg=COLORS["card"], fg=COLORS["text"],
                 insertbackground=COLORS["text"],
                 relief="flat", bd=0, width=22).pack(side="left", ipady=6, padx=(0, 8))

        # ── ZONE PRINCIPALE : tableau et formulaire ──
        main = tk.Frame(self, bg=COLORS["bg"])
        main.pack(fill="both", expand=True, padx=24, pady=8)

        # ── TABLEAU DES PRODUITS (à gauche) ──
        tbl = make_card(main, "Stock actuel")
        tbl.pack(side="left", fill="both", expand=True, padx=(0, 12))

        cols = ("Nom", "Reference", "Quantite", "Prix (EUR)", "Categorie")
        headers = ("Nom", "Référence", "Quantité", "Prix (€)", "Catégorie")
        self.tree = ttk.Treeview(tbl, columns=cols, show="headings", selectmode="browse")
        
        for c, h in zip(cols, headers):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=110, anchor="center")
        self.tree.column("Nom", width=180, anchor="w")
        
        style_tree(self.tree)  # Mon style sombre
        
        # Tag pour surligner les stocks faibles en orange
        self.tree.tag_configure("low", foreground=COLORS["warning"])

        # La barre de défilement + la sélection
        sb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        sb.pack(side="right", fill="y", pady=8)
        
        # Quand on clique sur une ligne, ça remplit le formulaire
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # ── BARRE D'ACTION : boutons + tri ──
        btn_row = tk.Frame(tbl, bg=COLORS["card"])
        btn_row.pack(fill="x", padx=8, pady=(0, 10))
        
        make_button(btn_row, "Supprimer", self._delete,
                    color=COLORS["danger"]).pack(side="right", padx=4)
        make_button(btn_row, "Modifier", self._update).pack(side="right", padx=4)

        # Les options de tri à gauche
        sort_f = tk.Frame(btn_row, bg=COLORS["card"])
        sort_f.pack(side="left")
        tk.Label(sort_f, text="Trier :", font=FONT_SMALL,
                 bg=COLORS["card"], fg=COLORS["subtext"]).pack(side="left", padx=4)
        
        self.sort_var = tk.StringVar(value="category")  # Tri par catégorie par défaut
        for val, lbl in [("category", "Categorie"), ("quantity", "Quantite"), ("name", "Nom")]:
            tk.Radiobutton(sort_f, text=lbl, variable=self.sort_var, value=val,
                           font=FONT_SMALL, bg=COLORS["card"], fg=COLORS["subtext"],
                           selectcolor=COLORS["entry"], activebackground=COLORS["card"],
                           command=self.refresh).pack(side="left")

        # ── FORMULAIRE (à droite) ──
        form = make_card(main, "Ajouter / Modifier un produit")
        form.pack(side="right", fill="y", ipadx=8, ipady=8)

        # Un sous-frame pour le formulaire (grid pour aligner proprement)
        fi = tk.Frame(form, bg=COLORS["card"])
        fi.pack(fill="both", expand=True)
        fi.columnconfigure(1, weight=1)

        # ── LES CHAMPS ──
        self.f_name     = make_entry(fi, "Nom",        0)
        self.f_ref      = make_entry(fi, "Reference",  1)
        self.f_qty      = make_entry(fi, "Quantite",   2, "0")
        self.f_price    = make_entry(fi, "Prix (EUR)", 3, "0.00")
        self.f_category = make_entry(fi, "Categorie",  4)

        # ── CATÉGORIES RAPIDES : des boutons pour remplir vite ──
        tk.Label(fi, text="Categories rapides :", font=FONT_SMALL,
                 bg=COLORS["card"], fg=COLORS["subtext"]).grid(
            row=5, column=0, sticky="w", padx=16, pady=(0, 4))
        cat_row = tk.Frame(fi, bg=COLORS["card"])
        cat_row.grid(row=5, column=1, sticky="ew", padx=(0, 16))
        
        for cat in ["Alimentaire", "Electronique", "Vetements", "Mobilier", "Autre"]:
            tk.Button(cat_row, text=cat, font=FONT_SMALL,
                      bg=COLORS["entry"], fg=COLORS["subtext"],
                      relief="flat", bd=0, cursor="hand2", padx=5, pady=2,
                      command=lambda c=cat: self.f_category.set(c)
                      ).pack(side="left", padx=2, pady=2)

        # ── LES BOUTONS D'ACTION PRINCIPAUX ──
        make_button(fi, "Ajouter au stock", self._add).grid(
            row=6, column=0, columnspan=2, padx=16, pady=(12, 4), sticky="ew")
        make_button(fi, "Mettre a jour", self._update,
                    color=COLORS["accent2"]).grid(
            row=7, column=0, columnspan=2, padx=16, pady=(0, 16), sticky="ew")

        # ── BADGE D'ALERTE STOCK FAIBLE ──
        self.low_badge = tk.Label(fi, text="", font=FONT_SMALL,
                                   bg=COLORS["card"], fg=COLORS["warning"])
        self.low_badge.grid(row=8, column=0, columnspan=2, padx=16, pady=(0, 8))

    def refresh(self):
        """
        Je rafraîchis le tableau des produits.
        
        Appelée au chargement, après chaque modif, quand on change le tri,
        ou quand on tape dans la recherche.
        """
        # Je vide le tableau
        self.tree.delete(*self.tree.get_children())
        
        # Je récupère les produits triés
        products = self.app.manager.get_products(self.sort_var.get())
        
        # Je filtre selon la recherche
        q = self.search_var.get().lower()
        if q:
            products = [p for p in products
                        if q in p["name"].lower() or q in p["reference"].lower()]
        
        # J'ajoute chaque produit au tableau
        for p in products:
            tag = "low" if p["quantity"] <= LOW_STOCK_THRESHOLD else ""
            self.tree.insert("", "end", iid=p["reference"],
                             values=(p["name"], p["reference"],
                                     p["quantity"], f"{p['price']:.2f}", p["category"]),
                             tags=(tag,))
        
        # Je mets à jour le badge d'alerte
        n_low = len(self.app.manager.low_stock_products())
        self.low_badge.configure(
            text=f"ALERTE: {n_low} produit(s) en stock faible" if n_low
                 else "Tous les stocks sont OK"
        )

    def _on_select(self, _e):
        """Quand je clique sur un produit, ça remplit le formulaire avec ses infos"""
        sel = self.tree.selection()
        if not sel:
            return
            
        p = self.app.manager._find_product(sel[0])
        if p:
            self.f_name.set(p["name"])
            self.f_ref.set(p["reference"])
            self.f_qty.set(p["quantity"])
            self.f_price.set(p["price"])
            self.f_category.set(p["category"])

    def _add(self):
        """J'ajoute un nouveau produit au stock"""
        try:
            self.app.manager.add_product(
                self.f_name.get(),
                self.f_ref.get(),
                self.f_qty.get(),
                self.f_price.get(),
                self.f_category.get()
            )
            
            # Je vide le formulaire pour le prochain ajout
            for v in (self.f_name, self.f_ref, self.f_qty, self.f_price, self.f_category):
                v.set("")
                
            self.refresh()
            messagebox.showinfo("Succes", "Produit ajoute avec succes.")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def _update(self):
        """Je mets à jour un produit existant avec les valeurs du formulaire"""
        ref = self.f_ref.get().strip()
        if not ref:
            messagebox.showwarning("Avertissement", "Selectionnez un produit d'abord.")
            return
            
        try:
            self.app.manager.update_product(
                ref,
                name=self.f_name.get().strip(),
                quantity=int(self.f_qty.get()),
                price=float(self.f_price.get()),
                category=self.f_category.get().strip(),
            )
            self.refresh()
            messagebox.showinfo("Succes", "Produit mis a jour.")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def _delete(self):
        """
        Je supprime un produit.
        
        ⚠️  C'est définitif ! Je demande toujours confirmation avant.
        """
        ref = self.f_ref.get().strip()
        if not ref:
            messagebox.showwarning("Avertissement", "Selectionnez un produit d'abord.")
            return
            
        if messagebox.askyesno("Confirmer", f"Supprimer le produit '{ref}' ?"):
            try:
                self.app.manager.delete_product(ref)
                
                for v in (self.f_name, self.f_ref, self.f_qty, self.f_price, self.f_category):
                    v.set("")
                    
                self.refresh()
            except Exception as e:
                messagebox.showerror("Erreur", str(e))


# ─────────────────────────────────────────────
#  PAGE VENTES : enregistrer et voir les transactions
# ─────────────────────────────────────────────
# Cette page permet de :
# - Enregistrer une vente (le stock se met à jour tout seul)
# - Voir l'historique complet
# - Sélectionner rapidement un produit
# - Calculer le chiffre d'affaires
# - Filtrer par période

class SalesFrame(tk.Frame):
    """
    La page pour gérer les ventes et voir l'historique.
    
    Elle contient :
    - Un formulaire pour enregistrer une vente
    - Une liste des produits pour une sélection rapide
    - Un tableau historique de toutes les ventes
    - Un compteur du CA en temps réel
    - Des alertes si le stock devient faible après une vente
    """
    
    def __init__(self, parent, app):
        """
        J'initialise la page des ventes.
        
        Args:
            parent: Le frame parent
            app: L'instance de l'app principale
        """
        super().__init__(parent, bg=COLORS["bg"])
        self.app = app
        self._build()

    def _build(self):
        """
        Je construis l'interface de la page des ventes.
        
        Organisation :
        1. Le titre
        2. Zone principale : historique à gauche, formulaire à droite
        """
        # ── LE TITRE ──
        tk.Label(self, text="Suivi des Ventes", font=FONT_TITLE,
                 bg=COLORS["bg"], fg=COLORS["text"]).pack(anchor="w", padx=24, pady=(20, 10))

        # ── ZONE PRINCIPALE ──
        main = tk.Frame(self, bg=COLORS["bg"])
        main.pack(fill="both", expand=True, padx=24, pady=8)

        # ── L'HISTORIQUE (à gauche) ──
        tbl = make_card(main, "Historique des ventes")
        tbl.pack(side="left", fill="both", expand=True, padx=(0, 12))

        cols = ("Date", "Produit", "Reference", "Quantite", "Prix unit.", "Total (EUR)", "Categorie")
        headers = ("Date", "Produit", "Référence", "Quantité", "Prix unit.", "Total (€)", "Catégorie")
        self.tree = ttk.Treeview(tbl, columns=cols, show="headings")
        
        for c, h in zip(cols, headers):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=95, anchor="center")
        self.tree.column("Produit", width=150, anchor="w")
        
        style_tree(self.tree)

        sb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        sb.pack(side="right", fill="y", pady=8)

        # ── LE CA TOTAL EN BAS ──
        self.total_lbl = tk.Label(tbl, text="Chiffre d'affaires : 0.00 EUR",
                                   font=("Segoe UI", 11, "bold"),
                                   bg=COLORS["card"], fg=COLORS["accent2"])
        self.total_lbl.pack(anchor="e", padx=16, pady=(0, 8))

        # ── LE FORMULAIRE (à droite) ──
        form = make_card(main, "Enregistrer une vente")
        form.pack(side="right", fill="y", ipadx=8, ipady=8)

        fi = tk.Frame(form, bg=COLORS["card"])
        fi.pack(fill="both", expand=True)
        fi.columnconfigure(1, weight=1)

        # ── LES CHAMPS ──
        self.s_ref = make_entry(fi, "Reference produit", 0)
        self.s_qty = make_entry(fi, "Quantite vendue",   1, "1")

        # ── SÉLECTION RAPIDE ──
        tk.Label(fi, text="Selection rapide :", font=FONT_SMALL,
                 bg=COLORS["card"], fg=COLORS["subtext"]).grid(
            row=2, column=0, sticky="nw", padx=16, pady=4)

        pf = tk.Frame(fi, bg=COLORS["entry"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        pf.grid(row=2, column=1, sticky="ew", padx=(0, 16), pady=4)
        
        self.picker = tk.Listbox(pf, font=FONT_SMALL,
                                  bg=COLORS["entry"], fg=COLORS["text"],
                                  selectbackground=COLORS["accent"],
                                  relief="flat", bd=0, height=8,
                                  exportselection=False)
        self.picker.pack(fill="both", expand=True)
        self.picker.bind("<<ListboxSelect>>", self._pick)

        # ── LE BOUTON POUR ENREGISTRER ──
        make_button(fi, "Enregistrer la vente", self._record,
                    color=COLORS["accent2"]).grid(
            row=3, column=0, columnspan=2, padx=16, pady=16, sticky="ew")

    def refresh(self):
        """
        Je rafraîchis la page des ventes.
        
        Met à jour :
        - L'historique
        - Le CA total
        - La liste de sélection rapide
        """
        self.tree.delete(*self.tree.get_children())
        
        # J'affiche les ventes les plus récentes en premier
        for s in reversed(self.app.manager.sales):
            self.tree.insert("", "end", values=(
                s["date"], s["product_name"], s["reference"],
                s["quantity"], f"{s['unit_price']:.2f}",
                f"{s['total']:.2f}", s["category"]
            ))
        
        total = self.app.manager.report_total_revenue()
        self.total_lbl.configure(text=f"Chiffre d'affaires total : {total:,.2f} EUR")

        # Je mets à jour la liste de sélection rapide
        self.picker.delete(0, "end")
        
        for p in self.app.manager.get_products("name"):
            display_text = f"{p['reference']} - {p['name']} (qte: {p['quantity']})"
            self.picker.insert("end", display_text)
            
            # En orange si le stock est faible
            if p["quantity"] <= LOW_STOCK_THRESHOLD:
                self.picker.itemconfigure("end", fg=COLORS["warning"])

    def _pick(self, _e):
        """Quand je clique sur un produit dans la liste, ça remplit la référence"""
        sel = self.picker.curselection()
        if sel:
            selected_text = self.picker.get(sel[0])
            reference = selected_text.split(" - ")[0]
            self.s_ref.set(reference)

    def _record(self):
        """
        J'enregistre une vente.
        
        Étapes :
        1. Je valide les champs
        2. J'appelle le manager pour enregistrer
        3. Le stock se met à jour tout seul
        4. Je rafraîchis les interfaces
        5. J'alerte si le stock devient faible
        """
        try:
            sale = self.app.manager.record_sale(
                self.s_ref.get().strip(),
                self.s_qty.get().strip()
            )
            
            # Je réinitialise le formulaire
            self.s_ref.set("")
            self.s_qty.set("1")
            
            self.refresh()
            self.app.frames["ProductsFrame"].refresh()
            
            # J'alerte si le stock est devenu faible
            p = self.app.manager._find_product(sale["reference"])
            if p and p["quantity"] <= LOW_STOCK_THRESHOLD:
                messagebox.showwarning(
                    "Alerte Stock Faible",
                    f"ATTENTION : Stock de '{p['name']}' = {p['quantity']} unite(s) restante(s).\n"
                    f"Pensez a reapprovisionner !"
                )
            else:
                messagebox.showinfo("Vente enregistree",
                                    f"{sale['quantity']} x {sale['product_name']}\n"
                                    f"Total : {sale['total']:.2f} EUR")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))


# ─────────────────────────────────────────────
#  PAGE FOURNISSEURS : gérer les relations commerciales
# ─────────────────────────────────────────────
# Cette page permet de :
# - Ajouter un fournisseur avec ses coordonnées
# - Modifier les infos existantes
# - Supprimer (avec confirmation)
# - Associer des produits à chaque fournisseur
# - Voir tous les fournisseurs triés

class SuppliersFrame(tk.Frame):
    """
    La page pour gérer les fournisseurs.
    
    Elle contient :
    - Un tableau avec tous les fournisseurs
    - Un formulaire pour ajouter/modifier
    - Un sélecteur de produits à associer
    - Des boutons d'action
    """
    
    def __init__(self, parent, app):
        """
        J'initialise la page des fournisseurs.
        
        Args:
            parent: Le frame parent
            app: L'instance de l'app principale
        """
        super().__init__(parent, bg=COLORS["bg"])
        self.app = app
        self._build()

    def _build(self):
        """
        Je construis l'interface de la page fournisseurs.
        
        Organisation :
        1. Le titre
        2. Zone principale : tableau à gauche, formulaire à droite
        """
        # ── LE TITRE ──
        tk.Label(self, text="Gestion des Fournisseurs", font=FONT_TITLE,
                 bg=COLORS["bg"], fg=COLORS["text"]).pack(anchor="w", padx=24, pady=(20, 10))

        # ── ZONE PRINCIPALE ──
        main = tk.Frame(self, bg=COLORS["bg"])
        main.pack(fill="both", expand=True, padx=24, pady=8)

        # ── LE TABLEAU (à gauche) ──
        tbl = make_card(main, "Liste des fournisseurs")
        tbl.pack(side="left", fill="both", expand=True, padx=(0, 12))

        cols = ("Nom", "Contact", "Email", "Telephone", "Produits fournis")
        self.tree = ttk.Treeview(tbl, columns=cols, show="headings", selectmode="browse")
        
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=130, anchor="w")
        self.tree.column("Produits fournis", width=240)
        
        style_tree(self.tree)

        sb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        sb.pack(side="right", fill="y", pady=8)
        
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # ── BARRE D'ACTION ──
        btn_row = tk.Frame(tbl, bg=COLORS["card"])
        btn_row.pack(fill="x", padx=8, pady=(0, 10))
        
        make_button(btn_row, "Supprimer", self._delete,
                    color=COLORS["danger"]).pack(side="right", padx=4)
        make_button(btn_row, "Modifier", self._update,
                    color=COLORS["accent3"]).pack(side="right", padx=4)

        # ── LE FORMULAIRE (à droite) ──
        form = make_card(main, "Ajouter un fournisseur")
        form.pack(side="right", fill="y", ipadx=8, ipady=8)

        fi = tk.Frame(form, bg=COLORS["card"])
        fi.pack(fill="both", expand=True)
        fi.columnconfigure(1, weight=1)

        # ── LES CHAMPS ──
        self.f_name    = make_entry(fi, "Nom",              0)
        self.f_contact = make_entry(fi, "Contact",          1)
        self.f_email   = make_entry(fi, "Email",            2)
        self.f_phone   = make_entry(fi, "Telephone",        3)
        self.f_prods   = make_entry(fi, "Produits fournis", 4)

        # ── INSTRUCTION ──
        tk.Label(fi, text="(references separees par virgules)", font=FONT_SMALL,
                 bg=COLORS["card"], fg=COLORS["subtext"]).grid(
            row=5, column=1, sticky="w", padx=(0, 16), pady=(0, 8))

        # ── SÉLECTEUR DE RÉFÉRENCES ──
        tk.Label(fi, text="Choisir references :", font=FONT_SMALL,
                 bg=COLORS["card"], fg=COLORS["subtext"]).grid(
            row=6, column=0, sticky="nw", padx=16, pady=4)
        
        pf = tk.Frame(fi, bg=COLORS["entry"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        pf.grid(row=6, column=1, sticky="ew", padx=(0, 16), pady=4)
        
        self.ref_picker = tk.Listbox(pf, font=FONT_SMALL,
                                      bg=COLORS["entry"], fg=COLORS["text"],
                                      selectbackground=COLORS["accent3"],
                                      relief="flat", bd=0, height=5,
                                      exportselection=False, selectmode="multiple")
        self.ref_picker.pack(fill="both", expand=True)
        self.ref_picker.bind("<<ListboxSelect>>", self._pick_refs)

        # ── LES BOUTONS ──
        make_button(fi, "Ajouter le fournisseur", self._add,
                    color=COLORS["accent3"]).grid(
            row=7, column=0, columnspan=2, padx=16, pady=(12, 4), sticky="ew")
        make_button(fi, "Mettre a jour", self._update,
                    color=COLORS["accent"]).grid(
            row=8, column=0, columnspan=2, padx=16, pady=(0, 16), sticky="ew")

    def refresh(self):
        """
        Je rafraîchis la page des fournisseurs.
        
        Met à jour le tableau et la liste des produits à associer.
        """
        self.tree.delete(*self.tree.get_children())
        
        for s in self.app.manager.get_suppliers():
            self.tree.insert("", "end", iid=s["name"],
                             values=(s["name"], s["contact"], s.get("email", ""),
                                     s.get("phone", ""),
                                     ", ".join(s.get("products", []))))
        
        self.ref_picker.delete(0, "end")
        for p in self.app.manager.get_products("name"):
            self.ref_picker.insert("end", f"{p['reference']} - {p['name']}")

    def _on_select(self, _e):
        """Quand je clique sur un fournisseur, ça remplit le formulaire"""
        sel = self.tree.selection()
        if not sel:
            return
            
        s = self.app.manager._find_supplier(sel[0])
        if s:
            self.f_name.set(s["name"])
            self.f_contact.set(s["contact"])
            self.f_email.set(s.get("email", ""))
            self.f_phone.set(s.get("phone", ""))
            self.f_prods.set(", ".join(s.get("products", [])))

    def _pick_refs(self, _e):
        """Quand je sélectionne des références, ça les met dans le champ produits"""
        selected = [self.ref_picker.get(i).split(" - ")[0]
                    for i in self.ref_picker.curselection()]
        self.f_prods.set(", ".join(selected))

    def _add(self):
        """J'ajoute un nouveau fournisseur"""
        try:
            prods = [p.strip() for p in self.f_prods.get().split(",") if p.strip()]
            
            self.app.manager.add_supplier(
                self.f_name.get(),
                self.f_contact.get(),
                self.f_email.get(),
                self.f_phone.get(),
                prods
            )
            
            for v in (self.f_name, self.f_contact, self.f_email, self.f_phone, self.f_prods):
                v.set("")
                
            self.refresh()
            messagebox.showinfo("Succes", "Fournisseur ajoute.")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def _update(self):
        """Je mets à jour un fournisseur existant"""
        sel = self.tree.selection()
        original = sel[0] if sel else self.f_name.get().strip()
        
        if not original:
            messagebox.showwarning("Avertissement", "Selectionnez un fournisseur d'abord.")
            return
            
        try:
            prods = [p.strip() for p in self.f_prods.get().split(",") if p.strip()]
            
            self.app.manager.update_supplier(
                original,
                name=self.f_name.get().strip(),
                contact=self.f_contact.get().strip(),
                email=self.f_email.get().strip(),
                phone=self.f_phone.get().strip(),
                products=prods,
            )
            
            self.refresh()
            messagebox.showinfo("Succes", "Fournisseur mis a jour.")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def _delete(self):
        """
        Je supprime un fournisseur.
        
        ⚠️  C'est définitif ! Je demande toujours confirmation avant.
        Les produits associés ne sont pas supprimés.
        """
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Avertissement", "Selectionnez un fournisseur d'abord.")
            return
            
        if messagebox.askyesno("Confirmer", f"Supprimer le fournisseur '{sel[0]}' ?"):
            try:
                self.app.manager.delete_supplier(sel[0])
                
                for v in (self.f_name, self.f_contact, self.f_email, self.f_phone, self.f_prods):
                    v.set("")
                    
                self.refresh()
            except Exception as e:
                messagebox.showerror("Erreur", str(e))


# ─────────────────────────────────────────────
#  PAGE RAPPORTS : stats et analyses commerciales
# ─────────────────────────────────────────────
# Cette page permet de :
# - Voir les produits les plus vendus
# - Analyser le CA par catégorie
# - Surveiller les niveaux de stock
# - Voir les alertes de stock faible
# - Filtrer tout ça par période

class ReportsFrame(tk.Frame):
    """
    La page des rapports et statistiques.
    
    Elle contient :
    - Des filtres par période
    - Le top 5 des produits vendus
    - Le CA par catégorie
    - Les niveaux de stock extrêmes
    - Les alertes détaillées pour les stocks critiques
    """
    
    def __init__(self, parent, app):
        """
        J'initialise la page des rapports.
        
        Args:
            parent: Le frame parent
            app: L'instance de l'app principale
        """
        super().__init__(parent, bg=COLORS["bg"])
        self.app = app
        self._build()

    def _build(self):
        """
        Je construis l'interface de la page des rapports.
        
        Organisation :
        1. Le titre
        2. Les filtres par période
        3. Trois cartes de rapports
        4. La zone d'alertes
        5. Le CA total en bas
        """
        # ── LE TITRE ──
        tk.Label(self, text="Rapports & Statistiques", font=FONT_TITLE,
                 bg=COLORS["bg"], fg=COLORS["text"]).pack(anchor="w", padx=24, pady=(20, 10))

        # ── LES FILTRES PAR PÉRIODE ──
        flt = tk.Frame(self, bg=COLORS["bg"])
        flt.pack(fill="x", padx=24, pady=(0, 12))
        
        tk.Label(flt, text="Periode :", font=FONT_LABEL,
                 bg=COLORS["bg"], fg=COLORS["subtext"]).pack(side="left")
        
        self.from_var = tk.StringVar()
        self.to_var   = tk.StringVar()
        
        for lbl, var in [("De :", self.from_var), ("A :", self.to_var)]:
            tk.Label(flt, text=lbl, font=FONT_LABEL,
                     bg=COLORS["bg"], fg=COLORS["subtext"]).pack(side="left", padx=(12, 4))
            tk.Entry(flt, textvariable=var, font=FONT_LABEL, width=12,
                     bg=COLORS["entry"], fg=COLORS["text"],
                     insertbackground=COLORS["text"],
                     relief="flat", bd=0).pack(side="left", ipady=5)
        
        tk.Label(flt, text="(AAAA-MM-JJ)", font=FONT_SMALL,
                 bg=COLORS["bg"], fg=COLORS["subtext"]).pack(side="left", padx=8)
        
        make_button(flt, "Actualiser", self.refresh).pack(side="left", padx=8)

        # ── TROIS CARTES DE RAPPORTS ──
        row1 = tk.Frame(self, bg=COLORS["bg"])
        row1.pack(fill="both", expand=True, padx=24, pady=(0, 8))

        def text_card(parent, title, pad=(0, 8)):
            """Petit helper pour créer une carte avec un champ texte monospace"""
            card = make_card(parent, title)
            card.pack(side="left", fill="both", expand=True, padx=pad)
            txt = tk.Text(card, font=FONT_MONO, wrap="word",
                          bg=COLORS["card"], fg=COLORS["text"],
                          relief="flat", bd=0, state="disabled", height=9)
            txt.pack(fill="both", expand=True, padx=12, pady=8)
            return txt

        self.top_text   = text_card(row1, "Top produits vendus",  pad=(0, 8))
        self.cat_text   = text_card(row1, "CA par categorie",      pad=(8, 8))
        self.stock_text = text_card(row1, "Niveaux de stock",      pad=(8, 0))

        # ── LA ZONE D'ALERTES ──
        row2 = tk.Frame(self, bg=COLORS["bg"])
        row2.pack(fill="x", padx=24, pady=(0, 8))

        alert_card = make_card(row2, "Alertes stock faible")
        alert_card.pack(side="left", fill="both", expand=True)
        
        self.alert_text = tk.Text(alert_card, font=FONT_MONO,
                                   bg=COLORS["card"], fg=COLORS["warning"],
                                   relief="flat", bd=0, state="disabled", height=3)
        self.alert_text.pack(fill="both", expand=True, padx=12, pady=8)

        # ── LE CA TOTAL EN BAS ──
        self.rev_lbl = tk.Label(self, text="", font=("Segoe UI", 15, "bold"),
                                 bg=COLORS["bg"], fg=COLORS["accent2"])
        self.rev_lbl.pack(pady=(0, 10))

    def refresh(self):
        """
        Je rafraîchis tous les rapports avec les filtres de période.
        
        Met à jour :
        - Le CA total
        - Le top 5 des produits
        - Le CA par catégorie
        - Les niveaux de stock
        - Les alertes stock faible
        """
        fd = self.from_var.get().strip() or None
        td = self.to_var.get().strip() or None

        # ── LE CA TOTAL ──
        total = self.app.manager.report_total_revenue(fd, td)
        self.rev_lbl.configure(text=f"Chiffre d'affaires total : {total:,.2f} EUR")

        # ── LE TOP PRODUITS ──
        top = self.app.manager.report_top_products(from_date=fd, to_date=td)
        top_text = "\n".join(f"{i+1}. {info['name']}\n"
                              f"   Qte: {info['qty']}  |  CA: {info['revenue']:.2f} EUR\n"
                              for i, (_, info) in enumerate(top)) or "(Aucune vente)"
        self._write(self.top_text, top_text)

        # ── LE CA PAR CATÉGORIE ──
        cats = self.app.manager.report_revenue_by_category(fd, td)
        cat_text = "\n".join(f"{cat:<20} {rev:>10.2f} EUR"
                              for cat, rev in sorted(cats.items(), key=lambda x: -x[1])
                              ) or "(Aucune vente)"
        self._write(self.cat_text, cat_text)

        # ── LES NIVEAUX DE STOCK ──
        low, high = self.app.manager.report_stock_levels()
        stock_text = "-- Stocks les plus bas --\n"
        stock_text += "\n".join(f"  {p['name']:<22} {p['quantity']:>4} unites" for p in low)
        stock_text += "\n\n-- Stocks les plus hauts --\n"
        stock_text += "\n".join(f"  {p['name']:<22} {p['quantity']:>4} unites" for p in high)
        self._write(self.stock_text, stock_text or "(Aucun produit)")

        # ── LES ALERTES STOCK FAIBLE ──
        low_list = self.app.manager.low_stock_products()
        if low_list:
            alert = "\n".join(
                f"  {'[RUPTURE]' if p['quantity']==0 else '[FAIBLE]'} "
                f"{p['name']} ({p['reference']}) -- {p['quantity']} restant(s)"
                for p in sorted(low_list, key=lambda x: x["quantity"])
            )
        else:
            alert = "  OK  Tous les produits ont un stock suffisant."
        self._write(self.alert_text, alert)

    @staticmethod
    def _write(widget, text):
        """
        J'écris dans un widget Text qui est normalement en lecture seule.
        
        Je l'active, j'écris, puis je le remets en lecture seule.
        """
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("end", text)
        widget.configure(state="disabled")


# ─────────────────────────────────────────────
#  POINT D'ENTRÉE
# ─────────────────────────────────────────────
# Ce code ne s'exécute que si on lance le fichier directement
# (pas quand on l'importe comme module)

if __name__ == "__main__":
    """
    Le point d'entrée de GestiStock.
    
    Ce code se lance uniquement avec `python gestion_stocks.py`.
    
    Ce qui se passe au démarrage :
    1. Je crée le dossier data/ s'il existe pas
    2. Je lance l'interface graphique
    3. Je démarre la boucle d'événements Tkinter
    """
    ensure_data_dir()
    app = App()
    app.mainloop()
