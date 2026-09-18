"""
Application de Gestion des Stocks — GestiStock
Mini-projet Python 2026
Auteur : Misbaou DIALLO
Extensions : Excel import/export, alertes stock faible, gestion fournisseurs
"""

# Bibliothèques Python nécessaires pour faire tourner notre application
import json  # Pour sauvegarder/lire les données en format JSON
import csv   # Pour exporter en CSV si besoin
import os    # Pour gérer les chemins de fichiers et dossiers
from datetime import datetime, date  # Pour gérer les dates des ventes
import tkinter as tk  # Interface graphique principale
from tkinter import ttk, messagebox, filedialog  # Composants graphiques avancés
from collections import defaultdict  # Pour les calculs de statistiques

# On essaie d'importer openpyxl pour les fonctionnalités Excel
# Si c'est pas installé, on désactive juste les boutons Excel
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    EXCEL_OK = True  # Excel est disponible !
except ImportError:
    EXCEL_OK = False  # Excel pas dispo, mais l'app tourne quand même


# ─────────────────────────────────────────────
#  COUCHE DONNÉES : Tout ce qui touche au stockage
# ─────────────────────────────────────────────

# Chemins vers nos fichiers de données
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")  # Dossier data/
PRODUCTS_FILE   = os.path.join(DATA_DIR, "produits.json")    # Nos produits
SALES_FILE      = os.path.join(DATA_DIR, "ventes.json")       # Historique des ventes
SUPPLIERS_FILE  = os.path.join(DATA_DIR, "fournisseurs.json") # Nos fournisseurs
LOW_STOCK_THRESHOLD = 5  # Seuil d'alerte : en dessous de 5 pièces, on prévient !


def ensure_data_dir():
    """Crée le dossier data/ s'il existe pas encore"""
    os.makedirs(DATA_DIR, exist_ok=True)


def load_json(path):
    """Charge un fichier JSON, retourne une liste vide s'il existe pas"""
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return []  # Fichier inexistant = liste vide


def save_json(path, data):
    """Sauvegarde des données en JSON avec joli formatage"""
    ensure_data_dir()  # S'assure que le dossier existe
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)  # Pretty print !


def export_csv(path, data, fieldnames):
    """Exporte des données en CSV pour Excel/Google Sheets"""
    ensure_data_dir()
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()  # Ligne d'en-têtes
        writer.writerows(data)  # Toutes les lignes de données


# ─────────────────────────────────────────────
#  FONCTIONS EXCEL : Pour les exports/import Excel avancés
# ─────────────────────────────────────────────

def _xl_header_style(ws, row, cols, fill_hex="4C3F9F"):
    """Style joli pour les en-têtes de tableaux Excel"""
    fill   = PatternFill("solid", fgColor=fill_hex)  # Fond de couleur
    font   = Font(bold=True, color="FFFFFF", size=11)  # Texte blanc gras
    border = Border(bottom=Side(style="medium", color="FFFFFF"))  # Ligne blanche en dessous
    
    # Applique le style à chaque colonne d'en-tête
    for col, title in enumerate(cols, 1):
        cell = ws.cell(row=row, column=col, value=title)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    ws.row_dimensions[row].height = 22  # Hauteur un peu plus grande pour les en-têtes


def _xl_autofit(ws):
    """Ajuste automatiquement la largeur des colonnes Excel"""
    for col in ws.columns:
        # Calcule la longueur max du contenu de la colonne
        max_len = max((len(str(c.value or "")) for c in col), default=8)
        # Ajuste la largeur (max 40 caractères pour pas trop large)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 40)


def export_excel(products, sales, suppliers, path):
    """Crée un fichier Excel avec 4 feuilles : Produits, Ventes, Fournisseurs, Stock Faible"""
    wb = openpyxl.Workbook()  # Nouveau classeur Excel

    # ── Feuille 1 : Produits ──
    ws_p = wb.active
    ws_p.title = "Produits"
    cols_p = ["Nom", "Reference", "Quantite", "Prix (EUR)", "Categorie"]
    _xl_header_style(ws_p, 1, cols_p)  # En-têtes stylés
    
    # Style pour les produits en stock faible
    low_fill  = PatternFill("solid", fgColor="FFF3CD")  # Fond jaune
    warn_font = Font(color="856404")  # Texte orangé
    
    # Ajoute tous les produits triés par catégorie
    for p in sorted(products, key=lambda x: x["category"]):
        row = ws_p.max_row + 1
        ws_p.append([p["name"], p["reference"], p["quantity"], p["price"], p["category"]])
        
        # Si stock faible, on met en évidence
        if p["quantity"] <= LOW_STOCK_THRESHOLD:
            for col in range(1, 6):
                ws_p.cell(row=row, column=col).fill = low_fill
                ws_p.cell(row=row, column=col).font = warn_font
    _xl_autofit(ws_p)  # Ajuste les largeurs

    # ── Feuille 2 : Ventes ──
    ws_s = wb.create_sheet("Ventes")
    cols_s = ["Date", "Produit", "Reference", "Quantite", "Prix unit. (EUR)", "Total (EUR)", "Categorie"]
    _xl_header_style(ws_s, 1, cols_s, fill_hex="1A7A4A")  # Vert pour les ventes
    
    # Ajoute toutes les ventes triées par date (plus récent en premier)
    for s in sorted(sales, key=lambda x: x["date"], reverse=True):
        ws_s.append([s["date"], s["product_name"], s["reference"],
                     s["quantity"], s["unit_price"], s["total"], s["category"]])
    
    # Ligne de total si y'a des ventes
    if sales:
        total_row = ws_s.max_row + 1
        ws_s.cell(total_row, 1, "TOTAL").font = Font(bold=True)
        ws_s.cell(total_row, 6, sum(s["total"] for s in sales)).font = Font(bold=True, color="1A7A4A")
    _xl_autofit(ws_s)

    # ── Feuille 3 : Fournisseurs ──
    ws_f = wb.create_sheet("Fournisseurs")
    cols_f = ["Nom", "Contact", "Email", "Telephone", "Produits fournis"]
    _xl_header_style(ws_f, 1, cols_f, fill_hex="8B3A8B")  # Violet pour fournisseurs
    
    # Ajoute tous les fournisseurs avec leurs infos
    for f in suppliers:
        ws_f.append([f["name"], f["contact"], f.get("email", ""),
                     f.get("phone", ""), ", ".join(f.get("products", []))])
    _xl_autofit(ws_f)

    # ── Feuille 4 : Alertes stock faible ──
    ws_a = wb.create_sheet("Stock Faible")
    cols_a = ["Nom", "Reference", "Quantite", "Categorie", "Statut"]
    _xl_header_style(ws_a, 1, cols_a, fill_hex="C0392B")  # Rouge pour alertes
    red_fill = PatternFill("solid", fgColor="FDECEA")  # Fond rouge clair
    
    # Ajoute seulement les produits en stock faible, triés par quantité croissante
    for p in sorted(products, key=lambda x: x["quantity"]):
        if p["quantity"] <= LOW_STOCK_THRESHOLD:
            status = "CRITIQUE" if p["quantity"] == 0 else "FAIBLE"  # Plus parlant !
            row = ws_a.max_row + 1
            ws_a.append([p["name"], p["reference"], p["quantity"], p["category"], status])
            # Met tout en rouge pour bien attirer l'attention
            for col in range(1, 6):
                ws_a.cell(row=row, column=col).fill = red_fill

    wb.save(path)  # Sauvegarde le fichier Excel
    return path


def import_excel_products(path):
    """Importe les produits depuis un fichier Excel (feuille 'Produits')"""
    wb = openpyxl.load_workbook(path, read_only=True)
    
    # Cherche une feuille qui contient "produit" dans son nom
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
    
    # Parcourt toutes les lignes de la feuille
    for row in ws.iter_rows(values_only=True):
        if headers is None:
            # Première ligne = en-têtes
            headers = [str(h).strip() if h else "" for h in row]
            continue
        if not any(row):
            continue  # Ligne vide, on passe
            
        d = dict(zip(headers, row))  # Associe en-têtes → valeurs
        
        # Conversion des noms de colonnes français → anglais (avec et sans accents)
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
        
        # Vérifie qu'on a les champs obligatoires
        if "name" in product and "reference" in product:
            # Valeurs par défaut si manquantes
            product.setdefault("quantity", 0)
            product.setdefault("price", 0.0)
            product.setdefault("category", "Autre")
            
            try:
                # Conversion des types
                product["quantity"] = int(product["quantity"])
                product["price"]    = float(product["price"])
                products.append(product)
            except (ValueError, TypeError):
                pass  # Ligne invalide, on ignore
    return products


# ─────────────────────────────────────────────
#  LOGIQUE MÉTIER : Le cerveau de notre application
# ─────────────────────────────────────────────

class StockManager:
    """Le gestionnaire principal qui gère tout : produits, ventes, fournisseurs..."""
    
    def __init__(self):
        """Au démarrage, charge toutes les données depuis les fichiers JSON"""
        self.products  = load_json(PRODUCTS_FILE)   # Charge les produits
        self.sales     = load_json(SALES_FILE)      # Charge l'historique des ventes
        self.suppliers = load_json(SUPPLIERS_FILE)  # Charge les fournisseurs

    # ── GESTION DES PRODUITS ──────────────────────────────

    def _find_product(self, ref):
        """Cherche un produit par sa référence (utilisé en interne)"""
        for p in self.products:
            if p["reference"] == ref:
                return p
        return None  # Pas trouvé

    def add_product(self, name, reference, quantity, price, category):
        """Ajoute un nouveau produit au stock"""
        # Validation des champs obligatoires
        if not str(name).strip():
            raise ValueError("Le nom du produit est obligatoire.")
        if not str(reference).strip():
            raise ValueError("La reference est obligatoire.")
        if self._find_product(str(reference).strip()):
            raise ValueError(f"Reference '{reference}' deja utilisee.")
        
        # Création du produit avec nettoyage des données
        product = {
            "name":      str(name).strip(),
            "reference": str(reference).strip(),
            "quantity":  int(quantity),
            "price":     float(price),
            "category":  str(category).strip() or "Autre",  # Catégorie par défaut
        }
        self.products.append(product)
        save_json(PRODUCTS_FILE, self.products)  # Sauvegarde immédiate
        return product

    def update_product(self, reference, **kwargs):
        """Modifie les informations d'un produit existant"""
        p = self._find_product(reference)
        if not p:
            raise ValueError(f"Produit '{reference}' introuvable.")
        
        # Met à jour seulement les champs autorisés
        for k, v in kwargs.items():
            if k in {"name", "quantity", "price", "category"}:
                p[k] = v
        save_json(PRODUCTS_FILE, self.products)
        return p

    def delete_product(self, reference):
        """Supprime un produit du stock"""
        p = self._find_product(reference)
        if not p:
            raise ValueError(f"Produit '{reference}' introuvable.")
        self.products.remove(p)
        save_json(PRODUCTS_FILE, self.products)

    def get_products(self, sort_by="category"):
        """Retourne tous les produits, triés selon le critère demandé"""
        reverse = (sort_by == "quantity")  # Tri décroissant seulement pour la quantité
        default_value = 0 if reverse else ""  # Évite le bug de type mixte
        return sorted(self.products,
                      key=lambda p: p.get(sort_by, default_value),
                      reverse=reverse)

    def low_stock_products(self):
        """Retourne la liste des produits en stock faible (≤ 5 unités)"""
        return [p for p in self.products if p["quantity"] <= LOW_STOCK_THRESHOLD]

    def import_products_from_excel(self, path):
        """
        Importe des produits depuis un fichier Excel et met à jour le stock.
        
        Args:
            path (str): Chemin vers le fichier Excel à importer
            
        Returns:
            tuple: (added, skipped) - nombre de produits ajoutés et ignorés
            
        Note:
            - Les produits avec référence déjà existante sont ignorés (pas d'écrasement)
            - Les nouveaux produits sont ajoutés à la fin de la liste
            - Le fichier produits.json est sauvegardé automatiquement
        """
        imported = import_excel_products(path)  # Appelle la fonction d'import Excel
        added, skipped = 0, 0  # Compteurs pour les statistiques
        
        # Parcourt tous les produits importés du fichier Excel
        for p in imported:
            if self._find_product(p["reference"]):
                skipped += 1  # Référence déjà existante → on ignore ce produit
            else:
                self.products.append(p)  # Nouveau produit → on l'ajoute au stock
                added += 1   # Incrémente le compteur de produits ajoutés
        
        # Sauvegarde immédiate pour persister les changements
        save_json(PRODUCTS_FILE, self.products)
        return added, skipped  # Retourne les statistiques d'import

    # ── GESTION DES VENTES ─────────────────────────────────
    # Cette section gère tout ce qui concerne les ventes :
    # - Enregistrement des ventes avec mise à jour automatique du stock
    # - Historique des ventes avec filtrage par dates
    # - Calculs automatiques (totaux, chiffre d'affaires)

    def record_sale(self, reference, quantity):
        """
        Enregistre une vente et met à jour le stock automatiquement.
        
        C'est la fonction la plus importante : elle fait deux choses en une :
        1. Déduit la quantité vendue du stock du produit
        2. Crée un enregistrement de vente dans l'historique
        
        Args:
            reference (str): Référence unique du produit vendu
            quantity (int): Quantité vendue (doit être positive)
            
        Returns:
            dict: L'enregistrement de vente créé
            
        Raises:
            ValueError: Si produit inexistant, quantité invalide ou stock insuffisant
            
        Exemple:
            >>> sale = manager.record_sale("ELEC-001", 2)
            >>> print(sale["total"])  # Affiche le prix total de la vente
        """
        # Étape 1 : Vérifier que le produit existe dans notre stock
        p = self._find_product(reference)
        if not p:
            raise ValueError(f"Produit '{reference}' introuvable.")
        
        # Étape 2 : Validation de la quantité (doit être un nombre positif)
        qty = int(quantity)
        if qty <= 0:
            raise ValueError("La quantite doit etre positive.")
        if p["quantity"] < qty:
            raise ValueError(f"Stock insuffisant. Disponible : {p['quantity']}.")
        
        # Étape 3 : Mettre à jour le stock (déduction immédiate)
        p["quantity"] -= qty  # Stock mis à jour en temps réel
        
        # Étape 4 : Créer l'enregistrement de vente pour l'historique
        sale = {
            "reference":    reference,        # Référence du produit
            "product_name": p["name"],       # Nom lisible du produit
            "quantity":     qty,              # Quantité vendue
            "unit_price":   p["price"],       # Prix unitaire au moment de la vente
            "total":        round(qty * p["price"], 2),  # Total calculé et arrondi
            "category":     p["category"],     # Catégorie pour les stats
            "date":         date.today().isoformat(),   # Date du jour (ISO: YYYY-MM-DD)
        }
        self.sales.append(sale)  # Ajoute à l'historique des ventes
        
        # Étape 5 : Sauvegarder les changements dans les deux fichiers
        save_json(PRODUCTS_FILE, self.products)  # Stock mis à jour
        save_json(SALES_FILE, self.sales)        # Historique sauvegardé
        return sale  # Retourne l'enregistrement créé (utile pour l'interface)

    def get_sales(self, from_date=None, to_date=None):
        """
        Retourne les ventes avec filtrage optionnel par période.
        
        Cette fonction est utilisée pour :
        - Afficher l'historique complet des ventes
        - Générer des rapports sur une période spécifique
        - Calculer des statistiques temporelles
        
        Args:
            from_date (str, optional): Date de début (format YYYY-MM-DD). Si None, pas de limite.
            to_date (str, optional): Date de fin (format YYYY-MM-DD). Si None, pas de limite.
            
        Returns:
            list: Liste des enregistrements de vente filtrés
            
        Exemples:
            >>> # Toutes les ventes
            >>> all_sales = manager.get_sales()
            >>> # Ventes du mois de janvier 2026
            >>> jan_sales = manager.get_sales("2026-01-01", "2026-01-31")
            >>> # Ventes depuis le 1er mars 2026
            >>> recent_sales = manager.get_sales(from_date="2026-03-01")
        """
        result = self.sales  # Commence avec toutes les ventes
        
        # Filtre par date de début si spécifiée
        if from_date:
            result = [s for s in result if s["date"] >= from_date]
            
        # Filtre par date de fin si spécifiée    
        if to_date:
            result = [s for s in result if s["date"] <= to_date]
            
        return result

    # ── RAPPORTS ET STATISTIQUES ───────────────────────────────
    # Ces fonctions transforment nos données brutes en informations utiles :
    # - Top produits vendus (pour savoir quoi commander en plus)
    # - Chiffre d'affaires par catégorie (pour savoir ce qui rapporte le plus)
    # - Niveaux de stock (pour éviter les ruptures)
    # - Alertes stock faible (pour réapprovisionner à temps)

    def report_top_products(self, n=5, from_date=None, to_date=None):
        """
        Calcule les n produits les plus vendus avec chiffre d'affaires.
        
        Cette fonction analyse l'historique des ventes pour déterminer :
        - Quels produits se vendent le plus (en quantité)
        - Quels produits génèrent le plus de chiffre d'affaires
        
        Args:
            n (int): Nombre de produits à retourner (défaut: 5)
            from_date (str, optional): Date de début pour l'analyse
            to_date (str, optional): Date de fin pour l'analyse
            
        Returns:
            list: Liste de tuples (reference, {"qty": int, "revenue": float, "name": str})
                  triée par quantité décroissante
                  
        Exemple:
            >>> top_3 = manager.report_top_products(3)
            >>> for ref, stats in top_3:
            ...     print(f"{stats['name']}: {stats['qty']} vendus, {stats['revenue']}€ CA")
        """
        # Utilise un defaultdict pour accumuler les stats par produit
        totals = defaultdict(lambda: {"qty": 0, "revenue": 0.0, "name": ""})
        
        # Parcourt toutes les ventes de la période et cumule les stats
        for s in self.get_sales(from_date, to_date):
            ref = s["reference"]
            totals[ref]["qty"]     += s["quantity"]      # Quantité totale vendue
            totals[ref]["revenue"] += s["total"]         # CA total généré
            totals[ref]["name"]     = s["product_name"]   # Nom du produit (gardé en dernier)
        
        # Convertit en liste et trie par quantité décroissante
        sorted_products = sorted(totals.items(), key=lambda x: x[1]["qty"], reverse=True)
        return sorted_products[:n]  # Retourne seulement les n premiers

    def report_revenue_by_category(self, from_date=None, to_date=None):
        """
        Calcule le chiffre d'affaires total par catégorie de produits.
        
        Utile pour :
        - Savoir quelles catégories sont les plus rentables
        - Prendre des décisions sur les catégories à développer
        - Analyser la répartition du chiffre d'affaires
        
        Args:
            from_date (str, optional): Date de début pour l'analyse
            to_date (str, optional): Date de fin pour l'analyse
            
        Returns:
            dict: Dictionnaire {categorie: chiffre_affaires} trié par ordre naturel
            
        Exemple:
            >>> revenue = manager.report_revenue_by_category()
            >>> print(f"Électronique: {revenue.get('Électronique', 0)}€")
        """
        cats = defaultdict(float)  # Accumule le CA par catégorie
        
        # Additionne le total de chaque vente à sa catégorie correspondante
        for s in self.get_sales(from_date, to_date):
            cats[s["category"]] += s["total"]  # Cumul du CA par catégorie
            
        return dict(cats)  # Convertit defaultdict en dict normal

    def report_total_revenue(self, from_date=None, to_date=None):
        """
        Calcule le chiffre d'affaires total sur une période donnée.
        
        C'est l'indicateur principal de performance commerciale.
        Utile pour :
        - Suivre les objectifs de vente
        - Comparer les performances entre périodes
        - Calculer la rentabilité globale
        
        Args:
            from_date (str, optional): Date de début (format YYYY-MM-DD)
            to_date (str, optional): Date de fin (format YYYY-MM-DD)
            
        Returns:
            float: Chiffre d'affaires total en euros
            
        Exemple:
            >>> # CA du mois dernier
            >>> ca_mois = manager.report_total_revenue("2026-02-01", "2026-02-28")
            >>> print(f"Chiffre d'affaires février: {ca_mois:.2f}€")
        """
        # Utilise get_sales() pour le filtrage, puis somme tous les totaux
        return sum(s["total"] for s in self.get_sales(from_date, to_date))

    def report_stock_levels(self):
        """
        Identifie les produits avec les niveaux de stock les plus extrêmes.
        
        Cette fonction aide à la gestion des stocks en montrant :
        - Les 5 produits avec le moins de stock (risque de rupture)
        - Les 5 produits avec le plus de stock (surstock potentiel)
        
        Returns:
            tuple: (lowest_stock, highest_stock)
                - lowest_stock (list): 5 produits avec le moins de stock (croissant)
                - highest_stock (list): 5 produits avec le plus de stock (décroissant)
                
        Note:
            Si moins de 5 produits existent, retourne tous les produits disponibles.
            
        Exemple:
            >>> low, high = manager.report_stock_levels()
            >>> print("Produits en risque:")
            >>> for p in low:
            ...     print(f"{p['name']}: {p['quantity']} unités")
        """
        if not self.products:
            return [], []  # Pas de produits = listes vides
            
        # Trie tous les produits par quantité (croissant)
        sorted_products = sorted(self.products, key=lambda x: x["quantity"])
        
        # Les 5 premiers = stock le plus bas
        lowest_stock = sorted_products[:5]
        
        # Les 5 derniers = stock le plus haut (inversé pour avoir le plus haut en premier)
        highest_stock = sorted_products[-5:][::-1]
        
        return lowest_stock, highest_stock

    # ── GESTION DES FOURNISSEURS (Extension 3) ───────────────
    # Les fournisseurs sont essentiels pour la gestion des achats :
    # - Ajouter de nouveaux fournisseurs avec leurs coordonnées
    # - Associer des produits à chaque fournisseur
    # - Gérer les informations de contact pour les commandes
    # - Suivre les relations commerciales

    def _find_supplier(self, name):
        """
        Cherche un fournisseur par son nom (insensible à la casse).
        
        Cette fonction interne est utilisée par les autres méthodes de gestion
        des fournisseurs pour éviter la duplication du code de recherche.
        
        Args:
            name (str): Nom du fournisseur à chercher
            
        Returns:
            dict or None: Objet fournisseur trouvé ou None si pas trouvé
            
        Note:
            La recherche est insensible à la casse : "Apple" == "apple" == "APPLE"
        """
        for s in self.suppliers:
            if s["name"].lower() == name.lower():  # Comparison insensible à la casse
                return s
        return None  # Fournisseur non trouvé

    def add_supplier(self, name, contact, email="", phone="", products=None):
        """
        Ajoute un nouveau fournisseur à la base de données.
        
        Crée un nouveau fournisseur avec toutes les informations nécessaires
        pour établir une relation commerciale et passer des commandes.
        
        Args:
            name (str): Nom unique du fournisseur (obligatoire)
            contact (str): Nom de la personne de contact (obligatoire)
            email (str, optional): Adresse email du fournisseur
            phone (str, optional): Numéro de téléphone
            products (list, optional): Liste des références produits fournis
            
        Returns:
            dict: Le fournisseur créé avec toutes ses informations
            
        Raises:
            ValueError: Si le nom ou le contact sont vides, ou si fournisseur existe déjà
            
        Exemple:
            >>> supplier = manager.add_supplier(
            ...     "TechDistrib",
            ...     "Jean Dupont",
            ...     email="contact@techdistrib.fr",
            ...     phone="01 23 45 67 89",
            ...     products=["ELEC-001", "ELEC-002"]
            ... )
        """
        # Validation des champs obligatoires
        if not str(name).strip():
            raise ValueError("Le nom du fournisseur est obligatoire.")
        if not str(contact).strip():
            raise ValueError("Le nom du contact est obligatoire.")
            
        # Vérifie que le fournisseur n'existe pas déjà
        if self._find_supplier(name):
            raise ValueError(f"Fournisseur '{name}' existe deja.")
        
        # Création de l'objet fournisseur avec nettoyage des données
        sup = {
            "name":     str(name).strip(),      # Nom nettoyé des espaces
            "contact":  str(contact).strip(),   # Contact nettoyé
            "email":    str(email).strip(),     # Email (peut être vide)
            "phone":    str(phone).strip(),     # Téléphone (peut être vide)
            "products": [p.strip() for p in (products or []) if p.strip()],  # Nettoie la liste
        }
        self.suppliers.append(sup)  # Ajoute à la liste des fournisseurs
        save_json(SUPPLIERS_FILE, self.suppliers)  # Sauvegarde immédiate
        return sup  # Retourne le fournisseur créé

    def update_supplier(self, original_name, **kwargs):
        """
        Modifie les informations d'un fournisseur existant.
        
        Permet de mettre à jour n'importe quelle information d'un fournisseur
        (coordonnées, produits, etc.) en spécifiant seulement les champs à modifier.
        
        Args:
            original_name (str): Nom actuel du fournisseur à modifier
            **kwargs: Champs à modifier (name, contact, email, phone, products)
            
        Returns:
            dict: Le fournisseur mis à jour
            
        Raises:
            ValueError: Si le fournisseur n'est pas trouvé
            
        Exemple:
            >>> # Change seulement l'email et le téléphone
            >>> updated = manager.update_supplier(
            ...     "TechDistrib",
            ...     email="nouveau@techdistrib.fr",
            ...     phone="01 99 88 77 66"
            ... )
        """
        # Recherche du fournisseur à modifier
        s = self._find_supplier(original_name)
        if not s:
            raise ValueError(f"Fournisseur '{original_name}' introuvable.")
            
        # Mise à jour seulement des champs autorisés et fournis
        for k, v in kwargs.items():
            if k in {"name", "contact", "email", "phone", "products"}:
                s[k] = v  # Met à jour le champ
                
        save_json(SUPPLIERS_FILE, self.suppliers)  # Sauvegarde les changements
        return s  # Retourne le fournisseur modifié

    def delete_supplier(self, name):
        """
        Supprime définitivement un fournisseur de la base de données.
        
        ⚠️  ATTENTION : Cette action est irréversible !
        Le fournisseur et toutes ses informations seront supprimés.
        
        Args:
            name (str): Nom du fournisseur à supprimer
            
        Raises:
            ValueError: Si le fournisseur n'est pas trouvé
            
        Note:
            Les produits associés au fournisseur ne sont PAS supprimés.
            Seule la relation fournisseur-produits est perdue.
            
        Exemple:
            >>> manager.delete_supplier("OldSupplier")
            >>> print("Fournisseur supprimé avec succès")
        """
        s = self._find_supplier(name)
        if not s:
            raise ValueError(f"Fournisseur '{name}' introuvable.")
            
        self.suppliers.remove(s)  # Suppression de la liste
        save_json(SUPPLIERS_FILE, self.suppliers)  # Sauvegarde pour persister

    def get_suppliers(self):
        """
        Retourne tous les fournisseurs triés par ordre alphabétique.
        
        Cette fonction est utilisée par l'interface pour afficher
        la liste complète des fournisseurs dans un ordre cohérent.
        
        Returns:
            list: Liste de tous les fournisseurs triés par nom (A-Z)
            
        Exemple:
            >>> suppliers = manager.get_suppliers()
            >>> for sup in suppliers:
            ...     print(f"{sup['name']} - {sup['contact']}")
        """
        return sorted(self.suppliers, key=lambda s: s["name"])  # Tri alphabétique

    # ── FONCTIONS D'EXPORT ────────────────────────────────
    # Ces fonctions permettent d'exporter nos données dans différents formats :
    # - CSV : Compatible Excel, Google Sheets, LibreOffice
    # - Excel : Format natif avec mise en forme avancée
    # - Utile pour les rapports, sauvegardes, et analyses externes

    def export_products_csv(self):
        """
        Exporte la liste complète des produits au format CSV.
        
        Crée un fichier produits.csv dans le dossier data/ avec :
        - Nom du produit
        - Référence unique
        - Quantité en stock
        - Prix unitaire
        - Catégorie
        
        Returns:
            str: Chemin complet du fichier CSV créé
            
        Note:
            Le fichier utilise l'encodage UTF-8 pour gérer les caractères spéciaux.
            Peut être ouvert directement avec Excel ou Google Sheets.
            
        Exemple:
            >>> path = manager.export_products_csv()
            >>> print(f"Produits exportés dans : {path}")
        """
        path = os.path.join(DATA_DIR, "produits.csv")  # Chemin du fichier
        # Colonnes exportées dans l'ordre logique
        columns = ["name", "reference", "quantity", "price", "category"]
        export_csv(path, self.products, columns)  # Appelle la fonction d'export générique
        return path  # Retourne le chemin pour l'interface

    def export_sales_csv(self):
        """
        Exporte l'historique complet des ventes au format CSV.
        
        Crée un fichier ventes.csv avec toutes les transactions :
        - Date de vente (format ISO)
        - Référence et nom du produit
        - Quantité vendue
        - Prix unitaire au moment de la vente
        - Total de la transaction
        - Catégorie du produit
        
        Returns:
            str: Chemin complet du fichier CSV créé
            
        Utilité:
        - Analyse des ventes dans Excel/Google Sheets
        - Création de graphiques et statistiques
        - Archivage des transactions commerciales
        
        Exemple:
            >>> path = manager.export_sales_csv()
            >>> print(f"Historique des ventes exporté dans : {path}")
        """
        path = os.path.join(DATA_DIR, "ventes.csv")  # Chemin du fichier
        # Colonnes dans l'ordre chronologique logique
        columns = [
            "date",        # Date de la vente
            "reference",   # Référence du produit
            "product_name", # Nom lisible du produit
            "quantity",    # Quantité vendue
            "unit_price",  # Prix unitaire (historique)
            "total",       # Total de la transaction
            "category"     # Catégorie pour les stats
        ]
        export_csv(path, self.sales, columns)  # Export avec les colonnes définies
        return path

    def export_suppliers_csv(self):
        """
        Exporte la liste des fournisseurs au format CSV.
        
        Crée un fichier fournisseurs.csv avec :
        - Nom du fournisseur
        - Personne de contact
        - Email et téléphone
        - Liste des produits fournis (séparés par des virgules)
        
        Returns:
            str: Chemin complet du fichier CSV créé
            
        Particularité:
        La liste des produits est transformée en chaîne de caractères
        pour être compatible avec le format CSV.
        
        Exemple:
            >>> path = manager.export_suppliers_csv()
            >>> print(f"Fournisseurs exportés dans : {path}")
        """
        path = os.path.join(DATA_DIR, "fournisseurs.csv")  # Chemin du fichier
        
        # Prépare les données : transforme la liste de produits en chaîne
        rows = []
        for s in self.suppliers:
            row = {
                **s,  # Copie toutes les infos du fournisseur
                "products": ", ".join(s.get("products", []))  # Liste → chaîne
            }
            rows.append(row)
        
        # Colonnes exportées
        columns = ["name", "contact", "email", "phone", "products"]
        export_csv(path, rows, columns)
        return path

    def export_all_excel(self):
        """
        Exporte TOUTES les données dans un seul fichier Excel multi-feuilles.
        
        C'est la fonction d'export la plus complète :
        - Feuille 1 : Produits (avec mise en évidence des stocks faibles)
        - Feuille 2 : Ventes (triées par date, avec total)
        - Feuille 3 : Fournisseurs (avec coordonnées complètes)
        - Feuille 4 : Alertes stock faible (produits critiques)
        
        Returns:
            str: Chemin complet du fichier Excel créé
            
        Avantages par rapport au CSV :
        - Mise en forme professionnelle (couleurs, polices)
        - Plusieurs feuilles dans un seul fichier
        - Filtres et tris automatiques dans Excel
        - Graphiques faciles à créer
        
        Note:
        Nécessite la bibliothèque openpyxl installée.
        
        Exemple:
            >>> path = manager.export_all_excel()
            >>> print(f"Export complet Excel : {path}")
            >>> # Le fichier peut être ouvert directement dans Excel
        """
        path = os.path.join(DATA_DIR, "gestion_stocks.xlsx")  # Nom du fichier
        export_excel(self.products, self.sales, self.suppliers, path)  # Appelle la fonction Excel
        return path


# ─────────────────────────────────────────────
#  INTERFACE GRAPHIQUE : Style et couleurs de notre app
# ─────────────────────────────────────────────

# Palette de couleurs moderne pour notre interface sombre
COLORS = {
    "bg":      "#0f1117",  # Fond principal (très sombre)
    "panel":   "#1a1d27",  # Panneaux latéraux
    "card":    "#22263a",  # Cartes et conteneurs
    "accent":  "#6c63ff",  # Couleur principale (violet)
    "accent2": "#00d4aa",  # Secondaire (vert turquoise)
    "accent3": "#c084fc",  # Tertiaire (violet clair)
    "danger":  "#ff5370",  # Actions dangereuses (rouge)
    "warning": "#ffcb6b",  # Alertes (orange)
    "text":    "#e8eaf6",  # Texte principal (blanc cassé)
    "subtext": "#8892b0",  # Texte secondaire (gris)
    "border":  "#2e3250",  # Bordures
    "entry":   "#151824",  # Champs de saisie
}

# Polices d'écriture pour une interface moderne et lisible
FONT_TITLE = ("Segoe UI", 22, "bold")  # Titres principaux
FONT_HEAD  = ("Segoe UI", 13, "bold")  # Sous-titres
FONT_LABEL = ("Segoe UI", 10)       # Labels et textes normaux
FONT_SMALL = ("Segoe UI", 9)        # Petits textes
FONT_MONO  = ("Consolas", 10)       # Texte monospace (pour les codes)


# ─────────────────────────────────────────────
#  COMPOSANTS RÉUTILISABLES : Nos briques d'interface
# ─────────────────────────────────────────────

def make_card(parent, title="", **kwargs):
    """Crée une carte stylisée avec bordure et titre optionnel"""
    outer = tk.Frame(parent, bg=COLORS["card"],
                     highlightbackground=COLORS["border"],
                     highlightthickness=1, **kwargs)
    if title:
        tk.Label(outer, text=title, font=FONT_HEAD,
                 bg=COLORS["card"], fg=COLORS["text"]).pack(anchor="w", padx=16, pady=(14, 6))
    return outer


def make_entry(parent, label, row, default=""):
    """Crée un champ de saisie avec son label, déjà stylisé"""
    # Label à gauche
    tk.Label(parent, text=label, font=FONT_LABEL,
             bg=COLORS["card"], fg=COLORS["subtext"]).grid(
        row=row, column=0, sticky="w", padx=(16, 8), pady=4)
    
    # Champ de saisie à droite
    var = tk.StringVar(value=default)
    tk.Entry(parent, textvariable=var, font=FONT_LABEL,
             bg=COLORS["entry"], fg=COLORS["text"],
             insertbackground=COLORS["text"],
             relief="flat", bd=0).grid(
        row=row, column=1, sticky="ew", padx=(0, 16), pady=4, ipady=6)
    return var


def make_button(parent, text, command, color=None, **kwargs):
    """Crée un bouton moderne avec hover et style cohérent"""
    color = color or COLORS["accent"]  # Couleur par défaut
    return tk.Button(parent, text=text, command=command,
                     font=("Segoe UI", 10, "bold"),
                     bg=color, fg="white", relief="flat", bd=0,
                     activebackground=COLORS["bg"], activeforeground=color,
                     cursor="hand2", padx=14, pady=7, **kwargs)


def style_tree(tree):
    """Applique notre style sombre aux tableaux Treeview"""
    s = ttk.Style()
    s.theme_use("clam")  # Theme de base modifiable
    
    # Style des cellules
    s.configure("Custom.Treeview",
                background=COLORS["card"], foreground=COLORS["text"],
                fieldbackground=COLORS["card"], rowheight=30,
                font=FONT_LABEL, borderwidth=0)
    
    # Style des en-têtes
    s.configure("Custom.Treeview.Heading",
                background=COLORS["panel"], foreground=COLORS["accent"],
                font=("Segoe UI", 10, "bold"), relief="flat")
    
    # Style quand on sélectionne une ligne
    s.map("Custom.Treeview",
          background=[("selected", COLORS["accent"])],
          foreground=[("selected", "white")])
    
    tree.configure(style="Custom.Treeview")


# ─────────────────────────────────────────────
#  FENÊTRE PRINCIPALE : Le cœur de notre application
# ─────────────────────────────────────────────

class App(tk.Tk):
    """Fenêtre principale de notre application GestiStock"""
    
    def __init__(self):
        super().__init__()
        self.title("GestiStock — Gestion des Stocks | Misbaou DIALLO")  # Titre de la fenêtre
        self.configure(bg=COLORS["bg"])  # Fond sombre
        self.geometry("1300x800")  # Taille par défaut
        self.minsize(980, 640)  # Taille minimale pour pas tout casser
        
        self.manager = StockManager()  # Notre gestionnaire de données
        self._build_ui()  # Construction de l'interface
        self.after(300, self._check_low_stock_on_start)  # Alertes après 300ms

    def _build_ui(self):
        """Construit toute l'interface graphique"""
        # Barre latérale à gauche
        self.sidebar = tk.Frame(self, bg=COLORS["panel"], width=215)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)  # Largeur fixe

        # Zone principale à droite
        self.content = tk.Frame(self, bg=COLORS["bg"])
        self.content.pack(side="left", fill="both", expand=True)

        self._build_sidebar()  # Construit le menu de navigation

        # Crée toutes les pages (mais n'en affiche qu'une à la fois)
        self.frames = {}
        for F in (ProductsFrame, SalesFrame, SuppliersFrame, ReportsFrame):
            frame = F(self.content, self)
            self.frames[F.__name__] = frame
            frame.place(relx=0, rely=0, relwidth=1, relheight=1)  # Plein écran

        self._show_frame("ProductsFrame")  # Page par défaut au démarrage

    def _build_sidebar(self):
        """Construit la barre latérale avec logo et navigation"""
        # Logo et nom de l'app
        tk.Label(self.sidebar, text="📦", font=("Segoe UI", 32),
                 bg=COLORS["panel"], fg=COLORS["accent"]).pack(pady=(24, 4))
        tk.Label(self.sidebar, text="GestiStock", font=("Segoe UI", 14, "bold"),
                 bg=COLORS["panel"], fg=COLORS["subtext"]).pack(pady=(0, 2))
        tk.Label(self.sidebar, text="Par Misbaou DIALLO", font=("Segoe UI", 9, "italic"),
                 bg=COLORS["panel"], fg=COLORS["subtext"]).pack(pady=(0, 16))

        # Ligne de séparation
        ttk.Separator(self.sidebar, orient="horizontal").pack(fill="x", padx=16)

        # Menu de navigation
        nav_items = [
            ("🗂   Produits",     "ProductsFrame"),
            ("💰   Ventes",       "SalesFrame"),
            ("🏭   Fournisseurs", "SuppliersFrame"),
            ("📊   Rapports",     "ReportsFrame"),
        ]
        self.nav_btns = {}
        
        # Crée un bouton pour chaque page
        for label, fname in nav_items:
            btn = tk.Button(
                self.sidebar, text=label, font=("Segoe UI", 11),
                anchor="w", padx=20, pady=10, bd=0, relief="flat",
                bg=COLORS["panel"], fg=COLORS["text"],
                activebackground=COLORS["accent"], activeforeground="white",
                cursor="hand2",
                command=lambda f=fname: self._show_frame(f),  # Au clic, affiche la page
            )
            btn.pack(fill="x", pady=2, padx=8)
            self.nav_btns[fname] = btn

        tk.Frame(self.sidebar, bg=COLORS["panel"]).pack(expand=True)
        ttk.Separator(self.sidebar, orient="horizontal").pack(fill="x", padx=16)

        btn_area = tk.Frame(self.sidebar, bg=COLORS["panel"])
        btn_area.pack(fill="x", pady=8, padx=8)

        make_button(btn_area, "⬇  Export CSV", self._export_csv,
                    color=COLORS["card"]).pack(fill="x", pady=3)

        if EXCEL_OK:
            make_button(btn_area, "📊  Export Excel (.xlsx)", self._export_excel,
                        color=COLORS["card"]).pack(fill="x", pady=3)
            make_button(btn_area, "📥  Import Excel (.xlsx)", self._import_excel,
                        color=COLORS["card"]).pack(fill="x", pady=3)
        else:
            tk.Label(btn_area, text="Installez openpyxl\npour Excel",
                     font=FONT_SMALL, bg=COLORS["panel"],
                     fg=COLORS["subtext"], justify="center").pack(pady=6)

    def _show_frame(self, name):
        for n, btn in self.nav_btns.items():
            btn.configure(bg=COLORS["accent"] if n == name else COLORS["panel"])
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
#  PAGE PRODUITS : Gestion complète du stock
# ─────────────────────────────────────────────
# C'est la page la plus utilisée de l'application :
# - Affichage de tous les produits avec leur stock
# - Recherche en temps réel
# - Tri par catégorie/nom/quantité
# - Ajout/modification/suppression de produits
# - Alertes visuelles pour les stocks faibles

class ProductsFrame(tk.Frame):
    """
    Page principale de gestion des produits.
    
    Cette frame contient :
    - Un tableau affichant tous les produits avec leurs informations
    - Une barre de recherche pour filtrer rapidement
    - Un formulaire pour ajouter/modifier des produits
    - Des boutons d'action (supprimer, modifier)
    - Des options de tri et des catégories rapides
    - Un système d'alerte pour les stocks faibles
    
    Hérité de tk.Frame pour s'intégrer dans l'interface principale.
    """
    
    def __init__(self, parent, app):
        """
        Initialise la frame des produits.
        
        Args:
            parent: Frame parent (zone de contenu principale)
            app: Instance de l'application principale (pour accéder au manager)
        """
        super().__init__(parent, bg=COLORS["bg"])  # Fond sombre comme toute l'app
        self.app = app  # Référence à l'app pour accéder au gestionnaire
        self._build()  # Construction de l'interface

    def _build(self):
        """
        Construit toute l'interface de la page produits.
        
        Organisation :
        1. En-tête avec titre et barre de recherche
        2. Zone principale avec tableau à gauche et formulaire à droite
        3. Tableau : liste des produits avec tri et actions
        4. Formulaire : champs pour ajouter/modifier un produit
        """
        # ── EN-TÊTE : Titre et recherche ──
        hdr = tk.Frame(self, bg=COLORS["bg"])
        hdr.pack(fill="x", padx=24, pady=(20, 10))  # Marge haute pour l'aération
        
        # Titre principal de la page
        tk.Label(hdr, text="Gestion des Produits", font=FONT_TITLE,
                 bg=COLORS["bg"], fg=COLORS["text"]).pack(side="left")

        # Barre de recherche stylisée à droite
        sf = tk.Frame(hdr, bg=COLORS["card"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        sf.pack(side="right")
        
        # Icône de recherche
        tk.Label(sf, text="🔍", bg=COLORS["card"], fg=COLORS["subtext"]).pack(side="left", padx=8)
        
        # Champ de recherche avec recherche en temps réel
        self.search_var = tk.StringVar()
        # Détecte chaque changement dans le champ et rafraîchit le tableau
        self.search_var.trace_add("write", lambda *_: self.refresh())
        tk.Entry(sf, textvariable=self.search_var, font=FONT_LABEL,
                 bg=COLORS["card"], fg=COLORS["text"],
                 insertbackground=COLORS["text"],
                 relief="flat", bd=0, width=22).pack(side="left", ipady=6, padx=(0, 8))

        # ── ZONE PRINCIPALE : Tableau et formulaire ──
        main = tk.Frame(self, bg=COLORS["bg"])
        main.pack(fill="both", expand=True, padx=24, pady=8)

        # ── TABLEAU DES PRODUITS (à gauche) ──
        tbl = make_card(main, "Stock actuel")  # Carte avec bordure et titre
        tbl.pack(side="left", fill="both", expand=True, padx=(0, 12))

        # Configuration du tableau Treeview
        cols = ("Nom", "Reference", "Quantite", "Prix (EUR)", "Categorie")  # IDs internes
        headers = ("Nom", "Référence", "Quantité", "Prix (€)", "Catégorie")   # Texte affiché
        self.tree = ttk.Treeview(tbl, columns=cols, show="headings", selectmode="browse")
        
        # Configure chaque colonne avec son en-tête et sa largeur
        for c, h in zip(cols, headers):
            self.tree.heading(c, text=h)  # Texte de l'en-tête
            self.tree.column(c, width=110, anchor="center")  # Largeur et alignement
        self.tree.column("Nom", width=180, anchor="w")  # Colonne nom plus large, alignée à gauche
        
        # Applique notre style sombre personnalisé
        style_tree(self.tree)
        
        # Configure les tags pour mettre en évidence les stocks faibles
        self.tree.tag_configure("low", foreground=COLORS["warning"])  # Texte orange pour stocks faibles

        # ── BARRE DE DÉFILEMENT ET SÉLECTION ──
        sb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)  # Lie la scrollbar au tableau
        self.tree.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        sb.pack(side="right", fill="y", pady=8)
        
        # Détecte quand l'utilisateur clique sur une ligne du tableau
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # ── BARRE D'ACTION : Boutons Modifier/Supprimer et tri ──
        btn_row = tk.Frame(tbl, bg=COLORS["card"])
        btn_row.pack(fill="x", padx=8, pady=(0, 10))
        
        # Boutons d'action à droite
        make_button(btn_row, "Supprimer", self._delete,
                    color=COLORS["danger"]).pack(side="right", padx=4)  # Rouge pour danger
        make_button(btn_row, "Modifier", self._update).pack(side="right", padx=4)  # Bleu par défaut

        # Options de tri à gauche
        sort_f = tk.Frame(btn_row, bg=COLORS["card"])
        sort_f.pack(side="left")
        tk.Label(sort_f, text="Trier :", font=FONT_SMALL,
                 bg=COLORS["card"], fg=COLORS["subtext"]).pack(side="left", padx=4)
        
        # Radio buttons pour choisir le critère de tri
        self.sort_var = tk.StringVar(value="category")  # Tri par catégorie par défaut
        for val, lbl in [("category", "Categorie"), ("quantity", "Quantite"), ("name", "Nom")]:
            tk.Radiobutton(sort_f, text=lbl, variable=self.sort_var, value=val,
                           font=FONT_SMALL, bg=COLORS["card"], fg=COLORS["subtext"],
                           selectcolor=COLORS["entry"], activebackground=COLORS["card"],
                           command=self.refresh).pack(side="left")  # Rafraîchit quand on change

        # ── FORMULAIRE (à droite) : Ajout/Modification de produits ──
        form = make_card(main, "Ajouter / Modifier un produit")
        form.pack(side="right", fill="y", ipadx=8, ipady=8)

        # Frame interne pour le formulaire (utilise grid pour alignement)
        # Séparé du titre qui utilise pack
        fi = tk.Frame(form, bg=COLORS["card"])
        fi.pack(fill="both", expand=True)
        fi.columnconfigure(1, weight=1)  # Colonne de droite s'étire automatiquement

        # ── CHAMPS DU FORMULAIRE ──
        self.f_name     = make_entry(fi, "Nom",        0)                    # Ligne 0
        self.f_ref      = make_entry(fi, "Reference",  1)                    # Ligne 1
        self.f_qty      = make_entry(fi, "Quantite",   2, "0")             # Ligne 2, défaut "0"
        self.f_price    = make_entry(fi, "Prix (EUR)", 3, "0.00")          # Ligne 3, défaut "0.00"
        self.f_category = make_entry(fi, "Categorie",  4)                    # Ligne 4

        # ── CATÉGORIES RAPIDES : Boutons pour remplir automatiquement ──
        tk.Label(fi, text="Categories rapides :", font=FONT_SMALL,
                 bg=COLORS["card"], fg=COLORS["subtext"]).grid(
            row=5, column=0, sticky="w", padx=16, pady=(0, 4))
        cat_row = tk.Frame(fi, bg=COLORS["card"])  # Frame pour les boutons de catégories
        cat_row.grid(row=5, column=1, sticky="ew", padx=(0, 16))
        
        # Boutons pour les catégories les plus courantes
        for cat in ["Alimentaire", "Electronique", "Vetements", "Mobilier", "Autre"]:
            tk.Button(cat_row, text=cat, font=FONT_SMALL,
                      bg=COLORS["entry"], fg=COLORS["subtext"],
                      relief="flat", bd=0, cursor="hand2", padx=5, pady=2,
                      command=lambda c=cat: self.f_category.set(c)  # Remplit le champ catégorie
                      ).pack(side="left", padx=2, pady=2)

        # ── BOUTONS D'ACTION PRINCIPAUX ──
        make_button(fi, "Ajouter au stock", self._add).grid(
            row=6, column=0, columnspan=2, padx=16, pady=(12, 4), sticky="ew")
        make_button(fi, "Mettre a jour", self._update,
                    color=COLORS["accent2"]).grid(  # Vert turquoise pour la mise à jour
            row=7, column=0, columnspan=2, padx=16, pady=(0, 16), sticky="ew")

        # ── BADGE D'ALERTE STOCK FAIBLE ──
        self.low_badge = tk.Label(fi, text="", font=FONT_SMALL,
                                   bg=COLORS["card"], fg=COLORS["warning"])  # Orange pour les alertes
        self.low_badge.grid(row=8, column=0, columnspan=2, padx=16, pady=(0, 8))

    def refresh(self):
        """
        Rafraîchit le tableau des produits.
        
        Cette fonction est appelée :
        - Au chargement de la page
        - Après chaque ajout/modification/suppression
        - Quand on change le critère de tri
        - Quand on tape dans la barre de recherche
        
        Elle vide et reconstruit complètement le tableau avec les données à jour.
        """
        # Vide le tableau (supprime toutes les lignes existantes)
        self.tree.delete(*self.tree.get_children())
        
        # Récupère les produits triés selon le critère sélectionné
        products = self.app.manager.get_products(self.sort_var.get())
        
        # Filtre par recherche si du texte est saisi
        q = self.search_var.get().lower()
        if q:
            products = [p for p in products
                        if q in p["name"].lower() or q in p["reference"].lower()]
        
        # Ajoute chaque produit au tableau
        for p in products:
            # Tag "low" si stock faible pour mettre en orange
            tag = "low" if p["quantity"] <= LOW_STOCK_THRESHOLD else ""
            self.tree.insert("", "end", iid=p["reference"],  # ID = référence (unique)
                             values=(p["name"], p["reference"],
                                     p["quantity"], f"{p['price']:.2f}", p["category"]),
                             tags=(tag,))
        
        # Met à jour le badge d'alerte stock faible
        n_low = len(self.app.manager.low_stock_products())
        self.low_badge.configure(
            text=f"ALERTE: {n_low} produit(s) en stock faible" if n_low
                 else "Tous les stocks sont OK"
        )

    def _on_select(self, _e):
        """
        Gère la sélection d'un produit dans le tableau.
        
        Quand l'utilisateur clique sur une ligne :
        - Remplit automatiquement le formulaire avec les infos du produit
        - Permet de modifier facilement le produit sélectionné
        
        Args:
            _e: Événement de sélection (non utilisé mais requis par tkinter)
        """
        sel = self.tree.selection()  # Récupère les lignes sélectionnées
        if not sel:
            return  # Pas de sélection = rien à faire
            
        # Récupère le produit sélectionné via sa référence
        p = self.app.manager._find_product(sel[0])
        if p:
            # Remplit tous les champs du formulaire avec les données du produit
            self.f_name.set(p["name"])        # Nom du produit
            self.f_ref.set(p["reference"])   # Référence (non modifiable)
            self.f_qty.set(p["quantity"])    # Quantité en stock
            self.f_price.set(p["price"])      # Prix unitaire
            self.f_category.set(p["category"]) # Catégorie

    def _add(self):
        """
        Ajoute un nouveau produit au stock.
        
        Récupère les valeurs du formulaire et les valide
        avant d'ajouter le produit à la base de données.
        """
        try:
            # Appelle la méthode du gestionnaire pour ajouter le produit
            self.app.manager.add_product(
                self.f_name.get(),     # Nom du produit
                self.f_ref.get(),      # Référence unique
                self.f_qty.get(),      # Quantité initiale
                self.f_price.get(),    # Prix unitaire
                self.f_category.get()  # Catégorie
            )
            
            # Vide tous les champs du formulaire pour le prochain ajout
            for v in (self.f_name, self.f_ref, self.f_qty, self.f_price, self.f_category):
                v.set("")
                
            self.refresh()  # Rafraîchit le tableau pour voir le nouveau produit
            messagebox.showinfo("Succes", "Produit ajoute avec succes.")
            
        except Exception as e:
            # Affiche l'erreur si quelque chose s'est mal passé
            messagebox.showerror("Erreur", str(e))

    def _update(self):
        """
        Met à jour un produit existant avec les nouvelles valeurs du formulaire.
        
        Le produit doit d'abord être sélectionné dans le tableau,
        puis l'utilisateur peut modifier les champs et cliquer sur "Mettre à jour".
        """
        ref = self.f_ref.get().strip()  # Récupère la référence (champ clé)
        if not ref:
            messagebox.showwarning("Avertissement", "Selectionnez un produit d'abord.")
            return
            
        try:
            # Appelle la méthode du gestionnaire pour mettre à jour le produit
            self.app.manager.update_product(
                ref,  # Référence du produit à modifier (ne change pas)
                name=self.f_name.get().strip(),        # Nouveau nom
                quantity=int(self.f_qty.get()),           # Nouvelle quantité
                price=float(self.f_price.get()),         # Nouveau prix
                category=self.f_category.get().strip(),   # Nouvelle catégorie
            )
            
            self.refresh()  # Rafraîchit le tableau pour voir les changements
            messagebox.showinfo("Succes", "Produit mis a jour.")
            
        except Exception as e:
            # Affiche l'erreur si la mise à jour a échoué
            messagebox.showerror("Erreur", str(e))

    def _delete(self):
        """
        Supprime définitivement un produit du stock.
        
        ⚠️  ACTION DANGEREUSE : Le produit sera complètement supprimé
        avec tout son historique de ventes associé.
        
        Demande une confirmation avant de supprimer pour éviter les erreurs.
        """
        ref = self.f_ref.get().strip()  # Récupère la référence du produit
        if not ref:
            messagebox.showwarning("Avertissement", "Selectionnez un produit d'abord.")
            return
            
        # Demande une confirmation explicite à l'utilisateur
        if messagebox.askyesno("Confirmer", f"Supprimer le produit '{ref}' ?"):
            try:
                # Supprime le produit via le gestionnaire
                self.app.manager.delete_product(ref)
                
                # Vide tous les champs du formulaire
                for v in (self.f_name, self.f_ref, self.f_qty, self.f_price, self.f_category):
                    v.set("")
                    
                self.refresh()  # Rafraîchit le tableau
                
            except Exception as e:
                # Affiche l'erreur si la suppression a échoué
                messagebox.showerror("Erreur", str(e))


# ─────────────────────────────────────────────
#  PAGE VENTES : Enregistrement et historique des transactions
# ─────────────────────────────────────────────
# Cette page permet de :
# - Enregistrer des ventes avec mise à jour automatique du stock
# - Voir l'historique complet des ventes avec dates et montants
# - Sélectionner rapidement des produits depuis une liste
# - Calculer le chiffre d'affaires total
# - Filtrer les ventes par période

class SalesFrame(tk.Frame):
    """
    Page de gestion des ventes et de l'historique des transactions.
    
    Cette frame contient :
    - Un formulaire pour enregistrer rapidement des ventes
    - Une liste des produits disponibles pour sélection rapide
    - Un tableau historique de toutes les ventes avec totaux
    - Un calculateur de chiffre d'affaires en temps réel
    - Des alertes visuelles pour les stocks faibles après vente
    
    Chaque vente met automatiquement à jour le stock du produit.
    """
    
    def __init__(self, parent, app):
        """
        Initialise la frame des ventes.
        
        Args:
            parent: Frame parent (zone de contenu principale)
            app: Instance de l'application principale
        """
        super().__init__(parent, bg=COLORS["bg"])  # Fond sombre cohérent
        self.app = app  # Référence à l'app pour accéder au gestionnaire
        self._build()  # Construction de l'interface

    def _build(self):
        """
        Construit l'interface de la page des ventes.
        
        Organisation :
        1. Titre de la page
        2. Zone principale avec tableau historique à gauche et formulaire à droite
        3. Tableau : historique complet des ventes avec chiffre d'affaires total
        4. Formulaire : enregistrement rapide avec sélection de produits
        """
        # ── TITRE DE LA PAGE ──
        tk.Label(self, text="Suivi des Ventes", font=FONT_TITLE,
                 bg=COLORS["bg"], fg=COLORS["text"]).pack(anchor="w", padx=24, pady=(20, 10))

        # ── ZONE PRINCIPALE ──
        main = tk.Frame(self, bg=COLORS["bg"])
        main.pack(fill="both", expand=True, padx=24, pady=8)

        # ── TABLEAU HISTORIQUE DES VENTES (à gauche) ──
        tbl = make_card(main, "Historique des ventes")
        tbl.pack(side="left", fill="both", expand=True, padx=(0, 12))

        # Configuration du tableau des ventes
        cols = ("Date", "Produit", "Reference", "Quantite", "Prix unit.", "Total (EUR)", "Categorie")
        headers = ("Date", "Produit", "Référence", "Quantité", "Prix unit.", "Total (€)", "Catégorie")
        self.tree = ttk.Treeview(tbl, columns=cols, show="headings")
        
        # Configure chaque colonne
        for c, h in zip(cols, headers):
            self.tree.heading(c, text=h)  # En-tête
            self.tree.column(c, width=95, anchor="center")  # Largeur et alignement
        self.tree.column("Produit", width=150, anchor="w")  # Colonne produit plus large
        
        style_tree(self.tree)  # Applique notre style sombre

        # Barre de défilement
        sb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        sb.pack(side="right", fill="y", pady=8)

        # ── CHIFFRE D'AFFAIRES TOTAL (en bas du tableau) ──
        self.total_lbl = tk.Label(tbl, text="Chiffre d'affaires : 0.00 EUR",
                                   font=("Segoe UI", 11, "bold"),
                                   bg=COLORS["card"], fg=COLORS["accent2"])  # Vert turquoise
        self.total_lbl.pack(anchor="e", padx=16, pady=(0, 8))  # Aligné à droite

        # ── FORMULAIRE D'ENREGISTREMENT (à droite) ──
        form = make_card(main, "Enregistrer une vente")
        form.pack(side="right", fill="y", ipadx=8, ipady=8)

        # Frame interne pour le formulaire (utilise grid)
        fi = tk.Frame(form, bg=COLORS["card"])
        fi.pack(fill="both", expand=True)
        fi.columnconfigure(1, weight=1)  # Colonne de droite s'étire

        # ── CHAMPS DU FORMULAIRE ──
        self.s_ref = make_entry(fi, "Reference produit", 0)      # Ligne 0
        self.s_qty = make_entry(fi, "Quantite vendue",   1, "1")  # Ligne 1, défaut "1"

        # ── SÉLECTION RAPIDE DES PRODUITS ──
        tk.Label(fi, text="Selection rapide :", font=FONT_SMALL,
                 bg=COLORS["card"], fg=COLORS["subtext"]).grid(
            row=2, column=0, sticky="nw", padx=16, pady=4)

        # Liste déroulante des produits disponibles
        pf = tk.Frame(fi, bg=COLORS["entry"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        pf.grid(row=2, column=1, sticky="ew", padx=(0, 16), pady=4)
        
        # Listbox avec tous les produits (référence, nom, quantité)
        self.picker = tk.Listbox(pf, font=FONT_SMALL,
                                  bg=COLORS["entry"], fg=COLORS["text"],
                                  selectbackground=COLORS["accent"],  # Violet au survol
                                  relief="flat", bd=0, height=8,
                                  exportselection=False)  # Permet la sélection multiple
        self.picker.pack(fill="both", expand=True)
        self.picker.bind("<<ListboxSelect>>", self._pick)  # Au clic, remplit la référence

        # ── BOUTON D'ENREGISTREMENT ──
        make_button(fi, "Enregistrer la vente", self._record,
                    color=COLORS["accent2"]).grid(  # Vert turquoise pour les ventes
            row=3, column=0, columnspan=2, padx=16, pady=16, sticky="ew")

    def refresh(self):
        """
        Rafraîchit la page des ventes.
        
        Met à jour :
        - Le tableau historique des ventes
        - Le chiffre d'affaires total
        - La liste des produits disponibles
        - Les alertes visuelles pour les stocks faibles
        """
        # Vide le tableau des ventes
        self.tree.delete(*self.tree.get_children())
        
        # Ajoute toutes les ventes (ordre anti-chronologique : plus récent en premier)
        for s in reversed(self.app.manager.sales):
            self.tree.insert("", "end", values=(
                s["date"], s["product_name"], s["reference"],
                s["quantity"], f"{s['unit_price']:.2f}",
                f"{s['total']:.2f}", s["category"]
            ))
        
        # Met à jour le chiffre d'affaires total avec formatage des milliers
        total = self.app.manager.report_total_revenue()
        self.total_lbl.configure(text=f"Chiffre d'affaires total : {total:,.2f} EUR")

        # ── MISE À JOUR DE LA LISTE DE SÉLECTION RAPIDE ──
        self.picker.delete(0, "end")  # Vide la liste actuelle
        
        # Ajoute tous les produits avec leur stock actuel
        for p in self.app.manager.get_products("name"):  # Tri par nom
            display_text = f"{p['reference']} - {p['name']} (qte: {p['quantity']})"
            self.picker.insert("end", display_text)
            
            # Met en orange les produits avec stock faible
            if p["quantity"] <= LOW_STOCK_THRESHOLD:
                self.picker.itemconfigure("end", fg=COLORS["warning"])

    def _pick(self, _e):
        """
        Gère la sélection d'un produit dans la liste rapide.
        
        Quand l'utilisateur clique sur un produit dans la liste :
        - Extrait la référence du produit
        - Remplit automatiquement le champ de référence du formulaire
        
        Args:
            _e: Événement de sélection (non utilisé mais requis par tkinter)
        """
        sel = self.picker.curselection()  # Récupère l'index sélectionné
        if sel:
            # Extrait la référence (partie avant le premier " - ")
            selected_text = self.picker.get(sel[0])
            reference = selected_text.split(" - ")[0]
            self.s_ref.set(reference)  # Remplit le champ de référence

    def _record(self):
        """
        Enregistre une nouvelle vente.
        
        Processus :
        1. Valide les champs du formulaire
        2. Appelle le gestionnaire pour enregistrer la vente
        3. Met à jour le stock automatiquement
        4. Rafraîchit toutes les interfaces
        5. Affiche des alertes si le stock devient faible
        """
        try:
            # Enregistre la vente via le gestionnaire
            sale = self.app.manager.record_sale(
                self.s_ref.get().strip(),    # Référence du produit
                self.s_qty.get().strip()     # Quantité vendue
            )
            
            # Réinitialise le formulaire pour la prochaine vente
            self.s_ref.set("")   # Vide la référence
            self.s_qty.set("1")   # Remet la quantité par défaut
            
            # Rafraîchit cette page et la page produits (pour voir le stock mis à jour)
            self.refresh()
            self.app.frames["ProductsFrame"].refresh()
            
            # Vérifie si le stock est devenu faible après cette vente
            p = self.app.manager._find_product(sale["reference"])
            if p and p["quantity"] <= LOW_STOCK_THRESHOLD:
                # Alerte orange pour stock faible
                messagebox.showwarning(
                    "Alerte Stock Faible",
                    f"ATTENTION : Stock de '{p['name']}' = {p['quantity']} unite(s) restante(s).\n"
                    f"Pensez a reapprovisionner !"
                )
            else:
                # Message de succès normal
                messagebox.showinfo("Vente enregistree",
                                    f"{sale['quantity']} x {sale['product_name']}\n"
                                    f"Total : {sale['total']:.2f} EUR")
                                    
        except Exception as e:
            # Affiche l'erreur si quelque chose s'est mal passé
            messagebox.showerror("Erreur", str(e))


# ─────────────────────────────────────────────
#  PAGE FOURNISSEURS : Gestion des relations commerciales
# ─────────────────────────────────────────────
# Cette page permet de :
# - Ajouter de nouveaux fournisseurs avec leurs coordonnées complètes
# - Modifier les informations des fournisseurs existants
# - Supprimer des fournisseurs (avec confirmation)
# - Associer des produits à chaque fournisseur
# - Voir tous les fournisseurs triés alphabétiquement

class SuppliersFrame(tk.Frame):
    """
    Page de gestion des fournisseurs et des relations commerciales.
    
    Cette frame contient :
    - Un tableau affichant tous les fournisseurs avec leurs coordonnées
    - Un formulaire pour ajouter/modifier les informations des fournisseurs
    - Un sélecteur de produits pour associer des références aux fournisseurs
    - Des boutons d'action pour modifier et supprimer
    - Une gestion des produits fournis par chaque fournisseur
    
    Les fournisseurs sont essentiels pour la gestion des achats
    et le réapprovisionnement des stocks.
    """
    
    def __init__(self, parent, app):
        """
        Initialise la frame des fournisseurs.
        
        Args:
            parent: Frame parent (zone de contenu principale)
            app: Instance de l'application principale
        """
        super().__init__(parent, bg=COLORS["bg"])  # Fond sombre cohérent
        self.app = app  # Référence à l'app pour accéder au gestionnaire
        self._build()  # Construction de l'interface

    def _build(self):
        """
        Construit l'interface de la page des fournisseurs.
        
        Organisation :
        1. Titre de la page
        2. Zone principale avec tableau à gauche et formulaire à droite
        3. Tableau : liste des fournisseurs avec coordonnées complètes
        4. Formulaire : ajout/modification avec sélecteur de produits
        """
        # ── TITRE DE LA PAGE ──
        tk.Label(self, text="Gestion des Fournisseurs", font=FONT_TITLE,
                 bg=COLORS["bg"], fg=COLORS["text"]).pack(anchor="w", padx=24, pady=(20, 10))

        # ── ZONE PRINCIPALE ──
        main = tk.Frame(self, bg=COLORS["bg"])
        main.pack(fill="both", expand=True, padx=24, pady=8)

        # ── TABLEAU DES FOURNISSEURS (à gauche) ──
        tbl = make_card(main, "Liste des fournisseurs")
        tbl.pack(side="left", fill="both", expand=True, padx=(0, 12))

        # Configuration du tableau des fournisseurs
        cols = ("Nom", "Contact", "Email", "Telephone", "Produits fournis")
        self.tree = ttk.Treeview(tbl, columns=cols, show="headings", selectmode="browse")
        
        # Configure chaque colonne
        for c in cols:
            self.tree.heading(c, text=c)  # En-tête (même nom que la colonne)
            self.tree.column(c, width=130, anchor="w")  # Largeur standard
        self.tree.column("Produits fournis", width=240)  # Colonne plus large pour les produits
        
        style_tree(self.tree)  # Applique notre style sombre

        # Barre de défilement
        sb = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        sb.pack(side="right", fill="y", pady=8)
        
        # Détecte la sélection d'un fournisseur
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # ── BARRE D'ACTION : Modifier/Supprimer ──
        btn_row = tk.Frame(tbl, bg=COLORS["card"])
        btn_row.pack(fill="x", padx=8, pady=(0, 10))
        
        # Boutons d'action à droite
        make_button(btn_row, "Supprimer", self._delete,
                    color=COLORS["danger"]).pack(side="right", padx=4)  # Rouge pour danger
        make_button(btn_row, "Modifier", self._update,
                    color=COLORS["accent3"]).pack(side="right", padx=4)  # Violet clair

        # ── FORMULAIRE (à droite) : Ajout/Modification de fournisseurs ──
        form = make_card(main, "Ajouter un fournisseur")
        form.pack(side="right", fill="y", ipadx=8, ipady=8)

        # Frame interne pour le formulaire
        fi = tk.Frame(form, bg=COLORS["card"])
        fi.pack(fill="both", expand=True)
        fi.columnconfigure(1, weight=1)  # Colonne de droite s'étire

        # ── CHAMPS DU FORMULAIRE ──
        self.f_name    = make_entry(fi, "Nom",              0)  # Nom du fournisseur
        self.f_contact = make_entry(fi, "Contact",          1)  # Personne de contact
        self.f_email   = make_entry(fi, "Email",            2)  # Email (optionnel)
        self.f_phone   = make_entry(fi, "Telephone",        3)  # Téléphone (optionnel)
        self.f_prods   = make_entry(fi, "Produits fournis", 4)  # Produits associés

        # ── INSTRUCTIONS POUR LES PRODUITS ──
        tk.Label(fi, text="(references separees par virgules)", font=FONT_SMALL,
                 bg=COLORS["card"], fg=COLORS["subtext"]).grid(
            row=5, column=1, sticky="w", padx=(0, 16), pady=(0, 8))

        # ── SÉLECTEUR DE RÉFÉRENCES (aide visuelle) ──
        tk.Label(fi, text="Choisir references :", font=FONT_SMALL,
                 bg=COLORS["card"], fg=COLORS["subtext"]).grid(
            row=6, column=0, sticky="nw", padx=16, pady=4)
        
        # Frame pour le sélecteur de références
        pf = tk.Frame(fi, bg=COLORS["entry"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        pf.grid(row=6, column=1, sticky="ew", padx=(0, 16), pady=4)
        
        # Listbox avec tous les produits disponibles (sélection multiple possible)
        self.ref_picker = tk.Listbox(pf, font=FONT_SMALL,
                                      bg=COLORS["entry"], fg=COLORS["text"],
                                      selectbackground=COLORS["accent3"],  # Violet clair au survol
                                      relief="flat", bd=0, height=5,
                                      exportselection=False, selectmode="multiple")  # Sélection multiple
        self.ref_picker.pack(fill="both", expand=True)
        self.ref_picker.bind("<<ListboxSelect>>", self._pick_refs)  # Met à jour le champ produits

        # ── BOUTONS D'ACTION PRINCIPAUX ──
        make_button(fi, "Ajouter le fournisseur", self._add,
                    color=COLORS["accent3"]).grid(  # Violet clair pour les fournisseurs
            row=7, column=0, columnspan=2, padx=16, pady=(12, 4), sticky="ew")
        make_button(fi, "Mettre a jour", self._update,
                    color=COLORS["accent"]).grid(  # Bleu pour la mise à jour
            row=8, column=0, columnspan=2, padx=16, pady=(0, 16), sticky="ew")

    def refresh(self):
        """
        Rafraîchit la page des fournisseurs.
        
        Met à jour :
        - Le tableau des fournisseurs
        - La liste des produits disponibles pour l'association
        """
        # Vide le tableau des fournisseurs
        self.tree.delete(*self.tree.get_children())
        
        # Ajoute tous les fournisseurs triés alphabétiquement
        for s in self.app.manager.get_suppliers():
            self.tree.insert("", "end", iid=s["name"],  # ID = nom du fournisseur
                             values=(s["name"], s["contact"], s.get("email", ""),
                                     s.get("phone", ""),
                                     ", ".join(s.get("products", []))))  # Liste des produits
        
        # Met à jour la liste des produits disponibles pour l'association
        self.ref_picker.delete(0, "end")  # Vide la liste actuelle
        for p in self.app.manager.get_products("name"):  # Tri par nom
            self.ref_picker.insert("end", f"{p['reference']} - {p['name']}")

    def _on_select(self, _e):
        """
        Gère la sélection d'un fournisseur dans le tableau.
        
        Quand l'utilisateur clique sur un fournisseur :
        - Remplit le formulaire avec toutes les informations du fournisseur
        - Permet de modifier facilement le fournisseur sélectionné
        
        Args:
            _e: Événement de sélection (non utilisé mais requis par tkinter)
        """
        sel = self.tree.selection()  # Récupère les lignes sélectionnées
        if not sel:
            return  # Pas de sélection = rien à faire
            
        # Récupère le fournisseur sélectionné via son nom
        s = self.app.manager._find_supplier(sel[0])
        if s:
            # Remplit tous les champs du formulaire avec les données du fournisseur
            self.f_name.set(s["name"])        # Nom du fournisseur
            self.f_contact.set(s["contact"])  # Personne de contact
            self.f_email.set(s.get("email", ""))   # Email (peut être vide)
            self.f_phone.set(s.get("phone", ""))   # Téléphone (peut être vide)
            # Liste des produits fournis (séparés par des virgules)
            self.f_prods.set(", ".join(s.get("products", [])))

    def _pick_refs(self, _e):
        """
        Gère la sélection multiple de références de produits.
        
        Quand l'utilisateur sélectionne des produits dans la liste :
        - Extrait les références des produits sélectionnés
        - Remplit automatiquement le champ "Produits fournis" avec les références
        - Sépare les références par des virgules
        
        Args:
            _e: Événement de sélection (non utilisé mais requis par tkinter)
        """
        # Récupère toutes les sélections et extrait les références
        selected = [self.ref_picker.get(i).split(" - ")[0]  # Extrait la partie avant " - "
                    for i in self.ref_picker.curselection()]
        
        # Remplit le champ avec les références séparées par des virgules
        self.f_prods.set(", ".join(selected))

    def _add(self):
        """
        Ajoute un nouveau fournisseur.
        
        Récupère les valeurs du formulaire, nettoie la liste de produits,
        et ajoute le fournisseur à la base de données.
        """
        try:
            # Nettoie la liste des produits (sépare par virgules et enlève les espaces)
            prods = [p.strip() for p in self.f_prods.get().split(",") if p.strip()]
            
            # Appelle la méthode du gestionnaire pour ajouter le fournisseur
            self.app.manager.add_supplier(
                self.f_name.get(),      # Nom du fournisseur
                self.f_contact.get(),   # Personne de contact
                self.f_email.get(),     # Email (peut être vide)
                self.f_phone.get(),     # Téléphone (peut être vide)
                prods                   # Liste des produits fournis
            )
            
            # Vide tous les champs du formulaire pour le prochain ajout
            for v in (self.f_name, self.f_contact, self.f_email, self.f_phone, self.f_prods):
                v.set("")
                
            self.refresh()  # Rafraîchit le tableau pour voir le nouveau fournisseur
            messagebox.showinfo("Succes", "Fournisseur ajoute.")
            
        except Exception as e:
            # Affiche l'erreur si quelque chose s'est mal passé
            messagebox.showerror("Erreur", str(e))

    def _update(self):
        """
        Met à jour un fournisseur existant.
        
        Le fournisseur doit d'abord être sélectionné dans le tableau,
        puis l'utilisateur peut modifier les champs et cliquer sur "Mettre à jour".
        """
        sel = self.tree.selection()
        original = sel[0] if sel else self.f_name.get().strip()  # Nom original du fournisseur
        
        if not original:
            messagebox.showwarning("Avertissement", "Selectionnez un fournisseur d'abord.")
            return
            
        try:
            # Nettoie la liste des produits
            prods = [p.strip() for p in self.f_prods.get().split(",") if p.strip()]
            
            # Appelle la méthode du gestionnaire pour mettre à jour le fournisseur
            self.app.manager.update_supplier(
                original,  # Nom original pour trouver le fournisseur
                name=self.f_name.get().strip(),        # Nouveau nom
                contact=self.f_contact.get().strip(),  # Nouveau contact
                email=self.f_email.get().strip(),     # Nouvel email
                phone=self.f_phone.get().strip(),     # Nouveau téléphone
                products=prods,                         # Nouvelle liste de produits
            )
            
            self.refresh()  # Rafraîchit le tableau pour voir les changements
            messagebox.showinfo("Succes", "Fournisseur mis a jour.")
            
        except Exception as e:
            # Affiche l'erreur si la mise à jour a échoué
            messagebox.showerror("Erreur", str(e))

    def _delete(self):
        """
        Supprime définitivement un fournisseur.
        
        ⚠️  ACTION DANGEREUSE : Le fournisseur sera complètement supprimé.
        Les produits associés ne sont PAS supprimés, seulement la relation est perdue.
        
        Demande une confirmation avant de supprimer.
        """
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Avertissement", "Selectionnez un fournisseur d'abord.")
            return
            
        # Demande une confirmation explicite
        if messagebox.askyesno("Confirmer", f"Supprimer le fournisseur '{sel[0]}' ?"):
            try:
                # Supprime le fournisseur via le gestionnaire
                self.app.manager.delete_supplier(sel[0])
                
                # Vide tous les champs du formulaire
                for v in (self.f_name, self.f_contact, self.f_email, self.f_phone, self.f_prods):
                    v.set("")
                    
                self.refresh()  # Rafraîchit le tableau
                
            except Exception as e:
                # Affiche l'erreur si la suppression a échoué
                messagebox.showerror("Erreur", str(e))


# ─────────────────────────────────────────────
#  PAGE RAPPORTS : Statistiques et analyses commerciales
# ─────────────────────────────────────────────
# Cette page permet de :
# - Voir les produits les plus vendus avec chiffre d'affaires
# - Analyser le chiffre d'affaires par catégorie de produits
# - Surveiller les niveaux de stock (plus haut et plus bas)
# - Consulter les alertes de stock faible avec statuts détaillés
# - Filtrer tous les rapports par période de temps

class ReportsFrame(tk.Frame):
    """
    Page de rapports et statistiques commerciales.
    
    Cette frame contient :
    - Filtres par période pour analyser des créneaux temporels spécifiques
    - Top 5 des produits les plus vendus avec revenus générés
    - Chiffre d'affaires ventilé par catégorie de produits
    - Niveaux de stock extrêmes (risques de rupture et surstocks)
    - Alertes détaillées pour les produits en stock critique
    
    Les rapports s'actualisent automatiquement et peuvent être filtrés
    par date pour des analyses périodiques (mensuel, trimestriel, etc.).
    """
    
    def __init__(self, parent, app):
        """
        Initialise la frame des rapports.
        
        Args:
            parent: Frame parent (zone de contenu principale)
            app: Instance de l'application principale
        """
        super().__init__(parent, bg=COLORS["bg"])  # Fond sombre cohérent
        self.app = app  # Référence à l'app pour accéder au gestionnaire
        self._build()  # Construction de l'interface

    def _build(self):
        """
        Construit l'interface de la page des rapports.
        
        Organisation :
        1. Titre de la page
        2. Filtres par période (dates de début/fin)
        3. Zone principale avec 3 cartes de rapports
        4. Zone d'alertes pour les stocks faibles
        5. Chiffre d'affaires total en bas
        """
        # ── TITRE DE LA PAGE ──
        tk.Label(self, text="Rapports & Statistiques", font=FONT_TITLE,
                 bg=COLORS["bg"], fg=COLORS["text"]).pack(anchor="w", padx=24, pady=(20, 10))

        # ── FILTRES PAR PÉRIODE ──
        flt = tk.Frame(self, bg=COLORS["bg"])
        flt.pack(fill="x", padx=24, pady=(0, 12))
        
        tk.Label(flt, text="Periode :", font=FONT_LABEL,
                 bg=COLORS["bg"], fg=COLORS["subtext"]).pack(side="left")
        
        # Champs de dates pour filtrer les rapports
        self.from_var = tk.StringVar()  # Date de début
        self.to_var   = tk.StringVar()  # Date de fin
        
        for lbl, var in [("De :", self.from_var), ("A :", self.to_var)]:
            tk.Label(flt, text=lbl, font=FONT_LABEL,
                     bg=COLORS["bg"], fg=COLORS["subtext"]).pack(side="left", padx=(12, 4))
            tk.Entry(flt, textvariable=var, font=FONT_LABEL, width=12,
                     bg=COLORS["entry"], fg=COLORS["text"],
                     insertbackground=COLORS["text"],
                     relief="flat", bd=0).pack(side="left", ipady=5)
        
        # Instructions pour le format des dates
        tk.Label(flt, text="(AAAA-MM-JJ)", font=FONT_SMALL,
                 bg=COLORS["bg"], fg=COLORS["subtext"]).pack(side="left", padx=8)
        
        # Bouton pour rafraîchir les rapports avec les filtres
        make_button(flt, "Actualiser", self.refresh).pack(side="left", padx=8)

        # ── ZONE PRINCIPALE : 3 CARTES DE RAPPORTS ──
        row1 = tk.Frame(self, bg=COLORS["bg"])
        row1.pack(fill="both", expand=True, padx=24, pady=(0, 8))

        # Fonction helper pour créer des cartes avec texte monospace
        def text_card(parent, title, pad=(0, 8)):
            """
            Crée une carte contenant un champ de texte monospace.
            
            Args:
                parent: Frame parent
                title: Titre de la carte
                pad: Padding horizontal
                
            Returns:
                tk.Text: Champ de texte configuré
            """
            card = make_card(parent, title)
            card.pack(side="left", fill="both", expand=True, padx=pad)
            txt = tk.Text(card, font=FONT_MONO, wrap="word",
                          bg=COLORS["card"], fg=COLORS["text"],
                          relief="flat", bd=0, state="disabled", height=9)  # Non éditable
            txt.pack(fill="both", expand=True, padx=12, pady=8)
            return txt

        # Crée les 3 cartes de rapports principaux
        self.top_text   = text_card(row1, "Top produits vendus",  pad=(0, 8))  # Produits les plus vendus
        self.cat_text   = text_card(row1, "CA par categorie",      pad=(8, 8))  # Chiffre d'affaires par catégorie
        self.stock_text = text_card(row1, "Niveaux de stock",      pad=(8, 0))  # Niveaux de stock extrêmes

        # ── ZONE D'ALERTE : Stocks faibles ──
        row2 = tk.Frame(self, bg=COLORS["bg"])
        row2.pack(fill="x", padx=24, pady=(0, 8))

        alert_card = make_card(row2, "Alertes stock faible")
        alert_card.pack(side="left", fill="both", expand=True)
        
        # Champ de texte pour les alertes (texte orange pour attirer l'attention)
        self.alert_text = tk.Text(alert_card, font=FONT_MONO,
                                   bg=COLORS["card"], fg=COLORS["warning"],  # Texte orange
                                   relief="flat", bd=0, state="disabled", height=3)
        self.alert_text.pack(fill="both", expand=True, padx=12, pady=8)

        # ── CHIFFRE D'AFFAIRES TOTAL (en bas) ──
        self.rev_lbl = tk.Label(self, text="", font=("Segoe UI", 15, "bold"),
                                 bg=COLORS["bg"], fg=COLORS["accent2"])  # Vert turquoise
        self.rev_lbl.pack(pady=(0, 10))

    def refresh(self):
        """
        Rafraîchit tous les rapports avec les filtres de période.
        
        Met à jour :
        - Le chiffre d'affaires total
        - Le top 5 des produits les plus vendus
        - Le chiffre d'affaires par catégorie
        - Les niveaux de stock extrêmes
        - Les alertes de stock faible
        """
        # Récupère les dates de filtre (vides si non spécifiées)
        fd = self.from_var.get().strip() or None  # Date de début
        td = self.to_var.get().strip() or None    # Date de fin

        # ── CHIFFRE D'AFFAIRES TOTAL ──
        total = self.app.manager.report_total_revenue(fd, td)
        self.rev_lbl.configure(text=f"Chiffre d'affaires total : {total:,.2f} EUR")

        # ── TOP PRODUITS VENDUS ──
        top = self.app.manager.report_top_products(from_date=fd, to_date=td)
        top_text = "\n".join(f"{i+1}. {info['name']}\n"
                              f"   Qte: {info['qty']}  |  CA: {info['revenue']:.2f} EUR\n"
                              for i, (_, info) in enumerate(top)) or "(Aucune vente)"
        self._write(self.top_text, top_text)

        # ── CHIFFRE D'AFFAIRES PAR CATÉGORIE ──
        cats = self.app.manager.report_revenue_by_category(fd, td)
        cat_text = "\n".join(f"{cat:<20} {rev:>10.2f} EUR"
                              for cat, rev in sorted(cats.items(), key=lambda x: -x[1])
                              ) or "(Aucune vente)"
        self._write(self.cat_text, cat_text)

        # ── NIVEAUX DE STOCK ──
        low, high = self.app.manager.report_stock_levels()
        stock_text = "-- Stocks les plus bas --\n"
        stock_text += "\n".join(f"  {p['name']:<22} {p['quantity']:>4} unites" for p in low)
        stock_text += "\n\n-- Stocks les plus hauts --\n"
        stock_text += "\n".join(f"  {p['name']:<22} {p['quantity']:>4} unites" for p in high)
        self._write(self.stock_text, stock_text or "(Aucun produit)")

        # ── ALERTES STOCK FAIBLE ──
        low_list = self.app.manager.low_stock_products()
        if low_list:
            alert = "\n".join(
                f"  {'[RUPTURE]' if p['quantity']==0 else '[FAIBLE]'} "
                f"{p['name']} ({p['reference']}) -- {p['quantity']} restant(s)"
                for p in sorted(low_list, key=lambda x: x["quantity"])  # Tri par quantité croissante
            )
        else:
            alert = "  OK  Tous les produits ont un stock suffisant."
        self._write(self.alert_text, alert)

    @staticmethod
    def _write(widget, text):
        """
        Écrit du texte dans un widget Text temporairement activé.
        
        Utilisé pour mettre à jour les champs de texte des rapports
        qui sont normalement en lecture seule (state="disabled").
        
        Args:
            widget: Widget tk.Text à mettre à jour
            text: Texte à écrire
        """
        widget.configure(state="normal")      # Active temporairement l'édition
        widget.delete("1.0", "end")           # Vide le contenu actuel
        widget.insert("end", text)             # Insère le nouveau texte
        widget.configure(state="disabled")    # Remet en lecture seule


# ─────────────────────────────────────────────
#  POINT D'ENTRÉE DE L'APPLICATION
# ─────────────────────────────────────────────
# Ce code ne s'exécute que lorsque le fichier est lancé directement
# (pas quand il est importé comme module)

if __name__ == "__main__":
    """
    Point d'entrée principal de l'application GestiStock.
    
    Ce code ne s'exécute que lorsque le fichier est lancé directement
    avec `python gestion_stocks.py`. Il ne s'exécute pas si le fichier
    est importé comme module dans un autre script.
    
    Processus de démarrage :
    1. Crée le dossier data/ s'il n'existe pas
    2. Lance l'interface graphique principale
    3. Démarre la boucle d'événements Tkinter
    """
    ensure_data_dir()  # S'assure que le dossier de données existe
    app = App()         # Crée et configure la fenêtre principale
    app.mainloop()    # Démarre la boucle d'événements de l'interface