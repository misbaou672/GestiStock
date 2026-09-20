import os
from PIL import Image, ImageDraw

ICON_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")
os.makedirs(ICON_DIR, exist_ok=True)

def create_icon(name, draw_fn):
    img = Image.new("RGBA", (20, 20), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw_fn(draw)
    img.save(os.path.join(ICON_DIR, f"{name}.png"))

# 1. Products / Catalogue Icon (Grid / Box)
def draw_products(d):
    c = "#475569"
    d.rectangle([2, 2, 8, 8], outline=c, width=2)
    d.rectangle([11, 2, 17, 8], outline=c, width=2)
    d.rectangle([2, 11, 8, 17], outline=c, width=2)
    d.rectangle([11, 11, 17, 17], outline=c, width=2)

# 2. Sales / Ventes Icon (Trending Chart)
def draw_sales(d):
    c = "#475569"
    d.line([(3, 16), (8, 11), (12, 14), (17, 4)], fill=c, width=2)
    d.line([(12, 4), (17, 4), (17, 9)], fill=c, width=2)

# 3. Suppliers / Fournisseurs Icon (Building / Users)
def draw_suppliers(d):
    c = "#475569"
    d.rectangle([4, 4, 16, 17], outline=c, width=2)
    d.rectangle([7, 7, 9, 9], fill=c)
    d.rectangle([11, 7, 13, 9], fill=c)
    d.rectangle([7, 12, 9, 14], fill=c)
    d.rectangle([11, 12, 13, 14], fill=c)

# 4. Reports / Rapports Icon (Bar Chart)
def draw_reports(d):
    c = "#475569"
    d.rectangle([3, 10, 6, 17], fill=c)
    d.rectangle([8, 5, 11, 17], fill=c)
    d.rectangle([13, 2, 16, 17], fill=c)

# 5. File / Export CSV Icon
def draw_csv(d):
    c = "#475569"
    d.rectangle([4, 2, 15, 17], outline=c, width=2)
    d.line([(7, 7), (12, 7)], fill=c, width=2)
    d.line([(7, 11), (12, 11)], fill=c, width=2)
    d.line([(7, 14), (10, 14)], fill=c, width=2)

# 6. Excel Icon
def draw_excel(d):
    c = "#475569"
    d.rectangle([3, 2, 16, 17], outline=c, width=2)
    d.line([(3, 7), (16, 7)], fill=c, width=1)
    d.line([(3, 12), (16, 12)], fill=c, width=1)
    d.line([(9, 2), (9, 17)], fill=c, width=1)

create_icon("products", draw_products)
create_icon("sales", draw_sales)
create_icon("suppliers", draw_suppliers)
create_icon("reports", draw_reports)
create_icon("csv", draw_csv)
create_icon("excel", draw_excel)

print("Light mode PNG icons generated.")
