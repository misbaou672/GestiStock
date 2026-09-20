import os
import sys
import time
import tkinter.messagebox as msgbox

# Prevent modal popup from blocking during automatic capture
msgbox.showwarning = lambda title, message: None
msgbox.showinfo = lambda title, message: None
msgbox.showerror = lambda title, message: None

from gestion_stocks import App

def take_screenshots():
    app = App()
    app.geometry("1400x850+50+50")
    
    # Disable after call so it doesn't pop up during test
    app.after_cancel(app._check_low_stock_on_start) if hasattr(app, '_check_low_stock_on_start') else None
    
    app.update_idletasks()
    app.update()
    time.sleep(1)
    
    frames = [
        ("ProductsFrame", "gestistock-1.jpg"),
        ("SalesFrame", "gestistock-2.jpg"),
        ("SuppliersFrame", "gestistock-3.jpg"),
        ("ReportsFrame", "gestistock-4.jpg")
    ]
    
    for fname, outfile in frames:
        print(f"Showing {fname}...")
        app._show_frame(fname)
        app.update_idletasks()
        app.update()
        time.sleep(1.5) # wait for render
        outpath = f"/home/msb/.gemini/antigravity/scratch/portfolio-react/src/components/projects/media/{outfile}"
        print(f"Capturing to {outpath}...")
        os.system(f"import -window root {outpath}")
        print(f"Captured {outfile}")
    
    # Also save gestistock.jpg as gestistock-1.jpg fallback
    os.system("cp /home/msb/.gemini/antigravity/scratch/portfolio-react/src/components/projects/media/gestistock-1.jpg /home/msb/.gemini/antigravity/scratch/portfolio-react/src/components/projects/media/gestistock.jpg")
    print("Captures complete!")
    app.destroy()

if __name__ == "__main__":
    take_screenshots()

