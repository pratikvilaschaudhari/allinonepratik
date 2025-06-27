import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import pandas as pd
import glob
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

if getattr(sys, 'frozen', False):
    APP_ROOT = sys._MEIPASS
else:
    APP_ROOT = os.path.abspath(os.path.dirname(__file__))
SALES_DIR = os.path.join(APP_ROOT, 'sales data')
IMAGE_DIR = os.path.join(APP_ROOT, 'images')
LOGO_PATH = os.path.join(APP_ROOT, 'Lazera Logo-02.png')
IMAGE_DISPLAY_SIZE = (100, 100)
FONT = ("Segoe UI", 10)
HEADER_FONT = ("Segoe UI", 10, "bold")

def get_latest_files(directory, pattern, count=5):
    files = glob.glob(os.path.join(directory, pattern))
    if 'salesdata' in pattern.lower():
        def get_num(p):
            name = os.path.basename(p)
            num = ''.join(filter(str.isdigit, name))
            return int(num) if num.isdigit() else -1
        files = [f for f in files if 'salesdata' in os.path.basename(f).lower()]
        files.sort(key=get_num)
    else:
        files.sort(key=os.path.getmtime)
    return files[-count:]

def load_sales_data():
    frames = []
    for i, path in enumerate(get_latest_files(SALES_DIR, 'salesdata*.xlsx', 5), 1):
        df = pd.read_excel(path)
        df.columns = df.columns.str.lower().str.replace(' ', '_')
        df.rename(columns={'colour':'color', 'quantity':'qty'}, inplace=True)
        if all(c in df.columns for c in ['article','store','color','size','qty','asp']):
            sub = df[['article','store','color','size','qty','asp']].copy()
            sub['week'] = f'Week {i}'
            frames.append(sub)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=['article','store','color','size','qty','asp','week'])

def export_to_excel_with_images(app):
    if not hasattr(app, 'df_pivot') or not app.articles:
        messagebox.showwarning("No Data", "No data to export.")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    # Add a placeholder column for "Image" (for reference)
    export_df = app.df_pivot.copy()
    export_df.insert(0, 'Image', 'See image above')

    # Export data to Excel
    export_df.to_excel(file_path, index=True)

    # Open the Excel file with openpyxl
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # Insert images as floating objects
    for row_idx, article in enumerate(app.articles, 2):  # Start from row 2 (header is row 1)
        img_path = next((os.path.join(IMAGE_DIR, f"{article}{ext}") for ext in ['.jpg', '.jpeg', '.png'] if os.path.exists(os.path.join(IMAGE_DIR, f"{article}{ext}"))), None)
        if not img_path:
            img_path = LOGO_PATH

        img = Image.open(img_path)
        img.thumbnail(IMAGE_DISPLAY_SIZE)
        temp_img_path = os.path.join(IMAGE_DIR, f"temp_{article}.png")
        img.save(temp_img_path, "PNG")  # Save as PNG (required by openpyxl)

        xl_img = OpenpyxlImage(temp_img_path)
        ws.add_image(xl_img, f'B{row_idx}')  # Place image in column B, current row

        # Clean up
        if os.path.exists(temp_img_path):
            os.remove(temp_img_path)

    # Save the workbook
    wb.save(file_path)
    messagebox.showinfo("Success", "Data and images exported to Excel successfully!")

def export_to_pdf_with_images(app):
    if not hasattr(app, 'df_pivot') or not app.articles:
        messagebox.showwarning("No Data", "No data to export.")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if not file_path:
        return

    # Prepare data for PDF
    export_df = app.df_pivot.copy()
    export_df.insert(0, 'Image', 'See image below')

    with PdfPages(file_path) as pdf:
        for article in app.articles:
            # Create a figure for each article
            fig = plt.figure(figsize=(11, 2))
            ax = fig.add_subplot(111)

            # Hide axes
            ax.axis('off')

            # Table data
            row_data = [[article] + [export_df.loc[article, f'Week {i}'] if f'Week {i}' in export_df.columns else 0 for i in range(1,6)]]
            col_labels = ['Article'] + [f'Week {i}' for i in range(1,6)]

            # Add table
            table = ax.table(cellText=row_data, colLabels=col_labels, loc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1.2, 1.2)

            # Add image
            img_path = next((os.path.join(IMAGE_DIR, f"{article}{ext}") for ext in ['.jpg', '.jpeg', '.png'] if os.path.exists(os.path.join(IMAGE_DIR, f"{article}{ext}"))), None)
            if not img_path:
                img_path = LOGO_PATH

            img = Image.open(img_path)
            img.thumbnail((200, 200))
            temp_img_path = os.path.join(IMAGE_DIR, f"temp_{article}_pdf.png")
            img.save(temp_img_path, "PNG")

            # Place image above or beside table (this is a simple example; adjust as needed)
            # For a more advanced layout, you might need to use reportlab or borb
            img_ax = fig.add_axes([0.1, 0.6, 0.3, 0.3])
            img_ax.imshow(img)
            img_ax.axis('off')

            # Save page
            pdf.savefig(fig)
            plt.close(fig)

            # Clean up
            if os.path.exists(temp_img_path):
                os.remove(temp_img_path)

    messagebox.showinfo("Success", "Data and images exported to PDF successfully!")

class ArticleSalesApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Article Sales | Lazera Shoes")
        self.state('zoomed')
        self.configure(bg="#f0f4f8")

        # Data
        self.df = load_sales_data()
        self.df_pivot = self._prepare_pivot()
        self.articles = self._get_sorted_articles()

        # UI
        self._build_ui()

    def _prepare_pivot(self):
        if self.df.empty:
            return pd.DataFrame()
        pivot = self.df.pivot_table(index='article', columns='week', values='qty', aggfunc='sum').fillna(0)
        return pivot

    def _get_sorted_articles(self):
        if self.df_pivot.empty:
            return []
        self.df_pivot['Total'] = self.df_pivot.sum(axis=1)
        self.df_pivot = self.df_pivot.sort_values('Total', ascending=False)
        return self.df_pivot.index.unique().tolist()

    def _build_ui(self):
        # Top controls
        top = tk.Frame(self, bg="#e3f2fd")
        top.pack(fill='x', padx=10, pady=8)
        tk.Button(top, text="Export to Excel with Images", command=lambda: export_to_excel_with_images(self), bg="#1976d2", fg="white", font=FONT).pack(side='left', padx=5)
        tk.Button(top, text="Export to PDF with Images", command=lambda: export_to_pdf_with_images(self), bg="#f44336", fg="white", font=FONT).pack(side='left', padx=5)

        # Table frame
        table_frame = tk.Frame(self, bg="#e3f2fd")
        table_frame.pack(fill='both', expand=True, padx=10, pady=5)

        # Canvas and scrollbar
        canvas = tk.Canvas(table_frame, bg="#e3f2fd")
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#e3f2fd")

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Table headers
        headers = ["Article", "Photo", "Week 1", "Week 2", "Week 3", "Week 4", "Week 5"]
        for i, h in enumerate(headers):
            tk.Label(scrollable_frame, text=h, font=HEADER_FONT, bg="#bbdefb", padx=10, pady=5).grid(row=0, column=i, sticky="ew")

        # Table data
        for row_idx, article in enumerate(self.articles, 1):
            # Article row frame (for border)
            row_frame = tk.Frame(scrollable_frame, bg="#e3f2fd", highlightbackground="#999", highlightthickness=3)
            row_frame.grid(row=row_idx, column=0, columnspan=len(headers), sticky="ew")
            row_frame.grid_rowconfigure(0, weight=1)
            for col in range(len(headers)):
                row_frame.grid_columnconfigure(col, weight=1)

            # Article
            lbl = tk.Label(row_frame, text=article, font=FONT, bg="#e3f2fd", padx=10, pady=5)
            lbl.grid(row=0, column=0, sticky="w")

            # Photo
            img_frame = tk.Frame(row_frame, bg="#e3f2fd")
            img_frame.grid(row=0, column=1, padx=5, pady=5)
            img_path = next((os.path.join(IMAGE_DIR, f"{article}{ext}") for ext in ['.jpg', '.jpeg', '.png'] if os.path.exists(os.path.join(IMAGE_DIR, f"{article}{ext}"))), None)
            img = Image.open(img_path if img_path else LOGO_PATH)
            img.thumbnail(IMAGE_DISPLAY_SIZE)
            img = ImageTk.PhotoImage(img)
            img_label = tk.Label(img_frame, image=img, bg="#e3f2fd")
            img_label.image = img
            img_label.pack()

            # Week sales
            for week_idx, week in enumerate([f"Week {i}" for i in range(1,6)], 2):
                qty = self.df_pivot.loc[article, week] if week in self.df_pivot.columns else 0
                lbl = tk.Label(row_frame, text=f"{qty:.0f}", font=FONT, bg="#e3f2fd", padx=10, pady=5)
                lbl.grid(row=0, column=week_idx, sticky="e")

if __name__ == '__main__':
    app = ArticleSalesApp()
    app.mainloop()
