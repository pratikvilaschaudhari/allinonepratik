import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import pandas as pd
import logging
import glob

if getattr(sys, 'frozen', False):
    APP_ROOT = os.path.dirname(sys.executable)
else:
    APP_ROOT = os.path.abspath(os.path.dirname(__file__))

SALES_DIR = os.path.join(APP_ROOT, 'sales data')
INVENTORY_DIR = os.path.join(APP_ROOT, 'inventory data')
PENDING_DIR = os.path.join(APP_ROOT, 'pending orders')
IMAGE_DIR = os.path.join(APP_ROOT, 'images')
LOGO_PATH = os.path.join(APP_ROOT, 'Lazera Logo-02.png')
ERROR_LOG_PATH = os.path.join(APP_ROOT, 'app_code', 'error_log.txt')
MRP_FIXED = 1999
WEEKS = [f'Week {i}' for i in range(1,6)] + ['Overall']
IMAGE_DISPLAY_SIZE = (280, 280)
LOGO_DISPLAY_SIZE = (180, 240)
FONT = ("Segoe UI", 10)
HEADER_FONT = ("Segoe UI", 10, "bold")

os.makedirs(os.path.dirname(ERROR_LOG_PATH), exist_ok=True)
logging.basicConfig(
    filename=ERROR_LOG_PATH,
    filemode='w',
    format='%(asctime)s %(levelname)s: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    level=logging.ERROR
)

def get_latest_files(directory, pattern, count=5):
    files = glob.glob(os.path.join(directory, pattern))
    if 'salesdata' in pattern.lower():
        def get_num(p):
            name = os.path.basename(p)
            num = ''.join(filter(str.isdigit, name))
            return int(num if num.isdigit() else -1)
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
        df.rename(columns={'colour':'color','quantity':'qty'}, inplace=True)
        if all(c in df.columns for c in ['article','store','color','size','qty','asp']):
            sub = df[['article','store','color','size','qty','asp']].copy()
            sub['week'] = f'Week {i}'
            frames.append(sub)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=['article','store','color','size','qty','asp','week'])

def load_inventory_data():
    files = get_latest_files(INVENTORY_DIR, '*.xlsx', 1)
    if not files:
        return pd.DataFrame(columns=['article','store','color','size','soh'])
    df = pd.read_excel(files[0])
    df.columns = df.columns.str.lower().str.replace(' ', '_')
    df.rename(columns={'quantity':'soh','colour':'color'}, inplace=True)
    if all(c in df.columns for c in ['article','store','color','size','soh']):
        return df[['article','store','color','size','soh']].copy()
    return pd.DataFrame(columns=['article','store','color','size','soh'])

def load_pending_data():
    path = os.path.join(PENDING_DIR, 'PENDING ORDERS.xlsx')
    if not os.path.exists(path):
        return pd.DataFrame(columns=['article','color','size','pending_qty','mrp'])
    df = pd.read_excel(path)
    df.columns = df.columns.str.lower().str.replace(' ', '_')
    rename_map = {'colour':'color', 'quantity':'pending_qty'}
    df.rename(columns=rename_map, inplace=True)
    for col in ['color','size','pending_qty','mrp','article']:
        if col not in df.columns:
            df[col] = 0 if col in ['pending_qty','mrp'] else ''
    return df[['article','color','size','pending_qty','mrp']].copy()

def calculate_asp_map(df):
    return {art: (g['asp']*g['qty']).sum()/g['qty'].sum() if g['qty'].sum()>0 else 0
            for art, g in df.groupby('article')}

def merge_data(sales_df, inv_df, asp_map):
    merged = pd.merge(sales_df, inv_df, on=['article','store','color','size'], how='outer')
    merged['soh'] = merged['soh'].fillna(0)
    merged['qty'] = merged['qty'].fillna(0)
    merged['asp'] = merged['asp'].fillna(0)
    merged['asp_calc'] = merged['article'].map(asp_map)
    merged['week'] = merged['week'].fillna('Overall')
    return merged

class AllInOneApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('AllInOne App | Lazera Shoes')
        self.state('zoomed')
        self.configure(bg="#f0f4f8")

        sales = load_sales_data()
        inv = load_inventory_data()
        pending = load_pending_data()

        self.asp_map = calculate_asp_map(sales)
        self.data = merge_data(sales, inv, self.asp_map)
        self.inv_data = inv
        self.pending_data = pending

        self.total_qty = self.data.groupby('article')['qty'].sum().to_dict()
        self.week_qty = self.data.groupby('week')['qty'].sum().to_dict()
        self.article_week_qty = self.data.groupby(['article','week'])['qty'].sum().to_dict()
        self.inv_map = inv.groupby('article')['soh'].sum().to_dict()
        self.store_inv = inv.groupby(['article','store'])['soh'].sum().to_dict()
        self.color_inv = inv.groupby(['article','color'])['soh'].sum().to_dict()
        self.size_inv = inv.groupby(['article','size'])['soh'].sum().to_dict()
        self.pending_total = pending.groupby('article')['pending_qty'].sum().to_dict()
        self.pending_color = pending.groupby(['article','color'])['pending_qty'].sum().to_dict()
        self.pending_size = pending.groupby(['article','size'])['pending_qty'].sum().to_dict()
        self.pending_colorsize = pending.groupby(['article','color','size'])['pending_qty'].sum().to_dict()
        self.mrp_map = pending.set_index('article')['mrp'].to_dict()

        self.overview = True
        self.articles = sorted(self.total_qty, key=self.total_qty.get, reverse=True)
        self.idx = 0
        self.week = 'Overall'

        self._build_ui()
        self._show()

    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('Treeview', rowheight=24, font=FONT, fieldbackground='#e3f2fd', background='#e3f2fd')
        style.configure('Treeview.Heading', font=HEADER_FONT, background='#1976d2', foreground='white')
        style.configure('Bold.TLabelframe', borderwidth=3, relief='solid', background='#e3f2fd')
        style.configure('Bold.TLabelframe.Label', font=HEADER_FONT, background='#1976d2', foreground='white')
        style.configure('PrevNext.TButton', background='#212121', foreground='#fff', font=HEADER_FONT, borderwidth=2)
        style.map('PrevNext.TButton', background=[('active', '#424242')])
        style.configure('FirstLast.TButton', background='#f44336', foreground='#212121', font=HEADER_FONT, borderwidth=2)
        style.map('FirstLast.TButton', background=[('active', '#d32f2f')])
        style.configure('Accent.TButton', background='#1976d2', foreground='white', font=HEADER_FONT)
        style.map('Accent.TButton', background=[('active', '#1565c0')])

        top = ttk.Frame(self, style='Bold.TLabelframe')
        top.pack(fill='x', padx=10, pady=8)
        search_frame = tk.Frame(top, bg='#e3f2fd')
        search_frame.pack(side='left', padx=0)
        self.search_var = tk.StringVar()
        entry = ttk.Entry(search_frame, textvariable=self.search_var, width=20, font=FONT)
        entry.pack(side='left', padx=0)
        entry.bind('<Return>', lambda e: self._search())
        ttk.Button(search_frame, text='Go', style='Accent.TButton', command=self._search).pack(side='left', padx=(2, 8))

        nav_frame = tk.Frame(top, bg='#e3f2fd')
        nav_frame.pack(side='left', padx=0)
        ttk.Button(nav_frame, text='Prev', style='PrevNext.TButton', command=self._prev).pack(side='left', padx=0)
        ttk.Button(nav_frame, text='Next', style='PrevNext.TButton', command=self._next).pack(side='left', padx=(2, 8))
        ttk.Button(nav_frame, text='First', style='FirstLast.TButton', command=self._first).pack(side='left', padx=0)
        ttk.Button(nav_frame, text='Last', style='FirstLast.TButton', command=self._last).pack(side='left', padx=2)
        self.store_count_label = tk.Label(top, text="", font=FONT, fg="#1976d2", bg='#e3f2fd')
        self.store_count_label.pack(side='left', padx=10)
        self.zero_sales_stores_label = tk.Label(top, text="", font=FONT, fg="#FF0000", bg='#e3f2fd')
        self.zero_sales_stores_label.pack(side='left', padx=10)

        sf = tk.Frame(self, bg='#bbdefb', bd=1, relief='groove')
        sf.pack(fill='x', pady=(0,10))
        headers = ['Article No','Rank','ASP',f'MRP','Sales','Revenue','Inventory','Pending']
        self.summary = {}
        for i,h in enumerate(headers):
            tk.Label(sf, text=h, font=HEADER_FONT, bg='#bbdefb', fg='#212121').grid(row=0,column=i,padx=5,pady=2)
            lbl = tk.Label(sf, text='', font=FONT, bg='#e3f2fd', fg='#212121', width=12)
            lbl.grid(row=1,column=i,padx=5,pady=2)
            self.summary[h] = lbl

        wf = tk.Frame(self, bg='#e3f2fd')
        wf.pack(fill='x', pady=(0,10))
        self.week_buttons = {}
        for w in WEEKS:
            cnt = self.week_qty.get(w, sum(self.week_qty.values())) if w!='Overall' else sum(self.week_qty.values())
            btn = ttk.Button(wf, text=f"{w} ({cnt})", style='Accent.TButton', command=lambda x=w: self._set_week(x))
            btn.pack(side='left', padx=5)
            self.week_buttons[w] = btn

        cf = tk.Frame(self, bg='#e3f2fd')
        cf.pack(fill='both', expand=True, padx=10, pady=5)
        left = tk.Frame(cf, bg='#e3f2fd')
        left.pack(side='left', fill='y', padx=(0,10))
        ip = tk.LabelFrame(left, text='Image Preview', bg='#e3f2fd', font=HEADER_FONT, bd=2, relief='solid', fg='#1976d2')
        ip.config(width=IMAGE_DISPLAY_SIZE[0], height=IMAGE_DISPLAY_SIZE[1])
        ip.pack()
        ip.pack_propagate(False)
        self.image_label = tk.Label(ip, bg='#e3f2fd')
        self.image_label.pack(fill='both', expand=True)
        logo_frame = tk.Frame(left, bg='#e3f2fd')
        logo_frame.pack(fill='x', pady=(12,0))
        if os.path.exists(LOGO_PATH):
            logo_img = Image.open(LOGO_PATH)
            logo_img = logo_img.resize(LOGO_DISPLAY_SIZE, Image.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(logo_img)
            tk.Label(logo_frame, image=self.logo_photo, bg='#e3f2fd').pack(pady=0)

        right = tk.Frame(cf, bg='#e3f2fd')
        right.pack(side='left', fill='both', expand=True)
        store_frame = ttk.LabelFrame(right, text='Store-wise', style="Bold.TLabelframe")
        store_frame.pack(side='left', fill='both', expand=True, padx=(0,10), pady=2)
        self.store_tree = self._make_table(store_frame, 'Store', height=18)
        self.store_tree.bind("<Double-1>", self._on_store_double_click)
        tables_frame = tk.Frame(right, bg='#e3f2fd')
        tables_frame.pack(side='left', fill='both', expand=True)
        color_frame = ttk.LabelFrame(tables_frame, text='Color-wise', style="Bold.TLabelframe")
        color_frame.pack(fill='x', expand=False, pady=(0,5))
        self.color_tree = self._make_table(color_frame, 'Color', height=6)
        size_frame = ttk.LabelFrame(tables_frame, text='Size-wise', style="Bold.TLabelframe")
        size_frame.pack(fill='x', expand=False, pady=(0,5))
        self.size_tree = self._make_table(size_frame, 'Size', height=6)
        detail_frame = ttk.LabelFrame(tables_frame, text='Color-Size-wise', style="Bold.TLabelframe")
        detail_frame.pack(fill='both', expand=True, pady=(0,5))
        self.detail_tree = self._make_detail_table(detail_frame, height=12)

        footer = tk.Frame(self, bg='#e3f2fd')
        footer.pack(side='bottom', fill='x')
        tk.Label(footer, text="For Lazera - made by kunal", font=("Segoe UI", 7), fg="#1976d2", bg='#e3f2fd', anchor='se').pack(side='right', padx=6, pady=2)

    def _make_table(self, parent, key, height=8):
        cols = (key,'Qty','Pending','SOH','Value') if key!='Store' else (key,'Qty','SOH','Value')
        tv = ttk.Treeview(parent, columns=cols, show='headings', height=height, style='Treeview')
        for c in cols:
            tv.heading(c, text=c, anchor='center')
            width = 200 if c==key else 80
            tv.column(c, width=width, anchor='center')
        tv.pack(fill='both', expand=True)
        setattr(self, f'{key.lower()}_tree', tv)
        return tv

    def _make_detail_table(self, parent, height=12):
        cols = ('Color','Size','Qty','Pending','SOH')
        tv = ttk.Treeview(parent, columns=cols, show='headings', height=height, style='Treeview')
        for c in cols:
            tv.heading(c, text=c, anchor='center')
            width = 120 if c in ('Color','Size') else 80
            tv.column(c, width=width, anchor='center')
        tv.pack(fill='both', expand=True)
        return tv

    def _show(self):
        if self.overview:
            if self.week == 'Overall':
                total_sales = self.data['qty'].sum()
            else:
                total_sales = self.data[self.data['week'] == self.week]['qty'].sum()
            total_inv = self.data['soh'].sum()
            total_pending = sum(self.pending_total.values())
            self.summary['Article No'].config(text='Overview')
            self.summary['Rank'].config(text='')
            self.summary['ASP'].config(text='')
            self.summary['MRP'].config(text='')
            self.summary['Sales'].config(text=int(total_sales))
            self.summary['Revenue'].config(text='')
            self.summary['Inventory'].config(text=int(total_inv))
            self.summary['Pending'].config(text=int(total_pending))
            self.store_count_label.config(text="")
            self.zero_sales_stores_label.config(text="")
            img = Image.open(LOGO_PATH)
            img.thumbnail(IMAGE_DISPLAY_SIZE)
            self.photo = ImageTk.PhotoImage(img)
            self.image_label.config(image=self.photo, text='')
            for tbl in ('store','color','size','detail'):
                getattr(self, f'{tbl}_tree').delete(*getattr(self, f'{tbl}_tree').get_children())
            return

        art = self.articles[self.idx]
        if self.week=='Overall':
            sold = self.total_qty.get(art,0)
            dfw = self.data[self.data['article']==art]
        else:
            sold = self.article_week_qty.get((art,self.week),0)
            dfw = self.data[(self.data['article']==art)&(self.data['week']==self.week)]
        asp = self.asp_map.get(art,0)
        mrp = self.mrp_map.get(art, MRP_FIXED)
        revenue = round(sold*asp,2)
        inv_tot = self.inv_map.get(art,0)
        pending_tot = self.pending_total.get(art,0)
        total = len(self.articles)

        # Store counts
        stores_with_stock = set(self.inv_data[(self.inv_data['article']==art) & (self.inv_data['soh']>0)]['store'].unique())
        store_count = len(stores_with_stock)
        self.store_count_label.config(text=f"Stores Available: {store_count}")

        # Stores with SOH > 0 and 0 sales in last 5 weeks
        sales_df = self.data[self.data['article'] == art]
        stores_with_sales = set(sales_df[sales_df['qty'] > 0]['store'].unique())
        zero_sales_stores = stores_with_stock - stores_with_sales
        self.zero_sales_stores_label.config(text=f"Stores with 0 Sales: {len(zero_sales_stores)}")

        for w,btn in self.week_buttons.items():
            cnt = self.total_qty.get(art,0) if w=='Overall' else self.article_week_qty.get((art,w),0)
            btn.config(text=f"{w} ({cnt})")
        s = self.summary
        s['Article No'].config(text=art)
        s['Rank'].config(text=f"{self.idx+1}/{total}")
        s['ASP'].config(text=f"₹{asp:.2f}")
        s['MRP'].config(text=f"₹{mrp}")
        s['Sales'].config(text=sold)
        s['Revenue'].config(text=f"₹{revenue:.2f}")
        s['Inventory'].config(text=inv_tot)
        s['Pending'].config(text=pending_tot)

        path = next((os.path.join(IMAGE_DIR,f'{art}{ext}') for ext in ['.jpg','.jpeg','.png'] if os.path.exists(os.path.join(IMAGE_DIR,f'{art}{ext}'))),None)
        if path:
            img = Image.open(path)
            img.thumbnail(IMAGE_DISPLAY_SIZE)
            self.photo = ImageTk.PhotoImage(img)
            self.image_label.config(image=self.photo, text='')
        else:
            self.image_label.config(image='', text='No Image', fg='#1976d2')

        # Store Table
        self.store_tree.delete(*self.store_tree.get_children())
        qty_map = dfw.groupby('store')['qty'].sum().to_dict()
        items = [i for (a,i) in self.store_inv.keys() if a==art]
        items.sort(key=lambda x: qty_map.get(x,0), reverse=True)
        for val in items:
            qty = qty_map.get(val,0)
            soh = self.store_inv.get((art,val),0)
            valp = round(qty*asp,2)
            self.store_tree.insert('', 'end', values=(val, qty, soh, f"₹{valp:.2f}"))

        # Color Table
        self.color_tree.delete(*self.color_tree.get_children())
        qty_map = dfw.groupby('color')['qty'].sum().to_dict()
        items = [i for (a,i) in self.color_inv.keys() if a==art]
        items.sort(key=lambda x: qty_map.get(x,0), reverse=True)
        for val in items:
            qty = qty_map.get(val,0)
            soh = self.color_inv.get((art,val),0)
            pend = self.pending_color.get((art,val),0)
            valp = round(qty*asp,2)
            self.color_tree.insert('', 'end', values=(val, qty, pend, soh, f"₹{valp:.2f}"))

        # Size Table
        self.size_tree.delete(*self.size_tree.get_children())
        qty_map = dfw.groupby('size')['qty'].sum().to_dict()
        items = [i for (a,i) in self.size_inv.keys() if a==art]
        items.sort(key=lambda x: qty_map.get(x,0), reverse=True)
        for val in items:
            qty = qty_map.get(val,0)
            soh = self.size_inv.get((art,val),0)
            pend = self.pending_size.get((art,val),0)
            valp = round(qty*asp,2)
            self.size_tree.insert('', 'end', values=(val, qty, pend, soh, f"₹{valp:.2f}"))

        # Color-Size Table
        self.detail_tree.delete(*self.detail_tree.get_children())
        detail = dfw.groupby(['color','size']).agg({'qty':'sum','soh':'sum'}).reset_index()
        for _, row in detail.iterrows():
            color, size = row['color'], row['size']
            qty = row['qty']
            soh = row['soh']
            pend = self.pending_colorsize.get((art, color, size), 0)
            self.detail_tree.insert('', 'end', values=(color, size, qty, pend, soh))

    def _on_store_double_click(self, event):
        item = self.store_tree.selection()
        if not item:
            return
        store = self.store_tree.item(item, "values")[0]
        art = self.articles[self.idx]
        sales_rows = self.data[(self.data['article'] == art) & (self.data['store'] == store)]
        inv_rows = self.inv_data[(self.inv_data['article'] == art) & (self.inv_data['store'] == store)]
        sales_group = sales_rows.groupby(['color','size'])['qty'].sum().reset_index()
        inv_group = inv_rows.groupby(['color','size'])['soh'].sum().reset_index()
        merged = pd.merge(sales_group, inv_group, on=['color','size'], how='outer').fillna(0)
        merged = merged[(merged['qty'] > 0) | (merged['soh'] > 0)]
        popup = tk.Toplevel(self)
        popup.title(f"{art} - {store} Details")
        popup.geometry(f"{min(600, self.winfo_screenwidth()//2)}x{min(400, self.winfo_screenheight()//2)}")
        ttk.Label(popup, text=f"Store: {store} | Article: {art}", font=HEADER_FONT, background="#e3f2fd", foreground="#1976d2").pack(pady=6)
        cols = ('Color','Size','Qty Sold','SOH')
        tree = ttk.Treeview(popup, columns=cols, show='headings', height=15, style='Treeview')
        for c in cols:
            tree.heading(c, text=c, anchor='center')
            tree.column(c, width=120 if c in ('Color','Size') else 80, anchor='center')
        tree.pack(fill='both', expand=True, padx=10, pady=10)
        for _, row in merged.iterrows():
            tree.insert('', 'end', values=(row['color'], row['size'], int(row['qty']), int(row['soh'])))
        ttk.Button(popup, text="Close", style='Accent.TButton', command=popup.destroy).pack(pady=6)

    def _prev(self):
        if self.overview:
            return
        if self.idx > 0:
            self.idx -= 1
        else:
            self.overview = True
        self._show()

    def _next(self):
        if self.overview:
            self.overview = False
            self.idx = 0
        elif self.idx < len(self.articles) - 1:
            self.idx += 1
        self._show()

    def _first(self):
        self.overview = True
        self._show()

    def _last(self):
        self.overview = False
        self.idx = len(self.articles)-1
        self._show()

    def _set_week(self, w):
        self.week = w
        if not self.overview:
            totals = self.total_qty if w=='Overall' else {a:self.article_week_qty.get((a,w),0) for a in self.articles}
            self.articles = sorted(totals, key=totals.get, reverse=True)
            self.idx = 0
        self._show()

    def _search(self):
        term = self.search_var.get().lower()
        if self.overview:
            return
        for i,a in enumerate(self.articles):
            if term in str(a).lower():
                self.idx = i
                break
        self._show()

if __name__=='__main__':
    AllInOneApp().mainloop()
