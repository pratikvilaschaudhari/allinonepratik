import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import pandas as pd
import logging
import glob

# --- Paths and Constants ---
if getattr(sys, 'frozen', False):
    APP_ROOT = sys._MEIPASS
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
FONT = ("Segoe UI", 10)
HEADER_FONT = ("Segoe UI", 10, "bold")

# --- Logging Setup ---
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
    merged = pd.merge(sales_df, inv_df, on=['article','store','color','size'], how='left')
    merged['soh'] = merged['soh'].fillna(0)
    merged['asp_calc'] = merged['article'].map(asp_map)
    return merged

class AllInOneApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('AllInOne App | Lazera Shoes')
        self.state('zoomed')

        sales = load_sales_data()
        inv = load_inventory_data()
        pending = load_pending_data()

        self.asp_map = calculate_asp_map(sales)
        self.data = merge_data(sales, inv, self.asp_map)
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
        style.configure('Treeview', rowheight=24, font=FONT, borderwidth=2, relief='solid')
        style.configure('Treeview.Heading', font=HEADER_FONT)
        style.configure('Bold.TLabelframe', borderwidth=3, relief='solid')
        style.configure('Bold.TLabelframe.Label', font=HEADER_FONT)

        # Top controls
        top = ttk.Frame(self); top.pack(fill='x', padx=10, pady=5)
        self.search_var = tk.StringVar()
        entry = ttk.Entry(top, textvariable=self.search_var, width=20)
        entry.pack(side='left'); entry.bind('<Return>', lambda e: self._search())
        ttk.Button(top, text='Go', command=self._search).pack(side='left', padx=2)
        ttk.Button(top, text='Prev', command=self._prev).pack(side='left', padx=2)
        ttk.Button(top, text='Next', command=self._next).pack(side='left', padx=2)

        # Summary
        sf = tk.Frame(self, bg='white', bd=1, relief='groove')
        sf.pack(fill='x', pady=(0,10))
        headers = ['Article No','Rank','ASP',f'MRP','Sales','Revenue','Inventory','Pending']
        self.summary = {}
        for i,h in enumerate(headers):
            tk.Label(sf, text=h, font=HEADER_FONT, bg='white').grid(row=0,column=i,padx=5,pady=2)
            lbl = tk.Label(sf, text='', font=FONT, bg='white', width=12)
            lbl.grid(row=1,column=i,padx=5,pady=2)
            self.summary[h] = lbl

        # Week buttons
        wf = tk.Frame(self, bg='white'); wf.pack(fill='x', pady=(0,10))
        self.week_buttons = {}
        for w in WEEKS:
            cnt = self.week_qty.get(w, sum(self.week_qty.values())) if w!='Overall' else sum(self.week_qty.values())
            btn = ttk.Button(wf, text=f"{w} ({cnt})", command=lambda x=w: self._set_week(x))
            btn.pack(side='left', padx=5);
            self.week_buttons[w] = btn

        # Main content
        cf = tk.Frame(self, bg='white'); cf.pack(fill='both', expand=True, padx=10, pady=5)
        # Image
        left = tk.Frame(cf, bg='white'); left.pack(side='left', fill='y', padx=(0,10))
        ip = tk.LabelFrame(left, text='Image Preview', bg='white', font=HEADER_FONT, bd=2, relief='solid')
        ip.config(width=IMAGE_DISPLAY_SIZE[0], height=IMAGE_DISPLAY_SIZE[1])
        ip.pack(); ip.pack_propagate(False)
        self.image_label = tk.Label(ip, bg='white'); self.image_label.pack(fill='both', expand=True)

        # Store table (vertical, 50%)
        right = tk.Frame(cf, bg='white'); right.pack(side='left', fill='both', expand=True)
        store_frame = ttk.LabelFrame(right, text='Store-wise', style="Bold.TLabelframe")
        store_frame.pack(side='left', fill='both', expand=True, padx=(0,10), pady=2)
        self.store_tree = self._make_table(store_frame, 'Store', height=18, borderwidth=2)

        # Right-side tables stacked vertically
        tables_frame = tk.Frame(right, bg='white')
        tables_frame.pack(side='left', fill='both', expand=True)

        color_frame = ttk.LabelFrame(tables_frame, text='Color-wise', style="Bold.TLabelframe")
        color_frame.pack(fill='x', expand=False, pady=(0,5))
        self.color_tree = self._make_table(color_frame, 'Color', height=6, borderwidth=2)

        size_frame = ttk.LabelFrame(tables_frame, text='Size-wise', style="Bold.TLabelframe")
        size_frame.pack(fill='x', expand=False, pady=(0,5))
        self.size_tree = self._make_table(size_frame, 'Size', height=6, borderwidth=2)

        detail_frame = ttk.LabelFrame(tables_frame, text='Color-Size-wise', style="Bold.TLabelframe")
        detail_frame.pack(fill='both', expand=True, pady=(0,5))
        self.detail_tree = self._make_detail_table(detail_frame, height=12, borderwidth=2)

    def _make_table(self, parent, key, height=8, borderwidth=2):
        cols = (key,'Qty','Pending','SOH','Value') if key!='Store' else (key,'Qty','SOH','Value')
        tv = ttk.Treeview(parent, columns=cols, show='headings', height=height)
        for c in cols:
            tv.heading(c, text=c, anchor='center')
            width = 200 if c==key else 80
            tv.column(c, width=width, anchor='center')
        tv.pack(fill='both', expand=True)
        tv['style'] = 'Treeview'
        setattr(self, f'{key.lower()}_tree', tv)
        return tv

    def _make_detail_table(self, parent, height=12, borderwidth=2):
        cols = ('Color','Size','Qty','Pending','SOH')
        tv = ttk.Treeview(parent, columns=cols, show='headings', height=height)
        for c in cols:
            tv.heading(c, text=c, anchor='center')
            width = 120 if c in ('Color','Size') else 80
            tv.column(c, width=width, anchor='center')
        tv.pack(fill='both', expand=True)
        tv['style'] = 'Treeview'
        return tv

    def _show(self):
        if self.overview:
            total_sales = sum(self.week_qty.values())
            total_inv = sum(self.inv_map.values())
            total_pending = sum(self.pending_total.values())
            self.summary['Article No'].config(text='Overview')
            self.summary['Rank'].config(text='')
            self.summary['ASP'].config(text='')
            self.summary['MRP'].config(text='')
            self.summary['Sales'].config(text=total_sales)
            self.summary['Revenue'].config(text='')
            self.summary['Inventory'].config(text=total_inv)
            self.summary['Pending'].config(text=total_pending)
            img = Image.open(LOGO_PATH); img.thumbnail(IMAGE_DISPLAY_SIZE)
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
            img = Image.open(path); img.thumbnail(IMAGE_DISPLAY_SIZE)
            self.photo = ImageTk.PhotoImage(img)
            self.image_label.config(image=self.photo, text='')
        else:
            self.image_label.config(image='', text='No Image', fg='grey')
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

    def _prev(self):
        if self.overview: return
        if self.idx>0: self.idx-=1
        else: self.overview=True
        self._show()

    def _next(self):
        if self.overview: self.overview=False
        elif self.idx<len(self.articles)-1: self.idx+=1
        self._show()

    def _set_week(self, w):
        self.week=w
        if not self.overview:
            totals = self.total_qty if w=='Overall' else {a:self.article_week_qty.get((a,w),0) for a in self.articles}
            self.articles = sorted(totals, key=totals.get, reverse=True)
            self.idx=0
        self._show()

    def _search(self):
        term = self.search_var.get().lower()
        if self.overview: return
        for i,a in enumerate(self.articles):
            if term in str(a).lower(): self.idx=i; break
        self._show()

if __name__=='__main__':
    AllInOneApp().mainloop()
