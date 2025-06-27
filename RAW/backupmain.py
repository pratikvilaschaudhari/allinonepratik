import os
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import pandas as pd
import logging
import glob

# --- Paths and Constants ---
APP_ROOT = "D:/allinone"
SALES_DIR = os.path.join(APP_ROOT, 'sales data')
INVENTORY_DIR = os.path.join(APP_ROOT, 'inventory data')
IMAGE_DIR = os.path.join(APP_ROOT, 'images')
LOGO_PATH = os.path.join(APP_ROOT, 'Lazera Logo-02.png')
ERROR_LOG_PATH = os.path.join(APP_ROOT, 'app_code', 'error_log.txt')
MRP_FIXED = 1999
WEEKS = ['Week 1', 'Week 2', 'Week 3', 'Week 4', 'Week 5', 'Overall']
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

# --- Data Loading ---
def get_latest_files(directory, pattern, count=5):
    try:
        files = glob.glob(os.path.join(directory, pattern))
        return sorted(files, key=os.path.getmtime, reverse=True)[:count]
    except Exception as e:
        logging.error(f"Error listing files in {directory}: {e}")
        return []


def load_sales_data():
    dfs = []
    for idx, path in enumerate(get_latest_files(SALES_DIR, 'salesdata*.xlsx', 5), 1):
        try:
            df = pd.read_excel(path)
            df.rename(columns={'Article':'article','store':'store','Colour':'Color','Size':'Size','Quantity':'Qty','ASP':'ASP'}, inplace=True)
            df['Week'] = f'Week {idx}'
            cols = ['article','store','Color','Size','Qty','ASP']
            if all(c in df.columns for c in cols):
                dfs.append(df[cols + ['Week']])
            else:
                missing=[c for c in cols if c not in df.columns]
                logging.error(f"Missing columns in {os.path.basename(path)}: {missing}")
        except Exception as e:
            logging.error(f"Error reading {path}: {e}")
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame(columns=['article','store','Color','Size','Qty','ASP','Week'])


def load_inventory_data():
    files = get_latest_files(INVENTORY_DIR, '*.xlsx', 1)
    if not files:
        logging.error("No inventory files found")
        return pd.DataFrame(columns=['article','store','Color','Size','SOH'])
    try:
        df = pd.read_excel(files[0])
        df.rename(columns={'store':'store','Article':'article','Size':'Size','Colour':'Color','Quantity':'SOH'}, inplace=True)
        cols=['article','store','Color','Size','SOH']
        if all(c in df.columns for c in cols):
            return df[cols]
        else:
            missing=[c for c in cols if c not in df.columns]
            logging.error(f"Missing columns in inventory: {missing}")
            return pd.DataFrame(columns=cols)
    except Exception as e:
        logging.error(f"Error reading inventory: {e}")
        return pd.DataFrame(columns=['article','store','Color','Size','SOH'])

# --- Business Logic ---
def calculate_article_asp(df):
    if df.empty:
        return {}
    try:
        agg = df.groupby('article').agg(
            total_revenue=pd.NamedAgg(column='ASP', aggfunc=lambda s: (s * df.loc[s.index, 'Qty']).sum()),
            total_qty=pd.NamedAgg(column='Qty', aggfunc='sum')
        )
        agg['ASP'] = agg.apply(lambda r: r['total_revenue'] / r['total_qty'] if r['total_qty'] > 0 else 0, axis=1)
        return agg['ASP'].to_dict()
    except Exception as e:
        logging.error(f"Error calculating ASP: {e}")
        return {}


def merge_data(sales_df, inv_df, asp_map):
    if sales_df.empty:
        return sales_df
    merged = pd.merge(sales_df, inv_df, on=['article','store','Color','Size'], how='left')
    merged['SOH'] = merged.get('SOH', 0)
    merged['ASP'] = merged['article'].map(asp_map)
    return merged

# --- UI Class ---
class AllInOneApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('AllInOne App | Lazera Shoes')
        self.state('zoomed')
        self.configure(bg='white')

        style = ttk.Style(self)
        style.configure('Treeview', font=FONT, rowheight=25)
        style.configure('Treeview.Heading', font=HEADER_FONT)
        style.configure('TButton', font=FONT)
        style.configure('TEntry', font=FONT)

        sales = load_sales_data()
        inv = load_inventory_data()
        asp_map = calculate_article_asp(sales)
        data = merge_data(sales, inv, asp_map)
        data['SOH'] = data.get('SOH', 0)

        # preserve raw inventory for correct SOH
        self.inv_data = inv
        self.data = data
        # maps for article/week sales
        self.article_week_map = data.groupby(['article','Week'])['Qty'].sum().to_dict()
        self.article_total_map = data.groupby('article')['Qty'].sum().to_dict()
        # correct inventory maps from raw inv_data
        self.article_inventory_map = inv.groupby('article')['SOH'].sum().to_dict()
        self.store_inventory_map = inv.groupby(['article','store'])['SOH'].sum().to_dict()
        self.color_inventory_map = inv.groupby(['article','Color'])['SOH'].sum().to_dict()
        self.size_inventory_map = inv.groupby(['article','Size'])['SOH'].sum().to_dict()

        # initial order by overall sales
        if self.article_total_map:
            self.articles = sorted(self.article_total_map, key=self.article_total_map.get, reverse=True)
        else:
            self.articles = []

        self.current_idx = 0 if self.articles else -1
        self.current_week = 'Overall'

        self._build_ui()
        if self.articles:
            self._display_article()

    def _build_ui(self):
        mf = tk.Frame(self, bg='white')
        mf.pack(fill='both', expand=True, padx=10, pady=10)

        # top
        top = tk.Frame(mf, bg='#f0f0f0')
        top.pack(fill='x', pady=(0,10))
        nav = tk.Frame(top, bg='#f0f0f0')
        nav.pack(side='left', padx=10, pady=5)
        self.search_var = tk.StringVar()
        se = ttk.Entry(nav, textvariable=self.search_var, width=30)
        se.pack(side='left', padx=5)
        se.bind('<Return>', lambda e: self.search_article())
        self.prev_btn = ttk.Button(nav, text='◀ Prev', command=self.prev_article)
        self.prev_btn.pack(side='left', padx=5)
        self.next_btn = ttk.Button(nav, text='Next ▶', command=self.next_article)
        self.next_btn.pack(side='left', padx=5)
        if os.path.exists(LOGO_PATH):
            try:
                logo = Image.open(LOGO_PATH).resize((180,60), Image.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(logo)
                tk.Label(top, image=self.logo_img, bg='#f0f0f0').pack(side='right', padx=10)
            except Exception as e:
                logging.error(f"Logo error: {e}")

        # summary
        sf = tk.Frame(mf, bg='white', bd=1, relief='groove')
        sf.pack(fill='x', pady=(0,10))
        headers = ["Article No","ASP",f"MRP (₹{MRP_FIXED})","Sales","Revenue","Inventory"]
        self.summary_labels = []
        for i,h in enumerate(headers):
            tk.Label(sf,text=h,font=HEADER_FONT,bg='white').grid(row=0,column=i,padx=5,pady=2)
            lbl=tk.Label(sf,text='',font=FONT,bg='white',width=20)
            lbl.grid(row=1,column=i,padx=5,pady=2)
            self.summary_labels.append(lbl)

        # weeks
        wf=tk.Frame(mf,bg='white')
        wf.pack(fill='x',pady=(0,10))
        self.week_buttons={}
        for w in WEEKS:
            btn=ttk.Button(wf,text=w,width=12,command=lambda x=w:self.set_week(x))
            btn.pack(side='left',padx=5)
            self.week_buttons[w]=btn

        # main area
        cf=tk.Frame(mf,bg='white')
        cf.pack(fill='both',expand=True)

        # image
        ip=tk.LabelFrame(cf,text='Image Preview',font=HEADER_FONT,bg='white')
        ip.pack(side='left',padx=(0,10))
        ip.config(width=IMAGE_DISPLAY_SIZE[0]+20,height=IMAGE_DISPLAY_SIZE[1]+20)
        ip.pack_propagate(False)
        self.image_container=tk.Frame(ip,width=IMAGE_DISPLAY_SIZE[0],height=IMAGE_DISPLAY_SIZE[1],bg='white')
        self.image_container.place(x=10,y=10)
        self.image_container.pack_propagate(False)
        self.image_label=tk.Label(self.image_container,text='Image will appear here',font=FONT,bg='white')
        self.image_label.pack(fill='both',expand=True)

        # tables
        tf=tk.Frame(cf,bg='white')
        tf.pack(side='right',fill='both',expand=True)

        # store-wise
        sf2=tk.LabelFrame(tf,text='Store-wise',font=HEADER_FONT,bg='white')
        sf2.pack(fill='both',expand=True,pady=(0,5))
        self.store_tree=ttk.Treeview(sf2,columns=('Store','Qty','SOH','Value'),show='headings')
        for col,w in [('Store',100),('Qty',80),('SOH',80),('Value',100)]:
            self.store_tree.heading(col,text=col)
            self.store_tree.column(col,width=w,anchor='center')
        self.store_tree.pack(fill='both',expand=True,padx=5,pady=5)

        bf=tk.Frame(tf,bg='white')
        bf.pack(fill='both',expand=True)

        # color-wise
        cf2=tk.LabelFrame(bf,text='Color-wise',font=HEADER_FONT,bg='white')
        cf2.pack(side='left',fill='both',expand=True,padx=(0,5))
        self.color_tree=ttk.Treeview(cf2,columns=('Color','Qty','SOH'),show='headings')
        for col,w in [('Color',120),('Qty',80),('SOH',80)]:
            self.color_tree.heading(col,text=col)
            self.color_tree.column(col,width=w,anchor='center')
        self.color_tree.pack(fill='both',expand=True,padx=5,pady=5)

        # size-wise
        cf3=tk.LabelFrame(bf,text='Size-wise',font=HEADER_FONT,bg='white')
        cf3.pack(side='right',fill='both',expand=True,padx=(5,0))
        self.size_tree=ttk.Treeview(cf3,columns=('Size','Qty','SOH'),show='headings')
        for col,w in [('Size',80),('Qty',80),('SOH',80)]:
            self.size_tree.heading(col,text=col)
            self.size_tree.column(col,width=w,anchor='center')
        self.size_tree.pack(fill='both',expand=True,padx=5,pady=5)

        # footer
        ff=tk.Frame(self,bg='#f0f0f0')
        ff.pack(side='bottom',fill='x')
        tk.Label(ff,text='AllInOne App | Lazera Shoes - Internal Use Only',font=("Segoe UI",9),bg='#f0f0f0').pack(pady=5)

    def set_week(self,week):
        self.current_week=week
        if week=='Overall':
            totals=self.article_total_map
        else:
            totals={art:self.article_week_map.get((art,week),0) for art in self.articles}
        self.articles=sorted(totals,key=totals.get,reverse=True)
        self.current_idx=0 if self.articles else -1
        self._display_article()

    def search_article(self):
        q=self.search_var.get().strip()
        matches=[i for i,a in enumerate(self.articles) if q.lower() in str(a).lower()]
        if matches:
            self.current_idx=matches[0]
            self._display_article()
        else:
            messagebox.showinfo('Not Found','No matching article')

    def prev_article(self):
        self.current_idx=(self.current_idx-1)%len(self.articles)
        self._display_article()

    def next_article(self):
        self.current_idx=(self.current_idx+1)%len(self.articles)
        self._display_article()

    def _display_article(self):
        art=self.articles[self.current_idx]
        # update week buttons
        for w,btn in self.week_buttons.items():
            cnt=self.article_total_map.get(art,0) if w=='Overall' else self.article_week_map.get((art,w),0)
            btn.config(text=f"{w} ({cnt})")

        # filter sales df for tables
        df=self.data[self.data['article']==art]
        if self.current_week!='Overall': df=df[df['Week']==self.current_week]

        asp=calculate_article_asp(df).get(art,0)
        sales_qty=df['Qty'].sum()
        sales_rev=round(sales_qty*asp,2)
        # correct summary inventory
        inv_total=self.article_inventory_map.get(art,0)

        # update summary labels
        lbls=self.summary_labels
        lbls[0].config(text=art)
        lbls[1].config(text=f"₹{asp:.2f}")
        lbls[2].config(text=f"₹{MRP_FIXED}")
        lbls[3].config(text=sales_qty)
        lbls[4].config(text=f"₹{sales_rev:.2f}")
        lbls[5].config(text=inv_total)

        # load image
        for ext in ['.jpg','.jpeg','.png']:
            path=os.path.join(IMAGE_DIR,f"{art}{ext}")
            if os.path.exists(path):
                img=Image.open(path)
                img.thumbnail(IMAGE_DISPLAY_SIZE,Image.LANCZOS)
                canvas=Image.new('RGB',IMAGE_DISPLAY_SIZE,(255,255,255))
                x=(IMAGE_DISPLAY_SIZE[0]-img.size[0])//2
                y=(IMAGE_DISPLAY_SIZE[1]-img.size[1])//2
                canvas.paste(img,(x,y))
                self.img_tk=ImageTk.PhotoImage(canvas)
                self.image_label.config(image=self.img_tk,text='')
                break
        else:
            self.image_label.config(image='',text=f"Image not found for\n{art}",fg='red',font=("Segoe UI",12))

        # clear existing rows
        for tr in (self.store_tree,self.color_tree,self.size_tree):
            for row in tr.get_children(): tr.delete(row)

        # populate store-wise table
        sales_by_store=df.groupby('store')['Qty'].sum().to_dict()
        for store,qty in sorted(sales_by_store.items(),key=lambda x:x[1],reverse=True):
            soh=self.store_inventory_map.get((art,store),0)
            val=round(qty*asp,2)
            self.store_tree.insert('', 'end', values=(store,qty,soh,val))

        # populate color-wise table
        sales_by_color=df.groupby('Color')['Qty'].sum().to_dict()
        for color,qty in sales_by_color.items():
            soh=self.color_inventory_map.get((art,color),0)
            self.color_tree.insert('', 'end', values=(color,qty,soh))

        # populate size-wise table
        sales_by_size=df.groupby('Size')['Qty'].sum().to_dict()
        for size,qty in sales_by_size.items():
            soh=self.size_inventory_map.get((art,size),0)
            self.size_tree.insert('', 'end', values=(size,qty,soh))

        # nav button states
        self.prev_btn.config(state='normal' if self.current_idx>0 else 'disabled')
        self.next_btn.config(state='normal' if len(self.articles)>1 else 'disabled')

if __name__=='__main__':
    AllInOneApp().mainloop()
