import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, font
from PIL import Image, ImageTk
import csv
import os
import pytesseract
import re
from pdf2image import convert_from_path
import time
import json
import shutil
import uuid
import webbrowser
from datetime import datetime

# ==========================================
# CONFIGURATION
# ==========================================
#
# Need to install:
# - pandas (pip install pandas)
# - pillow (pip install pillow)
# - pytesseract (pip install pytesseract)
# - pdf2image (pip install pdf2image)
# - Tesseract OCR (installer available at https://github.com/UB-Mannheim/tesseract/wiki)
# - Poppler (binary available at https://github.com/oschwartz10612/poppler-windows/releases/)
#
# 1. TESSERACT CONFIG (Windows Only)
# If Tesseract is not in your PATH, uncomment and point to the exe:
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
#
# 2. POPPLER CONFIG (Windows Only)
# If Poppler is not in your PATH, provide the path to the 'bin' folder:
POPPLER_PATH = r'C:\Program Files\poppler-25.12.0\Library\bin' 
#POPPLER_PATH = None # Set to None if installed via Brew/Apt or added to System PATH


DEFAULT_CSV_NAME = "contacts.csv"
IMG_FOLDER_NAME = "card_images"
MIN_ROW_HEIGHT = 40
TEXT_ROW_HEIGHT = 30 

CSV_HEADERS = [
    "ID", "First Name", "Last Name", "Company", "Job Title", 
    "E-mail Address", "Mobile Phone", "Business Phone", 
    "Address", "Notes Data", "Image Data"
]

ALL_DATA_COLUMNS = [
    "First Name", "Last Name", "Company", "Job Title", 
    "E-mail Address", "Mobile Phone", "Business Phone", "Address"
]

DEFAULT_VISIBLE = ["Select", "First Name", "Last Name", "Company", "E-mail Address", "Mobile Phone"]

class AutoScrollbar(ttk.Scrollbar):
    def set(self, lo, hi):
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            self.pack_forget()
        else:
            self.pack(side="right", fill="y")
        super().set(lo, hi)

class AutoHScrollbar(ttk.Scrollbar):
    def set(self, lo, hi):
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            self.pack_forget()
        else:
            self.pack(side="bottom", fill="x")
        super().set(lo, hi)

class RolodexApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Rolodex")
        self.root.geometry("1300x800")
        
        self.work_dir = os.getcwd()
        self.contacts = [] 
        self.contact_map = {} 
        self.selected_ids = set() 
        self.photo_cache = {} 
        self.active_filters = {} 
        self.current_row_height = 80
        
        # State
        self.drag_source_col = None
        self.is_dragging = False
        self.settings_popup = None
        self.visible_columns = list(DEFAULT_VISIBLE) 
        self.show_images = True
        self.block_menu = False
        self._click_job = None
        
        # Logic for Settings Toggle
        self.ignore_settings_open = False
        
        # UI Elements
        self.drop_marker = None 

        self.setup_styles()
        self.ensure_directories()
        
        self.create_toolbar()
        self.create_list_view()
        
        self.load_data()
        self.refresh_list()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        bg_color = "#f4f4f4"
        text_color = "#333333"
        
        self.root.configure(bg=bg_color)
        self.style.configure("TFrame", background=bg_color)
        self.style.configure("TLabel", background=bg_color, foreground="#666666", font=("Segoe UI", 9)) 
        self.style.configure("TButton", font=("Segoe UI", 9), padding=5)
        
        self.style.configure("Green.TButton", background="#4CAF50", foreground="white", font=("Segoe UI", 11, "bold"))
        self.style.map("Green.TButton", background=[('active', '#43A047')])
        self.style.configure("Red.TButton", background="#E53935", foreground="white", font=("Segoe UI", 9, "bold"))
        self.style.map("Red.TButton", background=[('active', '#D32F2F')])

        # Seamless Treeview Style
        self.style.configure("Treeview", background="white", foreground=text_color, 
                             rowheight=self.current_row_height, fieldbackground="white", font=("Segoe UI", 10),
                             borderwidth=0)
        self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#e0e0e0", 
                             borderwidth=1, relief="raised")
        self.style.map("Treeview", background=[('selected', '#e3f2fd')], foreground=[('selected', 'black')])

    def ensure_directories(self):
        img_dir = os.path.join(self.work_dir, IMG_FOLDER_NAME)
        if not os.path.exists(img_dir): os.makedirs(img_dir)

    # ==========================
    # UI Layout
    # ==========================
    def create_toolbar(self):
        self.toolbar = tk.Frame(self.root, bg="#e0e0e0", pady=8, padx=10, relief="flat")
        self.toolbar.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(self.toolbar, text="+", style="Green.TButton", width=4, 
                   command=self.show_add_options).pack(side=tk.LEFT, padx=(0, 15))

        tk.Label(self.toolbar, text="Search:", bg="#e0e0e0", fg="#666").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_contacts_search)
        tk.Entry(self.toolbar, textvariable=self.search_var, font=("Segoe UI", 10), width=20).pack(side=tk.LEFT, padx=5)

        tk.Label(self.toolbar, text="|   Dir:", bg="#e0e0e0", fg="#666").pack(side=tk.LEFT, padx=(15, 5))
        self.dir_label_var = tk.StringVar(value=self.work_dir)
        tk.Label(self.toolbar, textvariable=self.dir_label_var, bg="#e0e0e0", fg="#333", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)

        self.right_tools = tk.Frame(self.toolbar, bg="#e0e0e0")
        self.right_tools.pack(side=tk.RIGHT)
        
        self.btn_settings = ttk.Button(self.right_tools, text="☰", width=4, command=self.toggle_settings_popup)
        self.btn_settings.pack(side=tk.RIGHT, padx=5)

        self.batch_actions_frame = tk.Frame(self.right_tools, bg="#e0e0e0")
        self.batch_actions_frame.pack(side=tk.RIGHT, padx=5)

        self.btn_delete_selected = ttk.Button(self.batch_actions_frame, text="Delete Selected", style="Red.TButton", 
                                              command=self.delete_selected_contacts)
        self.btn_edit_selected = ttk.Button(self.batch_actions_frame, text="Edit Selected",
                                            command=self.edit_selected_contacts)

    def create_list_view(self):
        outer_frame = tk.Frame(self.root, bg="#f4f4f4")
        outer_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Border Container (Seamless Look)
        self.list_border_frame = tk.Frame(outer_frame, bg="#A0A0A0", bd=1) 
        self.list_border_frame.pack(fill=tk.BOTH, expand=True)

        # Inner Container - bd=0 for seamless
        self.list_container = tk.Frame(self.list_border_frame, bg="white", bd=0, highlightthickness=0)
        self.list_container.pack(fill=tk.BOTH, expand=True)

        # 1. Image Tree (Right)
        self.img_tree = ttk.Treeview(self.list_container, show="tree headings", selectmode="none")
        self.img_tree.heading("#0", text="Image")
        self.img_tree.column("#0", width=120, anchor="center")
        
        # 2. Main Data Tree (Left)
        # Use Frame to control width properly
        self.tree_frame = tk.Frame(self.list_container, bg="white", bd=0, highlightthickness=0)
        
        all_cols = ["Select"] + ALL_DATA_COLUMNS
        self.tree = ttk.Treeview(self.tree_frame, columns=all_cols, show="headings", selectmode="none")
        
        # Horizontal Scrollbar for Main Tree
        self.hsb = AutoHScrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=self.hsb.set)
        
        # 3. Vertical Scrollbar (Unified)
        self.vsb = AutoScrollbar(self.list_container, orient="vertical")
        
        def scroll_both(*args):
            self.tree.yview(*args)
            self.img_tree.yview(*args)
        
        self.vsb.config(command=scroll_both)
        self.tree.configure(yscroll=self.vsb.set) 
        
        # Packing Order
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        # Image tree packed in refresh_list based on visibility
        
        self.hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False) # Expand=False is key for shrinking

        self.drop_marker = tk.Canvas(self.tree, width=10, height=30, bg=None, bd=0, highlightthickness=0)

        self.setup_data_columns()

        # Bindings
        self.tree.bind("<Button-1>", self.on_click)
        self.tree.bind("<ButtonRelease-1>", self.on_release)
        self.tree.bind("<B1-Motion>", self.on_drag_motion)
        self.tree.bind("<Double-1>", self.on_double_click)
        
        def on_mousewheel(event):
            self.tree.yview_scroll(int(-1*(event.delta/120)), "units")
            self.img_tree.yview_scroll(int(-1*(event.delta/120)), "units")
            return "break"

        self.tree.bind("<MouseWheel>", on_mousewheel)
        self.img_tree.bind("<MouseWheel>", on_mousewheel)
        
        # Bind resize of image column
        self.img_tree.bind("<ButtonRelease-1>", self.on_img_col_resize)

    def setup_data_columns(self):
        if "Select" in self.visible_columns:
            if self.visible_columns[0] != "Select":
                self.visible_columns.remove("Select")
                self.visible_columns.insert(0, "Select")

        self.tree["displaycolumns"] = self.visible_columns

        all_cols = ["Select"] + ALL_DATA_COLUMNS
        for col in all_cols:
            if col == "Select":
                self.tree.heading(col, text="✓", command=self.toggle_select_all)
                self.tree.column(col, width=40, anchor="center", stretch=False)
            else:
                self.tree.heading(col, text=f"{col} ▼")
                self.tree.column(col, width=150, minwidth=50, stretch=False)

    def get_col_at_x(self, x):
        display_cols = self.tree["displaycolumns"]
        # Approximate check. Treeview doesn't expose horizontal scroll offset easily in pixels
        # But identify_column works with screen-relative-to-widget coordinates.
        col_id = self.tree.identify_column(x)
        if not col_id: return None, 0, 0
        
        idx = int(col_id.replace("#", "")) - 1
        all_defined = ["Select"] + ALL_DATA_COLUMNS
        if 0 <= idx < len(all_defined):
            return all_defined[idx], 0, 0 
        return None, 0, 0

    # ==========================
    # Logic: Interaction
    # ==========================
    def toggle_select_all(self):
        visible_ids = [self.tree.item(child)['tags'][0] for child in self.tree.get_children()]
        if not visible_ids: return
        all_selected = all(uid in self.selected_ids for uid in visible_ids)
        if all_selected:
            for uid in visible_ids: self.selected_ids.discard(uid)
        else:
            for uid in visible_ids: self.selected_ids.add(uid)
        self.refresh_list_visuals_only()
        self.update_batch_buttons()

    def on_click(self, event):
        self.is_dragging = False
        self.drag_start_x = event.x
        
        if self._click_job:
            self.root.after_cancel(self._click_job)
            self._click_job = None

        region = self.tree.identify("region", event.x, event.y)
        
        if region == "heading":
            col_name, _, _ = self.get_col_at_x(event.x)
            if col_name: self.drag_source_col = col_name
            else: self.drag_source_col = None

        elif region in ["tree", "cell"]:
            col_name, _, _ = self.get_col_at_x(event.x)
            if col_name == "Select":
                row_id = self.tree.identify_row(event.y)
                if row_id:
                    c_id = self.tree.item(row_id)['tags'][0]
                    if c_id in self.selected_ids: self.selected_ids.remove(c_id)
                    else: self.selected_ids.add(c_id)
                    self.refresh_list_visuals_only()
                    self.update_batch_buttons()

    def on_drag_motion(self, event):
        if not self.drag_source_col: return
        if abs(event.x - self.drag_start_x) > 5:
            self.is_dragging = True
            if self.tree.identify_region(event.x, event.y) == "heading":
                self.tree.config(cursor="sb_h_double_arrow")
                
                # Visual Marker
                # Identify column under mouse
                col_id = self.tree.identify_column(event.x)
                if col_id:
                    # We can't get exact bbox of header easily.
                    # Just place marker at mouse X
                    self.drop_marker.place(x=event.x - 5, y=0, height=24+5, width=11)
                    self.drop_marker.delete("all")
                    self.drop_marker.create_line(5, 0, 5, 24, width=2, fill="black")
                    self.drop_marker.create_polygon(0, 24, 10, 24, 5, 24+5, fill="black")
                    self.drop_marker.lift()
            else:
                self.drop_marker.place_forget()

    def on_release(self, event):
        self.tree.config(cursor="")
        self.drop_marker.place_forget()
        
        if self.block_menu:
            self.block_menu = False
            return

        region = self.tree.identify("region", event.x, event.y)

        # 1. Resize Logic (Main Table)
        # If released on separator, we must update the total width of the Main Table Frame
        if region == "separator":
            self.recalc_main_table_width()
            # Also triggers image resize as byproduct of space change?
            # Actually, we need to explicitly refresh if main table width changed image table space.
            self.refresh_list() 
            return

        # 2. Drag Reorder
        if self.is_dragging and self.drag_source_col:
            if region == "heading":
                # Find target
                col_name, _, _ = self.get_col_at_x(event.x)
                
                if col_name and col_name in self.visible_columns and self.drag_source_col in self.visible_columns:
                    if self.drag_source_col != "Select" and col_name != "Select":
                        current_idx = self.visible_columns.index(self.drag_source_col)
                        target_idx = self.visible_columns.index(col_name)
                        
                        self.visible_columns.pop(current_idx)
                        self.visible_columns.insert(target_idx, self.drag_source_col)
                        
                        self.setup_data_columns()
                        self.refresh_list()

        # 3. Filter Menu
        elif not self.is_dragging:
            if region == "heading":
                col_name, _, _ = self.get_col_at_x(event.x)
                if col_name and col_name != "Select":
                    self._click_job = self.root.after(200, lambda c=col_name: self.show_filter_menu(c))

        self.is_dragging = False
        self.drag_source_col = None

    def on_img_col_resize(self, event):
        if self.img_tree.identify_region(event.x, event.y) == "separator":
            self.refresh_list()

    def recalc_main_table_width(self):
        # Sum current column widths
        total = 0
        display_cols = self.tree["displaycolumns"]
        for col in display_cols:
            total += self.tree.column(col, "width")
        
        # Configure tree_frame to request this width
        # Note: tree_frame is pack(side=LEFT, fill=Y, expand=False)
        # We need to force its width.
        self.tree_frame.config(width=total)
        # We might need to propagate this change.
        self.tree_frame.update_idletasks()

    def on_double_click(self, event):
        if hasattr(self, '_click_job') and self._click_job:
            self.root.after_cancel(self._click_job)
            self._click_job = None
        
        self.block_menu = True
        self.root.after(500, lambda: setattr(self, 'block_menu', False))

        region = self.tree.identify("region", event.x, event.y)
        
        if region == "separator":
            col_name, _, _ = self.get_col_at_x(event.x - 5)
            if col_name:
                self.autosize_column(col_name)
                # After autosize, update main table width
                self.recalc_main_table_width()
                self.refresh_list()
            return

        if region in ["cell", "tree"]:
            row_id = self.tree.identify_row(event.y)
            if not row_id: return
            
            c_id = self.tree.item(row_id)['tags'][0]
            contact = self.contact_map.get(c_id)
            if not contact: return

            col_name, _, _ = self.get_col_at_x(event.x)
            
            if col_name == "E-mail Address":
                email = contact.get("E-mail Address")
                if email:
                    webbrowser.open(f"mailto:{email}")
                    return
            if col_name == "Select": return
            self.open_editor(contact)

    def autosize_column(self, col_name):
        font_obj = font.Font(font=("Segoe UI", 10))
        max_width = font_obj.measure(col_name + " ▼") + 20
        for c in self.contacts:
            val = c.get(col_name, "")
            w = font_obj.measure(val) + 20
            if w > max_width: max_width = w
            if max_width > 400:
                max_width = 400
                break
        self.tree.column(col_name, width=max_width)

    # ==========================
    # Logic: List Refresh
    # ==========================
    def refresh_list(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for item in self.img_tree.get_children(): self.img_tree.delete(item)
        self.photo_cache = {}

        if self.show_images:
            if not self.img_tree.winfo_ismapped():
                # Pack to LEFT of Scrollbar, Filling space
                self.img_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        else:
            self.img_tree.pack_forget()

        # Update Main Table Width explicitly
        self.recalc_main_table_width()

        filtered = self.contacts
        q = self.search_var.get().lower()
        if q: filtered = [c for c in filtered if q in "".join([str(v) for v in c.values()]).lower()]
        for col, allowed in self.active_filters.items():
            filtered = [c for c in filtered if c.get(col, "").strip() in allowed]

        if not self.show_images:
            if self.current_row_height != TEXT_ROW_HEIGHT:
                self.current_row_height = TEXT_ROW_HEIGHT
                self.style.configure("Treeview", rowheight=TEXT_ROW_HEIGHT)
                self.drop_marker.config(height=TEXT_ROW_HEIGHT)
        else:
            self.img_tree.update_idletasks() 
            col0_width = self.img_tree.column("#0", "width")
            
            target_w = col0_width - 15
            if target_w < 1: target_w = 1
            needed_height = int(target_w / 1.6)
            if needed_height < MIN_ROW_HEIGHT: needed_height = MIN_ROW_HEIGHT
            
            if abs(needed_height - self.current_row_height) > 2:
                self.current_row_height = needed_height
                self.style.configure("Treeview", rowheight=self.current_row_height)
                self.drop_marker.config(height=self.current_row_height)

        col0_width = self.img_tree.column("#0", "width")
        target_w = col0_width - 15
        if target_w < 1: target_w = 1

        for c in filtered:
            c_id = str(c["ID"])
            
            row_img = None
            if self.show_images:
                img_list = c.get("Image Data", [])
                if img_list and col0_width > 20:
                    path = img_list[0]["path"]
                    if os.path.exists(path):
                        cache_key = f"{path}_{target_w}"
                        if cache_key not in self.photo_cache:
                            try:
                                pil = Image.open(path)
                                w_percent = (target_w / float(pil.size[0]))
                                h_size = int((float(pil.size[1]) * float(w_percent)))
                                pil = pil.resize((target_w, h_size), Image.LANCZOS)
                                max_h = self.current_row_height - 4
                                if h_size > max_h:
                                    top = (h_size - max_h) // 2
                                    pil = pil.crop((0, top, target_w, top + max_h))
                                tk_img = ImageTk.PhotoImage(pil)
                                self.photo_cache[cache_key] = tk_img
                            except: pass
                        row_img = self.photo_cache.get(cache_key)

            full_vals = []
            all_cols = ["Select"] + ALL_DATA_COLUMNS
            for col in all_cols:
                if col == "Select":
                    full_vals.append("☑" if c_id in self.selected_ids else "☐")
                else:
                    full_vals.append(c.get(col, ""))

            try:
                self.tree.insert("", tk.END, iid=c_id, values=full_vals, tags=(c_id,))
                self.img_tree.insert("", tk.END, iid=c_id, text="", image=row_img, tags=(c_id,))
            except Exception as e:
                print(f"Insert Error: {e}")
        
        self.update_batch_buttons()

    def refresh_list_visuals_only(self):
        all_cols = ["Select"] + ALL_DATA_COLUMNS
        idx = all_cols.index("Select")
        for item_id in self.tree.get_children():
            c_id = self.tree.item(item_id)['tags'][0]
            vals = list(self.tree.item(item_id)['values'])
            vals[idx] = "☑" if c_id in self.selected_ids else "☐"
            self.tree.item(item_id, values=vals)

    def show_filter_menu(self, col_name):
        values = sorted(list(set([c.get(col_name, "").strip() for c in self.contacts])))
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label=f"Sort A-Z", command=lambda: self.sort_column(col_name, False))
        menu.add_command(label=f"Sort Z-A", command=lambda: self.sort_column(col_name, True))
        menu.add_separator()
        menu.add_command(label="Clear Filter", command=lambda: self.clear_filter(col_name))
        f_menu = tk.Menu(menu, tearoff=0)
        current = self.active_filters.get(col_name, set(values))
        def toggle_val(val):
            if val in current: current.remove(val)
            else: current.add(val)
            if len(current) == len(values): 
                if col_name in self.active_filters: del self.active_filters[col_name]
            else: self.active_filters[col_name] = current
            self.refresh_list()
        for v in values:
            label = v if v else "(Blank)"
            var = tk.BooleanVar(value=(v in current))
            f_menu.add_checkbutton(label=label, variable=var, command=lambda x=v: toggle_val(x))
        menu.add_cascade(label="Filter Values", menu=f_menu)
        menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())

    def toggle_settings_popup(self):
        # TOGGLE LOGIC: Check if we just closed it (via focus out)
        if self.ignore_settings_open:
            self.ignore_settings_open = False
            return

        if self.settings_popup and self.settings_popup.winfo_exists():
            self.settings_popup.destroy()
            return
            
        self.settings_popup = tk.Toplevel(self.root)
        self.settings_popup.overrideredirect(True)
        self.settings_popup.config(bg="white", relief="raised", bd=1)
        x = self.btn_settings.winfo_rootx() + self.btn_settings.winfo_width() - 300
        y = self.btn_settings.winfo_rooty() + self.btn_settings.winfo_height() + 5
        self.settings_popup.geometry(f"300x450+{x}+{y}")
        self.settings_popup.bind("<FocusOut>", lambda e: self.check_focus_out(e))
        self.build_settings_ui(self.settings_popup)
        self.settings_popup.focus_set()

    def check_focus_out(self, event):
        if self.settings_popup and self.settings_popup.winfo_exists():
            try:
                fw = self.root.focus_get()
                if fw and str(fw).startswith(str(self.settings_popup)): return
                if isinstance(fw, tk.Toplevel): return
            except: pass
            
            # If we are closing, set flag to ignore immediate reopen from button click
            self.ignore_settings_open = True
            # Reset flag after short delay
            self.root.after(200, lambda: setattr(self, 'ignore_settings_open', False))
            
            self.settings_popup.destroy()

    def build_settings_ui(self, win):
        tk.Label(win, text="Settings", font=("Segoe UI", 11, "bold"), bg="white").pack(anchor="w", padx=10, pady=10)
        tk.Label(win, text="Working Directory:", font=("Segoe UI", 9, "bold"), bg="white").pack(anchor="w", padx=10)
        d_frame = tk.Frame(win, bg="white")
        d_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        self.popup_dir_var = tk.StringVar(value=self.work_dir)
        tk.Entry(d_frame, textvariable=self.popup_dir_var, bg="#f9f9f9").pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(d_frame, text="...", width=3, command=self.browse_new_dir_popup).pack(side=tk.LEFT, padx=(5,0))
        ttk.Separator(win, orient="horizontal").pack(fill=tk.X, padx=10, pady=5)
        tk.Label(win, text="Visible Columns:", font=("Segoe UI", 9, "bold"), bg="white").pack(anchor="w", padx=10)
        canvas = tk.Canvas(win, bg="white", highlightthickness=0)
        sb = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
        col_frame = tk.Frame(canvas, bg="white")
        col_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=col_frame, anchor="nw", width=260)
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=5)
        sb.pack(side="right", fill="y", pady=5)
        self.popup_vars = {}
        
        var_img = tk.BooleanVar(value=self.show_images)
        self.popup_vars["Image"] = var_img
        tk.Checkbutton(col_frame, text="Image", variable=var_img, bg="white", anchor="w",
                       command=lambda: self.toggle_image_col()).pack(fill=tk.X)

        var_sel = tk.BooleanVar(value="Select" in self.visible_columns)
        self.popup_vars["Select"] = var_sel
        tk.Checkbutton(col_frame, text="Select Box (✓)", variable=var_sel, bg="white", anchor="w",
                       command=lambda: self.toggle_col_popup("Select")).pack(fill=tk.X)
        
        for col in ALL_DATA_COLUMNS:
            var = tk.BooleanVar(value=col in self.visible_columns)
            self.popup_vars[col] = var
            tk.Checkbutton(col_frame, text=col, variable=var, bg="white", anchor="w",
                           command=lambda c=col: self.toggle_col_popup(c)).pack(fill=tk.X)
        tk.Button(win, text="Close Menu", command=win.destroy, bg="#f0f0f0").pack(fill=tk.X, side=tk.BOTTOM)

    def browse_new_dir_popup(self):
        d = filedialog.askdirectory()
        if d:
            self.popup_dir_var.set(d)
            self.work_dir = d
            self.dir_label_var.set(d)
            self.ensure_directories()
            self.contacts = []; self.contact_map = {}; self.selected_ids = set()
            self.load_data(); self.refresh_list()

    def toggle_image_col(self):
        self.show_images = not self.show_images
        self.refresh_list()

    def toggle_col_popup(self, col):
        is_checked = self.popup_vars[col].get()
        if is_checked:
            if col not in self.visible_columns: self.visible_columns.append(col)
        else:
            if col in self.visible_columns: self.visible_columns.remove(col)
        self.setup_data_columns()
        self.refresh_list()

    def sort_column(self, col_name, reverse):
        self.contacts.sort(key=lambda x: x.get(col_name, "").lower(), reverse=reverse)
        self.refresh_list()
    
    def clear_filter(self, col):
        if col in self.active_filters: del self.active_filters[col]
        self.refresh_list()

    def filter_contacts_search(self, *args): self.refresh_list()

    def update_batch_buttons(self):
        count = len(self.selected_ids)
        if count > 0:
            self.btn_delete_selected.pack(side=tk.RIGHT, padx=5)
            self.btn_delete_selected.config(text=f"Delete Selected ({count})")
            
            self.btn_edit_selected.pack(side=tk.RIGHT, padx=5)
            self.btn_edit_selected.config(text=f"Edit Selected ({count})")
        else:
            self.btn_delete_selected.pack_forget()
            self.btn_edit_selected.pack_forget()

    def load_data(self):
        csv_path = os.path.join(self.work_dir, DEFAULT_CSV_NAME)
        if not os.path.exists(csv_path):
            self.contacts = []; self.contact_map = {}; return
        with open(csv_path, mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            self.contacts = list(reader)
        self.contact_map = {}
        for c in self.contacts:
            if not c.get("ID"): c["ID"] = str(uuid.uuid4())
            try: c["Image Data"] = json.loads(c.get("Image Data", "[]"))
            except: c["Image Data"] = []
            try: c["Notes Data"] = json.loads(c.get("Notes Data", "[]"))
            except: c["Notes Data"] = [{"name": "General", "content": c.get("Notes", "")}]
            self.contact_map[c["ID"]] = c

    def save_data(self):
        csv_path = os.path.join(self.work_dir, DEFAULT_CSV_NAME)
        export = []
        for c in self.contacts:
            copy = c.copy()
            copy["Image Data"] = json.dumps(c["Image Data"])
            copy["Notes Data"] = json.dumps(c["Notes Data"])
            export.append(copy)
        with open(csv_path, mode='w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)
            writer.writeheader()
            writer.writerows(export)

    def show_add_options(self):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Manual Creation", command=lambda: self.open_editor())
        menu.add_command(label="From Image", command=self.add_from_image)
        menu.add_command(label="From PDF", command=self.add_from_pdf)
        menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())

    def add_from_image(self):
        files = filedialog.askopenfilenames(filetypes=[("Images", "*.png;*.jpg;*.jpeg")])
        if files: self.process_files(files, False)

    def add_from_pdf(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if files: self.process_files(files, True)

    def process_files(self, files, is_pdf):
        full_text = ""
        img_data = []
        for fp in files:
            processed = []
            if is_pdf:
                try:
                    pil_imgs = convert_from_path(fp, poppler_path=POPPLER_PATH)
                    for i, img in enumerate(pil_imgs):
                        name = f"{os.path.basename(fp).split('.')[0]}_{int(time.time())}_{i}.jpg"
                        save = os.path.join(self.work_dir, IMG_FOLDER_NAME, name)
                        img.save(save, "JPEG")
                        processed.append(save)
                except: pass
            else:
                name = f"import_{int(time.time())}_{os.path.basename(fp)}"
                save = os.path.join(self.work_dir, IMG_FOLDER_NAME, name)
                shutil.copy2(fp, save)
                processed.append(save)
            
            for p in processed:
                try:
                    txt = pytesseract.image_to_string(Image.open(p))
                    full_text += txt + "\n"
                    label = "Back" if len(txt.strip()) < 40 else "Card"
                    img_data.append({"name": label, "path": p})
                except: pass
        
        data = self.heuristic_parse(full_text)
        data["Image Data"] = img_data
        data["Notes Data"] = [{"name": "Card Text", "content": full_text}]
        data["ID"] = str(uuid.uuid4())
        self.open_editor(data)

    def heuristic_parse(self, text):
        data = {k: "" for k in CSV_HEADERS}
        emails = re.findall(r'[\w\.-]+@[\w\.-]+\.\w+', text)
        if emails: data["E-mail Address"] = emails[0]
        phones = re.findall(r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', text)
        if phones: data["Mobile Phone"] = phones[0]
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        if lines:
            parts = lines[0].split()
            if len(parts) > 1:
                data["Last Name"] = parts[-1]
                data["First Name"] = " ".join(parts[:-1])
            else: data["First Name"] = lines[0]
        return data

    def update_contact(self, data):
        cid = data["ID"]
        if cid in self.contact_map:
            idx = self.contacts.index(self.contact_map[cid])
            self.contacts[idx] = data
        else: self.contacts.append(data)
        self.contact_map[cid] = data
        self.save_data()
        self.refresh_list()

    def delete_selected_contacts(self):
        if not self.selected_ids: return
        if messagebox.askyesno("Delete", f"Delete {len(self.selected_ids)} contacts?"):
            self.contacts = [c for c in self.contacts if c["ID"] not in self.selected_ids]
            self.contact_map = {c["ID"]: c for c in self.contacts}
            self.selected_ids.clear()
            self.save_data()
            self.refresh_list()

    def edit_selected_contacts(self):
        if not self.selected_ids: return
        for cid in list(self.selected_ids):
            contact = self.contact_map.get(cid)
            if contact:
                self.open_editor(contact)
    
    def delete_single(self, cid, win):
        if messagebox.askyesno("Confirm", "Delete this contact?"):
            self.contacts = [c for c in self.contacts if c["ID"] != cid]
            if cid in self.contact_map: del self.contact_map[cid]
            if cid in self.selected_ids: self.selected_ids.remove(cid)
            self.save_data()
            self.refresh_list()
            win.destroy()

    def open_editor(self, data=None):
        Editor(self, data)

class Editor:
    def __init__(self, app, data=None):
        self.app = app
        self.data = data if data else {k: "" for k in CSV_HEADERS}
        if "ID" not in self.data: self.data["ID"] = str(uuid.uuid4())
        if "Image Data" not in self.data: self.data["Image Data"] = []
        if "Notes Data" not in self.data: self.data["Notes Data"] = []

        self.win = tk.Toplevel(app.root)
        self.win.title(f"{self.data.get('First Name','')} {self.data.get('Last Name','')}")
        self.win.geometry("1000x800")
        self.create_ui()

    def create_ui(self):
        paned = tk.PanedWindow(self.win, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        left_frame = tk.Frame(paned)
        paned.add(left_frame, width=400, stretch="never")
        
        self.img_notebook = ttk.Notebook(left_frame)
        self.img_notebook.pack(fill=tk.BOTH, expand=True)
        self.img_notebook.bind("<Double-1>", lambda e: self.rename_tab_popup(e, self.img_notebook, "img"))
        self.img_notebook.bind("<Button-3>", lambda e: self.show_menu(e, self.img_notebook, "img"))
        self.render_images()
        tk.Button(left_frame, text="Add Image", command=self.add_image).pack(pady=5)

        right_frame = tk.Frame(paned)
        paned.add(right_frame, stretch="always")
        
        btn_frame = tk.Frame(right_frame, pady=20)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        tk.Button(btn_frame, text="Save", bg="#4CAF50", fg="white", font=("Segoe UI", 9), command=self.save).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn_frame, text="Cancel", font=("Segoe UI", 9), command=self.win.destroy).pack(side=tk.RIGHT, padx=5)
        
        cid = self.data.get("ID")
        is_existing = any(c["ID"] == cid for c in self.app.contacts)
        if is_existing:
             tk.Button(btn_frame, text="Delete", bg="#E53935", fg="white", font=("Segoe UI", 9),
                        command=lambda: self.app.delete_single(cid, self.win)).pack(side=tk.LEFT)

        canvas = tk.Canvas(right_frame, highlightthickness=0)
        sb = ttk.Scrollbar(right_frame, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas)
        
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_win = canvas.create_window((0,0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        
        def on_resize(e): canvas.itemconfig(canvas_win, width=e.width)
        canvas.bind("<Configure>", on_resize)
        
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.entries = {}
        fields = ["First Name", "Last Name", "Company", "Job Title", 
                  "E-mail Address", "Mobile Phone", "Business Phone"]
        for f in fields:
            tk.Label(scroll_frame, text=f, font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(5,0))
            ent = tk.Entry(scroll_frame, font=("Segoe UI", 10))
            ent.insert(0, self.data.get(f, ""))
            ent.pack(fill=tk.X)
            self.entries[f] = ent

        tk.Label(scroll_frame, text="Address", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(5,0))
        addr = tk.Text(scroll_frame, height=4, font=("Segoe UI", 10))
        addr.insert("1.0", self.data.get("Address", ""))
        addr.pack(fill=tk.X)
        self.entries["Address"] = addr
        
        tk.Label(scroll_frame, text="Notes", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(10,0))
        self.note_notebook = ttk.Notebook(scroll_frame)
        self.note_notebook.pack(fill=tk.BOTH, expand=True, pady=5)
        self.note_notebook.bind("<Button-1>", self.click_note_tab)
        self.note_notebook.bind("<Double-1>", lambda e: self.rename_tab_popup(e, self.note_notebook, "note"))
        self.note_notebook.bind("<Button-3>", lambda e: self.show_menu(e, self.note_notebook, "note"))
        self.note_widgets = []
        self.render_notes()

    def ask_tab_name(self, title, prompt, initial_value=""):
        dialog = tk.Toplevel(self.win)
        dialog.title(title)
        dialog.transient(self.win)
        dialog.grab_set()
        
        x = self.win.winfo_pointerx()
        y = self.win.winfo_pointery()
        dialog.geometry(f"+{x}+{y}")
        
        tk.Label(dialog, text=prompt).pack(padx=10, pady=5)
        entry = tk.Entry(dialog)
        entry.insert(0, initial_value)
        entry.select_range(0, tk.END)
        entry.pack(padx=10, pady=5)
        entry.focus_set()
        
        result = [None]
        
        def on_ok(event=None):
            result[0] = entry.get()
            dialog.destroy()
            
        def on_cancel(event=None):
            dialog.destroy()

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(fill=tk.X, pady=5)
        tk.Button(btn_frame, text="OK", width=8, command=on_ok).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Cancel", width=8, command=on_cancel).pack(side=tk.RIGHT, padx=10)
        
        entry.bind("<Return>", on_ok)
        entry.bind("<Escape>", on_cancel)
        
        self.win.wait_window(dialog)
        return result[0]

    def rename_tab_popup(self, event=None, notebook=None, type_=None, tab_index=None):
        try:
            if tab_index is None and event:
                try: 
                    if notebook.identify(event.x, event.y) == 'label':
                        tab_index = notebook.index(f"@{event.x},{event.y}")
                    else: return
                except: return
            
            if tab_index is None: return
            if type_ == "note" and tab_index == len(notebook.tabs()) - 1: return

            old_name = notebook.tab(tab_index, "text")
            new_name = self.ask_tab_name("Rename Tab", "New Name:", old_name)
            
            if new_name:
                notebook.tab(tab_index, text=new_name)
                if type_ == "img":
                    self.data["Image Data"][tab_index]["name"] = new_name
                elif type_ == "note":
                    self.data["Notes Data"][tab_index]["name"] = new_name
        except Exception as e:
            pass

    def show_menu(self, event, notebook, type_):
        try:
            index = notebook.index(f"@{event.x},{event.y}")
            if type_ == "note" and index == len(notebook.tabs()) - 1: return
            
            m = tk.Menu(self.win, tearoff=0)
            m.add_command(label="Rename", command=lambda: self.rename_tab_popup(None, notebook, type_, index))
            m.add_command(label="Delete", command=lambda: self.del_tab(notebook, index, type_))
            m.tk_popup(event.x_root, event.y_root)
        except: pass

    def del_tab(self, nb, idx, type_):
        if type_ == "note":
            del self.data["Notes Data"][idx]
            self.render_notes()
        else:
            del self.data["Image Data"][idx]
            self.render_images()

    def render_images(self):
        for t in self.img_notebook.tabs(): self.img_notebook.forget(t)
        imgs = self.data.get("Image Data", [])
        if not imgs: self.img_notebook.add(tk.Label(self.img_notebook, text="None"), text="No Img")
        for i, img in enumerate(imgs):
            f = tk.Frame(self.img_notebook)
            self.img_notebook.add(f, text=img.get("name", "Img"))
            if os.path.exists(img["path"]):
                try:
                    pil = Image.open(img["path"])
                    pil.thumbnail((350, 450))
                    photo = ImageTk.PhotoImage(pil)
                    l = tk.Label(f, image=photo); l.image=photo; l.pack()
                except: pass

    def add_image(self):
        # Allow multiple selection
        ftypes = [("All Files", "*.*"), ("Images", "*.png;*.jpg;*.jpeg"), ("PDF", "*.pdf")]
        files = filedialog.askopenfilenames(filetypes=ftypes)
        
        if not files: return

        # Ask for tab name base
        base_name = self.ask_tab_name("Tab Name", "Enter new image tab name:", "Doc")
        if not base_name: base_name = "Doc"

        new_images = []
        
        for f in files:
            if f.lower().endswith(".pdf"):
                try:
                    pil_imgs = convert_from_path(f, poppler_path=POPPLER_PATH)
                    for i, img in enumerate(pil_imgs):
                        name_suffix = f" ({i+1})" if len(pil_imgs) > 1 else ""
                        tab_name = f"{base_name}{name_suffix}"
                        
                        fname = f"doc_{int(time.time())}_{i}.jpg"
                        save_path = os.path.join(self.app.work_dir, IMG_FOLDER_NAME, fname)
                        img.save(save_path, "JPEG")
                        
                        new_images.append({"name": tab_name, "path": save_path})
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to convert PDF: {e}")
            else:
                # Image
                fname = f"img_{int(time.time())}_{os.path.basename(f)}"
                save_path = os.path.join(self.app.work_dir, IMG_FOLDER_NAME, fname)
                shutil.copy2(f, save_path)
                # If multiple files selected, apply numbering
                new_images.append({"name": base_name, "path": save_path})

        if len(new_images) > 1:
            for i, img_d in enumerate(new_images):
                img_d["name"] = f"{base_name} ({i+1})"
        
        self.data["Image Data"].extend(new_images)
        self.render_images()

    def render_notes(self):
        for t in self.note_notebook.tabs(): self.note_notebook.forget(t)
        self.note_widgets = []
        notes = self.data.get("Notes Data", [])
        if not notes: notes = [{"name":"Gen", "content":""}]
        for n in notes:
            f = tk.Frame(self.note_notebook)
            self.note_notebook.add(f, text=n.get("name", "Note"))
            t = tk.Text(f, height=10); t.insert("1.0", n.get("content","")); t.pack(fill=tk.BOTH, expand=True)
            self.note_widgets.append(t)
        self.note_notebook.add(tk.Frame(self.note_notebook), text="+")

    def click_note_tab(self, e):
        try:
            if self.note_notebook.index(f"@{e.x},{e.y}") == len(self.note_notebook.tabs())-1:
                new_name = datetime.now().strftime("%m/%d/%Y")
                self.data["Notes Data"].append({"name": new_name, "content":""})
                self.render_notes()
                self.note_notebook.select(len(self.note_widgets)-1)
        except: pass

    def save(self):
        for f, w in self.entries.items():
            if isinstance(w, tk.Entry): self.data[f] = w.get().strip()
            else: self.data[f] = w.get("1.0", tk.END).strip()
        
        nn = []
        for i, w in enumerate(self.note_widgets):
            nn.append({"name": self.data["Notes Data"][i]["name"], "content": w.get("1.0", tk.END).strip()})
        self.data["Notes Data"] = nn
        
        self.app.update_contact(self.data)
        self.win.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = RolodexApp(root)
    root.mainloop()
