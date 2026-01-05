import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from PIL import Image, ImageTk
import csv
import os
import pytesseract
import re
from pdf2image import convert_from_path
import time
import json
import shutil

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

CSV_HEADERS = [
    "First Name", "Last Name", "Company", "Job Title", 
    "E-mail Address", "Mobile Phone", "Business Phone", 
    "Address", "Notes Data", "Image Data"
]

class RolodexApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Modern Rolodex")
        self.root.geometry("1100x750")
        
        self.work_dir = os.getcwd()
        self.contacts = []
        self.sort_reverse = False
        self.selected_indices = set()

        self.setup_styles()
        self.ensure_directories()
        
        # --- UI Construction ---
        self.create_toolbar()
        self.create_list_view()
        
        # --- Data Load ---
        self.load_data()
        self.refresh_list()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        bg_color = "#f4f4f4"
        text_color = "#333333"
        
        self.root.configure(bg=bg_color)
        self.style.configure("TFrame", background=bg_color)
        self.style.configure("TLabel", background=bg_color, foreground=text_color, font=("Segoe UI", 10))
        self.style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"))
        self.style.configure("TButton", font=("Segoe UI", 9), padding=5)
        
        # Green Button (+)
        self.style.configure("Green.TButton", background="#4CAF50", foreground="white", font=("Segoe UI", 11, "bold"))
        self.style.map("Green.TButton", background=[('active', '#43A047')])
        
        # Red Button (Delete)
        self.style.configure("Red.TButton", background="#E53935", foreground="white", font=("Segoe UI", 9, "bold"))
        self.style.map("Red.TButton", background=[('active', '#D32F2F')])

        # Treeview
        self.style.configure("Treeview", background="white", foreground=text_color, rowheight=30, fieldbackground="white", font=("Segoe UI", 10))
        self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#e0e0e0")
        self.style.map("Treeview", background=[('selected', '#2196F3')], foreground=[('selected', 'white')])

    def ensure_directories(self):
        img_dir = os.path.join(self.work_dir, IMG_FOLDER_NAME)
        if not os.path.exists(img_dir):
            os.makedirs(img_dir)

    # ==========================
    # UI Layout
    # ==========================
    def create_toolbar(self):
        # One main toolbar frame
        self.toolbar = tk.Frame(self.root, bg="#e0e0e0", pady=8, padx=10, relief="flat")
        self.toolbar.pack(side=tk.TOP, fill=tk.X)

        # 1. Plus Button (Green)
        ttk.Button(self.toolbar, text="+", style="Green.TButton", width=4, 
                   command=self.show_add_options).pack(side=tk.LEFT, padx=(0, 15))

        # 2. Search
        tk.Label(self.toolbar, text="Search:", bg="#e0e0e0", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_contacts)
        tk.Entry(self.toolbar, textvariable=self.search_var, font=("Segoe UI", 10), width=20).pack(side=tk.LEFT, padx=5)

        # 3. Working Directory (Moved here)
        tk.Label(self.toolbar, text="|   Dir:", bg="#e0e0e0", fg="#666").pack(side=tk.LEFT, padx=(15, 5))
        self.dir_var = tk.StringVar(value=self.work_dir)
        # Using a read-only entry for display
        tk.Entry(self.toolbar, textvariable=self.dir_var, state="readonly", width=30, 
                 font=("Segoe UI", 8), disabledbackground="#f0f0f0").pack(side=tk.LEFT)
        ttk.Button(self.toolbar, text="Browse", width=6, command=self.change_directory).pack(side=tk.LEFT, padx=5)

        # 4. Delete Selected Button (Hidden by default, right aligned)
        self.del_btn_frame = tk.Frame(self.toolbar, bg="#e0e0e0")
        self.del_btn_frame.pack(side=tk.RIGHT)
        
        self.btn_delete_selected = ttk.Button(self.del_btn_frame, text="Delete Selected", style="Red.TButton", 
                                              command=self.delete_selected_contacts)
        # We don't pack it yet. We pack it when items are selected.

    def create_list_view(self):
        columns = ("Select", "Name", "Company", "Email", "Phone")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", selectmode="none") # selectmode none to handle custom selection logic
        
        self.tree.heading("Select", text="✓", command=self.toggle_select_all)
        self.tree.heading("Name", text="Name", command=lambda: self.sort_column("Name"))
        self.tree.heading("Company", text="Company", command=lambda: self.sort_column("Company"))
        self.tree.heading("Email", text="Email", command=lambda: self.sort_column("Email"))
        self.tree.heading("Phone", text="Phone")
        
        self.tree.column("Select", width=40, anchor="center", stretch=False)
        self.tree.column("Name", width=200)
        self.tree.column("Company", width=200)
        self.tree.column("Email", width=250)
        self.tree.column("Phone", width=150)

        self.tree.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        scrollbar = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        # Bindings
        self.tree.bind("<Button-1>", self.on_click)
        self.tree.bind("<Double-1>", self.on_double_click)

    # ==========================
    # Logic: Interaction
    # ==========================
    def change_directory(self):
        new_dir = filedialog.askdirectory()
        if new_dir:
            self.work_dir = new_dir
            self.dir_var.set(self.work_dir)
            self.ensure_directories()
            self.contacts = []
            self.selected_indices = set()
            self.load_data()
            self.refresh_list()

    def on_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            col = self.tree.identify_column(event.x)
            row_id = self.tree.identify_row(event.y)
            if not row_id: return
            
            # If they clicked the checkbox column ("#1")
            if col == "#1":
                item = self.tree.item(row_id)
                index = int(item['tags'][0]) # Get the real index from tag
                
                if index in self.selected_indices:
                    self.selected_indices.remove(index)
                else:
                    self.selected_indices.add(index)
                
                self.refresh_list_visuals_only()
                self.update_delete_button_visibility()

    def on_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        
        col = self.tree.identify_column(event.x)
        if col == "#1": return # Ignore double click on checkbox
        
        row_id = self.tree.identify_row(event.y)
        if row_id:
            item = self.tree.item(row_id)
            index = int(item['tags'][0])
            self.open_editor(self.contacts[index], contact_index=index)

    def toggle_select_all(self):
        # If all currently visible are selected, deselect all. Else select all.
        visible_ids = self.tree.get_children()
        visible_indices = [int(self.tree.item(x)['tags'][0]) for x in visible_ids]
        
        if not visible_indices: return

        # Check if all visible are in selected
        all_selected = all(idx in self.selected_indices for idx in visible_indices)
        
        if all_selected:
            for idx in visible_indices:
                self.selected_indices.discard(idx)
        else:
            for idx in visible_indices:
                self.selected_indices.add(idx)
        
        self.refresh_list_visuals_only()
        self.update_delete_button_visibility()

    def update_delete_button_visibility(self):
        if len(self.selected_indices) > 0:
            self.btn_delete_selected.pack(side=tk.RIGHT, padx=10)
            self.btn_delete_selected.config(text=f"Delete Selected ({len(self.selected_indices)})")
        else:
            self.btn_delete_selected.pack_forget()

    # ==========================
    # Logic: Data
    # ==========================
    def load_data(self):
        csv_path = os.path.join(self.work_dir, DEFAULT_CSV_NAME)
        if not os.path.exists(csv_path):
            self.contacts = []
            return

        with open(csv_path, mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            self.contacts = list(reader)

        # JSON parsing
        for c in self.contacts:
            try: c["Image Data"] = json.loads(c.get("Image Data", "[]"))
            except: c["Image Data"] = []
            try: c["Notes Data"] = json.loads(c.get("Notes Data", "[]"))
            except: c["Notes Data"] = [{"name": "General", "content": c.get("Notes", "")}]

    def save_data(self):
        csv_path = os.path.join(self.work_dir, DEFAULT_CSV_NAME)
        export_contacts = []
        for c in self.contacts:
            c_copy = c.copy()
            c_copy["Image Data"] = json.dumps(c["Image Data"])
            c_copy["Notes Data"] = json.dumps(c["Notes Data"])
            export_contacts.append(c_copy)

        with open(csv_path, mode='w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)
            writer.writeheader()
            writer.writerows(export_contacts)

    def refresh_list(self, filtered_contacts=None):
        for item in self.tree.get_children():
            self.tree.delete(item)

        data = filtered_contacts if filtered_contacts is not None else self.contacts

        for contact in data:
            real_index = self.contacts.index(contact) # Identify by index in master list
            full_name = f"{contact.get('First Name', '')} {contact.get('Last Name', '')}".strip()
            
            icon = "☑" if real_index in self.selected_indices else "☐"

            self.tree.insert("", tk.END, values=(
                icon,
                full_name,
                contact.get("Company", ""),
                contact.get("E-mail Address", ""),
                contact.get("Mobile Phone", "")
            ), tags=(str(real_index),))
        
        self.update_delete_button_visibility()

    def refresh_list_visuals_only(self):
        for item_id in self.tree.get_children():
            tag_index = int(self.tree.item(item_id)['tags'][0])
            vals = list(self.tree.item(item_id)['values'])
            vals[0] = "☑" if tag_index in self.selected_indices else "☐"
            self.tree.item(item_id, values=vals)

    def filter_contacts(self, *args):
        query = self.search_var.get().lower()
        if not query:
            self.refresh_list()
            return
        filtered = [c for c in self.contacts if query in "".join([str(v) for v in c.values()]).lower()]
        self.refresh_list(filtered)

    def sort_column(self, col_name):
        key_map = {"Name": "First Name", "Company": "Company", "Email": "E-mail Address"}
        key = key_map.get(col_name, "First Name")
        self.contacts.sort(key=lambda x: x.get(key, "").lower(), reverse=self.sort_reverse)
        self.sort_reverse = not self.sort_reverse
        self.refresh_list()

    # ==========================
    # Logic: Add/Delete/Merge
    # ==========================
    def show_add_options(self):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Manual Creation", command=lambda: self.open_editor())
        menu.add_command(label="From Image", command=self.add_from_image)
        menu.add_command(label="From PDF", command=self.add_from_pdf)
        try:
            # Locate button
            x, y, _, h = self.toolbar.bbox(self.toolbar.winfo_children()[0]) # Approximation
            # Better: use pointer
            menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())
        finally:
            menu.grab_release()

    def add_from_image(self):
        files = filedialog.askopenfilenames(filetypes=[("Images", "*.png;*.jpg;*.jpeg")])
        if files: self.process_files(files, is_pdf=False)

    def add_from_pdf(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if files: self.process_files(files, is_pdf=True)

    def process_files(self, file_paths, is_pdf):
        img_data_list = []
        full_text = ""
        
        for fp in file_paths:
            processed_images = []
            if is_pdf:
                try:
                    pil_images = convert_from_path(fp, poppler_path=POPPLER_PATH)
                    for i, img in enumerate(pil_images):
                        name = f"{os.path.basename(fp).split('.')[0]}_{int(time.time())}_{i}.jpg"
                        save_path = os.path.join(self.work_dir, IMG_FOLDER_NAME, name)
                        img.save(save_path, "JPEG")
                        processed_images.append(save_path)
                except Exception as e:
                    messagebox.showerror("Error", f"PDF Error: {e}")
                    continue
            else:
                name = f"import_{int(time.time())}_{os.path.basename(fp)}"
                save_path = os.path.join(self.work_dir, IMG_FOLDER_NAME, name)
                shutil.copy2(fp, save_path)
                processed_images.append(save_path)

            for img_path in processed_images:
                try:
                    text = pytesseract.image_to_string(Image.open(img_path))
                    full_text += text + "\n"
                    label = "Back" if len(text.strip()) < 40 else "Card"
                    img_data_list.append({"name": label, "path": img_path})
                except: pass

        data = self.heuristic_parse(full_text)
        data["Image Data"] = img_data_list
        data["Notes Data"] = [{"name": "OCR Text", "content": full_text}]
        self.check_and_open_editor(data)

    def heuristic_parse(self, text):
        data = {k: "" for k in CSV_HEADERS}
        data["Image Data"] = []
        data["Notes Data"] = []

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
            else:
                data["First Name"] = lines[0]
        return data

    def check_and_open_editor(self, new_data):
        new_name = f"{new_data.get('First Name','')} {new_data.get('Last Name','')}".strip()
        
        duplicate = None
        dup_idx = -1
        for i, c in enumerate(self.contacts):
            c_name = f"{c.get('First Name','')} {c.get('Last Name','')}".strip()
            if c_name.lower() == new_name.lower() and c_name:
                duplicate = c
                dup_idx = i
                break
        
        if duplicate:
            if messagebox.askyesno("Duplicate", f"'{new_name}' exists. Combine?"):
                self.resolve_merge(duplicate, new_data, dup_idx)
                return
        self.open_editor(existing_data=new_data)

    def resolve_merge(self, old_data, new_data, idx):
        merged = old_data.copy()
        conflicts = []
        
        merged["Image Data"] = old_data.get("Image Data", []) + new_data.get("Image Data", [])
        merged["Notes Data"] = old_data.get("Notes Data", []) + new_data.get("Notes Data", [])

        fields = ["Company", "Job Title", "E-mail Address", "Mobile Phone", "Business Phone", "Address"]
        for f in fields:
            o = old_data.get(f, "").strip()
            n = new_data.get(f, "").strip()
            if not o and n: merged[f] = n
            elif o and n and o != n: conflicts.append((f, o, n))

        if conflicts:
            dialog = tk.Toplevel(self.root)
            dialog.title("Merge Conflicts")
            tk.Label(dialog, text="Select info to keep:").pack(pady=10)
            vars_map = {}
            for f, o, n in conflicts:
                frm = tk.Frame(dialog); frm.pack(fill=tk.X, padx=10)
                tk.Label(frm, text=f"{f}:", width=15, anchor="e").pack(side=tk.LEFT)
                v = tk.StringVar(value="old")
                vars_map[f] = (v, o, n)
                tk.Radiobutton(frm, text=f"Old: {o}", variable=v, value="old").pack(side=tk.LEFT)
                tk.Radiobutton(frm, text=f"New: {n}", variable=v, value="new").pack(side=tk.LEFT)
            
            def finish():
                for f, (v, o, n) in vars_map.items(): merged[f] = o if v.get()=="old" else n
                self.contacts[idx] = merged
                self.save_data(); self.refresh_list(); dialog.destroy()
                self.open_editor(merged, contact_index=idx)
            
            tk.Button(dialog, text="Confirm", command=finish).pack(pady=10)
        else:
            self.contacts[idx] = merged
            self.save_data(); self.refresh_list()
            self.open_editor(merged, contact_index=idx)

    def open_editor(self, existing_data=None, contact_index=None):
        Editor(self, existing_data, contact_index)

    def delete_contact(self, index, window):
        if messagebox.askyesno("Confirm", "Delete this contact?"):
            # Remove from selected if it was there
            self.selected_indices.discard(index)
            # We need to shift selected indices that are higher than this one down by one?
            # Actually easier to just clear selection or handle rigorously.
            # Let's just clear selection to prevent index errors
            self.selected_indices.clear()
            
            del self.contacts[index]
            self.save_data()
            self.refresh_list()
            window.destroy()

    def delete_selected_contacts(self):
        if not self.selected_indices: return
        if messagebox.askyesno("Delete Selected", f"Delete {len(self.selected_indices)} contacts?"):
            for i in sorted(list(self.selected_indices), reverse=True):
                del self.contacts[i]
            self.selected_indices.clear()
            self.save_data()
            self.refresh_list()


# ==========================================
# EDITOR CLASS
# ==========================================
class Editor:
    def __init__(self, app, data=None, index=None):
        self.app = app
        self.data = data if data else {k: "" for k in CSV_HEADERS}
        if "Image Data" not in self.data: self.data["Image Data"] = []
        if "Notes Data" not in self.data: self.data["Notes Data"] = []
        
        self.index = index
        self.is_new = index is None

        self.win = tk.Toplevel(app.root)
        
        # Set Title
        fname = self.data.get("First Name", "").strip()
        lname = self.data.get("Last Name", "").strip()
        title = f"{fname} {lname}".strip()
        if not title: title = "New Contact"
        self.win.title(title)
        
        self.win.geometry("950x800")
        
        self.create_ui()

    def create_ui(self):
        main_paned = tk.PanedWindow(self.win, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Left: Images
        self.left_frame = tk.Frame(main_paned)
        main_paned.add(self.left_frame, width=400)
        
        self.img_notebook = ttk.Notebook(self.left_frame)
        self.img_notebook.pack(fill=tk.BOTH, expand=True)
        self.img_notebook.bind("<Double-1>", self.rename_tab)
        # Right Click Context Menu for Images
        self.img_notebook.bind("<Button-3>", lambda e: self.show_context_menu(e, self.img_notebook, "image"))
        
        self.render_images()

        btn_img_frame = tk.Frame(self.left_frame)
        btn_img_frame.pack(fill=tk.X, pady=5)
        # Note: Removing the 'Remove Image' button since we added right-click delete
        # but can keep it for accessibility if desired. Let's keep a simplified Add Image button?
        # User prompt didn't strictly ask to remove the button, just add right click context.
        # But previously we had a remove button. I'll leave it as fallback.
        tk.Button(btn_img_frame, text="Add Image to Contact", command=self.add_image_to_existing).pack(side=tk.LEFT)

        # Right: Inputs
        self.right_frame = tk.Frame(main_paned)
        main_paned.add(self.right_frame)

        # Scrollable Canvas
        canvas = tk.Canvas(self.right_frame, highlightthickness=0) # Fix visual glitch 7
        scrollbar = ttk.Scrollbar(self.right_frame, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas)

        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Make columns expandable
        scroll_frame.columnconfigure(0, weight=1)

        # Fields
        self.entries = {}
        row = 0
        fields = ["First Name", "Last Name", "Company", "Job Title", 
                  "E-mail Address", "Mobile Phone", "Business Phone"]
        
        for f in fields:
            tk.Label(scroll_frame, text=f, font=("Segoe UI", 9, "bold")).grid(row=row, column=0, sticky="w", pady=(5,0))
            ent = tk.Entry(scroll_frame, font=("Segoe UI", 10))
            ent.insert(0, self.data.get(f, ""))
            ent.grid(row=row+1, column=0, sticky="ew", padx=5)
            self.entries[f] = ent
            row += 2

        tk.Label(scroll_frame, text="Address", font=("Segoe UI", 9, "bold")).grid(row=row, column=0, sticky="w", pady=(5,0))
        addr_txt = tk.Text(scroll_frame, height=4, font=("Segoe UI", 10))
        addr_txt.insert("1.0", self.data.get("Address", ""))
        addr_txt.grid(row=row+1, column=0, sticky="ew", padx=5)
        self.entries["Address"] = addr_txt
        row += 2

        # Notes - Make it fill remaining space
        tk.Label(scroll_frame, text="Notes", font=("Segoe UI", 9, "bold")).grid(row=row, column=0, sticky="w", pady=(10,0))
        
        # We give the note notebook row a high weight to push it down if canvas expands?
        # Actually in a scroll frame, "remaining space" is just expanding downwards.
        # We'll set a larger default height and allow expansion.
        scroll_frame.rowconfigure(row+1, weight=1) 
        
        self.note_notebook = ttk.Notebook(scroll_frame)
        self.note_notebook.grid(row=row+1, column=0, sticky="nsew", padx=5, pady=5) # sticky nsew
        
        # Ensure the canvas and window allow the scroll frame to take height
        # This is tricky in Tkinter Scrollable frames. 
        # A compromise: Set a min-height for the notebook, and if the user resizes the window, 
        # the canvas expands, but the scroll_frame inside stays "top aligned" unless we config canvas window.
        # To truly make it grow with window, we need to resize the canvas window on configure.
        def on_canvas_configure(event):
            canvas.itemconfig(win_id, width=event.width)
            # Only force height if content is smaller than window? 
            # Standard behavior is usually sufficient: just fixed height notes or semi-large.
            # I will set a generous height for the text widgets inside.
        
        win_id = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.bind("<Configure>", on_canvas_configure)

        self.note_notebook.bind("<Button-1>", self.on_note_tab_click)
        self.note_notebook.bind("<Double-1>", self.rename_note_tab)
        # Right Click Context Menu for Notes
        self.note_notebook.bind("<Button-3>", lambda e: self.show_context_menu(e, self.note_notebook, "note"))

        self.note_widgets = [] 
        self.render_notes()

        # Bottom Actions
        action_frame = tk.Frame(self.win, pady=10, bg="#f4f4f4")
        action_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        tk.Button(action_frame, text="Save Contact", bg="#4CAF50", fg="white", font=("Segoe UI", 10, "bold"),
                  command=self.save).pack(side=tk.RIGHT, padx=10)
        tk.Button(action_frame, text="Cancel", command=self.win.destroy).pack(side=tk.RIGHT, padx=10)

        if not self.is_new:
            ttk.Button(action_frame, text="Delete Contact", style="Red.TButton", 
                       command=lambda: self.app.delete_contact(self.index, self.win)).pack(side=tk.LEFT, padx=10)

    # --- Context Menus ---
    def show_context_menu(self, event, notebook, context_type):
        # Identify tab
        try:
            index = notebook.index(f"@{event.x},{event.y}")
            # If notes, protect the "+" tab (last one)
            if context_type == "note" and index == len(notebook.tabs()) - 1:
                return 

            menu = tk.Menu(self.win, tearoff=0)
            menu.add_command(label="Delete Tab", command=lambda: self.delete_tab(notebook, index, context_type))
            menu.tk_popup(event.x_root, event.y_root)
        except:
            pass # Clicked outside a tab

    def delete_tab(self, notebook, index, context_type):
        if context_type == "note":
            # Remove from data
            if index < len(self.data["Notes Data"]):
                del self.data["Notes Data"][index]
            self.render_notes()
        elif context_type == "image":
            if index < len(self.data["Image Data"]):
                del self.data["Image Data"][index]
            self.render_images()

    # --- Image Logic ---
    def add_image_to_existing(self):
        f = filedialog.askopenfilename(filetypes=[("Images", "*.jpg;*.png")])
        if f:
            name = f"added_{int(time.time())}.jpg"
            save_path = os.path.join(self.app.work_dir, IMG_FOLDER_NAME, name)
            shutil.copy2(f, save_path)
            self.data["Image Data"].append({"name": "Card", "path": save_path})
            self.render_images()

    def render_images(self):
        for tab in self.img_notebook.tabs(): self.img_notebook.forget(tab)
        images = self.data.get("Image Data", [])
        if not images:
            self.img_notebook.add(tk.Label(self.img_notebook, text="No Images"), text="None")
            return
        
        for i, img_entry in enumerate(images):
            frame = tk.Frame(self.img_notebook)
            self.img_notebook.add(frame, text=img_entry.get("name", f"Img {i+1}"))
            path = img_entry["path"]
            if os.path.exists(path):
                try:
                    pil = Image.open(path)
                    pil.thumbnail((380, 500))
                    tk_img = ImageTk.PhotoImage(pil)
                    lbl = tk.Label(frame, image=tk_img)
                    lbl.image = tk_img
                    lbl.pack(expand=True)
                except: tk.Label(frame, text="Error").pack()
            else: tk.Label(frame, text="File Not Found").pack()

    def rename_tab(self, event):
        try:
            index = self.img_notebook.index(f"@{event.x},{event.y}")
            new_name = simpledialog.askstring("Rename", "New Name:")
            if new_name:
                self.img_notebook.tab(index, text=new_name)
                self.data["Image Data"][index]["name"] = new_name
        except: pass

    # --- Notes Logic ---
    def render_notes(self):
        for tab in self.note_notebook.tabs(): self.note_notebook.forget(tab)
        self.note_widgets = []
        notes = self.data.get("Notes Data", [])
        if not notes:
            notes = [{"name": "General", "content": ""}]
            self.data["Notes Data"] = notes
        
        for n in notes:
            frame = tk.Frame(self.note_notebook)
            self.note_notebook.add(frame, text=n.get("name", "Note"))
            # High default height to fill screen space roughly
            txt = tk.Text(frame, height=20, font=("Segoe UI", 10)) 
            txt.insert("1.0", n.get("content", ""))
            txt.pack(fill=tk.BOTH, expand=True)
            self.note_widgets.append(txt)
        
        # Plus Tab
        self.note_notebook.add(tk.Frame(self.note_notebook), text="+")

    def on_note_tab_click(self, event):
        try:
            index = self.note_notebook.index(f"@{event.x},{event.y}")
            if index == len(self.note_notebook.tabs()) - 1:
                self.data["Notes Data"].append({"name": "New Note", "content": ""})
                self.render_notes()
                self.note_notebook.select(len(self.note_widgets)-1)
        except: pass

    def rename_note_tab(self, event):
        try:
            index = self.note_notebook.index(f"@{event.x},{event.y}")
            if index == len(self.note_notebook.tabs()) - 1: return
            new_name = simpledialog.askstring("Rename", "Title:")
            if new_name:
                self.note_notebook.tab(index, text=new_name)
                self.data["Notes Data"][index]["name"] = new_name
        except: pass

    def save(self):
        # Update Data from UI
        for f, w in self.entries.items():
            if isinstance(w, tk.Entry): self.data[f] = w.get().strip()
            else: self.data[f] = w.get("1.0", tk.END).strip()
        
        # Update Notes
        new_notes = []
        for i, widget in enumerate(self.note_widgets):
            title = self.data["Notes Data"][i]["name"]
            content = widget.get("1.0", tk.END).strip()
            new_notes.append({"name": title, "content": content})
        self.data["Notes Data"] = new_notes

        if self.is_new: self.app.contacts.append(self.data)
        else: self.app.contacts[self.index] = self.data
        
        try:
            self.app.save_data()
            self.app.refresh_list()
            self.win.destroy() # Close silent save (Req 2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = RolodexApp(root)
    root.mainloop()