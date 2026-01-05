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

# Headers - Note: "Image Data" and "Notes Data" store JSON strings now
CSV_HEADERS = [
    "First Name", "Last Name", "Company", "Job Title", 
    "E-mail Address", "Mobile Phone", "Business Phone", 
    "Address", "Notes Data", "Image Data"
]

class RolodexApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Modern Rolodex")
        self.root.geometry("1000x750")
        
        # State
        self.work_dir = os.getcwd()
        self.contacts = []
        self.sort_reverse = False
        self.selection_mode = False
        self.selected_indices = set() # For bulk delete

        # Style Setup
        self.setup_styles()
        
        # UI Construction
        self.create_top_nav()
        self.create_main_toolbar()
        self.create_bulk_toolbar() # Hidden by default
        self.create_list_view()
        
        # Initial Load
        self.ensure_directories()
        self.load_data()
        self.refresh_list()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        # Colors
        bg_color = "#f4f4f4"
        accent_color = "#2196F3"
        green_color = "#4CAF50"
        red_color = "#E53935"
        text_color = "#333333"

        self.root.configure(bg=bg_color)
        
        # Generic Frame
        self.style.configure("TFrame", background=bg_color)
        self.style.configure("TLabel", background=bg_color, foreground=text_color, font=("Segoe UI", 10))
        self.style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"))
        
        # Buttons
        self.style.configure("TButton", font=("Segoe UI", 10), padding=6, background="#e0e0e0")
        
        # Green Button (+)
        self.style.configure("Green.TButton", background=green_color, foreground="white", font=("Segoe UI", 12, "bold"))
        self.style.map("Green.TButton", background=[('active', '#43A047')])
        
        # Red Button (Delete)
        self.style.configure("Red.TButton", background=red_color, foreground="white", font=("Segoe UI", 10, "bold"))
        self.style.map("Red.TButton", background=[('active', '#D32F2F')])

        # Treeview (Modern Look)
        self.style.configure("Treeview", 
                             background="white",
                             foreground=text_color, 
                             rowheight=30, 
                             fieldbackground="white",
                             font=("Segoe UI", 10))
        self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#e0e0e0")
        self.style.map("Treeview", background=[('selected', accent_color)], foreground=[('selected', 'white')])

    def ensure_directories(self):
        img_dir = os.path.join(self.work_dir, IMG_FOLDER_NAME)
        if not os.path.exists(img_dir):
            os.makedirs(img_dir)

    # ==========================
    # UI Layout
    # ==========================
    def create_top_nav(self):
        # Directory Selection Bar
        nav_frame = tk.Frame(self.root, bg="#ddd", height=40, padx=10, pady=5)
        nav_frame.pack(side=tk.TOP, fill=tk.X)
        
        tk.Label(nav_frame, text="Working Directory:", bg="#ddd", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        
        self.dir_var = tk.StringVar(value=self.work_dir)
        tk.Entry(nav_frame, textvariable=self.dir_var, state="readonly", width=60).pack(side=tk.LEFT, padx=10)
        
        tk.Button(nav_frame, text="Browse...", command=self.change_directory, bg="white", relief="flat").pack(side=tk.LEFT)

    def create_main_toolbar(self):
        self.main_toolbar = tk.Frame(self.root, bg="#f4f4f4", pady=10, padx=10)
        self.main_toolbar.pack(side=tk.TOP, fill=tk.X)

        # Plus Button
        self.add_btn = ttk.Button(self.main_toolbar, text="+", style="Green.TButton", width=3, command=self.show_add_options)
        self.add_btn.pack(side=tk.LEFT, padx=(0, 5))

        # Minus Button
        self.minus_btn = ttk.Button(self.main_toolbar, text="-", width=3, command=self.enter_selection_mode)
        self.minus_btn.pack(side=tk.LEFT, padx=(0, 15))

        # Search
        tk.Label(self.main_toolbar, text="Search:", bg="#f4f4f4").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_contacts)
        tk.Entry(self.main_toolbar, textvariable=self.search_var, font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

    def create_bulk_toolbar(self):
        self.bulk_toolbar = tk.Frame(self.root, bg="#ffebee", pady=10, padx=10)
        # Not packed initially
        
        tk.Label(self.bulk_toolbar, text="Selection Mode", bg="#ffebee", font=("Segoe UI", 12, "bold"), fg="#D32F2F").pack(side=tk.LEFT, padx=10)
        
        ttk.Button(self.bulk_toolbar, text="Delete Selected", style="Red.TButton", command=self.delete_selected_contacts).pack(side=tk.RIGHT, padx=5)
        ttk.Button(self.bulk_toolbar, text="Cancel", command=self.exit_selection_mode).pack(side=tk.RIGHT, padx=5)

    def create_list_view(self):
        # Treeview with a hidden 'select' column for checkboxes
        columns = ("Select", "Name", "Company", "Email", "Phone")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", selectmode="browse")
        
        self.tree.heading("Select", text="☐")
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
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", self.on_item_double_click)
        self.tree.bind("<ButtonRelease-1>", self.on_item_click)

    # ==========================
    # State Logic
    # ==========================
    def change_directory(self):
        new_dir = filedialog.askdirectory()
        if new_dir:
            self.work_dir = new_dir
            self.dir_var.set(self.work_dir)
            self.ensure_directories()
            self.contacts = []
            self.load_data()
            self.refresh_list()

    def enter_selection_mode(self):
        self.selection_mode = True
        self.selected_indices = set()
        self.main_toolbar.pack_forget()
        self.bulk_toolbar.pack(side=tk.TOP, fill=tk.X, before=self.tree)
        self.refresh_list() # To show checkboxes

    def exit_selection_mode(self):
        self.selection_mode = False
        self.selected_indices = set()
        self.bulk_toolbar.pack_forget()
        self.main_toolbar.pack(side=tk.TOP, fill=tk.X, before=self.tree)
        self.refresh_list()

    def on_item_click(self, event):
        if not self.selection_mode:
            return
            
        region = self.tree.identify("region", event.x, event.y)
        if region == "heading":
            # Handle header click (select all logic could go here)
            pass
        else:
            item_id = self.tree.identify_row(event.y)
            column = self.tree.identify_column(event.x)
            
            if item_id:
                # If clicking anywhere in row during selection mode, toggle
                index = self.tree.index(item_id)
                if index in self.selected_indices:
                    self.selected_indices.remove(index)
                else:
                    self.selected_indices.add(index)
                self.refresh_list_visuals_only()

    def on_item_double_click(self, event):
        if self.selection_mode:
            return # Ignore in select mode
        
        item_id = self.tree.selection()
        if not item_id: return
        
        # Get actual contact index (mapping needed if sorted/filtered)
        # Simplification: We store index in tag
        item = self.tree.item(item_id)
        index = int(item['tags'][0])
        self.open_editor(self.contacts[index], contact_index=index)

    # ==========================
    # Data Persistence
    # ==========================
    def load_data(self):
        csv_path = os.path.join(self.work_dir, DEFAULT_CSV_NAME)
        if not os.path.exists(csv_path):
            self.contacts = []
            return

        with open(csv_path, mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            self.contacts = list(reader)

        # Post-process: Parse JSON fields
        for c in self.contacts:
            # Images
            try:
                c["Image Data"] = json.loads(c.get("Image Data", "[]"))
            except:
                # Backward compatibility or fix raw paths
                raw = c.get("Image Data", "")
                if raw and not raw.startswith("["):
                    c["Image Data"] = [{"name": "Card", "path": raw}]
                else:
                    c["Image Data"] = []

            # Notes
            try:
                c["Notes Data"] = json.loads(c.get("Notes Data", "[]"))
            except:
                raw = c.get("Notes Data", "") or c.get("Notes", "") # Legacy field
                if raw and not raw.startswith("["):
                    c["Notes Data"] = [{"name": "General", "content": raw}]
                else:
                    c["Notes Data"] = []

    def save_data(self):
        csv_path = os.path.join(self.work_dir, DEFAULT_CSV_NAME)
        
        # Serialize JSON fields
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

        for i, contact in enumerate(data):
            full_name = f"{contact.get('First Name', '')} {contact.get('Last Name', '')}".strip()
            
            # Checkbox logic
            checkbox_char = ""
            if self.selection_mode:
                # Note: 'i' here might not match self.contacts index if filtered. 
                # For robust implementation, we need real ID. 
                # For this demo, assuming no filter when deleting is safest or mapping indices.
                # Let's map object identity for deletion.
                real_index = self.contacts.index(contact)
                checkbox_char = "☑" if real_index in self.selected_indices else "☐"

            self.tree.insert("", tk.END, values=(
                checkbox_char,
                full_name,
                contact.get("Company", ""),
                contact.get("E-mail Address", ""),
                contact.get("Mobile Phone", "")
            ), tags=(str(self.contacts.index(contact)),)) # Tag stores main list index

    def refresh_list_visuals_only(self):
        # Just update the checkboxes without reloading data
        for item_id in self.tree.get_children():
            tag_index = int(self.tree.item(item_id)['tags'][0])
            vals = list(self.tree.item(item_id)['values'])
            
            if self.selection_mode:
                vals[0] = "☑" if tag_index in self.selected_indices else "☐"
            else:
                vals[0] = ""
            
            self.tree.item(item_id, values=vals)

    def filter_contacts(self, *args):
        query = self.search_var.get().lower()
        if not query:
            self.refresh_list()
            return

        filtered = []
        for c in self.contacts:
            values = "".join([str(v) for v in c.values()]).lower()
            if query in values:
                filtered.append(c)
        self.refresh_list(filtered)

    def sort_column(self, col_name):
        key_map = {"Name": "First Name", "Company": "Company", "Email": "E-mail Address"}
        key = key_map.get(col_name, "First Name")
        
        self.contacts.sort(key=lambda x: x.get(key, "").lower(), reverse=self.sort_reverse)
        self.sort_reverse = not self.sort_reverse
        self.refresh_list()

    # ==========================
    # Logic: Add & Merge
    # ==========================
    def show_add_options(self):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Manual Creation", command=lambda: self.open_editor())
        menu.add_command(label="From Image", command=self.add_from_image)
        menu.add_command(label="From PDF", command=self.add_from_pdf)
        
        try:
            x = self.add_btn.winfo_rootx()
            y = self.add_btn.winfo_rooty() + self.add_btn.winfo_height()
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def add_from_image(self):
        files = filedialog.askopenfilenames(filetypes=[("Images", "*.png;*.jpg;*.jpeg")])
        if files: self.process_files(files, is_pdf=False)

    def add_from_pdf(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if files: self.process_files(files, is_pdf=True)

    def process_files(self, file_paths, is_pdf):
        # Multi-file processing assumes one contact per batch selection for simplicity,
        # or we create a new contact for the FIRST, and append others?
        # The prompt implies: "Open image or images... Create a new contact for EACH image" (original prompt)
        # But Prompt 3 says "Allow multiple business card images for a single contact".
        # Strategy: Ask user? No, let's treat the batch as ONE contact if user selected multiple at once, 
        # OR just process the first one as "Front" and others as attachments.
        
        # Better approach for workflow:
        # Create ONE contact. Analyze all images.
        
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
                # It's an image. Copy it to our folder for persistence?
                # Yes, let's copy to ensure persistence
                name = f"import_{int(time.time())}_{os.path.basename(fp)}"
                save_path = os.path.join(self.work_dir, IMG_FOLDER_NAME, name)
                shutil.copy2(fp, save_path)
                processed_images.append(save_path)

            for img_path in processed_images:
                # OCR
                try:
                    img = Image.open(img_path)
                    text = pytesseract.image_to_string(img)
                    full_text += text + "\n"
                    
                    # Back of card detection (Heuristic: Low text density)
                    # If text length is small, name it "Back", else "Card"
                    label = "Back" if len(text.strip()) < 40 else "Card"
                    img_data_list.append({"name": label, "path": img_path})
                except:
                    pass

        # Parse Info from the aggregated text of the first "Card" image usually
        data = self.heuristic_parse(full_text)
        data["Image Data"] = img_data_list
        data["Notes Data"] = [{"name": "OCR Text", "content": full_text}]
        
        self.check_and_open_editor(data)

    def heuristic_parse(self, text):
        data = {k: "" for k in CSV_HEADERS}
        data["Image Data"] = []
        data["Notes Data"] = []

        # Emails
        emails = re.findall(r'[\w\.-]+@[\w\.-]+\.\w+', text)
        if emails: data["E-mail Address"] = emails[0]

        # Phones
        phones = re.findall(r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', text)
        if phones: data["Mobile Phone"] = phones[0]

        # Name Splitting (Improvement 4)
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        if lines:
            # Assume first non-empty line is name
            full_name = lines[0]
            parts = full_name.split()
            if len(parts) > 1:
                data["Last Name"] = parts[-1]
                data["First Name"] = " ".join(parts[:-1])
            else:
                data["First Name"] = full_name
        
        return data

    def check_and_open_editor(self, new_data):
        # Duplicate Detection (Improvement 2)
        new_name = f"{new_data.get('First Name','')} {new_data.get('Last Name','')}".strip()
        
        duplicate_candidate = None
        candidate_idx = -1
        
        for i, c in enumerate(self.contacts):
            c_name = f"{c.get('First Name','')} {c.get('Last Name','')}".strip()
            if c_name.lower() == new_name.lower() and c_name != "":
                duplicate_candidate = c
                candidate_idx = i
                break
        
        if duplicate_candidate:
            resp = messagebox.askyesno("Duplicate Found", 
                                       f"Contact '{new_name}' already exists.\nDo you want to combine them?")
            if resp:
                self.resolve_merge(duplicate_candidate, new_data, candidate_idx)
                return

        # No duplicate, just open editor
        self.open_editor(existing_data=new_data)

    def resolve_merge(self, old_data, new_data, idx):
        # Logic: If old is empty and new has value -> auto fill.
        # If both have value and differ -> Ask user.
        merged = old_data.copy()
        conflicts = []

        # Merge Images (append)
        merged["Image Data"] = old_data.get("Image Data", []) + new_data.get("Image Data", [])
        # Merge Notes (append)
        merged["Notes Data"] = old_data.get("Notes Data", []) + new_data.get("Notes Data", [])

        fields = ["Company", "Job Title", "E-mail Address", "Mobile Phone", "Business Phone", "Address"]
        
        for f in fields:
            old_val = old_data.get(f, "").strip()
            new_val = new_data.get(f, "").strip()
            
            if not old_val and new_val:
                merged[f] = new_val
            elif old_val and new_val and old_val != new_val:
                conflicts.append((f, old_val, new_val))

        if conflicts:
            # Open Conflict Resolver Dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Merge Conflicts")
            tk.Label(dialog, text="Select which information to keep:", font=("bold")).pack(pady=10)
            
            vars_map = {}
            for f, old_v, new_v in conflicts:
                frame = tk.Frame(dialog, pady=5)
                frame.pack(fill=tk.X, padx=10)
                tk.Label(frame, text=f"{f}:", width=15, anchor="e").pack(side=tk.LEFT)
                
                v = tk.StringVar(value="old")
                vars_map[f] = (v, old_v, new_v)
                
                tk.Radiobutton(frame, text=f"Old: {old_v}", variable=v, value="old").pack(side=tk.LEFT)
                tk.Radiobutton(frame, text=f"New: {new_v}", variable=v, value="new").pack(side=tk.LEFT)

            def finish_merge():
                for f, (v, old_v, new_v) in vars_map.items():
                    merged[f] = old_v if v.get() == "old" else new_v
                self.contacts[idx] = merged
                self.save_data()
                self.refresh_list()
                dialog.destroy()
                self.open_editor(merged, contact_index=idx) # Open editor to verify

            tk.Button(dialog, text="Confirm Merge", command=finish_merge).pack(pady=10)
        else:
            self.contacts[idx] = merged
            self.save_data()
            self.refresh_list()
            self.open_editor(merged, contact_index=idx)

    # ==========================
    # Editor Window
    # ==========================
    def open_editor(self, existing_data=None, contact_index=None):
        Editor(self, existing_data, contact_index)

    def delete_contact(self, index, window):
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this contact?"):
            del self.contacts[index]
            self.save_data()
            self.refresh_list()
            window.destroy()

    def delete_selected_contacts(self):
        if not self.selected_indices:
            return
        
        if messagebox.askyesno("Delete Selected", f"Delete {len(self.selected_indices)} contacts?"):
            # Delete in reverse order to preserve indices during deletion
            for i in sorted(list(self.selected_indices), reverse=True):
                del self.contacts[i]
            self.save_data()
            self.exit_selection_mode()

# ==========================================
# SEPARATE EDITOR CLASS (To handle complex UI)
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
        self.win.title("Contact Details")
        self.win.geometry("900x800")
        
        self.create_ui()

    def create_ui(self):
        # Main Layout: Left (Images), Right (Fields)
        main_paned = tk.PanedWindow(self.win, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Left Side: Image Notebook ---
        self.left_frame = tk.Frame(main_paned)
        main_paned.add(self.left_frame, width=400)
        
        self.img_notebook = ttk.Notebook(self.left_frame)
        self.img_notebook.pack(fill=tk.BOTH, expand=True)
        self.img_notebook.bind("<Double-1>", self.rename_tab)
        
        # Load Images
        self.render_images()

        btn_img_frame = tk.Frame(self.left_frame)
        btn_img_frame.pack(fill=tk.X, pady=5)
        tk.Button(btn_img_frame, text="Remove Image", command=self.remove_current_image).pack(side=tk.LEFT)

        # --- Right Side: Inputs ---
        self.right_frame = tk.Frame(main_paned)
        main_paned.add(self.right_frame)

        # Scrollable Frame for inputs
        canvas = tk.Canvas(self.right_frame)
        scrollbar = ttk.Scrollbar(self.right_frame, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas)

        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Fields
        self.entries = {}
        row = 0
        
        fields = ["First Name", "Last Name", "Company", "Job Title", 
                  "E-mail Address", "Mobile Phone", "Business Phone"]
        
        for f in fields:
            tk.Label(scroll_frame, text=f, font=("Segoe UI", 9, "bold")).grid(row=row, column=0, sticky="w", pady=(10,0))
            ent = tk.Entry(scroll_frame, font=("Segoe UI", 10), width=40)
            ent.insert(0, self.data.get(f, ""))
            ent.grid(row=row+1, column=0, sticky="ew", padx=5)
            self.entries[f] = ent
            row += 2

        # Address (Multi-line)
        tk.Label(scroll_frame, text="Address", font=("Segoe UI", 9, "bold")).grid(row=row, column=0, sticky="w", pady=(10,0))
        addr_txt = tk.Text(scroll_frame, height=4, width=40, font=("Segoe UI", 10))
        addr_txt.insert("1.0", self.data.get("Address", ""))
        addr_txt.grid(row=row+1, column=0, sticky="ew", padx=5)
        self.entries["Address"] = addr_txt
        row += 2

        # --- Notes Notebook ---
        tk.Label(scroll_frame, text="Notes", font=("Segoe UI", 9, "bold")).grid(row=row, column=0, sticky="w", pady=(10,0))
        
        self.note_notebook = ttk.Notebook(scroll_frame, height=150)
        self.note_notebook.grid(row=row+1, column=0, sticky="ew", padx=5, pady=5)
        self.note_notebook.bind("<Button-1>", self.on_note_tab_click)
        self.note_notebook.bind("<Double-1>", self.rename_note_tab)

        self.note_widgets = [] # Keep track to save later
        self.render_notes()
        
        # --- Bottom Buttons ---
        action_frame = tk.Frame(self.win, pady=10, bg="#f4f4f4")
        action_frame.pack(side=tk.BOTTOM, fill=tk.X)

        tk.Button(action_frame, text="Save Contact", bg="#4CAF50", fg="white", font=("Segoe UI", 10, "bold"),
                  command=self.save).pack(side=tk.RIGHT, padx=10)
        
        tk.Button(action_frame, text="Cancel", command=self.win.destroy).pack(side=tk.RIGHT, padx=10)

        if not self.is_new:
            # Delete Button
            del_btn = ttk.Button(action_frame, text="Delete Contact", style="Red.TButton", 
                                 command=lambda: self.app.delete_contact(self.index, self.win))
            del_btn.pack(side=tk.LEFT, padx=10)

    # --- Image Logic ---
    def render_images(self):
        # Clear existing tabs
        for tab in self.img_notebook.tabs():
            self.img_notebook.forget(tab)
            
        images = self.data.get("Image Data", [])
        if not images:
            lbl = tk.Label(self.img_notebook, text="No Images")
            self.img_notebook.add(lbl, text="None")
            return

        for img_entry in images:
            path = img_entry["path"]
            name = img_entry.get("name", "Card")
            
            frame = tk.Frame(self.img_notebook)
            self.img_notebook.add(frame, text=name)
            
            if os.path.exists(path):
                try:
                    pil = Image.open(path)
                    pil.thumbnail((380, 500))
                    tk_img = ImageTk.PhotoImage(pil)
                    lbl = tk.Label(frame, image=tk_img)
                    lbl.image = tk_img
                    lbl.pack(expand=True)
                except:
                    tk.Label(frame, text="Error loading image").pack()
            else:
                tk.Label(frame, text="Image file missing").pack()

    def remove_current_image(self):
        idx = self.img_notebook.index("current")
        if idx < 0: return
        
        images = self.data.get("Image Data", [])
        if idx < len(images):
            del images[idx]
            self.data["Image Data"] = images
            self.render_images()

    def rename_tab(self, event):
        clicked_tab = self.img_notebook.tk.call(self.img_notebook._w, "identify", "tab", event.x, event.y)
        if clicked_tab == "": return
        index = self.img_notebook.index(f"@{event.x},{event.y}")
        
        new_name = simpledialog.askstring("Rename Tab", "Enter new name:")
        if new_name:
            self.img_notebook.tab(index, text=new_name)
            # Update data
            self.data["Image Data"][index]["name"] = new_name

    # --- Note Logic ---
    def render_notes(self):
        # Clear existing
        for tab in self.note_notebook.tabs():
            self.note_notebook.forget(tab)
        self.note_widgets = []

        notes = self.data.get("Notes Data", [])
        if not notes:
            notes = [{"name": "General", "content": ""}]
            self.data["Notes Data"] = notes

        for n in notes:
            frame = tk.Frame(self.note_notebook)
            self.note_notebook.add(frame, text=n.get("name", "Note"))
            txt = tk.Text(frame)
            txt.insert("1.0", n.get("content", ""))
            txt.pack(fill=tk.BOTH, expand=True)
            self.note_widgets.append(txt)

        # Add Green Plus Tab
        plus_frame = tk.Frame(self.note_notebook)
        self.note_notebook.add(plus_frame, text="+")

    def on_note_tab_click(self, event):
        clicked_tab = self.note_notebook.tk.call(self.note_notebook._w, "identify", "tab", event.x, event.y)
        if clicked_tab == "": return
        index = self.note_notebook.index(f"@{event.x},{event.y}")
        
        # If clicked the last tab (+)
        if index == len(self.note_notebook.tabs()) - 1:
            # Add new note
            self.data["Notes Data"].append({"name": "New Note", "content": ""})
            self.render_notes()
            self.note_notebook.select(len(self.note_widgets)-1)

    def rename_note_tab(self, event):
        index = self.note_notebook.index(f"@{event.x},{event.y}")
        # Don't rename the plus button
        if index == len(self.note_notebook.tabs()) - 1: return
        
        new_name = simpledialog.askstring("Rename Note", "Enter Title:")
        if new_name:
            self.note_notebook.tab(index, text=new_name)
            self.data["Notes Data"][index]["name"] = new_name

    # --- Saving ---
    def save(self):
        # Gather Fields
        for f, w in self.entries.items():
            if isinstance(w, tk.Entry):
                self.data[f] = w.get().strip()
            else:
                self.data[f] = w.get("1.0", tk.END).strip()

        # Gather Notes
        new_notes = []
        for i, widget in enumerate(self.note_widgets):
            # Get name from current data or tab text
            title = self.data["Notes Data"][i]["name"]
            content = widget.get("1.0", tk.END).strip()
            new_notes.append({"name": title, "content": content})
        self.data["Notes Data"] = new_notes

        # Send back to app
        if self.is_new:
            self.app.contacts.append(self.data)
        else:
            self.app.contacts[self.index] = self.data
        
        self.app.save_data()
        self.app.refresh_list()
        self.win.destroy()
        messagebox.showinfo("Saved", "Contact Saved.")

if __name__ == "__main__":
    root = tk.Tk()
    app = RolodexApp(root)
    root.mainloop()