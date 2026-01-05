import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import csv
import os
import pytesseract
import re
from pdf2image import convert_from_path
import time

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

CSV_FILE = "contacts.csv"
IMAGE_STORE_DIR = "card_images" # Folder to store converted PDF images

# Outlook Compatible Headers
CSV_HEADERS = [
    "First Name", "Last Name", "Company", "Job Title", 
    "E-mail Address", "Mobile Phone", "Business Phone", 
    "Address", "Notes", "Image Path"
]

class RolodexApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Python Rolodex")
        self.root.geometry("900x650")

        self.contacts = []
        self.sort_reverse = False
        
        # Ensure image storage directory exists
        if not os.path.exists(IMAGE_STORE_DIR):
            os.makedirs(IMAGE_STORE_DIR)
        
        self.load_data()
        self.create_gui()

    def create_gui(self):
        # --- Top Toolbar ---
        toolbar = tk.Frame(self.root, bd=1, relief=tk.RAISED)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        # Add Button (+)
        self.add_btn = tk.Button(toolbar, text="+", font=("Arial", 16, "bold"), 
                                 width=3, command=self.show_add_options)
        self.add_btn.pack(side=tk.LEFT, padx=5, pady=5)

        # Search / Filter
        tk.Label(toolbar, text="Search:").pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_contacts)
        self.search_entry = tk.Entry(toolbar, textvariable=self.search_var)
        self.search_entry.pack(side=tk.LEFT, padx=5)

        # --- Main List Area (Treeview) ---
        columns = ("Name", "Company", "Email", "Phone")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings")
        
        # Define headings and sorting
        self.tree.heading("Name", text="Name", command=lambda: self.sort_column("Name"))
        self.tree.heading("Company", text="Company", command=lambda: self.sort_column("Company"))
        self.tree.heading("Email", text="Email", command=lambda: self.sort_column("Email"))
        self.tree.heading("Phone", text="Phone")
        
        self.tree.column("Name", width=150)
        self.tree.column("Company", width=150)
        self.tree.column("Email", width=200)
        self.tree.column("Phone", width=120)

        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.tree.bind("<Double-1>", self.open_contact_detail)

        self.refresh_list()

    # ==========================
    # Logic: Data Handling
    # ==========================
    def load_data(self):
        if not os.path.exists(CSV_FILE):
            return

        with open(CSV_FILE, mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                self.contacts.append(row)

    def save_data(self):
        with open(CSV_FILE, mode='w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)
            writer.writeheader()
            writer.writerows(self.contacts)

    def refresh_list(self, data_subset=None):
        for item in self.tree.get_children():
            self.tree.delete(item)

        data = data_subset if data_subset is not None else self.contacts

        for contact in data:
            full_name = f"{contact.get('First Name', '')} {contact.get('Last Name', '')}".strip()
            self.tree.insert("", tk.END, values=(
                full_name,
                contact.get("Company", ""),
                contact.get("E-mail Address", ""),
                contact.get("Mobile Phone", "") or contact.get("Business Phone", "")
            ), tags=(contact,))

    def filter_contacts(self, *args):
        query = self.search_var.get().lower()
        if not query:
            self.refresh_list()
            return

        filtered = []
        for c in self.contacts:
            values = "".join(c.values()).lower()
            if query in values:
                filtered.append(c)
        self.refresh_list(filtered)

    def sort_column(self, col_name):
        key_map = {
            "Name": "First Name",
            "Company": "Company",
            "Email": "E-mail Address"
        }
        key = key_map.get(col_name, "First Name")
        self.contacts.sort(key=lambda x: x.get(key, "").lower(), reverse=self.sort_reverse)
        self.sort_reverse = not self.sort_reverse
        self.refresh_list()

    # ==========================
    # Logic: Add / OCR / PDF
    # ==========================
    def show_add_options(self):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Manual Creation", command=lambda: self.open_editor())
        menu.add_command(label="From Image (Business Card)", command=self.add_from_image)
        menu.add_command(label="From PDF (Business Card)", command=self.add_from_pdf)
        
        try:
            x = self.add_btn.winfo_rootx()
            y = self.add_btn.winfo_rooty() + self.add_btn.winfo_height()
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def add_from_image(self):
        file_paths = filedialog.askopenfilenames(
            title="Select Business Card Image(s)",
            filetypes=[("Images", "*.png;*.jpg;*.jpeg;*.bmp")]
        )
        if not file_paths: return

        for fp in file_paths:
            self.process_image_to_contact(fp)

    def add_from_pdf(self):
        file_paths = filedialog.askopenfilenames(
            title="Select Business Card PDF(s)",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if not file_paths: return

        for pdf_path in file_paths:
            try:
                # Convert PDF to images
                # poppler_path is required on Windows if not in system PATH
                images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)
                
                for i, img in enumerate(images):
                    # Save the converted image so we can reference it in CSV later
                    base_name = os.path.basename(pdf_path).split('.')[0]
                    # Create unique name: filename_timestamp_page.jpg
                    save_name = f"{base_name}_{int(time.time())}_p{i}.jpg"
                    save_path = os.path.join(IMAGE_STORE_DIR, save_name)
                    
                    img.save(save_path, 'JPEG')
                    
                    # Process this new image as a contact
                    self.process_image_to_contact(save_path)
                    
            except Exception as e:
                messagebox.showerror("PDF Error", f"Could not convert PDF:\n{e}\n\nMake sure Poppler is installed and configured.")

    def process_image_to_contact(self, image_path):
        # 1. OCR Processing
        try:
            img = Image.open(image_path)
            text = pytesseract.image_to_string(img)
        except Exception as e:
            messagebox.showerror("OCR Error", f"Could not process image:\n{e}\n\nMake sure Tesseract is installed.")
            return

        # 2. Heuristic Extraction
        data = {key: "" for key in CSV_HEADERS}
        data["Image Path"] = os.path.abspath(image_path) # Use absolute path for persistence
        data["Notes"] = "--- OCR RAW TEXT ---\n" + text

        # Regex for Email
        email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', text)
        if email_match: data["E-mail Address"] = email_match.group(0)

        # Regex for Phone
        phone_match = re.search(r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', text)
        if phone_match: data["Mobile Phone"] = phone_match.group(0)

        # Guess Name (First non-empty line)
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        if lines: data["First Name"] = lines[0] 

        self.open_editor(existing_data=data, auto_save=False)

    # ==========================
    # Logic: Details & Editing
    # ==========================
    def open_contact_detail(self, event):
        selected_id = self.tree.selection()
        if not selected_id: return
        
        item = self.tree.item(selected_id)
        vals = item['values'] 
        
        found_contact = None
        for c in self.contacts:
            c_name = f"{c.get('First Name', '')} {c.get('Last Name', '')}".strip()
            # Matching logic: Name AND Email (or phone if email missing)
            if c_name == vals[0] and c.get('E-mail Address') == vals[2]:
                found_contact = c
                break
        
        if found_contact:
            self.open_editor(existing_data=found_contact, readonly=True)

    def open_editor(self, existing_data=None, readonly=False, auto_save=False):
        window = tk.Toplevel(self.root)
        window.title("Contact Details" if readonly else "Edit Contact")
        window.geometry("600x750")

        frame = tk.Frame(window, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        entries = {}
        row_idx = 0

        if existing_data:
            current_data = existing_data.copy()
        else:
            current_data = {key: "" for key in CSV_HEADERS}

        # --- Image Display ---
        img_path = current_data.get("Image Path", "")
        img_label = tk.Label(frame, text="[No Image]", bg="#ccc", height=5)
        
        if img_path and os.path.exists(img_path):
            try:
                pil_img = Image.open(img_path)
                pil_img.thumbnail((250, 200)) # Resize for UI
                tk_img = ImageTk.PhotoImage(pil_img)
                img_label.config(image=tk_img, text="", height=0)
                img_label.image = tk_img 
            except:
                img_label.config(text="[Error loading image]")
        
        img_label.grid(row=row_idx, column=0, columnspan=2, pady=10)
        row_idx += 1

        # --- Fields ---
        fields_to_show = ["First Name", "Last Name", "Company", "Job Title", 
                          "E-mail Address", "Mobile Phone", "Business Phone", "Address"]
        
        for field in fields_to_show:
            tk.Label(frame, text=field + ":", font=("Arial", 10, "bold")).grid(row=row_idx, column=0, sticky="e", pady=2)
            ent = tk.Entry(frame, width=40)
            ent.insert(0, current_data.get(field, ""))
            if readonly: ent.config(state="readonly")
            ent.grid(row=row_idx, column=1, sticky="w", pady=2)
            entries[field] = ent
            row_idx += 1

        # --- Notes ---
        tk.Label(frame, text="Notes:", font=("Arial", 10, "bold")).grid(row=row_idx, column=0, sticky="ne", pady=2)
        notes_txt = tk.Text(frame, height=8, width=40)
        notes_txt.insert("1.0", current_data.get("Notes", ""))
        if readonly: notes_txt.config(state="disabled")
        notes_txt.grid(row=row_idx, column=1, sticky="w", pady=2)
        entries["Notes"] = notes_txt
        row_idx += 1

        # --- Buttons ---
        btn_frame = tk.Frame(window, pady=20)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X)

        if readonly:
            edit_btn = tk.Button(btn_frame, text="Edit", command=lambda: self.enable_editing(window, entries, btn_frame, current_data))
            edit_btn.pack(side=tk.RIGHT, padx=20)
            close_btn = tk.Button(btn_frame, text="Close", command=window.destroy)
            close_btn.pack(side=tk.RIGHT, padx=10)
        else:
            save_btn = tk.Button(btn_frame, text="Save Contact", bg="#e1f5fe",
                                 command=lambda: self.save_contact(window, entries, current_data, existing_data))
            save_btn.pack(side=tk.RIGHT, padx=20)
            cancel_btn = tk.Button(btn_frame, text="Cancel", command=window.destroy)
            cancel_btn.pack(side=tk.RIGHT, padx=10)

    def enable_editing(self, window, entries, btn_frame, current_data):
        window.title("Edit Contact")
        for key, widget in entries.items():
            if isinstance(widget, tk.Entry): widget.config(state="normal")
            elif isinstance(widget, tk.Text): widget.config(state="normal")
        
        for widget in btn_frame.winfo_children(): widget.destroy()

        save_btn = tk.Button(btn_frame, text="Save Changes", bg="#e1f5fe",
                                command=lambda: self.save_contact(window, entries, current_data, current_data))
        save_btn.pack(side=tk.RIGHT, padx=20)
        cancel_btn = tk.Button(btn_frame, text="Cancel", command=window.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=10)

    def save_contact(self, window, entries, data_obj, original_ref=None):
        for field, widget in entries.items():
            if field == "Notes":
                data_obj[field] = widget.get("1.0", tk.END).strip()
            else:
                data_obj[field] = widget.get().strip()
        
        if original_ref in self.contacts:
            original_ref.update(data_obj)
        else:
            self.contacts.append(data_obj)

        self.save_data()
        self.refresh_list()
        window.destroy()
        messagebox.showinfo("Success", "Contact saved successfully.")

import os

if __name__ == "__main__":
    root = tk.Tk()
    app = RolodexApp(root)
    root.mainloop()
    