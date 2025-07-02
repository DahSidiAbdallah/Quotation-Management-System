import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import os
import json
import sys
from datetime import datetime
from PIL import Image, ImageTk  # Added for logo support
from tkinter import filedialog
import pandas as pd  # For Excel export
from pdf_viewer import PDFPreviewWindow
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# Path to the SQLite database and logo image
DB_PATH = 'clients.db'
LOGO_PATH = 'MAFCI.png'  # Place your company logo here


def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    # Ensure client_type exists for backward compat
    c.execute('''CREATE TABLE IF NOT EXISTS clients (
        id INTEGER PRIMARY KEY,
        name TEXT UNIQUE,
        nif TEXT,
        rc TEXT,
        address TEXT,
        client_type TEXT
    )''')
    try:
        c.execute('ALTER TABLE clients ADD COLUMN client_type TEXT')
    except sqlite3.OperationalError:
        pass
    # Create quotations table
    c.execute('''
        CREATE TABLE IF NOT EXISTS quotations (
            id INTEGER PRIMARY KEY,
            client_id INTEGER,
            type TEXT,
            number TEXT,
            product TEXT,
            quantity REAL,
            unit_price REAL,
            date TEXT,
            purchase_order TEXT,
            FOREIGN KEY(client_id) REFERENCES clients(id)
        )
    ''')
    try:
        c.execute('ALTER TABLE quotations ADD COLUMN purchase_order TEXT')
    except sqlite3.OperationalError:
        pass
    conn.commit()
    conn.close()


def create_pdf(pdf_filename, client_name, nif, rc, address, client_preferences,
               doc_type, doc_number, purchase_order, product, quantity,
               unit_price, date_str):
    """Generate a PDF with consistent layout."""
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader
    from reportlab.lib import colors

    c = canvas.Canvas(pdf_filename, pagesize=A4)
    width, height = A4
    margin = 36
    section_gap = 18

    logo_x = margin
    logo_y = height - margin - 60
    logo_width = 120
    logo_height = 60
    company_info = [
        "Société Mauritano-Française des ciments",
        "Tel:+222 45 29 85 56 / mob:+222 45 29 48 17",
        "Email : info@mafci.mr",
        "Route de Rosso, Zone Port, Nouakchott-Mauritanie",
        "Capital: 431.000.000 MRU",
        "RC: 200721 / NIF: 30400224",
    ]
    if os.path.exists(LOGO_PATH):
        c.drawImage(ImageReader(LOGO_PATH), logo_x, logo_y, width=logo_width,
                    height=logo_height, mask='auto')
    c.setFont("Helvetica", 10)
    info_y = logo_y - 18
    for line in company_info:
        c.drawString(logo_x, info_y, line)
        info_y -= 13

    bar_height = 36
    bar_y = height - margin - 10
    c.setFillColorRGB(0.19, 0.44, 0.72)
    c.roundRect(width - 260 - margin, bar_y - bar_height, 250, bar_height, 8,
                fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(width - 250 - margin, bar_y - bar_height + 22, doc_type.upper())
    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(width - margin - 18, bar_y - bar_height + 22,
                      f"N° {doc_number}")
    c.setFillColorRGB(0, 0, 0)

    box_top = logo_y - logo_height - 35
    box_height = 72
    box_width = width - 2 * margin
    box_left = margin
    c.roundRect(box_left, box_top - box_height, box_width, box_height, 7,
                stroke=1, fill=0)
    c.setFont("Helvetica", 10)
    c.drawString(box_left + 16, box_top - 18, f"Date : {date_str}")
    c.drawString(box_left + 16, box_top - 34, f"Client : {client_name}")
    c.drawString(box_left + 16, box_top - 50, f"Adresse : {address}")
    c.drawString(box_left + box_width / 2 + 16, box_top - 18, f"RC : {rc}")
    c.drawString(box_left + box_width / 2 + 16, box_top - 34, f"NIF : {nif}")
    if purchase_order:
        c.drawString(box_left + box_width / 2 + 16, box_top - 50,
                     f"Bon de commande : {purchase_order}")

    table_y = box_top - box_height - section_gap
    col_widths = [170, 90, 110, 110]
    x_positions = [
        box_left,
        box_left + col_widths[0],
        box_left + col_widths[0] + col_widths[1],
        box_left + col_widths[0] + col_widths[1] + col_widths[2],
    ]
    headers = ["DÉSIGNATION", "Quantité (T)", "P.U. HT (MRU)", "MONTANT (MRU)"]
    header_bg = colors.HexColor("#eaf1fb")
    row_height = 22
    c.setFillColor(header_bg)
    c.rect(box_left, table_y - row_height, sum(col_widths), row_height,
           fill=1, stroke=0)
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 10)
    for i, header in enumerate(headers):
        c.drawString(x_positions[i] + 8, table_y - row_height + 7, header)
    c.setLineWidth(0.5)
    c.rect(box_left, table_y - row_height, sum(col_widths), row_height,
           stroke=1, fill=0)

    c.setFont("Helvetica", 10)
    c.drawString(x_positions[0] + 8, table_y - row_height * 2 + 6, product)
    c.drawRightString(x_positions[1] + col_widths[1] - 10,
                      table_y - row_height * 2 + 6, f"{quantity:,.2f}")
    c.drawRightString(x_positions[2] + col_widths[2] - 10,
                      table_y - row_height * 2 + 6, f"{unit_price:,.2f}")
    montant = quantity * unit_price
    c.drawRightString(x_positions[3] + col_widths[3] - 10,
                      table_y - row_height * 2 + 6, f"{montant:,.2f}")
    for i in range(5):
        xpos = box_left + sum(col_widths[:i])
        c.line(xpos, table_y - row_height, xpos, table_y - row_height * 3)
    c.line(box_left, table_y - row_height * 2,
           box_left + sum(col_widths), table_y - row_height * 2)
    c.rect(box_left, table_y - row_height * 2, sum(col_widths), row_height,
           stroke=1, fill=0)

    # Move the totals box slightly lower so it doesn't crowd the client details
    summary_y = table_y - row_height * 3 - section_gap - 40
    summary_box_width = 260
    summary_box_height = 54
    summary_x = width - margin - summary_box_width
    c.roundRect(summary_x, summary_y, summary_box_width, summary_box_height, 7,
                stroke=1, fill=0)
    c.setFont("Helvetica", 10)
    c.drawString(summary_x + 14, summary_y + summary_box_height - 16,
                 "Montant HT :")
    c.drawString(summary_x + 14, summary_y + summary_box_height - 32,
                 "TVA (16%) :")
    c.drawString(summary_x + 14, summary_y + summary_box_height - 48, "TTC :")
    c.setFont("Helvetica-Bold", 10)
    c.drawRightString(summary_x + summary_box_width - 16,
                      summary_y + summary_box_height - 16, f"{montant:,.2f}")
    tva = montant * 0.16
    c.drawRightString(summary_x + summary_box_width - 16,
                      summary_y + summary_box_height - 32, f"{tva:,.2f}")
    c.drawRightString(summary_x + summary_box_width - 16,
                      summary_y + summary_box_height - 48, f"{montant + tva:,.2f}")

    pay_y = margin + 70
    try:
        import qrcode
        from io import BytesIO
        qr = qrcode.QRCode(box_size=2, border=1)
        qr.add_data('MR130030000101006313901-73')
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        buf = BytesIO()
        img.save(buf, format='PNG')
        buf.seek(0)
        c.drawImage(ImageReader(buf), width - margin - 60, pay_y - 5,
                    width=52, height=52)
    except Exception:
        pass
    c.setFont("Helvetica-Bold", 10)
    c.drawString(margin, pay_y + 40,
                 "Paiement par virement bancaire uniquement :")
    c.setFont("Helvetica", 9)
    c.drawString(margin, pay_y + 28, "Banque : BAMIS")
    c.drawString(margin, pay_y + 16, "Compte :  00001 01006313901-73")
    c.drawString(margin, pay_y + 4, "IBAN : MR130030000101006313901-73")
    c.drawString(margin, pay_y - 8, "Devise : MRU")

    if client_preferences.get('afficher_pied', True):
        pied = client_preferences.get('pied_page', '')
        if pied:
            c.setStrokeColorRGB(0.7, 0.7, 0.7)
            c.setLineWidth(0.5)
            c.line(margin, 35, width - margin, 35)
            c.setFont("Helvetica-Oblique", 9)
            c.drawCentredString(width / 2, 25, pied)

    c.save()


class AddClientWindow(tb.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Ajouter un nouveau client")
        self.geometry("400x350")
        self.create_widgets()

    def setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use('clam')
        except:
            pass
        style.configure('TButton', font=('Segoe UI', 10), padding=6)
        style.configure('Accent.TButton', font=('Segoe UI', 10, 'bold'), padding=8, foreground='white', background='#0078D4')
        style.map('Accent.TButton', background=[('active', '#005A9E')])
        style.configure('TLabel', font=('Segoe UI', 10))
        style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'))
        style.configure('TEntry', font=('Segoe UI', 10), padding=4)
        style.configure('TLabelframe.Label', font=('Segoe UI', 10, 'bold'))
        self.option_add("*Font", ("Segoe UI", 10))

    def create_widgets(self):
        frm = tb.Frame(self)
        frm.pack(fill='both', expand=True, padx=20, pady=15)
        tb.Label(frm, text="Nom :").grid(row=0, column=0, sticky='e', pady=5)
        self.name_entry = tb.Entry(frm)
        self.name_entry.grid(row=0, column=1, pady=5, sticky='ew')
        tb.Label(frm, text="NIF :").grid(row=1, column=0, sticky='e', pady=5)
        self.nif_entry = tb.Entry(frm)
        self.nif_entry.grid(row=1, column=1, pady=5, sticky='ew')
        tb.Label(frm, text="RC :").grid(row=2, column=0, sticky='e', pady=5)
        self.rc_entry = tb.Entry(frm)
        self.rc_entry.grid(row=2, column=1, pady=5, sticky='ew')
        tb.Label(frm, text="Adresse :").grid(row=3, column=0, sticky='e', pady=5)
        self.addr_entry = tb.Entry(frm)
        self.addr_entry.grid(row=3, column=1, pady=5, sticky='ew')
        tb.Label(frm, text="Type de client :").grid(row=4, column=0, sticky='e', pady=5)
        self.client_type_var = tk.StringVar()
        self.ciment_radio = tb.Radiobutton(frm, text="Ciment", variable=self.client_type_var, value='ciment')
        self.ciment_radio.grid(row=4, column=1, sticky='w', pady=5)
        self.beton_radio = tb.Radiobutton(frm, text="Béton", variable=self.client_type_var, value='beton')
        self.beton_radio.grid(row=4, column=1, sticky='e', pady=5)
        tb.Button(frm, text="Enregistrer", command=self.save_client, bootstyle="success").grid(row=5, column=0, columnspan=2, pady=15, sticky='ew')

        self.geometry("350x280")
        self.parent = parent
        # Form fields
        tk.Label(self, text="Nom :").grid(row=0, column=0, sticky='e')
        self.name_entry = tk.Entry(self)
        self.name_entry.grid(row=0, column=1)
        tk.Label(self, text="NIF :").grid(row=1, column=0, sticky='e')
        self.nif_entry = tk.Entry(self)
        self.nif_entry.grid(row=1, column=1)
        tk.Label(self, text="RC :").grid(row=2, column=0, sticky='e')
        self.rc_entry = tk.Entry(self)
        self.rc_entry.grid(row=2, column=1)
        tk.Label(self, text="Adresse :").grid(row=3, column=0, sticky='e')
        self.addr_entry = tk.Entry(self)
        self.addr_entry.grid(row=3, column=1)
        tk.Label(self, text="Type de client :").grid(row=4, column=0, sticky='e')
        self.client_type_var = tk.StringVar(value='ciment')
        self.ciment_radio = tk.Radiobutton(self, text="Ciment", variable=self.client_type_var, value='ciment')
        self.ciment_radio.grid(row=4, column=1, sticky='w')
        self.beton_radio = tk.Radiobutton(self, text="Béton", variable=self.client_type_var, value='beton')
        self.beton_radio.grid(row=4, column=1, sticky='e')
        # Preferences button
        self.preferences = {}
        tk.Button(self, text="Préférences...", command=self.open_preferences).grid(row=5, column=1, pady=5)
        # Save button
        tk.Button(self, text="Enregistrer", command=self.save_client).grid(row=6, column=1, pady=10)

    def open_preferences(self):
        PreferencesWindow(self)

    def save_client(self):
        import json
        name = self.name_entry.get().strip()
        nif = self.nif_entry.get().strip()
        rc = self.rc_entry.get().strip()
        addr = self.addr_entry.get().strip()
        client_type = self.client_type_var.get()
        preferences = json.dumps(self.preferences)  # Save as JSON
        if not name:
            messagebox.showerror("Erreur", "Le nom du client est obligatoire")
            return
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        try:
            c.execute('INSERT INTO clients (name, nif, rc, address, client_type, preferences) VALUES (?,?,?,?,?,?)',
                      (name, nif, rc, addr, client_type, preferences))
            conn.commit()
        except sqlite3.IntegrityError:
            messagebox.showerror("Erreur", "Le client existe déjà")
        conn.close()
        self.parent.refresh_clients()
        self.destroy()

class EditClientWindow(tb.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Modifier un client")
        self.geometry("400x350")
        self.parent = parent
        # Select client
        tb.Label(self, text="Choisir le client à modifier :").grid(row=0, column=0, sticky='e')
        self.client_var = tk.StringVar()
        self.client_dropdown = tb.Combobox(self, textvariable=self.client_var, state='readonly')
        self.client_dropdown['values'] = self.parent.load_clients()
        self.client_dropdown.grid(row=0, column=1)
        tb.Button(self, text="Charger", command=self.load_client, bootstyle="primary").grid(row=0, column=2)
        # Info fields
        tb.Label(self, text="Nom :").grid(row=1, column=0, sticky='e')
        self.name_entry = tb.Entry(self)
        self.name_entry.grid(row=1, column=1)
        tb.Label(self, text="NIF :").grid(row=2, column=0, sticky='e')
        self.nif_entry = tb.Entry(self)
        self.nif_entry.grid(row=2, column=1)
        tb.Label(self, text="RC :").grid(row=3, column=0, sticky='e')
        self.rc_entry = tb.Entry(self)
        self.rc_entry.grid(row=3, column=1)
        tb.Label(self, text="Adresse :").grid(row=4, column=0, sticky='e')
        self.addr_entry = tb.Entry(self)
        self.addr_entry.grid(row=4, column=1)
        tb.Label(self, text="Type de client :").grid(row=5, column=0, sticky='e')
        self.client_type_var = tk.StringVar()
        self.ciment_radio = tb.Radiobutton(self, text="Ciment", variable=self.client_type_var, value='ciment')
        self.ciment_radio.grid(row=5, column=1, sticky='w')
        self.beton_radio = tb.Radiobutton(self, text="Béton", variable=self.client_type_var, value='beton')
        self.beton_radio.grid(row=5, column=1, sticky='e')
        # Preferences
        self.preferences = {}
        tb.Button(self, text="Préférences...", command=self.open_preferences, bootstyle="secondary").grid(row=6, column=1, pady=5)
        # Save
        tb.Button(self, text="Enregistrer les modifications", command=self.save_client, bootstyle="success").grid(row=7, column=1, pady=10)

    def load_client(self):
        import json
        name = self.client_var.get().strip()
        if not name:
            messagebox.showerror("Erreur", "Veuillez sélectionner un client")
            return
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute('SELECT nif, rc, address, client_type, preferences FROM clients WHERE name=?', (name,))
        row = c.fetchone()
        conn.close()
        if not row:
            messagebox.showerror("Erreur", "Client non trouvé")
            return
        nif, rc, addr, client_type, preferences = row
        self.name_entry.delete(0, 'end'); self.name_entry.insert(0, name)
        self.nif_entry.delete(0, 'end'); self.nif_entry.insert(0, nif)
        self.rc_entry.delete(0, 'end'); self.rc_entry.insert(0, rc)
        self.addr_entry.delete(0, 'end'); self.addr_entry.insert(0, addr)
        self.client_type_var.set(client_type)
        try:
            self.preferences = json.loads(preferences) if preferences else {}
        except Exception:
            self.preferences = {}

    def open_preferences(self):
        PreferencesWindow(self)

    def save_client(self):
        import json
        name = self.name_entry.get().strip()
        nif = self.nif_entry.get().strip()
        rc = self.rc_entry.get().strip()
        addr = self.addr_entry.get().strip()
        client_type = self.client_type_var.get()
        preferences = json.dumps(self.preferences)  # Save as JSON
        if not name:
            messagebox.showerror("Erreur", "Le nom du client est obligatoire")
            return
        old_name = self.client_var.get().strip()
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute('UPDATE clients SET name=?, nif=?, rc=?, address=?, client_type=?, preferences=? WHERE name=?',
                  (name, nif, rc, addr, client_type, preferences, old_name))
        conn.commit()
        conn.close()
        self.parent.refresh_clients()
        self.destroy()

class PreferencesWindow(tb.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Préférences du client")
        self.geometry("400x400")
        self.parent = parent
        prefs = self.parent.preferences if hasattr(self.parent, 'preferences') else {}

        # Champs supplémentaires
        tb.Label(self, text="Adresse de livraison :").pack(anchor='w', padx=10, pady=(10,0))
        self.delivery_entry = tb.Entry(self)
        self.delivery_entry.pack(fill='x', padx=10)
        self.delivery_entry.insert(0, prefs.get('adresse_livraison', ''))

        tb.Label(self, text="Informations fiscales :").pack(anchor='w', padx=10, pady=(10,0))
        self.tax_entry = tb.Entry(self)
        self.tax_entry.pack(fill='x', padx=10)
        self.tax_entry.insert(0, prefs.get('infos_fiscales', ''))

        tb.Label(self, text="Notes :").pack(anchor='w', padx=10, pady=(10,0))
        self.notes_entry = tb.Entry(self)
        self.notes_entry.pack(fill='x', padx=10)
        self.notes_entry.insert(0, prefs.get('notes', ''))

        # Sections à afficher/masquer
        tb.Label(self, text="Sections à afficher :").pack(anchor='w', padx=10, pady=(10,0))
        self.show_footer_var = tk.BooleanVar(value=prefs.get('afficher_pied', True))
        tb.Checkbutton(self, text="Pied de page personnalisé", variable=self.show_footer_var).pack(anchor='w', padx=20)

        # Pied de page personnalisé
        tb.Label(self, text="Pied de page :").pack(anchor='w', padx=10, pady=(10,0))
        self.footer_entry = tb.Entry(self)
        self.footer_entry.pack(fill='x', padx=10)
        self.footer_entry.insert(0, prefs.get('pied_page', ''))

        # Logo personnalisé (just text path for now)
        tb.Label(self, text="Logo personnalisé (chemin) :").pack(anchor='w', padx=10, pady=(10,0))
        self.logo_entry = tb.Entry(self)
        self.logo_entry.pack(fill='x', padx=10)
        self.logo_entry.insert(0, prefs.get('logo', ''))

        # Branding
        tb.Label(self, text="Couleur principale (ex: #FF0000) :").pack(anchor='w', padx=10, pady=(10,0))
        self.color_entry = tb.Entry(self)
        self.color_entry.pack(fill='x', padx=10)
        self.color_entry.insert(0, prefs.get('couleur', ''))

        tb.Button(self, text="Enregistrer", command=self.save, bootstyle="success").pack(pady=15)

    def save(self):
        # Save all preferences to parent
        self.parent.preferences = {
            'adresse_livraison': self.delivery_entry.get(),
            'infos_fiscales': self.tax_entry.get(),
            'notes': self.notes_entry.get(),
            'afficher_pied': self.show_footer_var.get(),
            'pied_page': self.footer_entry.get(),
            'logo': self.logo_entry.get(),
            'couleur': self.color_entry.get()
        }
        self.destroy()

        super().__init__(parent)
        self.title("Ajouter un nouveau client")
        self.geometry("350x280")
        self.parent = parent
        # Form fields
        tk.Label(self, text="Nom :").grid(row=0, column=0, sticky='e')
        self.name_entry = tk.Entry(self)
        self.name_entry.grid(row=0, column=1)
        tk.Label(self, text="NIF :").grid(row=1, column=0, sticky='e')
        self.nif_entry = tk.Entry(self)
        self.nif_entry.grid(row=1, column=1)
        tk.Label(self, text="RC :").grid(row=2, column=0, sticky='e')
        self.rc_entry = tk.Entry(self)
        self.rc_entry.grid(row=2, column=1)
        tk.Label(self, text="Adresse :").grid(row=3, column=0, sticky='e')
        self.addr_entry = tk.Entry(self)
        self.addr_entry.grid(row=3, column=1)
        tk.Label(self, text="Type de client :").grid(row=4, column=0, sticky='e')
        self.client_type_var = tk.StringVar(value='ciment')
        self.ciment_radio = tk.Radiobutton(self, text="Ciment", variable=self.client_type_var, value='ciment')
        self.ciment_radio.grid(row=4, column=1, sticky='w')
        self.beton_radio = tk.Radiobutton(self, text="Béton", variable=self.client_type_var, value='beton')
        self.beton_radio.grid(row=4, column=1, sticky='e')

        # Preferences button
        self.preferences = {}
        tk.Button(self, text="Préférences...", command=self.open_preferences).grid(row=5, column=1, pady=5)

        # Save button
        tk.Button(self, text="Enregistrer", command=self.save_client).grid(row=6, column=1, pady=10)

    def save_client(self):
        import json
        name = self.name_entry.get().strip()
        nif = self.nif_entry.get().strip()
        rc = self.rc_entry.get().strip()
        addr = self.addr_entry.get().strip()
        client_type = self.client_type_var.get()
        preferences = json.dumps(self.preferences)  # Save as JSON
        if not name:
            messagebox.showerror("Erreur", "Le nom du client est obligatoire")
            return
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        try:
            c.execute('INSERT INTO clients (name, nif, rc, address, client_type, preferences) VALUES (?,?,?,?,?,?)',
                      (name, nif, rc, addr, client_type, preferences))
            conn.commit()
        except sqlite3.IntegrityError:
            messagebox.showerror("Erreur", "Le client existe déjà")
        conn.close()
        self.parent.refresh_clients()
        self.destroy()

    def open_preferences(self):
        PreferencesWindow(self)


class HistoryWindow(tb.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Historique des devis et factures")
        self.geometry("900x400")
        self.parent = parent
        self.conn = sqlite3.connect(DB_PATH)

        # Filters
        filter_frame = tb.Frame(self)
        filter_frame.pack(fill='x', padx=10, pady=5)

        tb.Label(filter_frame, text="Client :").pack(side='left')
        self.client_var = tk.StringVar()
        self.client_dropdown = tb.Combobox(filter_frame, textvariable=self.client_var, state='readonly')
        self.client_dropdown.pack(side='left', padx=5)
        self.client_dropdown['values'] = self.get_clients()
        self.client_dropdown.bind('<<ComboboxSelected>>', lambda e: self.refresh_tree())

        tb.Label(filter_frame, text="Type :").pack(side='left', padx=(10,0))
        self.type_var = tk.StringVar()
        self.type_dropdown = tb.Combobox(filter_frame, textvariable=self.type_var, state='readonly')
        self.type_dropdown['values'] = ("", "devis", "facture")
        self.type_dropdown.pack(side='left', padx=5)
        self.type_dropdown.bind('<<ComboboxSelected>>', lambda e: self.refresh_tree())

        tb.Label(filter_frame, text="Date (AAAA-MM-JJ) :").pack(side='left', padx=(10,0))
        self.date_entry = tb.Entry(filter_frame)
        self.date_entry.pack(side='left', padx=5)
        self.date_entry.bind('<Return>', lambda e: self.refresh_tree())

        tb.Button(filter_frame, text="Réinitialiser", command=self.reset_filters, bootstyle="secondary").pack(side='left', padx=10)

        # Treeview
        columns = ("Client", "Type", "Numéro", "Produit", "Quantité", "Prix Unitaire", "Date", "Bon de commande")
        self.tree = tb.Treeview(self, columns=columns, show='headings')
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        self.tree.pack(fill='both', expand=True, padx=10, pady=5)

        # Export button
        tb.Button(self, text="Exporter vers Excel", command=self.export_to_excel, bootstyle="success").pack(pady=5)

        self.refresh_tree()

    def get_clients(self):
        c = self.conn.cursor()
        c.execute('SELECT name FROM clients')
        return [r[0] for r in c.fetchall()]

    def refresh_tree(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        c = self.conn.cursor()
        query = '''SELECT clients.name, quotations.type, quotations.number, quotations.product, quotations.quantity, quotations.unit_price, quotations.date, quotations.purchase_order
                   FROM quotations JOIN clients ON quotations.client_id = clients.id WHERE 1=1'''
        params = []
        if self.client_var.get():
            query += ' AND clients.name = ?'
            params.append(self.client_var.get())
        if self.type_var.get():
            query += ' AND quotations.type = ?'
            params.append(self.type_var.get())
        if self.date_entry.get():
            query += ' AND quotations.date = ?'
            params.append(self.date_entry.get())
        c.execute(query, params)
        for row in c.fetchall():
            self.tree.insert('', 'end', values=row)

    def reset_filters(self):
        self.client_var.set('')
        self.type_var.set('')
        self.date_entry.delete(0, 'end')
        self.refresh_tree()

    def export_to_excel(self):
        # Export only the visible rows in the treeview
        import pandas as pd
        from tkinter import filedialog
        import os
        rows = [self.tree.item(item)['values'] for item in self.tree.get_children()]
        if not rows:
            messagebox.showinfo("Info", "Aucune donnée à exporter.")
            return
        cols = [self.tree.heading(col)['text'] for col in self.tree['columns']]
        df = pd.DataFrame(rows, columns=cols)
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel", "*.xlsx")],
            initialfile="historique_devis_factures.xlsx",
            title="Choisissez l'emplacement pour enregistrer le fichier Excel"
        )
        if not filename:
            return
        df.to_excel(filename, index=False)
        messagebox.showinfo("Succès", f"Historique exporté dans {os.path.basename(filename)}")

class QuotationApp(tb.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("Générateur de Devis et Factures")
        try:
            self.iconbitmap("MAFCI.ico")
        except Exception as e:
            print("Avertissement : Impossible de charger l'icône MAFCI.ico pour la fenêtre :", e)
        self.geometry("800x600")
        init_db()
        self.create_widgets()
        self.ask_doc_type_and_number()

    def ask_doc_type_and_number(self):
        import tkinter as tk
        from tkinter import messagebox
        import ttkbootstrap as tb
        
        # Create the dialog window
        dialog = tb.Toplevel(self)
        dialog.title("Choix du document")
        dialog.geometry("400x250")  # Made it smaller since we don't need that much space
        dialog.resizable(False, False)
        dialog.transient(self)
        
        # Make the dialog modal
        self.attributes('-disabled', True)
        dialog.grab_set()
        
        # Function to clean up and close the dialog
        def close_dialog():
            try:
                dialog.grab_release()
            except Exception:
                pass
            self.attributes('-disabled', False)
            dialog.destroy()
            
        dialog.protocol("WM_DELETE_WINDOW", close_dialog)
        
        # Center the dialog on screen
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - dialog.winfo_width()) // 2
        y = (dialog.winfo_screenheight() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # Create main frame
        frm = tb.Frame(dialog, padding=20)
        frm.pack(fill='both', expand=True)
        
        # Document type selection
        tb.Label(frm, text="Type de document :").pack(anchor='w', pady=(5, 2))
        self._doc_type_var = tk.StringVar()
        doc_types = ["devis", "facture"]
        type_menu = tb.Combobox(frm, textvariable=self._doc_type_var, values=doc_types, state='readonly')
        type_menu.pack(fill='x', pady=(0, 10))
        
        # Document number entry
        tb.Label(frm, text="Numéro du document :").pack(anchor='w', pady=(5, 2))
        self._doc_number_var = tk.StringVar()
        entry = tb.Entry(frm, textvariable=self._doc_number_var)
        entry.pack(fill='x', pady=(0, 20))
        
        # Set focus to the entry field
        entry.focus()
        
        # OK button handler
        def on_ok():
            doc_type = self._doc_type_var.get().strip()
            doc_number = self._doc_number_var.get().strip()
            
            if not doc_type:
                messagebox.showerror("Erreur", "Veuillez sélectionner le type de document.", parent=dialog)
                return
                
            if not doc_number:
                messagebox.showerror("Erreur", "Veuillez saisir le numéro du document.", parent=dialog)
                return
                
            # If we get here, validation passed
            self.document_type = doc_type
            self.document_number = doc_number
            
            # Update header label with doc info
            self.doc_info_var.set(f"{doc_type.capitalize()} n° {doc_number}")

            # Clean up and close
            close_dialog()

            # Enable the main window widgets
            self.enable_widgets()
        
        # Buttons frame
        btn_frame = tb.Frame(frm)
        btn_frame.pack(fill='x', pady=(10, 0))
        
        # OK button
        ok_btn = tb.Button(btn_frame, text="Valider", command=on_ok, bootstyle="success", width=10)
        ok_btn.pack(side='right', padx=5)
        
        # Cancel button
        cancel_btn = tb.Button(btn_frame, text="Annuler", command=close_dialog, bootstyle="secondary", width=10)
        cancel_btn.pack(side='right')
        
        # Bind Enter key to OK button
        dialog.bind('<Return>', lambda e: on_ok())
        
        # Make sure the main window is disabled while dialog is open
        self.attributes('-disabled', True)
        
        # Wait for the dialog to be closed
        self.wait_window(dialog)


    def enable_widgets(self):
        """Enable the widgets in the main window after the startup dialog."""
        self.attributes('-disabled', False)
        # Combo boxes should remain read-only but not disabled
        self.main_client_type_dropdown.config(state='readonly')
        self.client_dropdown.config(state='readonly')
        self.product_type_dropdown.config(state='readonly')
        # Text entry fields
        self.purchase_order_entry.config(state='normal')
        self.quantity_entry.config(state='normal')
        self.unit_price_entry.config(state='normal')
        # Action buttons
        for btn in (
            self.add_client_btn,
            self.edit_client_btn,
            self.details_client_btn,
            self.generate_pdf_btn,
            self.preview_pdf_btn,
            self.history_btn,
        ):
            btn.config(state='normal')

    def _set_state_recursive(self, widget, state='normal'):
        """Recursively set the state for widget and all of its children."""
        try:
            if 'state' in widget.configure():
                widget.configure(state=state)
        except Exception:
            pass
        for child in widget.winfo_children():
            self._set_state_recursive(child, state)

    def setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use('clam')
        except:
            pass
        style.configure('TButton', font=('Segoe UI', 10), padding=6)
        style.configure('Accent.TButton', font=('Segoe UI', 10, 'bold'), padding=8, foreground='white', background='#0078D4')
        style.map('Accent.TButton', background=[('active', '#005A9E')])
        style.configure('TLabel', font=('Segoe UI', 10))
        style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'))
        style.configure('TEntry', font=('Segoe UI', 10), padding=4)
        style.configure('TLabelframe.Label', font=('Segoe UI', 10, 'bold'))
        # Optionally set branding color from preferences
        # Example: style.configure('Accent.TButton', background=branding_color)
        self.option_add("*Font", ("Segoe UI", 10))
        
    def generate_pdf(self):
        from tkinter import filedialog
        import json

        client_name = self.client_var.get()
        if not client_name:
            messagebox.showerror("Erreur", "Veuillez sélectionner un client")
            return

        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute('SELECT id, nif, rc, address, preferences FROM clients WHERE name=?', (client_name,))
        client = c.fetchone()
        conn.close()
        if not client:
            messagebox.showerror("Erreur", "Client non trouvé")
            return
        client_id, nif, rc, address, client_preferences = client
        if client_preferences:
            try:
                client_preferences = json.loads(client_preferences)
            except Exception:
                client_preferences = {}
        else:
            client_preferences = {}

        doc_type = self.document_type
        doc_number = self.document_number
        purchase_order = self.purchase_order_var.get().strip()
        product = self.product_type_var.get()
        try:
            quantity = float(self.quantity_entry.get())
            unit_price = float(self.unit_price_entry.get())
        except ValueError:
            messagebox.showerror("Erreur", "La quantité et le prix unitaire doivent être des nombres")
            return

        date_str = datetime.now().strftime("%Y-%m-%d")

        # Save to DB
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute(
            '''INSERT INTO quotations (client_id, type, number, product, quantity, unit_price, date, purchase_order)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
            (client_id, doc_type, doc_number, product, quantity, unit_price, date_str, purchase_order),
        )
        conn.commit()
        conn.close()

        default_name = f"{doc_type}_{client_name}_{doc_number}_{date_str}.pdf"
        pdf_filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Fichiers PDF", "*.pdf")],
            initialfile=default_name,
            title="Choisissez l'emplacement pour enregistrer le PDF",
        )
        if not pdf_filename:
            return

        try:
            create_pdf(
                pdf_filename,
                client_name,
                nif,
                rc,
                address,
                client_preferences,
                doc_type,
                doc_number,
                purchase_order,
                product,
                quantity,
                unit_price,
                date_str,
            )
            messagebox.showinfo("Succès", f"PDF généré avec succès : {os.path.basename(pdf_filename)}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la génération du PDF : {e}")

    def on_product_selected(self, event=None):
        """Handle product selection and update unit price automatically."""
        selected_type = self.product_type_var.get()
        if selected_type in self.CONCRETE_PRICES:
            self.unit_price_entry.delete(0, 'end')
            self.unit_price_entry.insert(0, str(self.CONCRETE_PRICES[selected_type]))
            self.update_totals()

    def update_totals(self, event=None):
        try:
            quantity = float(self.quantity_entry.get())
            unit_price = float(self.unit_price_entry.get())
            montant_ht = quantity * unit_price
            tva = montant_ht * 0.16
            ttc = montant_ht + tva
            self.ht_var.set(f"{montant_ht:,.2f}")
            self.tva_var.set(f"{tva:,.2f}")
            self.ttc_var.set(f"{ttc:,.2f}")
        except Exception:
            self.ht_var.set("0.00")
            self.tva_var.set("0.00")
            self.ttc_var.set("0.00")

    def load_clients(self, client_type=None):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        if client_type:
            c.execute('SELECT name FROM clients WHERE client_type=?', (client_type,))
        else:
            c.execute('SELECT name FROM clients')
        rows = [r[0] for r in c.fetchall()]
        conn.close()
        return rows

    def create_widgets(self):
        # --- MAFCI Logo at the top using ttkbootstrap ---
        logo_frame = tb.Frame(self)
        logo_frame.pack(side='top', fill='x', pady=(10, 0))
        try:
            logo_img = Image.open(LOGO_PATH)
            try:
                resample = Image.Resampling.LANCZOS
            except AttributeError:
                resample = Image.ANTIALIAS
            logo_img = logo_img.resize((120, 60), resample)
            self.logo_photo = ImageTk.PhotoImage(logo_img)
            tb.Label(logo_frame, image=self.logo_photo, bootstyle="light").pack(side='left', pady=5)
        except Exception as e:
            print(f"Échec du chargement du logo MAFCI: {e}")

        self.doc_info_var = tk.StringVar(value="")
        self.doc_info_label = tb.Label(logo_frame, textvariable=self.doc_info_var, style='Header.TLabel')
        self.doc_info_label.pack(side='left', padx=10)

        # --- Main content frame for the rest of the UI ---
        content_frame = tb.Frame(self)
        content_frame.pack(side='top', fill='both', expand=True)


        # Client type selector
        client_frame = ttk.LabelFrame(content_frame, text="Client")
        client_frame.grid(row=1, column=0, columnspan=2, padx=15, pady=10, sticky='nsew')
        ttk.Label(client_frame, text="Type de client :").grid(row=0, column=0, sticky='e', padx=5, pady=4)
        self.main_client_type_var = tk.StringVar()
        self.main_client_type_dropdown = ttk.Combobox(client_frame, textvariable=self.main_client_type_var, state='disabled')
        self.main_client_type_dropdown['values'] = ('', 'ciment', 'beton')
        self.main_client_type_dropdown.grid(row=0, column=1, sticky='w', padx=5, pady=4)
        self.main_client_type_dropdown.bind('<<ComboboxSelected>>', self.update_clients_for_type)

        ttk.Label(client_frame, text="Client :").grid(row=1, column=0, sticky='e', padx=5, pady=4)
        self.client_var = tk.StringVar()
        self.client_dropdown = ttk.Combobox(client_frame, textvariable=self.client_var, state='disabled')
        self.client_dropdown.grid(row=1, column=1, sticky='w', padx=5, pady=4)
        self.client_dropdown['values'] = self.load_clients()
        self.client_dropdown.bind('<<ComboboxSelected>>', self.update_product_types)

        self.add_client_btn = ttk.Button(client_frame, text="Ajouter un client...", command=self.open_add_client, style='Accent.TButton', state='disabled')
        self.add_client_btn.grid(row=2, column=0, columnspan=2, padx=5, pady=7, sticky='ew')
        self.edit_client_btn = ttk.Button(client_frame, text="Modifier un client...", command=self.open_edit_client, state='disabled')
        self.edit_client_btn.grid(row=3, column=0, columnspan=2, padx=5, pady=7, sticky='ew')
        self.details_client_btn = ttk.Button(client_frame, text="Voir les détails", command=self.show_client_details, state='disabled')
        self.details_client_btn.grid(row=4, column=0, columnspan=2, padx=5, pady=7, sticky='ew')

        # Product type selector
        product_frame = ttk.LabelFrame(content_frame, text="Produit")
        product_frame.grid(row=1, column=2, columnspan=2, padx=15, pady=10, sticky='nsew')
        ttk.Label(product_frame, text="Désignation :").grid(row=0, column=0, sticky='e', padx=5, pady=4)
        self.product_type_var = tk.StringVar()
        self.product_type_dropdown = ttk.Combobox(product_frame, textvariable=self.product_type_var, state='disabled')
        self.product_type_dropdown.grid(row=0, column=1, sticky='w', padx=5, pady=4)
        self.product_type_dropdown.bind('<<ComboboxSelected>>', self.on_product_selected)

        ttk.Label(product_frame, text="Bon de commande :").grid(row=1, column=0, sticky='e', padx=5, pady=4)
        self.purchase_order_var = tk.StringVar()
        self.purchase_order_entry = ttk.Entry(product_frame, textvariable=self.purchase_order_var, width=18, state='disabled')
        self.purchase_order_entry.grid(row=1, column=1, sticky='w', padx=5, pady=4)

        ttk.Label(product_frame, text="Quantité :").grid(row=2, column=0, sticky='e', padx=5, pady=4)
        self.quantity_entry = ttk.Entry(product_frame, width=18, state='disabled')
        self.quantity_entry.grid(row=2, column=1, sticky='w', padx=5, pady=4)

        ttk.Label(product_frame, text="Prix unitaire :").grid(row=3, column=0, sticky='e', padx=5, pady=4)
        self.unit_price_entry = ttk.Entry(product_frame, width=18, state='disabled')
        self.unit_price_entry.grid(row=3, column=1, sticky='w', padx=5, pady=4)

        # Auto-calculated fields
        totals_frame = ttk.LabelFrame(content_frame, text="Montants")
        # Align totals with the client frame for a cleaner look
        totals_frame.grid(row=2, column=0, columnspan=2, padx=15, pady=10, sticky='nsew')
        self.ht_var = tk.StringVar(value="0.00")
        self.tva_var = tk.StringVar(value="0.00")
        self.ttc_var = tk.StringVar(value="0.00")
        ttk.Label(totals_frame, text="Montant HT :").grid(row=0, column=0, sticky='e', padx=5, pady=4)
        ttk.Label(totals_frame, textvariable=self.ht_var, style='Header.TLabel').grid(row=0, column=1, sticky='w', padx=5, pady=4)
        ttk.Label(totals_frame, text="TVA (16%) :").grid(row=1, column=0, sticky='e', padx=5, pady=4)
        ttk.Label(totals_frame, textvariable=self.tva_var, style='Header.TLabel').grid(row=1, column=1, sticky='w', padx=5, pady=4)
        ttk.Label(totals_frame, text="Montant TTC :").grid(row=2, column=0, sticky='e', padx=5, pady=4)
        ttk.Label(totals_frame, textvariable=self.ttc_var, style='Header.TLabel').grid(row=2, column=1, sticky='w', padx=5, pady=4)

        # Bind events for live calculation
        self.quantity_entry.bind('<KeyRelease>', self.update_totals)
        self.unit_price_entry.bind('<KeyRelease>', self.update_totals)

        # Generate button
        actions_frame = ttk.LabelFrame(content_frame, text="Actions")
        actions_frame.grid(row=3, column=0, columnspan=4, padx=15, pady=10, sticky='nsew')
        self.generate_pdf_btn = ttk.Button(actions_frame, text="Générer le PDF", command=self.generate_pdf, style='Accent.TButton', state='disabled')
        self.generate_pdf_btn.grid(row=0, column=0, padx=10, pady=8, sticky='ew')
        self.preview_pdf_btn = ttk.Button(actions_frame, text="Prévisualiser le PDF", command=self.preview_pdf, state='disabled')
        self.preview_pdf_btn.grid(row=0, column=1, padx=10, pady=8, sticky='ew')
        self.history_btn = ttk.Button(actions_frame, text="Historique", command=self.open_history_window, state='disabled')
        self.history_btn.grid(row=1, column=0, columnspan=2, padx=10, pady=8, sticky='ew')

    def update_clients_for_type(self, event=None):
        ctype = self.main_client_type_var.get()
        self.client_dropdown.set('')
        self.product_type_dropdown.set('')
        self.product_type_dropdown['values'] = []
        if ctype:
            clients = self.load_clients(ctype)
            self.client_dropdown['values'] = clients
        else:
            self.client_dropdown['values'] = []

    def get_client_type(self, client_name):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute('SELECT client_type FROM clients WHERE name=?', (client_name,))
        row = c.fetchone()
        conn.close()
        return row[0] if row else None

    # Concrete type to price mapping (in MRU/m³)
    CONCRETE_PRICES = {
        'Béton C30 SR': 7772,
        'Béton C25 SR': 5700,
        'Béton C20 SR': 4700,
        'Béton C15 SR': 4300,
        'Béton C30': 0000,
        'Béton C25': 4700,
        'Béton C20': 4500,
        'Béton C15': 4300,
        'Béton C35 SR': 8352,
        'Béton C45 SR': 0000,

    }

    def update_product_types(self, event=None):
        cement_types = ['Ciment 42.5', 'Ciment 32.5', 'Ciment SR']
        concrete_types = ['Béton C15', 'Béton C20', 'Béton C20 SR', 'Béton C25', 'Béton C25 SR', 'Béton C45 SR', 'Béton C30', 'Béton C30 SR', 'Béton C35 SR', 'Béton C15 SR']
        client_name = self.client_var.get()
        self.product_type_dropdown.set('')
        if not client_name:
            self.product_type_dropdown['values'] = []
            self.product_type_var.set('')
            return
        client_type = self.get_client_type(client_name)
        if client_type == 'ciment':
            self.product_type_dropdown['values'] = cement_types
        elif client_type == 'beton':
            self.product_type_dropdown['values'] = concrete_types
        else:
            self.product_type_dropdown['values'] = []
        if self.product_type_dropdown['values']:
            self.product_type_var.set(self.product_type_dropdown['values'][0])
            # Auto-set the unit price when a product type is selected
            selected_type = self.product_type_var.get()
            if selected_type in self.CONCRETE_PRICES:
                self.unit_price_entry.delete(0, 'end')
                self.unit_price_entry.insert(0, str(self.CONCRETE_PRICES[selected_type]))
                self.update_totals()
        else:
            self.product_type_var.set('')
            self.unit_price_entry.delete(0, 'end')

    def open_add_client(self):
        AddClientWindow(self)

    def open_edit_client(self):
        EditClientWindow(self)

    def on_doc_type_selected(self, event=None):
        doc_type = self.document_type_var.get()
        enable = bool(doc_type)
        state = 'normal' if enable else 'disabled'
        # Enable/disable all main fields
        self.main_client_type_dropdown.config(state=state)
        self.client_dropdown.config(state=state)
        self.product_type_dropdown.config(state=state)
        self.purchase_order_entry.config(state=state)
        self.quantity_entry.config(state=state)
        self.unit_price_entry.config(state=state)
        self.add_client_btn.config(state=state)
        self.edit_client_btn.config(state=state)
        self.details_client_btn.config(state=state)
        self.generate_pdf_btn.config(state=state)
        self.preview_pdf_btn.config(state=state)
        self.history_btn.config(state=state)

    def show_client_details(self):
        client_name = self.client_var.get()
        if not client_name:
            messagebox.showerror("Erreur", "Veuillez sélectionner un client")
            return
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute('SELECT name, nif, rc, address, client_type, preferences FROM clients WHERE name=?', (client_name,))
        row = c.fetchone()
        conn.close()
        if not row:
            messagebox.showerror("Erreur", "Client non trouvé")
            return
        name, nif, rc, address, client_type, preferences = row
        import json
        try:
            preferences = json.loads(preferences) if preferences else {}
        except Exception:
            preferences = {}
        details = f"Nom : {name}\nNIF : {nif}\nRC : {rc}\nAdresse : {address}\nType : {client_type}\n"
        if preferences:
            details += "\nPréférences :\n"
            for k, v in preferences.items():
                details += f"  {k} : {v}\n"
        detail_win = tk.Toplevel(self)
        detail_win.title(f"Détails du client : {name}")
        detail_win.geometry("400x350")
        tk.Label(detail_win, text=details, justify='left', anchor='nw').pack(fill='both', expand=True, padx=15, pady=15)

    def open_history_window(self):
        HistoryWindow(self)

    def refresh_clients(self):
        self.client_dropdown['values'] = self.load_clients()

    def preview_pdf(self):
        import tempfile

        client_name = self.client_var.get()
        if not client_name:
            messagebox.showerror("Erreur", "Veuillez sélectionner un client")
            return

        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute('SELECT id, nif, rc, address, preferences FROM clients WHERE name=?', (client_name,))
        client = c.fetchone()
        conn.close()
        if not client:
            messagebox.showerror("Erreur", "Client non trouvé")
            return
        client_id, nif, rc, address, client_preferences = client
        if client_preferences:
            try:
                client_preferences = json.loads(client_preferences)
            except Exception:
                client_preferences = {}
        else:
            client_preferences = {}

        doc_type = self.document_type
        doc_number = self.document_number
        purchase_order = self.purchase_order_var.get().strip()
        product = self.product_type_var.get()
        try:
            quantity = float(self.quantity_entry.get())
            unit_price = float(self.unit_price_entry.get())
        except ValueError:
            messagebox.showerror("Erreur", "La quantité et le prix unitaire doivent être des nombres")
            return

        date_str = datetime.now().strftime("%Y-%m-%d")

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        pdf_filename = tmp.name
        tmp.close()

        try:
            create_pdf(
                pdf_filename,
                client_name,
                nif,
                rc,
                address,
                client_preferences,
                doc_type,
                doc_number,
                purchase_order,
                product,
                quantity,
                unit_price,
                date_str,
            )
            PDFPreviewWindow(self, pdf_filename)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de générer l'aperçu : {e}")

    def export_history_to_excel(self):
        import pandas as pd
        from tkinter import filedialog
        import os
        rows = [self.tree.item(item)['values'] for item in self.tree.get_children()]
        if not rows:
            messagebox.showinfo("Info", "Aucune donnée à exporter.")
            return
        cols = [self.tree.heading(col)['text'] for col in self.tree['columns']]
        df = pd.DataFrame(rows, columns=cols)
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel", "*.xlsx")],
            initialfile="historique_devis_factures.xlsx",
            title="Choisissez l'emplacement pour enregistrer le fichier Excel"
        )
        if not filename:
            return
        df.to_excel(filename, index=False)
        messagebox.showinfo("Succès", f"Historique exporté dans {os.path.basename(filename)}")


if __name__ == '__main__':
    app = QuotationApp()
    app.mainloop()

