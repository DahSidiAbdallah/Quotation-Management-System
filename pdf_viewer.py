import tkinter as tk
from tkinter import Toplevel, Button, messagebox
from PIL import Image, ImageTk
from pdf2image import convert_from_path
import os

class PDFPreviewWindow(Toplevel):
    def __init__(self, parent, pdf_path, on_generate_callback=None):
        super().__init__(parent)
        self.title("Aperçu du PDF")
        self.geometry("500x650")
        self.pdf_path = pdf_path
        self.on_generate_callback = on_generate_callback
        self.images = []
        self.img_labels = []
        self.current_page = 0
        self.pages = []
        try:
            # You may need to set poppler_path if not in PATH
            self.pages = convert_from_path(pdf_path, dpi=120, poppler_path=r"C:\poppler\Library\bin")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire le PDF : {e}")
            self.destroy()
            return
        self.zoom = 1.0
        self.canvas = tk.Canvas(self, bg='gray')
        self.canvas.pack(fill='both', expand=True)
        self.btn_frame = tk.Frame(self)
        self.btn_frame.pack(fill='x', pady=6)
        self.prev_btn = Button(self.btn_frame, text="Page précédente", command=self.show_prev_page)
        self.prev_btn.pack(side='left', padx=5)
        self.next_btn = Button(self.btn_frame, text="Page suivante", command=self.show_next_page)
        self.next_btn.pack(side='left', padx=5)
        self.zoom_in_btn = Button(self.btn_frame, text="Zoom +", command=self.zoom_in)
        self.zoom_in_btn.pack(side='left', padx=5)
        self.zoom_out_btn = Button(self.btn_frame, text="Zoom -", command=self.zoom_out)
        self.zoom_out_btn.pack(side='left', padx=5)
        self.gen_btn = Button(self.btn_frame, text="Générer ce PDF", command=self.generate_pdf)
        self.gen_btn.pack(side='right', padx=5)
        self.bind('<Configure>', self.on_resize)
        self.show_page(0)

    def show_page(self, page_num):
        if not self.pages:
            return
        self.current_page = page_num
        self.display_image()
        self.title(f"Aperçu du PDF - Page {page_num+1} / {len(self.pages)}")

    def display_image(self):
        img = self.pages[self.current_page]
        w, h = img.size
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        # Fit to window with zoom
        scale = min(canvas_w/w, canvas_h/h) * self.zoom
        new_w = int(w * scale)
        new_h = int(h * scale)
        img_resized = img.resize((new_w, new_h), Image.LANCZOS)
        self.tk_img = ImageTk.PhotoImage(img_resized)
        self.canvas.delete('all')
        self.canvas.create_image((canvas_w-new_w)//2, (canvas_h-new_h)//2, anchor='nw', image=self.tk_img)
        self.canvas.config(scrollregion=self.canvas.bbox('all'))

    def on_resize(self, event):
        self.display_image()

    def zoom_in(self):
        self.zoom *= 1.2
        self.display_image()

    def zoom_out(self):
        self.zoom /= 1.2
        self.display_image()

    def show_prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.display_image()
            self.title(f"Aperçu du PDF - Page {self.current_page+1} / {len(self.pages)}")

    def show_next_page(self):
        if self.current_page < len(self.pages) - 1:
            self.current_page += 1
            self.display_image()
            self.title(f"Aperçu du PDF - Page {self.current_page+1} / {len(self.pages)}")

    def generate_pdf(self):
        from tkinter import filedialog, messagebox
        import shutil
        dest = filedialog.asksaveasfilename(defaultextension='.pdf', filetypes=[('PDF files', '*.pdf')], title='Enregistrer le PDF')
        if dest:
            try:
                shutil.copy(self.pdf_path, dest)
                messagebox.showinfo('Succès', 'PDF enregistré avec succès.')
            except Exception as e:
                messagebox.showerror('Erreur', f'Impossible d\'enregistrer le PDF : {e}')
        self.destroy()
