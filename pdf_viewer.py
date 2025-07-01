import tkinter as tk
from tkinter import Toplevel, Button, messagebox
from PIL import Image, ImageTk
from pdf2image import convert_from_path
import os
import sys

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
            # Try to find poppler in common locations
            poppler_paths = [
                r"C:\Program Files\poppler-23.11.0\Library\bin",  # Common Windows path
                r"C:\Program Files\poppler\Library\bin",
                r"C:\poppler\Library\bin",
                r"C:\Program Files (x86)\poppler\Library\bin",
                r"C:\Program Files\poppler-0.68.0\bin",  # Version might vary
                r"C:\poppler-0.68.0\bin",
                r"C:\poppler-23.11.0\Library\bin"  # Latest version as of now
            ]
            
            # Try each path until one works
            for path in poppler_paths:
                if os.path.exists(path):
                    try:
                        self.pages = convert_from_path(pdf_path, dpi=120, poppler_path=path)
                        break
                    except Exception:
                        continue
            else:
                # If no path worked, try without specifying poppler_path (in case it's in PATH)
                try:
                    self.pages = convert_from_path(pdf_path, dpi=120)
                except Exception as e:
                    messagebox.showerror("Erreur", f"Impossible de lire le PDF. Assurez-vous que Poppler est installé.\n\nDétails: {e}")
                    self.destroy()
                    return
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
        if not hasattr(self, 'pages') or not self.pages or self.current_page >= len(self.pages):
            return
            
        img = self.pages[self.current_page]
        w, h = img.size
        
        # Get canvas dimensions with minimum size
        canvas_w = max(100, self.canvas.winfo_width())
        canvas_h = max(100, self.canvas.winfo_height())
        
        # Calculate scale with safety checks
        if w <= 0 or h <= 0 or canvas_w <= 0 or canvas_h <= 0:
            return
            
        scale = min(canvas_w/w, canvas_h/h) * self.zoom
        new_w = max(1, int(w * scale))  # Ensure at least 1 pixel
        new_h = max(1, int(h * scale))  # Ensure at least 1 pixel
        
        try:
            img_resized = img.resize((new_w, new_h), Image.LANCZOS)
            self.tk_img = ImageTk.PhotoImage(img_resized)
            self.canvas.delete('all')
            self.canvas.create_image((canvas_w-new_w)//2, (canvas_h-new_h)//2, anchor='nw', image=self.tk_img)
            self.canvas.config(scrollregion=self.canvas.bbox('all'))
        except Exception as e:
            print(f"Error displaying image: {e}")

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
