import os
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from pdf2docx import Converter as pdf_to_docx
from docx2pdf import convert as docx_to_pdf
from PIL import Image
from pdf2image import convert_from_path

class DocuShiftLocal(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuración de la Ventana
        self.title("DocuShift Pro - Gestión de Documentos")
        self.geometry("900x700")
        ctk.set_appearance_mode("dark")
        
        # --- LAYOUT ---
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Sidebar
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        
        self.logo = ctk.CTkLabel(self.sidebar, text="DocuShift\nLOCAL", font=("Roboto", 24, "bold"))
        self.logo.pack(pady=30)

        ctk.CTkButton(self.sidebar, text="Modo Claro/Oscuro", command=self.toggle_theme).pack(pady=10, padx=10)
        
        # Main View
        self.main_view = ctk.CTkScrollableFrame(self, corner_radius=0, fg_color="transparent")
        self.main_view.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

        self.render_ui()

    def toggle_theme(self):
        mode = "light" if ctk.get_appearance_mode() == "Dark" else "dark"
        ctk.set_appearance_mode(mode)

    def render_ui(self):
        # 1. CONVERTIR
        self.create_card("🔄 CONVERSIÓN", [
            ("Word a PDF (Carpeta)", self.word_to_pdf, "#3498db"),
            ("PDF a Word", self.pdf_to_word, "#3498db"),
            ("Imágenes a PDF", self.img_to_pdf, "#2ecc71"),
            ("PDF a Imágenes (JPG)", self.pdf_to_img, "#d35400")
        ])

        # 2. ORGANIZAR
        self.create_card("📂 ORGANIZAR Y EDITAR", [
            ("Unir varios PDFs", self.merge_pdfs, "#9b59b6"),
            ("Dividir PDF (por páginas)", self.split_pdf, "#9b59b6"),
            ("Girar PDF (90°)", self.rotate_pdf, "#9b59b6"),
            ("Eliminar página específica", self.delete_page, "#c0392b")
        ])

        # 3. SEGURIDAD
        self.create_card("🔒 SEGURIDAD", [
            ("Proteger con Contraseña", self.lock_pdf, "#2c3e50"),
            ("Quitar Restricciones", self.unlock_pdf, "#2c3e50")
        ])

    def create_card(self, title, buttons):
        lbl = ctk.CTkLabel(self.main_view, text=title, font=("Roboto", 18, "bold"), text_color="#3498db")
        lbl.pack(pady=(20, 10), anchor="w")

        frame = ctk.CTkFrame(self.main_view, fg_color="transparent")
        frame.pack(fill="x", padx=10)
        
        for i, (text, cmd, color) in enumerate(buttons):
            btn = ctk.CTkButton(frame, text=text, command=cmd, fg_color=color, height=50)
            btn.grid(row=i//2, column=i%2, padx=10, pady=10, sticky="ew")
            frame.grid_columnconfigure(i%2, weight=1)

    # --- LÓGICA ---

    def word_to_pdf(self):
        folder = filedialog.askdirectory()
        if folder:
            try:
                docx_to_pdf(folder)
                messagebox.showinfo("Éxito", "Documentos convertidos.")
            except Exception as e: messagebox.showerror("Error", str(e))

    def pdf_to_word(self):
        f = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if f:
            cv = pdf_to_docx(f); cv.convert(f.replace(".pdf", ".docx")); cv.close()
            messagebox.showinfo("Éxito", "Word creado.")

    def merge_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if files:
            m = PdfMerger()
            for f in files: m.append(f)
            out = filedialog.asksaveasfilename(defaultextension=".pdf")
            if out: m.write(out); m.close(); messagebox.showinfo("OK", "PDFs unidos.")

    def img_to_pdf(self):
        files = filedialog.askopenfilenames(filetypes=[("Imágenes", "*.jpg *.png")])
        if files:
            imgs = [Image.open(f).convert("RGB") for f in files]
            out = filedialog.asksaveasfilename(defaultextension=".pdf")
            if out: imgs[0].save(out, save_all=True, append_images=imgs[1:])

    def pdf_to_img(self):
        f = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if f:
            try:
                images = convert_from_path(f)
                folder = filedialog.askdirectory()
                if folder:
                    for i, img in enumerate(images): img.save(f"{folder}/p_{i+1}.jpg", "JPEG")
                    messagebox.showinfo("OK", "Imágenes guardadas.")
            except: messagebox.showerror("Error", "Necesitas instalar Poppler para esta función.")

    def split_pdf(self):
        f = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if f:
            r = PdfReader(f); folder = filedialog.askdirectory()
            if folder:
                for i, p in enumerate(r.pages):
                    w = PdfWriter(); w.add_page(p)
                    with open(f"{folder}/p_{i+1}.pdf", "wb") as out: w.write(out)
                messagebox.showinfo("OK", "PDF dividido.")

    def rotate_pdf(self):
        f = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if f:
            r = PdfReader(f); w = PdfWriter()
            for p in r.pages: p.rotate(90); w.add_page(p)
            with open(f.replace(".pdf", "_rotado.pdf"), "wb") as out: w.write(out)
            messagebox.showinfo("OK", "Rotado 90°.")

    def delete_page(self):
        f = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if f:
            num = simpledialog.askinteger("Eliminar", "¿Qué número de página quieres quitar?")
            if num:
                r = PdfReader(f); w = PdfWriter()
                for i, p in enumerate(r.pages):
                    if i != (num - 1): w.add_page(p)
                with open(f.replace(".pdf", "_modificado.pdf"), "wb") as out: w.write(out)
                messagebox.showinfo("OK", f"Página {num} eliminada.")

    def lock_pdf(self):
        f = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if f:
            pw = simpledialog.askstring("Seguridad", "Nueva contraseña:", show="*")
            if pw:
                r = PdfReader(f); w = PdfWriter()
                for p in r.pages: w.add_page(p)
                w.encrypt(pw)
                with open(f.replace(".pdf", "_protegido.pdf"), "wb") as out: w.write(out)

    def unlock_pdf(self):
        f = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if f:
            pw = simpledialog.askstring("Desbloquear", "Introduce la contraseña actual:", show="*")
            if pw:
                r = PdfReader(f); w = PdfWriter()
                if r.is_encrypted: r.decrypt(pw)
                for p in r.pages: w.add_page(p)
                with open(f.replace(".pdf", "_desbloqueado.pdf"), "wb") as out: w.write(out)

if __name__ == "__main__":
    app = DocuShiftLocal()
    app.mainloop()