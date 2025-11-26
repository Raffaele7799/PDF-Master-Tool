# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
import sys
import subprocess
import shutil
import threading
import time
import webbrowser  # <--- Nuova libreria per aprire il link Ko-fi

# Gestione dipendenze mancanti
try:
    from pypdf import PdfWriter, PdfReader
    from pypdf.generic import NameObject, DecodedStreamObject, DictionaryObject, NumberObject, ArrayObject, NullObject
    import fitz  # PyMuPDF
    from PIL import Image, ImageTk
    import ttkbootstrap as ttk
    from ttkbootstrap.constants import *
except ImportError as e:
    import tkinter.messagebox
    root = tk.Tk()
    root.withdraw()
    tkinter.messagebox.showerror("Librerie Mancanti", f"Errore: {e}\n\nInstalla le librerie necessarie:\npip install pypdf pymupdf pillow ttkbootstrap")
    sys.exit()

# --- FUNZIONE RICERCA GHOSTSCRIPT (SOLO LOCALE) ---
def find_ghostscript():
    """
    Cerca Ghostscript SOLO localmente per massima velocità.
    Non scansiona il sistema operativo.
    
    Cerca in ordine:
    1. Interno all'EXE (_MEIPASS) -> Se impacchettato con --add-data
    2. Accanto all'EXE (o allo script)
    3. Sottocartella 'bin' o 'gs' accanto all'EXE
    """
    gs_name = "gswin64c.exe"
    potential_paths = []

    # 1. Controllo interno (PyInstaller Bundle)
    if hasattr(sys, '_MEIPASS'):
        potential_paths.append(os.path.join(sys._MEIPASS, gs_name))

    # 2. Determina cartella base (EXE o Script)
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    # 3. Aggiungi percorsi relativi esterni
    potential_paths.append(os.path.join(base_path, gs_name))            # Es: ./gswin64c.exe
    potential_paths.append(os.path.join(base_path, "bin", gs_name))    # Es: ./bin/gswin64c.exe
    potential_paths.append(os.path.join(base_path, "gs", gs_name))     # Es: ./gs/gswin64c.exe

    for p in potential_paths:
        if os.path.exists(p):
            return p

    return None

def convert_to_pdfa_ghostscript(input_pdf, output_pdf, gs_exe_path):
    """Usa Ghostscript per convertire in PDF/A-2b Reale."""
    if not gs_exe_path:
        return False, "Ghostscript non trovato."

    # Comandi per Ghostscript (PDF/A-2b RGB)
    cmd = [
        gs_exe_path,
        "-dPDFA=2",
        "-dBATCH",
        "-dNOPAUSE",
        "-dNOOUTERSAVE",
        "-sColorConversionStrategy=RGB", 
        "-sProcessColorModel=DeviceRGB",
        "-sDEVICE=pdfwrite",
        "-dPDFACompatibilityPolicy=1", 
        f"-sOutputFile={output_pdf}",
        input_pdf
    ]

    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        
        process = subprocess.run(
            cmd, 
            stdout=subprocess.PIPE, 
            stderr=subprocess.PIPE, 
            startupinfo=startupinfo,
            text=True
        )
        
        if process.returncode == 0:
            return True, "Conversione OK"
        else:
            return False, f"Errore GS: {process.stderr}"
    except Exception as e:
        return False, str(e)

# --- CLASSE DIALOGO ROTAZIONE ---
class RotatePreviewDialog(tk.Toplevel):
    def __init__(self, parent, pdf_path, callback_update_list, start_angle=0):
        super().__init__(parent)
        self.title(f"Ruota PDF - {os.path.basename(pdf_path)}")
        self.geometry("900x850")
        self.pdf_path = pdf_path
        self.callback = callback_update_list
        self.angle = start_angle % 360 
        self.doc = None
        self.zoom = 0.4 
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # --- TOOLBAR ---
        toolbar = ttk.Frame(self, padding=10, bootstyle="secondary")
        toolbar.pack(fill=X, side=TOP)
        
        ttk.Label(toolbar, text="Angolo:", bootstyle="inverse-secondary").pack(side=LEFT)
        self.lbl_angle = ttk.Label(toolbar, text=f"{self.angle}°", font=("Segoe UI", 12, "bold"), bootstyle="inverse-secondary", width=5)
        self.lbl_angle.pack(side=LEFT, padx=5)
        
        # --- PULSANTI ROTAZIONE ---
        ttk.Button(toolbar, text="↪️ Ruota Sx (-90°)", command=lambda: self.rotate_view(-90), bootstyle="info-outline").pack(side=LEFT, padx=5)
        ttk.Button(toolbar, text="↪️ Ruota Dx (+90°)", command=lambda: self.rotate_view(90), bootstyle="info").pack(side=LEFT, padx=5)

        ttk.Separator(toolbar, orient=VERTICAL).pack(side=LEFT, fill=Y, padx=10)
        
        # Pulsanti Zoom
        ttk.Button(toolbar, text="🔍 -", command=self.zoom_out, bootstyle="secondary-outline", width=4).pack(side=LEFT)
        self.lbl_zoom = ttk.Label(toolbar, text=f"{int(self.zoom*100)}%", bootstyle="inverse-secondary", width=6, anchor=CENTER)
        self.lbl_zoom.pack(side=LEFT, padx=2)
        ttk.Button(toolbar, text="🔍 +", command=self.zoom_in, bootstyle="secondary-outline", width=4).pack(side=LEFT)
        
        ttk.Button(toolbar, text="💾 CONFERMA", command=self.save_changes, bootstyle="success").pack(side=RIGHT)

        # --- AREA CANVAS ---
        self.canvas = tk.Canvas(self, bg="#404040")
        self.scrollbar_y = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollbar_x = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.scroll_frame = ttk.Frame(self.canvas)
        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="center")
        self.canvas.configure(yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set)
        self.scrollbar_y.pack(side=RIGHT, fill=Y)
        self.scrollbar_x.pack(side=BOTTOM, fill=X)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)
        
        self.image_label = tk.Label(self.scroll_frame, bg="#404040")
        self.image_label.pack()
        
        try:
            self.doc = fitz.open(pdf_path)
            self.render_preview()
        except Exception as e:
            messagebox.showerror("Errore Apertura", str(e))
            self.destroy()

    def rotate_view(self, delta_angle):
        self.angle = (self.angle + delta_angle) % 360
        self.lbl_angle.config(text=f"{self.angle}°")
        self.render_preview()

    def zoom_in(self): 
        self.zoom = round(self.zoom + 0.1, 2)
        self.render_preview()
    
    def zoom_out(self): 
        if self.zoom > 0.1: 
            self.zoom = round(self.zoom - 0.1, 2)
            self.render_preview()

    def render_preview(self):
        if not self.doc: return
        try:
            self.lbl_zoom.config(text=f"{int(self.zoom*100)}%")
            page = self.doc.load_page(0)
            mat = fitz.Matrix(self.zoom, self.zoom).prerotate(self.angle)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img_data = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            self.tk_img = ImageTk.PhotoImage(img_data)
            self.image_label.config(image=self.tk_img)
            self.canvas.update_idletasks()
            self.canvas.config(scrollregion=self.canvas.bbox("all"))
        except: pass

    def on_close(self):
        if self.doc:
            self.doc.close()
        self.destroy()

    def save_changes(self):
        if self.angle == 0:
            if not messagebox.askyesno("Info", "Nessuna rotazione applicata. Chiudere senza salvare?"): return
            self.on_close()
            return
        
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=f"Ruotato_{os.path.basename(self.pdf_path)}")
        if save_path:
            try:
                # Chiudiamo fitz per liberare il file
                if self.doc: self.doc.close()

                reader = PdfReader(self.pdf_path)
                writer = PdfWriter()
                for page in reader.pages:
                    page.rotate(self.angle)
                    writer.add_page(page)
                
                with open(save_path, "wb") as f:
                    writer.write(f)
                
                if messagebox.askyesno("Fatto", "Aggiornare il file nella lista con quello ruotato?"): 
                    self.callback(self.pdf_path, save_path)
                
                self.destroy()
            except Exception as e:
                messagebox.showerror("Errore", str(e))

# --- CLASSE ANTEPRIMA ZOOM ---
class ZoomPreviewWindow(tk.Toplevel):
    def __init__(self, parent, pdf_list, start_index):
        super().__init__(parent)
        self.geometry("800x700")
        self.pdf_list = pdf_list
        self.current_index = start_index
        self.zoom = 0.25 
        self.doc = None
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        nav_frame = ttk.Frame(self, padding=5, bootstyle="secondary")
        nav_frame.pack(fill=X, side=TOP)
        
        ttk.Button(nav_frame, text="⬅ Prev", command=self.prev_file, bootstyle="light").pack(side=LEFT, padx=5)
        ttk.Button(nav_frame, text="Next ➡", command=self.next_file, bootstyle="light").pack(side=LEFT, padx=5)
        
        self.lbl_title = ttk.Label(nav_frame, text="", font=("Segoe UI", 10, "bold"), bootstyle="inverse-secondary")
        self.lbl_title.pack(side=LEFT, padx=20)
        
        ttk.Button(nav_frame, text="🔍 +", command=self.zoom_in, bootstyle="info-outline", width=5).pack(side=RIGHT, padx=2)
        ttk.Button(nav_frame, text="🔍 -", command=self.zoom_out, bootstyle="info-outline", width=5).pack(side=RIGHT, padx=2)
        self.lbl_zoom = ttk.Label(nav_frame, text="25%", width=6, anchor=CENTER, bootstyle="inverse-secondary")
        self.lbl_zoom.pack(side=RIGHT, padx=5)

        self.canvas = tk.Canvas(self, bg="#404040")
        self.scrollbar_y = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollbar_x = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.scroll_frame = ttk.Frame(self.canvas)
        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="center")
        self.canvas.configure(yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set)
        self.scrollbar_y.pack(side=RIGHT, fill=Y)
        self.scrollbar_x.pack(side=BOTTOM, fill=X)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)
        
        self.image_label = tk.Label(self.scroll_frame, bg="#404040")
        self.image_label.pack()
        
        self.load_current_file()

    def load_current_file(self):
        if self.doc: self.doc.close()
        path = self.pdf_list[self.current_index]
        if not os.path.exists(path): return
        self.title(f"Anteprima [{self.current_index + 1}/{len(self.pdf_list)}] - {os.path.basename(path)}")
        self.lbl_title.config(text=f"{self.current_index + 1}/{len(self.pdf_list)}: {os.path.basename(path)}")
        try:
            self.doc = fitz.open(path)
            self.render_page()
        except Exception as e: messagebox.showerror("Errore", str(e))

    def render_page(self):
        if not self.doc: return
        try:
            page = self.doc.load_page(0)
            mat = fitz.Matrix(self.zoom, self.zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img_data = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            self.tk_img = ImageTk.PhotoImage(img_data)
            self.image_label.config(image=self.tk_img)
            self.lbl_zoom.config(text=f"{int(self.zoom*100)}%")
            self.canvas.update_idletasks()
            self.canvas.config(scrollregion=self.canvas.bbox("all"))
        except: pass

    def next_file(self):
        if self.current_index < len(self.pdf_list) - 1: self.current_index += 1; self.load_current_file()
    def prev_file(self):
        if self.current_index > 0: self.current_index -= 1; self.load_current_file()
    
    def zoom_in(self): 
        self.zoom = round(self.zoom + 0.10, 2) 
        self.render_page()
        
    def zoom_out(self): 
        if self.zoom > 0.15: 
            self.zoom = round(self.zoom - 0.10, 2)
            self.render_page()
            
    def on_close(self):
        if self.doc: self.doc.close()
        self.destroy()

# --- APP PRINCIPALE v17.0 (Logic Lab Edition) ---
class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.style = ttk.Style()
        self.configure_table_style()

        self.root.title("PDF Master Tool v17.0 - Powered by Logic Lab ⚡")
        self.root.geometry("1200x850")
        self.pdf_list = [] 
        self.is_processing = False
        
        # --- CONTROLLO GHOSTSCRIPT IMMEDIATO (SOLO LOCALE) ---
        # Essendo locale è velocissimo, non serve thread background
        self.gs_path = find_ghostscript()

        # SIDEBAR
        sidebar = ttk.Frame(root, padding=15, bootstyle="secondary")
        sidebar.pack(side=LEFT, fill=Y)
        ttk.Label(sidebar, text="LISTA", font=("Segoe UI", 9, "bold"), bootstyle="inverse-secondary").pack(anchor=W)
        ttk.Button(sidebar, text="➕ Aggiungi PDF", command=self.add_pdfs, bootstyle="primary", width=22).pack(pady=2)
        ttk.Button(sidebar, text="🗑️ Rimuovi", command=self.remove_pdf, bootstyle="danger-outline", width=22).pack(pady=2)
        ttk.Button(sidebar, text="♻️ Svuota Tutto", command=self.clear_all, bootstyle="warning-outline", width=22).pack(pady=(5, 10))
        ttk.Separator(sidebar).pack(fill=X, pady=5)
        btn_frame = ttk.Frame(sidebar, bootstyle="secondary")
        btn_frame.pack(fill=X)
        ttk.Button(btn_frame, text="⬆ Su", command=self.move_up, bootstyle="light-outline", width=10).pack(side=LEFT, padx=2)
        ttk.Button(btn_frame, text="⬇ Giù", command=self.move_down, bootstyle="light-outline", width=10).pack(side=RIGHT, padx=2)
        ttk.Separator(sidebar).pack(fill=X, pady=10)
        ttk.Label(sidebar, text="STRUMENTI", font=("Segoe UI", 9, "bold"), bootstyle="inverse-secondary").pack(anchor=W)
        ttk.Button(sidebar, text="🔄 Ruota (Dialog)", command=lambda: self.launch_rotate_dialog(0), bootstyle="info-outline", width=22).pack(pady=2)
        ttk.Button(sidebar, text="✂️ Estrai Pagine", command=self.extract_pages_ui, bootstyle="info-outline", width=22).pack(pady=2)
        ttk.Button(sidebar, text="📄 Estrai Testo (.txt)", command=self.extract_text_ui, bootstyle="info-outline", width=22).pack(pady=2)
        
        ttk.Separator(sidebar).pack(fill=X, pady=15)

        # OPZIONI
        ttk.Label(sidebar, text="OPZIONI FINALI", font=("Segoe UI", 9, "bold"), bootstyle="inverse-secondary").pack(anchor=W, pady=(1,5))
        self.pdf_a_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(sidebar, text="Converti a PDF/A-2b", variable=self.pdf_a_var, bootstyle="info-round-toggle").pack(anchor=W)
        
        # --- OPZIONE NORMALIZZA A4 ---
        self.normalize_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(sidebar, text="Ridimensiona tutto (A4)", variable=self.normalize_var, bootstyle="info-round-toggle").pack(anchor=W, pady=5)

        ttk.Label(sidebar, text="Titolo:", bootstyle="inverse-secondary", font=("Segoe UI", 8)).pack(anchor=W, pady=(5,0))
        self.entry_title = ttk.Entry(sidebar, width=22)
        self.entry_title.pack(pady=2)
        ttk.Label(sidebar, text="Autore:", bootstyle="inverse-secondary", font=("Segoe UI", 8)).pack(anchor=W)
        self.entry_author = ttk.Entry(sidebar, width=22)
        self.entry_author.pack(pady=2)
        ttk.Label(sidebar, text="Password (blocca PDF/A):", bootstyle="inverse-secondary", font=("Segoe UI", 8)).pack(anchor=W, pady=(5,0))
        self.entry_pwd = ttk.Entry(sidebar, width=22, show="*") 
        self.entry_pwd.pack(pady=2)
        
        # --- SELETTORE TEMA ---
        ttk.Separator(sidebar).pack(fill=X, pady=10)
        ttk.Label(sidebar, text="🎨 TEMA APPLICAZIONE", font=("Segoe UI", 9, "bold"), bootstyle="inverse-secondary").pack(anchor=W)
        self.current_theme = tk.StringVar(value=self.style.theme.name)
        self.theme_combo = ttk.Combobox(sidebar, textvariable=self.current_theme, values=self.style.theme_names(), state="readonly", width=20)
        self.theme_combo.pack(pady=5)
        self.theme_combo.bind("<<ComboboxSelected>>", self.change_theme)

        # --- PULSANTE DONAZIONI LOGIC LAB ---
        ttk.Separator(sidebar).pack(fill=X, pady=10)
        
        def apri_donazione():
            webbrowser.open("https://ko-fi.com/logiclab")

        btn_donate = ttk.Button(
            sidebar, 
            text="☕ Offri un Caffè a Logic Lab", 
            command=apri_donazione, 
            bootstyle="warning-outline", # Giallo/Arancio elegante
            width=22
        )
        btn_donate.pack(pady=5)

        ttk.Separator(sidebar).pack(fill=X, pady=15)
        self.merge_btn = ttk.Button(sidebar, text="✏ UNISCI E SALVA ", command=self.start_merge_thread, bootstyle="success", width=22)
        self.merge_btn.pack(pady=(5, 0), ipady=8)

        # MAIN AREA
        main_area = ttk.Frame(root, padding=20)
        main_area.pack(side=RIGHT, fill=BOTH, expand=True)
        top_frame = ttk.Frame(main_area)
        top_frame.pack(fill=X)
        ttk.Label(top_frame, text="PDF MASTER TOOL 🖍📝", font=("Segoe UI", 15, "bold"), bootstyle="primary").pack(side=LEFT)
        hint_lbl = ttk.Label(main_area, text="💡 Clicca sull'intestazione per ordinare A-Z | Doppio Click = Apri File", font=("Segoe UI", 9), bootstyle="secondary")
        hint_lbl.pack(anchor=W, pady=(5, 10))

        # --- Progress Bar ---
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_area, variable=self.progress_var, maximum=100, bootstyle="success-striped")
        self.progress_bar.pack(fill=X, pady=(0, 10))

        # --- TABELLA ---
        tree_frame = ttk.Frame(main_area)
        tree_frame.pack(fill=BOTH, expand=True)
        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side=RIGHT, fill=Y)
        
        cols = ("Nome", "Pagine", "Percorso")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", yscrollcommand=tree_scroll.set, selectmode="browse", height=15)
        
        self.tree.heading("Nome", text="Nome File", anchor=W, command=lambda: self.sort_column("Nome", False))
        self.tree.heading("Pagine", text="Pagine", anchor=CENTER, command=lambda: self.sort_column("Pagine", False))
        self.tree.heading("Percorso", text="Percorso Completo", anchor=W, command=lambda: self.sort_column("Percorso", False))
        
        self.tree.column("Nome", width=450, anchor=W)
        self.tree.column("Pagine", width=80, minwidth=50, anchor=CENTER, stretch=False)
        self.tree.column("Percorso", width=250, anchor=W)
        
        self.tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.config(command=self.tree.yview)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)

        self.context_menu = tk.Menu(root, tearoff=0)
        self.context_menu.add_command(label="👁️ Anteprima Zoom", command=self.launch_zoom_preview)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="↪️ Ruota Destra 90°", command=lambda: self.launch_rotate_dialog(90))
        self.context_menu.add_command(label="↩️ Ruota Sinistra 90°", command=lambda: self.launch_rotate_dialog(-90))
        self.context_menu.add_separator()
        self.context_menu.add_command(label="✂️ Estrai Pagine", command=self.extract_pages_ui)
        self.context_menu.add_command(label="📄 Estrai Testo (.txt)", command=self.extract_text_ui)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="🗑️ Rimuovi", command=self.remove_pdf)

        self.tree.bind('<Delete>', lambda e: self.remove_pdf())
        self.tree.bind('<Double-1>', self.open_external_file) 
        self.tree.bind("<Button-3>", self.show_context_menu)
        
        # --- BARRA DI STATO AVANZATA ---
        self.bottom_bar = ttk.Frame(main_area, bootstyle="secondary", padding=5)
        self.bottom_bar.pack(side=BOTTOM, fill=X, pady=(5,0))

        self.status_msg_lbl = ttk.Label(self.bottom_bar, text="Pronto.", bootstyle="inverse-secondary", font=("Segoe UI", 9))
        self.status_msg_lbl.pack(side=LEFT, padx=5)

        badges_frame = ttk.Frame(self.bottom_bar, bootstyle="secondary")
        badges_frame.pack(side=RIGHT)

        # AGGIORNAMENTO BADGES (Senza Thread, è istantaneo ora)
        if self.gs_path:
            # Se trovato localmente è PER FORZA Portable/Bundled
            ttk.Label(badges_frame, text=" ✔ GS Locale Attivo ", bootstyle="success-inverse", font=("Segoe UI", 9, "bold")).pack(side=LEFT, padx=2)
        else:
            ttk.Label(badges_frame, text=" ❌ Ghostscript Non Trovato ", bootstyle="danger-inverse", font=("Segoe UI", 9, "bold")).pack(side=LEFT, padx=2)
            self.status_msg_lbl.config(text="⚠️ Attenzione: PDF/A sarà simulato.")

    # --- UTILITIES ---
    def configure_table_style(self):
        self.style.configure('Treeview', rowheight=40, font=("Segoe UI", 10))
        self.style.configure('Treeview.Heading', font=("Segoe UI", 10, "bold"), relief="flat")
        self.style.map('Treeview.Heading', 
                       relief=[('active', 'groove')], 
                       background=[('active', self.style.colors.secondary)], 
                       foreground=[('active', self.style.colors.selectfg)])

    def change_theme(self, event):
        new_theme = self.theme_combo.get()
        self.style.theme_use(new_theme)
        self.configure_table_style()

    def sort_column(self, col, reverse):
        if not self.pdf_list: return
        if col == "Nome":
            self.pdf_list.sort(key=lambda x: os.path.basename(x).lower(), reverse=reverse)
        elif col == "Pagine":
            def page_key(p):
                c = self.get_page_count(p)
                return int(c) if isinstance(c, int) else 0
            self.pdf_list.sort(key=page_key, reverse=reverse)
        elif col == "Percorso":
            self.pdf_list.sort(key=str.lower, reverse=reverse)

        self.refresh_list()
        arrow = " ▼" if reverse else " ▲"
        clean_cols = {"Nome": "Nome File", "Pagine": "Pagine", "Percorso": "Percorso Completo"}
        self.tree.heading("Nome", text="Nome File")
        self.tree.heading("Pagine", text="Pagine")
        self.tree.heading("Percorso", text="Percorso Completo")
        self.tree.heading(col, text=clean_cols[col] + arrow, command=lambda: self.sort_column(col, not reverse))

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.tk_popup(event.x_root, event.y_root)

    def launch_zoom_preview(self):
        selected = self.tree.selection()
        if not selected: return
        ZoomPreviewWindow(self.root, self.pdf_list, self.tree.index(selected[0]))

    def launch_rotate_dialog(self, start_angle=0):
        selected = self.tree.selection()
        if not selected: return
        idx = self.tree.index(selected[0])
        RotatePreviewDialog(self.root, self.pdf_list[idx], self.update_file_in_list, start_angle)

    def update_file_in_list(self, old_path, new_path):
        try:
            idx = self.pdf_list.index(old_path)
            self.pdf_list[idx] = new_path
            self.refresh_list()
            child_id = self.tree.get_children()[idx]
            self.tree.selection_set(child_id)
        except: pass

    def clear_all(self):
        if self.pdf_list and messagebox.askyesno("Conferma", "Svuotare lista?"):
            self.pdf_list = []
            self.refresh_list()

    def get_page_count(self, filepath):
        try: return len(PdfReader(filepath).pages)
        except: return "?"

    def refresh_list(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for f in self.pdf_list:
            self.tree.insert("", tk.END, values=(f"📝 {os.path.basename(f)}", self.get_page_count(f), f))
        self.status_msg_lbl.config(text=f"Totale: {len(self.pdf_list)} file")

    def add_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if files:
            self.pdf_list.extend(files)
            self.refresh_list()

    def remove_pdf(self):
        sel = self.tree.selection()
        if sel:
            idx = self.tree.index(sel[0])
            self.pdf_list.pop(idx)
            self.refresh_list()

    def move_up(self):
        sel = self.tree.selection()
        if not sel: return
        idx = self.tree.index(sel[0])
        if idx > 0:
            self.pdf_list[idx], self.pdf_list[idx-1] = self.pdf_list[idx-1], self.pdf_list[idx]
            self.refresh_list()
            self.tree.selection_set(self.tree.get_children()[idx-1])

    def move_down(self):
        sel = self.tree.selection()
        if not sel: return
        idx = self.tree.index(sel[0])
        if idx < len(self.pdf_list)-1:
            self.pdf_list[idx], self.pdf_list[idx+1] = self.pdf_list[idx+1], self.pdf_list[idx]
            self.refresh_list()
            self.tree.selection_set(self.tree.get_children()[idx+1])

    def open_external_file(self, event):
        sel = self.tree.selection()
        if sel:
            try: os.startfile(self.pdf_list[self.tree.index(sel[0])])
            except: pass

    def extract_pages_ui(self):
        sel = self.tree.selection()
        if not sel: return
        idx = self.tree.index(sel[0])
        path = self.pdf_list[idx]
        
        istruzioni = "Inserisci i numeri delle pagine da estrarre.\n\n" \
                     "Esempi:\n" \
                     "• 5         (Solo pagina 5)\n" \
                     "• 1,3       (Pagine 1 e 3)\n" \
                     "• 1-5       (Dalla 1 alla 5)\n" \
                     "• 1-3, 8    (Dalla 1 alla 3 + la 8)"
                     
        pgs = simpledialog.askstring("Estrai Pagine", istruzioni, parent=self.root)
        
        if not pgs: return
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=f"Ext_{os.path.basename(path)}")
        if save_path:
            try:
                r, w = PdfReader(path), PdfWriter()
                keep = set()
                for part in pgs.split(','):
                    part = part.strip()
                    if '-' in part:
                        try:
                            s, e = map(int, part.split('-'))
                            keep.update(range(s-1, e))
                        except: pass
                    else: 
                        try: keep.add(int(part)-1)
                        except: pass
                
                if not keep: return
                
                for i in sorted(keep):
                    if 0 <= i < len(r.pages): w.add_page(r.pages[i])
                with open(save_path, "wb") as f: w.write(f)
                
                if messagebox.askyesno("Lista File", "Vuoi sostituire il file originale nella lista con quello estratto?"): 
                    self.update_file_in_list(path, save_path)

                if messagebox.askyesno("Fatto", "Estrazione completata! ✅\nVuoi aprire la cartella di salvataggio?"):
                    try: os.startfile(os.path.dirname(save_path))
                    except: pass

            except Exception as e: messagebox.showerror("Errore", str(e))

    def extract_text_ui(self):
        sel = self.tree.selection()
        if not sel: return
        idx = self.tree.index(sel[0])
        path = self.pdf_list[idx]
        save_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("File di Testo", "*.txt")],
            initialfile=f"Testo_{os.path.basename(path).rsplit('.', 1)[0]}.txt"
        )
        if save_path:
            try:
                doc = fitz.open(path)
                with open(save_path, "w", encoding="utf-8") as f:
                    f.write(f"TESTO ESTRATTO DA: {os.path.basename(path)}\n")
                    f.write("="*50 + "\n\n")
                    for i, page in enumerate(doc):
                        text = page.get_text()
                        f.write(f"--- PAGINA {i+1} ---\n")
                        f.write(text)
                        f.write("\n\n")
                doc.close()
                if messagebox.askyesno("Fatto", "Testo estratto con successo! 📄\nVuoi aprire il file ora?"):
                    os.startfile(save_path)
            except Exception as e:
                messagebox.showerror("Errore", str(e))

    def make_fake_pdfa(self, writer):
        metadata = DictionaryObject()
        metadata.update({NameObject("/Title"): NameObject("PDF"), NameObject("/Creator"): NameObject("Python")})
        xml = b"""<?xpacket begin="" id="W5M0MpCehiHzreSzNTczkc9d"?><x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description rdf:about="" xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/"><pdfaid:part>1</pdfaid:part><pdfaid:conformance>B</pdfaid:conformance></rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end="w"?>"""
        stream = DecodedStreamObject()
        stream.set_data(xml)
        stream.update({NameObject("/Type"): NameObject("/Metadata"), NameObject("/Subtype"): NameObject("/XML")})
        writer._root_object.update({NameObject("/Metadata"): stream})

    # --- THREAD HELPERS ---
    def update_ui(self, progress=None, status=None, cursor=None):
        """Aggiorna la GUI dal thread principale."""
        if progress is not None: self.progress_var.set(progress)
        if status is not None: self.status_msg_lbl.config(text=status)
        if cursor is not None: self.root.config(cursor=cursor)

    def finish_merge(self, success, msg, path):
        """Chiamato alla fine del thread per mostrare i messaggi."""
        self.is_processing = False
        self.merge_btn.config(state=NORMAL)
        self.root.config(cursor="")
        self.progress_var.set(100 if success else 0)
        self.status_msg_lbl.config(text="Operazione Terminata.") 
        
        if success:
            if messagebox.askyesno("Fatto", f"{msg}\nAprire cartella?"):
                try: os.startfile(os.path.dirname(path))
                except: pass
            self.progress_var.set(0)
        else:
            messagebox.showerror("Errore", msg)

    # --- MERGE E CONVERSIONE (THREADED) ---
    def start_merge_thread(self):
        if self.is_processing: return
        if len(self.pdf_list) < 2: 
            # Eccezione: se voglio solo convertire un singolo file in PDF/A o normalizzarlo
            if not (len(self.pdf_list) == 1 and (self.normalize_var.get() or self.pdf_a_var.get())):
                return messagebox.showwarning("Info", "Servono almeno 2 file per unire (o 1 per convertire/ridimensionare).")
        
        out = filedialog.asksaveasfilename(defaultextension=".pdf")
        if not out: return

        pwd = self.entry_pwd.get().strip()
        
        # USA IL PATH GIA TROVATO ALL'AVVIO
        gs_found = self.gs_path
        
        if self.pdf_a_var.get() and pwd and gs_found:
            if not messagebox.askyesno("Attenzione", "PDF/A e Password sono incompatibili.\nLa password sarà ignorata durante la conversione GS.\nProcedere?"):
                return
        
        # Avvia il thread (DAEMON = True per chiusura sicura)
        self.is_processing = True
        self.merge_btn.config(state=DISABLED)
        self.root.config(cursor="watch")
        t = threading.Thread(target=self._worker_merge, args=(out, pwd, gs_found), daemon=True)
        t.start()

    def _worker_merge(self, out, pwd, gs_found):
        temp_out = ""
        try:
            self.root.after(0, lambda: self.update_ui(0, "Inizio elaborazione...", "watch"))
            
            # 1. FASE UNIONE (pypdf)
            merger = PdfWriter()
            
            # Dimensioni A4 Standard (Punti)
            A4_W, A4_H = 595.0, 842.0
            
            total_files = len(self.pdf_list)
            step = 50 / total_files if total_files > 0 else 0 

            title = self.entry_title.get().strip()
            author = self.entry_author.get().strip()

            for idx, f in enumerate(self.pdf_list):
                self.root.after(0, lambda s=f"Elaborazione: {os.path.basename(f)}...": self.update_ui(status=s))
                
                reader = PdfReader(f)
                for page in reader.pages:
                    if self.normalize_var.get():
                        try:
                            curr_w = float(page.mediabox.width)
                            curr_h = float(page.mediabox.height)
                            
                            if curr_w > 0 and curr_h > 0:
                                scale_w = A4_W / curr_w
                                scale_h = A4_H / curr_h
                                scale = min(scale_w, scale_h)
                                page.scale_by(scale)
                        except Exception: pass 

                    merger.add_page(page)
                
                prog = (idx + 1) * step
                self.root.after(0, lambda p=prog: self.update_ui(progress=p))

            # Zoom 100% e Layout
            if self.normalize_var.get() and len(merger.pages) > 0:
                dest = ArrayObject([merger.pages[0].indirect_reference, NameObject("/XYZ"), NullObject(), NullObject(), NumberObject(1)])
                merger._root_object.update({NameObject("/OpenAction"): dest})

            if title or author:
                merger.add_metadata({'/Title': title if title else "PDF", '/Author': author if author else "Utente"})
            
            should_convert = self.pdf_a_var.get() and gs_found

            if pwd and not should_convert:
                merger.encrypt(pwd)
            
            simulated_pdfa = False
            if self.pdf_a_var.get() and not gs_found:
                self.make_fake_pdfa(merger)
                simulated_pdfa = True

            temp_out = out if not should_convert else out.replace(".pdf", "_temp.pdf")
            
            with open(temp_out, "wb") as f_out:
                merger.write(f_out)
            merger.close()

            self.root.after(0, lambda: self.update_ui(60, "File base creato..."))
            
            # 2. FASE CONVERSIONE (Ghostscript)
            final_msg = "Operazione completata! ✅"
            if simulated_pdfa: 
                final_msg += "\n(Attenzione: PDF/A Simulato - GS non trovato)"

            if should_convert:
                self.root.after(0, lambda: self.update_ui(status="Conversione PDF/A (Ghostscript)..."))
                
                success, msg_err = convert_to_pdfa_ghostscript(temp_out, out, gs_found)
                
                if success:
                    try: os.remove(temp_out)
                    except: pass
                    final_msg += "\n(Convertito in PDF/A-2b Reale)"
                else:
                    if os.path.exists(out):
                        try: os.remove(out)
                        except: pass
                    os.rename(temp_out, out)
                    final_msg = f"Attenzione: Unione OK, ma conversione fallita.\n{msg_err}"

            self.root.after(0, lambda: self.finish_merge(True, final_msg, out))

        except Exception as e:
            if temp_out and os.path.exists(temp_out) and temp_out != out:
                try: os.remove(temp_out)
                except: pass
            self.root.after(0, lambda: self.finish_merge(False, str(e), out))

if __name__ == "__main__":
    root = ttk.Window(themename="flatly")
    app = PDFMergerApp(root)
    root.mainloop()