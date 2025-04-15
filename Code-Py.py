import os
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from io import BytesIO
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
import re

def extract_nis_from_pdf(pdf_path):
    """Extrait le numéro NIS d'un PDF"""
    with open(pdf_path, 'rb') as f:
        reader = PdfReader(f)
        text = ""
        for page in reader.pages:
            text += page.extract_text()

        nis = None
        for line in text.split('\n'):
            if "NIS" in line or "INS" in line:
                parts = line.split()
                for part in parts:
                    if part.isdigit() and len(part) >= 8:
                        nis = part
                        break
                if nis:
                    break
        return nis

def add_nis_to_pdf(input_pdf_path, output_pdf_path, nis):
    """Ajoute le numéro NIS à la fin du PDF avec un cadre autour"""
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()

    # Créer un buffer pour le nouveau contenu avec reportlab
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=letter)

    # Coordonnées d'affichage
    x_position = 44
    y_position = 150

    # Ajouter le texte sans spécifier de police personnalisée
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x_position, y_position, "IDENTITE NATIONALE DE SANTE")

    y_position -= 20
    frame_x = x_position - 5
    frame_y = y_position - 5
    frame_width = 250
    frame_height = 20
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)
    c.rect(frame_x, frame_y, frame_width, frame_height)

    c.drawString(x_position, y_position, f"NIS: {nis}")

    c.save()

    packet.seek(0)
    new_pdf = PdfReader(packet)
    last_page = reader.pages[-1]
    last_page.merge_page(new_pdf.pages[0])

    # Ajouter toutes les pages
    for page in reader.pages[:-1]:
        writer.add_page(page)
    writer.add_page(last_page)

    with open(output_pdf_path, 'wb') as f:
        writer.write(f)

    return output_pdf_path

def sanitize_filename(nis):
    """Supprime les caractères non désirés du nom du fichier (si nécessaire)"""
    return re.sub(r'\W+', '', nis)

class PDFModifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ajout du NIS dans PDF")
        self.root.geometry("750x700")  # Taille de la fenêtre plus grande
        self.root.config(bg="#f4f7fc")

        # Titre principal
        self.title_label = tk.Label(root, text="Modification de PDF avec ajout du NIS", font=("Helvetica", 18, "bold"), bg="#f4f7fc", fg="#2e3b4e")
        self.title_label.pack(pady=25)

        # Frame pour la sélection des fichiers
        self.frame_files = tk.Frame(root, bg="#f4f7fc")
        self.frame_files.pack(pady=20, padx=30)

        # Label et boutons pour la sélection de fichiers
        self.label_ins = tk.Label(self.frame_files, text="Sélectionner le fichier contenant le NIS", font=("Helvetica", 12), bg="#f4f7fc")
        self.label_ins.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.button_select_ins = tk.Button(self.frame_files, text="Sélectionner le fichier NIS", command=self.select_ins, relief="flat", bg="#4CAF50", fg="white", font=("Helvetica", 12), height=2, width=30)
        self.button_select_ins.grid(row=0, column=1, padx=20, pady=10)

        self.label_admin = tk.Label(self.frame_files, text="Sélectionner le bulletin d'admission", font=("Helvetica", 12), bg="#f4f7fc")
        self.label_admin.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.button_select_admin = tk.Button(self.frame_files, text="Sélectionner le fichier administratif", command=self.select_admin, relief="flat", bg="#4CAF50", fg="white", font=("Helvetica", 12), height=2, width=30)
        self.button_select_admin.grid(row=1, column=1, padx=20, pady=10)

        self.label_output = tk.Label(self.frame_files, text="Sélectionner le dossier de sortie", font=("Helvetica", 12), bg="#f4f7fc")
        self.label_output.grid(row=2, column=0, padx=10, pady=10, sticky="w")

        self.button_select_output = tk.Button(self.frame_files, text="Sélectionner le dossier de sortie", command=self.select_output, relief="flat", bg="#4CAF50", fg="white", font=("Helvetica", 12), height=2, width=30)
        self.button_select_output.grid(row=2, column=1, padx=20, pady=10)

        # Frame pour les actions
        self.frame_actions = tk.Frame(root, bg="#f4f7fc")
        self.frame_actions.pack(pady=30)

        self.process_button = tk.Button(self.frame_actions, text="Modifier le PDF", command=self.process_files, bg="#007BFF", fg="white", font=("Helvetica", 14), relief="flat", width=25, height=3)
        self.process_button.grid(row=0, column=0)

        # ProgressBar pour le processus
        self.progress = Progressbar(root, orient="horizontal", length=500, mode="determinate")
        self.progress.pack(pady=10)

        # Label pour afficher le message de résultat
        self.result_label = tk.Label(root, text="", fg="green", font=("Helvetica", 12), bg="#f4f7fc")
        self.result_label.pack(pady=20)

        # Initialiser les variables
        self.chemin_nis = None
        self.chemin_admin = None
        self.output_folder = None

    def select_ins(self):
        self.chemin_nis = filedialog.askopenfilename(title="Sélectionner le fichier contenant le NIS", filetypes=[("Fichiers PDF", "*.pdf")])
        if self.chemin_nis:
            self.label_ins.config(text=f"✅ Fichier NIS sélectionné")

    def select_admin(self):
        self.chemin_admin = filedialog.askopenfilename(title="Sélectionner le fichier administratif", filetypes=[("Fichiers PDF", "*.pdf")])
        if self.chemin_admin:
            self.label_admin.config(text=f"✅ Fichier administratif sélectionné")

    def select_output(self):
        self.output_folder = filedialog.askdirectory(title="Sélectionner le dossier de sortie")
        if self.output_folder:
            self.label_output.config(text=f"✅ Dossier de sortie sélectionné")

    def process_files(self):
        if not self.chemin_nis or not self.chemin_admin or not self.output_folder:
            messagebox.showerror("Erreur", "Veuillez sélectionner tous les fichiers et le dossier de sortie.")
            return

        nis = extract_nis_from_pdf(self.chemin_nis)
        if not nis:
            messagebox.showerror("Erreur", "Aucun NIS trouvé dans le fichier.")
            return

        sanitized_nis = sanitize_filename(nis)
        output_pdf_path = os.path.join(self.output_folder, f"{sanitized_nis}.pdf")

        # Démarrer la barre de progression
        self.progress['value'] = 0
        self.root.update_idletasks()

        # Étape 1 : Extraction du NIS
        self.progress['value'] = 50
        self.root.update_idletasks()

        # Étape 2 : Ajout du NIS au PDF
        modified_pdf_path = add_nis_to_pdf(self.chemin_admin, output_pdf_path, nis)

        # Étape 3 : Mise à jour à 100% et fin
        self.progress['value'] = 100
        self.root.update_idletasks()

        self.result_label.config(text=f"✅ PDF modifié avec le NIS {nis} sauvegardé sous :\n{modified_pdf_path}", fg="green")

# Lancer l'application
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFModifierApp(root)
    root.mainloop()
