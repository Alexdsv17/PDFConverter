# PDFConverter


import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
from docx import Document
from docx.shared import Inches

# Désactiver certaines fonctionnalités de OpenCV pour éviter l'erreur avec PyInstaller
os.environ["OPENCV_VIDEOIO_PRIORITY_MSMF"] = "0"

def convert_pdf_to_docx():
    # Demander à l'utilisateur de sélectionner un fichier PDF
    pdf_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not pdf_file:
        return  # L'utilisateur a annulé la sélection

    # Demander où enregistrer le fichier DOCX
    docx_file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Files", "*.docx")])
    if not docx_file:
        return  # L'utilisateur a annulé l'enregistrement

    try:
        # Conversion PDF → DOCX
        cv = Converter(pdf_file)
        cv.convert(docx_file)
        cv.close()

        # Demander si l'utilisateur veut ajouter une image
        add_image = messagebox.askyesno("Ajouter une image", "Voulez-vous ajouter une image au document ?")
        if add_image:
            image_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
            if image_path:
                doc = Document(docx_file)
                doc.add_paragraph("Voici une image :")
                doc.add_picture(image_path, width=Inches(4))
                doc.save(docx_file)
        
        messagebox.showinfo("Succès", f"Conversion réussie ! Le fichier est enregistré sous {docx_file}")
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue : {str(e)}")

# Interface Tkinter
root = tk.Tk()
root.title("Convertisseur PDF → DOCX")

convert_button = tk.Button(root, text="Convertir un PDF en DOCX", command=convert_pdf_to_docx)
convert_button.pack(pady=20)

root.mainloop()
