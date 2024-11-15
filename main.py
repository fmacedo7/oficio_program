import tkinter as tk
from tkinter import messagebox
import csv
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

def generate_pdf(director, school, date, observations, ofice_number):
    pdf_filename = f"Oficio_{ofice_number}.pdf"
    c = canvas.Canvas(pdf_filename, pagesize=A4)
    width, height = A4

    c.drawString(100, height - 100, f"Nome do Diretor(a): {director}")
    c.drawString(100, height - 120, f"Nome da Escola: {school}")
    c.drawString(100, height - 140, f"Data: {date}")
    c.drawString(100, height - 160, f"Observações: {observations}")
    
    c.save()
    messagebox.showinfo("PDF Gerado", f"Documento salvo como {pdf_filename}")

def save_history(director, school, date, observations):
    with open("historico_oficios.csv", mode="a", newline="") as file:
        writer = csv.writer(file)
        writer = writerow([director, school, date, observations, datetime.now().strftime("%d/%m/%Y, %H:%M:%S")])

def export_ofice():
    director = entry_director.get()
    school = entry_school.get()
    date = entry_date.get()
    observations = entry_observations.get()
    ofice_number = datetime.now().strftime("%Y%m%d%H%M%S")

    if director and school and date:
        generate_pdf(director, school, date, observations, ofice_number)
        save_history(director, school, date, observations)
    else:
        messagebox.showerror("Erro", "Por favor, preencha todos os campos obrigatórios.")

root = tk.Tk()
root.title("Gerador de Oficios")

tk.Label(root, text="Nome do diretor(a)").grid(row=0, column=0)
entry_director = tk.Entry(root, width=50)
entry_director.grid(row=0, column=1)

tk.Label(root, text="Nome da escola").grid(row=1, column=0)
entry_school = tk.Entry(root, width=50)
entry_school.grid(row=1, column=1)

tk.Label(root, text="Data").grid(row=2, column=0)
entry_date = tk.Entry(root, width=50)
entry_date.grid(row=2, column=1)

tk.Label(root, text="Observações").grid(row=3, column=0)
entry_observations = tk.Entry(root, width=50)
entry_observations.grid(row=3, column=1)

export_ofice_button = tk.Button(root, text="Emitir Ofício", command=export_ofice)
export_ofice_button.grid(row=4, column=0, columnspan=2, pady=10)

root.mainloop()