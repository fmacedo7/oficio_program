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

