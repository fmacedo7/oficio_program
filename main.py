import tkinter as tk
from tkinter import messagebox, filedialog
import json
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Função para preencher o documento com os dados fornecidos
def preencher_documento(campos):
    caminho_modelo = "oficio_modelo_2.docx"

    try:
        doc = Document(caminho_modelo)
    except FileNotFoundError:
        messagebox.showerror("Erro", f"Arquivo {caminho_modelo} não encontrado.")
        return

    # Substituir os campos delimitados por colchetes
    for p in doc.paragraphs:
        for key, value in campos.items():
            if key in p.text:
                p.text = p.text.replace(f"[{key}]", value)
            for run in p.runs:
                run.font.name = "Arial"
                run.font.size = Pt(11)

    # Salvar o documento preenchido
    caminho_arquivo = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Documentos Word", "*.docx")],
        initialfile="Documento_Preenchido.docx"
    )
    if not caminho_arquivo:
        return

    doc.save(caminho_arquivo)
    salvar_historico(campos)
    messagebox.showinfo("Documento Gerado", f"Ofício salvo em {caminho_arquivo}")


# Função para salvar o histórico
def salvar_historico(campos):
    try:
        with open("historico.json", "r") as file:
            historico = json.load(file)
    except (FileNotFoundError, json.JSONDecodeError):
        historico = []

    historico.append({**campos, "data_criacao": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})

    with open("historico.json", "w") as file:
        json.dump(historico, file, indent=4)


# Função para exibir a tela de preenchimento de dados
def emitir_documento():
    def salvar_dados():
        campos = {
            "data": entry_data.get(),
            "nome_diretor": entry_diretor.get(),
            "escola": entry_escola.get(),
            "nome_servidor": entry_servidor.get(),
            "matricula": entry_matricula.get(),
            "situacao_funcional": entry_situacao.get(),
            "funcao": entry_funcao.get(),
            "turnos": entry_turno.get(),
            "termino_contrato": entry_termino.get(),
            "observacoes": entry_observacoes.get()
        }
        if all([campos[key] for key in campos if key != "observacoes"]):
            preencher_documento(campos)
            limpar_campos()
        else:
            messagebox.showerror("Erro", "Por favor, preencha todos os campos obrigatórios, exceto Observações.")

    def limpar_campos():
        for entry in entries:
            entry.delete(0, tk.END)

    # Tela de preenchimento de dados
    tela_preenchimento = tk.Toplevel(root)
    tela_preenchimento.title("Preencher Dados")
    tela_preenchimento.geometry("600x500")

    labels = [
        "Data", "Nome do Diretor(a)", "Escola", "Nome do Servidor(a)",
        "Matrícula", "Situação Funcional", "Função", "Turnos",
        "Término do Contrato", "Observações"
    ]
    entries = []

    for i, label in enumerate(labels):
        tk.Label(tela_preenchimento, text=label).grid(row=i, column=0, padx=5, pady=5)
        entry = tk.Entry(tela_preenchimento, width=50)
        entry.grid(row=i, column=1, padx=5, pady=5)
        entries.append(entry)

    (entry_data, entry_diretor, entry_escola, entry_servidor, entry_matricula,
     entry_situacao, entry_funcao, entry_turno, entry_termino, entry_observacoes) = entries

    btn_salvar = tk.Button(tela_preenchimento, text="Salvar e Gerar Documento", command=salvar_dados)
    btn_salvar.grid(row=len(labels), column=0, columnspan=2, pady=10)


# Função para exibir a tela de histórico (sem funcionalidade ainda)
def exibir_historico():
    messagebox.showinfo("Histórico", "Funcionalidade em desenvolvimento.")


# Tela inicial com dois botões
root = tk.Tk()
root.title("Gerador de Ofícios")
root.geometry("400x300")

btn_historico = tk.Button(root, text="Histórico", command=exibir_historico, width=30, height=2)
btn_historico.pack(pady=20)

btn_emitir = tk.Button(root, text="Emitir Documento", command=emitir_documento, width=30, height=2)
btn_emitir.pack(pady=20)

root.mainloop()