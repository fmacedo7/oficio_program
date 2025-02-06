import tkinter as tk
from tkinter import messagebox, filedialog
import json
from datetime import datetime
from docx import Document
from docx.shared import Pt

def preencher_documento(campos):
    caminho_modelo = "oficio_modelo_2.docx"

    try:
        doc = Document(caminho_modelo)
    except FileNotFoundError:
        messagebox.showerror("Erro", f"Arquivo {caminho_modelo} não encontrado.")
        return

    for p in doc.paragraphs:
        if "[data1]" in p.text:
            p.text = f"Piripiri – PI, {campos['data1']}"
            p.alignment = 2  # Alinhado à direita
        elif "[diretor]" in p.text:
            p.text = f"Ilmo(a). Sr(a) Diretor(a): {campos['diretor']}"
        elif "[escola]" in p.text:
            p.text = f"Escola Municipal: {campos['escola']}"
        elif "[servidor]" in p.text:
            p.text = (f"Encaminhamos, através deste, o(a) servidor(a) {campos['servidor']}, "
                      f"matrícula nº {campos['matricula']}, na situação funcional {campos['situacao']}, "
                      f"para exercer a função de {campos['funcao']} no(s) turno(s) {campos['turno']}. "
                      f"O contrato será válido até {campos['termino']}.")
        elif "[observacao]" in p.text:
            p.text = f"OBSERVAÇÕES: {campos['observacao']}"

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

def salvar_historico(campos):
    try:
        with open("historico.json", "r", encoding="utf-8") as file:
            historico = json.load(file)
    except (FileNotFoundError, json.JSONDecodeError):
        historico = []

    historico.append({**campos, "data_criacao": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})

    with open("historico.json", "w", encoding="utf-8") as file:
        json.dump(historico, file, indent=4, ensure_ascii=False)

def exibir_historico():
    try:
        with open("historico.json", "r", encoding="utf-8") as file:
            historico = json.load(file)
    except (FileNotFoundError, json.JSONDecodeError):
        messagebox.showinfo("Histórico", "Nenhum histórico encontrado.")
        return

    # Criar uma nova janela para exibir o histórico
    janela_historico = tk.Toplevel(root)
    janela_historico.title("Histórico de Documentos")
    janela_historico.geometry("600x400")

    # Mostrar os dados do histórico em formato de texto
    texto_historico = tk.Text(janela_historico, wrap=tk.WORD)
    texto_historico.pack(expand=True, fill=tk.BOTH)

    for idx, item in enumerate(historico, start=1):
        texto_historico.insert(tk.END, f"Documento {idx}\n")
        for chave, valor in item.items():
            texto_historico.insert(tk.END, f"  {chave}: {valor}\n")
        texto_historico.insert(tk.END, "\n")

def emitir_documento():
    def salvar_dados():
        campos = {
            "data1": entry_data.get(),
            "diretor": entry_diretor.get(),
            "escola": entry_escola.get(),
            "servidor": entry_servidor.get(),
            "matricula": entry_matricula.get(),
            "situacao": entry_situacao.get(),
            "funcao": entry_funcao.get(),
            "turno": entry_turno.get(),
            "termino": entry_termino.get(),
            "observacao": entry_observacoes.get()
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

# Tela inicial com dois botões
root = tk.Tk()
root.title("Gerador de Ofícios")
root.geometry("400x300")

btn_historico = tk.Button(root, text="Histórico", command=exibir_historico, width=30, height=2)
btn_historico.pack(pady=20)

btn_emitir = tk.Button(root, text="Emitir Documento", command=emitir_documento, width=30, height=2)
btn_emitir.pack(pady=20)

root.mainloop()