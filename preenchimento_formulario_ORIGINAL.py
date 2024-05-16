import tkinter as tk
from tkinter import ttk, messagebox
from docx import Document
from docx.shared import Pt

def limpar_campos():
    entry_nome.delete(0, tk.END)
    entry_valor.delete(0, tk.END)
    entry_objeto.delete(0, tk.END)
    entry_nome_documento.delete(0, tk.END)
    entry_nacionalidade.delete(0, tk.END)
    entry_estado_civil.delete(0, tk.END)
    entry_profissao.delete(0, tk.END)
    entry_rg.delete(0, tk.END)
    entry_orgao_emissor_rg.delete(0, tk.END)
    entry_cpf.delete(0, tk.END)
    entry_endereco.delete(0, tk.END)
    entry_numero_endereco.delete(0, tk.END)
    entry_complemento_endereco.delete(0, tk.END)
    entry_bairro.delete(0, tk.END)
    entry_cidade.delete(0, tk.END)
    entry_estado.delete(0, tk.END)
    entry_cep.delete(0, tk.END)
    entry_nome.focus_set()

def preencher_documento():
    nome = entry_nome.get()
    valor = entry_valor.get()
    objeto = entry_objeto.get() # campo que preciso ter a opção em negrito
    nacionalidade = entry_nacionalidade.get()
    estado_civil = entry_estado_civil.get()
    profissao = entry_profissao.get()
    rg = entry_rg.get()
    orgao_emissor_rg = entry_orgao_emissor_rg.get()
    cpf = entry_cpf.get()
    endereco = entry_endereco.get()
    numero_endereco = entry_numero_endereco.get()
    complemento_endereco = entry_complemento_endereco.get()
    bairro = entry_bairro.get()
    cidade = entry_cidade.get()
    estado = entry_estado.get()
    cep = entry_cep.get()

    confirmacao = messagebox.askyesno("Confirmação", f"Nome: {nome}\nValor: {valor}\nObjeto: {objeto}\n\nAs informações estão corretas?")

    if confirmacao:
        document = Document('contrato.docx')
        for paragraph in document.paragraphs:
            if '{nome}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{nome}', nome)
            if '{valor}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{valor}', valor)
            if '{objeto}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{objeto}', objeto)
            if '{nacionalidade}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{nacionalidade}', nacionalidade)
            if '{estado_civil}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{estado_civil}', estado_civil)
            if '{profissao}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{profissao}', profissao)
            if '{rg}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{rg}', rg)
            if '{orgao_emissor_rg}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{orgao_emissor_rg}', orgao_emissor_rg)
            if '{cpf}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{cpf}', cpf)
            if '{endereco}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{endereco}', endereco)
            if '{numero_endereco}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{numero_endereco}', numero_endereco)
            if '{complemento_endereco}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{complemento_endereco}', complemento_endereco)
            if '{bairro}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{bairro}', bairro)
            if '{cidade}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{cidade}', cidade)
            if '{estado}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{estado}', estado)
            if '{cep}' in paragraph.text:
                paragraph.text = paragraph.text.replace('{cep}', cep)

            
            paragraph.style.font.name = 'Arial Narrow'
            paragraph.style.font.size = Pt(12)

        nome_documento = entry_nome_documento.get()
        document.save(f'{nome_documento}.docx')
        messagebox.showinfo("Sucesso", "Documento criado com sucesso!")
        limpar_campos()

root = tk.Tk()
root.title("Preencher Contrato")
root.geometry("400x800") 

frame_principal = ttk.Frame(root)
frame_principal.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

info_frame = ttk.LabelFrame(frame_principal, text="Informações")
info_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

label_nome = tk.Label(info_frame, text="NOME:", font=("Helvetica", 10, "bold"))
label_nome.grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_nome = tk.Entry(info_frame)
entry_nome.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

label_valor = tk.Label(info_frame, text="VALOR:", font=("Helvetica", 10, "bold"))
label_valor.grid(row=1, column=0, padx=5, pady=5, sticky="w")
entry_valor = tk.Entry(info_frame)
entry_valor.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

label_objeto = tk.Label(info_frame, text="OBJETO:", font=("Helvetica", 10, "bold"))
label_objeto.grid(row=2, column=0, padx=5, pady=5, sticky="w")
entry_objeto = tk.Entry(info_frame)
entry_objeto.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

label_nome_documento = tk.Label(info_frame, text="NOME DO DOCUMENTO:", font=("Helvetica", 10, "bold"))
label_nome_documento.grid(row=3, column=0, padx=5, pady=5, sticky="w")
entry_nome_documento = tk.Entry(info_frame)
entry_nome_documento.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

label_nacionalidade = tk.Label(info_frame, text="NACIONALIDADE:", font=("Helvetica", 10, "bold"))
label_nacionalidade.grid(row=4, column=0, padx=5, pady=5, sticky="w")
entry_nacionalidade = tk.Entry(info_frame)
entry_nacionalidade.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

label_estado_civil = tk.Label(info_frame, text="ESTADO CIVIL:", font=("Helvetica", 10, "bold"))
label_estado_civil.grid(row=5, column=0, padx=5, pady=5, sticky="w")
entry_estado_civil = tk.Entry(info_frame)
entry_estado_civil.grid(row=5, column=1, padx=5, pady=5, sticky="ew")

label_profissao = tk.Label(info_frame, text="PROFISSÃO:", font=("Helvetica", 10, "bold"))
label_profissao.grid(row=6, column=0, padx=5, pady=5, sticky="w")
entry_profissao = tk.Entry(info_frame)
entry_profissao.grid(row=6, column=1, padx=5, pady=5, sticky="ew")

label_rg = tk.Label(info_frame, text="RG:", font=("Helvetica", 10, "bold"))
label_rg.grid(row=7, column=0, padx=5, pady=5, sticky="w")
entry_rg = tk.Entry(info_frame)
entry_rg.grid(row=7, column=1, padx=5, pady=5, sticky="ew")

label_orgao_emissor_rg = tk.Label(info_frame, text="ÓRGÃO EMISSOR RG:", font=("Helvetica", 10, "bold"))
label_orgao_emissor_rg.grid(row=8, column=0, padx=5, pady=5, sticky="w")
entry_orgao_emissor_rg = tk.Entry(info_frame)
entry_orgao_emissor_rg.grid(row=8, column=1, padx=5, pady=5, sticky="ew")

label_cpf = tk.Label(info_frame, text="CPF:", font=("Helvetica", 10, "bold"))
label_cpf.grid(row=9, column=0, padx=5, pady=5, sticky="w")
entry_cpf = tk.Entry(info_frame)
entry_cpf.grid(row=9, column=1, padx=5, pady=5, sticky="ew")

label_endereco = tk.Label(info_frame, text="ENDEREÇO:", font=("Helvetica", 10, "bold"))
label_endereco.grid(row=10, column=0, padx=5, pady=5, sticky="w")
entry_endereco = tk.Entry(info_frame)
entry_endereco.grid(row=10, column=1, padx=5, pady=5, sticky="ew")

label_numero_endereco = tk.Label(info_frame, text="NÚMERO ENDEREÇO:", font=("Helvetica", 10, "bold"))
label_numero_endereco.grid(row=11, column=0, padx=5, pady=5, sticky="w")
entry_numero_endereco = tk.Entry(info_frame)
entry_numero_endereco.grid(row=11, column=1, padx=5, pady=5, sticky="ew")

label_complemento_endereco = tk.Label(info_frame, text="COMPLEMENTO ENDEREÇO:", font=("Helvetica", 10, "bold"))
label_complemento_endereco.grid(row=12, column=0, padx=5, pady=5, sticky="w")
entry_complemento_endereco = tk.Entry(info_frame)
entry_complemento_endereco.grid(row=12, column=1, padx=5, pady=5, sticky="ew")

label_bairro = tk.Label(info_frame, text="BAIRRO:", font=("Helvetica", 10, "bold"))
label_bairro.grid(row=13, column=0, padx=5, pady=5, sticky="w")
entry_bairro = tk.Entry(info_frame)
entry_bairro.grid(row=13, column=1, padx=5, pady=5, sticky="ew")

label_cidade = tk.Label(info_frame, text="CIDADE:", font=("Helvetica", 10, "bold"))
label_cidade.grid(row=14, column=0, padx=5, pady=5, sticky="w")
entry_cidade = tk.Entry(info_frame)
entry_cidade.grid(row=14, column=1, padx=5, pady=5, sticky="ew")

label_estado = tk.Label(info_frame, text="ESTADO:", font=("Helvetica", 10, "bold"))
label_estado.grid(row=15, column=0, padx=5, pady=5, sticky="w")
entry_estado = tk.Entry(info_frame)
entry_estado.grid(row=15, column=1, padx=5, pady=5, sticky="ew")

label_cep = tk.Label(info_frame, text="CEP:", font=("Helvetica", 10, "bold"))
label_cep.grid(row=16, column=0, padx=5, pady=5, sticky="w")
entry_cep = tk.Entry(info_frame)
entry_cep.grid(row=16, column=1, padx=5, pady=5, sticky="ew")

ttk.Button(frame_principal, text="Gerar documento", command=preencher_documento).pack(pady=10)

root.mainloop()
