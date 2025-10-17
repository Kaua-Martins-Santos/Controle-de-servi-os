import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime, timedelta
import csv

# ----------------------
# Banco de dados
# ----------------------
conexao = sqlite3.connect("tickets.db")
cursor = conexao.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS tickets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    data_solicitacao TEXT,
    setor TEXT,
    predio TEXT,
    local TEXT,
    descricao TEXT,
    status TEXT,
    colaborador TEXT,
    solucao TEXT,
    data_conclusao TEXT,
    horas TEXT
)
""")
conexao.commit()

# ----------------------
# Estilo
# ----------------------
def configurar_estilo():
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("TButton", font=("Helvetica", 11), padding=6, background="#8fbcd4", foreground="white")
    style.map("TButton",
              foreground=[('pressed', 'white'), ('active', 'white')],
              background=[('pressed', '#76a3c3'), ('active', '#76a3c3')])
    style.configure("TLabel", font=("Helvetica", 10), background="#f5f5f5")
    style.configure("Treeview", font=("Helvetica", 10), rowheight=25, background="white", fieldbackground="white")
    style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"), background="#d1e0e0")
    style.map("Treeview", background=[('selected', '#b3cde0')])

# ----------------------
# Funções de tickets
# ----------------------
def criar_ticket():
    def salvar_ticket():
        setor_valor = entrada_setor.get()
        predio_valor = entrada_predio.get()
        local_valor = entrada_local.get()
        descricao_valor = texto_descricao.get("1.0", tk.END).strip()
        data_solicitacao = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status_valor = "Aberto"

        if not setor_valor or not predio_valor or not local_valor or not descricao_valor:
            messagebox.showwarning("Aviso", "Preencha todos os campos!")
            return

        cursor.execute("""
            INSERT INTO tickets (data_solicitacao, setor, predio, local, descricao, status)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (data_solicitacao, setor_valor, predio_valor, local_valor, descricao_valor, status_valor))
        conexao.commit()
        messagebox.showinfo("Sucesso", "Ticket criado com sucesso!")
        janela_ticket.destroy()
        atualizar_abas()

    janela_ticket = tk.Toplevel(janela_principal)
    janela_ticket.title("Novo Ticket")
    janela_ticket.geometry("400x350")
    janela_ticket.configure(bg="#f5f5f5")
    janela_ticket.resizable(False, False)

    frame = ttk.Frame(janela_ticket, padding=20)
    frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frame, text="Setor:").grid(row=0, column=0, sticky="w", pady=5)
    entrada_setor = ttk.Entry(frame, width=30)
    entrada_setor.grid(row=0, column=1, pady=5)

    ttk.Label(frame, text="Prédio:").grid(row=1, column=0, sticky="w", pady=5)
    entrada_predio = ttk.Entry(frame, width=30)
    entrada_predio.grid(row=1, column=1, pady=5)

    ttk.Label(frame, text="Local:").grid(row=2, column=0, sticky="w", pady=5)
    entrada_local = ttk.Entry(frame, width=30)
    entrada_local.grid(row=2, column=1, pady=5)

    ttk.Label(frame, text="Descrição:").grid(row=3, column=0, sticky="nw", pady=5)
    texto_descricao = tk.Text(frame, width=30, height=5, font=("Helvetica", 10), bd=1, relief="solid")
    texto_descricao.grid(row=3, column=1, pady=5)

    ttk.Button(frame, text="Salvar", command=salvar_ticket).grid(row=4, column=0, columnspan=2, pady=15, ipadx=50)

def atualizar_abas():
    for arvore in [arvore_todos, arvore_abertos]:
        for item in arvore.get_children():
            arvore.delete(item)

    cursor.execute("SELECT id, data_solicitacao, setor, predio, local, status FROM tickets")
    for row in cursor.fetchall():
        arvore_todos.insert("", tk.END, values=row)
        if row[5] == "Aberto":
            arvore_abertos.insert("", tk.END, values=row)

def finalizar_ticket(id_ticket):
    def salvar_finalizacao():
        colaborador_valor = entrada_colaborador.get()
        solucao_valor = texto_solucao.get("1.0", tk.END).strip()
        horas_valor = entrada_horas.get()
        data_conclusao_valor = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if not colaborador_valor or not solucao_valor or not horas_valor:
            messagebox.showwarning("Aviso", "Preencha todos os campos!")
            return

        cursor.execute("""
            UPDATE tickets SET 
                status='Finalizado',
                colaborador=?,
                solucao=?,
                data_conclusao=?,
                horas=?
            WHERE id=?
        """, (colaborador_valor, solucao_valor, data_conclusao_valor, horas_valor, id_ticket))
        conexao.commit()
        messagebox.showinfo("Sucesso", "Ticket finalizado!")
        janela_finalizar.destroy()
        atualizar_abas()

    cursor.execute("SELECT * FROM tickets WHERE id=?", (id_ticket,))
    ticket = cursor.fetchone()
    if not ticket:
        messagebox.showerror("Erro", "Ticket não encontrado!")
        return

    janela_finalizar = tk.Toplevel(janela_principal)
    janela_finalizar.title(f"Finalizar Ticket #{id_ticket}")
    janela_finalizar.geometry("450x400")
    janela_finalizar.configure(bg="#f5f5f5")

    frame = ttk.Frame(janela_finalizar, padding=15)
    frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frame, text=f"Setor: {ticket[2]}").pack(anchor="w", pady=2)
    ttk.Label(frame, text=f"Prédio: {ticket[3]}").pack(anchor="w", pady=2)
    ttk.Label(frame, text=f"Local: {ticket[4]}").pack(anchor="w", pady=2)
    ttk.Label(frame, text=f"Descrição: {ticket[5]}").pack(anchor="w", pady=2)

    ttk.Label(frame, text="Colaborador que realizou:").pack(anchor="w", pady=5)
    entrada_colaborador = ttk.Entry(frame, width=40)
    entrada_colaborador.pack(anchor="w", pady=2)

    ttk.Label(frame, text="Solução encontrada:").pack(anchor="w", pady=5)
    texto_solucao = tk.Text(frame, width=40, height=5, font=("Helvetica", 10), bd=1, relief="solid")
    texto_solucao.pack(anchor="w", pady=2)

    ttk.Label(frame, text="Horas gastas:").pack(anchor="w", pady=5)
    entrada_horas = ttk.Entry(frame, width=40)
    entrada_horas.pack(anchor="w", pady=2)

    ttk.Button(frame, text="Finalizar Ticket", command=salvar_finalizacao).pack(pady=15, ipadx=50)

# ----------------------
# Busca
# ----------------------
def buscar_ticket():
    termo = entrada_busca.get().strip()
    if not termo:
        atualizar_abas()
        return

    for item in arvore_todos.get_children():
        arvore_todos.delete(item)

    if termo.isdigit():
        cursor.execute("SELECT id, data_solicitacao, setor, predio, local, status FROM tickets WHERE id=?", (int(termo),))
    else:
        termo_like = f"%{termo}%"
        cursor.execute("""
            SELECT id, data_solicitacao, setor, predio, local, status FROM tickets
            WHERE setor LIKE ? OR predio LIKE ? OR local LIKE ? OR descricao LIKE ?
        """, (termo_like, termo_like, termo_like, termo_like))

    for row in cursor.fetchall():
        arvore_todos.insert("", tk.END, values=row)

# ----------------------
# Relatórios
# ----------------------
def gerar_relatorio(periodo):
    hoje = datetime.now()
    if periodo == "semanal":
        inicio = hoje - timedelta(days=7)
    elif periodo == "mensal":
        inicio = hoje - timedelta(days=30)
    elif periodo == "anual":
        inicio = hoje - timedelta(days=365)
    else:
        return

    # Tickets abertos no período
    cursor.execute("SELECT id, data_solicitacao, setor, predio, local, descricao, status FROM tickets WHERE data_solicitacao>=?", (inicio.strftime("%Y-%m-%d %H:%M:%S"),))
    abertos = cursor.fetchall()

    # Tickets finalizados no período
    cursor.execute("SELECT id, data_solicitacao, setor, predio, local, descricao, status, colaborador, solucao, data_conclusao, horas FROM tickets WHERE data_conclusao>=?", (inicio.strftime("%Y-%m-%d %H:%M:%S"),))
    finalizados = cursor.fetchall()

    # Limpa tabelas
    for t in [arvore_abertos_rel, arvore_finalizados_rel]:
        for item in t.get_children():
            t.delete(item)

    # Preenche tabelas
    for row in abertos:
        arvore_abertos_rel.insert("", tk.END, values=row)
    for row in finalizados:
        arvore_finalizados_rel.insert("", tk.END, values=row)

def exportar_relatorio(treeview, nome_arquivo):
    linhas = []
    colunas = [treeview.heading(c)['text'] for c in treeview["columns"]]
    linhas.append(colunas)
    for item in treeview.get_children():
        linhas.append(treeview.item(item)['values'])

    caminho = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=nome_arquivo, filetypes=[("CSV", "*.csv")])
    if caminho:
        with open(caminho, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(linhas)
        messagebox.showinfo("Sucesso", f"Relatório salvo em {caminho}")

# ----------------------
# Interface Principal
# ----------------------
janela_principal = tk.Tk()
janela_principal.title("Gerenciador de Tickets")
janela_principal.geometry("950x600")
janela_principal.configure(bg="#f5f5f5")
janela_principal.resizable(False, False)

configurar_estilo()

frame_principal = ttk.Frame(janela_principal, padding=10)
frame_principal.pack(fill=tk.BOTH, expand=True)

ttk.Button(frame_principal, text="Novo Ticket", command=criar_ticket).pack(side=tk.TOP, pady=10, ipadx=50)

# Caixa de busca
frame_busca = ttk.Frame(frame_principal)
frame_busca.pack(fill=tk.X, pady=5)
entrada_busca = ttk.Entry(frame_busca, width=40)
entrada_busca.pack(side=tk.LEFT, padx=5)
ttk.Button(frame_busca, text="Buscar", command=buscar_ticket).pack(side=tk.LEFT, padx=5)
ttk.Button(frame_busca, text="Limpar", command=atualizar_abas).pack(side=tk.LEFT, padx=5)

# Notebook (abas)
notebook = ttk.Notebook(frame_principal)
notebook.pack(fill=tk.BOTH, expand=True)

# Função auxiliar para criar Treeview com scrollbars
def criar_treeview(frame, colunas):
    container = ttk.Frame(frame)
    container.pack(fill=tk.BOTH, expand=True)
    tree = ttk.Treeview(container, columns=colunas, show="headings")
    for c in colunas:
        tree.heading(c, text=c)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scroll_vert = ttk.Scrollbar(container, orient=tk.VERTICAL, command=tree.yview)
    scroll_vert.pack(side=tk.RIGHT, fill=tk.Y)
    tree.configure(yscrollcommand=scroll_vert.set)

    scroll_hor = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
    scroll_hor.pack(fill=tk.X)
    tree.configure(xscrollcommand=scroll_hor.set)

    return tree

# Aba Todos
aba_todos = ttk.Frame(notebook)
notebook.add(aba_todos, text="Todos os Tickets")
arvore_todos = criar_treeview(aba_todos, ["ID","Data","Setor","Prédio","Local","Status"])

# Aba Finalizar Tickets
aba_abertos = ttk.Frame(notebook)
notebook.add(aba_abertos, text="Finalizar Tickets")
arvore_abertos = criar_treeview(aba_abertos, ["ID","Data","Setor","Prédio","Local","Status"])

def clicar_abertos(event):
    selecionado = arvore_abertos.selection()
    if selecionado:
        item = arvore_abertos.item(selecionado)
        id_ticket = item["values"][0]
        finalizar_ticket(id_ticket)

arvore_abertos.bind("<Double-1>", clicar_abertos)

# Aba Relatórios
aba_relatorios = ttk.Frame(notebook)
notebook.add(aba_relatorios, text="Relatórios")

frame_rel = ttk.Frame(aba_relatorios, padding=10)
frame_rel.pack(fill=tk.BOTH, expand=True)

# Botões período
ttk.Label(frame_rel, text="Gerar relatório por período:").pack(anchor="w", pady=5)
frame_botoes_rel = ttk.Frame(frame_rel)
frame_botoes_rel.pack(anchor="w", pady=5)
ttk.Button(frame_botoes_rel, text="Semanal", command=lambda: gerar_relatorio("semanal")).pack(side=tk.LEFT, padx=5)
ttk.Button(frame_botoes_rel, text="Mensal", command=lambda: gerar_relatorio("mensal")).pack(side=tk.LEFT, padx=5)
ttk.Button(frame_botoes_rel, text="Anual", command=lambda: gerar_relatorio("anual")).pack(side=tk.LEFT, padx=5)

# Treeviews Relatórios
ttk.Label(frame_rel, text="Tickets Abertos:").pack(anchor="w", pady=5)
arvore_abertos_rel = criar_treeview(frame_rel, ["ID","Data","Setor","Prédio","Local","Descrição","Status"])
ttk.Button(frame_rel, text="Exportar Abertos", command=lambda: exportar_relatorio(arvore_abertos_rel, "relatorio_abertos")).pack(pady=5)

ttk.Label(frame_rel, text="Tickets Finalizados:").pack(anchor="w", pady=5)
arvore_finalizados_rel = criar_treeview(frame_rel, ["ID","Data","Setor","Prédio","Local","Descrição","Status","Colaborador","Solução","Data Conclusão","Horas"])
ttk.Button(frame_rel, text="Exportar Finalizados", command=lambda: exportar_relatorio(arvore_finalizados_rel, "relatorio_finalizados")).pack(pady=5)

# Inicializa abas
atualizar_abas()

janela_principal.mainloop()
