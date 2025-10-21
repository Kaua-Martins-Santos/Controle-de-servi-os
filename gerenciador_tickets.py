import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import csv
import re
import win32print

# Importando as classes dos outros arquivos para organizar melhor o projeto
from database import BancoDeDados
from ticket_manager import GerenciadorTickets

class App:
    def __init__(self, root):
        # Configurações iniciais da janela principal e do banco de dados
        self.root = root
        self.db = BancoDeDados()
        self.gerenciador = GerenciadorTickets(self.db)
        self.theme_name = "dark" # Tema padrão é o escuro

        # Dicionário com as cores para os temas claro e escuro
        self.themes = {
            "light": {"BG_COLOR": "#f0f2f5", "PRIMARY_COLOR": "#008080", "SECONDARY_COLOR": "#006666", "TEXT_COLOR": "#333333", "HEADER_BG_COLOR": "#4a4a4a", "WHITE_COLOR": "#ffffff", "SELECTED_COLOR": "#4682B4", "FIELD_BG_COLOR": "#ffffff"},
            "dark": {"BG_COLOR": "#202124", "PRIMARY_COLOR": "#4682B4", "SECONDARY_COLOR": "#5a9bd3", "TEXT_COLOR": "#e8eaed", "HEADER_BG_COLOR": "#3c4043", "WHITE_COLOR": "#e8eaed", "SELECTED_COLOR": "#4682B4", "FIELD_BG_COLOR": "#303030"}
        }
        self.style = ttk.Style()
        self.style.theme_use('clam') # Usamos o tema 'clam' como base que é mais moderno

        # Funções para iniciar a aplicação
        self.criar_widgets()
        self.apply_theme()
        self.atualizar_abas()

    # Aplica o tema de cores escolhido (claro ou escuro) a todos os componentes
    def apply_theme(self):
        colors = self.themes[self.theme_name]
        LABEL_BOLD_FONT = ("Calibri", 11, "bold")
        self.root.configure(bg=colors["BG_COLOR"])

        # Configurações de estilo para os diferentes widgets (botões, tabelas, etc.)
        self.style.configure(".", background=colors["BG_COLOR"], foreground=colors["TEXT_COLOR"], font=("Calibri", 11), bordercolor=colors["FIELD_BG_COLOR"])
        self.style.configure("TButton", background=colors["PRIMARY_COLOR"], foreground=colors["WHITE_COLOR"], font=("Calibri", 11, "bold"), padding=6, borderwidth=0)
        self.style.map("TButton", background=[('active', colors["SECONDARY_COLOR"])])
        self.style.configure("TLabel", background=colors["BG_COLOR"], foreground=colors["TEXT_COLOR"])
        self.style.configure("Bold.TLabel", font=LABEL_BOLD_FONT, background=colors["BG_COLOR"])
        self.style.configure("Status.TLabel", font=("Calibri", 10), background=colors["BG_COLOR"])
        self.style.configure("TFrame", background=colors["BG_COLOR"])
        self.style.configure("TNotebook", background=colors["BG_COLOR"], borderwidth=0)
        self.style.configure("TNotebook.Tab", background=colors["HEADER_BG_COLOR"], padding=[10, 5], font=("Calibri", 11, "bold"), foreground=colors["TEXT_COLOR"])
        self.style.map("TNotebook.Tab", background=[("selected", colors["SELECTED_COLOR"])], foreground=[("selected", colors["WHITE_COLOR"])])
        self.style.configure("Treeview", background=colors["FIELD_BG_COLOR"], fieldbackground=colors["FIELD_BG_COLOR"], foreground=colors["TEXT_COLOR"], rowheight=28, font=("Calibri", 11))
        self.style.configure("Treeview.Heading", background=colors["HEADER_BG_COLOR"], foreground=colors["WHITE_COLOR"], font=("Calibri", 11, "bold"), padding=8)
        self.style.map("Treeview", background=[('selected', colors["SELECTED_COLOR"])], foreground=[('selected', colors["BG_COLOR"])])
        self.style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])
        self.style.configure("TEntry", fieldbackground=colors["FIELD_BG_COLOR"], foreground=colors["TEXT_COLOR"], insertcolor=colors["TEXT_COLOR"])
        self.style.map('TCombobox', fieldbackground=[('readonly', colors["FIELD_BG_COLOR"])], selectbackground=[('readonly', colors["SELECTED_COLOR"])], foreground=[('readonly', colors["TEXT_COLOR"])], bordercolor=[('readonly', colors["HEADER_BG_COLOR"])], lightcolor=[('readonly', colors["HEADER_BG_COLOR"])], darkcolor=[('readonly', colors["HEADER_BG_COLOR"])], arrowcolor=[('readonly', colors["TEXT_COLOR"])])

    # Alterna entre o tema claro e escuro
    def toggle_theme(self):
        self.theme_name = "light" if self.theme_name == "dark" else "dark"
        self.theme_button.config(text="🌙" if self.theme_name == "light" else "☀️")
        self.apply_theme()

    # Cria todos os elementos visuais da tela principal
    def criar_widgets(self):
        self.root.title("Sistema de Controle de Serviços - UNASP")
        self.root.geometry("1100x700")

        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ***** INÍCIO DA ALTERAÇÃO *****
        # Frame superior com os botões de ação e busca
        self.frame_topo = ttk.Frame(main_frame, padding=10)
        self.frame_topo.pack(fill=tk.X)
        ttk.Button(self.frame_topo, text="Novo Ticket", command=self.abrir_janela_ticket).pack(side=tk.LEFT, padx=5)
        # Adicionamos o botão de Reimprimir
        ttk.Button(self.frame_topo, text="Reimprimir Ticket", command=self.reimprimir_ticket_selecionado).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.frame_topo, text="Excluir Ticket", command=self.excluir_ticket_selecionado).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.frame_topo, text="Relatório", command=self.abrir_janela_relatorio).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.frame_topo, text="Gerenciar", command=self.abrir_janela_gerenciamento).pack(side=tk.LEFT, padx=5)
        self.entrada_busca = ttk.Entry(self.frame_topo, width=30, font=("Calibri", 11))
        self.entrada_busca.pack(side=tk.LEFT, padx=(20, 5), ipady=4)
        ttk.Button(self.frame_topo, text="Buscar", command=self.buscar_ticket).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.frame_topo, text="Limpar", command=self.limpar_busca).pack(side=tk.LEFT, padx=5)
        self.theme_button = ttk.Button(self.frame_topo, text="☀️", command=self.toggle_theme, width=4)
        self.theme_button.pack(side=tk.RIGHT, padx=5)
        # ***** FIM DA ALTERAÇÃO *****

        # Notebook para organizar os tickets em abas
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=10, padx=10)

        # Barra de status no rodapé da janela
        self.status_bar = ttk.Label(self.root, text="", anchor=tk.W, style="Status.TLabel", padding=(5, 2))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Definição das colunas para as tabelas de tickets
        colunas_gerais = ("ID", "Data Solicitação", "Setor", "Prédio", "Local", "Status", "Prioridade")
        self.arvore_todos = self.criar_aba(self.notebook, "Todos os Tickets", colunas_gerais)
        self.arvore_abertos = self.criar_aba(self.notebook, "Tickets Abertos", colunas_gerais)
        colunas_fechados = ("ID", "Setor", "Local", "Colaborador", "Início Serviço", "Término Serviço", "Horas")
        self.arvore_fechados = self.criar_aba(self.notebook, "Tickets Fechados", colunas_fechados, [80, 150, 180, 150, 150, 150, 100])

        # Adiciona um evento de duplo clique para abrir detalhes do ticket
        self.arvore_abertos.bind("<Double-1>", self.ao_clicar_aberto)
        self.arvore_todos.bind("<Double-1>", self.ao_clicar_todos)
        self.arvore_fechados.bind("<Double-1>", self.ao_clicar_todos)

    # Exibe uma mensagem na barra de status por alguns segundos
    def mostrar_status(self, mensagem):
        self.status_bar.config(text=mensagem)
        self.root.after(3000, self.limpar_status)

    # Limpa a mensagem da barra de status
    def limpar_status(self):
        self.status_bar.config(text="")

    # Função genérica para criar uma nova aba com uma tabela (Treeview)
    def criar_aba(self, parent, texto_aba, colunas, larguras=None):
        frame = ttk.Frame(parent)
        if isinstance(parent, ttk.Notebook):
            parent.add(frame, text=texto_aba)
        else:
            frame.pack(fill=tk.BOTH, expand=True)

        arvore = ttk.Treeview(frame, columns=colunas, show="headings", selectmode="browse")
        if not larguras:
            larguras = [140] * len(colunas)
            larguras[0] = 80
            larguras[1] = 160
        for i, col in enumerate(colunas):
            arvore.heading(col, text=col)
            arvore.column(col, width=larguras[i], anchor='center')
        arvore.pack(fill=tk.BOTH, expand=True)
        return arvore

    # Formata a data vinda do banco (YYYY-MM-DD) para o formato brasileiro (DD/MM/YYYY)
    def formatar_data_para_exibicao(self, data_db):
        if not data_db: return ""
        try:
            data_obj = datetime.strptime(data_db.split()[0], "%Y-%m-%d")
        except (ValueError, TypeError, IndexError):
            return data_db
        return data_obj.strftime("%d/%m/%Y")

    # Atualiza as informações exibidas nas abas, buscando os dados mais recentes do banco
    def atualizar_abas(self, data=None):
        # Limpa as tabelas antes de preenchê-las novamente
        for arvore in [self.arvore_todos, self.arvore_abertos, self.arvore_fechados]:
            for item in arvore.get_children():
                arvore.delete(item)

        # Se não for uma busca, pega todos os tickets
        tickets_gerais = data if data is not None else self.gerenciador.get_todos_tickets()
        if tickets_gerais:
            for row in tickets_gerais:
                row = list(row)
                row[1] = self.formatar_data_para_exibicao(row[1]) # Formata a data
                self.arvore_todos.insert("", tk.END, values=tuple(row))
                if row[5] == "Aberto": # Se o status for 'Aberto', adiciona na aba de abertos
                    self.arvore_abertos.insert("", tk.END, values=tuple(row))

        # Pega os tickets finalizados para a aba de fechados
        tickets_fechados = self.gerenciador.get_fechados_tickets()
        if tickets_fechados:
            for row in tickets_fechados:
                row = list(row)
                row[4] = self.formatar_data_para_exibicao(row[4]) # Formata data de início
                row[5] = self.formatar_data_para_exibicao(row[5]) # Formata data de término
                self.arvore_fechados.insert("", tk.END, values=tuple(row))

    # Abre a janela para criar um novo ticket ou editar um existente
    def abrir_janela_ticket(self, ticket_id=None):
        janela = tk.Toplevel(self.root)
        janela.title("Novo Ticket" if ticket_id is None else f"Editando Ticket #{ticket_id}")
        janela.geometry("450x520")
        janela.resizable(False, False)
        janela.transient(self.root)
        janela.grab_set()
        colors = self.themes[self.theme_name]
        janela.configure(bg=colors["BG_COLOR"])
        frame = ttk.Frame(janela, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # Se um ID for passado, busca os dados do ticket para preencher os campos
        ticket = self.gerenciador.get_ticket_por_id(ticket_id) if ticket_id else None

        # --- Campos do formulário ---
        ttk.Label(frame, text="ID:").grid(row=0, column=0, sticky="w", pady=5)
        entrada_id = ttk.Entry(frame, width=33)
        entrada_id.grid(row=0, column=1, pady=5)
        if ticket:
            entrada_id.insert(0, ticket[0])
            entrada_id.config(state='readonly') # Trava o campo ID na edição

        ttk.Label(frame, text="Data:").grid(row=1, column=0, sticky="w", pady=5)
        entrada_data = ttk.Entry(frame, width=20)
        entrada_data.grid(row=1, column=1, sticky="w", pady=5)
        if ticket:
            entrada_data.insert(0, self.formatar_data_para_exibicao(ticket[1]))
        # Botão "Hoje" para preencher a data atual
        ttk.Button(frame, text="Hoje", width=6, command=lambda: (entrada_data.delete(0, tk.END), entrada_data.insert(0, datetime.now().strftime("%d/%m/%Y")))).grid(row=1, column=1, sticky="e", padx=5)

        ttk.Label(frame, text="Setor:").grid(row=2, column=0, sticky="w", pady=5)
        combo_setor = ttk.Combobox(frame, values=["Eletrica", "Hidraulica", "Construção", "Pintura", "Refrigeração", "Serralheria"], state="readonly", width=30)
        combo_setor.grid(row=2, column=1, pady=5)
        if ticket:
            combo_setor.set(ticket[2])
        else:
            combo_setor.current(0)

        ttk.Label(frame, text="Prédio:").grid(row=3, column=0, sticky="w", pady=5)
        entrada_predio = ttk.Entry(frame, width=33)
        entrada_predio.grid(row=3, column=1, pady=5)
        lista_predios_completa = [row[1] for row in self.gerenciador.get_predios() or []]
        lista_sugestoes = tk.Listbox(janela, height=4, bg=colors["FIELD_BG_COLOR"], fg=colors["TEXT_COLOR"], highlightbackground=colors["PRIMARY_COLOR"])

        # Lógica para autocompletar o nome do prédio enquanto o usuário digita
        def on_predio_keyup(event):
            if event.keysym in ("Up", "Down", "Return", "Escape"): return
            texto = entrada_predio.get().lower()
            if not texto: lista_sugestoes.place_forget(); return
            sugestoes = [p for p in lista_predios_completa if texto in p.lower()]
            lista_sugestoes.delete(0, tk.END)
            for s in sugestoes: lista_sugestoes.insert(tk.END, s)
            if sugestoes:
                x = frame.winfo_x() + entrada_predio.winfo_x()
                y = frame.winfo_y() + entrada_predio.winfo_y() + entrada_predio.winfo_height()
                lista_sugestoes.place(x=x, y=y); lista_sugestoes.config(width=0, height=len(sugestoes) if len(sugestoes) <=4 else 4); lista_sugestoes.tkraise(); lista_sugestoes.selection_set(0)
            else: lista_sugestoes.place_forget()

        def select_suggestion(event):
            if lista_sugestoes.winfo_viewable() and lista_sugestoes.curselection():
                valor = lista_sugestoes.get(lista_sugestoes.curselection()); entrada_predio.delete(0, tk.END); entrada_predio.insert(0, valor); lista_sugestoes.place_forget(); entrada_local.focus(); return "break"

        def navigate_suggestions(event):
            if not lista_sugestoes.winfo_viewable(): return
            current_selection = lista_sugestoes.curselection(); idx = current_selection[0] if current_selection else -1
            if event.keysym == "Down": idx = 0 if idx == lista_sugestoes.size() - 1 else idx + 1
            elif event.keysym == "Up": idx = lista_sugestoes.size() - 1 if idx <= 0 else idx - 1
            if idx != -1: lista_sugestoes.selection_clear(0, tk.END); lista_sugestoes.selection_set(idx); lista_sugestoes.see(idx)
            return "break"

        entrada_predio.bind("<KeyRelease>", on_predio_keyup); entrada_predio.bind("<Return>", select_suggestion); entrada_predio.bind("<Up>", navigate_suggestions); entrada_predio.bind("<Down>", navigate_suggestions); lista_sugestoes.bind("<Button-1>", select_suggestion); lista_sugestoes.bind("<Return>", select_suggestion)
        if ticket: entrada_predio.insert(0, ticket[3])


        ttk.Label(frame, text="Local:").grid(row=4, column=0, sticky="w", pady=5)
        entrada_local = ttk.Entry(frame, width=33)
        entrada_local.grid(row=4, column=1, pady=5)
        if ticket: entrada_local.insert(0, ticket[4])

        ttk.Label(frame, text="Prioridade:").grid(row=5, column=0, sticky="w", pady=5)
        combo_prioridade = ttk.Combobox(frame, values=["Baixa", "Média", "Alta"], state="readonly", width=30)
        combo_prioridade.grid(row=5, column=1, pady=5)
        if ticket and ticket[7]: combo_prioridade.set(ticket[7])
        else: combo_prioridade.set("Média")

        ttk.Label(frame, text="Descrição:").grid(row=6, column=0, sticky="nw", pady=5)
        texto_descricao = tk.Text(frame, width=35, height=8, font=("Calibri", 11), bg=colors["FIELD_BG_COLOR"], fg=colors["TEXT_COLOR"], insertbackground=colors["TEXT_COLOR"], relief="solid", borderwidth=1)
        texto_descricao.grid(row=6, column=1, pady=5)
        if ticket: texto_descricao.insert("1.0", ticket[5])

        # Função auxiliar para limpar os campos após criar um ticket
        def limpar_campos_para_novo_ticket():
            entrada_id.delete(0, tk.END)
            entrada_data.delete(0, tk.END)
            #combo_setor.current(0) # Mantem o setor selecionado
            entrada_predio.delete(0, tk.END)
            entrada_local.delete(0, tk.END)
            #combo_prioridade.set("Média") # Mantem a prioridade
            texto_descricao.delete("1.0", tk.END)
            # Coloca o foco no campo ID para o próximo ticket
            entrada_id.focus()

        # Função para salvar o ticket (cria um novo ou atualiza um existente)
        def salvar():
            # Pega os valores dos campos
            id_val_str = entrada_id.get().strip(); data_str = entrada_data.get(); setor_val = combo_setor.get(); predio_val = entrada_predio.get(); local_val = entrada_local.get(); descricao_val = texto_descricao.get("1.0", tk.END).strip(); prioridade_val = combo_prioridade.get()

            # Validações para garantir que os campos foram preenchidos corretamente
            if not id_val_str: messagebox.showwarning("Aviso", "O campo 'ID' é obrigatório.", parent=janela); return
            try: id_val = int(id_val_str)
            except ValueError: messagebox.showwarning("Aviso", "O ID deve ser um número.", parent=janela); return
            if not all([data_str, setor_val, predio_val, local_val, descricao_val, prioridade_val]): messagebox.showwarning("Aviso", "Preencha todos os campos!", parent=janela); return
            try: datetime.strptime(data_str, "%d/%m/%Y")
            except ValueError: messagebox.showwarning("Aviso", "Data inválida! Use o formato DD/MM/YYYY.", parent=janela); return

            if ticket_id is None: # Se não tem ID, é um novo ticket
                if self.gerenciador.get_ticket_por_id(id_val): messagebox.showerror("Erro", f"O ID #{id_val} já existe no sistema.", parent=janela); return
                self.gerenciador.criar_ticket(id_val, data_str, setor_val, predio_val, local_val, descricao_val, prioridade_val)
                self.mostrar_status(f"Ticket #{id_val} criado com sucesso!")

                # Prepara as informações para impressão
                ticket_info_impressao = {
                    "id": id_val, "data_solicitacao": data_str, "setor": setor_val,
                    "predio": predio_val, "local": local_val, "descricao": descricao_val
                }
                self.imprimir_ticket(ticket_info_impressao) # Chama a função de impressão

                # Em vez de fechar a janela, limpa os campos para o próximo ticket
                limpar_campos_para_novo_ticket()

            else: # Se tem ID, está editando
                self.gerenciador.atualizar_ticket(ticket_id, data_str, setor_val, predio_val, local_val, descricao_val, prioridade_val)
                self.mostrar_status(f"Ticket #{ticket_id} atualizado com sucesso!")

                # Fecha a janela SOMENTE quando está editando
                janela.destroy()

            self.atualizar_abas() # Atualiza a lista de tickets na tela principal

        ttk.Button(frame, text="Salvar", command=salvar).grid(row=7, column=0, columnspan=2, pady=20, ipadx=50)

    # Função para formatar e enviar o ticket para a impressora térmica
    def imprimir_ticket(self, ticket_info):
        try:
            printer_name = win32print.GetDefaultPrinter() # Pega a impressora padrão do Windows
            h_printer = win32print.OpenPrinter(printer_name)
            try:
                # --- Comandos da impressora (ESC/POS) ---
                ESC = b'\x1b'
                GS = b'\x1d'
                LF = b'\x0a' # Pular linha

                CMD_INIT = ESC + b'@' # Reseta a impressora
                CMD_ALIGN_CENTER = ESC + b'a\x01' # Alinhar no centro
                CMD_ALIGN_LEFT = ESC + b'a\x00' # Alinhar à esquerda
                CMD_FONT_BOLD_ON = ESC + b'E\x01' # Negrito
                CMD_FONT_BOLD_OFF = ESC + b'E\x00'
                CMD_DOUBLE_HW_ON = GS + b'!\x11'  # Letra grande
                CMD_DOUBLE_HW_OFF = GS + b'!\x00'
                CMD_CUT = GS + b'V\x01' # Cortar papel

                # --- Montando o conteúdo do ticket ---
                ticket_data = CMD_INIT
                ticket_data += CMD_ALIGN_CENTER

                # Título
                ticket_data += CMD_DOUBLE_HW_ON
                ticket_data += CMD_FONT_BOLD_ON
                ticket_data += b'ORDEM DE SERVICO\n'
                ticket_data += CMD_DOUBLE_HW_OFF
                ticket_data += b'UNASP-EC\n'
                ticket_data += CMD_FONT_BOLD_OFF
                ticket_data += b'========================================\n'

                # ID e Data
                id_str = str(ticket_info["id"])
                data_str = ticket_info["data_solicitacao"]
                cabecalho = f'{id_str} {data_str.rjust(40 - len(id_str))}\n'
                ticket_data += cabecalho.encode("cp850") # cp850 ajuda com acentos

                ticket_data += b'========================================\n'
                ticket_data += LF

                ticket_data += CMD_ALIGN_LEFT

                # Corpo do ticket
                ticket_data += CMD_FONT_BOLD_ON
                ticket_data += b'PREDIO: '
                ticket_data += CMD_FONT_BOLD_OFF
                ticket_data += f'{ticket_info["predio"]}\n'.encode('cp850', errors='replace')

                ticket_data += CMD_FONT_BOLD_ON
                ticket_data += b'LOCAL: '
                ticket_data += CMD_FONT_BOLD_OFF
                ticket_data += f'{ticket_info["local"]}\n'.encode('cp850', errors='replace')
                ticket_data += LF

                ticket_data += CMD_FONT_BOLD_ON
                ticket_data += f'SERVICO SOLICITADO: ({ticket_info["data_solicitacao"]})\n'.encode('cp850', errors='replace')
                ticket_data += CMD_FONT_BOLD_OFF
                ticket_data += f'{ticket_info["descricao"]}\n'.encode('cp850', errors='replace')
                ticket_data += LF

                ticket_data += b'----------------------------------------\n'
                ticket_data += CMD_FONT_BOLD_ON
                ticket_data += b'SETOR SOLICITADO: '
                ticket_data += CMD_FONT_BOLD_OFF
                ticket_data += f'{ticket_info["setor"]}\n'.encode('cp850', errors='replace')

                # Espaços para preenchimento manual
                ticket_data += CMD_FONT_BOLD_ON
                ticket_data += b'SOLUCAO:\n'
                ticket_data += CMD_FONT_BOLD_OFF
                ticket_data += LF * 4

                ticket_data += b'----------------------------------------\n'
                ticket_data += CMD_FONT_BOLD_ON
                ticket_data += b'COLABORADOR RESPONSAVEL:\n'
                ticket_data += CMD_FONT_BOLD_OFF
                ticket_data += LF

                ticket_data += b'----------------------------------------\n'
                ticket_data += b'HORAS GASTAS: __:__\n'
                ticket_data += b'INICIO DO SERVICO: __/__/____\n'
                ticket_data += b'TERMINO DO SERVICO: __/__/____\n'
                ticket_data += LF

                ticket_data += b'========================================\n'
                ticket_data += CMD_FONT_BOLD_ON
                ticket_data += b'RECEBIMENTO DO SERVICO\n'
                ticket_data += CMD_FONT_BOLD_OFF
                ticket_data += LF
                ticket_data += b'NOME:\n\n'
                ticket_data += b'ASS.:___________________________________\n'
                ticket_data += LF * 4

                ticket_data += CMD_CUT # Corta o papel no final

                # Envia os dados para a impressora
                h_job = win32print.StartDocPrinter(h_printer, 1, ("Ticket de Servico", None, "RAW"))
                try:
                    win32print.WritePrinter(h_printer, ticket_data)
                finally:
                    win32print.EndDocPrinter(h_printer)
            finally:
                win32print.ClosePrinter(h_printer)

            self.mostrar_status(f"Ticket #{ticket_info['id']} enviado para a impressora.")

        except Exception as e:
            messagebox.showerror("Erro de Impressão", f"Não foi possível imprimir o ticket: {e}")

    # Abre a janela para finalizar um ticket ou ver os detalhes de um já finalizado
    def abrir_janela_finalizar(self, ticket_id):
        ticket = self.gerenciador.get_ticket_por_id(ticket_id)
        if not ticket: messagebox.showerror("Erro", "Ticket não encontrado!"); return
        is_finalizado = (ticket[6] == 'Finalizado')

        janela = tk.Toplevel(self.root); janela.title(f"Detalhes do Ticket #{ticket_id}"); janela.resizable(False, False); janela.transient(self.root); janela.grab_set()
        colors = self.themes[self.theme_name]; janela.configure(bg=colors["BG_COLOR"])
        frame = ttk.Frame(janela, padding=15); frame.pack(fill=tk.BOTH, expand=True)

        if is_finalizado: # Se já está finalizado, apenas mostra os detalhes
            janela.geometry("480x450")
            ttk.Label(frame, text="Descrição do Problema:", style="Bold.TLabel").pack(anchor="w", pady=(5, 2))
            desc_text = tk.Text(frame, height=4, width=50, font=("Calibri", 11), relief="flat")
            desc_text.insert("1.0", ticket[5] or "Não informado"); desc_text.config(state="disabled", bg=colors["BG_COLOR"], fg=colors["TEXT_COLOR"]); desc_text.pack(anchor="w", fill="x", pady=(0, 5))

            ttk.Label(frame, text="Solução Aplicada:", style="Bold.TLabel").pack(anchor="w", pady=(5, 2))
            solucao_text = tk.Text(frame, height=5, width=50, font=("Calibri", 11), relief="flat")
            solucao_text.insert("1.0", ticket[9] or "Não informada"); solucao_text.config(state="disabled", bg=colors["BG_COLOR"], fg=colors["TEXT_COLOR"]); solucao_text.pack(anchor="w", fill="x", pady=(0, 10))

            info_frame = ttk.Frame(frame); info_frame.pack(fill="x", anchor="w")
            ttk.Label(info_frame, text="Colaborador:", style="Bold.TLabel").grid(row=0, column=0, sticky="w", pady=2); ttk.Label(info_frame, text=ticket[8] or "Não informado").grid(row=0, column=1, sticky="w", padx=10)
            ttk.Label(info_frame, text="Início do Serviço:", style="Bold.TLabel").grid(row=1, column=0, sticky="w", pady=2); ttk.Label(info_frame, text=self.formatar_data_para_exibicao(ticket[12])).grid(row=1, column=1, sticky="w", padx=10)
            ttk.Label(info_frame, text="Término do Serviço:", style="Bold.TLabel").grid(row=2, column=0, sticky="w", pady=2); ttk.Label(info_frame, text=self.formatar_data_para_exibicao(ticket[13])).grid(row=2, column=1, sticky="w", padx=10)
            ttk.Label(info_frame, text="Horas Gastas:", style="Bold.TLabel").grid(row=3, column=0, sticky="w", pady=2); ttk.Label(info_frame, text=ticket[11] or "Não informado").grid(row=3, column=1, sticky="w", padx=10)

            ttk.Button(frame, text="Fechar", command=janela.destroy).pack(side="bottom", pady=(20, 0))
        else: # Se está aberto, mostra os campos para finalizar
            janela.geometry("450x520")
            ttk.Label(frame, text=f"Descrição: {ticket[5]}", wraplength=400).pack(anchor="w", pady=2)
            dates_frame = ttk.Frame(frame); dates_frame.pack(fill=tk.X, pady=(15,0))
            ttk.Label(dates_frame, text="Data de Início:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
            entrada_data_inicio = ttk.Entry(dates_frame, width=15); entrada_data_inicio.grid(row=0, column=1, pady=2)
            ttk.Button(dates_frame, text="Hoje", width=6, command=lambda: (entrada_data_inicio.delete(0, tk.END), entrada_data_inicio.insert(0, datetime.now().strftime("%d/%m/%Y")))).grid(row=0, column=2, padx=5)
            ttk.Label(dates_frame, text="Data de Término:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
            entrada_data_termino = ttk.Entry(dates_frame, width=15); entrada_data_termino.grid(row=1, column=1, pady=2)
            ttk.Button(dates_frame, text="Hoje", width=6, command=lambda: (entrada_data_termino.delete(0, tk.END), entrada_data_termino.insert(0, datetime.now().strftime("%d/%m/%Y")))).grid(row=1, column=2, padx=5)
            ttk.Label(frame, text="Colaborador:").pack(anchor="w", pady=(15, 0))
            lista_nomes = [row[0] for row in self.gerenciador.get_colaboradores_por_setor(ticket[2]) or []]
            combo_colaborador = ttk.Combobox(frame, values=lista_nomes, state="readonly", width=38); combo_colaborador.pack(anchor="w")
            ttk.Label(frame, text="Solução:").pack(anchor="w", pady=(10, 0)); texto_solucao = tk.Text(frame, width=40, height=5, font=("Calibri", 11), bg=colors["FIELD_BG_COLOR"], fg=colors["TEXT_COLOR"], insertbackground=colors["TEXT_COLOR"], relief="solid", borderwidth=1); texto_solucao.pack(anchor="w")

            horas_frame = ttk.Frame(frame); horas_frame.pack(anchor="w", pady=(10,0))
            ttk.Label(horas_frame, text="Horas gastas:").pack(side=tk.LEFT)
            entrada_horas = ttk.Entry(horas_frame, width=15); entrada_horas.pack(side=tk.LEFT, padx=5)

            # Valida o formato das horas (HH:MM)
            def validar_horas(valor): return re.fullmatch(r'\d{1,2}:\d{2}', valor) is not None

            # Salva os dados de finalização
            def salvar():
                data_inicio_str = entrada_data_inicio.get(); data_termino_str = entrada_data_termino.get(); horas_str = entrada_horas.get().strip()
                if not data_inicio_str or not data_termino_str: messagebox.showwarning("Aviso", "As datas de início e término são obrigatórias.", parent=janela); return
                try: data_inicio_obj = datetime.strptime(data_inicio_str, "%d/%m/%Y"); data_termino_obj = datetime.strptime(data_termino_str, "%d/%m/%Y")
                except ValueError: messagebox.showwarning("Aviso", "Formato de data inválido! Use DD/MM/YYYY.", parent=janela); return
                if data_termino_obj < data_inicio_obj: messagebox.showwarning("Aviso", "A data de término não pode ser anterior à data de início.", parent=janela); return
                if not validar_horas(horas_str): messagebox.showwarning("Aviso", "Formato de horas inválido! Use HH:MM.", parent=janela); return
                outros_campos = [combo_colaborador.get(), texto_solucao.get("1.0", tk.END).strip(), horas_str]
                if not all(outros_campos): messagebox.showwarning("Aviso", "Preencha todos os campos de finalização!", parent=janela); return
                self.gerenciador.finalizar_ticket(ticket_id, *outros_campos, data_inicio_str, data_termino_str); self.mostrar_status(f"Ticket #{ticket_id} finalizado com sucesso!"); janela.destroy(); self.atualizar_abas()
            ttk.Button(frame, text="Finalizar Ticket", command=salvar).pack(pady=20, ipadx=50)

    # Abre a janela de gerenciamento de Colaboradores e Prédios
    def abrir_janela_gerenciamento(self):
        janela = tk.Toplevel(self.root); janela.title("Gerenciamento"); janela.geometry("650x450"); janela.resizable(False, False); janela.transient(self.root); janela.grab_set()
        colors = self.themes[self.theme_name]; janela.configure(bg=colors["BG_COLOR"])
        notebook = ttk.Notebook(janela); notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Aba de Colaboradores ---
        frame_colab = ttk.Frame(notebook); notebook.add(frame_colab, text="Colaboradores")
        frame_lista_colab = ttk.Frame(frame_colab, padding=10); frame_lista_colab.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        colunas_colab = ("ID", "Nome", "Setor"); arvore_colab = self.criar_aba(frame_lista_colab, "", colunas_colab, [40, 200, 150])
        frame_form_colab = ttk.Frame(frame_colab, padding=20); frame_form_colab.pack(side=tk.RIGHT, fill=tk.Y)
        ttk.Label(frame_form_colab, text="Nome:").pack(anchor="w"); entrada_nome_colab = ttk.Entry(frame_form_colab, width=30); entrada_nome_colab.pack(anchor="w", pady=(0, 10))
        ttk.Label(frame_form_colab, text="Setor:").pack(anchor="w"); combo_setor_colab = ttk.Combobox(frame_form_colab, values=["Eletrica", "Hidraulica", "Construção", "Pintura", "Refrigeração", "Serralheria"], state="readonly", width=28); combo_setor_colab.pack(anchor="w")
        def atualizar_lista_colab():
            for i in arvore_colab.get_children(): arvore_colab.delete(i)
            for row in self.gerenciador.get_colaboradores() or []: arvore_colab.insert("", "end", values=row)
        def limpar_campos_colab(): entrada_nome_colab.delete(0, tk.END); combo_setor_colab.set(''); arvore_colab.selection_remove(arvore_colab.selection())
        def adicionar_colab():
            nome=entrada_nome_colab.get().strip(); setor=combo_setor_colab.get()
            if not nome or not setor: messagebox.showwarning("Aviso", "Preencha todos os campos.", parent=janela); return
            self.gerenciador.adicionar_colaborador(nome, setor); atualizar_lista_colab(); limpar_campos_colab()
        def editar_colab():
            selecionado = arvore_colab.focus()
            if not selecionado: messagebox.showwarning("Aviso", "Selecione um colaborador para editar.", parent=janela); return
            id_colab=arvore_colab.item(selecionado)['values'][0]; nome=entrada_nome_colab.get().strip(); setor=combo_setor_colab.get()
            if not nome or not setor: messagebox.showwarning("Aviso", "Preencha todos os campos.", parent=janela); return
            self.gerenciador.editar_colaborador(id_colab, nome, setor); atualizar_lista_colab(); limpar_campos_colab()
        def remover_colab():
            selecionado = arvore_colab.focus()
            if not selecionado: messagebox.showwarning("Aviso", "Selecione um colaborador para remover.", parent=janela); return
            id_colab = arvore_colab.item(selecionado)['values'][0]
            if messagebox.askyesno("Confirmar", "Tem certeza que deseja remover o colaborador?", parent=janela): self.gerenciador.remover_colaborador(id_colab); atualizar_lista_colab(); limpar_campos_colab()
        def ao_selecionar_colab(event):
            selecionado = arvore_colab.focus()
            if not selecionado: return
            dados=arvore_colab.item(selecionado)['values']; entrada_nome_colab.delete(0,tk.END); entrada_nome_colab.insert(0,dados[1]); combo_setor_colab.set(dados[2])
        arvore_colab.bind("<<TreeviewSelect>>", ao_selecionar_colab)
        frame_botoes_colab = ttk.Frame(frame_form_colab); frame_botoes_colab.pack(pady=20)
        ttk.Button(frame_botoes_colab, text="Adicionar", command=adicionar_colab).pack(fill=tk.X, pady=2); ttk.Button(frame_botoes_colab, text="Editar", command=editar_colab).pack(fill=tk.X, pady=2); ttk.Button(frame_botoes_colab, text="Remover", command=remover_colab).pack(fill=tk.X, pady=2)
        ttk.Button(frame_form_colab, text="Limpar Campos", command=limpar_campos_colab).pack(pady=10)
        atualizar_lista_colab()

        # --- Aba de Prédios ---
        frame_predio = ttk.Frame(notebook); notebook.add(frame_predio, text="Prédios")
        frame_lista_predio = ttk.Frame(frame_predio, padding=10); frame_lista_predio.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        colunas_predio = ("ID", "Nome"); arvore_predio = self.criar_aba(frame_lista_predio, "", colunas_predio, [40, 250])
        frame_form_predio = ttk.Frame(frame_predio, padding=20); frame_form_predio.pack(side=tk.RIGHT, fill=tk.Y)
        ttk.Label(frame_form_predio, text="Nome do Prédio:").pack(anchor="w"); entrada_nome_predio = ttk.Entry(frame_form_predio, width=30); entrada_nome_predio.pack(anchor="w")
        def atualizar_lista_predio():
            for i in arvore_predio.get_children(): arvore_predio.delete(i)
            for row in self.gerenciador.get_predios() or []: arvore_predio.insert("", "end", values=row)
        def limpar_campos_predio(): entrada_nome_predio.delete(0, tk.END); arvore_predio.selection_remove(arvore_predio.selection())
        def adicionar_predio():
            nome = entrada_nome_predio.get().strip()
            if not nome: messagebox.showwarning("Aviso", "O nome não pode ser vazio.", parent=janela); return
            if any(nome.lower() == p[1].lower() for p in self.gerenciador.get_predios() or []): messagebox.showerror("Erro", "Este nome de prédio já existe.", parent=janela); return
            if self.gerenciador.adicionar_predio(nome): atualizar_lista_predio(); limpar_campos_predio()
        def editar_predio():
            selecionado = arvore_predio.focus()
            if not selecionado: messagebox.showwarning("Aviso", "Selecione um prédio para editar.", parent=janela); return
            id_predio = arvore_predio.item(selecionado)['values'][0]; nome = entrada_nome_predio.get().strip()
            if not nome: messagebox.showwarning("Aviso", "O nome não pode ser vazio.", parent=janela); return
            self.gerenciador.editar_predio(id_predio, nome); atualizar_lista_predio(); limpar_campos_predio()
        def remover_predio():
            selecionado = arvore_predio.focus()
            if not selecionado: messagebox.showwarning("Aviso", "Selecione um prédio para remover.", parent=janela); return
            id_predio = arvore_predio.item(selecionado)['values'][0]
            if messagebox.askyesno("Confirmar", "Tem certeza que deseja remover o prédio?", parent=janela): self.gerenciador.remover_predio(id_predio); atualizar_lista_predio(); limpar_campos_predio()
        def ao_selecionar_predio(event):
            selecionado = arvore_predio.focus()
            if not selecionado: return
            dados = arvore_predio.item(selecionado)['values']; entrada_nome_predio.delete(0, tk.END); entrada_nome_predio.insert(0, dados[1])
        arvore_predio.bind("<<TreeviewSelect>>", ao_selecionar_predio)
        frame_botoes_predio = ttk.Frame(frame_form_predio); frame_botoes_predio.pack(pady=20)
        ttk.Button(frame_botoes_predio, text="Adicionar", command=adicionar_predio).pack(fill=tk.X, pady=2); ttk.Button(frame_botoes_predio, text="Editar", command=editar_predio).pack(fill=tk.X, pady=2); ttk.Button(frame_botoes_predio, text="Remover", command=remover_predio).pack(fill=tk.X, pady=2)
        ttk.Button(frame_form_predio, text="Limpar Campo", command=limpar_campos_predio).pack(pady=10)
        atualizar_lista_predio()

    # Abre a janela para gerar relatórios
    def abrir_janela_relatorio(self):
        janela = tk.Toplevel(self.root); janela.title("Relatório de Tickets"); janela.geometry("1000x600"); janela.transient(self.root); janela.grab_set()
        colors = self.themes[self.theme_name]; janela.configure(bg=colors["BG_COLOR"])

        # Frame para os filtros
        frame_filtros = ttk.Frame(janela, padding=10); frame_filtros.pack(fill=tk.X)

        # Linha 1 de filtros
        ttk.Label(frame_filtros, text="Período:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        combo_periodo = ttk.Combobox(frame_filtros, values=["Semanal", "Mensal", "Anual"], state="readonly", width=15)
        combo_periodo.grid(row=0, column=1, sticky="w", padx=5, pady=5); combo_periodo.current(0)

        ttk.Label(frame_filtros, text="Status:").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        combo_status = ttk.Combobox(frame_filtros, values=["Todos", "Aberto", "Finalizado"], state="readonly", width=15)
        combo_status.grid(row=0, column=3, sticky="w", padx=5, pady=5); combo_status.set("Todos")

        ttk.Label(frame_filtros, text="Setor:").grid(row=0, column=4, sticky="w", padx=5, pady=5)
        combo_setor = ttk.Combobox(frame_filtros, values=["Todos", "Eletrica", "Hidraulica", "Construção", "Pintura", "Refrigeração", "Serralheria"], state="readonly", width=15)
        combo_setor.grid(row=0, column=5, sticky="w", padx=5, pady=5); combo_setor.set("Todos")

        # Linha 2 de filtros
        predios = ["Todos"] + [p[1] for p in self.gerenciador.get_predios()]
        ttk.Label(frame_filtros, text="Prédio:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        combo_predio = ttk.Combobox(frame_filtros, values=predios, state="readonly", width=15)
        combo_predio.grid(row=1, column=1, sticky="w", padx=5, pady=5); combo_predio.set("Todos")

        colaboradores = ["Todos"] + [c[1] for c in self.gerenciador.get_colaboradores()]
        ttk.Label(frame_filtros, text="Colaborador:").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        combo_colaborador = ttk.Combobox(frame_filtros, values=colaboradores, state="readonly", width=15)
        combo_colaborador.grid(row=1, column=3, sticky="w", padx=5, pady=5); combo_colaborador.set("Todos")

        ttk.Label(frame_filtros, text="Prioridade:").grid(row=1, column=4, sticky="w", padx=5, pady=5)
        combo_prioridade = ttk.Combobox(frame_filtros, values=["Todas", "Baixa", "Média", "Alta"], state="readonly", width=15)
        combo_prioridade.grid(row=1, column=5, sticky="w", padx=5, pady=5); combo_prioridade.set("Todas")

        # Linha 3 de filtros
        ttk.Label(frame_filtros, text="Local (busca):").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        entry_local = ttk.Entry(frame_filtros, width=18)
        entry_local.grid(row=2, column=1, sticky="w", padx=5, pady=5)

        # Frame para os botões
        frame_botoes = ttk.Frame(janela, padding=10); frame_botoes.pack(fill=tk.X)
        ttk.Button(frame_botoes, text="Gerar Relatório", command=lambda: gerar()).pack(side=tk.LEFT, padx=10)
        ttk.Button(frame_botoes, text="Exportar para CSV", command=lambda: exportar_csv()).pack(side=tk.LEFT, padx=10)

        # Frame para a tabela de resultados
        frame_tabela = ttk.Frame(janela, padding=10); frame_tabela.pack(fill=tk.BOTH, expand=True)
        colunas_relatorio = ("ID", "Data", "Setor", "Prédio", "Local", "Status", "Prioridade", "Colaborador")
        arvore_relatorio = ttk.Treeview(frame_tabela, columns=colunas_relatorio, show="headings", selectmode="browse")
        for col in colunas_relatorio:
            arvore_relatorio.heading(col, text=col)
            arvore_relatorio.column(col, width=115, anchor='center')
        arvore_relatorio.pack(fill=tk.BOTH, expand=True)

        def gerar():
            # Limpa o relatório anterior
            for item in arvore_relatorio.get_children(): arvore_relatorio.delete(item)

            # Cálculo do período de tempo
            tipo = combo_periodo.get(); hoje = datetime.now()
            if tipo == "Semanal":
                inicio = hoje - timedelta(days=hoje.weekday())
                fim = inicio + timedelta(days=4)
            elif tipo == "Mensal":
                inicio = hoje.replace(day=1)
                proximo_mes = (hoje.replace(day=28) + timedelta(days=4))
                fim = proximo_mes - timedelta(days=proximo_mes.day)
            else: # Anual
                inicio = hoje - timedelta(days=365)
                fim = hoje

            # Coleta dos filtros
            filtros = {
                "predio": combo_predio.get() if combo_predio.get() != "Todos" else None,
                "colaborador": combo_colaborador.get() if combo_colaborador.get() != "Todos" else None,
                "status": combo_status.get() if combo_status.get() != "Todos" else None,
                "setor": combo_setor.get() if combo_setor.get() != "Todos" else None,
                "local": entry_local.get().strip() or None,
                "prioridade": combo_prioridade.get() if combo_prioridade.get() != "Todas" else None
            }

            # Busca os registros no banco de dados com as datas e filtros
            registros = self.gerenciador.get_tickets_relatorio(
                inicio.strftime("%Y-%m-%d 00:00:00"),
                fim.strftime("%Y-%m-%d 23:59:59"),
                **filtros
            )

            if not registros:
                messagebox.showinfo("Relatório", "Nenhum ticket encontrado para estes filtros.", parent=janela)
                return

            # Preenche a tabela do relatório com os dados encontrados
            for reg in registros:
                reg = list(reg)
                reg[1] = self.formatar_data_para_exibicao(reg[1]) # Formata a data
                # Se o colaborador for None (posição 7), exibe um campo vazio
                if reg[7] is None:
                    reg[7] = ""
                arvore_relatorio.insert("", tk.END, values=tuple(reg))


        def exportar_csv():
            if not arvore_relatorio.get_children(): messagebox.showwarning("Aviso", "Gere um relatório primeiro.", parent=janela); return
            path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV file", "*.csv"), ("All files", "*.*")])
            if not path: return
            try:
                with open(path, "w", newline="", encoding="utf-8") as file:
                    writer = csv.writer(file); writer.writerow([arvore_relatorio.heading(col)["text"] for col in arvore_relatorio["columns"]])
                    for item_id in arvore_relatorio.get_children(): writer.writerow(arvore_relatorio.item(item_id)['values'])
                messagebox.showinfo("Sucesso", f"Relatório exportado para {path}", parent=janela)
            except IOError as e: messagebox.showerror("Erro de Exportação", f"Não foi possível salvar o arquivo: {e}", parent=janela)

    # Busca um ticket com base no termo digitado
    def buscar_ticket(self):
        termo = self.entrada_busca.get().strip()
        if not termo: self.atualizar_abas(); return
        resultados = self.gerenciador.get_todos_tickets()
        if resultados: resultados = [t for t in resultados if termo.lower() in str(t).lower()]
        if not resultados: messagebox.showinfo("Busca", f"Nenhum ticket encontrado para o termo '{termo}'.")
        self.atualizar_abas(resultados)

    # Limpa o campo de busca e recarrega todos os tickets
    def limpar_busca(self):
        self.entrada_busca.delete(0, tk.END); self.atualizar_abas()

    # Pega o ID do ticket que está selecionado na aba ativa
    def get_id_selecionado(self):
        try:
            aba_selecionada = self.notebook.tab(self.notebook.select(), "text"); arvore_ativa = None
            if aba_selecionada == "Todos os Tickets": arvore_ativa = self.arvore_todos
            elif aba_selecionada == "Tickets Abertos": arvore_ativa = self.arvore_abertos
            elif aba_selecionada == "Tickets Fechados": arvore_ativa = self.arvore_fechados
            if arvore_ativa and arvore_ativa.focus(): return arvore_ativa.item(arvore_ativa.focus())['values'][0]
        except (IndexError, KeyError): return None
        return None

    # Exclui o ticket que está selecionado
    def excluir_ticket_selecionado(self):
        ticket_id = self.get_id_selecionado()
        if not ticket_id: messagebox.showwarning("Aviso", "Selecione um ticket para excluir."); return
        if messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir o Ticket #{ticket_id}?"):
            self.gerenciador.deletar_ticket(ticket_id); self.mostrar_status(f"Ticket #{ticket_id} excluído com sucesso!"); self.atualizar_abas()

    # ***** INÍCIO DA NOVA FUNÇÃO *****
    # Reimprime um ticket selecionado, se ele estiver aberto
    def reimprimir_ticket_selecionado(self):
        ticket_id = self.get_id_selecionado()
        if not ticket_id:
            messagebox.showwarning("Aviso", "Selecione um ticket para reimprimir.")
            return

        # Busca todas as informações do ticket
        ticket_info = self.gerenciador.get_ticket_por_id(ticket_id)
        if not ticket_info:
            messagebox.showerror("Erro", "Não foi possível encontrar os dados do ticket.")
            return

        # Verifica se o ticket está aberto (status é o 7º item, índice 6)
        if ticket_info[6] != 'Aberto':
            messagebox.showinfo("Aviso", "Apenas tickets com status 'Aberto' podem ser reimpressos.")
            return

        # Prepara os dados para a função de impressão
        ticket_para_impressao = {
            "id": ticket_info[0],
            "data_solicitacao": self.formatar_data_para_exibicao(ticket_info[1]),
            "setor": ticket_info[2],
            "predio": ticket_info[3],
            "local": ticket_info[4],
            "descricao": ticket_info[5]
        }
        # Chama a função que já existe para imprimir
        self.imprimir_ticket(ticket_para_impressao)
    # ***** FIM DA NOVA FUNÇÃO *****

    # Evento de duplo clique na aba de tickets abertos (para finalizar)
    def ao_clicar_aberto(self, event):
        ticket_id = self.get_id_selecionado()
        if ticket_id: self.abrir_janela_finalizar(ticket_id)

    # Evento de duplo clique na aba "Todos" (decide se abre a janela de edição ou de detalhes)
    def ao_clicar_todos(self, event):
        ticket_id = self.get_id_selecionado()
        if ticket_id:
            ticket_info = self.gerenciador.get_ticket_por_id(ticket_id)
            if ticket_info:
                if ticket_info[6] == 'Aberto': self.abrir_janela_ticket(ticket_id) # Se aberto, edita
                else: self.abrir_janela_finalizar(ticket_id) # Se fechado, mostra detalhes

# Ponto de entrada do programa
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    # Garante que a conexão com o banco de dados seja fechada ao fechar a janela
    root.protocol("WM_DELETE_WINDOW", lambda: (app.db.fechar(), root.destroy()))
    root.mainloop()