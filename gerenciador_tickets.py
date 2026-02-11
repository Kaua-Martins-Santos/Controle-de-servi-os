import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import csv
import re
import json  # Importante para a fila de impressão

# Importando as classes dos outros arquivos
from database import BancoDeDados
from ticket_manager import GerenciadorTickets

class App:
    def __init__(self, root):
        self.root = root
        self.db = BancoDeDados()
        self.gerenciador = GerenciadorTickets(self.db)
        self.theme_name = "dark" 

        self.themes = {
            "light": {
                "BG_COLOR": "#F0F2F5", "FRAME_COLOR": "#FFFFFF", "PRIMARY_COLOR": "#005A9C",
                "SECONDARY_COLOR": "#4682B4", "TEXT_COLOR": "#333333", "HEADER_BG_COLOR": "#FFFFFF",
                "SELECTED_COLOR": "#005A9C", "WHITE_COLOR": "#FFFFFF", "FIELD_BG_COLOR": "#FFFFFF",
                "PLACEHOLDER_COLOR": "#A9A9A9"
            },
            "dark": {
                "BG_COLOR": "#202124", "FRAME_COLOR": "#2D2E30", "PRIMARY_COLOR": "#4682B4",
                "SECONDARY_COLOR": "#005A9C", "TEXT_COLOR": "#E8EAED", "HEADER_BG_COLOR": "#2D2E30",
                "SELECTED_COLOR": "#4682B4", "WHITE_COLOR": "#E8EAED", "FIELD_BG_COLOR": "#3c4043",
                "PLACEHOLDER_COLOR": "#9E9E9E"
            }
        }
        self.style = ttk.Style()
        self.style.theme_use('clam')

        self.criar_widgets()
        self.apply_theme()
        self.atualizar_abas()

    def apply_theme(self):
        colors = self.themes[self.theme_name]
        LABEL_BOLD_FONT = ("Calibri", 11, "bold")
        self.root.configure(bg=colors["BG_COLOR"])

        self.style.configure(".", background=colors["BG_COLOR"], foreground=colors["TEXT_COLOR"], font=("Calibri", 11), bordercolor=colors["FIELD_BG_COLOR"])
        self.style.configure("TButton", background=colors["PRIMARY_COLOR"], foreground=colors["WHITE_COLOR"], font=("Calibri", 11, "bold"), padding=6, borderwidth=0)
        self.style.map("TButton", background=[('active', colors["SECONDARY_COLOR"])])
        self.style.configure("TLabel", background=colors["BG_COLOR"], foreground=colors["TEXT_COLOR"])
        self.style.configure("Bold.TLabel", font=LABEL_BOLD_FONT, background=colors["BG_COLOR"])
        self.style.configure("Status.TLabel", font=("Calibri", 10), background=colors["BG_COLOR"])
        self.style.configure("Header.TFrame", background=colors["HEADER_BG_COLOR"])
        self.style.configure("Header.TLabel", background=colors["HEADER_BG_COLOR"], foreground=colors["TEXT_COLOR"])
        self.style.configure("TFrame", background=colors["BG_COLOR"])
        self.style.configure("TNotebook", background=colors["BG_COLOR"], borderwidth=0)
        self.style.configure("TNotebook.Tab", background=colors["FRAME_COLOR"], padding=[10, 5], font=("Calibri", 11, "bold"), foreground=colors["TEXT_COLOR"])
        self.style.map("TNotebook.Tab", background=[("selected", colors["SELECTED_COLOR"])], foreground=[("selected", colors["WHITE_COLOR"])])
        self.style.configure("Treeview", background=colors["FRAME_COLOR"], fieldbackground=colors["FIELD_BG_COLOR"], foreground=colors["TEXT_COLOR"], rowheight=28, font=("Calibri", 11))
        self.style.configure("Treeview.Heading", background=colors["PRIMARY_COLOR"], foreground=colors["WHITE_COLOR"], font=("Calibri", 11, "bold"), padding=8)
        self.style.map("Treeview", background=[('selected', colors["SELECTED_COLOR"])], foreground=[('selected', colors["WHITE_COLOR"])])
        self.style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])
        self.style.configure("TEntry", fieldbackground=colors["FIELD_BG_COLOR"], foreground=colors["TEXT_COLOR"], insertcolor=colors["TEXT_COLOR"])
        self.style.map('TCombobox', fieldbackground=[('readonly', colors["FIELD_BG_COLOR"])], selectbackground=[('readonly', colors["SELECTED_COLOR"])], foreground=[('readonly', colors["TEXT_COLOR"])])

        if hasattr(self, 'entrada_busca'):
            self.on_search_focus_out(None)

    def toggle_theme(self):
        self.theme_name = "light" if self.theme_name == "dark" else "dark"
        self.theme_button.config(text="🌙" if self.theme_name == "light" else "☀️")
        self.apply_theme()

    def criar_widgets(self):
        self.root.title("UNASP - Sistema de Controle de Serviços")
        self.root.geometry("1200x750")

        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        header_frame = ttk.Frame(main_frame, padding=(10, 10), style="Header.TFrame")
        header_frame.pack(fill=tk.X)
        
        ttk.Label(header_frame, text="UNASP", font=("Montserrat", 18, "bold"), foreground=self.themes[self.theme_name]["PRIMARY_COLOR"], background=self.themes[self.theme_name]["HEADER_BG_COLOR"]).pack(side=tk.LEFT, padx=(5,0))
        ttk.Label(header_frame, text="Controle de Serviços", font=("Calibri", 16), style="Header.TLabel").pack(side=tk.LEFT, padx=(10, 0))
        
        right_header_frame = ttk.Frame(header_frame, style="Header.TFrame")
        right_header_frame.pack(side=tk.RIGHT)

        self.search_placeholder = "Digite para buscar..."
        self.entrada_busca = ttk.Entry(right_header_frame, width=30, font=("Calibri", 11))
        self.entrada_busca.pack(side=tk.LEFT, padx=(0, 5), ipady=4)
        self.entrada_busca.bind("<FocusIn>", self.on_search_focus_in)
        self.entrada_busca.bind("<FocusOut>", self.on_search_focus_out)
        self.entrada_busca.bind("<KeyRelease>", self.on_search_change)
        self.on_search_focus_out(None)
        
        self.theme_button = ttk.Button(right_header_frame, text="☀️", command=self.toggle_theme, width=4)
        self.theme_button.pack(side=tk.RIGHT, padx=(10, 5))

        toolbar_frame = ttk.Frame(main_frame, padding=(10, 5))
        toolbar_frame.pack(fill=tk.X)
        ttk.Button(toolbar_frame, text="Novo Ticket", command=self.abrir_janela_ticket, style="TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar_frame, text="Reimprimir Ticket", command=self.reimprimir_ticket_selecionado, style="TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar_frame, text="Excluir Ticket", command=self.excluir_ticket_selecionado, style="TButton").pack(side=tk.LEFT, padx=5)
        ttk.Separator(toolbar_frame, orient='vertical').pack(side=tk.LEFT, padx=10, fill='y', pady=5)
        ttk.Button(toolbar_frame, text="Relatório", command=self.abrir_janela_relatorio, style="TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar_frame, text="Gerenciar", command=self.abrir_janela_gerenciamento, style="TButton").pack(side=tk.LEFT, padx=5)
        
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=10, padx=10)

        self.status_bar = ttk.Label(self.root, text="", anchor=tk.W, style="Status.TLabel", padding=(5, 2))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        colunas_gerais = ("ID", "Data Solicitação", "Setor", "Prédio", "Local", "Status", "Prioridade")
        self.arvore_todos = self.criar_aba(self.notebook, "Todos os Tickets", colunas_gerais)
        self.arvore_abertos = self.criar_aba(self.notebook, "Tickets Abertos", colunas_gerais)
        colunas_fechados = ("ID", "Setor", "Local", "Colaborador", "Início Serviço", "Término Serviço", "Horas")
        self.arvore_fechados = self.criar_aba(self.notebook, "Tickets Fechados", colunas_fechados, [80, 150, 180, 150, 150, 150, 100])

        self.arvore_abertos.bind("<Double-1>", self.ao_clicar_aberto)
        self.arvore_todos.bind("<Double-1>", self.ao_clicar_todos)
        self.arvore_fechados.bind("<Double-1>", self.ao_clicar_todos)

    def on_search_focus_in(self, event):
        if self.entrada_busca.get() == self.search_placeholder:
            self.entrada_busca.delete(0, tk.END)
            self.entrada_busca.config(foreground=self.themes[self.theme_name]["TEXT_COLOR"])

    def on_search_focus_out(self, event):
        if not self.entrada_busca.get():
            self.entrada_busca.insert(0, self.search_placeholder)
            self.entrada_busca.config(foreground=self.themes[self.theme_name]["PLACEHOLDER_COLOR"])

    def on_search_change(self, event):
        termo = self.entrada_busca.get().strip()
        if termo == self.search_placeholder:
            termo = ""
        resultados = self.gerenciador.buscar_tickets_por_termo(termo) if termo else None
        self.atualizar_abas(resultados)

    def mostrar_status(self, mensagem):
        self.status_bar.config(text=mensagem)
        self.root.after(3000, self.limpar_status)

    def limpar_status(self):
        self.status_bar.config(text="")

    def criar_aba(self, parent, texto_aba, colunas, larguras=None):
        frame = ttk.Frame(parent)
        parent.add(frame, text=texto_aba)
        arvore = ttk.Treeview(frame, columns=colunas, show="headings", selectmode="browse")
        if not larguras:
            larguras = [140] * len(colunas)
            larguras[0] = 80; larguras[1] = 160
        for i, col in enumerate(colunas):
            arvore.heading(col, text=col)
            arvore.column(col, width=larguras[i], anchor='center')
        arvore.pack(fill=tk.BOTH, expand=True)
        return arvore

    def formatar_data_para_exibicao(self, data_db):
        if not data_db: return ""
        try:
            return datetime.strptime(data_db.split()[0], "%Y-%m-%d").strftime("%d/%m/%Y")
        except (ValueError, TypeError, IndexError):
            return data_db

    def atualizar_abas(self, data=None):
        for arvore in [self.arvore_todos, self.arvore_abertos, self.arvore_fechados]:
            for item in arvore.get_children():
                arvore.delete(item)

        tickets_gerais = data if data is not None else self.gerenciador.get_todos_tickets()
        
        if tickets_gerais:
            for row in tickets_gerais:
                row_lista = list(row)
                if len(row_lista) > 2:
                    row_lista[2] = self.formatar_data_para_exibicao(row_lista[2])
                
                valores_para_exibir = tuple(row_lista[1:])
                ticket_db_id = row_lista[0]
                
                self.arvore_todos.insert("", tk.END, iid=ticket_db_id, values=valores_para_exibir)
                
                if len(row_lista) > 6 and row_lista[6] == "Aberto":
                    self.arvore_abertos.insert("", tk.END, iid=ticket_db_id, values=valores_para_exibir)

        if data is None:
            tickets_fechados = self.gerenciador.get_fechados_tickets()
            if tickets_fechados:
                for row in tickets_fechados:
                    row_lista = list(row)
                    if len(row_lista) > 6:
                        row_lista[5] = self.formatar_data_para_exibicao(row_lista[5])
                        row_lista[6] = self.formatar_data_para_exibicao(row_lista[6])
                    
                    valores_para_exibir = tuple(row_lista[1:])
                    ticket_db_id = row_lista[0]
                    self.arvore_fechados.insert("", tk.END, iid=ticket_db_id, values=valores_para_exibir)
    
    # ---------------------------------------------------------
    # MODIFICAÇÃO PRINCIPAL: Adição de Scrollbar na Janela de Tickets
    # ---------------------------------------------------------
    def abrir_janela_ticket(self, ticket_id=None):
        janela = tk.Toplevel(self.root)
        janela.title("Novo Ticket" if ticket_id is None else "Editando Ticket")
        janela.geometry("470x550") # Ajustado para caber scroll e campos
        janela.transient(self.root)
        janela.grab_set()
        
        colors = self.themes[self.theme_name]
        janela.configure(bg=colors["BG_COLOR"])

        # Container Principal
        container = ttk.Frame(janela)
        container.pack(fill=tk.BOTH, expand=True)

        # Canvas para rolagem
        canvas = tk.Canvas(container, bg=colors["BG_COLOR"], highlightthickness=0)
        
        # Barra de rolagem
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        
        # Frame interno (onde os widgets vão ficar)
        frame = ttk.Frame(canvas, padding=20)
        
        # Janela dentro do Canvas
        canvas_window = canvas.create_window((0, 0), window=frame, anchor="nw")

        # Funções para configurar a rolagem
        def _configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _configure_window_width(event):
            canvas.itemconfig(canvas_window, width=event.width)

        def _on_mousewheel(event):
            if canvas.winfo_exists():
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        frame.bind("<Configure>", _configure_scroll_region)
        canvas.bind("<Configure>", _configure_window_width)
        janela.bind("<MouseWheel>", _on_mousewheel) # Habilita rolagem com roda do mouse

        # Empacotamento
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)

        # --- FIM DA CONFIGURAÇÃO DE SCROLL ---

        ticket = self.gerenciador.get_ticket_por_id(ticket_id) if ticket_id else None
        if ticket:
            janela.title(f"Editando Ticket #{ticket[1]}")

        ttk.Label(frame, text="ID:").grid(row=0, column=0, sticky="w", pady=5)
        entrada_id = ttk.Entry(frame, width=33)
        entrada_id.grid(row=0, column=1, pady=5)
        if ticket:
            entrada_id.insert(0, ticket[1])

        def formatar_data_entry(event, entry_widget):
            if event.keysym == 'BackSpace': return
            text = entry_widget.get().replace("/", "")[:8]
            new_text = ""
            if len(text) > 2: new_text += text[:2] + "/" + text[2:]
            else: new_text = text
            if len(text) > 4: new_text = new_text[:5] + "/" + new_text[5:]
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, new_text)
            entry_widget.icursor(len(new_text))

        ttk.Label(frame, text="Data:").grid(row=1, column=0, sticky="w", pady=5)
        entrada_data = ttk.Entry(frame, width=20)
        entrada_data.grid(row=1, column=1, sticky="w", pady=5)
        entrada_data.bind("<KeyRelease>", lambda e: formatar_data_entry(e, entrada_data))
        
        if ticket:
            entrada_data.insert(0, self.formatar_data_para_exibicao(ticket[2]))
        ttk.Button(frame, text="Hoje", width=6, command=lambda: (entrada_data.delete(0, tk.END), entrada_data.insert(0, datetime.now().strftime("%d/%m/%Y")))).grid(row=1, column=1, sticky="e", padx=5)

        ttk.Label(frame, text="Setor:").grid(row=2, column=0, sticky="w", pady=5)
        combo_setor = ttk.Combobox(frame, values=["Eletrica", "Hidraulica", "Construção", "Pintura", "Refrigeração", "Serralheria", "Marcenaria"], state="readonly", width=30)
        combo_setor.grid(row=2, column=1, pady=5)
        
        if ticket: 
            combo_setor.set(ticket[3]) 
        
        ttk.Label(frame, text="Prédio:").grid(row=3, column=0, sticky="w", pady=5)
        entrada_predio = ttk.Entry(frame, width=33)
        entrada_predio.grid(row=3, column=1, pady=5)
        lista_predios_completa = [row[1] for row in self.gerenciador.get_predios() or []]
        lista_sugestoes = tk.Listbox(janela, height=4, bg=colors["FIELD_BG_COLOR"], fg=colors["TEXT_COLOR"], highlightbackground=colors["PRIMARY_COLOR"])

        def on_predio_keyup(event):
            if event.keysym in ("Up", "Down", "Return", "Escape"): return
            texto = entrada_predio.get().lower()
            if not texto: lista_sugestoes.place_forget(); return
            sugestoes = [p for p in lista_predios_completa if texto in p.lower()]
            lista_sugestoes.delete(0, tk.END)
            for s in sugestoes: lista_sugestoes.insert(tk.END, s)
            if sugestoes:
                x = frame.winfo_x() + entrada_predio.winfo_x()
                # Ajuste de coordenadas considerando o canvas
                y = frame.winfo_y() + entrada_predio.winfo_y() + entrada_predio.winfo_height() + 5
                # Listbox precisa ser relativa à janela principal (Toplevel) para não ficar cortada pelo frame
                # Mas aqui é tricky. Se a posição for complexa, melhor manter simples.
                # Tentativa de posicionamento relativo ao container global
                lista_sugestoes.place(in_=frame, x=entrada_predio.winfo_x(), y=entrada_predio.winfo_y() + entrada_predio.winfo_height(), width=entrada_predio.winfo_width())
                lista_sugestoes.config(height=len(sugestoes) if len(sugestoes) <=4 else 4)
                lista_sugestoes.lift()
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
        if ticket: entrada_predio.insert(0, ticket[4]) 

        ttk.Label(frame, text="Local:").grid(row=4, column=0, sticky="w", pady=5)
        entrada_local = ttk.Entry(frame, width=33)
        entrada_local.grid(row=4, column=1, pady=5)
        if ticket: entrada_local.insert(0, ticket[5]) 

        ttk.Label(frame, text="Prioridade:").grid(row=5, column=0, sticky="w", pady=5)
        combo_prioridade = ttk.Combobox(frame, values=["Baixa", "Média", "Alta"], state="readonly", width=30)
        combo_prioridade.grid(row=5, column=1, pady=5)
        if ticket and ticket[8]: combo_prioridade.set(ticket[8]) 
        else: combo_prioridade.set("Média")

        ttk.Label(frame, text="Descrição:").grid(row=6, column=0, sticky="nw", pady=5)
        texto_descricao = tk.Text(frame, width=35, height=8, font=("Calibri", 11), bg=colors["FIELD_BG_COLOR"], fg=colors["TEXT_COLOR"], insertbackground=colors["TEXT_COLOR"], relief="solid", borderwidth=1)
        texto_descricao.grid(row=6, column=1, pady=5)
        if ticket: texto_descricao.insert("1.0", ticket[6]) 
        
        # Botão Salvar dentro do fluxo normal (agora com scroll, ele sempre será alcançável)
        ttk.Button(frame, text="Salvar Ticket", command=lambda: salvar(), style="TButton").grid(row=7, column=1, pady=20, sticky="e")

        def limpar_campos_para_novo_ticket():
            entrada_id.delete(0, tk.END)
            entrada_data.delete(0, tk.END)
            entrada_predio.delete(0, tk.END)
            entrada_local.delete(0, tk.END)
            texto_descricao.delete("1.0", tk.END)
            combo_setor.set('') 
            entrada_id.focus()

        def salvar():
            id_externo_val = entrada_id.get().strip()
            data_str = entrada_data.get()
            setor_val = combo_setor.get()
            predio_val = entrada_predio.get()
            local_val = entrada_local.get()
            descricao_val = texto_descricao.get("1.0", tk.END).strip()
            prioridade_val = combo_prioridade.get()
            
            campos_obrigatorios = [data_str, setor_val, predio_val, descricao_val, prioridade_val]
            if not all(campos_obrigatorios):
                messagebox.showwarning("Aviso", "Preencha todos os campos obrigatórios!", parent=janela)
                return
            
            try:
                if len(data_str) != 10: raise ValueError
                datetime.strptime(data_str, "%d/%m/%Y")
            except ValueError:
                messagebox.showwarning("Aviso", "Data inválida! Use o formato DD/MM/YYYY.", parent=janela)
                return

            if ticket_id is None:
                self.gerenciador.criar_ticket(id_externo_val, data_str, setor_val, predio_val, local_val, descricao_val, prioridade_val)
                msg_id = id_externo_val if id_externo_val else "(sem ID)"
                self.mostrar_status(f"Ticket #{msg_id} criado com sucesso!")
                
                if messagebox.askyesno("Imprimir", f"Deseja imprimir o Ticket #{msg_id}?", parent=janela):
                    ticket_info_impressao = { "id": id_externo_val, "data_solicitacao": data_str, "setor": setor_val, "predio": predio_val, "local": local_val, "descricao": descricao_val }
                    self.imprimir_ticket(ticket_info_impressao)
                limpar_campos_para_novo_ticket()
            else:
                self.gerenciador.atualizar_ticket(ticket_id, id_externo_val, data_str, setor_val, predio_val, local_val, descricao_val, prioridade_val)
                self.mostrar_status(f"Ticket #{id_externo_val} atualizado com sucesso!")
                janela.destroy()

            self.atualizar_abas()

    def imprimir_ticket(self, ticket_info):
        # --- ATUALIZAÇÃO DA FILA DE IMPRESSÃO ---
        try:
            # Em vez de imprimir, salvamos o pedido no banco
            # Convertemos o dicionário de dados para texto (JSON)
            dados_json = json.dumps(ticket_info, ensure_ascii=False)
            
            # Adapte '?' para '%s' se seu banco for Postgres e seu driver não fizer automático (mas nossa classe Database já faz!)
            self.db.executar_query(
                "INSERT INTO fila_impressao (dados_ticket, status) VALUES (?, 'Pendente')", 
                (dados_json,)
            )
            
            self.mostrar_status(f"Impressão do Ticket #{ticket_info['id']} enviada para o servidor!")
            messagebox.showinfo("Sucesso", "O ticket foi enviado para a fila de impressão do servidor.")
            
        except Exception as e:
            messagebox.showerror("Erro de Impressão", f"Não foi possível enviar para a fila: {e}")

    def abrir_janela_finalizar(self, ticket_id):
        ticket = self.gerenciador.get_ticket_por_id(ticket_id)
        if not ticket: messagebox.showerror("Erro", "Ticket não encontrado!"); return
        
        is_finalizado = (ticket[7] == 'Finalizado')

        janela = tk.Toplevel(self.root)
        janela.title(f"Detalhes do Ticket #{ticket[1]}")
        janela.resizable(False, False); janela.transient(self.root); janela.grab_set()
        colors = self.themes[self.theme_name]; janela.configure(bg=colors["BG_COLOR"])
        frame = ttk.Frame(janela, padding=15); frame.pack(fill=tk.BOTH, expand=True)

        if is_finalizado:
            janela.geometry("480x450")
            ttk.Label(frame, text="Descrição do Problema:", style="Bold.TLabel").pack(anchor="w", pady=(5, 2))
            desc_text = tk.Text(frame, height=4, width=50, font=("Calibri", 11), relief="flat")
            desc_text.insert("1.0", ticket[6] or "Não informado"); desc_text.config(state="disabled", bg=colors["BG_COLOR"], fg=colors["TEXT_COLOR"]); desc_text.pack(anchor="w", fill="x", pady=(0, 5)) 

            ttk.Label(frame, text="Solução Aplicada:", style="Bold.TLabel").pack(anchor="w", pady=(5, 2))
            solucao_text = tk.Text(frame, height=5, width=50, font=("Calibri", 11), relief="flat")
            solucao_text.insert("1.0", ticket[10] or "Não informada"); solucao_text.config(state="disabled", bg=colors["BG_COLOR"], fg=colors["TEXT_COLOR"]); solucao_text.pack(anchor="w", fill="x", pady=(0, 10)) 

            info_frame = ttk.Frame(frame); info_frame.pack(fill="x", anchor="w")
            
            ttk.Label(info_frame, text="Colaborador(es):", style="Bold.TLabel").grid(row=0, column=0, sticky="nw", pady=2)
            ttk.Label(info_frame, text=ticket[9] or "Não informado", wraplength=350).grid(row=0, column=1, sticky="w", padx=10)
            
            ttk.Label(info_frame, text="Início do Serviço:", style="Bold.TLabel").grid(row=1, column=0, sticky="w", pady=2); ttk.Label(info_frame, text=self.formatar_data_para_exibicao(ticket[13])).grid(row=1, column=1, sticky="w", padx=10) 
            ttk.Label(info_frame, text="Término do Serviço:", style="Bold.TLabel").grid(row=2, column=0, sticky="w", pady=2); ttk.Label(info_frame, text=self.formatar_data_para_exibicao(ticket[14])).grid(row=2, column=1, sticky="w", padx=10) 
            ttk.Label(info_frame, text="Horas Gastas:", style="Bold.TLabel").grid(row=3, column=0, sticky="w", pady=2); ttk.Label(info_frame, text=ticket[12] or "Não informado").grid(row=3, column=1, sticky="w", padx=10) 

            ttk.Button(frame, text="Fechar", command=janela.destroy).pack(side="bottom", pady=(20, 0))
        else:
            janela.geometry("450x550")
            ttk.Label(frame, text=f"Descrição: {ticket[6]}", wraplength=400).pack(anchor="w", pady=2) 
            dates_frame = ttk.Frame(frame); dates_frame.pack(fill=tk.X, pady=(15,0))
            ttk.Label(dates_frame, text="Data de Início:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
            entrada_data_inicio = ttk.Entry(dates_frame, width=15); entrada_data_inicio.grid(row=0, column=1, pady=2)
            ttk.Button(dates_frame, text="Hoje", width=6, command=lambda: (entrada_data_inicio.delete(0, tk.END), entrada_data_inicio.insert(0, datetime.now().strftime("%d/%m/%Y")))).grid(row=0, column=2, padx=5)
            ttk.Label(dates_frame, text="Data de Término:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
            entrada_data_termino = ttk.Entry(dates_frame, width=15); entrada_data_termino.grid(row=1, column=1, pady=2)
            ttk.Button(dates_frame, text="Hoje", width=6, command=lambda: (entrada_data_termino.delete(0, tk.END), entrada_data_termino.insert(0, datetime.now().strftime("%d/%m/%Y")))).grid(row=1, column=2, padx=5)
            
            ttk.Label(frame, text="Colaborador(es):").pack(anchor="w", pady=(15, 0))
            lista_nomes = [row[0] for row in self.gerenciador.get_colaboradores_por_setor(ticket[3]) or []] 
            
            colab_frame = ttk.Frame(frame)
            colab_frame.pack(anchor="w", fill="x", expand=True)
            
            colab_scrollbar = ttk.Scrollbar(colab_frame, orient=tk.VERTICAL)
            lista_colaborador = tk.Listbox(
                colab_frame,
                selectmode=tk.MULTIPLE,
                height=5,
                width=38,
                font=("Calibri", 11),
                bg=colors["FIELD_BG_COLOR"],
                fg=colors["TEXT_COLOR"],
                selectbackground=colors["SELECTED_COLOR"],
                selectforeground=colors["WHITE_COLOR"],
                relief="solid",
                borderwidth=1,
                yscrollcommand=colab_scrollbar.set
            )
            colab_scrollbar.config(command=lista_colaborador.yview)
            colab_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            lista_colaborador.pack(side=tk.LEFT, fill=tk.X, expand=True)

            for nome in lista_nomes:
                lista_colaborador.insert(tk.END, nome)

            ttk.Label(frame, text="Solução:").pack(anchor="w", pady=(10, 0)); texto_solucao = tk.Text(frame, width=40, height=5, font=("Calibri", 11), bg=colors["FIELD_BG_COLOR"], fg=colors["TEXT_COLOR"], insertbackground=colors["TEXT_COLOR"], relief="solid", borderwidth=1); texto_solucao.pack(anchor="w")
            horas_frame = ttk.Frame(frame); horas_frame.pack(anchor="w", pady=(10,0))
            ttk.Label(horas_frame, text="Horas gastas:").pack(side=tk.LEFT)
            entrada_horas = ttk.Entry(horas_frame, width=15); entrada_horas.pack(side=tk.LEFT, padx=5)

            def validar_horas(valor): return re.fullmatch(r'\d{1,2}:\d{2}', valor) is not None

            def salvar():
                data_inicio_str = entrada_data_inicio.get(); data_termino_str = entrada_data_termino.get(); horas_str = entrada_horas.get().strip()
                if not data_inicio_str or not data_termino_str: messagebox.showwarning("Aviso", "As datas de início e término são obrigatórias.", parent=janela); return
                try: data_inicio_obj = datetime.strptime(data_inicio_str, "%d/%m/%Y"); data_termino_obj = datetime.strptime(data_termino_str, "%d/%m/%Y")
                except ValueError: messagebox.showwarning("Aviso", "Formato de data inválido! Use DD/MM/YYYY.", parent=janela); return
                if data_termino_obj < data_inicio_obj: messagebox.showwarning("Aviso", "A data de término não pode ser anterior à data de início.", parent=janela); return
                if not validar_horas(horas_str): messagebox.showwarning("Aviso", "Formato de horas inválido! Use HH:MM.", parent=janela); return
                
                selected_indices = lista_colaborador.curselection()
                if not selected_indices:
                    messagebox.showwarning("Aviso", "Selecione pelo menos um colaborador.", parent=janela)
                    return
                
                nomes_selecionados = [lista_colaborador.get(i) for i in selected_indices]
                outros_campos = [nomes_selecionados, texto_solucao.get("1.0", tk.END).strip(), horas_str]
                
                if not all(outros_campos[1:]):
                     messagebox.showwarning("Aviso", "Preencha todos os campos de finalização!", parent=janela); return
                
                self.gerenciador.finalizar_ticket(ticket_id, *outros_campos, data_inicio_str, data_termino_str)
                self.mostrar_status(f"Ticket #{ticket[1]} finalizado com sucesso!")
                janela.destroy()
                self.atualizar_abas()
            
            ttk.Button(frame, text="Finalizar Ticket", command=salvar).pack(pady=20, ipadx=50)

    def abrir_janela_gerenciamento(self):
        janela = tk.Toplevel(self.root); janela.title("Gerenciamento"); janela.geometry("650x450"); janela.resizable(False, False); janela.transient(self.root); janela.grab_set()
        colors = self.themes[self.theme_name]; janela.configure(bg=colors["BG_COLOR"])
        notebook = ttk.Notebook(janela); notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        frame_colab = ttk.Frame(notebook); notebook.add(frame_colab, text="Colaboradores")
        frame_lista_colab = ttk.Frame(frame_colab, padding=10); frame_lista_colab.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        colunas_colab = ("ID", "Nome", "Setor")
        larguras_colab = [40, 200, 150]
        arvore_colab = ttk.Treeview(frame_lista_colab, columns=colunas_colab, show="headings", selectmode="browse")
        for i, col in enumerate(colunas_colab):
            arvore_colab.heading(col, text=col)
            arvore_colab.column(col, width=larguras_colab[i], anchor='center')
        arvore_colab.pack(fill=tk.BOTH, expand=True)

        frame_form_colab = ttk.Frame(frame_colab, padding=20); frame_form_colab.pack(side=tk.RIGHT, fill=tk.Y)
        ttk.Label(frame_form_colab, text="Nome:").pack(anchor="w"); entrada_nome_colab = ttk.Entry(frame_form_colab, width=30); entrada_nome_colab.pack(anchor="w", pady=(0, 10))
        ttk.Label(frame_form_colab, text="Setor:").pack(anchor="w"); combo_setor_colab = ttk.Combobox(frame_form_colab, values=["Eletrica", "Hidraulica", "Construção", "Pintura", "Refrigeração", "Serralheria", "Marcenaria"], state="readonly", width=28); combo_setor_colab.pack(anchor="w")
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

        frame_predio = ttk.Frame(notebook); notebook.add(frame_predio, text="Prédios")
        frame_lista_predio = ttk.Frame(frame_predio, padding=10); frame_lista_predio.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        colunas_predio = ("ID", "Nome")
        larguras_predio = [40, 250]
        arvore_predio = ttk.Treeview(frame_lista_predio, columns=colunas_predio, show="headings", selectmode="browse")
        for i, col in enumerate(colunas_predio):
            arvore_predio.heading(col, text=col)
            arvore_predio.column(col, width=larguras_predio[i], anchor='center')
        arvore_predio.pack(fill=tk.BOTH, expand=True)

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

    def abrir_janela_relatorio(self):
        janela = tk.Toplevel(self.root); janela.title("Relatório de Tickets"); janela.geometry("1200x600"); janela.transient(self.root); janela.grab_set()
        colors = self.themes[self.theme_name]; janela.configure(bg=colors["BG_COLOR"])
        frame_filtros = ttk.Frame(janela, padding=10); frame_filtros.pack(fill=tk.X)
        
        ttk.Label(frame_filtros, text="Período:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        combo_periodo = ttk.Combobox(frame_filtros, values=["Hoje", "Esta Semana", "Este Mês", "Mês Passado", "Este Ano", "Personalizado"], state="readonly", width=15)
        combo_periodo.grid(row=0, column=1, sticky="w", padx=5, pady=5); combo_periodo.set("Este Mês")

        ttk.Label(frame_filtros, text="De:").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        entrada_data_inicio = ttk.Entry(frame_filtros, width=15)
        entrada_data_inicio.grid(row=0, column=3, sticky="w", padx=5, pady=5)
        
        ttk.Label(frame_filtros, text="Até:").grid(row=0, column=4, sticky="w", padx=5, pady=5)
        entrada_data_fim = ttk.Entry(frame_filtros, width=15)
        entrada_data_fim.grid(row=0, column=5, sticky="w", padx=5, pady=5)

        def formatar_data_entry(event, entry_widget):
            if event.keysym == 'BackSpace': return
            text = entry_widget.get().replace("/", "")[:8]
            new_text = ""
            if len(text) > 2: new_text += text[:2] + "/" + text[2:]
            else: new_text = text
            if len(text) > 4: new_text = new_text[:5] + "/" + new_text[5:]
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, new_text)
            entry_widget.icursor(len(new_text))
            if combo_periodo.get() != "Personalizado":
                combo_periodo.set("Personalizado")

        entrada_data_inicio.bind("<KeyRelease>", lambda e: formatar_data_entry(e, entrada_data_inicio))
        entrada_data_fim.bind("<KeyRelease>", lambda e: formatar_data_entry(e, entrada_data_fim))

        def atualizar_datas_por_periodo(event=None):
            escolha = combo_periodo.get()
            hoje = datetime.now()
            ini, fim = "", ""

            if escolha == "Hoje":
                ini = fim = hoje.strftime("%d/%m/%Y")
            elif escolha == "Esta Semana":
                inicio_sem = hoje - timedelta(days=hoje.weekday())
                fim_sem = inicio_sem + timedelta(days=6)
                ini = inicio_sem.strftime("%d/%m/%Y")
                fim = fim_sem.strftime("%d/%m/%Y")
            elif escolha == "Este Mês":
                inicio_mes = hoje.replace(day=1)
                proximo_mes = (inicio_mes + timedelta(days=32)).replace(day=1)
                fim_mes = proximo_mes - timedelta(days=1)
                ini = inicio_mes.strftime("%d/%m/%Y")
                fim = fim_mes.strftime("%d/%m/%Y")
            elif escolha == "Mês Passado":
                inicio_mes_atual = hoje.replace(day=1)
                fim_mes_passado = inicio_mes_atual - timedelta(days=1)
                inicio_mes_passado = fim_mes_passado.replace(day=1)
                ini = inicio_mes_passado.strftime("%d/%m/%Y")
                fim = fim_mes_passado.strftime("%d/%m/%Y")
            elif escolha == "Este Ano":
                ini = hoje.replace(day=1, month=1).strftime("%d/%m/%Y")
                fim = hoje.replace(day=31, month=12).strftime("%d/%m/%Y")
            
            if escolha != "Personalizado":
                entrada_data_inicio.delete(0, tk.END); entrada_data_inicio.insert(0, ini)
                entrada_data_fim.delete(0, tk.END); entrada_data_fim.insert(0, fim)

        combo_periodo.bind("<<ComboboxSelected>>", atualizar_datas_por_periodo)
        atualizar_datas_por_periodo()

        ttk.Label(frame_filtros, text="Status:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        combo_status = ttk.Combobox(frame_filtros, values=["Todos", "Aberto", "Finalizado"], state="readonly", width=15)
        combo_status.grid(row=1, column=1, sticky="w", padx=5, pady=5); combo_status.set("Todos")
        
        ttk.Label(frame_filtros, text="Setor:").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        combo_setor = ttk.Combobox(frame_filtros, values=["Todos", "Eletrica", "Hidraulica", "Construção", "Pintura", "Refrigeração", "Serralheria"], state="readonly", width=15)
        combo_setor.grid(row=1, column=3, sticky="w", padx=5, pady=5); combo_setor.set("Todos")
        
        predios = ["Todos"] + [p[1] for p in self.gerenciador.get_predios()]
        ttk.Label(frame_filtros, text="Prédio:").grid(row=1, column=4, sticky="w", padx=5, pady=5)
        combo_predio = ttk.Combobox(frame_filtros, values=predios, state="readonly", width=15)
        combo_predio.grid(row=1, column=5, sticky="w", padx=5, pady=5); combo_predio.set("Todos")
        
        colaboradores = ["Todos"] + [c[1] for c in self.gerenciador.get_colaboradores()]
        ttk.Label(frame_filtros, text="Colaborador:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        combo_colaborador = ttk.Combobox(frame_filtros, values=colaboradores, state="readonly", width=15)
        combo_colaborador.grid(row=2, column=1, sticky="w", padx=5, pady=5); combo_colaborador.set("Todos")
        
        ttk.Label(frame_filtros, text="Prioridade:").grid(row=2, column=2, sticky="w", padx=5, pady=5)
        combo_prioridade = ttk.Combobox(frame_filtros, values=["Todas", "Baixa", "Média", "Alta"], state="readonly", width=15)
        combo_prioridade.grid(row=2, column=3, sticky="w", padx=5, pady=5); combo_prioridade.set("Todas")
        
        ttk.Label(frame_filtros, text="ID (busca):").grid(row=2, column=4, sticky="w", padx=5, pady=5)
        entry_id_externo = ttk.Entry(frame_filtros, width=18)
        entry_id_externo.grid(row=2, column=5, sticky="w", padx=5, pady=5)
        
        ttk.Label(frame_filtros, text="Local (busca):").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        entry_local = ttk.Entry(frame_filtros, width=18)
        entry_local.grid(row=3, column=1, sticky="w", padx=5, pady=5)
        
        
        frame_botoes = ttk.Frame(janela, padding=10); frame_botoes.pack(fill=tk.X)
        ttk.Button(frame_botoes, text="Gerar Relatório", command=lambda: gerar()).pack(side=tk.LEFT, padx=10)
        ttk.Button(frame_botoes, text="Exportar para CSV", command=lambda: exportar_csv()).pack(side=tk.LEFT, padx=10)
        frame_tabela = ttk.Frame(janela, padding=10); frame_tabela.pack(fill=tk.BOTH, expand=True)
        
        colunas_relatorio = ("ID", "Data", "Setor", "Prédio", "Local", "Status", "Prioridade", "Colaborador", "Horas", "Data Término", "Solução")
        arvore_relatorio = ttk.Treeview(frame_tabela, columns=colunas_relatorio, show="headings", selectmode="browse")
        larguras = [40, 90, 100, 110, 130, 80, 80, 120, 60, 90, 150]
        
        for i, col in enumerate(colunas_relatorio):
            arvore_relatorio.heading(col, text=col)
            arvore_relatorio.column(col, width=larguras[i], anchor='center')
        arvore_relatorio.pack(fill=tk.BOTH, expand=True)

        def gerar():
            for item in arvore_relatorio.get_children(): arvore_relatorio.delete(item)
            data_ini_str = entrada_data_inicio.get()
            data_fim_str = entrada_data_fim.get()

            try:
                dt_ini = datetime.strptime(data_ini_str, "%d/%m/%Y")
                dt_fim = datetime.strptime(data_fim_str, "%d/%m/%Y")
                inicio_query = dt_ini.strftime("%Y-%m-%d 00:00:00")
                fim_query = dt_fim.strftime("%Y-%m-%d 23:59:59")
            except ValueError:
                messagebox.showwarning("Data Inválida", "Verifique se as datas estão no formato DD/MM/AAAA.", parent=janela)
                return

            filtros = {
                "predio": combo_predio.get() if combo_predio.get() != "Todos" else None,
                "colaborador": combo_colaborador.get() if combo_colaborador.get() != "Todos" else None,
                "status": combo_status.get() if combo_status.get() != "Todos" else None,
                "setor": combo_setor.get() if combo_setor.get() != "Todos" else None,
                "local": entry_local.get().strip() or None,
                "prioridade": combo_prioridade.get() if combo_prioridade.get() != "Todas" else None,
                "id_externo": entry_id_externo.get().strip() or None
            }
            
            registros = self.gerenciador.get_tickets_relatorio(inicio_query, fim_query, **filtros)
            if not registros:
                messagebox.showinfo("Relatório", "Nenhum ticket encontrado para estes filtros.", parent=janela)
                return
            
            for reg in registros:
                reg_lista = list(reg[1:])
                reg_lista[1] = self.formatar_data_para_exibicao(reg_lista[1])
                status_do_ticket = reg_lista[5]
                horas_do_ticket = reg_lista[8]
                if status_do_ticket != 'Finalizado': reg_lista[8] = "-" 
                else: reg_lista[8] = horas_do_ticket if horas_do_ticket else "N/A"
                if reg_lista[7] is None: reg_lista[7] = "" 
                reg_lista[9] = self.formatar_data_para_exibicao(reg_lista[9]) 
                if reg_lista[10] is None: reg_lista[10] = ""
                arvore_relatorio.insert("", tk.END, values=tuple(reg_lista))

        def exportar_csv():
            if not arvore_relatorio.get_children(): 
                messagebox.showwarning("Aviso", "Gere um relatório primeiro.", parent=janela)
                return
            path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV file", "*.csv"), ("All files", "*.*")])
            if not path: return
            try:
                with open(path, "w", newline="", encoding="utf-8-sig") as file:
                    writer = csv.writer(file, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                    writer.writerow([arvore_relatorio.heading(col)["text"] for col in arvore_relatorio["columns"]])
                    for item_id in arvore_relatorio.get_children():
                        row_values = arvore_relatorio.item(item_id)['values']
                        clean_values = []
                        for v in row_values:
                            if isinstance(v, str): clean_values.append(v.replace('\n', ' ').replace('\r', '').strip())
                            else: clean_values.append(v)
                        writer.writerow(clean_values)
                messagebox.showinfo("Sucesso", f"Relatório exportado para {path}", parent=janela)
            except IOError as e: 
                messagebox.showerror("Erro de Exportação", f"Não foi possível salvar o arquivo: {e}", parent=janela)

    def get_id_selecionado(self):
        try:
            aba_selecionada = self.notebook.tab(self.notebook.select(), "text")
            arvore_ativa = None
            if aba_selecionada == "Todos os Tickets": arvore_ativa = self.arvore_todos
            elif aba_selecionada == "Tickets Abertos": arvore_ativa = self.arvore_abertos
            elif aba_selecionada == "Tickets Fechados": arvore_ativa = self.arvore_fechados
            if arvore_ativa: return arvore_ativa.focus() 
        except (IndexError, KeyError): return None
        return None

    def excluir_ticket_selecionado(self):
        try:
            aba_selecionada = self.notebook.tab(self.notebook.select(), "text")
            arvore_ativa = None
            if aba_selecionada == "Todos os Tickets": arvore_ativa = self.arvore_todos
            elif aba_selecionada == "Tickets Abertos": arvore_ativa = self.arvore_abertos
            elif aba_selecionada == "Tickets Fechados": arvore_ativa = self.arvore_fechados
            if not arvore_ativa: return
        except Exception: return

        ticket_db_id = arvore_ativa.focus() 
        if not ticket_db_id: messagebox.showwarning("Aviso", "Selecione um ticket para excluir."); return
        id_externo_para_exibir = arvore_ativa.item(ticket_db_id)['values'][0]
        if messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir o Ticket #{id_externo_para_exibir}?"):
            self.gerenciador.deletar_ticket(ticket_db_id) 
            self.mostrar_status(f"Ticket #{id_externo_para_exibir} excluído com sucesso!")
            self.atualizar_abas()

    def reimprimir_ticket_selecionado(self):
        ticket_db_id = self.get_id_selecionado() 
        if not ticket_db_id: messagebox.showwarning("Aviso", "Selecione um ticket para reimprimir."); return
        ticket_info = self.gerenciador.get_ticket_por_id(ticket_db_id)
        if not ticket_info: messagebox.showerror("Erro", "Não foi possível encontrar os dados do ticket."); return
        if ticket_info[7] != 'Aberto': messagebox.showinfo("Aviso", "Apenas tickets com status 'Aberto' podem ser reimpressos."); return
        ticket_para_impressao = {
            "id": ticket_info[1], 
            "data_solicitacao": self.formatar_data_para_exibicao(ticket_info[2]), 
            "setor": ticket_info[3],
            "predio": ticket_info[4],
            "local": ticket_info[5],
            "descricao": ticket_info[6]
        }
        self.imprimir_ticket(ticket_para_impressao)

    def ao_clicar_aberto(self, event):
        ticket_db_id = self.get_id_selecionado() 
        if ticket_db_id: self.abrir_janela_finalizar(ticket_db_id)

    def ao_clicar_todos(self, event):
        ticket_db_id = self.get_id_selecionado() 
        if ticket_db_id:
            ticket_info = self.gerenciador.get_ticket_por_id(ticket_db_id)
            if ticket_info:
                if ticket_info[7] == 'Aberto': self.abrir_janela_ticket(ticket_db_id) 
                else: self.abrir_janela_finalizar(ticket_db_id) 

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.protocol("WM_DELETE_WINDOW", lambda: (app.db.fechar(), root.destroy()))
    root.mainloop()