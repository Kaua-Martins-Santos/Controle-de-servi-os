from datetime import datetime

class GerenciadorTickets:
    def __init__(self, db):
        self.db = db 

    # --- Funções para gerenciar Tickets ---

    def criar_ticket(self, id_externo, data, setor, predio, local, descricao, prioridade):
        data_iso = datetime.strptime(data, "%d/%m/%Y").strftime("%Y-%m-%d %H:%M:%S")
        
        # --- ALTERAÇÃO ---
        # Agora inserimos na coluna 'id_externo'. O 'ticket_db_id' será gerado automaticamente.
        self.db.executar_query("INSERT INTO tickets (id_externo, data_solicitacao, setor, predio, local, descricao, status, prioridade) VALUES (?, ?, ?, ?, ?, ?, 'Aberto', ?)", 
                               (id_externo, data_iso, setor, predio, local, descricao, prioridade))

    def atualizar_ticket(self, ticket_db_id, id_externo, data, setor, predio, local, descricao, prioridade):
        # --- ALTERAÇÃO ---
        # O 'ticket_db_id' é usado no WHERE para encontrar o ticket exato.
        # Agora também atualizamos o 'id_externo'.
        data_iso = datetime.strptime(data, "%d/%m/%Y").strftime("%Y-%m-%d %H:%M:%S")
        self.db.executar_query("UPDATE tickets SET id_externo=?, data_solicitacao=?, setor=?, predio=?, local=?, descricao=?, prioridade=? WHERE ticket_db_id=?", 
                               (id_externo, data_iso, setor, predio, local, descricao, prioridade, ticket_db_id))

    def finalizar_ticket(self, ticket_db_id, colaboradores, solucao, horas, data_inicio, data_termino):
        colaboradores_str = ", ".join(colaboradores) 
        data_conclusao = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        data_inicio_iso = datetime.strptime(data_inicio, "%d/%m/%Y").strftime("%Y-%m-%d")
        data_termino_iso = datetime.strptime(data_termino, "%d/%m/%Y").strftime("%Y-%m-%d")
        
        # --- ALTERAÇÃO ---
        # O WHERE agora usa o ID interno 'ticket_db_id'
        self.db.executar_query("UPDATE tickets SET status='Finalizado', colaborador=?, solucao=?, data_conclusao=?, horas=?, data_inicio_servico=?, data_termino_servico=? WHERE ticket_db_id=?", 
                               (colaboradores_str, solucao, data_conclusao, horas, data_inicio_iso, data_termino_iso, ticket_db_id))

    def deletar_ticket(self, ticket_db_id):
        # --- ALTERAÇÃO ---
        self.db.executar_query("DELETE FROM tickets WHERE ticket_db_id=?", (ticket_db_id,))

    def get_ticket_por_id(self, ticket_db_id):
        # --- ALTERAÇÃO ---
        # Busca pelo 'ticket_db_id' e retorna TODAS as colunas, incluindo os dois IDs.
        query = """
            SELECT ticket_db_id, id_externo, data_solicitacao, setor, predio, local, descricao, status,
                   prioridade, colaborador, solucao, data_conclusao, horas,
                   data_inicio_servico, data_termino_servico
            FROM tickets WHERE ticket_db_id=?
        """
        return self.db.executar_query(query, (ticket_db_id,), fetch='one')

    def get_todos_tickets(self):
        # --- ALTERAÇÃO ---
        # Seleciona o 'ticket_db_id' (para uso interno) e o 'id_externo' (para exibição).
        return self.db.executar_query("SELECT ticket_db_id, id_externo, data_solicitacao, setor, predio, local, status, prioridade FROM tickets ORDER BY ticket_db_id DESC", fetch='all')

    def get_fechados_tickets(self):
        # --- ALTERAÇÃO ---
        return self.db.executar_query("SELECT ticket_db_id, id_externo, setor, local, colaborador, data_inicio_servico, data_termino_servico, horas FROM tickets WHERE status = 'Finalizado' ORDER BY data_termino_servico DESC, ticket_db_id DESC", fetch='all')

    def buscar_tickets_por_termo(self, termo):
        # --- ALTERAÇÃO ---
        # Busca também no 'id_externo'
        query = """
            SELECT ticket_db_id, id_externo, data_solicitacao, setor, predio, local, status, prioridade
            FROM tickets
            WHERE
                id_externo LIKE ? OR
                setor LIKE ? OR
                predio LIKE ? OR
                local LIKE ? OR
                descricao LIKE ?
            ORDER BY ticket_db_id DESC
        """
        termo_busca = f"%{termo}%"
        params = (termo_busca, termo_busca, termo_busca, termo_busca, termo_busca)
        return self.db.executar_query(query, params, fetch='all')

    # --- Funções de Colaboradores e Prédios (sem alterações) ---

    def get_colaboradores(self):
        return self.db.executar_query("SELECT id, nome, setor FROM colaboradores ORDER BY nome", fetch='all')
    def get_colaboradores_por_setor(self, setor):
        return self.db.executar_query("SELECT nome FROM colaboradores WHERE setor=? ORDER BY nome", (setor,), fetch='all')
    def adicionar_colaborador(self, nome, setor):
        self.db.executar_query("INSERT INTO colaboradores (nome, setor) VALUES (?, ?)", (nome, setor))
    def editar_colaborador(self, id_colab, nome, setor):
        self.db.executar_query("UPDATE colaboradores SET nome=?, setor=? WHERE id=?", (nome, setor, id_colab))
    def remover_colaborador(self, id_colab):
        self.db.executar_query("DELETE FROM colaboradores WHERE id=?", (id_colab,))

    def get_predios(self):
        return self.db.executar_query("SELECT id, nome FROM predios ORDER BY nome", fetch='all')
    def adicionar_predio(self, nome):
        return self.db.executar_query("INSERT INTO predios (nome) VALUES (?)", (nome,))
    def editar_predio(self, id_predio, nome):
        self.db.executar_query("UPDATE predios SET nome=? WHERE id=?", (nome, id_predio))
    def remover_predio(self, id_predio):
        self.db.executar_query("DELETE FROM predios WHERE id=?", (id_predio,))

    # --- Função para Relatórios ---
    def get_tickets_relatorio(self, inicio, fim, **filtros):
        # --- ALTERAÇÃO ---
        # Adicionamos 'id_externo' na seleção
        query = """
            SELECT ticket_db_id, id_externo, data_solicitacao, setor, predio, local, status, prioridade, 
                   colaborador, horas, data_termino_servico, solucao 
            FROM tickets 
            WHERE data_solicitacao BETWEEN ? AND ?
        """
        params = [inicio, fim] 

        if filtros.get("predio"):
            query += " AND predio = ?"
            params.append(filtros["predio"])
        if filtros.get("colaborador"):
            query += " AND colaborador LIKE ?" 
            params.append(f'%{filtros["colaborador"]}%')
        if filtros.get("status"):
            query += " AND status = ?"
            params.append(filtros["status"])
        if filtros.get("setor"):
            query += " AND setor = ?"
            params.append(filtros["setor"])
        if filtros.get("local"):
            query += " AND local LIKE ?"
            params.append(f'%{filtros["local"]}%')
        if filtros.get("prioridade"):
            query += " AND prioridade = ?"
            params.append(filtros["prioridade"])
        
        # --- ALTERAÇÃO ---
        # Adicionamos filtro por id_externo
        if filtros.get("id_externo"):
            query += " AND id_externo LIKE ?"
            params.append(f'%{filtros["id_externo"]}%')

        query += " ORDER BY data_solicitacao DESC"
        return self.db.executar_query(query, tuple(params), fetch='all')