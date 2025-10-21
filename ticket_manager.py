from datetime import datetime

# Esta classe gerencia a lógica de negócios relacionada aos tickets.
# Ela usa a classe BancoDeDados para interagir com o banco, mas
# a lógica principal (o que fazer com os dados) fica aqui.
# Isso separa as responsabilidades e deixa o código mais limpo.
class GerenciadorTickets:
    def __init__(self, db):
        self.db = db # Recebe a conexão com o banco de dados já aberta

    # --- Funções para gerenciar Tickets ---

    def criar_ticket(self, id_externo, data, setor, predio, local, descricao, prioridade):
        # Converte a data do formato brasileiro (DD/MM/YYYY) para o formato do banco (YYYY-MM-DD)
        data_iso = datetime.strptime(data, "%d/%m/%Y").strftime("%Y-%m-%d %H:%M:%S")
        # Executa o comando SQL para inserir um novo ticket
        self.db.executar_query("INSERT INTO tickets (id, data_solicitacao, setor, predio, local, descricao, status, prioridade) VALUES (?, ?, ?, ?, ?, ?, 'Aberto', ?)", (id_externo, data_iso, setor, predio, local, descricao, prioridade))

    def atualizar_ticket(self, id_ticket, data, setor, predio, local, descricao, prioridade):
        data_iso = datetime.strptime(data, "%d/%m/%Y").strftime("%Y-%m-%d %H:%M:%S")
        self.db.executar_query("UPDATE tickets SET data_solicitacao=?, setor=?, predio=?, local=?, descricao=?, prioridade=? WHERE id=?", (data_iso, setor, predio, local, descricao, prioridade, id_ticket))

    def finalizar_ticket(self, id_ticket, colaborador, solucao, horas, data_inicio, data_termino):
        data_conclusao = datetime.now().strftime("%Y-%m-%d %H:%M:%S") # Pega a data e hora atuais
        data_inicio_iso = datetime.strptime(data_inicio, "%d/%m/%Y").strftime("%Y-%m-%d")
        data_termino_iso = datetime.strptime(data_termino, "%d/%m/%Y").strftime("%Y-%m-%d")
        # Atualiza o status do ticket para 'Finalizado' e preenche os campos de conclusão
        self.db.executar_query("UPDATE tickets SET status='Finalizado', colaborador=?, solucao=?, data_conclusao=?, horas=?, data_inicio_servico=?, data_termino_servico=? WHERE id=?", (colaborador, solucao, data_conclusao, horas, data_inicio_iso, data_termino_iso, id_ticket))

    def deletar_ticket(self, id_ticket):
        self.db.executar_query("DELETE FROM tickets WHERE id=?", (id_ticket,))

    def get_ticket_por_id(self, id_ticket):
        # Busca todas as informações de um ticket específico pelo seu ID
        query = """
            SELECT id, data_solicitacao, setor, predio, local, descricao, status,
                   prioridade, colaborador, solucao, data_conclusao, horas,
                   data_inicio_servico, data_termino_servico
            FROM tickets WHERE id=?
        """
        return self.db.executar_query(query, (id_ticket,), fetch='one')

    def get_todos_tickets(self):
        # Busca os campos principais de todos os tickets para exibir na lista geral
        return self.db.executar_query("SELECT id, data_solicitacao, setor, predio, local, status, prioridade FROM tickets ORDER BY id DESC", fetch='all')

    def get_fechados_tickets(self):
        # Busca os tickets finalizados para a aba de "Fechados"
        return self.db.executar_query("SELECT id, setor, local, colaborador, data_inicio_servico, data_termino_servico, horas FROM tickets WHERE status = 'Finalizado' ORDER BY data_termino_servico DESC, id DESC", fetch='all')

    # --- Funções para gerenciar Colaboradores ---

    def get_colaboradores(self):
        return self.db.executar_query("SELECT id, nome, setor FROM colaboradores ORDER BY nome", fetch='all')

    def get_colaboradores_por_setor(self, setor):
        # Útil para filtrar os colaboradores na hora de finalizar um ticket
        return self.db.executar_query("SELECT nome FROM colaboradores WHERE setor=? ORDER BY nome", (setor,), fetch='all')

    def adicionar_colaborador(self, nome, setor):
        self.db.executar_query("INSERT INTO colaboradores (nome, setor) VALUES (?, ?)", (nome, setor))

    def editar_colaborador(self, id_colab, nome, setor):
        self.db.executar_query("UPDATE colaboradores SET nome=?, setor=? WHERE id=?", (nome, setor, id_colab))

    def remover_colaborador(self, id_colab):
        self.db.executar_query("DELETE FROM colaboradores WHERE id=?", (id_colab,))


    # --- Funções para gerenciar Prédios ---

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
        # ***** INÍCIO DA ALTERAÇÃO *****
        # Adicionamos 'colaborador' na lista de campos a serem selecionados
        query = "SELECT id, data_solicitacao, setor, predio, local, status, prioridade, colaborador FROM tickets WHERE data_solicitacao BETWEEN ? AND ?"
        # ***** FIM DA ALTERAÇÃO *****

        params = [inicio, fim] # Parâmetros iniciais (datas)

        # Adicionamos os outros filtros dinamicamente, se eles foram preenchidos
        if filtros.get("predio"):
            query += " AND predio = ?"
            params.append(filtros["predio"])

        if filtros.get("colaborador"):
            query += " AND colaborador = ?"
            params.append(filtros["colaborador"])

        if filtros.get("status"):
            query += " AND status = ?"
            params.append(filtros["status"])

        if filtros.get("setor"):
            query += " AND setor = ?"
            params.append(filtros["setor"])

        if filtros.get("local"):
            # Usamos o 'LIKE' para buscar por parte do texto no campo local
            query += " AND local LIKE ?"
            params.append(f'%{filtros["local"]}%')

        if filtros.get("prioridade"):
            query += " AND prioridade = ?"
            params.append(filtros["prioridade"])

        # Ordenamos o resultado para ficar mais organizado
        query += " ORDER BY data_solicitacao DESC"

        # Executa a busca com a query e os parâmetros que montamos
        return self.db.executar_query(query, tuple(params), fetch='all')