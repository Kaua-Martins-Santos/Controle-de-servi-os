import sqlite3
from tkinter import messagebox

# Esta classe cuida de toda a comunicação com o banco de dados.
# Isso ajuda a manter o código organizado, pois todas as operações
# de banco de dados (criar, ler, atualizar, deletar) ficam em um só lugar.
class BancoDeDados:
    # O construtor é chamado quando criamos um novo objeto BancoDeDados.
    # Ele se conecta ao arquivo do banco de dados (ou o cria se não existir).
    def __init__(self, db_name="tickets.db"):
        self.conexao = sqlite3.connect(db_name)
        self.cursor = self.conexao.cursor()
        self.criar_tabelas() # Garante que as tabelas existam.

    # Esta função cria as tabelas do banco de dados se elas ainda não existirem.
    # Também verifica se alguma coluna nova precisa ser adicionada,
    # o que é útil para atualizar o programa sem perder dados.
    def criar_tabelas(self):
        try:
            # Cria a tabela principal de 'tickets'
            self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS tickets (
                id INTEGER PRIMARY KEY,
                data_solicitacao TEXT, setor TEXT, predio TEXT, local TEXT, descricao TEXT,
                status TEXT, prioridade TEXT, colaborador TEXT, solucao TEXT,
                data_conclusao TEXT, horas TEXT, data_inicio_servico TEXT, data_termino_servico TEXT
            )
            """)
            # Cria tabelas auxiliares para 'colaboradores' e 'predios'
            self.cursor.execute("CREATE TABLE IF NOT EXISTS colaboradores (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL, setor TEXT NOT NULL)")
            self.cursor.execute("CREATE TABLE IF NOT EXISTS predios (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL UNIQUE)")

            # Verifica se colunas mais recentes já existem, e as adiciona se não.
            # Isso evita erros se o usuário estiver atualizando de uma versão antiga do programa.
            self.cursor.execute("PRAGMA table_info(tickets)")
            colunas = [col[1] for col in self.cursor.fetchall()]
            if 'prioridade' not in colunas: self.cursor.execute("ALTER TABLE tickets ADD COLUMN prioridade TEXT DEFAULT 'Média'")
            if 'data_inicio_servico' not in colunas: self.cursor.execute("ALTER TABLE tickets ADD COLUMN data_inicio_servico TEXT")
            if 'data_termino_servico' not in colunas: self.cursor.execute("ALTER TABLE tickets ADD COLUMN data_termino_servico TEXT")
            self.conexao.commit() # Salva as alterações no banco
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Não foi possível criar as tabelas: {e}")

    # Uma função central para executar qualquer comando (query) no banco de dados.
    # 'params' são os valores a serem inseridos de forma segura para evitar SQL injection.
    # 'fetch' define se queremos um resultado ('one'), todos ('all'), ou nenhum.
    def executar_query(self, query, params=(), fetch=None):
        try:
            self.cursor.execute(query, params)
            if fetch == 'one':
                return self.cursor.fetchone() # Retorna apenas um resultado
            if fetch == 'all':
                return self.cursor.fetchall() # Retorna uma lista de resultados
            self.conexao.commit() # Se não for uma busca, salva as alterações
            return self.cursor
        except sqlite3.Error as e:
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao executar a operação: {e}")
            return None

    # Fecha a conexão com o banco de dados. É importante fazer isso
    # quando o programa for encerrado para não corromper o arquivo.
    def fechar(self):
        if self.conexao:
            self.conexao.close()