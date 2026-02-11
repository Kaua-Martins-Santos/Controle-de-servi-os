import psycopg2
from tkinter import messagebox
import os

class BancoDeDados:
    
    def __init__(self):
        # --- CONFIGURAÇÃO DO POSTGRESQL ---
        self.db_config = {
            "host": "10.132.16.185",       
            "database": "controle_servicos", # O nome que você escolheu
            "user": "postgres",
            "password": "Un@sp",    # ATENÇÃO: Coloque sua senha aqui
            "port": "5432"
        }

        try:
            self.conexao = psycopg2.connect(**self.db_config)
            self.cursor = self.conexao.cursor()
            
        except Exception as e:
            messagebox.showerror("Erro de Conexão", 
                f"Não foi possível conectar ao servidor.\nErro: {e}\n\nVerifique se o PostgreSQL está rodando.")
            raise SystemExit(f"Erro ao conectar ao DB: {e}")

    def executar_query(self, query, params=(), fetch=None):
        try:
            # TRUQUE DE COMPATIBILIDADE:
            # O código antigo usa '?' (SQLite), o Postgres usa '%s'.
            # Trocamos automaticamente para não precisar mexer no resto do programa.
            query_postgres = query.replace('?', '%s')
            
            self.cursor.execute(query_postgres, params)
            
            if fetch == 'one':
                return self.cursor.fetchone()
            if fetch == 'all':
                return self.cursor.fetchall()
            
            self.conexao.commit()
            return self.cursor
            
        except psycopg2.Error as e:
            self.conexao.rollback() # Cancela a operação se der erro para não travar
            messagebox.showerror("Erro de Banco de Dados", f"Erro ao executar operação: {e}")
            return None

    def fechar(self):
        if self.conexao:
            self.conexao.close()