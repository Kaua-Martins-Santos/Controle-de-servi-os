import sqlite3
import psycopg2
import os

# --- SUAS CONFIGURAÇÕES (EDITAR AQUI) ---
HOST = "localhost"
DB_NAME = "controle_servicos"  # Coloque o nome que você criou no passo 1
USER = "postgres"
PASSWORD = "Un@sp"    # A senha que você criou na instalação

# Tenta achar o banco antigo automaticamente
home_dir = os.path.expanduser('~')
db_path_antigo = os.path.join(home_dir, 'Documents', 'ControleDeServicosData', 'tickets.db')

def migrar():
    if not os.path.exists(db_path_antigo):
        print(f"ERRO: Não encontrei o banco antigo em: {db_path_antigo}")
        return

    print("--- INICIANDO MIGRAÇÃO ---")
    
    # Conecta no Velho (SQLite)
    try:
        conn_old = sqlite3.connect(db_path_antigo)
        cursor_old = conn_old.cursor()
        print("Conectado ao SQLite (Origem).")
    except Exception as e:
        print(f"Erro ao abrir SQLite: {e}")
        return

    # Conecta no Novo (PostgreSQL)
    try:
        conn_new = psycopg2.connect(host=HOST, database=DB_NAME, user=USER, password=PASSWORD)
        cursor_new = conn_new.cursor()
        print("Conectado ao PostgreSQL (Destino).")
    except Exception as e:
        print(f"Erro ao conectar no PostgreSQL: {e}")
        return

    # 1. Migrar Prédios
    print("\nMigrando Prédios...")
    cursor_old.execute("SELECT nome FROM predios")
    for row in cursor_old.fetchall():
        cursor_new.execute("INSERT INTO predios (nome) VALUES (%s) ON CONFLICT DO NOTHING", row)

    # 2. Migrar Colaboradores
    print("Migrando Colaboradores...")
    cursor_old.execute("SELECT nome, setor FROM colaboradores")
    for row in cursor_old.fetchall():
        cursor_new.execute("INSERT INTO colaboradores (nome, setor) VALUES (%s, %s)", row)

    # 3. Migrar Tickets
    print("Migrando Tickets...")
    # Lendo as colunas na ordem exata
    cols = "id_externo, data_solicitacao, setor, predio, local, descricao, status, prioridade, colaborador, solucao, data_conclusao, horas, data_inicio_servico, data_termino_servico"
    cursor_old.execute(f"SELECT {cols} FROM tickets")
    tickets = cursor_old.fetchall()

    query_insert = f"INSERT INTO tickets ({cols}) VALUES ({'%s, ' * 13}%s)"
    
    for t in tickets:
        cursor_new.execute(query_insert, t)

    conn_new.commit()
    print(f"\nSUCESSO! {len(tickets)} tickets migrados.")
    conn_old.close()
    conn_new.close()

if __name__ == "__main__":
    migrar()