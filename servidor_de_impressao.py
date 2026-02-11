import time
import json
import win32print
import psycopg2
from database import BancoDeDados

def imprimir_fisicamente(ticket_info):
    try:
        printer_name = win32print.GetDefaultPrinter()
        h_printer = win32print.OpenPrinter(printer_name)
        try:
            # --- CÓDIGO DE IMPRESSÃO ORIGINAL (Formatado para o Script) ---
            ESC = b'\x1b'; GS = b'\x1d'; LF = b'\x0a'
            CMD_INIT = ESC + b'@'; CMD_ALIGN_CENTER = ESC + b'a\x01'; CMD_ALIGN_LEFT = ESC + b'a\x00'
            CMD_FONT_BOLD_ON = ESC + b'E\x01'; CMD_FONT_BOLD_OFF = ESC + b'E\x00'
            CMD_DOUBLE_HW_ON = GS + b'!\x11'; CMD_DOUBLE_HW_OFF = GS + b'!\x00'; CMD_CUT = GS + b'V\x01'

            ticket_data = CMD_INIT + CMD_ALIGN_CENTER + CMD_DOUBLE_HW_ON + CMD_FONT_BOLD_ON + b'ORDEM DE SERVICO\n' + CMD_DOUBLE_HW_OFF + b'UNASP-EC\n' + CMD_FONT_BOLD_OFF + b'========================================\n'
            
            id_str = str(ticket_info["id"])
            data_str = ticket_info["data_solicitacao"]
            cabecalho = f'{id_str} {data_str.rjust(40 - len(id_str))}\n'
            
            # Codificação correta para acentos na impressora térmica (cp850 ou pc850)
            ticket_data += cabecalho.encode("cp850") + b'========================================\n' + LF + CMD_ALIGN_LEFT
            ticket_data += CMD_FONT_BOLD_ON + b'PREDIO: ' + CMD_FONT_BOLD_OFF + f'{ticket_info["predio"]}\n'.encode('cp850', errors='replace')
            ticket_data += CMD_FONT_BOLD_ON + b'LOCAL: ' + CMD_FONT_BOLD_OFF + f'{ticket_info["local"]}\n'.encode('cp850', errors='replace') + LF
            ticket_data += CMD_FONT_BOLD_ON + f'SERVICO SOLICITADO: ({ticket_info["data_solicitacao"]})\n'.encode('cp850', errors='replace') + CMD_FONT_BOLD_OFF + f'{ticket_info["descricao"]}\n'.encode('cp850', errors='replace') + LF
            ticket_data += b'----------------------------------------\n' + CMD_FONT_BOLD_ON + b'SETOR SOLICITADO: ' + CMD_FONT_BOLD_OFF + f'{ticket_info["setor"]}\n'.encode('cp850', errors='replace')
            ticket_data += CMD_FONT_BOLD_ON + b'SOLUCAO:\n' + CMD_FONT_BOLD_OFF + LF * 4
            ticket_data += b'----------------------------------------\n' + CMD_FONT_BOLD_ON + b'COLABORADOR RESPONSAVEL:\n' + CMD_FONT_BOLD_OFF + LF
            ticket_data += b'----------------------------------------\n' + b'HORAS GASTAS: __:__\n' + b'INICIO DO SERVICO: __/__/____\n' + b'TERMINO DO SERVICO: __/__/____\n' + LF
            ticket_data += b'========================================\n' + CMD_FONT_BOLD_ON + b'RECEBIMENTO DO SERVICO\n' + CMD_FONT_BOLD_OFF + LF + b'NOME:\n\n' + b'ASS.:___________________________________\n' + LF * 4 + CMD_CUT

            h_job = win32print.StartDocPrinter(h_printer, 1, ("Ticket Remoto", None, "RAW"))
            try:
                win32print.WritePrinter(h_printer, ticket_data)
            finally:
                win32print.EndDocPrinter(h_printer)
        finally:
            win32print.ClosePrinter(h_printer)
        return True
    except Exception as e:
        print(f"Erro ao imprimir: {e}")
        return False

def iniciar_servidor():
    print("--- SERVIDOR DE IMPRESSÃO INICIADO ---")
    print("Aguardando pedidos de impressão...")
    
    db = BancoDeDados() # Reusa sua conexão já configurada

    while True:
        try:
            # Busca pedidos 'Pendente'
            # Note o uso de %s pois estamos usando o psycopg2 direto ou via sua classe adaptada
            db.cursor.execute("SELECT id, dados_ticket FROM fila_impressao WHERE status = 'Pendente' ORDER BY id ASC LIMIT 1")
            pedido = db.cursor.fetchone()

            if pedido:
                id_pedido = pedido[0]
                dados_json = pedido[1]
                print(f"Processando pedido #{id_pedido}...")

                # Converte o texto JSON de volta para dicionário
                ticket_info = json.loads(dados_json)

                # Tenta imprimir
                sucesso = imprimir_fisicamente(ticket_info)

                if sucesso:
                    db.cursor.execute("UPDATE fila_impressao SET status = 'Impresso' WHERE id = %s", (id_pedido,))
                    db.conexao.commit()
                    print(f"Pedido #{id_pedido} impresso com sucesso!")
                else:
                    print(f"Falha ao imprimir pedido #{id_pedido}.")
            
            # Espera 2 segundos antes de checar de novo para não travar o PC
            time.sleep(2)

        except Exception as e:
            print(f"Erro no loop: {e}")
            # Se cair a conexão, tenta reconectar (recriando o objeto db)
            try:
                db.fechar()
            except: 
                pass
            time.sleep(5)
            try:
                db = BancoDeDados()
            except:
                pass

if __name__ == "__main__":
    iniciar_servidor()