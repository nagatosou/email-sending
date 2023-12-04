import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
import PySimpleGUI as sg
import psycopg2
from datetime import datetime

# Lista de domínios de e-mail válidos
dominios_validos = ["gmail.com", "hotmail.com", "outlook.com", "yahoo.com", "mail.com", "icloud.com", "bol.com.br", "magalhaesadv.adv.br"]

def is_valid_email(email):
    # Verifique se o e-mail de destino possui um dos domínios válidos
    for dominio in dominios_validos:
        if email.endswith("@" + dominio):
            return True
    return False



# -----------------------------------------------------------------------------------------------------------------------------------------------


def enviar_emails(email_de_envio, senha, assunto, campanha, mensagem, arquivo_excel):
    enviados_com_sucesso = []
    nao_enviados = []

    try:
        df = pd.read_excel(arquivo_excel)
    except pd.errors.ParserError:
        print(f"Erro ao carregar o arquivo Excel: {arquivo_excel}")
        return enviados_com_sucesso, nao_enviados

    # Percorra as linhas do DataFrame para obter os endereços de e-mail e mensagens
    for index, row in df.iterrows():
        email_destino = row['Email']
        nome_cliente = row['Cliente']

        # Verifique se o nome do cliente não está em branco
        if not isinstance(nome_cliente, str) or not nome_cliente.strip():
            print(f"Nome do cliente em branco para e-mail: {email_destino}. Mensagem não enviada.")
            continue

        if pd.notna(email_destino) and isinstance(mensagem, str) and mensagem.strip() and is_valid_email(email_destino):
            # Crie o objeto MIMEMultipart
            msg = MIMEMultipart()
            msg['From'] = email_de_envio
            msg['To'] = email_destino
            msg['Subject'] = assunto

            # Substitua a tag [cliente] na mensagem pelo nome do cliente
            mensagem_personalizada = mensagem.replace('[cliente]', nome_cliente)
            print(f"Mensagem Personalizada: {mensagem_personalizada}")

            # Crie o objeto MIMEText após a personalização da mensagem
            mensagem_final = MIMEText(mensagem_personalizada, 'plain', 'utf-8')

            # Adicione a mensagem personalizada ao objeto MIMEMultipart
            msg.attach(mensagem_final)

            try:
                server = smtplib.SMTP('smtp.office365.com', 587)
                server.starttls()
                server.login(email_de_envio, senha)
                server.sendmail(email_de_envio, email_destino, msg.as_string())
                server.quit()
                print(f"E-mail enviado para {email_destino}")
                enviados_com_sucesso.append(email_destino)

                # Adicione um atraso de 5 segundos
                time.sleep(20)

            except Exception as e:
                print(f"Erro ao enviar e-mail para {email_destino}: {str(e)}")
                nao_enviados.append(email_destino)
        else:
            print(f"E-mail ou mensagem inválidos para e-mail: {email_destino}")
            nao_enviados.append(email_destino)

    return enviados_com_sucesso, nao_enviados

# -----------------------------------------------------------------------------------------------------------------------------------------------

def validar_credenciais(email, senha):
    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(email, senha)
        server.quit()
        print("Credenciais válidas. Conexão bem-sucedida.")
        return True
    except smtplib.SMTPAuthenticationError:
        print("Erro de autenticação. Credenciais inválidas.")
        return False
    except Exception as e:
        print(f"Erro ao validar credenciais: {str(e)}")
        return False
    


# Aqui fica parte de banco de dados onde cria e salva dados dos envios
# -----------------------------------------------------------------------------------------------------------------------------------------------

def criar_janela_resumo(enviados, erros, connection):
    try:
        cursor = connection.cursor()

        # Criar a tabela de relatório se não existir
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS relatorio (
                id SERIAL PRIMARY KEY,
                email TEXT,
                campanha TEXT,
                status TEXT,
                data_envio TIMESTAMP,
                valor DECIMAL DEFAULT 0.0
            )
        """)

        connection.commit()

        for email in enviados:
            # Inserir dados na tabela de relatórios para e-mails enviados com sucesso
            cursor.execute("INSERT INTO relatorio (email, campanha, status, data_envio, valor) VALUES (%s, %s, %s, %s, %s)",
                           (email, 'enviado', datetime.now(), 0.01))

        for email in erros:
            # Inserir dados na tabela de relatórios para e-mails não enviados
            cursor.execute("INSERT INTO relatorio (email, campanha, status, data_envio, valor) VALUES (%s, %s, %s, %s, %s)",
                           (email, 'nao enviado', datetime.now(), 0.0))

        connection.commit()

        cursor.close()

        sg.popup("Relatório salvo no banco de dados com sucesso!")

    except (Exception, psycopg2.Error) as error:
        sg.popup_error(f"Erro ao salvar relatório no banco de dados: {str(error)}")


# -----------------------------------------------------------------------------------------------------------------------------------------------


def salvar_relatorio(connection, emails, status, valor, campanha, carteira):
    try:
        cursor = connection.cursor()

        # Inserir dados na tabela de relatórios
        for email in emails:
            cursor.execute("INSERT INTO relatorio (email, campanha, status, data_envio, valor, carteira) VALUES (%s, %s, %s, %s, %s, %s)",
                           (email, campanha, status, datetime.now(), valor, carteira))

        connection.commit()

        cursor.close()

        print("Relatório salvo no banco de dados com sucesso!")

    except (Exception, psycopg2.Error) as error:
        raise Exception(f"Erro ao salvar relatório no banco de dados: {str(error)}")


    
# -----------------------------------------------------------------------------------------------------------------------------------------------

def baixar_relatorio_por_periodo(connection, date_start, date_end):
    try:
        cursor = connection.cursor()

        # Consulta para obter emails enviados e não enviados no intervalo de datas especificado
        cursor.execute("""
            SELECT email, status, data_envio, campanha
            FROM relatorio
            WHERE data_envio::date BETWEEN %s AND %s
        """, (date_start, date_end))

        result = cursor.fetchall()

        # Separar os dados recuperados em listas distintas
        enviados = [row[0] + f" ({row[2].strftime('%Y-%m-%d %H:%M:%S')})" for row in result if row[1] == 'enviado']
        nao_enviados = [row[0] + f" ({row[2].strftime('%Y-%m-%d %H:%M:%S')})" for row in result if row[1] == 'nao enviado']

        cursor.close()

        return enviados, nao_enviados

    except (Exception, psycopg2.Error) as error:
        raise Exception(f"Erro ao baixar relatório por período: {str(error)}")
    



#------------------------------------------------------------------------------------------------------------------------------------------------------------------


def criar_tabela_relatorio(connection):
    try:
        cursor = connection.cursor()

        # Verificar se a tabela 'relatorio' já existe
        cursor.execute("SELECT EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = 'relatorio')")
        table_exists = cursor.fetchone()[0]

        if not table_exists:
          
            cursor.execute("""
                CREATE TABLE relatorio (
                    id SERIAL PRIMARY KEY,
                    carteira TEXT,
                    campanha TEXT,
                    email TEXT,
                    status TEXT,
                    data_envio TIMESTAMP,
                    valor DECIMAL DEFAULT 0.0
                )
            """)

            connection.commit()

            print("Tabela 'relatorio' criada com sucesso.")
        else:
            print("A tabela 'relatorio' já existe.")

        cursor.execute("SELECT EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = 'variaveis')")
        table_exists = cursor.fetchone()[0]

        if not table_exists:
            # Criar a tabela 'variaveis'
            cursor.execute("""
                CREATE TABLE variaveis (
                    id SERIAL PRIMARY KEY,
                    carteiras TEXT
                )
            """)
            print("Tabela 'variaveis' criada com sucesso.")
        else:
            print("A tabela 'variaveis' já existe.")

        connection.commit()

        cursor.close()

    except (Exception, psycopg2.Error) as error:
        raise Exception(f"Erro ao criar tabelas: {str(error)}")
    


#------------------------------------------------------------------------------------------------------------------------------------------------------



def registrar_carteira(connection, nome_carteira):
    try:
        cursor = connection.cursor()

        # Inserir a nova carteira na tabela 'variaveis'
        cursor.execute("INSERT INTO variaveis (carteiras) VALUES (%s) RETURNING id", (nome_carteira,))

        # Obter o ID retornado pela consulta RETURNING
        carteira_id = cursor.fetchone()[0]

    except (Exception, psycopg2.Error) as error:
        # Se ocorrer um erro, o rollback é executado para desfazer a transação
        connection.rollback()
        raise Exception(f"Erro ao registrar carteira: {str(error)}")

    finally:
        # O commit é feito no bloco finally para garantir que seja executado
        cursor.close()

    print(f"Carteira '{nome_carteira}' registrada com sucesso. ID: {carteira_id}")
    connection.commit()




#-------------------------------------------------------------------------------------------------------------------

def obter_carteiras(connection):
    try:
        cursor = connection.cursor()

        # Consulta ao banco de dados para obter a lista de carteiras
        cursor.execute("SELECT carteiras FROM variaveis")
        carteiras = cursor.fetchall()
        carteiras = [carteira[0] for carteira in carteiras]

        cursor.close()

        return carteiras

    except (Exception, psycopg2.Error) as error:
        raise Exception(f"Erro ao obter a lista de carteiras: {str(error)}")
