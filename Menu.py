import PySimpleGUI as sg
import Backend
import pandas as pd
import psycopg2
from datetime import datetime



try:
    connection = psycopg2.connect(
        user="postgres",
        password="admin",
        host="localhost",
        port="5432",
        database="admin",
        options="-c client_encoding=utf-8"
    )

    Backend.criar_tabela_relatorio(connection)
    carteiras = Backend.obter_carteiras(connection)

except (Exception, psycopg2.Error) as error:
    sg.popup_error(f"Erro ao conectar ao banco de dados: {str(error)}")
    exit()


layout = [
    [
        sg.Column([
            [sg.Text('Envie seus Emails')],
            [sg.Text('Email de Envio:'), sg.Input(key='-EMAIL-')],
            [sg.Text('Senha:'), sg.Input(key='-SENHA-', password_char='*')],
            [sg.Text('Campanha:'), sg.Input(key='-CAMPANHA-')],
            [sg.Text('Assunto:'), sg.Input(key='-ASSUNTO-')],
            [sg.Text('Mensagem de Envio:'), sg.Multiline(key='-MENSAGEM-', size=(30, 5))],
            [sg.Text('Endereço do Excel:'), sg.Input(key='-ARQUIVO_EXCEL-')],
            [sg.Text('Selecione a Carteira:'), sg.Combo(values=carteiras, key='-CARTEIRA_SELECIONADA-')],
            [sg.Button('Enviar')],
            [sg.Button('Baixar Relatório por Data')],
            [sg.Button('Fechar')],
        ], element_justification='center', vertical_alignment='top'),
        sg.Column([
            [sg.Text('Cadastro Carteiras')],
            [sg.Text('Insira Nome da Carteira'), sg.Input(key='-CARTEIRA-')],
            [sg.Button('Cadastrar')]
        ], element_justification='center', vertical_alignment='top')
    ]
]

size = (800, 500)
window = sg.Window('Envio Certo', layout, size=size)


while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED or event == 'Fechar':
        break
    
    try:
        carteiras = Backend.obter_carteiras(connection)
        window['-CARTEIRA_SELECIONADA-'].update(values=carteiras)
    except Exception as error:
        sg.popup_error(f"Erro ao obter a lista de carteiras: {str(error)}")




    if event == 'Enviar':
        email = values['-EMAIL-']
        senha = values['-SENHA-']
        assunto = values['-ASSUNTO-']
        campanha = values['-CAMPANHA-']
        carteira = values['-CARTEIRA_SELECIONADA-']
        mensagem = values['-MENSAGEM-']
        arquivo_excel = values['-ARQUIVO_EXCEL-']

        if Backend.validar_credenciais(email, senha):
            enviados, erros = Backend.enviar_emails(email, senha, assunto, campanha, mensagem, arquivo_excel)

            carteira_selecionada = values['-CARTEIRA_SELECIONADA-']

            Backend.salvar_relatorio(connection, enviados, 'enviado', 0.01, campanha, carteira)
            Backend.salvar_relatorio(connection, erros, 'nao enviado', 0.0, campanha, carteira)

        else:
            sg.popup_error("Credenciais inválidas. Por favor, verifique o e-mail e a senha.")

    elif event == 'Baixar Relatório por Data':
        layout_data = [
            [sg.Text('Selecione o Período para o Relatório')],
            [sg.CalendarButton('Data Inicial', target='-DATA_INICIAL-', close_when_date_chosen=True),
             sg.Input(key='-DATA_INICIAL-', visible=False)],
            [sg.CalendarButton('Data Final', target='-DATA_FINAL-', close_when_date_chosen=True),
             sg.Input(key='-DATA_FINAL-', visible=False)],
            [sg.Button('Baixar Relatório')],
        ]

        window_data = sg.Window('Escolher Data', layout_data)
        event_data, values_data = window_data.read()

        if event_data == sg.WINDOW_CLOSED:
            window_data.close()
            continue

        if event_data == 'Baixar Relatório':
            date_start = values_data['-DATA_INICIAL-']
            date_end = values_data['-DATA_FINAL-']

            if not date_start or not date_end:
                sg.popup_error("Por favor, selecione um período antes de baixar o relatório.")
                continue

            try:
                date_start = datetime.strptime(date_start.split()[0], "%Y-%m-%d").date()
                date_end = datetime.strptime(date_end.split()[0], "%Y-%m-%d").date()

                enviados, nao_enviados = Backend.baixar_relatorio_por_periodo(connection, date_start, date_end)

            except Exception as e:
                sg.popup_error(f"Erro ao baixar relatório por período: {str(e)}")
                continue

            layout_resumo_data = [
                [sg.Text('E-mails Enviados com Sucesso:')],
                [sg.Listbox(enviados, size=(40, 10))],
                [sg.Text('E-mails com Erro:')],
                [sg.Listbox(nao_enviados, size=(40, 10))],
                [sg.Button('Download Excel'), sg.Button('Download TXT'), sg.Button('Fechar')],
            ]

            window_resumo_data = sg.Window('Resumo de Envio por Data', layout_resumo_data)

            event_resumo_data, _ = window_resumo_data.read()

            if event_resumo_data == sg.WINDOW_CLOSED or event_resumo_data == 'Fechar':
                window_resumo_data.close()
                continue

            if event_resumo_data == 'Download Excel':
                max_length = max(len(enviados), len(nao_enviados))
                enviados += [''] * (max_length - len(enviados))
                nao_enviados += [''] * (max_length - len(nao_enviados))

                # Criar um DataFrame com os dados
                df = pd.DataFrame({'E-mails Enviados': enviados, 'E-mails com Erro': nao_enviados})

                file_path = sg.popup_get_file('Salvar Relatório Excel', save_as=True, file_types=(('Arquivos Excel', '*.xlsx'),))

                if file_path:
                    try:
                        df.to_excel(file_path, index=False)
                        sg.popup(f"Relatório Excel salvo em: {file_path}")
                    except Exception as excel_error:
                        sg.popup_error(f"Erro ao salvar o relatório em Excel: {str(excel_error)}")

            elif event_resumo_data == 'Download TXT':
                file_path = sg.popup_get_file('Salvar Relatório TXT', save_as=True, file_types=(('Arquivos de Texto', '*.txt'),))

                if file_path:
                    content = f'E-mails Enviados com Sucesso:\n{", ".join(enviados)}\n\nE-mails com Erro:\n{", ".join(nao_enviados)}'
                    with open(file_path, 'w') as file:
                        file.write(content)
                    sg.popup(f"Relatório TXT salvo em: {file_path}")

    elif event == 'Cadastrar':
        nome_carteira = values['-CARTEIRA-']

        try:
            Backend.registrar_carteira(connection, nome_carteira)
            sg.popup("Carteira cadastrada com sucesso!")
        except Exception as e:
            sg.popup_error(f"Erro ao cadastrar carteira: {str(e)}")

window.close()

