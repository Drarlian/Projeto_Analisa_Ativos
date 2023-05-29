from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

with open('id_da_planilha.txt', 'r') as arquivo:
    id_da_planilha = arquivo.read()


def autenticar_acesso():
    # ETAPA DE AUTENTICAÇÃO COM O GOOGLE:
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'id_do_cliente.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds


def pegar_arquivo(creds):
    try:
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()  # -> Pega tod0 o arquivo.

        return sheet

    except HttpError as err:
        print(err)


def pegar_dados_planilha(intervalo: str = 'Página1!A:Z'):
    creds = autenticar_acesso()

    # LER OS DADOS DA PLANILHA: (GET)
    try:
        sheet = pegar_arquivo(creds)

        result = sheet.values().get(spreadsheetId=id_da_planilha,
                                    range=intervalo).execute()  # -> Leio o arquivo, passando as informações.

        valores = result.get('values', [])  # -> Obtendo os dados da planilha que foi pega anteriormente.

        return valores

    except HttpError as err:
        print(err)


def atualizar_dados_intervalo_planilha(valores_adicionar: list, intervalo: str):
    creds = autenticar_acesso()

    # ADICIONAR/EDITAR DADOS NA PLANILHA: (UPDATE)
    try:
        sheet = pegar_arquivo(creds)

        sheet.values().update(spreadsheetId=id_da_planilha,  # -> Atualizo o arquivo, passando as informações
                              range=intervalo, valueInputOption='USER_ENTERED',
                              body={'values': valores_adicionar}).execute()

    except HttpError as err:
        print(err)


def adicionar_dados_fim_planilha(valores_adicionar: list, intervalo: str = 'Página1!A:A'):
    creds = autenticar_acesso()

    # ETAPA DE MANIPULAÇÃO DA PLANILHA:
    try:
        sheet = pegar_arquivo(creds)

        sheet.values().append(spreadsheetId=id_da_planilha, range=intervalo,
                              valueInputOption='USER_ENTERED', body={'values': valores_adicionar}).execute()

    except HttpError as err:
        print(err)


def adicionar_dados_intervalo_planilha(valores_adicionar: list, intervalo: str):
    creds = autenticar_acesso()

    # ETAPA DE MANIPULAÇÃO DA PLANILHA:
    try:
        sheet = pegar_arquivo(creds)

        sheet.values().append(spreadsheetId=id_da_planilha, range=intervalo,
                              valueInputOption='USER_ENTERED', body={'values': valores_adicionar}).execute()

    except HttpError as err:
        print(err)


def remover_dados_planilha(intervalo: str = 'Página1!A1:Z99'):
    creds = autenticar_acesso()

    # ETAPA DE MANIPULAÇÃO DA PLANILHA:
    try:
        sheet = pegar_arquivo(creds)

        sheet.values().clear(spreadsheetId=id_da_planilha, range=intervalo, body={}).execute()

    except HttpError as err:
        print(err)


def atualizar_formatacao_planilha(base: bool = True, request: bool = False):
    """
    formatacao=False -> Se a formatação for False, cell_format e request NÃO devem ser informados,
    pois os dados padrão da função SERÃO usados.
    formatacao=True -> Se a formatação for True, cell_format e request DEVEM ser informados,
    pois os dados padrão da função NÃO serão usados.
    """
    creds = autenticar_acesso()

    if base:
        cell_format = {
            'horizontalAlignment': 'LEFT',
            'backgroundColor': {
                'red': 1,
                'green': 1,
                'blue': 1
            }
        }

        request = {
            'repeatCell': {
                'range': {
                    'sheetId': 0,
                    'startRowIndex': 0,
                    'endRowIndex': 99,
                    'startColumnIndex': 0,
                    'endColumnIndex': 99
                },
                'cell': {
                    'userEnteredFormat': cell_format
                },
                'fields': 'userEnteredFormat.horizontalAlignment,userEnteredFormat.backgroundColor'
            }
        }

    # ETAPA DE MANIPULAÇÃO DA PLANILHA:
    try:
        sheet = pegar_arquivo(creds)

        sheet.batchUpdate(spreadsheetId=id_da_planilha, body={'requests': [request]}).execute()

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    pass
    # atualizar_dados_intervalo_planilha(lista_completa, 'Página1!A3')
    # atualizar_formatacao_planilha()
    # print(pegar_dados_planilha(intervalo='Página1!A2:F'))
