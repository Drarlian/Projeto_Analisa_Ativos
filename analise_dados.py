from manipula_planilha import *


def analisar_pvp():
    """
    Recebe uma lista dos ativos presentes na planilha.
    """
    lista_ativos = pegar_dados_planilha()

    if lista_ativos[0][0] == 'Ativo':
        lista_ativos.pop(0)

    for c in range(len(lista_ativos)):
        if float(lista_ativos[c][4].replace(',', '.')) <= 1.05:
            cell = criar_celula(base=False, cores=(0, 1, 0))
        else:
            cell = criar_celula(base=False, cores=(1, 0, 0))

        atualizar_formatacao_planilha(base=False,
                                      request=criar_request(cell, base=False, linhas=(c+1, c+2), colunas=(4, 5)))


def criar_celula(base: bool = True, alinhamento: str = 'LEFT', cores: tuple = (1, 1, 1)):
    if base:
        cell_format = {
            'horizontalAlignment': 'LEFT',
            'backgroundColor': {
                'red': 1,
                'green': 1,
                'blue': 1
            }
        }
    else:
        cell_format = {
            'horizontalAlignment': alinhamento,
            'backgroundColor': {
                'red': cores[0],
                'green': cores[1],
                'blue': cores[2]
            }
        }

    return cell_format


def criar_request(cell_format, base: bool = True, linhas: tuple = (0, 99), colunas: tuple = (0, 99)):
    if base:
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
    else:
        request = {
            'repeatCell': {
                'range': {
                    'sheetId': 0,
                    'startRowIndex': linhas[0],
                    'endRowIndex': linhas[1],
                    'startColumnIndex': colunas[0],
                    'endColumnIndex': colunas[1]
                },
                'cell': {
                    'userEnteredFormat': cell_format
                },
                'fields': 'userEnteredFormat.horizontalAlignment,userEnteredFormat.backgroundColor'
            }
        }

    return request


def descobre_range(lista: list) -> str:
    """
    Descobre o range que vai ser usado pela lista recebida.
    """
    import string
    letras = list(string.ascii_uppercase)

    return f'Página1!A2:{letras[len(lista[0]) - 1]}'


if __name__ == '__main__':
    pass
    # analisar_pvp()
