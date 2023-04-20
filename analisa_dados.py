from manipula_planilha import *


def analisar_pvp(tipo_ativo: str) -> None:
    """
    Pega a lista dos ativos informados presentes na planilha e analisa o valor do P/PV, de acordo com a analise altera
    o fundo do valor na planilha. (verde = bom | vermelho = ruim)
    :param tipo_ativo: Determina o tipo de ativo -> acoes | fiis
    :return: None
    """
    indicador_positivo = 1.05

    if tipo_ativo == 'acoes':
        lista_ativos = pegar_dados_planilha('Página1!A2:G')
        posicao_do_elemento_pvp = 4
        colunas = (4, 5)  # -> Começa na coluna 4 e termina na coluna 5.
        # -> (A primeira coluna não é alterada, apenas as colunas seguintes da primeira)
    elif tipo_ativo == 'fiis':
        lista_ativos = pegar_dados_planilha('Página1!I2:O')
        posicao_do_elemento_pvp = 3
        colunas = (11, 12)  # -> Começa na coluna 11 e termina na coluna 12.
        # -> (A primeira coluna não é alterada, apenas as colunas seguintes da primeira)
    else:
        return None

    if lista_ativos[0][0] == 'Ativo':  # -> Nunca deve acontecer.
        lista_ativos.pop(0)

    for c in range(len(lista_ativos)):
        if float(lista_ativos[c][posicao_do_elemento_pvp].replace(',', '.')) <= indicador_positivo:
            cell = criar_celula(base=False, cores=(0, 1, 0))
        else:
            cell = criar_celula(base=False, cores=(1, 0, 0))

        atualizar_formatacao_planilha(base=False,
                                      request=criar_request(cell, base=False, linhas=(c+1, c+2), colunas=colunas))


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
    analisar_pvp('acoes')
    analisar_pvp('fiis')
