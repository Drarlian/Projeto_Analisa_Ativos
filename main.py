from manipula_planilha import *
from raspagem_dados import *
from analisa_dados import *  # -> Importando o "manipula_planilha" dentro de "analise_dados".


def adicionar_acoes(lista_acoes: list, primeira_vez: bool = False) -> None:
    """
    Recebe uma lista de ações e adiciona os ativos da lista na planilha juntamente com seus indicadores.
    É possível passar o parâmetro "primeira_vez" para avisar que a parte de acoes da planilha está vazia e por isso
    os títulos devem ser adicionados.
    :param lista_acoes: Lista de ações.
    :param primeira_vez: Determina se a planilha contém ou não ações.
    :return: None
    """
    if primeira_vez:
        remover_dados_planilha('Página1!A:G')

        atualizar_formatacao_planilha(False, criar_request(criar_celula(), False, linhas=(0, 99), colunas=(0, 7)))
    else:
        lista_acoes, lista_acoes_existentes = verificar_ativo_existe(lista_acoes, 'acoes')

        if len(lista_acoes) == 0:
            return None


    lista_completa = []

    for c in range(len(lista_acoes)):
        if c == 0 and primeira_vez:
            dados = pegar_dados_ativo('acoes', lista_acoes[c], titulo=True)
        else:
            dados = pegar_dados_ativo('acoes', lista_acoes[c])

        lista_completa += dados

    adicionar_dados_fim_planilha(lista_completa, 'Página1!A:A')

    analisar_pvp('acoes')
    registrar_data_hora()


def atualizar_acoes() -> None:
    """
    Atualiza os indicadores das ações presentes na planilha.
    :return: None
    """
    lista = pegar_dados_planilha(intervalo='Página1!A2:G')
    lista_ativos = []

    for ativo in lista:
        lista_ativos.append(ativo[0])

    lista_completa = []

    for c in range(len(lista_ativos)):
        dados = pegar_dados_ativo('acoes', lista_ativos[c])

        lista_completa += dados

    atualizar_dados_intervalo_planilha(lista_completa, 'Página1!A2')

    analisar_pvp('acoes')
    registrar_data_hora()


def adicionar_fiis(lista_fiis: list, primeira_vez: bool = False) -> None:
    """
    Recebe uma lista de fiis e adiciona os ativos da lista na planilha juntamente com seus indicadores.
    É possível passar o parâmetro "primeira_vez" para avisar que a parte de fiis da planilha está vazia e por isso
    os títulos devem ser adicionados.
    :param lista_fiis: Lista de ações.
    :param primeira_vez: Determina se a planilha contém ou não ações.
    :return: None
    """
    if primeira_vez:
        remover_dados_planilha('Página1!I:O')

        atualizar_formatacao_planilha(False, criar_request(criar_celula(), False, linhas=(0, 99), colunas=(8, 15)))
    else:
        lista_fiis, lista_fiis_existentes = verificar_ativo_existe(lista_fiis, 'fiis')

        if len(lista_fiis) == 0:
            return None

    lista_completa = []

    for c in range(len(lista_fiis)):
        if c == 0 and primeira_vez:
            dados = pegar_dados_ativo('fiis', lista_fiis[c], titulo=True)
        else:
            dados = pegar_dados_ativo('fiis', lista_fiis[c])

        lista_completa += dados

    adicionar_dados_fim_planilha(lista_completa, 'Página1!I:I')

    analisar_pvp('fiis')
    registrar_data_hora()


def atualizar_fiis() -> None:
    """
    Atualiza os indicadores dos fiis presentes na planilha.
    :return: None
    """
    lista = pegar_dados_planilha(intervalo='Página1!I2:O')
    lista_ativos = []

    for ativo in lista:
        lista_ativos.append(ativo[0])

    lista_completa = []

    for c in range(len(lista_ativos)):
        dados = pegar_dados_ativo('fiis', lista_ativos[c])

        lista_completa += dados

    atualizar_dados_intervalo_planilha(lista_completa, 'Página1!I2')

    analisar_pvp('fiis')
    registrar_data_hora()


def verificar_ativo_existe(lista_ativos: list, tipo_ativo: str) -> tuple:
    """
    Verifica se algum dos ativos informados já está presente na planilha, caso esteja remove esse ativo da lista_acoes.
    Coloca os ativos JÁ presentes na planilha na lista "ativos_existentes".
    Coloca os ativos NÃO presentes na planilha na lista "ativos_nao_existentes".
    :param lista_ativos: Lista de ativos vinda do usuário.
    :param tipo_ativo: Tipo de ativo que será tratado.
    :return: Retorna duas listas, uma contendo os ativos NÃO presentes na planilha e outra com os ativos JÁ presentes
    na planilha.
    """
    if tipo_ativo == 'acoes':
        ativos_planilha = pegar_dados_planilha(intervalo='Página1!A2:A')
    elif tipo_ativo == 'fiis':
        ativos_planilha = pegar_dados_planilha(intervalo='Página1!I2:I')
    else:
        # return None
        raise TypeError('O tipo do ativo não existe.')

    ativos_planilha_formatado = []
    for lista in ativos_planilha:
        ativos_planilha_formatado.extend(lista)  # -> Pegando cada ativo dentro das listas de ativos_planilha.
        # ativos_planilha_formatado += lista  # -> Mesma coisa que a linha de cima, porém usando concatenação de lista.

    ativos_nao_existentes = lista_ativos.copy()  # -> Aqui vão ficar os ativos REALMENTE NOVOS para a planilha.
    ativos_existentes = []  # -> Aqui vão ficar os ativos que JÁ EXISTEM na planilha e devem ser atualizados.

    for elemento in lista_ativos:
        if elemento in ativos_planilha_formatado:
            ativos_existentes.append(elemento)  # -> Recebendo os ativos que já estão na planilha.
            ativos_nao_existentes.remove(elemento)  # -> Removendo os ativos que já estão na planilha.

    if len(ativos_existentes) == 0:
        ativos_existentes = None

    return ativos_nao_existentes, ativos_existentes


def registrar_data_hora() -> None:
    from datetime import datetime
    """
    Escreve na planilha a data e hora da última modificação.
    Formato: dd/mm/yyyy - hh:mm:ss
    :return: None
    """
    texto: str = 'Ultima atualização:'
    data_atualizacao: datetime = f'{datetime.now().strftime("%d/%m/%Y")}'
    hora_atualizacao: datetime = f'{datetime.now().strftime("%H:%M:%S")}'

    if not pegar_dados_planilha('Página1!Q1:Q3'):  # -> Caso não tenha nenhuma data_hora registrada.
        adicionar_dados_fim_planilha([[texto], [data_atualizacao], [hora_atualizacao]], 'Página1!Q1:Q3')
    else:  # -> Caso tenha alguma data_hora registrada.
        atualizar_dados_intervalo_planilha([[data_atualizacao], [hora_atualizacao]], 'Página1!Q2:Q3')


if __name__ == '__main__':
    # TESTES:
    lista1 = ['DEVA11', 'CPTS11', 'RBVA11', 'HGLG11', 'ARRI11']
    # lista1 = ['DEVA11', 'CPTS11']  # -> Lista Teste
    # lista2 = ['PETR4', 'VALE3', 'ITUB4', 'BBAS3', 'SUZB3', 'JBSS3', 'RAIZ4', 'MRFG3', 'UNIP6', 'CMIN3', 'EKTR3',
    #           'TAEE4']
    # lista2 = ['PETR4', 'VALE3']  # -> Lista Teste
    adicionar_fiis(lista1)
    # adicionar_acoes(lista2)
    # atualizar_fiis()
    # atualizar_acoes()
