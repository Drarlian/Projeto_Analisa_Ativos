from manipula_planilha import *
from raspagem_dados import *
from analisa_dados import *  # -> Importando o "manipula_planilha" dentro de "analise_dados".


def pegar_acoes(lista_acoes: list, primeira_vez: bool = False) -> None:
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

    lista_completa = []


    for c in range(len(lista_acoes)):
        if c == 0 and primeira_vez:
            dados = pegar_dados_ativo('acoes', lista_acoes[c], titulo=True)
        else:
            dados = pegar_dados_ativo('acoes', lista_acoes[c])

        lista_completa += dados

    adicionar_dados_fim_planilha(lista_completa, 'Página1!A:A')

    analisar_pvp('acoes')


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


def pegar_fiis(lista_fiis: list, primeira_vez: bool = False) -> None:
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

    lista_completa = []

    for c in range(len(lista_fiis)):
        if c == 0 and primeira_vez:
            dados = pegar_dados_ativo('fiis', lista_fiis[c], titulo=True)
        else:
            dados = pegar_dados_ativo('fiis', lista_fiis[c])

        lista_completa += dados

    adicionar_dados_fim_planilha(lista_completa, 'Página1!I:I')

    analisar_pvp('fiis')


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


if __name__ == '__main__':
    # TESTES:
    lista1 = ['DEVA11', 'CPTS11', 'RBVA11', 'HGLG11']
    # lista1 = ['DEVA11', 'CPTS11']  # -> Lista Teste
    lista2 = ['PETR4', 'VALE3', 'ITUB4', 'BBAS3', 'SUZB3', 'JBSS3', 'RAIZ4', 'MRFG3', 'UNIP6', 'CMIN3', 'EKTR3']
    # lista2 = ['PETR4', 'VALE3']  # -> Lista Teste
    pegar_fiis(lista1)
    pegar_acoes(lista2)
    # atualizar_fiis()
    # atualizar_acoes()
