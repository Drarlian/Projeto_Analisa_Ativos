from manipula_planilha_excel import *
from raspagem_dados import *
from analisa_dados_excel import *  # -> Importando o "manipula_planilha" dentro de "analise_dados".

def adicionar_acoes(lista_acoes: list, primeira_vez: bool = False) -> None:
    """
    Recebe uma lista de ações e adiciona os ativos da lista na planilha juntamente com seus indicadores.
    É possível passar o parâmetro "primeira_vez" para avisar que a parte de acoes da planilha está vazia e por isso
    os títulos devem ser adicionados.
    :param lista_acoes: Lista de ações.
    :param primeira_vez: Determina se a planilha contém ou não ações.
    :return: None
    """
    # Pensando em remover o fator primeira_vez
    if primeira_vez:
        pass
    else:
        lista_acoes, lista_acoes_existentes = verificar_ativo_existe('acoes', lista_acoes)

        if len(lista_acoes) == 0:
            return None

    lista_completa = []

    for c, _ in enumerate(lista_acoes):
        if c == 0 and primeira_vez:
            dados = pegar_dados_ativo('acoes', lista_acoes[c], titulo=True)
        else:
            dados = pegar_dados_ativo('acoes', lista_acoes[c])

        lista_completa += dados

    adicionar_dados_fim_coluna(lista_completa, 'A')

    lista_pvp = [ativo[0] for ativo in lista_completa]
    analisar_pvp_excel('acoes', lista_pvp)

    registrar_ativos_atualizados('acoes', lista_acoes)


def atualizar_acoes_todas() -> None:
    """
    Atualiza os indicadores de todas as ações presentes na planilha.
    :return: None
    """
    lista = pegar_dados_intervalo_planilha('A2:A', ultima_linha=True)
    lista_ativos = []

    for ativo in lista:
        lista_ativos.append(ativo[0])

    lista_atualizada_formatada = []

    for c in range(len(lista_ativos)):
        dados = pegar_dados_ativo('acoes', lista_ativos[c])

        lista_atualizada_formatada += dados

    adicionar_dados_intervalo_planilha(lista_atualizada_formatada, intervalo='A2:F', ultima_linha=True)

    lista_pvp = [ativo[0] for ativo in lista_atualizada_formatada]
    analisar_pvp_excel('acoes', lista_pvp)

    registrar_ativos_atualizados('acoes', lista_ativos)


def atualizar_acoes_especificas(lista_acoes: list) -> None:
    """
    Atualiza os indicadores das ações especificadas presentes na planilha.
    :param lista_acoes: Lista de ações que devem ser atualizados.
    :return: None
    """
    _, lista_ativos_existentes = verificar_ativo_existe('acoes', lista_acoes)
    # -> Descarto os ativos que não estão presentes na planilha.

    if lista_ativos_existentes is None:  # -> Se não tiver ativos já presentes encerra a atualização.
        return None

    ativos_planilha: list = [ativo[0] for ativo in pegar_dados_intervalo_planilha(intervalo='A2:A', ultima_linha=True)]

    for ativo in lista_ativos_existentes:
        dado = pegar_dados_ativo('acoes', ativo)
        posicao_elemento: int = ativos_planilha.index(ativo)
        atualizar_dados_intervalo_planilha(dado, f'A{posicao_elemento+2}:F{posicao_elemento+2}')

    analisar_pvp_excel('acoes', lista_ativos_existentes)
    registrar_ativos_atualizados('acoes', lista_ativos_existentes)


def adicionar_fiis():
    """
    Recebe uma lista de fiis e adiciona os ativos da lista na planilha juntamente com seus indicadores.
    É possível passar o parâmetro "primeira_vez" para avisar que a parte de fiis da planilha está vazia e por isso
    os títulos devem ser adicionados.
    :param lista_fiis: Lista de ações.
    :param primeira_vez: Determina se a planilha contém ou não ações.
    :return: None
    """
    pass


def atualizar_fiis_todos():
    """
    Atualiza os indicadores de todos os fiis presentes na planilha.
    :return: None
    """
    pass


def atualizar_fiis_especificos():
    """
    Atualiza os indicadores dos fiis especificados presentes na planilha.
    :param lista_fiis: Lista de fiis que devem ser atualizados.
    :return: None
    """
    pass


def verificar_ativo_existe(tipo_ativo: str, lista_ativos: list) -> tuple:  # -> Talvez em um arquivo utilitários?
    """
    Verifica se algum dos ativos informados já está presente na planilha, caso esteja remove esse ativo da lista_ativos.
    Coloca os ativos JÁ presentes na planilha na lista "ativos_existentes".
    Coloca os ativos NÃO presentes na planilha na lista "ativos_nao_existentes".
    :param lista_ativos: Lista de ativos vinda do usuário.
    :param tipo_ativo: Tipo de ativo que será tratado.
    :return: Retorna duas listas, uma contendo os ativos NÃO presentes na planilha e outra com os ativos JÁ presentes
    na planilha.
    """
    if tipo_ativo == 'acoes':
        ativos_planilha = pegar_dados_intervalo_planilha('A2:A', ultima_linha=True)
    elif tipo_ativo == 'fiis':
        ativos_planilha = pegar_dados_intervalo_planilha('I2:I', ultima_linha=True)
    else:
        raise TypeError('O tipo do ativo não existe.')

    ativos_planilha_formatado = []
    for lista in ativos_planilha:
        ativos_planilha_formatado.extend(lista)  # -> Pegando cada ativo dentro das listas de ativos_planilha.
        # ativos_planilha_formatado += lista  # -> Mesma coisa que a linha de cima, porém usando concatenação de lista.

    ativos_nao_existentes = lista_ativos.copy()  # -> Aqui vão ficar os ativos REALMENTE NOVOS para a planilha.
    ativos_existentes = []  # -> Aqui vão ficar os ativos que JÁ EXISTEM na planilha.

    for elemento in lista_ativos:
        if elemento in ativos_planilha_formatado:
            ativos_existentes.append(elemento)  # -> Recebendo os ativos que já estão na planilha.
            ativos_nao_existentes.remove(elemento)  # -> Removendo os ativos que já estão na planilha.

    if len(ativos_existentes) == 0:
        ativos_existentes = None

    return ativos_nao_existentes, ativos_existentes



def registrar_data_hora(tipo_ativo: str) -> None:  # -> Talvez em um arquivo utilitários?
    """
    Escreve na planilha a data e hora da última modificação.
    Formato: dd/mm/yyyy - hh:mm:ss
    :param tipo_ativo: Determina o tipo de ativo -> acoes | fiis
    :return: None
    """
    from datetime import datetime

    if tipo_ativo == 'acoes':
        texto: str = 'Ultima atualização - Ações:'
        intervalo = 'Q1:Q3'
    elif tipo_ativo == 'fiis':
        texto: str = 'Ultima atualização - FII\'s:'
        intervalo = 'S1:S3'
    else:
        raise TypeError('O tipo do ativo não existe.')

    data_atualizacao: datetime = f'{datetime.now().strftime("%d/%m/%Y")}'
    hora_atualizacao: datetime = f'{datetime.now().strftime("%H:%M:%S")}'

    adicionar_dados_intervalo_planilha([[texto], [data_atualizacao], [hora_atualizacao]], intervalo)


def registrar_ativos_atualizados(tipo_ativo: str, lista_ativos: list) -> None:  # -> Talvez em um arquivo utilitários?
    """
    Escreve na planilha os ativos que foram atualizados durante a última modificação.
    :param tipo_ativo: Determina o tipo de ativo -> acoes | fiis
    :param lista_ativos: Lista de ativos foram atualizados ou adicionados na planilha.
    :return: None
    """
    if tipo_ativo == 'acoes':
        texto: str = 'Ultimas Ações Atualizadas:'
        intervalo = 'Q6:Q' + descobrir_ultima_linha_planilha_excel('Q')
    elif tipo_ativo == 'fiis':
        texto = 'Ultimos FII\'s Atualizadas:'
        intervalo = 'S6:S' + descobrir_ultima_linha_planilha_excel('S')
    else:
        raise TypeError('O tipo do ativo não existe.')


    lista_formatada = [[x] for x in lista_ativos]
    lista_formatada.insert(0, [texto])

    if pegar_dados_intervalo_planilha(intervalo):  # -> Verifico se já existe alguma atualização passada.
        remover_dados_intervalo_planilha(intervalo)  # -> Removendo apenas os dados dos ativos atualizados.

    adicionar_dados_intervalo_planilha(lista_formatada, intervalo)

    registrar_data_hora(tipo_ativo)


if __name__ == '__main__':
    # adicionar_acoes(['TAEE4', 'ITUB4', 'BBDC4', 'VALE3', 'JBSS3'])
    adicionar_acoes(['BBDC4'])
