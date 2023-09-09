import PlanilhaExcel.manipula_planilha_excel as manipula_excel
import PlanilhaExcel.analisa_dados_excel as analisa_excel
import RaspagemDados.raspagem_dados as raspagem


def adicionar_acoes(lista_acoes: list, primeira_vez: bool = False) -> None:
    """
    Recebe uma lista de ações e adiciona os ativos da lista na planilha juntamente com seus indicadores.
    É possível passar o parâmetro "primeira_vez" para avisar que a parte de acoes da planilha está vazia e por isso
    os títulos devem ser adicionados.
    :param lista_acoes: Lista de ações.
    :param primeira_vez: Determina se a planilha contém ou não ações.
    :return: None
    """
    lista_acoes: list = list(dict.fromkeys(lista_acoes).keys())

    if primeira_vez:
        new_lista_acoes: list = lista_acoes.copy()
    else:
        new_lista_acoes, lista_acoes_existentes = verificar_ativo_existe('acoes', lista_acoes)

    if len(new_lista_acoes) == 0:
        return None

    if primeira_vez:
        lista_completa = raspagem.new_pegar_dados_ativo('acoes', new_lista_acoes, titulo=True)
    else:
        lista_completa = raspagem.new_pegar_dados_ativo('acoes', new_lista_acoes)

    manipula_excel.adicionar_dados_fim_coluna(lista_completa, 'A', 'F')

    if primeira_vez:
        lista_completa = lista_completa[1:]

    registrar_ativos_atualizados('acoes', new_lista_acoes)

    lista_pvp = [ativo[0] for ativo in lista_completa]
    analisa_excel.analisar_pvp_excel('acoes', lista_pvp)


def atualizar_acoes_todas() -> None:
    """
    Atualiza os indicadores de todas as ações presentes na planilha.
    :return: None
    """
    lista = manipula_excel.pegar_dados_intervalo_planilha('A2:A', ultima_linha=True)
    lista_ativos = []

    for ativo in lista:
        lista_ativos.append(ativo[0])

    lista_atualizada_formatada = raspagem.new_pegar_dados_ativo('acoes', lista_ativos)

    manipula_excel.adicionar_dados_intervalo_planilha(lista_atualizada_formatada, intervalo='A2:F', ultima_linha=True)

    registrar_ativos_atualizados('acoes', lista_ativos)

    lista_pvp = [ativo[0] for ativo in lista_atualizada_formatada]
    analisa_excel.analisar_pvp_excel('acoes', lista_pvp)


def atualizar_acoes_especificas(lista_acoes: list) -> None:
    """
    Atualiza os indicadores das ações especificadas presentes na planilha.
    :param lista_acoes: Lista de ações que devem ser atualizados.
    :return: None
    """
    lista_acoes: list = list(dict.fromkeys(lista_acoes).keys())

    _, lista_ativos_existentes = verificar_ativo_existe('acoes', lista_acoes)
    # -> Descarto os ativos que não estão presentes na planilha.

    if lista_ativos_existentes is None:  # -> Se não tiver ativos já presentes encerra a atualização.
        return None

    ativos_planilha: list = [ativo[0] for ativo in manipula_excel.pegar_dados_intervalo_planilha(intervalo='A2:A', ultima_linha=True)]

    dados = raspagem.new_pegar_dados_ativo('acoes', lista_ativos_existentes)

    for indice, ativo in enumerate(lista_ativos_existentes):
        posicao_elemento: int = ativos_planilha.index(ativo)  # -> Posição na planilha. (Com 2 posições a menos)
        manipula_excel.atualizar_dados_intervalo_planilha([dados[indice]], f'A{posicao_elemento+2}:F{posicao_elemento+2}')


    ativos_atualizados = [ativo[0] for ativo in dados]

    registrar_ativos_atualizados('acoes', ativos_atualizados)

    analisa_excel.analisar_pvp_excel('acoes', ativos_atualizados)


def adicionar_fiis(lista_fiis: list, primeira_vez: bool = False) -> None:
    """
    Recebe uma lista de fiis e adiciona os ativos da lista na planilha juntamente com seus indicadores.
    É possível passar o parâmetro "primeira_vez" para avisar que a parte de fiis da planilha está vazia e por isso
    os títulos devem ser adicionados.
    :param lista_fiis: Lista de ações.
    :param primeira_vez: Determina se a planilha contém ou não ações.
    :return: None
    """
    lista_fiis: list = list(dict.fromkeys(lista_fiis).keys())

    if primeira_vez:
        new_lista_fiis: list = lista_fiis.copy()
    else:
        new_lista_fiis, lista_acoes_existentes = verificar_ativo_existe('fiis', lista_fiis)

    if len(new_lista_fiis) == 0:
        return None

    if primeira_vez:
        lista_completa = raspagem.new_pegar_dados_ativo('fiis', new_lista_fiis, titulo=True)
    else:
        lista_completa = raspagem.new_pegar_dados_ativo('fiis', new_lista_fiis)

    manipula_excel.adicionar_dados_fim_coluna(lista_completa, 'I', 'N')

    if primeira_vez:
        lista_completa = lista_completa[1:]

    registrar_ativos_atualizados('fiis', new_lista_fiis)

    lista_pvp = [ativo[0] for ativo in lista_completa]
    analisa_excel.analisar_pvp_excel('fiis', lista_pvp)


def atualizar_fiis_todos():
    """
    Atualiza os indicadores de todos os fiis presentes na planilha.
    :return: None
    """
    lista = manipula_excel.pegar_dados_intervalo_planilha('I2:I', ultima_linha=True)
    lista_ativos = []

    for ativo in lista:
        lista_ativos.append(ativo[0])

    lista_atualizada_formatada = raspagem.new_pegar_dados_ativo('fiis', lista_ativos)

    manipula_excel.adicionar_dados_intervalo_planilha(lista_atualizada_formatada, intervalo='I2:N', ultima_linha=True)

    registrar_ativos_atualizados('fiis', lista_ativos)

    lista_pvp = [ativo[0] for ativo in lista_atualizada_formatada]
    analisa_excel.analisar_pvp_excel('fiis', lista_pvp)


def atualizar_fiis_especificos(lista_fiis: list) -> None:
    """
    Atualiza os indicadores dos fiis especificados presentes na planilha.
    :param lista_fiis: Lista de fiis que devem ser atualizados.
    :return: None
    """
    lista_fiis: list = list(dict.fromkeys(lista_fiis).keys())

    _, lista_ativos_existentes = verificar_ativo_existe('fiis', lista_fiis)
    # -> Descarto os ativos que não estão presentes na planilha.

    if lista_ativos_existentes is None:  # -> Se não tiver ativos já presentes encerra a atualização.
        return None

    ativos_planilha: list = [ativo[0] for ativo in
                             manipula_excel.pegar_dados_intervalo_planilha(intervalo='I2:I', ultima_linha=True)]

    dados = raspagem.new_pegar_dados_ativo('fiis', lista_ativos_existentes)

    for indice, ativo in enumerate(lista_ativos_existentes):
        posicao_elemento: int = ativos_planilha.index(ativo)  # -> Posição na planilha. (Com 2 posições a menos)
        manipula_excel.atualizar_dados_intervalo_planilha([dados[indice]],
                                                          f'I{posicao_elemento + 2}:N{posicao_elemento + 2}')

    ativos_atualizados = [ativo[0] for ativo in dados]

    registrar_ativos_atualizados('fiis', ativos_atualizados)

    analisa_excel.analisar_pvp_excel('fiis', ativos_atualizados)


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
        ativos_planilha = manipula_excel.pegar_dados_intervalo_planilha('A2:A', ultima_linha=True)
    elif tipo_ativo == 'fiis':
        ativos_planilha = manipula_excel.pegar_dados_intervalo_planilha('I2:I', ultima_linha=True)
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
        intervalo = 'T1:T3'
    else:
        raise TypeError('O tipo do ativo não existe.')

    data_atualizacao: datetime = f'{datetime.now().strftime("%d/%m/%Y")}'
    hora_atualizacao: datetime = f'{datetime.now().strftime("%H:%M:%S")}'

    manipula_excel.adicionar_dados_intervalo_planilha([[texto], [data_atualizacao], [hora_atualizacao]], intervalo)


def registrar_ativos_atualizados(tipo_ativo: str, lista_ativos: list) -> None:  # -> Talvez em um arquivo utilitários?
    """
    Escreve na planilha os ativos que foram atualizados durante a última modificação.
    :param tipo_ativo: Determina o tipo de ativo -> acoes | fiis
    :param lista_ativos: Lista de ativos que foram atualizados ou adicionados na planilha.
    :return: None
    """
    if tipo_ativo == 'acoes':
        texto: str = 'Ultimas Ações Atualizadas:'
        intervalo = f'Q6:Q{len(lista_ativos)+6}'
        intervalo_incompleto = 'Q6:Q'
    elif tipo_ativo == 'fiis':
        texto = 'Ultimos FII\'s Atualizados:'
        intervalo = f'T6:T{len(lista_ativos)+6}'
        intervalo_incompleto = 'T6:T'
    else:
        raise TypeError('O tipo do ativo não existe.')


    lista_formatada = [[x] for x in lista_ativos]
    lista_formatada.insert(0, [texto])

    if manipula_excel.pegar_dados_intervalo_planilha(intervalo_incompleto, ultima_linha=True):  # -> Verifico se já existe alguma atualização passada.
        manipula_excel.remover_dados_intervalo_planilha(intervalo_incompleto, ultima_linha=True)  # -> Removendo apenas os dados dos ativos atualizados.

    manipula_excel.adicionar_dados_intervalo_planilha(lista_formatada, intervalo)

    registrar_data_hora(tipo_ativo)


if __name__ == '__main__':
    # adicionar_acoes(['TAEE4', 'ITUB4', 'BBDC4', 'VALE3', 'JBSS3'])
    adicionar_acoes(['BBDC4'])
