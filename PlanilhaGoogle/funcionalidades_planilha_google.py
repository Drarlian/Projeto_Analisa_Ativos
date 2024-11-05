import PlanilhaGoogle.manipula_planilha_google as manipula_google
import RaspagemDados.raspagem_dados as raspagem
import PlanilhaGoogle.analisa_dados_google as analisa_google


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
        manipula_google.remover_dados_planilha('Página1!A:G')

        manipula_google.atualizar_formatacao_planilha(False, analisa_google.criar_request(
            analisa_google.criar_celula(), False, linhas=(0, 99), colunas=(0, 7)))

        lista_ativos = lista_acoes.copy()
    else:
        lista_ativos, lista_acoes_existentes = verificar_ativo_existe(lista_acoes, 'acoes')

    if len(lista_ativos) == 0:
        return None

    lista_completa = []

    for c in range(len(lista_ativos)):
        if c == 0 and primeira_vez:
            dados = raspagem.pegar_dados_ativo('acoes', lista_ativos[c], titulo=True)
        else:
            dados = raspagem.pegar_dados_ativo('acoes', lista_ativos[c])

        lista_completa += dados

    manipula_google.adicionar_dados_fim_planilha(lista_completa, 'Página1!A:A')

    registrar_ativos_atualizados('acoes', lista_ativos)

    analisa_google.analisar_pvp('acoes')


def atualizar_acoes_todas() -> None:
    """
    Atualiza os indicadores de todas as ações presentes na planilha.
    :return: None
    """
    lista = manipula_google.pegar_dados_planilha(intervalo='Página1!A2:A')
    lista_ativos = []

    for ativo in lista:
        lista_ativos.append(ativo[0])

    lista_atualizada_formatada = []

    for c in range(len(lista_ativos)):
        dados = raspagem.pegar_dados_ativo('acoes', lista_ativos[c])

        lista_atualizada_formatada += dados

    manipula_google.atualizar_dados_intervalo_planilha(lista_atualizada_formatada, 'Página1!A2')

    registrar_ativos_atualizados('acoes', lista_ativos)

    analisa_google.analisar_pvp('acoes')


def atualizar_acoes_especificas(lista_acoes: list) -> None:
    """
    Atualiza os indicadores das ações especificadas presentes na planilha.
    :param lista_acoes: Lista de ações que devem ser atualizados.
    :return: None
    """
    _, lista_ativos = verificar_ativo_existe(lista_acoes, 'acoes')
    # -> Descarto os ativos que não estão presentes na planilha.

    if lista_ativos is None:  # -> Se não tiver ativos já presentes encerra a atualização.
        return None

    ativos_planilha = [ativo[0] for ativo in manipula_google.pegar_dados_planilha(intervalo='Página1!A2:A')]

    for ativo in lista_ativos:
        dado = raspagem.pegar_dados_ativo('acoes', ativo)
        posicao_elemento = ativos_planilha.index(ativo)
        manipula_google.atualizar_dados_intervalo_planilha(dado, f'Página1!A{posicao_elemento+2}')

    registrar_ativos_atualizados('acoes', lista_ativos)

    analisa_google.analisar_pvp('acoes')


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
        manipula_google.remover_dados_planilha('Página1!I:O')

        manipula_google.atualizar_formatacao_planilha(False, analisa_google.criar_request(
            analisa_google.criar_celula(), False, linhas=(0, 99), colunas=(8, 15)))
    else:
        lista_fiis, lista_fiis_existentes = verificar_ativo_existe(lista_fiis, 'fiis')

        if len(lista_fiis) == 0:
            return None

    lista_atualizada_formatada = []

    for c in range(len(lista_fiis)):
        if c == 0 and primeira_vez:
            dados = raspagem.pegar_dados_ativo('fiis', lista_fiis[c], titulo=True)
        else:
            dados = raspagem.pegar_dados_ativo('fiis', lista_fiis[c])

        lista_atualizada_formatada += dados

    manipula_google.adicionar_dados_fim_planilha(lista_atualizada_formatada, 'Página1!I:I')

    registrar_ativos_atualizados('fiis', lista_fiis)

    analisa_google.analisar_pvp('fiis')


def atualizar_fiis_todos() -> None:
    """
    Atualiza os indicadores de todos os fiis presentes na planilha.
    :return: None
    """
    lista = manipula_google.pegar_dados_planilha(intervalo='Página1!I2:I')
    lista_ativos = []

    for ativo in lista:
        lista_ativos.append(ativo[0])

    lista_atualizada_formatada = []

    for c in range(len(lista_ativos)):
        dados = raspagem.pegar_dados_ativo('fiis', lista_ativos[c])

        lista_atualizada_formatada += dados

    manipula_google.atualizar_dados_intervalo_planilha(lista_atualizada_formatada, 'Página1!I2')

    registrar_ativos_atualizados('fiis', lista_ativos)

    analisa_google.analisar_pvp('fiis')


def atualizar_fiis_especificos(lista_fiis: list) -> None:
    """
    Atualiza os indicadores dos fiis especificados presentes na planilha.
    :param lista_fiis: Lista de fiis que devem ser atualizados.
    :return: None
    """
    _, lista_ativos = verificar_ativo_existe(lista_fiis, 'fiis')
    # -> Descarto os ativos que não estão presentes na planilha.

    if lista_ativos is None:  # -> Se não tiver ativos já presentes encerra a atualização.
        return None

    ativos_planilha = [ativo[0] for ativo in manipula_google.pegar_dados_planilha(intervalo='Página1!I2:I')]

    for ativo in lista_ativos:  # -> Pego a posição do ativo na lista_ativos e atualizo APENAS o ativo dessa posição.
        dado = raspagem.pegar_dados_ativo('fiis', ativo)
        posicao_elemento = ativos_planilha.index(ativo)
        manipula_google.atualizar_dados_intervalo_planilha(dado, f'Página1!I{posicao_elemento+2}')

    registrar_ativos_atualizados('fiis', lista_ativos)

    analisa_google.analisar_pvp('fiis')


def verificar_ativo_existe(lista_ativos: list, tipo_ativo: str) -> tuple:
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
        ativos_planilha = manipula_google.pegar_dados_planilha(intervalo='Página1!A2:A')
    elif tipo_ativo == 'fiis':
        ativos_planilha = manipula_google.pegar_dados_planilha(intervalo='Página1!I2:I')
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


def registrar_data_hora(tipo_ativo: str) -> None:
    """
    Escreve na planilha a data e hora da última modificação.
    Formato: dd/mm/yyyy - hh:mm:ss
    :param tipo_ativo: Determina o tipo de ativo -> acoes | fiis
    :return: None
    """
    from datetime import datetime

    if tipo_ativo == 'acoes':
        texto: str = 'Ultima atualização - Ações:'
        intervalo = 'Página1!Q1:Q3'
    elif tipo_ativo == 'fiis':
        texto: str = 'Ultima atualização - FII\'s:'
        intervalo = 'Página1!S1:S3'
    else:
        raise TypeError('O tipo do ativo não existe.')

    data_atualizacao: datetime = f'{datetime.now().strftime("%d/%m/%Y")}'
    hora_atualizacao: datetime = f'{datetime.now().strftime("%H:%M:%S")}'

    manipula_google.atualizar_dados_intervalo_planilha([[texto], [data_atualizacao], [hora_atualizacao]], intervalo)


def registrar_ativos_atualizados(tipo_ativo: str, lista_ativos: list) -> None:
    """
    Escreve na planilha os ativos que foram atualizados durante a última modificação.
    :param tipo_ativo: Determina o tipo de ativo -> acoes | fiis
    :param lista_ativos: Lista de ativos foram atualizados ou adicionados na planilha.
    :return: None
    """
    if tipo_ativo == 'acoes':
        texto: str = 'Ultimas Ações Atualizadas:'
        intervalo = 'Página1!Q6:Q'
    elif tipo_ativo == 'fiis':
        texto = 'Ultimos FII\'s Atualizadas:'
        intervalo = 'Página1!S6:S'
    else:
        raise TypeError('O tipo do ativo não existe.')


    lista_formatada = [[x] for x in lista_ativos]
    lista_formatada.insert(0, [texto])

    if manipula_google.pegar_dados_planilha(intervalo):  # -> Verifico se já existe alguma atualização passada.
        manipula_google.remover_dados_planilha(intervalo.replace('6', '7'))  # -> Removendo apenas os dados dos ativos atualizados.

    manipula_google.atualizar_dados_intervalo_planilha(lista_formatada, intervalo)

    registrar_data_hora(tipo_ativo)


if __name__ == '__main__':
    # TESTES:
    # lista1 = ['DEVA11', 'CPTS11', 'RBVA11', 'HGLG11', 'ARRI11']
    # lista1 = ['SNCI11']  # -> Lista Teste
    # lista1 = ['NSLU11', 'KCRE11', 'XPML11']  # -> Lista Teste
    # lista1 = ['ARRI11', 'HSML11', 'NSLU11', 'HABT11', 'BRCO11', 'DEVA11', 'HFOF11', 'HGLG11', 'HGRE11', 'VIFI11',
    #           'VISC11', 'IRDM11', 'RBRY11', 'CPTS11', 'VGIR11', 'SNCI11', 'XPML11', 'RBVA11']
    lista1 = ['VINO11']
    # lista1 = ['SNCI11']
    # lista2 = ['PETR4', 'VALE3', 'ITUB4', 'BBAS3', 'SUZB3', 'JBSS3', 'RAIZ4', 'MRFG3', 'UNIP6', 'CMIN3', 'EKTR3',
    #           'TAEE4']
    lista2 = ['SUZB3', 'JBSS3']  # -> Lista Teste
    # adicionar_fiis(lista1)
    # adicionar_acoes(lista2)
    # atualizar_fiis_especificos(['LVBI11'])
    # atualizar_fiis_todos()
    # atualizar_acoes_especificas(['MGLU3'])
    # atualizar_acoes_todas()
    atualizar_acoes_especificas(['BBAS3', 'SUZB3', 'JBSS3', 'WEGE3', 'WEGE4', 'UNIP6'])
