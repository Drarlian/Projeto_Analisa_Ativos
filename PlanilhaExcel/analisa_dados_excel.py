import PlanilhaExcel.manipula_planilha_excel as manipula_excel


def analisar_pvp_excel(tipo_ativo: str, lista_ativos: list,  todos: bool = False) -> None:
    """
    Recebe uma lista com os ativos que terão seu P/VP analisado.
    :param tipo_ativo: Tipo de ativo que será analisado | Possibilidades: 'fiis' ou 'acoes'
    :param lista_ativos: Lista com ativos | Formato: ['ativo1', 'ativo2']
    :param todos: Define se TODOS os ativos do tipo informado serão atualizados.
    :return: None
    """
    from openpyxl.styles import PatternFill

    if tipo_ativo == 'acoes':
        intervalo: str = 'A2:F'
        posicao_do_elemento_pvp: int = 4
    elif tipo_ativo == 'fiis':
        intervalo: str = 'I2:N'
        posicao_do_elemento_pvp: int = 3
    else:
        return None

    # Descobrindo o número da última linha preenchida:
    ultima_linha: str = manipula_excel.descobrir_linha_vazia_planilha_excel(intervalo[0])
    intervalo: str = intervalo + ultima_linha

    if todos:
        ativos = manipula_excel.pegar_dados_intervalo_planilha(intervalo)
        lista_ativos: list = [elemento[0] for elemento in ativos]

    valor_maximo_aceitavel: float = 1.05

    planilha = manipula_excel.iniciar_planilha()
    aba_ativa = planilha.active

    try:
        for celula in aba_ativa[intervalo]:
            # Se o ativo atual da planilha estiver dentro da lista_ativos, ele analisa o ativo:
            if celula[0].value in lista_ativos:
                try:
                    valor_celula: float = float(celula[posicao_do_elemento_pvp].value.replace(',', '.'))
                except AttributeError:
                    valor_celula: float = float(celula[posicao_do_elemento_pvp].value)
                # Se o valor do P/VP estiver acima do valor máximo aceitável: (vermelho)
                if valor_celula >= valor_maximo_aceitavel:
                    celula[posicao_do_elemento_pvp].fill = PatternFill(start_color='FFFF0000', end_color=None,
                                                                       fill_type='solid')
                # Se o valor do P/VP estiver abaixo do valor máximo aceitável: (verde)
                else:
                    celula[posicao_do_elemento_pvp].fill = PatternFill(start_color='008000', end_color=None,
                                                                       fill_type='solid')

                lista_ativos.remove(celula[0].value)  # -> Remove o ativo que foi analisado da lista_ativos.
                if len(lista_ativos) == 0:  # -> Se não existem mais ativos para serem analisados, encerra o loop.
                    break
    except:
        print('Error!')
    else:
        planilha.save('PlanilhaExcel\\Arquivo.xlsx')
    finally:
        planilha.close()


def analisa_pl_excel(lista_acoes: list, todos: bool = False) -> None:
    """
    Recebe uma lista com as AÇÕES que terão seu P/L analisados.
    :param lista_acoes: Lista com ativos | Formato: ['ativo1', 'ativo2']
    :param todos: Define se TODOS os ativos do tipo informado serão atualizados.
    :return: None
    """
    from openpyxl.styles import PatternFill

    intervalo: str = 'A2:F'
    posicao_do_elemento_pl: int = 3

    # Descobrindo o número da última linha preenchida:
    ultima_linha: str = manipula_excel.descobrir_linha_vazia_planilha_excel(intervalo[0])
    intervalo: str = intervalo + ultima_linha

    if todos:
        outro_intervalo: str = 'A2:A' + ultima_linha  # -> Pode ser assim pois sempre vai ser uma AÇÃO.
        ativos: list = manipula_excel.pegar_dados_intervalo_planilha(outro_intervalo)  # -> Pegando apenas o nome dos ativos.
        lista_acoes: list = [elemento[0] for elemento in ativos]

    valor_maximo_aceitavel: float = 10.0

    planilha = manipula_excel.iniciar_planilha()
    aba_ativa = planilha.active

    try:
        for celula in aba_ativa[intervalo]:
            # Se o ativo atual da planilha estiver dentro da lista_ativos, ele analisa o ativo:
            if celula[0].value in lista_acoes:
                try:
                    valor_celula: float = float(celula[posicao_do_elemento_pl].value.replace(',', '.'))
                except AttributeError:
                    valor_celula: float = float(celula[posicao_do_elemento_pl].value)
                # Se o valor do P/L estiver acima do valor máximo aceitável: (vermelho)
                if valor_celula >= valor_maximo_aceitavel:
                    celula[posicao_do_elemento_pl].fill = PatternFill(start_color='FFFF0000', end_color=None,
                                                                      fill_type='solid')
                # Se o valor do P/L for negativo: (amarelo)
                elif valor_celula < 0:
                    celula[posicao_do_elemento_pl].fill = PatternFill(start_color='FFFF00', end_color=None,
                                                                      fill_type='solid')
                # Se o valor do P/L estiver abaixo do valor máximo aceitável: (verde)
                else:
                    celula[posicao_do_elemento_pl].fill = PatternFill(start_color='008000', end_color=None,
                                                                      fill_type='solid')

                lista_acoes.remove(celula[0].value)  # -> Remove o ativo que foi analisado da lista_ativos.
                if len(lista_acoes) == 0:  # -> Se não existem mais ativos para serem analisados, encerra o loop.
                    break
    except:
        print('Error!')
    else:
        planilha.save('PlanilhaExcel\\Arquivo.xlsx')
    finally:
        planilha.close()


if __name__ == '__main__':
    # analisar_pvp_excel('fiis', ['RBRY11', 'VGIR11'])
    # analisar_pvp_excel('fiis', ['DEVA11', 'HGLG11'])
    analisar_pvp_excel('fiis', lista_ativos=None, todos=True)

    # analisar_pvp_excel('acoes', ['TAEE4', 'BBDC4'])
    # analisar_pvp_excel('acoes', ['JBSS3', 'EKTR3'])
    analisar_pvp_excel('acoes', lista_ativos=None, todos=True)
