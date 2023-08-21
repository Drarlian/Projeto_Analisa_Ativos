import PlanilhaExcel.manipula_planilha_excel as manipula_excel


def analisar_pvp_excel(tipo_ativo: str, lista_ativos: list,  todos: bool = False) -> None:
    """
    Recebe uma lista com os ativos que serão atualizados.
    :param tipo_ativo: Tipo de ativo que será analisado | Possibilidades: 'fiis' ou 'acoes'
    :param lista_ativos: Lista com ativos | Formato: ['ativo1', 'ativo2']
    :param todos: Define se TODOS os ativos do tipo informado serão atualizados.
    :return: None
    """
    from openpyxl.styles import PatternFill

    if tipo_ativo == 'acoes':
        intervalo = 'A2:F'
        posicao_do_elemento_pvp = 4
    elif tipo_ativo == 'fiis':
        intervalo = 'I2:N'
        posicao_do_elemento_pvp = 3
    else:
        return None

    intervalo = intervalo + manipula_excel.descobrir_ultima_linha_planilha_excel(intervalo[0])  # Descobrindo o número da última linha preenchida.

    if todos:
        ativos = manipula_excel.pegar_dados_intervalo_planilha(intervalo)
        lista_ativos = [elemento[0] for elemento in ativos]

    indicador_positivo = 1.05

    planilha = manipula_excel.iniciar_planilha()
    aba_ativa = planilha.active

    cont = 0
    for celula in aba_ativa[intervalo]:
        if celula[0].value is None:
            break
        else:
            if celula[0].value == lista_ativos[cont]:
                try:
                    valor_celula = float(celula[posicao_do_elemento_pvp].value.replace(',', '.'))
                except AttributeError:
                    valor_celula = float(celula[posicao_do_elemento_pvp].value)
                # Se o indicador for negativo:
                if valor_celula >= indicador_positivo:
                    celula[posicao_do_elemento_pvp].fill = PatternFill(start_color='FFFF0000', end_color=None,
                                                                       fill_type='solid')
                # Se o indicador for positivo:
                else:
                    celula[posicao_do_elemento_pvp].fill = PatternFill(start_color='008000', end_color=None,
                                                                       fill_type='solid')

                if lista_ativos[cont] == lista_ativos[-1]:
                    break
                else:
                    cont += 1

    planilha.save('PlanilhaExcel\\Arquivo.xlsx')
    planilha.close()


if __name__ == '__main__':
    # analisar_pvp_excel('fiis', ['RBRY11', 'VGIR11'])
    # analisar_pvp_excel('fiis', ['DEVA11', 'HGLG11'])
    analisar_pvp_excel('fiis', lista_ativos=None, todos=True)

    # analisar_pvp_excel('acoes', ['TAEE4', 'BBDC4'])
    # analisar_pvp_excel('acoes', ['JBSS3', 'EKTR3'])
    analisar_pvp_excel('acoes', lista_ativos=None, todos=True)
