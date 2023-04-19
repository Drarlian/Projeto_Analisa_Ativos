from manipula_planilha import *
from raspagem_dados import *
from analise_dados import *

if __name__ == '__main__':
    primeira_vez = True

    if primeira_vez:
        # Adiciona os ativos informados em uma planilha vazia.

        lista = ['PETR4', 'VALE3', 'ITUB4', 'BBAS3', 'SUZB3', 'JBSS3', 'RAIZ4', 'MRFG3', 'UNIP6', 'CMIN3', 'EKTR3']
        # lista = ['PETR4', 'VALE3', 'ITUB4']  # -> Lista Teste

        remover_dados_planilha()
        atualizar_formatacao_planilha()

        lista_completa = []

        for c in range(len(lista)):
            if c == 0:
                dados = pegar_dados_acao(lista[c], titulo=True)
            else:
                dados = pegar_dados_acao(lista[c])

            lista_completa += dados

        adicionar_dados_fim_planilha(lista_completa)

        analisar_pvp()
    else:
        # Atualiza os valores dos Ativos presentes na planilha.

        lista = pegar_dados_planilha(intervalo='Página1!A2:Z')
        lista_ativos = []

        for ativo in lista:
            lista_ativos.append(ativo[0])

        lista_completa = []

        for c in range(len(lista_ativos)):
            dados = pegar_dados_acao(lista_ativos[c])

            lista_completa += dados

        atualizar_dados_intervalo_planilha(lista_completa, 'Página1!A2')

        analisar_pvp()
