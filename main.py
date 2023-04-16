from manipula_planilha import *
from raspagem_dados import *

if __name__ == '__main__':
    dados = pegar_dados_acao('PETR4')
    adicionar_dados_fim_planilha(dados)
