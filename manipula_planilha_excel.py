from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


planilha = load_workbook('Arquivo.xlsx')
aba_ativa = planilha.active


def pegar_dados_intervalo_planilha(intervalo: str) -> list:
    """
    Retorna os valores presentes no intervalo informado.
    Os valores são retornados dentro de uma lista.
    A lista retornada contém listas para cada linha do intervalo informado.
    Ex: [['Pessoa1', 46, 2500, 'Jogador', '987654321'], ['Pessoa2', 22, 8000, 'Streamer', '768948302']]
    :param intervalo: Intervalo da planilha.
    :return: Retorna uma lista contendo os valores.
    """
    try:
        valores: list = []
        valores_linha: list = []

        linha = int(intervalo.split(':')[0][-1])
        for celula in aba_ativa[intervalo]:
            for elemento in celula:
                if linha == elemento.row:
                    valores_linha.append(elemento.value)
                else:
                    linha += 1
                    valores_linha.append(elemento.value)
            valores.append(valores_linha.copy())
            valores_linha.clear()
    except:
        print('Error!')
    else:
        return valores



def atualizar_dados_intervalo_planilha(intervalo: str):
    pass

def adicionar_dados_fim_planilha(intervalo: str):
    pass


def adicionar_dados_intervalo_planilha(valores_adicionar: list, intervalo: str) -> None:
    """
    Adiciona os dados informados no intervalo especificado da planilha.
    :param valores_adicionar: Dados a serem adicionados.
    :param intervalo: Intervalo da planilha.
    :return: None
    """
    try:
        for celula in aba_ativa[intervalo]:
            cont = 0
            for elemento in celula:
                elemento.value = valores_adicionar[cont]
                cont += 1
    except:
        print('Error!')
    else:
        planilha.save('Arquivo.xlsx')
        planilha.close()


def remover_dados_intervalo_planilha(intervalo: str):
    pass


def atualizar_formatacao_planilha():
    pass



if __name__ == '__main__':
    # adicionar_dados_intervalo_planilha(['Pessoa1', 22, 8000, 'Streamer', '768948302'], 'A4:E4')
    print(pegar_dados_intervalo_planilha('A3:E4'))
