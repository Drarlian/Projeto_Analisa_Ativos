from bs4 import BeautifulSoup
from selenium import webdriver

def pegar_dados_acao(nome_acao: str, titulo: bool = False) -> dict:
    url = f'https://investidor10.com.br/acoes/{nome_acao}/'

    navegador = webdriver.Edge()
    navegador.get(url)

    navegador.implicitly_wait(5)

    soup = BeautifulSoup(navegador.page_source, 'html.parser')

    elementos_titulo = soup.find_all('div', {'class': '_card-header'})
    elementos_valores = soup.find_all('div', attrs={'class': '_card-body'})

    # navegador.quit()

    titulos = ['Ativo']

    for elemento in elementos_titulo[:5]:
        titulo_span_tag = elemento.find('span')
        if titulo_span_tag:
            titulos.append(titulo_span_tag.text)


    indicadores = [nome_acao]

    for elemento in elementos_valores:
        span_tag = elemento.find('span')
        if span_tag:
            indicadores.append(span_tag.text)

    dados = {}

    for c in range(6):
        dados[titulos[c]] = indicadores[c]

    return formata_para_planilha(dados, titulo=titulo)


def formata_para_planilha(dicionario: dict, titulo: bool = False) -> list:
    lista = []
    posicao = 0

    if titulo:
        posicao = 1

        lista.append([])
        for chave in dicionario.keys():
            lista[0].append(chave)

    lista.append([])
    for valor in dicionario.values():
        lista[posicao].append(valor)

    return lista


if __name__ == '__main__':
    teste = pegar_dados_acao('TAEE4')
    print(teste)
