from bs4 import BeautifulSoup
from selenium import webdriver

def pegar_dados_ativo(tipo_ativo: str, nome_ativo: str, titulo: bool = False) -> list:
    # tipo_ativo = acoes | fiis
    if tipo_ativo not in ('acoes', 'fiis'):
        raise TypeError('O tipo do ativo não existe.')

    url = f'https://investidor10.com.br/{tipo_ativo}/{nome_ativo}/'

    edge_configs = webdriver.EdgeOptions()
    edge_configs.add_argument("--headless")  # -> Tornando o processo de pesquisa do site invisível.
    # edge_configs.add_argument("--disable-gpu")  # -> Desativar a aceleração de GPU.
    navegador = webdriver.Edge(options=edge_configs)
    # navegador = webdriver.Edge()
    navegador.get(url)

    navegador.implicitly_wait(5)

    soup = BeautifulSoup(navegador.page_source, 'html.parser')

    elementos_valores = soup.find_all('div', attrs={'class': '_card-body'})

    indicadores = [nome_ativo]

    for elemento in elementos_valores:
        span_tag = elemento.find('span')
        if span_tag:
            indicadores.append(span_tag.text)

    navegador.quit()

    if titulo:
        elementos_titulo = soup.find_all('div', {'class': '_card-header'})  # Procura pelo titulo apenas se for True
        titulos = ['Ativo']

        for elemento in elementos_titulo[:5]:
            titulo_span_tag = elemento.find('span')
            if titulo_span_tag:  # -> Caso encontre a tag span adiciona o texto dela na lista.
                titulos.append(titulo_span_tag.text)

        if tipo_ativo == 'fiis':
            titulos[1] = titulos[1].split(' ')[1]  # -> Preciso editar pois esse titulo vem assim: 'SNCI11 Cotação'
            titulos[2] = ' '.join(
                titulos[2].split(' ')[1:])  # -> Preciso editar, esse titulo vem assim: 'SNCI11 DY (12M)'

        lista_formatada = [titulos, indicadores]
    else:
        lista_formatada = [indicadores]

    return lista_formatada


def new_pegar_dados_ativo(tipo_ativo: str, lista_ativos: list, titulo: bool = False) -> list:
    """
    Procura pelas informações do(s) ativo(s) informado(s) de forma rápida e otimizada.
    Formato do Retorno:
    [['ativo1', 'informacao1', 'informacao2'], ['ativo2', 'outra_informacao1', 'outra_informacao2']]
    :param tipo_ativo: Tipo do ativo que será pesquisado. Opções: acoes | fiis
    :param lista_ativos: Lista contendo os ativos. A lista deve conter apenas ativos do mesmo tipo.
    :param titulo: Define se o cabeçalho das informações devem ser pegos.
    :return: Retorna uma lista contendo uma lista para cada ativo recebido.
    """
    # tipo_ativo = acoes | fiis
    if tipo_ativo not in ('acoes', 'fiis'):
        raise TypeError('O tipo do ativo não existe.')

    lista_completa = []

    lista_urls = []
    for nome_ativo in lista_ativos:
        lista_urls.append(f'https://investidor10.com.br/{tipo_ativo}/{nome_ativo}/')


    for indice, url in enumerate(lista_urls):
        if indice == 0:
            edge_configs = webdriver.EdgeOptions()
            edge_configs.add_argument("--headless")  # -> Tornando o processo de pesquisa do site invisível.
            # edge_configs.add_argument("--disable-gpu")  # -> Desativar a aceleração de GPU.
            navegador = webdriver.Edge(options=edge_configs)
            # navegador = webdriver.Edge()
            navegador.get(url)

            navegador.implicitly_wait(5)
        else:
            # Abrir uma nova aba
            navegador.execute_script("window.open('', '_blank');")

            # Mudar para a segunda aba
            navegador.switch_to.window(navegador.window_handles[indice])

            # Abrir o segundo link e pegar dados
            navegador.get(url)

        soup = BeautifulSoup(navegador.page_source, 'html.parser')

        elementos_valores = soup.find_all('div', attrs={'class': '_card-body'})

        indicadores = [lista_ativos[indice]]

        for elemento in elementos_valores:
            span_tag = elemento.find('span')
            if span_tag:
                indicadores.append(span_tag.text)

        if titulo and indice == 0:
            elementos_titulo = soup.find_all('div', {'class': '_card-header'})  # Procura pelo titulo apenas se for True
            titulos = ['Ativo']

            for elemento in elementos_titulo[:5]:
                titulo_span_tag = elemento.find('span')
                if titulo_span_tag:  # -> Caso encontre a tag span adiciona o texto dela na lista.
                    titulos.append(titulo_span_tag.text)

            if tipo_ativo == 'fiis':
                titulos[1] = titulos[1].split(' ')[1]  # -> Preciso editar pois esse titulo vem assim: 'SNCI11 Cotação'
                titulos[2] = ' '.join(
                    titulos[2].split(' ')[1:])  # -> Preciso editar, esse titulo vem assim: 'SNCI11 DY (12M)'

            lista_completa.append(titulos.copy())
            titulos.clear()

        lista_completa.append(indicadores.copy())
        indicadores.clear()


    navegador.quit()
    return lista_completa


if __name__ == '__main__':
    teste = pegar_dados_ativo('fiis', 'SNCI11', True)
    print(teste)

    teste2 = new_pegar_dados_ativo('fiis', ['SNCI11', 'CPTS11'], True)
    print(teste2)
