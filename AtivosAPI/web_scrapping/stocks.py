from bs4 import BeautifulSoup
from selenium import webdriver


def get_stocks_from_web(tipo_ativo: str, lista_ativos: list, titulo: bool = False) -> list:
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

    if tipo_ativo == 'acoes':
        for indice, url in enumerate(lista_urls):
            if indice == 0:
                # Navegador Chrome:
                chrome_configs = webdriver.ChromeOptions()
                chrome_configs.add_argument("--headless")

                navegador = webdriver.Chrome(options=chrome_configs)
                # navegador = webdriver.Chrome()
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

            # Validando se a página retornada trouxe um elemento inválido (Code 404)
            elemento_erro = soup.find_all('div', attrs={'class': 'code'})
            if elemento_erro:
                return [None]

            # PEGANDO OS DADOS DO CARD INICIAL DA PÁGINA:
            elementos_valores = soup.find_all('div', attrs={'class': '_card-body'})

            informacoes_titulo = [elemento.find('span').text for elemento in elementos_valores if elemento.find('span')]
            informacoes_titulo.insert(0, lista_ativos[indice])

            novas_informacoes = {
                'TITULO': informacoes_titulo[0],
                'COTAÇÃO': informacoes_titulo[1],
                'VARIAÇÃO (12M)': informacoes_titulo[2],
                'P/L': informacoes_titulo[3],
                'P/VP': informacoes_titulo[4],
                'DY (12M)': informacoes_titulo[5],
            }

            # PEGANDO OS DADOS DO CARD PRINCIPAL DE INFORMAÇÃO DA PÁGINA:
            principais_informacoes = soup.find('div', attrs={'id': 'table-indicators'}).find_all('div', attrs={'class': 'cell'})

            for elemento in principais_informacoes:
                temp = elemento.find('div', attrs={'class': 'value'})
                novas_informacoes[elemento.find('span').text.strip()] = temp.find('span').text.strip()

            # PEGANDO AS INFORMAÇÕES ESPECIFICAS DA EMPRESA:
            indicadores = soup.find('div', attrs={'id': 'table-indicators-company'}).find_all('div', attrs={'class': 'cell'})

            for indicador in indicadores:
                temp = indicador.find('span', attrs={'class': 'value'}).find('div', attrs={'class': 'detail-value'})

                if temp is not None:
                    novas_informacoes[indicador.find('span').text.strip()] = temp.text.strip()
                else:
                    novas_informacoes[indicador.find('span', attrs={'class': 'title'}).text.strip()] = indicador.find('span', attrs={'class': 'value'}).text.strip()

            lista_completa.append(novas_informacoes.copy())
            novas_informacoes.clear()

        navegador.quit()
        return lista_completa

    elif tipo_ativo == 'fiis':
        pass


if __name__ == '__main__':
    temp_result = get_stocks_from_web('acoes', ['ITUB4'])

    for item in temp_result:
        for chave, valor in item.items():
            print(f'{chave}: {valor}')
