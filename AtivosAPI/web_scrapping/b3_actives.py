from bs4 import BeautifulSoup
from selenium import webdriver
from AtivosAPI.functions.utilities_functions.utilities import normalizar_texto
from datetime import datetime


def b3_actives_from_web(tipo_ativo: str, lista_ativos: list, is_new: bool = False) -> dict:
    """
    Procura pelas informações do(s) ativo(s) informado(s) de forma rápida e otimizada.
    Formato do Retorno:
    [['ativo1', 'informacao1', 'informacao2'], ['ativo2', 'outra_informacao1', 'outra_informacao2']]
    :param tipo_ativo: Tipo do ativo que será pesquisado. Opções: acoes | fiis
    :param lista_ativos: Lista contendo os ativos. A lista deve conter apenas ativos do mesmo tipo.
    :param is_new: Define se os campos "views", "nota" e "indicadores_postivos" devem ser adicionados.
    :return: Retorna uma lista de dicionários contendo as informações dos ativos solicitados.
    """
    try:
        # tipo_ativo = acoes | fiis
        if tipo_ativo not in ('acoes', 'fiis'):
            return {"status": False, "message": "O tipo de ativo informado não existe ou não está disponível."}

        lista_completa = []

        lista_urls = []
        for nome_ativo in lista_ativos:
            lista_urls.append(f'https://investidor10.com.br/{tipo_ativo}/{nome_ativo}/')

        # Navegador Chrome:
        chrome_configs = webdriver.ChromeOptions()
        chrome_configs.add_argument("--headless")

        navegador = webdriver.Chrome(options=chrome_configs)
        # navegador = webdriver.Chrome()

        for indice, url in enumerate(lista_urls):
            if indice == 0:
                navegador.get(url)

                navegador.implicitly_wait(5)
            else:
                # Abrir uma nova aba
                navegador.execute_script("window.open('', '_blank');")

                # Mudar para a segunda aba
                navegador.switch_to.window(navegador.window_handles[indice])

                # Abrir o segundo link e pegar dados
                navegador.get(url)

                navegador.implicitly_wait(5)

            soup = BeautifulSoup(navegador.page_source, 'html.parser')

            # Pegando o nomes da empresa:
            elemento_nome = soup.find('h2', attrs={'class': 'name-company'}).text

            # PEGANDO OS DADOS DO CARD INICIAL DA PÁGINA:
            elementos_valores = soup.find_all('div', attrs={'class': '_card-body'})

            informacoes_titulo = [elemento.find('span').text for elemento in elementos_valores if elemento.find('span')]

            if len(informacoes_titulo) == 0:
                continue

            informacoes_titulo.insert(0, elemento_nome)
            informacoes_titulo.insert(1, lista_ativos[indice])

            novas_informacoes = dict()

            if tipo_ativo == 'acoes':
                # Adicionando o texto "S.A." no nome das empresas (Apenas para Ações):
                if 'S.A.' not in informacoes_titulo[0].upper():
                    informacoes_titulo[0] = f'{informacoes_titulo[0].upper()} S.A.'

                novas_informacoes['nome'] = informacoes_titulo[0].upper()
                novas_informacoes['ticker'] = informacoes_titulo[1].upper()
                novas_informacoes['cotacao'] = informacoes_titulo[2].upper()
                novas_informacoes['variacao_12m'] = informacoes_titulo[3]
                novas_informacoes['p/l'] = informacoes_titulo[4]
                novas_informacoes['p/vp'] = informacoes_titulo[5]
                novas_informacoes['dy_12m'] = informacoes_titulo[6]

                # PEGANDO OS DADOS DO CARD PRINCIPAL DE INFORMAÇÃO DA PÁGINA:
                principais_informacoes = soup.find('div', attrs={'id': 'table-indicators'}).find_all('div', attrs={'class': 'cell'})

                for elemento in principais_informacoes:
                    temp_title = normalizar_texto(elemento.find('span').text.strip())

                    if temp_title.startswith('dividend_yield'):
                        continue

                    temp_value = elemento.find('div', attrs={'class': 'value'}).find('span').text.strip()

                    novas_informacoes[temp_title] = temp_value

                # PEGANDO AS INFORMAÇÕES ESPECIFICAS DA EMPRESA:
                indicadores = soup.find('div', attrs={'id': 'table-indicators-company'}).find_all('div', attrs={'class': 'cell'})

                for indicador in indicadores:
                    temp_value = indicador.find('span', attrs={'class': 'value'}).find('div', attrs={'class': 'detail-value'})

                    if temp_value is not None:
                        temp_title = normalizar_texto(indicador.find('span').text.strip())
                        novas_informacoes[temp_title] = temp_value.text.strip()
                    else:
                        temp_title = normalizar_texto(indicador.find('span', attrs={'class': 'title'}).text.strip())
                        novas_informacoes[temp_title] = indicador.find('span', attrs={'class': 'value'}).text.strip()

            elif tipo_ativo == 'fiis':
                novas_informacoes['nome'] = informacoes_titulo[0].upper()
                novas_informacoes['ticker'] = informacoes_titulo[1].upper()
                novas_informacoes['cotacao'] = informacoes_titulo[2].upper()
                novas_informacoes['dy_12M'] = informacoes_titulo[3]
                novas_informacoes['p/vp'] = informacoes_titulo[4]
                novas_informacoes['liquidez_diaria'] = informacoes_titulo[5]
                novas_informacoes['variacao_12m'] = informacoes_titulo[6]

                # PEGANDO OS DADOS DO CARD PRINCIPAL DE INFORMAÇÃO DA PÁGINA:
                principais_informacoes = soup.find('div', attrs={'id': 'table-indicators'}).find_all('div', attrs={
                    'class': 'cell'})

                for elemento in principais_informacoes:
                    temp_title = normalizar_texto(elemento.find('div', attrs={'class': 'desc'}).find('span').text.strip())
                    temp_value = elemento.find('div', attrs={'class': 'desc'}).find('div', attrs={'class': 'value'})
                    novas_informacoes[temp_title] = temp_value.find('span').text.strip()

            if is_new:
                novas_informacoes['views'] = 0
                novas_informacoes['nota'] = 'N/A'
                novas_informacoes['indicadores_postivos'] = []

            novas_informacoes['ultima_atualizacao'] = datetime.now().strftime("%d/%m/%Y - %H:%M")

            lista_completa.append(novas_informacoes.copy())
            novas_informacoes.clear()

        navegador.quit()

        if len(lista_completa) > 0:
            return {"status": True, "informations": lista_completa}
        else:
            return {"status": False, "message": "O(s) ativo(s) informado(s) não existe(m) ou não foi(ram) encontrado(s)."}
    except:
        return {"status": None, "message": "Erro interno durante a obtenção dos dados."}


if __name__ == '__main__':
    import timeit
    import cProfile

    # temp_result = b3_actives_from_web('fiis', ['CPTS11', 'RBVA11', 'NSLU11', 'XPML11'])
    # print(temp_result)

    # tempo = timeit.timeit(lambda: b3_actives_from_web(tipo_ativo='fiis', lista_ativos=['CPTS11', 'RBVA11', 'NSLU11', 'XPML11']), number=10)  # Executa 10 vezes
    # print(f"Tempo médio por execução: {tempo / 10:.6f} segundos")  # -> Tempo médio por execução: 14.038316 segundos

    # cProfile.run('b3_actives_from_web(tipo_ativo="fiis", lista_ativos=["CPTS11", "RBVA11", "NSLU11", "XPML11"])')

    temp_result = b3_actives_from_web(tipo_ativo='acoes', lista_ativos=['WEGE3', 'PETR3', 'ITSA3', 'KLBN3'])

    if temp_result['status']:
        for item in temp_result['informations']:
            print(item)
            for chave, valor in item.items():
                print(f'{chave}: {valor}')
            print('-' * 30)
