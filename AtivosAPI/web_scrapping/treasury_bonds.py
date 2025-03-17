from typing import List
from bs4 import BeautifulSoup
from selenium import webdriver


def get_treasury_bonds_from_web() -> List[dict] | None:
    """

    :return:
    """

    url: str = 'https://www.tesourodireto.com.br/titulos/precos-e-taxas.htm'

    # PEGANDO TODOS OS TITULOS

    chrome_configs = webdriver.ChromeOptions()
    chrome_configs.add_argument("--headless")  # -> Tornando o processo de pesquisa do site invisível.

    navegador = webdriver.Chrome(options=chrome_configs)
    navegador.get(url)

    navegador.implicitly_wait(5)

    soup = BeautifulSoup(navegador.page_source, 'html.parser')

    # Validando se a página retornada trouxe um elemento inválido (Code 404)
    elemento_erro = soup.find_all('div', attrs={'class': 'code'})
    if elemento_erro:
        return None

    tabela_titulos = soup.find('table', attrs={'class': 'td-invest-table'})  # Pegando a primeira tabela com os titulos do tesouro.

    titulos_nome = tabela_titulos.find_all('span', attrs={'class': 'td-invest-table__name__text'})  # Pegando o texto que contem o nome do titulo.
    # titulos_ano = tabela_titulos.find_all('span', attrs={'class': 'td-invest-table__name__year'})
    titulos_valor = tabela_titulos.find_all('span', attrs={'class': 'td-invest-table__col__text'})

    new_titulos_nome = [titulo.get("aria-label") for titulo in titulos_nome]

    # print('-' * 30)

    # print(titulos_ano)
    # for ano in titulos_ano:
    #     print(ano.text)

    # print('-' * 30)

    # Pegando os valores gerais dos títulos: (Forma 1 - Slice)
    # valores = [valor.text.strip() for valor in titulos_valor]
    # # Agrupa os valores em sublistas de 4 elementos
    # new_titulos_valor = [valores[i:i + 4] for i in range(0, len(valores), 4)]
    # print(new_titulos_valor)

    # Pegando os valores gerais dos títulos: (Forma 2 - Verificação de Substring)
    new_titulos_valor = []
    posicao_atual: int = -1
    for valor in titulos_valor:
        if "%" in valor.text:
            new_titulos_valor.append([valor.text])
            posicao_atual += 1
        else:
            new_titulos_valor[posicao_atual].append(valor.text)

    # Classes possíveis para os dados:
    # td-invest-table__name__text
    # td-invest-table__name__year
    # td-invest-table__col__text

    final_response = []
    for indice in range(len(new_titulos_nome)):
        temp_dict = {'titulo': new_titulos_nome[indice], 'rentabilidade_anual': new_titulos_valor[indice][0],
                     'investimento_minimo': new_titulos_valor[indice][1],
                     'preco_unitario': new_titulos_valor[indice][2], 'vencimento': new_titulos_valor[indice][3]}

        final_response.append(temp_dict.copy())
        temp_dict.clear()

    # (Lista de listas é legacy, HashMap é mais otimizado)
    # Juntando todas as informações em um lugar só.
    # final_titulos = list(zip(new_titulos_nome, new_titulos_valor))
    # for titulo in final_titulos:
    #     print(titulo)

    return final_response


if __name__ == '__main__':
    print(get_treasury_bonds_from_web())
