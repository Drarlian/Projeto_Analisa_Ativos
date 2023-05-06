from bs4 import BeautifulSoup
from selenium import webdriver

def pegar_dados_ativo(tipo_ativo: str, nome_ativo: str, titulo: bool = False) -> dict:
    # tipo_ativo = acoes | fiis
    if tipo_ativo not in ('acoes', 'fiis'):
        raise TypeError('O tipo do ativo não existe.')

    url = f'https://investidor10.com.br/{tipo_ativo}/{nome_ativo}/'

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
        if titulo_span_tag:  # -> Caso encontre a tag span adiciona o texto dela na lista.
            titulos.append(titulo_span_tag.text)


    indicadores = [nome_ativo]

    for elemento in elementos_valores:
        span_tag = elemento.find('span')
        if span_tag:
            indicadores.append(span_tag.text)

    if tipo_ativo == 'fiis':
        titulos[1] = titulos[1].split(' ')[1]  # -> Preciso editar pois esse titulo vem assim: 'SNCI11 Cotação'
        titulos[2] = ' '.join(titulos[2].split(' ')[1:])  # -> Preciso editar, esse titulo vem assim: 'SNCI11 DY (12M)'

    if titulo:
        lista_formatada = [titulos, indicadores]
    else:
        lista_formatada = [indicadores]

    return lista_formatada


if __name__ == '__main__':
    teste = pegar_dados_ativo('fiis', 'SNCI11', True)
    print(teste)
