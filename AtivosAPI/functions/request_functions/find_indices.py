import requests


def get_selic() -> dict:
    """
    Pega o valor da taxa selic atual através da API do Banco Central.
    :return: Retorna o valor da taxa selic encontrada na API do Banco Central.
    """

    numero_tentativas: int = 1
    while True:
        try:
            # Obtendo a Selic:
            # url_selic = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.11/dados?formato=json'
            url_selic = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.432/dados/ultimos/1?formato=json'

            response_selic = requests.get(url_selic)
            selic_data = response_selic.json()
            print(selic_data)

            print(f"Taxa Selic Atual: {float(selic_data[0]['valor']):.2f}%")  # -> Taxa Selic Atual: 13.25%
        except:
            print(f'Tentativa Atual: {numero_tentativas}')

            if numero_tentativas == 5:
                return {"status": False, "value": None}

            numero_tentativas += 1
        else:
            return {"status": True, "value": round(float(selic_data[0]['valor']), 2)}


def get_ipca() -> dict:
    """
    Pega o valor do IPCA dos últimos 12 meses da API do Banco Central e calcula o IPCA atual para obter o IPCA acumulado.
    :return: Retorna o valor do IPCA acumulado nos últimos 12 meses com base nos dados da API do Banco Central.
    """

    numero_tentativas: int = 1
    while True:
        try:
            # Obtendo o IPCA:
            # url_ipca = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.10844/dados/ultimos/12?formato=json'
            url_ipca = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.433/dados/ultimos/12?formato=json'
            response = requests.get(url_ipca)
            dados_ipca = response.json()
            print(dados_ipca)

            # Inicializa o fator acumulado
            fator_acumulado = 1.0

            # Para cada mês, converte o valor para decimal e multiplica
            for item in dados_ipca:
                valor_mensal = float(item['valor'])
                fator_acumulado *= (1 + valor_mensal / 100)

            # Subtrai 1 para obter a variação acumulada
            ipca_acumulado_decimal = fator_acumulado - 1
            ipca_acumulado_percent = ipca_acumulado_decimal * 100

            print(f"IPCA Atual: {ipca_acumulado_percent:.2f}%")  # -> IPCA Atual: 4.56%
        except:
            print(f'Tentativa Atual: {numero_tentativas}')

            if numero_tentativas == 5:
                return {"status": False, "value": None}

            numero_tentativas += 1
        else:
            return {"status": True, "value": round(ipca_acumulado_percent, 2)}


if __name__ == "__main__":
    selic = get_selic()
    print(selic)  # -> {'status': True, 'value': 13.25}

    ipca = get_ipca()
    print(ipca)   # -> {'status': True, 'value': 4.56}
