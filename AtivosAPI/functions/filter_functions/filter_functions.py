from typing import List
from AtivosAPI.functions.utilities_functions.utilities import format_cotacao

def filter_actives(actives: List[dict], quantity: int):
    filtered_actives = sorted(actives, key=lambda active: format_cotacao(active['cotacao']), reverse=True)
    return filtered_actives[:quantity]


if __name__ == '__main__':
    teste = [1, 2, 3, 4, 5]
    print(teste[:200])
