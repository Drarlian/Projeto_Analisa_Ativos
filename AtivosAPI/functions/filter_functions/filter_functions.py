from typing import List
from AtivosAPI.functions.utilities_functions.utilities import format_cotacao

def filter_actives_by_cotacao(actives: List[dict], quantity: int):
    filtered_actives = sorted(actives, key=lambda active: format_cotacao(active['cotacao']), reverse=True)
    return filtered_actives[:quantity]


def filter_active_by_setores(actives: List[dict]):
    filtered_actives = sorted(list(set([active['segmento'] for active in actives])))
    return filtered_actives


def order_actives_by_views(actives: List[dict]):
    ordered_actives = sorted(actives, key=lambda active: active['views'], reverse=True)
    ordered_actives = [{"titulo": active["titulo"], "cotacao": active["cotacao"], "views": active["views"]} for active in ordered_actives]
    return ordered_actives


if __name__ == '__main__':
    teste = [1, 2, 3, 4, 5]
    print(teste[:200])
