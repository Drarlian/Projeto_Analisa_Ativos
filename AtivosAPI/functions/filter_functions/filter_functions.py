from typing import List
from AtivosAPI.functions.utilities_functions.utilities import format_cotacao

def filter_actives_by_cotacao(actives: List[dict], quantity: int):
    filtered_actives = sorted(actives, key=lambda active: format_cotacao(active['cotacao']), reverse=True)
    return filtered_actives[:quantity]


def filter_actives(actives: List[dict], filters: List[str]):
    def update_filter_item_to_title_case(filter_title, filter_value):
        if filter_title == 'tipo_de_fundo':
            if len(filter_value.split(' ')) == 1:
                return filter_value.title()
            else:
                temp = filter_value.split(' ')
                return ' '.join(temp[:-1]) + ' ' + temp[-1].title()
        else:
            return filter_value

    response = dict()
    for filter_item in filters:
        response[filter_item] = sorted(list(set([update_filter_item_to_title_case(filter_item, active[filter_item]) for active in actives])))

    return response


def order_actives_by_views(actives: List[dict]):
    ordered_actives = sorted(actives, key=lambda active: active['views'], reverse=True)
    ordered_actives = [{"ticker": active["ticker"], "cotacao": active["cotacao"], "views": active["views"]} for active in ordered_actives]
    return ordered_actives


if __name__ == '__main__':
    teste = [1, 2, 3, 4, 5]
    print(teste[:200])
    print('teste'.title())
