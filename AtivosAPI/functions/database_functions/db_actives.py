from AtivosAPI.functions.database_functions.db_manipulation import get_connection
from typing import List
# from bson import ObjectId
import re

client = get_connection()
db_acoes = client["acoes_informations"]  # -> Nome do Banco de Dados do Projeto.
db_fiis = client["fiis_informations"]  # -> Nome do Banco de Dados do Projeto.


async def get_all_actives(type_active: str, remove_ids: bool = True):
    if type_active == 'acoes':
        # Acessando a coleção "acoes":
        actives = await db_acoes.acoes.find().to_list(length=None)
    else:
        # Acessando a coleção "fiis":
        actives = await db_fiis.fiis.find().to_list(length=None)

    """
    Converte o ObjectId para string.
    Isso é necessário pois o ObjectId é um tipo específico de dado usado pelo MongoDB para identificar documentos 
    de forma única. No entanto, o ObjectId não é diretamente serializável para JSON.
    Para corrigir esse erro, devemos converter explicitamente o ObjectId para uma string antes de retornar os dados.
    """
    for active in actives:
        if remove_ids:
            del active["_id"]
        else:
            active["_id"] = str(active["_id"])  # -> Convertendo o _id para uma String.

    return actives


async def get_actives_by_title(type_active: str, titulos: List[str]):
    # Construindo a consulta com o operador $in:

    if type_active == 'acoes':
        # Acessando a coleção "acoes":
        actives = await db_acoes.acoes.find({"titulo": {"$in": titulos}}).to_list(length=None)
    else:
        # Acessando a coleção "fiis":
        actives = await db_fiis.fiis.find({"titulo": {"$in": titulos}}).to_list(length=None)

    # print(acoes)

    # Removendo o ID do itens:
    for active in actives:
        del active["_id"]

    return actives


async def find_active_by_approximation(type_active: str, termo: str, threshold: int = 65):
    """
    Função de Busca utilizando lógica Fuzzy.

    :param type_active: Define qual banco de dados acessar (acoes ou fiis).
    :param termo: Termo sendo procurado.
    :param threshold: Nivel de Similiaridade.
    :return: Retorna uma lista contendo todos os ativos similares encontrados.
    """
    from rapidfuzz import fuzz

    # 1. Pegando todas as ações do banco (isso aqui carrega tudo na memória):
    # Pode ser um problema futuro caso a base cresça muito.
    if type_active == 'acoes':
        # Acessando a coleção "acoes":
        actives = await db_acoes.acoes.find().to_list(length=None)
    else:
        # Acessando a coleção "fiis":
        actives = await db_fiis.fiis.find().to_list(length=None)

    resultados = []

    # 2. Iterando por todos os ativos encontradas no banco selecionado:
    for active in actives:
        titulo = active.get("titulo", "")  # Evita erro se não tiver o campo "titulo".

        # 3. Calcula o grau de similaridade entre o termo buscado e o título do ativo:

        # Busca apenas com o ".partial_ratio":
        # score = fuzz.partial_ratio(termo.lower(), titulo.lower())

        # Busca com abordagem híbrida utilizando ".ratio" e ".partial_ratio":
        # score = (fuzz.ratio(termo.lower(), titulo.lower()) + fuzz.partial_ratio(termo.lower(), titulo.lower())) / 2

        # Busca com abordagem tripla utilizando ".ratio", ".partial_ratio" e "token_sort_ratio":
        score = (
                        fuzz.ratio(termo.lower(), titulo.lower()) +
                        fuzz.partial_ratio(termo.lower(), titulo.lower()) +
                        fuzz.token_sort_ratio(termo.lower(), titulo.lower())
                ) / 3

        # print(f'Ativo: {active["titulo"]} | Score: {score}')

        # 4. Se a similaridade for maior que o limite (threshold), adiciona ao resultado
        if score >= threshold:
            active.pop("_id", None)  # Remove o campo "_id" do MongoDB

            # Adiciona a nota do ativo e o tipo do ativo:
            resultados.append({**active, "similaridade": score, "tipo_ativo": 'acoes' if type_active == 'acoes' else 'fiis'})

    # Ordenando os resultados pela "similaridade", do maior para o menor:
    resultados.sort(key=lambda active: active["similaridade"], reverse=True)

    # for active in resultados:
    #     print(f'Ativo: {active["titulo"]} | Score: {active["similaridade"]}')

    #     print('-' * 30)

    # Removendo a chave 'similaridade' antes de retornar:
    for active in resultados:
        active.pop("similaridade", None)

    # Retornando os 5 primeiros resultados.
    return resultados[:5]


async def find_active_by_title(type_active: str, partial_title: str):
    """
    Função de Busca com lógica de REGEX.
    """
    # Criando a regex para busca parcial (case-insensitive):
    regex = re.compile(partial_title, re.IGNORECASE)

    if type_active == 'acoes':
        # Busca usando regex no campo "titulo":
        actives = await db_acoes.acoes.find({"titulo": {"$regex": regex}}).to_list(length=None)
    else:
        # Busca usando regex no campo "titulo":
        actives = await db_fiis.fiis.find({"titulo": {"$regex": regex}}).to_list(length=None)

    # Remove o campo "_id" de cada documento retornado
    for active in actives:
        active.pop("_id", None)

    return actives


#  Adicionar múltiplas ações ao mesmo tempo
async def add_multiple_actives(type_active: str, actives: List[dict]):
    if type_active == 'acoes':
        # Acessando a coleção "acoes":
        result = await db_acoes.acoes.insert_many(actives)
    else:
        # Acessando a coleção "fiis":
        result = await db_fiis.fiis.insert_many(actives)

    return [str(id) for id in result.inserted_ids]  # Retorna uma lista de IDs


async def teste():
    resultado_acoes = await find_active_by_approximation('acoes', 'hg')
    resultado_fiis = await find_active_by_approximation('fiis', 'hg')

    print(resultado_acoes)
    print(len(resultado_acoes))
    for item in resultado_acoes:
        print(item)

    print('-' * 50)

    print(resultado_fiis)
    print(len(resultado_fiis))
    for item in resultado_fiis:
        print(item)


if __name__ == '__main__':
    import asyncio

    asyncio.run(teste())
