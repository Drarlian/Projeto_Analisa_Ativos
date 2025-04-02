from AtivosAPI.functions.database_functions.db_manipulation import get_connection
from typing import List
# from bson import ObjectId

client = get_connection()
db_acoes = client["acoes_informations"]  # -> Nome do Banco de Dados do Projeto.


async def get_all_stocks(remove_ids: bool = True):
    # pessoas = await db.people.find().to_list(length=None)
    # return pessoas

    # Acessar a coleção "people":
    acoes = await db_acoes.acoes.find().to_list(length=None)

    """
    Converte o ObjectId para string.
    Isso é necessário pois o ObjectId é um tipo específico de dado usado pelo MongoDB para identificar documentos 
    de forma única. No entanto, o ObjectId não é diretamente serializável para JSON.
    Para corrigir esse erro, devemos converter explicitamente o ObjectId para uma string antes de retornar os dados.
    """
    for acao in acoes:
        if remove_ids:
            del acao["_id"]
        else:
            acao["_id"] = str(acao["_id"])  # -> Convertendo o _id para uma String.

    return acoes


async def get_acoes_por_titulos(titulos: List[str]):
    # Construindo a consulta com o operador $in
    acoes = await db_acoes.acoes.find({"titulo": {"$in": titulos}}).to_list(length=None)
    # print(acoes)

    # Removendo o ID do itens:
    for acao in acoes:
        del acao["_id"]

    return acoes


#  Adicionar múltiplas ações ao mesmo tempo
async def add_multiplas_acoes(acoes: List[dict]):
    result = await db_acoes.acoes.insert_many(acoes)
    return [str(id) for id in result.inserted_ids]  # Retorna uma lista de IDs


if __name__ == '__main__':
    import asyncio

    temp_test = ['WEGE3', 'VAMO3', 'AMER3']
    teste = asyncio.run(get_acoes_por_titulos(temp_test))

    print(teste)
    print(len(teste))

    print(temp_test)
    print(len(temp_test))
