from AtivosAPI.functions.database_functions.db_manipulation import get_connection
from typing import List
# from bson import ObjectId

client = get_connection()
db_fiis = client["fiis_informations"]  # -> Nome do Banco de Dados do Projeto.


async def get_all_fiis(remove_ids: bool = True):
    # pessoas = await db.people.find().to_list(length=None)
    # return pessoas

    # Acessar a coleção "people":
    fiis = await db_fiis.fiis.find().to_list(length=None)

    """
    Converte o ObjectId para string.
    Isso é necessário pois o ObjectId é um tipo específico de dado usado pelo MongoDB para identificar documentos 
    de forma única. No entanto, o ObjectId não é diretamente serializável para JSON.
    Para corrigir esse erro, devemos converter explicitamente o ObjectId para uma string antes de retornar os dados.
    """
    for fii in fiis:
        if remove_ids:
            del fii["_id"]
        else:
            fii["_id"] = str(fii["_id"])  # -> Convertendo o _id para uma String.

    return fiis


async def get_fiis_por_titulos(titulos: List[str]):
    # Construindo a consulta com o operador $in
    fiis = await db_fiis.fiis.find({"titulo": {"$in": titulos}}).to_list(length=None)
    # print(fiis)

    # Removendo o ID do itens:
    for fii in fiis:
        del fii["_id"]

    return fiis


#  Adicionar múltiplas ações ao mesmo tempo
async def add_multiplos_fiis(fiis: List[dict]):
    result = await db_fiis.fiis.insert_many(fiis)
    return [str(id) for id in result.inserted_ids]  # Retorna uma lista de IDs


if __name__ == '__main__':
    import asyncio

    temp_test = ['KNRI11', 'RBVA11', 'NSLU11']
    # teste = asyncio.run(get_fiis_por_titulos(temp_test))
    # teste = asyncio.run(add_multiplos_fiis([{'titulo': 'teste', 'cotacao': 'R$ 0,00'}]))
    teste = asyncio.run(get_all_fiis())

    print(teste)
    print(len(teste))

    print(temp_test)
    print(len(temp_test))
