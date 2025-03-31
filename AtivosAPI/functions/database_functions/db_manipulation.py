import asyncio

from motor.motor_asyncio import AsyncIOMotorClient
# from entities.requests import Informations
from typing import List

from dotenv import load_dotenv
import os

# Carrega as variáveis de ambiente do arquivo .env:
load_dotenv()

# Pegando as variáveis de ambiente carregadas:
DB_USER = os.getenv('DB_USER')
DB_PASSWORD = os.getenv('DB_PASSWORD')
DB_CLUSTER = os.getenv('DB_CLUSTER')

MONGODB_URL = f"mongodb+srv://{DB_USER}:{DB_PASSWORD}@{DB_CLUSTER.lower()}.crypllj.mongodb.net/?retryWrites=true&w=majority&appName={DB_CLUSTER}"

client = AsyncIOMotorClient(MONGODB_URL)  # ->  Cria um cliente Assíncrono para se Conectar ao banco de dados MongoDB.

db_acoes = client["acoes_informations"]  # -> Nome do Banco de Dados do Projeto.
db_fiis = client["fiis_informations"]  # -> Nome do Banco de Dados do Projeto.
db_indicadores = client["indicadores_informations"]  # -> Nome do Banco de Dados do Projeto.
db_tesouro = client["tesouro_informations"]  # -> Nome do Banco de Dados do Projeto.


async def get_all_stocks():
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
        acao["_id"] = str(acao["_id"])  # -> Convertendo o _id para uma String.

    return acoes


async def get_acoes_por_titulos(titulos: List[str]):
    # Construindo a consulta com o operador $in
    acoes = await db_acoes.acoes.find({"titulo": {"$in": titulos}}).to_list(length=None)
    print(acoes)

    # Removendo o ID do itens:
    for acao in acoes:
        del acao["_id"]

    return acoes


if __name__ == '__main__':
    import asyncio

    temp_test = ['WEGE3', 'VAMO3', 'AMER3']
    teste = asyncio.run(get_acoes_por_titulos(temp_test))

    print(teste)
    print(len(teste))

    print(temp_test)
    print(len(temp_test))
