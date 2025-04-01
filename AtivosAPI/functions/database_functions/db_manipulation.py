from motor.motor_asyncio import AsyncIOMotorClient
# from entities.requests import Informations
from dotenv import load_dotenv
import os

# Carrega as variáveis de ambiente do arquivo .env:
load_dotenv()

# Pegando as variáveis de ambiente carregadas:
DB_USER = os.getenv('DB_USER')
DB_PASSWORD = os.getenv('DB_PASSWORD')
DB_CLUSTER = os.getenv('DB_CLUSTER')

MONGODB_URL = f"mongodb+srv://{DB_USER}:{DB_PASSWORD}@{DB_CLUSTER.lower()}.crypllj.mongodb.net/?retryWrites=true&w=majority&appName={DB_CLUSTER}"


def get_connection():
    client = AsyncIOMotorClient(MONGODB_URL)  # ->  Cria um cliente Assíncrono para se Conectar ao banco de dados MongoDB.

    return client


# db_acoes = client["acoes_informations"]  # -> Nome do Banco de Dados do Projeto.
# db_fiis = client["fiis_informations"]  # -> Nome do Banco de Dados do Projeto.
# db_indicadores = client["indicadores_informations"]  # -> Nome do Banco de Dados do Projeto.
# db_tesouro = client["tesouro_informations"]  # -> Nome do Banco de Dados do Projeto.
