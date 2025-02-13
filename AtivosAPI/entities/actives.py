from pydantic import BaseModel
from typing import TypedDict


class Ativos(BaseModel):
    lista_ativos: list


class Fii(TypedDict):
    tipo: str
    ativo: str
    cotacao: str
    dy_12M: str
    pvp: str
    liquidez_diaria: str
    variacao_12M: str


class Acao(TypedDict):
    tipo: str
    ativo: str
    cotacao: str
    variacao_12M: str
    pl: str
    pvp: str
    dy: str