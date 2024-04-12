from fastapi import FastAPI
from pydantic import BaseModel
import RaspagemDados.raspagem_dados as raspagem
from typing import TypedDict

app = FastAPI()


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


@app.post('/fiis')
def get_fiis(ativos: Ativos):
    for elemento in ativos.lista_ativos:
        if len(elemento) < 6:
            return {"message": "Os dados fornecidos estão incorretos!"}

    try:
        ativos_inicio: list = raspagem.new_pegar_dados_ativo('fiis', ativos.lista_ativos, False)
    except:
        return {"message": "Erro Interno"}
    else:
        if ativos_inicio[0] is None:
            return {"message": "Algum fii foi informado incorretamente!"}
        else:
            ativos_final: list = []

            for elemento in ativos_inicio:
                ativo_temporario: Fii = {
                    "tipo": "fii",
                    "ativo": elemento[0].upper(),
                    "cotacao": elemento[1],
                    "dy_12M": elemento[2],
                    "pvp": elemento[3],
                    "liquidez_diaria": elemento[4],
                    "variacao_12M": elemento[5]
                }

                ativos_final.append(ativo_temporario.copy())
                # ativo_temporario.clear()

            return ativos_final


@app.post('/acoes')
def get_acoes(ativos: Ativos):
    for elemento in ativos.lista_ativos:
        if len(elemento) < 5:
            return {"message": "Os dados fornecidos estão incorretos!"}

    try:
        ativos_inicio: list = raspagem.new_pegar_dados_ativo("acoes", ativos.lista_ativos, False)
    except:
        return {"message": "Erro Interno"}
    else:
        if ativos_inicio[0] is None:
            return {"message": "Alguma ação foi informada incorretamente!"}
        else:
            ativos_final: list = []

            for elemento in ativos_inicio:
                ativo_temporario: Acao = {
                    "tipo": "ação",
                    "ativo": elemento[0].upper(),
                    "cotacao": elemento[1],
                    "variacao_12M": elemento[2],
                    "pl": elemento[3],
                    "pvp": elemento[4],
                    "dy": elemento[5]
                }

                ativos_final.append(ativo_temporario.copy())

            return ativos_final
