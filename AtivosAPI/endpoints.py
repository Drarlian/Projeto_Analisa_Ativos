from fastapi import FastAPI
from pydantic import BaseModel
import RaspagemDados.raspagem_dados as raspagem

app = FastAPI()


class Fii(BaseModel):
    ativo: str
    cotacao: str
    dy_12M: str
    pvp: str
    liquidez_diaria: str
    variacao_12M: str


class Acao(BaseModel):
    ativo: str
    cotacao: str
    dy_12M: str
    liquidez_diaria: str
    variacao_12M: str


@app.get('/fiis/{ativos}')
def get_fiis(ativos: str):
    if (',' not in ativos) and (len(ativos) < 6):
        return {"message": "Os dados fornecidos estão incorretos!"}

    ativos = ativos.split(',')

    if len(ativos[0]) >= 6:
        try:
            ativos_inicio: list = raspagem.new_pegar_dados_ativo('fiis', ativos, False)
        except:
            return {"message": "Erro Interno"}
        else:
            if ativos_inicio[0] is None:
                return {"message": "Algum fii foi informado incorretamente!"}
            else:
                ativos_final: list = []

                for elemento in ativos_inicio:
                    ativo_temporario: dict = {
                        'tipo': 'fii',
                        'ativo': elemento[0].upper(),
                        'cotacao': elemento[1],
                        'dy_12M': elemento[2],
                        'pvp': elemento[3],
                        'liquidez_diaria': elemento[4],
                        'variacao_12M': elemento[5]
                    }

                    ativos_final.append(ativo_temporario.copy())
                    ativo_temporario.clear()

                return ativos_final
    else:
        return {"message": "Os dados estão incorretos!"}

@app.get('/acoes/{ativos}')
def get_acoes(ativos: str):
    if (',' not in ativos) and (len(ativos) < 5):
        return {"message": "Os dados fornecidos estão incorretos!"}

    ativos = ativos.split(',')

    if len(ativos[0]) >= 5:
        try:
            ativos_inicio: list = raspagem.new_pegar_dados_ativo('acoes', ativos, False)
        except:
            return {"message": "Erro Interno"}
        else:
            if ativos_inicio[0] is None:
                return {"message": "Alguma ação foi informada incorretamente!"}
            else:
                ativos_final: list = []

                for elemento in ativos_inicio:
                    ativo_temporario = {
                        'tipo': 'ação',
                        'ativo': elemento[0].upper(),
                        'cotacao': elemento[1],
                        'variacao_12M': elemento[2],
                        'pl': elemento[3],
                        'pvp': elemento[4],
                        'dy': elemento[5]
                    }

                    ativos_final.append(ativo_temporario.copy())
                    ativo_temporario.clear()

                return ativos_final
    else:
        return {"message": "Os dados estão incorretos!"}
