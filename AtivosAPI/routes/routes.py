import RaspagemDados.raspagem_dados as raspagem
import uvicorn
from fastapi import FastAPI, Query
from typing import List
from AtivosAPI.entities.actives import Fii, Acao

app = FastAPI()

"""
# Caso seja passado um valor, esse valor será considerado "default" e usado caso não seja passado nenhum valor.
...,            # Parâmetro obrigatório
"valor_default" # Valor Default

min_length = 3,   # Mínimo de 3 caracteres
max_length = 50,  # Máximo de 50 caracteres
pattern = "^[a-zA-Z0-9 ]+$",  # Apenas letras, números e espaços

# Campos usados para uma melhor visualização/explicação na documentação do Swagger UI:
description = "Termo de pesquisa para filtrar os itens",
examples = ["produto123"]  # Exemplos para a documentação
"""


@app.get('/fiis')
def get_fiis(ativos: str = Query(..., min_length=6, max_length=100,
                                 description="Fii's separados por vírgula",
                                 examples=["XPLG11,KNRI11,ALZR11,BTLG11,HGLG11"])):
    lista_ativos: List[str] = ativos.split(',')

    for elemento in lista_ativos:
        if len(elemento) < 6:
            return {"message": "Os dados fornecidos estão incorretos!"}

    try:
        ativos_inicio: list = raspagem.new_pegar_dados_ativo('fiis', lista_ativos, False)
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


@app.get('/acoes')
def get_acoes(ativos: str = Query(..., min_length=5, max_length=100,
                                 description="Ações separados por vírgula",
                                 examples=["WEGE3,ITSA4,CSNA3,PETR4,BBSE3"])):
    lista_ativos: List[str] = ativos.split(',')

    for elemento in lista_ativos:
        if len(elemento) < 5:
            return {"message": "Os dados fornecidos estão incorretos!"}

    try:
        ativos_inicio: list = raspagem.new_pegar_dados_ativo("acoes", lista_ativos, False)
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


if __name__ == "__main__":
    uvicorn.run(app, host='0.0.0.0', port=8000)
