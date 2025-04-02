import uvicorn
from fastapi import FastAPI, Query
from fastapi.responses import JSONResponse
from typing import List
from AtivosAPI.entities.actives import Fii, Acao
from AtivosAPI.web_scrapping.b3_actives import b3_actives_from_web
from AtivosAPI.web_scrapping.treasury_bonds import get_treasury_bonds_from_web
from AtivosAPI.functions.database_functions import db_acoes, db_fiis

app = FastAPI()

"""
# Caso seja passado um valor, esse valor passado será usado.
# Caso não seja passado um valor, será usado o "default".
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
async def get_fiis(ativos: str = Query(..., min_length=5, max_length=100,
                                 description="Fii's separados por vírgula",
                                 examples=["XPLG11,KNRI11,ALZR11,BTLG11,HGLG11"])):
    lista_ativos: List[str] = ativos.split(',')

    for elemento in lista_ativos:
        if len(elemento) < 5:
            return {"message": "Os dados fornecidos estão incorretos!"}

    try:
        # Consultando no banco se os ativos existem:
        response = await db_fiis.get_fiis_por_titulos(lista_ativos)

        # Se os ativos existirem, retorno eles:
        if len(response) == len(lista_ativos):
            return JSONResponse(status_code=200, content=response)

        # Se algum ativo não existir, faço o scrapping do ativo não existente:
        else:
            # Encontrando os ativos que não existem no banco de dados.
            ativos_no_banco: List[str] = [ativo['titulo'] for ativo in response]

            ativos_faltantes = list(set(lista_ativos) - set(ativos_no_banco))
            print(f'Ativos Faltantes: {ativos_faltantes}')

            # Fazendo scrapping dos ativos não existentes no banco:
            ativos_response: dict = b3_actives_from_web("fiis", ativos_faltantes)

            # Adicionando os ativos do scrapping no banco: (Isso facilita para proximas buscas por ele)
            # (Não preciso me preocupar com os dados ficarem desatualizados pois a próxima schedule vai apagar ele)
            await db_fiis.add_multiplos_fiis(ativos_response["informations"])

            # Removendo o campo _id adicionado pelo mongo nos ativos adicionados ao banco.
            for ativo in ativos_response["informations"]:
                ativo.pop("_id", None)  # Remove a chave "_id" se existir, sem gerar erro caso não exista

            # Juntando os ativos do scrapping com os encontrados no banco.
            ativos_response["informations"] = ativos_response["informations"] + response
    except:
        return JSONResponse(status_code=404, content={"message": "Erro interno durante a obtenção dos dados"})
    else:
        if ativos_response["status"]:
            return JSONResponse(status_code=200, content=ativos_response["informations"])
        else:
            return JSONResponse(status_code=404, content={"message": ativos_response["message"]})

@app.get('/get-all-fiis')
async def get_all_fiis():
    try:
        response_fiis = await db_fiis.get_all_fiis()
    except:
        return JSONResponse(status_code=404, content={"message": "Erro interno!"})
    else:
        return JSONResponse(status_code=200, content=response_fiis)


@app.get('/acoes')
async def get_acoes(ativos: str = Query(..., min_length=5, max_length=100,
                                 description="Ações separados por vírgula",
                                 examples=["WEGE3,ITSA4,CSNA3,PETR4,BBSE3"])):
    lista_ativos: List[str] = ativos.split(',')

    for elemento in lista_ativos:
        if len(elemento) < 5:
            return {"message": "Os dados fornecidos estão incorretos!"}

    try:
        # Consultando no banco se os ativos existem:
        response = await db_acoes.get_acoes_por_titulos(lista_ativos)

        # Se os ativos existirem, retorno eles:
        if len(response) == len(lista_ativos):
            return JSONResponse(status_code=200, content=response)

        # Se algum ativo não existir, faço o scrapping do ativo não existente:
        else:
            # Encontrando os ativos que não existem no banco de dados.
            ativos_no_banco: List[str] = [ativo['titulo'] for ativo in response]

            ativos_faltantes = list(set(lista_ativos) - set(ativos_no_banco))
            print(f'Ativos Faltantes: {ativos_faltantes}')

            # Fazendo scrapping dos ativos não existentes no banco:
            ativos_response: dict = b3_actives_from_web("acoes", ativos_faltantes)

            # Adicionando os ativos do scrapping no banco: (Isso facilita para proximas buscas por ele)
            # (Não preciso me preocupar com os dados ficarem desatualizados pois a próxima schedule vai apagar ele)
            await db_acoes.add_multiplas_acoes(ativos_response["informations"])

            # Removendo o campo _id adicionado pelo mongo nos ativos adicionados ao banco.
            for ativo in ativos_response["informations"]:
                ativo.pop("_id", None)  # Remove a chave "_id" se existir, sem gerar erro caso não exista

            # Juntando os ativos do scrapping com os encontrados no banco.
            ativos_response["informations"] = ativos_response["informations"] + response
    except:
        return JSONResponse(status_code=404, content={"message": "Erro interno durante a obtenção dos dados"})
    else:
        if ativos_response["status"]:
            return JSONResponse(status_code=200, content=ativos_response["informations"])
        else:
            return JSONResponse(status_code=404, content={"message": ativos_response["message"]})


@app.get('/get-all-acoes/')
async def get_all_acoes():
    try:
        response_acoes = await db_acoes.get_all_stocks()
    except:
        return JSONResponse(status_code=404, content={"message": "Erro interno!"})
    else:
        return JSONResponse(status_code=200, content=response_acoes)



@app.get('/tesouro-direto')
def get_all_treasury_bonds():
    try:
        trasury_bonds_response = get_treasury_bonds_from_web()
    except:
        return JSONResponse(status_code=404, content={"message": "Erro interno durante a obtenção dos dados."})
    else:
        if trasury_bonds_response:
            return trasury_bonds_response
        else:
            return JSONResponse(status_code=404, content={"message": "Erro interno durante a obtenção dos dados."})


if __name__ == "__main__":
    uvicorn.run(app, host='0.0.0.0', port=8000)
