import uvicorn
from fastapi import FastAPI, Query
from fastapi.responses import JSONResponse
from typing import List
from AtivosAPI.web_scrapping.b3_actives import b3_actives_from_web
from AtivosAPI.web_scrapping.treasury_bonds import get_treasury_bonds_from_web
from AtivosAPI.functions.database_functions import db_actives, db_images
from fastapi.middleware.cors import CORSMiddleware
from AtivosAPI.functions.filter_functions.filter_functions import (filter_actives_by_cotacao, filter_active_by_setores,
                                                                   order_actives_by_views)

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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
    lista_ativos: List[str] = ativos.upper().split(',')
    lista_ativos = list(set(lista_ativos))

    for elemento in lista_ativos:
        if len(elemento) < 5:
            return {"message": "Os dados fornecidos estão incorretos!"}

    try:
        # Consultando no banco se os ativos existem:
        # response = await db_actives.get_actives_by_title('fiis', lista_ativos)
        response = []
        for fii in lista_ativos:
            response_fii = await db_actives.get_one_and_increment_views('fiis', fii)
            if response_fii is not None:
                response.append(response_fii)

        # Se os ativos existirem, retorno eles:
        if len(response) == len(lista_ativos):
            return JSONResponse(status_code=200, content=response)

        # Se algum ativo não existir, faço o scrapping do ativo não existente:
        else:
            # Encontrando os ativos que não existem no banco de dados.
            ativos_no_banco: List[str] = [ativo['ticker'] for ativo in response]

            ativos_faltantes = list(set(lista_ativos) - set(ativos_no_banco))
            print(f'Ativos Faltantes: {ativos_faltantes}')

            # Fazendo scrapping dos ativos não existentes no banco:
            ativos_response: dict = b3_actives_from_web("fiis", ativos_faltantes, is_new=True)

            if not ativos_response["status"]:
                return JSONResponse(status_code=400, content={"message": ativos_response["message"]})

            # Os novos ativos devem iniciar com view = 1:
            for acao in ativos_response["informations"]:
                acao["views"] = 1

            # Adicionando os ativos do scrapping no banco: (Isso facilita para proximas buscas por ele)
            # (Não preciso me preocupar com os dados ficarem desatualizados pois a próxima schedule vai atualizar ele)
            await db_actives.add_multiple_actives('fiis', ativos_response["informations"])

            # Removendo o campo _id adicionado pelo mongo nos ativos adicionados ao banco.
            for ativo in ativos_response["informations"]:
                ativo.pop("_id", None)  # Remove a chave "_id" se existir, sem gerar erro caso não exista

            # Juntando os ativos do scrapping com os encontrados no banco.
            ativos_response["informations"] = response + ativos_response["informations"]
    except Exception as e:
        print(e)
        return JSONResponse(status_code=404, content={"message": "Erro interno durante a obtenção dos dados"})
    else:
        if ativos_response["status"]:
            return JSONResponse(status_code=200, content=ativos_response["informations"])
        else:
            return JSONResponse(status_code=404, content={"message": ativos_response["message"]})


@app.get('/get-all-fiis')
async def get_all_fiis():
    try:
        response_fiis = await db_actives.get_all_actives('fiis')
    except:
        return JSONResponse(status_code=404, content={"message": "Erro interno!"})
    else:
        return JSONResponse(status_code=200, content=response_fiis)


@app.get('/get-top-fiis/{fiis_quantity}')
async def get_top_fiis(fiis_quantity: str):
    try:
        response_fiis = await db_actives.get_all_actives('fiis')

        response_filter = filter_actives_by_cotacao(response_fiis, int(fiis_quantity))
    except:
        return JSONResponse(status_code=404, content={"message": "Erro interno!"})
    else:
        return JSONResponse(status_code=200, content=response_filter)


@app.get('/acoes')
async def get_acoes(ativos: str = Query(..., min_length=5, max_length=100,
                                 description="Ações separados por vírgula",
                                 examples=["WEGE3,ITSA4,CSNA3,PETR4,BBSE3"])):
    lista_ativos: List[str] = ativos.upper().split(',')
    lista_ativos = list(set(lista_ativos))

    for elemento in lista_ativos:
        if len(elemento) < 5:
            return {"message": "Os dados fornecidos estão incorretos!"}

    try:
        # Consultando no banco se os ativos existem:
        # response = await db_actives.get_actives_by_title('acoes', lista_ativos)
        response = []
        for acao in lista_ativos:
            response_acao = await db_actives.get_one_and_increment_views('acoes', acao)
            if response_acao is not None:
                response.append(response_acao)

        ativos_response = None

        # Se os ativos existirem, retorno eles:
        if len(response) == len(lista_ativos):
            pass

        # Se algum ativo não existir, faço o scrapping do ativo não existente:
        else:
            # Encontrando os ativos que não existem no banco de dados.
            ativos_no_banco: List[str] = [ativo['ticker'] for ativo in response]

            ativos_faltantes = list(set(lista_ativos) - set(ativos_no_banco))
            print(f'Ativos Faltantes: {ativos_faltantes}')

            # Fazendo scrapping dos ativos não existentes no banco:
            ativos_response: dict = b3_actives_from_web("acoes", ativos_faltantes, is_new=True)

            if not ativos_response["status"]:
                return JSONResponse(status_code=400, content={"message": ativos_response["message"]})

            # Adicionando os ativos do scrapping no banco: (Isso facilita para proximas buscas por ele)
            # (Não preciso me preocupar com os dados ficarem desatualizados pois a próxima schedule vai atualizar ele)
            await db_actives.add_multiple_actives('acoes', ativos_response["informations"])

            # Removendo o campo _id adicionado pelo mongo nos ativos adicionados ao banco.
            for ativo in ativos_response["informations"]:
                ativo.pop("_id", None)  # Remove a chave "_id" se existir, sem gerar erro caso não exista

            # Juntando os ativos do scrapping com os encontrados no banco.
            response = response + ativos_response["informations"]

        response_all_images = await db_images.get_images_by_ticker([acao['ticker'] for acao in response])
        for acao in response:
            for image in response_all_images:
                if acao["ticker"] == image["ticker"]:
                    acao["img"] = image["img"]
                    break

            if acao.get("img", None) is None:
                acao["img"] = ""

    except:
        return JSONResponse(status_code=404, content={"message": "Erro interno durante a obtenção dos dados"})
    else:
        if len(response) == len(lista_ativos):
            return JSONResponse(status_code=200, content=response)
        elif ativos_response is not None and ativos_response["status"]:
            return JSONResponse(status_code=200, content=response)
        else:
            return JSONResponse(status_code=404, content={"message": ativos_response["message"]})


@app.get('/get-all-acoes')
async def get_all_acoes():
    try:
        response_acoes = await db_actives.get_all_actives('acoes')
    except:
        return JSONResponse(status_code=404, content={"message": "Erro interno!"})
    else:
        return JSONResponse(status_code=200, content=response_acoes)


@app.get('/get-top-acoes/{acoes_quantity}')
async def get_top_acoes(acoes_quantity: str):
    try:
        response_acoes = await db_actives.get_all_actives('acoes')

        response_filter = filter_actives_by_cotacao(response_acoes, int(acoes_quantity))
    except:
        return JSONResponse(status_code=404, content={"message": "Erro interno!"})
    else:
        return JSONResponse(status_code=200, content=response_filter)


@app.get('/search-actives')
async def search_actives(term: str):
    try:
        response_acoes = await db_actives.find_active_by_approximation('acoes', term)
        response_fiis = await db_actives.find_active_by_approximation('fiis', term)
        response = response_acoes + response_fiis
    except:
        return JSONResponse(status_code=404, content={"message": "Erro interno!"})
    else:
        return JSONResponse(status_code=200, content=response)


@app.get('/get-all-sectors/{type_active}/{actives_quantity}')
async def get_all_sectors(type_active: str, actives_quantity: int):
    try:
        if type_active == 'acoes':
            response = await db_actives.get_all_actives('acoes')
            final_response = {'acoes': filter_active_by_setores(response)[:actives_quantity], 'fiis': []}

        elif type_active == 'fiis':
            response = await db_actives.get_all_actives('fiis')
            final_response = {'acoes': [], 'fiis': filter_active_by_setores(response)[:actives_quantity]}

        elif type_active == 'all':
            response_acoes = await db_actives.get_all_actives('acoes')
            response_acoes = filter_active_by_setores(response_acoes)

            response_fiis = await db_actives.get_all_actives('fiis')
            response_fiis = filter_active_by_setores(response_fiis)

            final_response = {'acoes': response_acoes[:actives_quantity], 'fiis': response_fiis[:actives_quantity]}

        else:
            return JSONResponse(status_code=404, content={"message": "Tipo de ativo inválido!"})
    except:
        return JSONResponse(status_code=404, content={"message": "Erro interno!"})
    else:
        return JSONResponse(status_code=200, content=final_response)


@app.get('/get-most-viewed/{type_active}/{actives_quantity}')
async def get_most_viewed(type_active: str, actives_quantity: int):
    try:
        if type_active == 'acoes':
            response = await db_actives.get_all_actives('acoes')
            final_response = {'acoes': order_actives_by_views(response)[:actives_quantity], 'fiis': []}

        elif type_active == 'fiis':
            response = await db_actives.get_all_actives('fiis')
            final_response = {'acoes': [], 'fiis': order_actives_by_views(response)[:actives_quantity]}

        elif type_active == 'all':
            response_acoes = await db_actives.get_all_actives('acoes')
            response_acoes = order_actives_by_views(response_acoes)

            response_fiis = await db_actives.get_all_actives('fiis')
            response_fiis = order_actives_by_views(response_fiis)

            final_response = {'acoes': response_acoes[:actives_quantity], 'fiis': response_fiis[:actives_quantity]}

        else:
            return JSONResponse(status_code=404, content={"message": "Tipo de ativo inválido!"})
    except:
        return JSONResponse(status_code=404, content={"message": "Erro interno!"})
    else:
        return JSONResponse(status_code=200, content=final_response)


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
