from AtivosAPI.functions.database_functions.db_manipulation import get_connection
from typing import List

client = get_connection()
db_images = client["images"]  # -> Nome do Banco de Dados do Projeto.


async def get_all_images(remove_ids: bool = True) -> List[dict]:
    images = await db_images.acoes.find().to_list(length=None)

    # Convertendo ObjectId para string
    for image in images:
        if remove_ids:
            del image["_id"]
        else:
            image["_id"] = str(image["_id"])  # -> Convertendo o _id para uma String.

    return images


async def get_image_by_ticker(ticker: str, remove_id: bool = True) -> dict:
    image = await db_images.acoes.find_one({"ticker": ticker}).to_list(length=None)

    if image:
        if remove_id:
            del image["_id"]
        else:
            image["_id"] = str(image["_id"])  # Convertendo ObjectId para string

    return image


async def get_images_by_ticker(tickers: List[str]) -> List[dict]:
    # Construindo a consulta com o operador $in:

    actives = await db_images.acoes.find({"ticker": {"$in": tickers}}).to_list(length=None)

    # Removendo o ID do itens:
    for active in actives:
        del active["_id"]

    return actives


async def update_one_image(new_data: dict) -> int:
    # upsert=True -> Serve para adicionar o item caso ele não exista.
    result = await db_images.acoes.update_one({"ticker": new_data["ticker"].upper()}, {"$set": new_data}, upsert=True)

    return result.modified_count  # Retorna quantos documentos foram modificados.


if __name__ == "__main__":
    import asyncio

    response = asyncio.run(get_all_images())
    print(response)
