import unicodedata
import re


def normalizar_texto(texto: str) -> str | None:
    # Se a chave for relacionada a dividend_yield, ignora-a (retorna None)
    # if re.match(r'^dividend_yield', texto):
    #     return None

    # Decomposição Unicode e remoção de acentos
    texto = unicodedata.normalize('NFKD', texto)
    texto = re.sub(r'[\u0300-\u036f]', '', texto)

    # Converte para minúsculas
    texto = texto.lower()

    # Substitui espaços por underline
    texto = texto.replace(" ", "_")

    # Remove parênteses e pontos
    texto = re.sub(r'[().]', '', texto)

    # Substitui padrão de underscore antes e depois de uma barra por apenas a barra
    texto = re.sub(r'_\s*/\s*_', '/', texto)
    # Caso ainda haja _ antes ou depois da barra, remove
    texto = re.sub(r'_/', '/', texto)
    texto = re.sub(r'/_', '/', texto)

    # Se começar com "no_", substitui por "n_"
    if texto.startswith("no_"):
        texto = "n_" + texto[3:]

    return texto


def format_cotacao(cotacao: str):
    cotacao = cotacao.split(' ')[1]
    cotacao = cotacao.replace('.', '')
    cotacao = cotacao.replace(',', '.')

    if '-' in cotacao:
        return 0.0
    else:
        return float(cotacao)
