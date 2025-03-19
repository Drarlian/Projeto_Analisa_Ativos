import unicodedata
import re


def normalizar_texto(texto: str) -> str:
    texto = unicodedata.normalize('NFKD', texto)  # Decomposição Unicode
    texto = re.sub(r'[\u0300-\u036f]', '', texto)  # Remove os acentos
    texto = texto.lower()  # Converte para minúsculas
    texto = texto.replace(" ", "_")  # Substitui espaços por _
    texto = re.sub(r'[().]', '', texto)  # Remove parênteses e pontos
    return texto
