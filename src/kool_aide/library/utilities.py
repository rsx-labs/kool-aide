import json


def get_salt() -> str:
    return '1234567890123456'


def encrypt(clear_text: str, salt: str) -> str:
    return clear_text


def decrypt(cypher_text: str, salt : str) -> str:
    return cypher_text
