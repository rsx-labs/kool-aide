import json
from datetime import datetime, time

from kool_aide.library.constants import *

def get_salt() -> str:
    # dont use
    return '1234567890'


def encrypt(clear_text: str, salt: str) -> str:
    return clear_text


def decrypt(cypher_text: str, salt: str) -> str:
    return cypher_text


def append_date_to_file_name(file_name: str) -> str:
    long_months = LONG_MONTHS[datetime.now().month-1]
    short_months = SHORT_MONTHS[datetime.now().month-1]
    day = str(datetime.now().day).zfill(2)
    month = str(datetime.now().month).zfill(2)
    year = str(datetime.now().year)

    result = file_name.replace('[LM]',long_months)
    result = result.replace('[SM]', short_months)
    result = result.replace('[M]', month)
    result = result.replace('[D]',day)
    result = result.replace('[Y]', year) 

    return result

# using Vigenère cipher  
def encrypt(value, key):
    enc = []
    for i in range(len(value)):
        key_c = key[i % len(key)]
        enc_c = chr((ord(value[i]) + ord(key_c)) % 256)
        enc.append(enc_c)
    return base64.urlsafe_b64encode(("".join(enc)).encode())

# using Vigenère cipher
def decrypt(value, key):
    dec = []
    value = base64.urlsafe_b64decode(value).decode("utf-8")
    for i in range(len(value)):
        key_c = key[i % len(key)]
        dec_c = chr((256 + ord(value[i]) - ord(key_c)) % 256)
        dec.append(dec_c)
    return "".join(dec)

def print_to_screen(message: str, quiet: bool, show_time = True) -> None:
    if not  quiet:
        if show_time:
            print(f'[{datetime.now().strftime("%H:%M:%S")}] {message}')
        else:
            print(f'{message}')


