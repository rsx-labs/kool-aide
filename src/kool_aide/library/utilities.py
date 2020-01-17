import json
from datetime import datetime, time

from kool_aide.library.constants import *
from kool_aide.assets.resources.version import *

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
            print(f'{datetime.now().strftime("%d-%b-%y %H:%M:%S")} [CONSOLE] {message}')
        else:
            print(f'{message}')

def get_start_date(date: datetime) -> datetime:
    return datetime(date.year, date.month, date.day, 0,0,1)

def get_end_date(date: datetime) -> datetime:
    return datetime(date.year, date.month, date.day, 23,59,59)

def get_date(string_date) -> datetime:
    try:
        return datetime.strptime(string_date, '%m/%d/%Y')
    except:
        return None

def get_param_value(parameter_name: str, json_parameters, default_value = None)-> str:
    if parameter_name not in json_parameters:
        return default_value
    else:
        return json_parameters[parameter_name]

def get_version() -> str:
    try:
        return APP_TITLE + ' v' + APP_VERSION
    except:
        return ""

def get_cell_address(col:int, row:int, absolute=True)->str:
    columns='ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    column = columns[col]

    if absolute:
        return f'${column}${row}'
    else:
        return f'{column}{row}'

def get_cell_range_address(address1:str, address2:str, sheet='')-> str:
    if len(sheet)>0:
        return f'{sheet}!{address1}:{address2}'
    else:
        return f'{address1}:{address2}'


