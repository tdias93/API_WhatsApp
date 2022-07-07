from tqdm import tqdm
from requests.exceptions import Timeout

import json
import requests
import PY_ReadWriteExcel


def ApiPayload(session, number, text):

    """ Gera Body da API

        Args:
            session (str) : Nome da sessão
            number (str)  : Número do destinatario
            text (str)    : Texto da menssagem
            
    """

    payload = json.dumps({
        "session": f"{session}",
        "number": f"{number}",
        "text": f"{text}"
    })

    return payload

def ApiHeaders(sessionkey):

    """ Gera Body da API

        Args:
            sessionkey (str) : Chave da sessão
            
    """

    headers = {
    'sessionkey': f'{sessionkey}',
    'Content-Type': 'application/json'
    }

    return headers

 
msg =  'Olá, \n\n'
msg += 'Estamos realizando a atualização cadastral dos celulares corporativos da BR Samor. '
msg += 'Para que esse processo seja realizado, solicito que responda o formulário presente no link abaixo. \n\n' 
msg += 'Em caso de dúvidas, estou a disposição. \n\n'
msg += 'Obrigado.\n\n'
msg += 'Thiago Dias \n'
msg += 'Gerente de TI\n\n'
msg += 'https://forms.office.com/r/cHXif5pDi4'

data = PY_ReadWriteExcel.WorkbookRead()

url = 'http://192.168.85.132:3333/sendText'

counter = 0
counterError = 0

for indice in tqdm(data):

    payload = ApiPayload(session = 'SESSION1', number = indice, text = msg)
    headers = ApiHeaders('12345')
    
    try:

        response = requests.request("POST", url, headers = headers, data = payload, timeout = 2)

        if response.status_code == 200:   

            PY_ReadWriteExcel.WorkbookWrite(number = indice, value = '200 - OK')

            counter += 1

        else:

            PY_ReadWriteExcel.WorkbookWrite(number = indice, value = str(response.status_code) + ' - ' + response.text)

            counterError += 1

    except Timeout as err:

        PY_ReadWriteExcel.WorkbookWrite(number = indice, value = 'Server Connection Timeout')

        counterError += 1
           

print(f'\nSuccessfully Sent : {str(counter)}')
print(f'Sent With Error   : {str(counterError)}')