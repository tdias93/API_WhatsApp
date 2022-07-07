from openpyxl import load_workbook
from datetime import datetime

import os
import time

def WorkbookWrite(number, value = '200 - OK'):

    """ 
        Input de dados na planilha

        Args:
            event (str) : Evento Executado

    """

    # Registra Data e Hora do Evento
    registry = datetime.now()
    registry = registry.strftime('%d/%m/%Y - %H:%M:%S')

    plan = load_workbook(f"{os.path.dirname(os.path.realpath(__file__))}\\system\\LinhasVivo.xlsx")         
    dados = plan['LinhasVivo']                

    for cell in dados:

        if str(cell[0].value) == number and str(cell[1].value) == 'None':
            dados[f'B{cell[0].row}'] = value
            dados[f'C{cell[0].row}'] = registry

            break

    plan.save(f"{os.path.dirname(os.path.realpath(__file__))}\\system\\LinhasVivo.xlsx")

    time.sleep(2)


def WorkbookRead():

    """ 
        Coleta dados da planilha

            Returns:
                number : Lista com os números coletados.
    """
    
    planValue = []
    header = True

    plan = load_workbook(f"{os.path.dirname(os.path.realpath(__file__))}\\system\\LinhasVivo.xlsx")
    plan = plan['LinhasVivo']

    for cell in plan.values:

        if header:
            header = False

        elif str(cell[0]) == 'None':
            next

        else:
            planValue.append(str(cell[0]))
                            
    return planValue
    