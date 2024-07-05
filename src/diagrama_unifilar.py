''' Módulo responsável por criar um Diagrama Unifilar no AutoCAD a partir de uma planilha Excel.'''

from pyautocad import Autocad, APoint
from autocad_utils import conectar_ou_criar_autocad, abrir_dwg, add_text
from excel_utils import ler_planilha
import os

def criar_diagrama_unifilar(planilha_path, dwg_path):
    acad_app = conectar_ou_criar_autocad()
    acad = Autocad(create_if_not_exists=True)

    if not os.path.exists(dwg_path):
        print(f"Erro: O arquivo DWG '{dwg_path}' não foi encontrado.")
        return

    abrir_dwg(acad_app, dwg_path)

    