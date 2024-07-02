"""     Módulo com funções utilitárias para interação com o AutoCAD.    """

import os
import comtypes.client
from pyautocad import Autocad, APoint
import time

# Conecta ao AutoCAD existente ou cria uma nova instância
def conectar_ou_criar_autocad():
    try:
        acad = comtypes.client.GetActiveObject("AutoCAD.Application")
        print("Conectado ao AutoCAD existente.")
    except OSError:
        acad = comtypes.client.CreateObject("AutoCAD.Application")
        acad.Visible = True
        print("Nova instância do AutoCAD criada.")
    return acad

# Fecha o documento DWG aberto no AutoCAD
def fechar_documento_aberto(acad_app, dwg_path):
    for doc in acad_app.Documents:
        if doc.FullName.lower() == dwg_path.lower():
            doc.Close(False)
            print(f"Arquivo {dwg_path} fechado.")

# Abre um arquivo DWG no AutoCAD
def abrir_dwg(acad_app, dwg_path, tentativas=3, intervalo=2):
    if not os.path.exists(dwg_path):
        raise FileNotFoundError(f"O arquivo {dwg_path} não foi encontrado.")
    
    fechar_documento_aberto(acad_app, dwg_path)

    # Tenta abrir o arquivo DWG várias vezes, com um intervalo de tempo
    for tentativa in range(tentativas):
        try:
            acad_app.Documents.Open(dwg_path)
            print(f"Arquivo {dwg_path} aberto com sucesso.")
            return
        except Exception as e:
            print(f"Erro ao abrir o arquivo DWG na tentativa {tentativa + 1}: {e}")
            time.sleep(intervalo)
    
    raise Exception(f"Não foi possível abrir o arquivo {dwg_path} após {tentativas} tentativas.")

# Adiciona um texto no AutoCAD
def add_text(acad, text, point):
    acad.model.AddText(text, point, 2.5)
