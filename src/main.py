''' Script principal para criar um quadro de cargas e Diagrama unifilar no AutoCAD a partir de uma planilha Excel. '''

from pyautocad import Autocad, APoint
from autocad_utils import conectar_ou_criar_autocad, abrir_dwg, add_text
from excel_utils import ler_planilha
import os

def criar_quadro_de_cargas(planilha_path, dwg_path):
    acad_app = conectar_ou_criar_autocad()
    acad = Autocad(create_if_not_exists=True)

    if not os.path.exists(dwg_path):
        print(f"O arquivo {dwg_path} não foi encontrado.")
        return

    abrir_dwg(acad_app, dwg_path)

    df_filtered = ler_planilha(planilha_path)

    start_point = APoint(0, 0)
    row_height = 5
    col_width = 50

    add_text(acad, "Quadro de Cargas", start_point + APoint(100, 0))
    add_text(acad, "Circuito", start_point + APoint(col_width, -15))
    add_text(acad, "Número de Fases", start_point + APoint(col_width * 2, -15))
    add_text(acad, "Equilíbrio de Fases", start_point + APoint(col_width * 3, -15))
    add_text(acad, "Disjuntores", start_point + APoint(col_width * 4, -15))

    for i, row in df_filtered.iterrows():
        circuito = row['Circuito']
        numero_fase = row['NumeroFase']
        equilibrio_fases = row['EquilibrioFases']
        disjuntores = row['Disjuntores']

        add_text(acad, circuito, start_point + APoint(col_width, -row_height * (i + 1)))
        add_text(acad, numero_fase, start_point + APoint(col_width * 2, -row_height * (i + 1)))
        add_text(acad, equilibrio_fases, start_point + APoint(col_width * 3, -row_height * (i + 1)))
        add_text(acad, disjuntores, start_point + APoint(col_width * 4, -row_height * (i + 1)))

    print("Quadro de cargas criado com sucesso no AutoCAD.")

if __name__ == "__main__":
    planilha_path = "Quadro_de_Cargas.xlsx"
    dwg_path = os.path.abspath("DiagAuto.dwg")
    criar_quadro_de_cargas(planilha_path, dwg_path)
