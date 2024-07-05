''' Script principal que chama as funções para criar o quadro de cargas e o diagrama unifilar. '''

import os
from quadro_de_cargas import criar_quadro_de_cargas
from diagrama_unifilar import criar_diagrama_unifilar

if __name__ == "__main__":
    planilha_path = os.path.abspath("Quadro_de_Cargas.xlsx")
    dwg_path = os.path.abspath("DiagAuto.dwg")
    
    # Chama a função para criar o quadro de cargas
    criar_quadro_de_cargas(planilha_path, dwg_path)
    
    # Chama a função para criar o diagrama unifilar
    criar_diagrama_unifilar(planilha_path, dwg_path)
