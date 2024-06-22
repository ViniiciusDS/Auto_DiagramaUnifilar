""" Automatizando o Quadro de Cargas com Python e AutoCAD """

# Importa as bibliotecas necessárias
import pandas as pd
from pyautocad import Autocad, APoint


# Lendo a planilha de cargas
df = pd.read_excel("Quadro de Cargas.xlsx")

# Inicializa a conexão com o AutoCAD
acad = Autocad(create_if_not_exists=True)
print(acad.doc.Name)

# Define o ponto inicial do quadro de cargas
start_point = APoint(0, 0)

# Filtrando os dados a partir da linha 6 e colunas B, C e D
df_filtered = df.iloc[4:15, [1, 2, 3]]
df_filtered.columns = ['Circuito', 'NumeroFase', 'EquilibrioFases']

# Espaçamento entre linhas e colunas
row_height = 10
col_width = 50

# Adicionar texto no AutoCAD
def add_text(text, point):
    acad.model.AddText(text, point, 2.5)

# Adiciona o cabeçalho do quadro de cargas
add_text("Quadro de Cargas", start_point )
add_text("Circuito", start_point + APoint(col_width, 10))
add_text("Número de Fases", start_point + APoint(col_width * 2, 10))
add_text("Equilíbrio de Fases", start_point + APoint(col_width * 3, 10))

# Adiciona os dados do quadro de cargas
for i, row in df_filtered.iterrows():
    add_text(row['Circuito'], start_point + APoint(0, -row_height * (i + 1)))
    add_text(str(row['NumeroFase']), start_point + APoint(col_width, -row_height * (i + 1)))
    add_text(str(row['EquilibrioFases']), start_point + APoint(col_width * 2, -row_height * (i + 1)))



print("Quadro de cargas criado com sucesso no AutoCAD.")
