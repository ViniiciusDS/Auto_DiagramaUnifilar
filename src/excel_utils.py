'''  Utilitários para manipulação de planilhas excel'''
import pandas as pd

def ler_planilha(planilha_path):
    df = pd.read_excel(planilha_path)
    df_filtered = df.iloc[5:, [1, 2, 3, 4]]
    df_filtered.columns = ['Circuito', 'NumeroFase', 'EquilibrioFases', 'Disjuntores']
    df_filtered = df_filtered.fillna('')
    return df_filtered
