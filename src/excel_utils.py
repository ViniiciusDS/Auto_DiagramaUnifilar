'''  Utilitários para manipulação de planilhas excel'''
import pandas as pd
import os

def ler_planilha(planilha_path):
    df = pd.read_excel(planilha_path)
    df_filtered = df.iloc[5:, [1, 2, 3, 4]]
    df_filtered.columns = ['Circuito', 'NumeroFase', 'EquilibrioFases', 'Disjuntores']
    df_filtered = df_filtered.fillna('')
    return df_filtered



# Função para Identificar e Filtrar quantos circuitos há no Quadro de Cargas
def Filtro_QTD_Circuitos(planilha_path):
    # Carrega a planilha
    df = pd.read_excel(planilha_path, header=None)

    # Identifica a célula "Circuitos" na coluna B
    inicio_circuitos = df[df[1] == 'CIRCUITO'].index[0]
    
    # Identifica a célula "Totais" mesclada de B até P
    fim_circuitos = df[df[1] == 'TOTAIS'].index[0]
    
    # Extrai os dados dos circuitos entre "Circuitos" e "Totais"
    circuitos_df = df.iloc[inicio_circuitos + 1:fim_circuitos, 1].dropna().reset_index(drop=True)
    
    # Remove as linhas que não possuem informações úteis
    circuitos_df = circuitos_df.dropna(how='all').reset_index(drop=True)
    
    return circuitos_df

# Função para Identificar e Filtrar o Número de Fase de cada circuito
def Filtro_Fase(planilha_path):
    # Carrega a planilha
    df = pd.read_excel(planilha_path, header=None)

    # Identifica a célula Disjuntores na coluna E
    inicio_fase = df[df[2] == 'N DE FASES'].index[0]

    # Identifica a célula "Totais" mesclada de B até P
    fim_fase = df[df[1] == 'TOTAIS'].index[0]

    # Extrai os dados dos disjuntores entre "Disjuntores" e "Totais"
    fase_df = df.iloc[inicio_fase + 1:fim_fase, 2].dropna().reset_index(drop=True)

    # Remove as linhas que não possuem informações úteis
    fase_df = fase_df.dropna(how='all').reset_index(drop=True)

    return fase_df

# Função para Identificar e Filtrar o Equilibrio de Fases de cada circuito
def Filtro_Equi_Fase(planilha_path):
    # Carrega a planilha
    df = pd.read_excel(planilha_path, header=None)

    # Identifica a célula Disjuntores na coluna E
    inicio_equi_fase = df[df[3] == 'EQUILIBRIO\nDE FASES'].index[0]

    # Identifica a célula "Totais" mesclada de B até P
    fim_equi_fase = df[df[1] == 'TOTAIS'].index[0]

    # Extrai os dados dos disjuntores entre "Disjuntores" e "Totais"
    equi_fase_df = df.iloc[inicio_equi_fase + 1:fim_equi_fase, 3].dropna().reset_index(drop=True)

    # Remove as linhas que não possuem informações úteis
    equi_fase_df = equi_fase_df.dropna(how='all').reset_index(drop=True)

    return equi_fase_df

# Função para Identificar e Filtra os Disjuntores de cada circuito
def Filtro_Disjuntores(planilha_path):
    # Carrega a planilha
    df = pd.read_excel(planilha_path, header=None)

    # Identifica a célula Disjuntores na coluna E
    inicio_disjuntores = df[df[4] == 'DISJUNTOR (A)'].index[0]

    # Identifica a célula "Totais" mesclada de B até P
    fim_disjuntores = df[df[1] == 'TOTAIS'].index[0]

    # Extrai os dados dos disjuntores entre "Disjuntores" e "Totais"
    disjuntores_df = df.iloc[inicio_disjuntores + 1:fim_disjuntores, 4].dropna().reset_index(drop=True)

    # Remove as linhas que não possuem informações úteis
    disjuntores_df = disjuntores_df.dropna(how='all').reset_index(drop=True)

    # Adiciona o Sufixo "A" para os disjuntores
    disjuntores_df = disjuntores_df.astype(str) + 'A'

    return disjuntores_df

# Função para Identificar e Filtra a Corrente de Curto de cada circuito
def Filtro_Corrente_Curto(planilha_path):
    # Carrega a planilha
    df = pd.read_excel(planilha_path, header=None)

    # Identifica a célula Disjuntores na coluna E
    inicio_corrente_curto = df[df[6] == 'CORRENTE DE CURTO'].index[0]

    # Identifica a célula "Totais" mesclada de B até P
    fim_corrente_curto = df[df[1] == 'TOTAIS'].index[0]

    # Extrai os dados dos disjuntores entre "Disjuntores" e "Totais"
    corrente_curto_df = df.iloc[inicio_corrente_curto + 1:fim_corrente_curto, 6].dropna().reset_index(drop=True)

    # Remove as linhas que não possuem informações úteis
    corrente_curto_df = corrente_curto_df.dropna(how='all').reset_index(drop=True)

    return corrente_curto_df

# Função para Identificar e Filtrar a curva de disjuntores de cada circuito
def Filtro_Curva_Disj(planilha_path):
    # Carrega a planilha
    df = pd.read_excel(planilha_path, header=None)

    # Identifica a célula Disjuntores na coluna E
    inicio_curva_disj = df[df[7] == 'CURVA'].index[0]

    # Identifica a célula "Totais" mesclada de B até P
    fim_curva_disj = df[df[1] == 'TOTAIS'].index[0]

    # Extrai os dados dos disjuntores entre "Disjuntores" e "Totais"
    curva_disj_df = df.iloc[inicio_curva_disj + 1:fim_curva_disj, 7].dropna().reset_index(drop=True)

    # Remove as linhas que não possuem informações úteis
    curva_disj_df = curva_disj_df.dropna(how='all').reset_index(drop=True)

    return curva_disj_df

# Função para Identinficar e Filtra a fiação de cada circuito
def Filtro_Fiacao(planilha_path):
    # Carrega a planilha
    df = pd.read_excel(planilha_path, header=None)

   # Preenche as células mescladas com o valor correto
    df = df.ffill(axis='columns')
    
    # Identifica a célula Disjuntores na coluna E
    inicio_fiacao = df[df[10] == 'FASE'].index[0]

    # Identifica a célula "Totais" mesclada de B até P
    fim_fiacao = df[df[1] == 'TOTAIS'].index[0]

    # Extrai os dados dos disjuntores entre "Disjuntores" e "Totais"
    fiacao_df = df.iloc[inicio_fiacao + 1:fim_fiacao, 10].dropna().reset_index(drop=True)

    # Remove as linhas que não possuem informações úteis
    fiacao_df = fiacao_df.dropna(how='all').reset_index(drop=True)

    return fiacao_df

# Exemplo de uso
if __name__ == "__main__":
    planilha_path = os.path.abspath("Quadro_de_Cargas.xlsx")
    #circuitos = Filtro_QTD_Circuitos(planilha_path)
    #print(circuitos)
    #fase = Filtro_Fase(planilha_path)
    #print(fase)
    #equi_fase = Filtro_Equi_Fase(planilha_path)
    #print(equi_fase)
    #disjuntores = Filtro_Disjuntores(planilha_path)
    #print(disjuntores)
    #corrente_curto = Filtro_Corrente_Curto(planilha_path)
    #print(corrente_curto)
    #curva_disj = Filtro_Curva_Disj(planilha_path)
    #print(curva_disj)
    #fiacao = Filtro_Fiacao(planilha_path)
    #print(fiacao)
    #df = pd.read_excel(planilha_path)
    #print(df)

    



