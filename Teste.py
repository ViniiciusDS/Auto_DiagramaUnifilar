from pyautocad import Autocad, APoint

# Inicializa a conexão com o AutoCAD
acad = Autocad(create_if_not_exists=True)
print(acad.doc.Name)

# Define pontos para a linha
p1 = APoint(0, 0)
p2 = APoint(100, 100)

# Adiciona a linha ao documento ativo
acad.model.AddLine(p1, p2)
print("Linha criada de (0,0) até (100,100)")