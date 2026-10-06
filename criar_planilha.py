from openpyxl import Workbook
from openpyxl.styles import Alignment

planilha_funcionarios = Workbook()
pagina = planilha_funcionarios.active

#Criar cabeçalho
pagina.append(['Nome', 'Email', 'Setor'])

#Centralizar células do cabeçalho
for celula in pagina[1]:
    celula.alignment = Alignment(horizontal="center") 

#Cadastrar funcionarios
funcionarios = [
    ("Ana Souza", "ana.souza@example.com", "Vendas"),
    ("Bruno Lima", "bruno.lima@example.com", "Marketing"),
    ("Carla Mendes", "carla.mendes@example.com", "Financeiro"),
    ("Diego Alves", "diego.alves@example.com", "Vendas"),
    ("Eduarda Rocha", "eduarda.rocha@example.com", "Logística"),
    ("Felipe Costa", "felipe.costa@example.com", "Marketing"),
    ("Gabriela Nunes", "gabriela.nunes@example.com", "Financeiro"),
    ("Henrique Dias", "henrique.dias@example.com", "Diretoria"),
]

#Formatar tamanho da coluna
pagina.column_dimensions["A"].width = 22
pagina.column_dimensions["B"].width = 34
pagina.column_dimensions["C"].width = 16

#Colocar na planilha
for funcionario in funcionarios:
    pagina.append(funcionario)

planilha_funcionarios.save('funcionarios.xlsx')