import re, glob
from openpyxl import Workbook
from PyPDF2 import PdfReader
import tkinter as tk
from tkinter import filedialog

# Função que retorna o caminho da pasta selecionada para a busca dos boletins
def selecionar_pasta():
    root = tk.Tk()
    root.withdraw() # Oculta a janela principal
    pasta_selecionada = filedialog.askdirectory()  # Exibe a caixa de diálogo para selecionar uma pasta

    if pasta_selecionada:
        print("Pasta selecionada:", pasta_selecionada)
        return pasta_selecionada
    else:
        exit("Nenhuma pasta selecionada.")

# Função que extrai os campos de interesse com a data a partir do texto extraído do PDF
def extraiDados(texto, campos_busca):
    resultados = {}  # Dicionário para armazenar os resultados com o nome da tabela como chave
    
    for campo in campos_busca:
        # Regex que busca pela data no texto extraído do PDF
        data = re.search('(\d{2}-\d{2}-\d{4})[\s]*[/][\s]*Semana\s\d+', texto)
        # Regex que busca pelos campos de interesse no texto extraído do PDF
        dados = re.search(f"({campo}).\s(-?\d+[,]\d+)\s(-?\d+[,]\d+[%]?)\s\/chevron-\w+\s(-?\d+[,]\d+[%]?)\s\/chevron-\w+\s(-?\d+[,]\d+[%]?)", texto)
        
        # Dicionário com os campos de interesse
        dicionario = {
            'data': data.group(1),
            'indice': dados.group(2),
            'variacao_semanal': dados.group(3),
            'variacao_mensal': dados.group(4),
            'variacao_anual': dados.group(5)
        }
        resultados[campo] = dicionario  # Adiciona os campos de interesse ao dicionário usando o nome da tabela como chave
    
    return resultados

# Criar um novo workbook e remove a planilha padrão "Sheet"
workbook = Workbook()
workbook.remove(workbook['Sheet'])

campos_interesse = ['Convencional Trimestre', 'Convencional Longo Prazo', 'Incentivada 50% Trimestre', 'Incentivada 50% Longo Prazo']

# Cria uma planilha para cada campo de interesse
for campo in campos_interesse:
    # Cria a planilha com os nomes dos campos
    planilha = workbook.create_sheet(campo)
    
    # Adiciona cabeçalhos
    cabecalhos = ['Data', 'Índice R$/MWh', 'Variação Semanal', 'Variação Mensal', 'Variação Anual']
    planilha.append(cabecalhos)

# Pega a pasta selecionada com o padrão de arquivos para pdf e guarda o caminho para cada arquivo da pasta
pasta_selecionada = selecionar_pasta()
padrao_arquivos = pasta_selecionada + '/*.pdf'
caminhos_arquivos = glob.glob(padrao_arquivos)

# Faz a iteração em cada arquivo no diretório informado e extrai os dados
for caminho in caminhos_arquivos:
    reader = PdfReader(caminho) # Lê cada caminho de pdf na lista de caminhos
    texto = reader.pages[0].extract_text() # Extrai o texto de cada pdf
    dados = extraiDados(texto, campos_interesse) # Extrai os campos de interesse de cada pdf
    
    # Adiciona os dados na planilha correspondente
    for campo, resultado in dados.items():
        planilha = workbook[campo]
        planilha.append(list(resultado.values()))  # Adiciona os valores como uma linha na planilha

# Salvar e fechar o workbook
workbook.save('dados_agrupados.xlsx')
workbook.close()
