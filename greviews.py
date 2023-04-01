import googlemaps
from unidecode import unidecode
from datetime import datetime
import openpyxl
from tkinter import Tk
from tkinter.filedialog import asksaveasfilename
from googletrans import Translator

# Coloque sua chave da API do Google Maps aqui
gmaps = googlemaps.Client(key='AIzaSyCO6rPf01s9fvy7LnT3lrTFQrDL66IvkZg')

# Solicite que o usuário insira a cidade e o nome do estabelecimento
cidade = input('Digite a cidade: ').lower()
cidade = unidecode(cidade)
nome_estabelecimento = input('Digite o nome do estabelecimento: ').lower()
nome_estabelecimento = unidecode(nome_estabelecimento)

# Pesquise pelo estabelecimento específico na cidade especificada
place_result = gmaps.places(f'{nome_estabelecimento} {cidade}')

if not place_result['results']:
    print(f'Nenhum estabelecimento encontrado com o nome "{nome_estabelecimento}" na cidade de "{cidade}".')
else:
    # Obtenha o ID do lugar
    place_id = place_result['results'][0]['place_id']

    # Obtenha as avaliações do lugar
    reviews = gmaps.place(place_id)['result']['reviews']

    # Classifique as avaliações por data (mais recente primeiro)
    reviews.sort(key=lambda x: x['time'], reverse=True)

# Crie uma nova planilha do Excel
wb = openpyxl.Workbook()
ws = wb.active

# Adicione os cabeçalhos das colunas
ws.append(['Autor', 'Avaliação', 'Texto', 'Data'])

# Solicite que o usuário insira o número de avaliações a serem mostradas
num_reviews = int(input('Digite o número de avaliações a serem mostradas (0 para mostrar todas): '))

# Adicione as avaliações à planilha
translator = Translator()
for i, review in enumerate(reviews):
    # Pare de adicionar avaliações se o número desejado for atingido
    if num_reviews != 0 and i >= num_reviews:
        break

    # Converta o timestamp em uma data legível
    date = datetime.fromtimestamp(review['time']).strftime('%Y-%m-%d')

    # Traduza o texto da avaliação para o português brasileiro
    translated_text = translator.translate(review['text'], dest='pt').text

    ws.append([review['author_name'], review['rating'], translated_text, date])

# Crie uma instância de Tkinter e oculte a janela principal
root = Tk()
root.withdraw()

# Abra a caixa de diálogo para selecionar o local do arquivo
file_path = asksaveasfilename(defaultextension='.xlsx')

# Salve a planilha no local selecionado pelo usuário
if file_path:
    wb.save(file_path)