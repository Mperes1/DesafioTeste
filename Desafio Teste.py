from openpyxl import Workbook
from selenium.webdriver import Firefox
from time import sleep

"""

Para o perfeito funcionamento do codigo, é necessario definir as sequintes variaveis corretamente.

Lista_Busca = É a lista de produtos que será procurado no site da Magazine Luiza
CaminhoArquivo = Caminho onde será salva a planilha com as informações dos produtos buscado. 
NomeArquivo = Deve ser informado o nome da planilha.

"""

Lista_busca = ['Samsung S21', 'Xaiomi 9T', 'Motorola', 'alexa']
Lista_NomeProduto = []
Lista_PrecoProduto = []
Lista_DescontoProduto = []
Lista_Links = []

CaminhoArquivo = 'C:/Users/Lauana/Desktop/Mauricio/Desafio Paschoalotto/'
NomeArquivo = 'Planilha_de_preço'

url = 'https://www.magazineluiza.com.br'
navegador = Firefox()
navegador.get(url)
sleep(1)


for item in Lista_busca:
    produto = Lista_busca[0]
    pesquisa = navegador.find_element_by_id('inpHeaderSearch')
    pesquisa.send_keys(item)
    navegador.find_element_by_id('btnHeaderSearch').click()
    sleep(1)

    navegador.find_element_by_xpath("/html/body/div[2]/div[3]/div[2]/div/div[3]/div/div[2]/ul/li[1]/a[1]").click()

    Lista_Links.append(navegador.current_url)

    BuscaNome = navegador.find_elements_by_tag_name('h1')
    Lista_NomeProduto.append(BuscaNome[0].text)

    BuscaPreco = navegador.find_elements_by_css_selector('.price-template__text')
    Lista_PrecoProduto.append(BuscaPreco[0].text)

    if navegador.find_elements_by_css_selector('.price-template__discount-text'):
        BuscaDesconto = navegador.find_elements_by_css_selector('.price-template__discount-text')
        DescontoProduto = BuscaDesconto[0].text
    else:
        DescontoProduto = 'Produto não possui desconto'

    Lista_DescontoProduto.append(DescontoProduto)

navegador.quit()
wb = Workbook()
celulares = wb.worksheets[0]

celulares['A1'] = 'Nome'
celulares['B1'] = 'Preço'
celulares['C1'] = 'Desconto'
celulares['D1'] = 'Link do produto'

celulares.title = 'Celulares'
Coluna = 2
for produtos in Lista_NomeProduto:
    celulares[f'A{Coluna}'] = produtos
    Coluna += 1

Coluna = 2
for links in Lista_Links:
    celulares[f'D{Coluna}'] = links
    Coluna += 1

Coluna = 2
for precos in Lista_PrecoProduto:
    celulares[f'B{Coluna}'] = f'R$ {precos}'
    Coluna += 1

Coluna = 2
for Descontos in Lista_DescontoProduto:
    celulares[f'C{Coluna}'] = Descontos
    Coluna += 1

wb.save(f"{CaminhoArquivo}{NomeArquivo}.xlsx")

print(f'Planilha salva com sucesso no caminho: {CaminhoArquivo}')
