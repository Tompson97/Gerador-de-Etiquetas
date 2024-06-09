# Biblioteca para construir a interface gráfica
import os
import uuid
import sys
import webbrowser
import PySimpleGUI as sg
from PIL import Image, ImageFont, ImageDraw

# Verifica se estamos executando o código como um script congelado (empacotado pelo PyInstaller)
if getattr(sys, 'frozen', False):
    # Se estivermos executando como um script congelado, o arquivo .png estará no diretório do executável
    png_path = os.path.join(sys._MEIPASS, 'oferta.png')
    fonte_arial = os.path.join(sys._MEIPASS,'arial_black.ttf')
    fonte_oferta = os.path.join(sys._MEIPASS,'fonte_good.ttf')
else:
    # Se não estivermos executando como um script congelado, o arquivo .png estará no diretório de trabalho atual
    png_path = 'oferta.png'
    fonte_arial = 'arial_black.ttf'
    fonte_oferta = 'fonte_good.ttf'

# Agora você pode abrir o arquivo .png usando o caminho em png_path
with open(png_path, 'rb') as f:
    pass  # Faça algo com o arquivo

# Declare as variáveis como globais
produto = None
preco = None
qtd = None
descricao = None

# Alterne seu tema para usar o recém-adicionado
sg.theme('Dark2')

layout = [
    [sg.Text('Insira o nome do produto  (Até 16 caracteres)')],
    [sg.Input(key='nome_produto')],
    [sg.Text('PORÇÃO (UND, PCT, KG,ETC)')],
    [sg.Input(key='descricao')],
    [sg.Text('Inseira o preço do produto.')],
    [sg.Input(key='preco_produto')],
    [sg.Text('Quantidade de etiquetas')],
    [sg.Input(key='qtd_etiquetas')],
    [sg.Button('Gerar')],
    [sg.Button('Ajuda')],
    [sg.Text('', key='mensagem')]
    ]

def edicao_etiqueta():
    global produto, preco, qtd, descricao
    imagem = Image.open(png_path)
    desenho = ImageDraw.Draw(imagem)
    largura_imagem, altura_imagem = imagem.size

    if  len(produto) <= 4:
        tam_nome = 190
        posi_nome = 300

    elif  len(produto) <= 6:
        tam_nome = 190
        posi_nome = 100

    elif len(produto) <= 8:
        tam_nome = 150
        posi_nome = 190

    elif len(produto) == 9:
        tam_nome = 130
        posi_nome = 110

    elif len(produto) == 10:
        tam_nome = 130
        posi_nome = 70

    elif len(produto) == 11:
        tam_nome = 100
        posi_nome = 120

    elif len(produto) == 12:
        tam_nome = 100
        posi_nome = 80

    elif len(produto) == 13:
        tam_nome = 100
        posi_nome = 80
    else:
        tam_nome = 80
        posi_nome = 85

    fonte_nome = fonte_arial
    fonte_nome = ImageFont.truetype(fonte_nome, tam_nome)
    cor_fonte_nome = '#1C0E0E'
    coord_nome = (posi_nome, 300)

    if preco >= 1000:
        posicao_preco = 0.79
        tam_preco = 290
        preco = float(preco)
        preco = "{:.2f}".format(preco)

    elif preco > 100:
        posicao_preco = 0.76
        tam_preco = 350
        preco = float(preco)
        preco = "{:.2f}".format(preco)

    elif preco > 20:
        posicao_preco = 0.80
        tam_preco = 420
        preco = float(preco)
        preco = "{:.2f}".format(preco)

    elif preco > 10:
        posicao_preco = 0.85
        tam_preco = 470
        preco = float(preco)
        preco = "{:.2f}".format(preco)

    elif preco < 10:
        posicao_preco = 0.85
        tam_preco = 500
        preco = float(preco)
        preco = "{:.2f}".format(preco)

    else:
        posicao_preco = 0.79
        tam_preco = 420
        preco = float(preco)
        preco = "{:.2f}".format(preco)

    fonte_preco = fonte_oferta
    fonte_preco = ImageFont.truetype(fonte_preco, tam_preco)
    cor_fonte_preco = (227, 58, 63)
    largura_preco = 5 * 150 // posicao_preco  # Ajuste este valor conforme necessário
    coord_preco = ((largura_imagem - largura_preco) / 2, 880)

    fonte_desc = r'arial_black.ttf'
    fonte_desc = ImageFont.truetype(fonte_desc, 100)
    cor_desc = '#1C0E0E'
    coord_desc = (430, 500)

    desenho.text(coord_nome, produto, font=fonte_nome, fill=cor_fonte_nome)
    desenho.text(coord_preco, str(preco).replace('.', ','), font=fonte_preco, fill=cor_fonte_preco)
    desenho.text(coord_desc, descricao.upper(), font=fonte_desc, fill=cor_desc)

    # Verifica se a pasta 'etiquetas' existe
    if not os.path.exists('etiquetas'):
        # Se não existir, cria a pasta 'etiquetas'
        os.makedirs('etiquetas')

    # Salva a imagem na pasta 'etiquetas'
    for i in range(qtd):
        imagem.save(f'etiquetas/{produto}_copia{i+1}.png')

    # Exibe a imagem
    imagem.show()
def interface():
    window = sg.Window('Gerador de Etiquetas', layout=layout, size=(300, 330))

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        elif event == 'Gerar':
            def gerar():
                caracter = len(str(values['nome_produto']))
                limite_preco = float(values['preco_produto'].replace(',', '.'))
                limite_qtd = int(values['qtd_etiquetas'])
                descricao = values['descricao']
                def converter_valores():
                    global produto, preco, qtd, descricao
                    try:
                        produto = values['nome_produto'].upper()
                        preco = float(values['preco_produto'].replace(',', '.'))
                        qtd = int(values['qtd_etiquetas'])
                        descricao = values['descricao']
                        edicao_etiqueta()
                        window['mensagem'].update('Etiqueta gerada com sucesso!')
                        return True
                    except Exception as e:
                        window['mensagem'].update('Dados inválidos')
                        return False
                    if not converter_valores():
                        window['mensagem'].update('Dados inválidos')
                if (caracter != 0 < 17 and limite_preco > 0 and limite_preco < 10000 and limite_qtd > 0
                        and limite_qtd <= 20 and descricao != '' and len(descricao) <= 5):
                    converter_valores()

                elif descricao == '' or len(descricao) > 5:
                    window['mensagem'].update('A descrição deve conter de 1 à 5 letras!')

                elif limite_qtd <= 0 or limite_qtd > 20:
                    window['mensagem'].update('A quantidade de etiquetas deve ser de 1 à 20.')

                elif caracter > 0 and caracter >= 17:
                    window['mensagem'].update('O nome do produto deve ter até 16 caracteres')

                elif limite_preco <= 0 or limite_preco >= 10000:
                    window['mensagem'].update('Preço limite de 0 até R$ 9.999,90')

                else:
                    window['mensagem'].update('Dados inválidos')
            try:
                gerar()
            except:
                window['mensagem'].update('Dados inválidos')

        elif event == 'Ajuda':
            webbrowser.open('https://drive.google.com/drive/folders/1NTtcKWVSfhTkiY9LiJvJPnQcmU1fw4x3?usp=sharing')

interface()