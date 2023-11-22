import openpyxl
import pyperclip
import pyautogui

# Entrar na planilha 
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):

    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1462,398, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1459,476, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1453,617, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1453,706, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1447,793, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1455,871, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    break
    # codigo_produto = linha[3].value
    # peso = linha[4].value
    # dimensoes = linha[5].value
    # preco = linha[6].value
    # quantidade_em_estoque = linha[7].value
    # data_de_validadde = linha[8].value
    # cor = linha[9].value
    # tamanho = linha[10].value
    # material = linha[11].value
    # fabricante = linha[12].value
    # observacoes = linha[13].value
    # codigo_de_barras = linha[15].value
    # localizacao_armazem = linha[16].value