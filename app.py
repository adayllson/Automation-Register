import openpyxl
import pyperclip
import pyautogui
from time import sleep

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

    pyautogui.click(1441,943, duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1471,420, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1539,503, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1490,593, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1474,670, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    if tamanho == 'Pequeno':
        pyautogui.click(1435,797, duration=1)
        
    elif tamanho == 'Médio':
        pyautogui.click(1439,814, duration=1)
        
    else:
        pyautogui.click(1437,837, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1431,855, duration=1)
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