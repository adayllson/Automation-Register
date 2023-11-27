import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha 
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):

    # Nome do Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1462,398, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    # Descrição
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1459,476, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1453,617, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Código do Produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1453,706, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1447,793, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1455,871, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão Próximo
    pyautogui.click(1441,943, duration=1)
    sleep(3)

    # Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1471,420, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1539,503, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1490,593, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1474,670, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Tamanho
    tamanho = linha[10].value
    pyautogui.click(1511,760, duration=1)

    if tamanho == 'Pequeno':
        pyautogui.click(1435,797, duration=1)
        
    elif tamanho == 'Médio':
        pyautogui.click(1439,814, duration=1)
        
    else:
        pyautogui.click(1437,837, duration=1)

    # Material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1431,855, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão Próximo página 2
    pyautogui.click(1434,915, duration=1)
    sleep(3)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1448,441, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_de_origem = linha[13].value
    pyperclip.copy(pais_de_origem)
    pyautogui.click(1425,526, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1452,608, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1443,741, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1482,820, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão Concluir
    pyautogui.click(1428,893, duration=1)

    # Botão Confirmar Inclusão
    pyautogui.click(1824,213, duration=1)

    # Inicia um novo cadastro
    pyautogui.click(1641,672, duration=1)


        

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