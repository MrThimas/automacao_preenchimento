import openpyxl 
import pyperclip
import pyautogui
import time

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(872,235, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(853,318, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(873,428, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(883,505, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(868,584, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    dimensao = linha[5].value
    pyperclip.copy(dimensao)
    pyautogui.click(906,662, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    #Proximo
    pyautogui.click(856,718, duration = 1)
    time.sleep(5)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(871,243, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    quant_estoque = linha[7].value
    pyperclip.copy(quant_estoque)
    pyautogui.click(897,322, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.click(871,397, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(875,475, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyautogui.click(891,570, duration = 1)

    if tamanho == 'Pequeno':
        pyautogui.click(892,594, duration = 1)
    elif tamanho == 'Médio':
        pyautogui.click(892,614, duration = 1)
    else:
        pyautogui.click(892, 633, duration = 1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(863,642, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    #Proximo novamente
    pyautogui.click(852,698, duration = 1)
    time.sleep(5)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(866,272, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(877,351, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(879,435, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(881,549, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(876,623, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    #Concluir totalmente
    pyautogui.click(857,682, duration = 1)
    pyautogui.click(1226,192, duration = 2)
    time.sleep(5)
    pyautogui.click(1043,474, duration = 1)
    time.sleep(5)
