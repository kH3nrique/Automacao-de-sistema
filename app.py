#entrar na planilha
import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

#copiar informações de um campo em um seu correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    produtoNome = linha[0].value
    pyperclip.copy(produtoNome)
    pyautogui.click(852,225, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(850,308, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(845,428, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codProduto = linha[3].value
    pyperclip.copy(codProduto)
    pyautogui.click(845,509, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(847,582, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    dimencoes = linha[5].value
    pyperclip.copy(dimencoes)
    pyautogui.click(847,664, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(821,722, duration=1)
    sleep(1)
    
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(849,250, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    quantidade = linha[7].value
    pyperclip.copy(quantidade)
    pyautogui.click(848,331, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    validade = linha[8].value
    pyperclip.copy(validade)
    pyautogui.click(851,399, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(856,482, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    tamanho = linha[10].value
    pyautogui.click(846,558, duration=1)
    #ler informação da planilha
    #clicar no respectivo tamanho
    if tamanho == 'Pequeno':
        pyautogui.click(846,593, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(846,622, duration=1)
    else:
        pyautogui.click(846,652, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(843,637, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(829,688, duration=1)
    sleep(1)
    
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(850,254, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    paisOrigem = linha[13].value
    pyperclip.copy(paisOrigem)
    pyautogui.click(850,329, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(853,415, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    codBarra = linha[15].value
    pyperclip.copy(codBarra)
    pyautogui.click(854,548, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    localizacao = linha[16].value
    pyperclip.copy(localizacao)
    pyautogui.click(853,635, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    pyautogui.click(825,677, duration=1)
    sleep(1)

    pyautogui.click(1190,453, duration=1)
    sleep(1)
    
    pyautogui.click(1032,470, duration=1)
    sleep(1)