import openpyxl
import pyperclip
import pyautogui
from time import sleep


open = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet = open['Produtos']

for linha in sheet.iter_rows(min_row=2, max_row=2):
    nome_produto = (linha[0].value)
    pyperclip.copy(nome_produto)
    pyautogui.click(-537,174, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    descricao = (linha[1].value)
    pyperclip.copy(descricao)
    pyautogui.click(-491,278, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    categoria = (linha[2].value)
    pyperclip.copy(categoria)
    pyautogui.click(-197,396, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    codigo_produto = (linha[3].value)
    pyperclip.copy(codigo_produto)
    pyautogui.click(-120,482, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    peso = (linha[4].value)
    pyperclip.copy(peso)
    pyautogui.click(-166,568, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    dimensoes = (linha[5].value)
    pyperclip.copy(dimensoes)
    pyautogui.click(-154,655, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    pyautogui.click(-535,701, duration=1)
    sleep(3)
    
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(-344,199, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    quantidade = linha[7].value
    pyperclip.copy(quantidade)
    pyautogui.click(-337,288, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.click(-310,364, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(-296,457, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    tamanho = linha[10].value
    pyautogui.click(-321,544, duration=1)
    if tamanho == "Pequeno":
        pyautogui.click(-321.571, duration=1)
    elif tamanho == "Médio":
        pyautogui.click(-331,599, duration=1)
    else:
        pyautogui.click(-336,624, duration=1)
    
   
    
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(-269,633, duration=1)
    pyautogui.hotkey('ctrl', 'v')
  
    pyautogui.click(-534,689, duration=1)
    sleep(3)
    
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(-419,216, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(-401,302, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(-387,397, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    codigo_barras = linha[15].value
    pyperclip.copy(codigo_barras)
    pyautogui.click(-377,526, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(-371,616, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    pyautogui.click(-528,673, duration=1)
    pyautogui.click(-145,138, duration=1)
    pyautogui.click(-314,441, duration=1)
    sleep(3)