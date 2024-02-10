import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
selected_page = workbook['Produtos']

for linha in selected_page.iter_rows(min_row=2):
    
    # Nome do Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1397,362,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # DescriçãoNenhuma
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1392,463,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1386,586,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Código do produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1382,670,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Peso (em kg)
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1382,759,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Dimensões (L x A x P)
    dimensao = linha[5].value
    pyperclip.copy(dimensao)
    pyautogui.click(1382,839,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # BOTÃO PRÓXIMO
    pyautogui.click(1397,901,duration=1)
    sleep(3)
    
    # Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1400,387,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Quantidade em estoque
    quantidade_estoque = linha[7].value
    pyperclip.copy(quantidade_estoque)
    pyautogui.click(1390,474,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Data de validade
    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.click(1382,565,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1396,651,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Tamanho
    tamanho = linha[10].value
    pyautogui.click(1368,731,duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(1400,763,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1400,786,duration=1)
    else:
        pyautogui.click(1400,808,duration=1)
    
    # Material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1419,823,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # BOTÃO PRÓXIMO
    pyautogui.click(1392,882,duration=1)
    sleep(3)
    
    # Fabricante
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1388,406,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Pais de Origem
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1385,489,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Observações
    observacao = linha[14].value
    pyperclip.copy(observacao)
    pyautogui.click(1385,604,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Codigo de Barras
    codigo_barras = linha[15].value
    pyperclip.copy(codigo_barras)
    pyautogui.click(1387,714,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Localização no Armazém
    localizacao = linha[16].value
    pyperclip.copy(localizacao)
    pyautogui.click(1385,795,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # BOTÃO CONCLUIR
    pyautogui.click(1402,863,duration=1)
    sleep(3)
    
    #BOTÃO OK
    pyautogui.click(1792,178,duration=1)
    sleep(3)
    
    #BOTÃO OK
    pyautogui.click(1620,638,duration=1)
    sleep(3)