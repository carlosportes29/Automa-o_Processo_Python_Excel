import pyautogui
import time
import openpyxl
#import pandas as pd

pyautogui.alert('O processo de Automação será iniciado, favor não mexer.')

# pyautogui.write -> escrever um texto
# pyautogui.press -> apertar 1 tecla
# pyautogui.click -> clicar em algum lugar da tela
# pyautogui.hotkey -> combinação de teclas
pyautogui.PAUSE = 0.6

# abrir o navegador (chrome)
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
time.sleep(1)

# entrar no link 
pyautogui.write("https://cadastroprodutos-devaprender.netlify.app/")
pyautogui.press("enter")
time.sleep(1)

# Passo 3: Importar a base de produtos pra cadastrar

#tabela = pd.read_excel("produtos_teste.xlsx")

workbook = openpyxl.load_workbook (r'C:\Automa-o_Processo_Python_Excel\produtos_teste.xlsx')
sheet_produtos = workbook['produtos']

# Passo 4: Cadastrar um produto Excel x Web
for linha in sheet_produtos.iter_rows(min_row=2, max_row=501):
        produto = linha[0].value
        fornecedor = linha[1].value
        categoria = linha[2].value
        quantidade = linha[3].value
        valor_unitario= linha[4].value
        notificarvenda = linha[5].value
        # Produto
        pyautogui.click(x=735, y=471)
        pyautogui.write(produto)
        # Fornecedor
        pyautogui.click(x=1030, y=471)
        pyautogui.write(fornecedor)
        # Categoria
        pyautogui.click(x=699, y=598)
        pyautogui.write(categoria)
        # valor_unitario
        pyautogui.click(x=1126, y=592)
        pyautogui.write(valor_unitario)
        # notificar a venda sim ou não
        if notificarvenda == "Sim":  
            pyautogui.click(x=622, y=706)
        elif notificarvenda == "Não":
            # click na notificação
            pyautogui.click(753,788)
        # click no botão confirmar
        pyautogui.click(777,791)
        # clicar no OK após o cadastro
        pyautogui.click(1182,209)
        time.sleep(1)