import openpyxl
import pyautogui
import time
import webbrowser

pyautogui.alert('O processo de Automação será iniciado, favor não mexer.')

workbook = openpyxl.load_workbook (r'C:\Automação de Processos\produtos.xlsx')
sheet_produtos = workbook['produtos']
for linha in sheet_produtos.iter_rows(min_row=2, max_row=501):
    produto = linha[0].value
    fornecedor = linha[1].value
    categoria = linha[2].value
    quantidade = linha[3].value
    valor_unitario= linha[4].value
    notificarvenda = linha[5].value
    print(linha[0].value,linha[1].value,linha[2].value,linha[3].value,linha[4].value,linha[5].value)
    pyautogui.click(1371,404,duration=1)
    pyautogui.write(produto)
    pyautogui.click(1677,421,duration=1)
    pyautogui.write(fornecedor)
    pyautogui.click(1367,541,duration=1)
    pyautogui.write(categoria)
    pyautogui.click(1730,562,duration=1)
    pyautogui.write(valor_unitario)
# The `if` statement is used to check a condition. If the condition is true, the code inside the `if`
# block will be executed. In this case, the condition being checked is `notificar_venda == "Sim"`. If
# the value of `notificar_venda` is equal to "Sim", the variable `result` will be set to `True`.
    if notificarvenda == "Sim":  
        pyautogui.click(1224,661,duration=1)
    elif notificarvenda == "Não": 
        pyautogui.click(1411,763,duration=1)
    pyautogui.click(1377,742,duration=1)
    pyautogui.click(1793,232,duration=1)