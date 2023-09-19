import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('Programavendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_row=2):
    pyautogui.click(935,516,duration=1.5)
    pyautogui.write(linha[0].value)
    pyautogui.click(956,545,duration=1.5)
    pyautogui.write(linha[1].value)
    pyautogui.click(939,568,duration=1.5)
    pyautogui.write(str(linha[2].value))
    pyautogui.click(994,588,duration=1.5)
    pyautogui.write(linha[3].value) 
    pyautogui.click(875,621,duration=1.5)
    pyautogui.click(945,585,duration=1.5)
    linha [0].value
    linha [1].value
    linha [2].value
    linha [3].value