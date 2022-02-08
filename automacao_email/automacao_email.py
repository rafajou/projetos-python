# %%

import pyautogui
import pyperclip
import pandas as pd
import time
import os

# %% [markdown]
# 1º Passo - Acessar o sistema e baixar a base de dados

# %%
pyautogui.hotkey('win', 'r')
pyautogui.write('chrome')
pyautogui.press('enter')
# %%
if os.path.exists(r'C:\Users\Rafael\Documents\Vendas - Dez.xlsx') == False:
    time.sleep(3)
    pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.click(x=550, y=465, clicks=2)
    time.sleep(3)
    pyautogui.click(x=554, y=589, button='right')
    time.sleep(2)
    pyautogui.click(x=755, y=931)
    time.sleep(2)
    pyautogui.sleep(5)
    pyautogui.click(x=495, y=64)
    pyperclip.copy(r"C:\Users\Rafael\Documentos")
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    pyautogui.click(x=690, y=648)

# %% [markdown]
# 2º Passo - Importar a base de dados para o Python

# %%
time.sleep(5)
tabela = pd.read_excel(r'C:\Users\Rafael\Documents\Vendas - Dez.xlsx')
#display(tabela)
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()
#display(faturamento)
#display(quantidade)

# %% [markdown]
# 3º Passo - Enviar email para a equipe

# %%
pyautogui.hotkey('ctrl', 't')
time.sleep(2)
pyperclip.copy('http://gmail.com.br')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(2)
pyautogui.click(x=140, y=307)
time.sleep(1)
pyperclip.copy('rafaelgrisantedallaqua@hotmail.com')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('tab')
pyperclip.copy('Relatório de Vendas')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

texto = f"""Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Abs
Rafael Grisante Dallaqua"""

pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('ctrl', 'enter')

# %%
time.sleep(5)
pyautogui.position()

# %%



