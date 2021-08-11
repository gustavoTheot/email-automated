import pyautogui
import webbrowser
import time
import pandas as pd
from IPython.display import display

# Passo 1 - Entrar no link

## Abrindo o browser com o link determinado
webbrowser.open('https://drive.google.com/drive/folders/14oLE59U1RqyRqlBbKpsyymW-mitvbtoh')

# Passo 2 - Navegar pelo sistema até instalar o desejado 

## Navegando pelo sistema 
time.sleep(2) # esperar 3s
### Sempre que for utilizado o click, é preciso passar a posição x,y. Isso pode ser feito atraveis pyautogui.position()
pyautogui.click(x=2860, y=431) # clique no arquivo
time.sleep(1)
pyautogui.click(x=3744, y=214) # clique nas ações
time.sleep(2)
pyautogui.click(x=3473, y=720) # click em download 
time.sleep(4)
pyautogui.click(x=3769, y=888) # clickl em confirmar download

time.sleep(3) # esperar para fazer o donwload

# Passo 3 - Lê e interpretar o arquivo instalado 

## Interpretando dados
table = pd.read_excel(r'C:\Users\gusta\Downloads\Vendas - Dez.xlsx') # pd.read_tipo_do_arquivo = representando o modelo do arquivo. 'r' serve o pyautogui leia da forma que foi atribuido
financeiro = table['Valor Final'].sum() # entrando dentro do arquivo e somando a tabela valor
quantidade = table['Quantidade'] .sum() # entrando dentro do arquivo e somando a tabela quantidade
display(table) # disferente do print, ele mostra tabelas de forma mai organizada

time.sleep(3) # esperar a informação ser processada

# Passo 4 - Enviar informações atráves do email
pyautogui.hotkey('ctrl', 't') # hotkey -> quando for utilizar mais de 1 tecla
pyautogui.write('https://mail.google.com/mail/u/0/#inbox')
pyautogui.press('enter') # press -> quando for utilizar apenas 1 tecla
time.sleep(3) # esperar o email abrir 
pyautogui.click(x=2484, y=239)
time.sleep(1) # esperar abrir a mensagem
pyautogui.write('gustavotheotonio46@gmail.com') # write -> escrever
pyautogui.press('tab') 
pyautogui.press('tab')
time.sleep(1)
pyautogui.write('Email automatizado')
pyautogui.press('tab')
text = f'''Bom dia!
Quem fala é a pfizer. Estou aqui para informar que vacina salva vidas. 
Foram vendidas {quantidade:,.2f} pelo valor de R$ {financeiro:,.2f}
''' # 'f' diz que o texto tera variáveis 
pyautogui.write(text)
pyautogui.press('tab')
pyautogui.press('enter')



