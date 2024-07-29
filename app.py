"""

Para mandar mensagem direta pra pessoa tem que usar o whatsapp web
Lá vai ter uma URL que vai conseguir enviar a mensagem

Links e dados da planilha são falsos
Projeto usado pra teste de código

Use o comando que evita erro na hora de enviar no terminal
Comando "pip install pillow"

"""

import openpyxl
import pyautogui

#Importando criptografor para conseguir mandar mensagem usando link personalizado
from urllib.parse import quote

#Abrindo o navegador
import webbrowser

#Importando dormencia para dar tempo para logar
from time import sleep




"""
Caso a pessoa não esteja logada
Fazer o login com o código abaixo
Isso é só para ele não esteja logado
Se estiver o programa vai ficar parado na tela pelo mesmo tempo
"""

webbrowser.open('https://web.whatsapp.com/')
sleep(300)
#     300 segundo = 5 min


workBook = openpyxl.load_workbook('clientes.xlsx')
#Abrir a página certa da planilha
workSheet = workBook['Sheet1']
for linha in workSheet.iter_rows(min_row=2, max_row=2):
    #Nome
    nome =linha[0].value
    #Telefone
    telefone = linha[1].value
    #Data de vencimento
    dataVencimento = linha[2].value
    #Mensagem que vai ser mandada

    mensagem = f'Olá {nome} seu boleto vence no dia {dataVencimento.strftime('%d/%m/%Y')}'
    #Formatando a mensagem para conseguir mandar no Whatsapp
    
    #Aqui é onde pode ocorrer o erro de que um dado não foi formatado direito ou não conseguiu mandar a mensagem

    try:
        #Link especial
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        
        """
        Exemplo de numero de telefoneo
        5511978412433
        55 -> numero do país que está
        11 -> DDD da região
            978412433 -> numero de telefone
        Tudo tem que estar formatado para acessar o contato certo
        """
        #Abrindo o whatsapp web
        webbrowser.open(link_mensagem_whatsapp)

        #Pausa para carregar a página
        sleep(60)
        #Vai achar o botão de envio na tela
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(60)
        #Vai clicar no botão
        pyautogui.click(seta[0], seta[1])
        sleep(60)
        #Fechar a quia no navegador chrome
        pyautogui.hotkey('ctrl','w')
        sleep(60)
    except:
        print(f'Erro a mandar a mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')
