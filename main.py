from selenium import webdriver
import time
import pandas as pd
import credenciais
from bs4 import BeautifulSoup # biblioteca para fazer o webscrapping
import requests
import shutil
import os
from datetime import datetime


class sistema():
    # Inicio da função
    def __init__(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--start-maximized") # Maximizar a janela ao abrir 
        self.driver = webdriver.Chrome(options=options)

    # abre o site do sistema interno
    def abrir_site(self):
       self.driver.get('https://rpeotta.e-clic.net/Account/Login?ReturnUrl=%2f')
       print("abrir_site")
       time.sleep(4)
    
    #importando as crendencias e faz login no sistema
    def crendenciais(self):
        #caminho do campo login
        campo_login = self.driver.find_element("xpath",'//*[@id="LoginUserName"]') 
        #time.sleep(2)
        campo_login.click() 
        time.sleep(2)
        #Insere o login no campo login
        campo_login.send_keys(credenciais.login)
        #time.sleep(5)
        #Campo da senha é especificado
        campo_senha = self.driver.find_element("xpath",'//*[@id="LoginPassword"]') 
        campo_senha.click()
        #time.sleep(2)
        #Inserindo a senha no campo senha.
        campo_senha.send_keys(credenciais.senha)
        time.sleep(1)
        #Clica no para acessar o sistema
        botao_clique = self.driver.find_element("xpath",'/html/body/div[1]/div[1]/div[3]/div[1]/form/div/div[1]/button')
        botao_clique.click()
        print("Logando no sistema")
        time.sleep(5)
        #self.driver.quit()

    def extrair_relatorio(self):
        print("Extraindo o relatorio")
        self.driver.get('https://rpeotta.e-clic.net/documento/Report/relatorio_documento.aspx')
        time.sleep(4)

        campo_cliente = self.driver.find_element("xpath",'//*[@id="combocliente_chzn"]/a/span') # Instanciando o campo cliente
        campo_cliente.click()
        time.sleep(4)
        #Selecionando a opção "TODOS" 
        click_todos_cliente = self.driver.find_element("xpath",'//*[@id="combocliente_chzn_o_1"]')
        click_todos_cliente.click()
        time.sleep(5)
        #Clicando no campo "PROJETO"
        campo_projeto = self.driver.find_element('xpath','//*[@id="comboprojeto_chzn"]/a/span')
        campo_projeto.click()
        time.sleep(5)
        #Selecionando a opção "TODOS"
        click_todos_projeto = self.driver.find_element("xpath",'//*[@id="comboprojeto_chzn_o_1"]')
        click_todos_projeto.click()
        time.sleep(4)
        #Clicando na opção do disquete para exibir a opção de baixar
        campo_disquete = self.driver.find_element('xpath','//*[@id="ReportViewer1_ctl09_ctl04_ctl00_ButtonLink"]')
        campo_disquete.click()
        time.sleep(2)
        #Clicando na opção 
        botao_excel = self.driver.find_element('xpath','//*[@id="ReportViewer1_ctl09_ctl04_ctl00_Menu"]')
        time.sleep(2)
        botao_excel.click()
        print("Relatório baixado")
        time.sleep(10)
    #Extrair a tabela via html
    def web_scrapping(self):
        pageSource = self.driver.page_source
        bs = BeautifulSoup(pageSource, 'html.parser')
        print(bs.find(id='content').get_text())
        time.sleep(10)
    
    def manipulacao_arquivo(self):
        #Obtem a data e hora atual
        now = datetime.now()
        #Formata a data e a hora para incluir no nome do arquivo (ex: "2024-11-04_15-30-45")
        timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
                
        #Caminho do arquivo no downloads
        arquivo_origem = 'C:\\Users\\arthur.lopes\Downloads\\relatorio documento.xlsx'

        #Realizando a copia do arquivo para a pasta destino (solicitado pela Aline)
        #OBS: O arquivo precisa sempre se substituidp pelo atual 
        shutil.copy(arquivo_origem,'T:\\13.Planejamento\\01. LD e CR\\0000 - BASE DE DADOS\\relatorio_eclick.xlsx')
        print("Arquivo salvo na base de dados")

        #Salvando uma copia no pasta (OBSOLETO)
        destino_obsoleto = 'T:\\13.Planejamento\\01. LD e CR\\0000 - BASE DE DADOS\\Obsoleto\\'
        nome_arquivo_timestamp = f"relatorio_eclick_{timestamp}.xlsx"
        destino_arquivo_obsoleto = destino_obsoleto + nome_arquivo_timestamp 

        #shutil.copy(arquivo_origem,'T:\\13.Planejamento\\01. LD e CR\\0000 - BASE DE DADOS\\Obsoleto\\relatorio_eclick.xlsx')
        shutil.copy(arquivo_origem,destino_arquivo_obsoleto)
        print("Arquivo salvo na pasta obsoleto")

        time.sleep(5)
        #Apagando o arquivo baixado para não haver duplicação do nome
        os.remove(arquivo_origem)
        print("Arquivo deletado da pasta download")
        time.sleep(5)
        
    def fim_automação(self):
        print("Fim da automação")
        self.driver.quit()



        



        
sistema = sistema()
sistema.abrir_site() 
sistema.crendenciais()
sistema.extrair_relatorio()
#sistema.web_scrapping()
sistema.manipulacao_arquivo()
sistema.fim_automação()