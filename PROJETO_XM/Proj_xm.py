# PROJETO :  EXTRATOR VOLUME DE VENDAS XM 
# VERSÃO : BETA 
# AUTHOR: ALAN COSTA
#

# IMPORTAÇÕES DAS BIBLIOTECAS NECESSÁRIAS PARA O FUNCIONAMENTO DO EXTRATOR
from selenium import webdriver
from bs4 import BeautifulSoup
import itertools
import os
import xlwings as xw
from selenium.webdriver import Ie
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import pandas as pd
import time
from datetime import datetime,date,timedelta
dateTimeObj = datetime.now()
hora2 = dateTimeObj.strftime("%H:%M:%S")
url='http://xmlink.grupopetropolis.com.br:8080/XMLinkWM/relatorios/consolidado_det.aspx'
login = ''
senha = ''
login = login
senha = senha
#FUNÇÃO QUE EXTRAI AS INFORMAÇÕES DO VOLUME DE VENDAS 
def xmlink():
    
    #se por acaso queira pegar o dia anterior utiliza a variavel passado no campo dataini e fim
    hoje = date.today()
    intervalo = timedelta(1)
    passado = hoje - intervalo
    passado = passado.strftime("%d/%m/%Y")
    
    hoje = hoje.strftime("%d/%m/%Y")
    
    revendas = pd.DataFrame()
    COLUNAS = [
    'REVENDA',
'HL'
    ]
   
    und = ['2','1','5','4','6','7','8','9','10' ]#ENUMERA DE ACORDO COM A QUANTIDADE DE CADA REVENDA DA SUA REGIONAL
    #ARRAY COM O NOME DE CADA REVENDA , PARA GRAVAR NO DATA SET 
    rev = ['435 - BARREIRAS','431 - BOM JESUS DA LAPA','375 - BRUMADO','414 - VITÓRIA DA CONQUISTA','422 - EUNÁPOLIS','425 - JEQUIÉ','202 - TEIXEIRA DE FREITAS','414 - VITÓRIA DA CONQUISTA','037 - ITABUNA','038 - SANTO ANTONIO DE JESUS' ]
    
    try:#INICIO DA TRATIVA DE ERRO
        
        for x,y in zip(und,rev):#FOR DUPLO X CORRESPONDE ARRAY ORDEM DE FILTRO DE 1 A 10 CORRESPONDENTE A REVENDA NA COMBO DO XM
    

            #----------CONECTA NO XM LINK----------------------
        
            browser = webdriver.PhantomJS()#CHAMA O WEBDRIVER PHANTON
            #AQUI PASSA O LINK DO XM PARA 
            browser.get(url)
            print("CONECTANDO AO XMLINK")
            time.sleep(5)

            #VARIAVEL USERNAME RECEBE O CAMPO INPUT DO LOGIN DO XMLINL
            username = browser.find_element_by_id('ctl00_ContentPlaceHolder1_Login1_UserName')
            #VARIAVEL PASSWORD RECEBE O INPUT CORRESPONDENTE A SENHA
            password = browser.find_element_by_id('ctl00_ContentPlaceHolder1_Login1_Password')
            
            username.send_keys(login)
            time.sleep(5)
            password.send_keys(senha)

            browser.find_element_by_id('ctl00_ContentPlaceHolder1_Login1_LoginButton').click()


            time.sleep(5)
            print("LOGIN COM SUCESSO")
            #------------FIM DO TRCHO QUE ACESSA E LOGA NO XM--------------
            
            #-----INICIO DO LOOP DE FILTRO DE DATA E SELECIONAR REVENDA ESTE SERÁ REPETIDO CINCO VEZES---------- 
            
            dataini = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_Filtro1_txtDataInicial"]')

            datafim = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_Filtro1_txtDataFinal"]')
            time.sleep(5)
            dataini.clear()
            dataini.send_keys(hoje)#ESCREVA PASSADO SE QUISER PEGAR O DIA DE ONTEM OU HOJE SE QUISER PEGAR O DIA ATUAL
            time.sleep(5)
            datafim.clear()
            datafim.send_keys(hoje)#ESCREVA PASSADO SE QUISER PEGAR O DIA DE ONTEM OU HOJE SE QUISER PEGAR O DIA ATUAL
            time.sleep(15)
            select = Select(browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_Filtro1_cboEmpresa"]'))
   
            select.select_by_index(x)
            time.sleep(5)
            browser.find_element_by_name('ctl00$ContentPlaceHolder1$Filtro1$btnExibir').click()
            time.sleep(30)
            
            tbl = browser.find_element_by_xpath("""//*[@id="ctl00_ContentPlaceHolder1_updRep"]""") 
            html = tbl.get_attribute("innerHTML")
            soup = BeautifulSoup(html, "html.parser")#BEAUTFULSOUP ENTRA EM AÇÃO PEGANDO A TABELA E DISTRINCHANDO OS DADOS
            table = soup.find('table', attrs={'class':'dataTable'})
            table_rows = table.find_all('tr')
            
                
            res = []#ARRAY VAZIO
            for tr in table_rows:#FOR QUE PERCORRE TODAS LINHAS E CELULAS DA TABELA EXTRAINDO OS DADOS E JOGANDO NO ARRAY RES
                
                td = tr.find_all('td')
                row = [tr.text.strip() for tr in td if tr.text.strip()]
                if row:
                
                    res.append(row)

            #AQUI JOGAMOS TODOS OS DADDOS DO ARRAY RES PARA UM DATAFRAME A TRATAMOS E PEGAMOS A LINHA TOTAL QUE É O QUE INTERESSA
            df = pd.DataFrame(res, columns=["Cod. Produto", "Descrição","Qtd. Bonificada","Qtd. Vendida","Preço Médio","Positivação","Volume HL"])
            hl = df
            hl["revenda"] = y
            hl = hl.tail(1)
            hl1 = hl[["revenda","Qtd. Vendida"]]
           
            revendas = revendas.append(pd.DataFrame(hl1, columns=['revenda','Qtd. Vendida']),ignore_index=True,sort=True) 
            print("DADOS DE "+ y +" IMPORTADO COM SUCESSO!")
            
            #revendas.to_excel(r"c:\""+y+".xlsx")  
            #revendas.to_json (r'C:\VENDADIA\teste.json')
            
            
        else:
            revendas.to_csv(r"c:\VENDADIA/"+"VENDA.csv", index = False)
          
            #-----AQUI INICIA COPIAR O DATAFRAME PARA A NUVEM PARA EXIBICIÇÃO DOS VENDA DIA ONLINA ----------                   
            
        
            print("EXTRATOR TERMINOU O CICLO E E RETORNA APÓS 3 MINUTOS ")
         
    except:
        revendas.to_csv(r"c:\VENDADIA/"+"VENDA.csv", index = False)
       
        pass
     