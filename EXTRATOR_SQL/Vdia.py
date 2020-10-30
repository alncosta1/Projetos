import time
from datetime import datetime
import pygsheets
import pandas as pd
from sqlalchemy import create_engine
import pyodbc
import xlwings as xw

from numbers import Number

#----------------------------------INCIO PEGA AS INFORMAÇÕES DO BI DA PLANILHA VENDA DIA E GRAVA NO SHEETS ----------------------------------------
print("INICIANDO , ABRINDO MACRO E COPIANDO DADOS DO BI")
wb = xw.Book(r'C:\Users\00254780\OneDrive - CERVEJARIA PETROPOLIS SA\Anexos\bkp\DIÁRIO\FELÍCIO\ROTINAS\8 - VENDA DIA FELÍCIO v2.xlsm')
sht=wb.sheets('BI')

bi = pd.DataFrame()
COLUNAS = ['cemp','Cod Nome Emp','Empresas','Dias Úteis','Dias Trabalhados','Meta Venda Hecto','Real Venda Hecto','Meta Hl Cerveja','Real Hl Cerveja','Meta Faturamento','Real Faturamento','cabarehl']




list = sht.range('a25:l29').options(ndim=2).value


#aqui gravamos os dados de vendas        



bi = bi.append(pd.DataFrame(list, columns=['cemp','Cod Nome Emp','Empresas','Dias Úteis','Dias Trabalhados','Meta Venda Hecto','Real Venda Hecto','Meta Hl Cerveja','Real Hl Cerveja','Meta Faturamento','Real Faturamento','cabarehl']),ignore_index=True)
#bi['Real Venda Hecto'] = bi['Real Venda Hecto'].astype(int)

bi['Dias Úteis'] = bi['Dias Úteis'].astype(int)
bi['Dias Trabalhados'] = bi['Dias Trabalhados'].astype(int)
bi['Dias Trabalhados'] = bi['Dias Trabalhados'].astype(int)
gc = pygsheets.authorize(service_file='C:\\raw\\cred.json')
    #ABRINDO A PLANILHA DO GOOGLE 
sh = gc.open('venda_dia')

    #SELECIONANDO A GUIA DA PLANILHA É UM ARRAY 
wks = sh[1]
wks.clear()
    #COMANDO QUE IRAR ESCREVER NA LINHA 1 COLUNA 1 DA PLANILHA 
wks.set_dataframe(bi,(1,1))


print('FINALIZADO E DADOS IMPORTADOS')
print("INIANDO O LOOP A CADA 2 MINUTOS")
#----------------------------------------FIM----------------------------------------------------------------------------------
#----------------------AQUI INICIA A CONSULTA AO BANCO DE DE DADOS JA DEFININDO O TIME QUE O LOOP VAI RODAR-------------------        
def xvdia():
    


    try:
 
        print("Começando a Coleta de Dados")
    
        sql_conn = pyodbc.connect('DRIVER=SQL Server;SERVER=10.65.6.142, 0502;UID=login;PWD=senha;APP=Microsoft Office 2003;WSID=FRSRVBDSQL31;DATABASE=DCGP;')
        query = """SELECT
	Subordinacao.Codsup AS 'Superv.',

	Pedido.codven,
	Itens_pedido.codcat,
	replace(SUM ( Itens_pedido.qtde/ Produtos.Volven ),'.',',') AS 'QTDE',
	Transacao.Tipo,
	replace(SUM ( Itens_pedido.valbru ),'.',',') AS 'valbru',
	Itens_pedido.Nompro,
	replace(Produtos.Litros,'.',','),
	Itens_pedido.Trade,
	Itens_pedido.Codemp,
	CONVERT(DATE, Itens_pedido.datped, 121)  AS "DATA",
    replace((CONVERT(FLOAT,Produtos.Litros,2)* (SUM ( Itens_pedido.qtde/ Produtos.Volven ))/ 	100 ),'.',',') AS HL,
    replace(SUM ( Itens_pedido.valbru )/SUM ( Itens_pedido.qtde/ Produtos.Volven ),'.',',') AS PMVENDA,
	(CASE WHEN Itens_pedido.codcat IN ('15790','15792','15793','15794','1158',	'1159',	'1168',	'1169',	'1170',	'1176',	'1177',	'1180',	'1186',	'1190',	'1191',	'1192',	'1198',	'1199',	'12272',	'1237',	'12626',	'12705',	'14754',	'14807',	'14808',	'15138',	'15140',	'15148',	'15149',	'15152',	'2379',	'2382',	'2383',	'6494',	'8500',	'8737',	'8739',	'10014',	'1043',	'11002',	'11005',	'1173',	'1174',	'1175',	'1183',	'1187',	'1188',	'1189',	'1193',	'1194',	'12276',	'14440',	'14443',	'14860',	'15144',	'15145',	'15154',	'2323',	'2375',	'2378',	'2922',	'2984',	'6510',	'6511',	'6512',	'6515',	'6516',	'6518',	'6521',	'6525',	'6527',	'6530',	'6531',	'6533',	'7121',	'8849',	'8852',	'10029',	'10031',	'10048',	'11004',	'11007',	'1103',	'1150',	'1151',	'1152',	'1153',	'1155',	'1156',	'1157',	'1160',	'1161',	'1162',	'1163',	'1165',	'1166',	'1167',	'1171',	'1184',	'1185',	'1195',	'12274',	'1236',	'1238',	'14753',	'14796',	'14797',	'14813',	'14814',	'14815',	'14853',	'15082',	'15084',	'15136',	'15137',	'15147',	'15155',	'15156',	'2352',	'2354',	'2380',	'2381',	'5254',	'5273',	'5274',	'6189',	'6514',	'6517',	'6519',	'6520',	'6522',	'6524',	'6526',	'6528',	'6529',	'6532',	'7122',	'8851',	'8854',	'1129',	'1130',	'1131',	'1132',	'1239',	'1240',	'1382',	'1426',	'1427',	'1428',	'1429',	'1430',	'1431',	'1432',	'1433',	'1434',	'1435',	'1436',	'1437',	'14483',	'14498',	'14840',	'14841',	'14842',	'14843',	'14844',	'14845',	'14846',	'14847',	'14848',	'14849',	'14850',	'14851',	'1608',	'2350',	'8517',	'8518',	'991705',	'991798',	'991799',	'1154',	'1164',	'15063',	'2846',	'2847',	'2992',	'5812',	'5813',	'5814',	'6513',	'6523',	'1013',	'1196',	'13098',	'14437',	'14438',	'14444',	'14636',	'14640',	'14794',	'14795',	'15139',	'15150',	'15151',	'8738',	'1197',	'2351',	'2353',	'2373',	'2374',	'2958',	'2959',	'2960',	'2961',	'3267',	'3268',	'5275',	'5276',	'5277',	'5278',	'9332',	'9333',	'9452',	'9453',	'14734',	'14735',	'14736',	'14738',	'15083',	'15100',	'15622',	'15623',	'15624',	'3174',	'3175',	'7382',	'7383',	'7384',	'2448',	'2449',	'2450',	'3176',	'3177',	'3178',	'2990',	'2991',	'5250',	'5260',	'5270',	'9495',	'2980',	'2981',	'2982',	'2983',	'5623',	'5624',	'5809',	'5810',	'5811',	'8740',	'10055',	'1020',	'12273',	'14585',	'5856',	'9311',	'9312',	'12278',	'8300',	'8301',	'11003',	'11006',	'2924',	'2925',	'2926',	'2927',	'2985',	'6755',	'6756',	'6757',	'6758',	'8462',	'8463',	'8850',	'8853',	'5252',	'5262',	'5272',	'7381',	'8741',	'5251',	'5261',	'5271',	'9496',	'8414',	'8416',	'10063',	'3173',	'3199',	'3229',	'3301',	'3302',	'8422',	'8423',	'8424',	'8425',	'8426',	'8427',	'8428',	'8429',	'12275',	'14441',	'14442',	'15134',	'15135',	'15141',	'15142',	'4544',	'4545',	'4546',	'7410',	'7412',	'12277',	'12330',	'2451',	'2452',	'2453',	'2994',	'2995',	'4645',	'4646',	'4647',	'4648',	'5390',	'8432',	'8433',	'8434',	'8435',	'8585',	'8586',	'8587',	'8588',	'8589',	'8591',	'8592',	'8593',	'8594',	'15449',	'2447',	'2456',	'2457',	'2458',	'2459',	'7554',	'8590',	'14154',	'14744',	'15759',	'2455',	'4775',	'12777',	'2454',	'2697',	'2698',	'3807',	'3300',	'3309',	'14436',	'14439',	'14445',	'14449',	'14637',	'14638',	'14639',	'14446',	'15143',	'15146',	'15164',	'15165',	'15056'
) THEN 'CERVEJA' ELSE 'DEMAIS' END) AS TIPO_BEBIDA,
(case when trade = 2 then 't'
else 'n'
end) as vendatrade

		
FROM
	Itens_pedido Itens_pedido,
Pedido Pedido,
Produtos Produtos,
	Subordinacao Subordinacao,
Subordinacao Subordinacao_1,
	Transacao Transacao 
WHERE
concat(Pedido.codemp,Pedido.codven) not in ('3583584','35819','3583775','3583987','35811','35812','35814','35816','35845','358907','3582022','3582112','35850','35813','3583524 ','3582618','358120')
and
	Itens_pedido.Codemp = Pedido.codemp 
	AND Itens_pedido.nummap = Pedido.nummap 
	AND Itens_pedido.numped = Pedido.numped 
	AND Itens_pedido.tipped = Pedido.tipped 
	AND Produtos.Codemp = Itens_pedido.Codemp 
	AND Produtos.Codemp = Pedido.codemp 
	AND Itens_pedido.pecodi = Produtos.Codpro 
	AND Subordinacao.CodEmp = Itens_pedido.Codemp 
	AND Subordinacao.CodEmp = Pedido.codemp 
	AND Subordinacao.CodEmp = Produtos.Codemp 
	AND Pedido.codven = Subordinacao.Codsub 
	AND Subordinacao_1.CodEmp = Itens_pedido.Codemp 
	AND Subordinacao_1.CodEmp = Pedido.codemp 
	AND Subordinacao_1.CodEmp = Produtos.Codemp 
	AND Subordinacao_1.CodEmp = Subordinacao.CodEmp 
	AND Subordinacao.Codsup = Subordinacao_1.Codsub 
	AND Transacao.Codemp = Itens_pedido.Codemp 
	AND Transacao.Codemp = Pedido.codemp 
	AND Transacao.Codemp = Produtos.Codemp 
	AND Transacao.Codemp = Subordinacao.CodEmp 
	AND Transacao.Codemp = Subordinacao_1.CodEmp 
	AND Itens_pedido.tipoco = Transacao.Codigo 
	AND ((
			CONVERT(DATE, Itens_pedido.datped, 121)  = CONVERT(DATE,GETDATE() ,121) 
			) 
		AND (
		Transacao.Tipo IN ( 'v', 'b' )) 
		AND (
		Itens_pedido.Codemp IN ( '353', '354','358','366','405' ))) 
        and (Itens_pedido.Trade <> 2 OR Itens_pedido.Trade IS NULL)
        
	
GROUP BY

	Subordinacao.Codsup,
	
	Pedido.codven,
	Itens_pedido.codcat,
	Transacao.Tipo,
	Itens_pedido.Nompro,
	Produtos.Litros,
	Itens_pedido.Trade,
	Itens_pedido.Codemp,
	Itens_pedido.qtddev,
	CONVERT(DATE, Itens_pedido.datped, 121)
HAVING
	(
	Itens_pedido.qtddev= 0)
	"""
#----------------------------------FIM DA CONSULTA AO BANCO DE DADOS------------------------------------------------------------        

#----------------------------------A CONSULTA É PASSADO PARA UM DATAFRAME , ONDE É VERIFICADO SE ELE ESTÁ VAZIO PRIMEIRO-------
#-----------------------SE VAZIO DEIXAMOS NAS NUVENS O ULTIMO DATASET VÁLIDO--------------------------------------------------
        df = pd.read_sql(query, sql_conn)
        if df.empty:
            print("DataFrame Vazio")
        else:
#-----------------------SE O DATASET ESTIVER COM DADOS INICIA AQUI A TRATAIVA E A IMPORTAÇÃO PARA AS NUVENS------------------            
            
            print('Consulta Realizada, inicando importação no sheets')
#aqui termina a consulta 


            dateTimeObj = datetime.now()

    #wb = xw.Book(r'C:\Users\00254780\OneDrive - CERVEJARIA PETROPOLIS SA\Anexos\bkp\DIÁRIO\FELÍCIO\ROTINAS\8 - VENDA DIA FELÍCIO v2.xlsm')
    #sht=wb.sheets('Resultado')

    
    

#aqui gravamos os dados de vendas        
            base = df
            base['Data_atual'] = dateTimeObj.strftime("%d/%m/%Y %H:%M:%S")
            gc = pygsheets.authorize(service_file='C:\\raw\\cred.json')
    #ABRINDO A PLANILHA DO GOOGLE 
            sh = gc.open('venda_dia')

    #SELECIONANDO A GUIA DA PLANILHA É UM ARRAY 
            wks = sh[0]
            wks.clear()
    #COMANDO QUE IRAR ESCREVER NA LINHA 1 COLUNA 1 DA PLANILHA 
            wks.set_dataframe(df,(1,1))
            print('dados importados ')
    
    except:
        dateTimeObj = datetime.now()

    #wb = xw.Book(r'C:\Users\00254780\OneDrive - CERVEJARIA PETROPOLIS SA\Anexos\bkp\DIÁRIO\FELÍCIO\ROTINAS\8 - VENDA DIA FELÍCIO v2.xlsm')
    #sht=wb.sheets('Resultado')

        print('Verifique se sua Conexão está ok ou proxy da maquina está ativo,importando ultimos dados validos ')
        
            
while True:
    print("passou 15 minutos iniciando a extração de dados ")
    xvdia()
    time.sleep(100) #A cada hora ele executa o print        
    
    continue            
#-------------------------------------FIM DATASET IMPORTADO PARA AS NUVENS------------------------------------------------  