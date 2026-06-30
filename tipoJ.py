#teste para projeto de estágio

#inicialmente vamos testar a conexão com o SAP GUI Scripting e a execução de uma transação simples, como ME53N (Visualizar Requisição de Compra).

import win32com.client
import pandas as pd

#adicionar uma coluna nova no excel chamada Comentários para preencher com as informações do SAP (coluna entre Centro de Custo e Cost Center Description)
sapguiauto = win32com.client.GetObject("SAPGUI")
application = sapguiauto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

session.findById("wnd[0]/tbar[0]/okcd").text = "ME53N"
session.findById("wnd[0]").sendVKey (0) #enter para executar a transação ME53N
session.findById("wnd[0]").sendVKey (17) #shift + F5 para abrir a tela de variantes

#preciso abrir o excel com as PRs, filtrar as tipo J e fazer a leitura de cada uma delas para verificar se é investimento (variavel)

#
#
# criar um for (loop) para ler as PRs em ordem usar pd 

session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-BANFN").text = "9554310715" #colocar o número da PR para testar (variavel)
session.findById("wnd[1]").sendVKey (0) #enter para executar a variante e abrir a PR 


#proximo passo é abrir ordem de compra para verificar se é investimento ou não, e depois exportar os dados para o excel.


## abertura da ordem 
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/lblCOBL-AUFNR").caretPosition = 5
session.findById("wnd[0]").sendVKey (2) #Ctrl + setas pra baixo até chegar na ordem de compra, depois disso f2 pra abrir a ordem 


#verificar se é investimento 
#aperta ctrl seta pra direita 3 vezes para chegar em atribuições
#session.findById("wnd[0]/tbar[1]/btn[6]").press
session.findById("wnd[0]").sendVKey (6)

#elemeno PEP
try:
    elemento_pep = session.findById("wnd[0]/usr/tabsTABSTRIP_600/tabpBUT2/ssubAREA_FOR_601:SAPMKAUF:0601/subAREA1:SAPMKAUF:0315/ctxtCOAS-PSPEL").text
    ##print("Elemento PEP:", elemento_pep) #se tiver elemento pep vai dar o valor
    investimento = "Investimento"
except Exception as e:
   investimento = ""
    ##print("Elemento PEP não encontrado:", e) #se não tiver elemento pep, vai dar a mensagem de que não foi encontrado, ou seja, não é investimento.