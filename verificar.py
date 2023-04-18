from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
from datetime import datetime,timedelta 
import numpy as np
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import gspread
import glob
import os.path
from selenium.webdriver.common.action_chains import ActionChains
import warnings

warnings.filterwarnings("ignore")

filename = "service_account.json"

def dia_da_semana():
    
    today = datetime.now()
    today = today.isoweekday()
    return (today)

def data_sexta():
    data_sexta = datetime.now() - timedelta(3)
    ts = pd.Timestamp(data_sexta)
    data_sexta = data_sexta.strftime('%d/%m/%Y')

    return(data_sexta)

def data_sabado():
    data_sabado = datetime.now() - timedelta(2)
    ts = pd.Timestamp(data_sabado)
    data_sabado = data_sabado.strftime('%d/%m/%Y')

    return(data_sabado)

def data_ontem():
    data_ontem = datetime.datetime.now() - timedelta(1)
    ts = pd.Timestamp(data_ontem)
    data_ontem = data_ontem.strftime('%d/%m/%Y')

    return(data_ontem)

def data_antes_ontem():
    data_ontem = datetime.now() - timedelta(2)
    ts = pd.Timestamp(data_ontem)
    data_ontem = data_ontem.strftime('%d/%m/%Y')

    return(data_ontem)

def data_hoje():
    data_hoje = datetime.now()
    ts = pd.Timestamp(data_hoje)
    data_hoje = data_hoje.strftime('%d/%m/%Y')
    
    return(data_hoje)

def hora_atual():
    hora_atual = datetime.now()
    ts = pd.Timestamp(hora_atual)
    hora_atual = hora_atual.strftime('%H:%M:%S')
    
    return(hora_atual)

def acessar_innovaro():
    
    #options = webdriver.ChromeOptions()
    #options.add_argument("--headless")
    
    #link1 = "http://192.168.3.141/"
    link1 = 'http://cemag.innovaro.com.br/sistema'
    #link1 = 'http://devcemag.innovaro.com.br:81/sistema'
    nav = webdriver.Chrome()#chrome_options=options)
    nav.maximize_window()
    time.sleep(2)
    nav.get(link1)

    return(nav)

########### LOGIN ###########

def login(nav):
    #logando 
    
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="username"]'))).send_keys("Francisco Lucas")
    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys("lucas7")

    # WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="username"]'))).send_keys("Francisco Lucas")
    # WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys("lucas6")

    time.sleep(2)

    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys(Keys.ENTER)

    time.sleep(2)

########### MENU ###########

def menu_innovaro(nav):
    #abrindo menu

    try:
        nav.switch_to.default_content()
    except:
        pass

    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.CLASS_NAME, 'menuBar-button-label'))).click()
    #WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.ID, 'bt_1892603865'))).click()

    time.sleep(2)

def listar(nav, classe):
    lista_menu = nav.find_elements(By.CLASS_NAME, classe)
    
    elementos_menu = []

    for x in range (len(lista_menu)):
        a = lista_menu[x].text
        elementos_menu.append(a)

    test_lista = pd.DataFrame(elementos_menu)
    test_lista = test_lista.loc[test_lista[0] != ""].reset_index()

    return(lista_menu, test_lista)

def listar_botoes_f(nav):
    
    listar_botoes = nav.find_elements(By.CLASS_NAME, 'button')
    
    elementos_menu = []

    for x in range (len(listar_botoes)):
        a = listar_botoes[x].text
        elementos_menu.append(a)

    test_lista = pd.DataFrame(elementos_menu)
    test_lista = test_lista.loc[test_lista[0] != ""].reset_index()

    return(listar_botoes, test_lista)

def menu_apontamento(nav):

    try:
        nav.switch_to.default_content()
    except:
        pass

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Produção'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em producao
    time.sleep(1)

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Controle de fábrica (SFC)'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em SFC
    time.sleep(1)

    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    click_producao = test_list.loc[test_list[0] == 'Apontamentos da produção'].reset_index(drop=True)['index'][0]
    
    lista_menu[click_producao].click() ##clicando em SFC
    time.sleep(1)

def preechendo_apontamentos(nav, dataInicioInnovaro, dataFimInnovaro):
    
    mudanca_iframe(nav)

    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + 'a')
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.DELETE)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(dataInicioInnovaro)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.TAB)

    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + 'a')
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.DELETE)
    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(dataFimInnovaro)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.TAB)

    time.sleep(1)
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).click()
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + 'a')
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.DELETE)

    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + Keys.SHIFT + 'E')
    
    try:
        nav.switch_to.default_content()
    except:
        pass

    time.sleep(2)

    try:
        while WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.ID, 'progressMessageBox'))):
            print('carregando')
    except:
        pass

    time.sleep(2)
    nav.switch_to.default_content()

    listar_botoes, test_list = listar_botoes_f(nav)
    click_producao = test_list.loc[test_list[0] == 'Exportar'].reset_index(drop=True)['index'][0]
    
    listar_botoes[click_producao].click() ##clicando em producao
    listar_botoes[click_producao].click() ##clicando em producao
    listar_botoes[click_producao].click() ##clicando em producao
    
    time.sleep(1.5)
    
    WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.ID, 'answers_1'))).click()

    time.sleep(2)

    listar_botoes, test_list = listar_botoes_f(nav)
    click_producao = test_list.loc[test_list[0] == 'Executar'].reset_index(drop=True)['index'][0]
    
    listar_botoes[click_producao].click() ##clicando em producao
    listar_botoes[click_producao].click()
    listar_botoes[click_producao].click()
    time.sleep(1.5)
        
    mudanca_iframe(nav)

    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.ID, '_download_elt'))).click()

    time.sleep(1.5)

########### ACESSANDO PLANILHAS DE TRANSFERÊNCIA ###########

def mudanca_iframe(nav):

    iframe_list = nav.find_elements(By.CLASS_NAME, 'tab-frame')

    for iframe in range(len(iframe_list)):
        time.sleep(1)
        try:
            nav.switch_to.default_content()
            nav.switch_to.frame(iframe_list[iframe])
            print(iframe)
        except:
            pass

def ultimo_arquivo():

    folder_path = r"C:\Users\pcp2\Downloads"
    file_type = '\*csv' # se nao quiser filtrar por extenção deixe apenas *
    files = glob.glob(folder_path + file_type)
    max_file = max(files, key=os.path.getctime)

    return max_file

def etl():

    max_file = ultimo_arquivo()

    df = pd.read_csv(max_file, encoding ='ISO-8859-1' ,sep=';')

    df = df.replace('"','',regex=True)
    df = df.replace('=','',regex=True)

    df.columns = df.columns.str.replace('=','')
    df.columns = df.columns.str.replace('"','')

    df['Recurso'] = df['Recurso'].replace('-','_', regex=True)
    df['n_underscore'] = df['Recurso'].str.find('_')
    df['Recurso'] = df['Recurso'].replace('-','_', regex=True)
    
    for i in range(len(df)+5):
        try:
            df['Recurso'][i] = df['Recurso'][i][:df['n_underscore'][i]]
        except:
            pass
    
    df = df.drop(columns=['n_underscore'])

    return (df)

def planilha_serra(filename, dataInicio, dataFim):

    #SERRA#

    sheet = 'RQ PCP-001-000 (APONTAMENTO SERRA)'
    worksheet1 = 'USO DO PCP'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks_serra = sh.worksheet(worksheet1)

    headers = wks_serra.row_values(2)

    base = wks_serra.get()
    base = pd.DataFrame(base)
    base = base.set_axis(headers, axis=1, inplace=False)[2:]

    ########### Tratando planilhas ###########

    #ajustando datas
    for i in range(len(base)+2):
        try:
            if base['DATA'][i] == "":
                base['DATA'][i] = base['DATA'][i-1]
        except:
            pass

    i = None

    #filtrando peças que não foram apontadas
    #base_filtrada  = base[base['APONTAMENTO'].isnull()]

    #filtrando data de hoje
    base = base[base['DATA'] != '']
    base['DATA'] = pd.to_datetime(base['DATA'], format='%d/%m/%Y')

    base_filtrada = base.query('DATA >= @dataInicio and DATA < @dataFim')    

    #base_filtrada = base[base['DATA'] == data]

    base_filtrada['Máquina'] = 'Serra' 

    df_apontamento = base_filtrada.copy()

    base_filtrada = base_filtrada.loc[base_filtrada.APONTAMENTO != '']
    
    base_filtrada = base_filtrada[base_filtrada['APONTAMENTO'].str.contains("OK ROB", na=False)]

    base_filtrada = base_filtrada[['DATA','CÓDIGO','QNT', 'Máquina']]
    df_apontamento = df_apontamento[['DATA','CÓDIGO','QNT', 'Máquina']]

    base_filtrada = base_filtrada.rename(columns={'DATA':'Data','CÓDIGO':'Recurso','QNT':'Produzido'})
    df_apontamento = df_apontamento.rename(columns={'DATA':'Data','CÓDIGO':'Recurso','QNT':'Produzido'})

    base_filtrada = base_filtrada[['Data','Recurso', 'Máquina','Produzido']]
    df_apontamento = df_apontamento[['Data','Recurso', 'Máquina','Produzido']]

    return(base_filtrada, df_apontamento, wks_serra)

def planilha_usinagem(filename, dataInicio, dataFim):

    #USINAGEM#

    sheet = 'RQ PCP-001-000 (APONTAMENTO SERRA)'
    worksheet1 = 'USINAGEM 2022'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks_usinagem = sh.worksheet(worksheet1)

    headers = wks_usinagem.row_values(5)

    base = wks_usinagem.get()
    base = pd.DataFrame(base)
    base = base.iloc[:,0:11]
    base = base.set_axis(headers, axis=1, inplace=False)[5:]

    ########### Tratando planilhas ###########

    #ajustando datas
    for i in range(len(base)+5):
        try:
            if base['DATA'][i] == "":
                base['DATA'][i] = base['DATA'][i-1]
        except:
            pass

    i = None

    #filtrando peças que não foram apontadas
    #base_filtrada  = base[base['PCP'].isnull()]

    #filtrando data de hoje
    base = base[base['DATA'] != '']
    base['DATA'] = pd.to_datetime(base['DATA'], format='%d/%m/%Y')

    base_filtrada = base.query('DATA >= @dataInicio and DATA < @dataFim')    
    
    #base_filtrada = base[base['DATA'] == data]

    base_filtrada['Máquina'] = 'Usinagem' 
    
    df_apontamento = base_filtrada.copy()

    #inserindo 0 antes do código da peca
    base_filtrada['CÓDIGO'] = base_filtrada['CÓDIGO'].astype(str)

    for i in range(len(base)):
        try:
            if len(base_filtrada['CÓDIGO'][i]) == 5:
                base_filtrada['CÓDIGO'][i] = "0" + base_filtrada['CÓDIGO'][i] 
        except:
            pass

    i = None

    base_filtrada = base_filtrada[base_filtrada['PCP'].str.contains("OK ROB", na=False)]

    base_filtrada = base_filtrada[['DATA','CÓDIGO','QNT', 'Máquina']]
    df_apontamento = df_apontamento[['DATA','CÓDIGO','QNT', 'Máquina']]

    base_filtrada = base_filtrada.rename(columns={'DATA':'Data','CÓDIGO':'Recurso','QNT':'Produzido'})
    df_apontamento = df_apontamento.rename(columns={'DATA':'Data','CÓDIGO':'Recurso','QNT':'Produzido'})

    base_filtrada = base_filtrada[['Data','Recurso', 'Máquina','Produzido']]
    df_apontamento = df_apontamento[['Data','Recurso', 'Máquina','Produzido']]

    return(base_filtrada, df_apontamento, wks_usinagem)

def planilha_corte(filename, dataInicio, dataFim):

    #CORTE#

    sheet = 'Banco de dados OP'
    worksheet1 = 'Finalizadas'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks_corte = sh.worksheet(worksheet1)

    base = wks_corte.get()
    base = pd.DataFrame(base)
    #base = base.iloc

    headers = wks_corte.row_values(5)#[0:3]

    base = base.set_axis(headers, axis=1, inplace=False)[5:]

    base['Data finalização'] = base['Data finalização'].str[:10]

    ########### Tratando planilhas ###########

    #filtrando peças que não foram apontadas
    #base_filtrada  = base[base['PCP'].isnull()]

    #filtrando data de hoje
    base['Data finalização'] = pd.to_datetime(base['Data finalização'], format='%d/%m/%Y')
    base = base[base['Data finalização'] != '']
    
    base = base.rename(columns={"Data finalização":'Data'})

    base_filtrada = base.query('Data >= @dataInicio and Data < @dataFim')
    
    # base_filtrada = base[base['Data'] == data]

    base_filtrada['Máquina'] = 'Corte'

    df_apontamento = base_filtrada.copy()

    #filtrando linhas que tem código de chapa
    base_filtrada = base_filtrada[base_filtrada['Código Chapa'] != '']

    #extraindo código
    base_filtrada["Peça"] = base_filtrada["Peça"].str[:6]

    for i in range(len(base)):
        try:
            if len(base_filtrada['Peça'][i]) == 5:
                base_filtrada['Peça'][i] = "0" + base_filtrada['Peça'][i] 
        except:
            pass

    base_filtrada = base_filtrada[base_filtrada['Apont. peças'].str.contains("OK ROB", na=False)]

    base_filtrada = base_filtrada[['Data','Peça','Total Prod.', 'Máquina']]
    df_apontamento = df_apontamento[['Data','Peça','Total Prod.', 'Máquina']]

    base_filtrada = base_filtrada.rename(columns={'Peça':'Recurso','Total Prod.':'Produzido'})
    df_apontamento = df_apontamento.rename(columns={'Peça':'Recurso','Total Prod.':'Produzido'})

    base_filtrada = base_filtrada[['Data','Recurso', 'Máquina','Produzido']]
    df_apontamento = df_apontamento[['Data','Recurso', 'Máquina','Produzido']]

    return(base_filtrada, df_apontamento, wks_corte)

def planilha_estamparia(filename, dataInicio, dataFim):

    #ESTAMPARIA#

    sheet = 'RQ PCP-003-001 (APONTAMENTO ESTAMPARIA) e RQ PCP-009-000 (SEQUENCIAMENTO ESTAMPARIA)'
    worksheet1 = 'APONTAMENTO PCP (RQ PCP 003 001)'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks_estamparia = sh.worksheet(worksheet1)

    base = wks_estamparia.get()
    base = pd.DataFrame(base)
    base = base.iloc[:,0:15]

    headers = wks_estamparia.row_values(5)[0:15]

    base = base.set_axis(headers, axis=1, inplace=False)[5:]

    ########### Tratando planilhas ###########

    #filtrando data de hoje

    base = base[base['DATA'] != '']

    base['DATA'] = pd.to_datetime(base['DATA'], format='%d/%m/%Y')

    base_filtrada = base.query('DATA >= @dataInicio and DATA < @dataFim')
    
    # base_filtrada = base[base['DATA'] == data]

    df_apontamento = base_filtrada.copy()

    base_filtrada = base_filtrada.fillna('')

    #filtrando linhas sem observação
    base_filtrada = base_filtrada.loc[base_filtrada.CÓDIGO != '']

    base_filtrada['Máquina'] = 'Estamparia'

    df_apontamento = base_filtrada.copy()

    #MATRICULA ALEX: 4322
    for i in range(len(base)):
        try:
            if base_filtrada['MATRÍCULA'][i] == '': 
                base_filtrada['MATRÍCULA'][i] = '4322'
        except:
            pass

    base_filtrada['MATRÍCULA'] = base_filtrada['MATRÍCULA'].str[:4]

    base_filtrada = base_filtrada.loc[(base_filtrada['QTD REALIZADA'] != '')]
        
    base_filtrada = base_filtrada[base_filtrada['STATUS'].str.contains("OK ROB", na=False)]

    base_filtrada = base_filtrada[['DATA','CÓDIGO','QTD REALIZADA','Máquina']]
    df_apontamento = df_apontamento[['DATA','CÓDIGO','QTD REALIZADA','Máquina']]

    base_filtrada = base_filtrada.rename(columns={'DATA':'Data','CÓDIGO':'Recurso','QTD REALIZADA':'Produzido'})
    df_apontamento = df_apontamento.rename(columns={'DATA':'Data','CÓDIGO':'Recurso','QTD REALIZADA':'Produzido'})

    base_filtrada = base_filtrada[['Data','Recurso', 'Máquina','Produzido']]
    df_apontamento = df_apontamento[['Data','Recurso', 'Máquina','Produzido']]

    return(base_filtrada, df_apontamento, wks_estamparia)

def planilha_montagem(filename, dataInicio, dataFim):

    #MONTAGEM#

    sheet = 'RQ PCP-004-000 APONTAMENTO MONTAGEM M22'
    worksheet1 = 'APONTAMENTO'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks_montagem = sh.worksheet(worksheet1)

    base = wks_montagem.get()
    base = pd.DataFrame(base)
    base = base.iloc[:,0:9]

    headers = wks_montagem.row_values(5)[0:9]

    base = base.set_axis(headers, axis=1, inplace=False)[5:]

    ########### Tratando planilhas ###########
    
    base['CARIMBO'] = base['CARIMBO'].str[:10]

    base = base[base['CARIMBO'] != '#REF!']
    base = base[base['CARIMBO'] != '']

    base['FUNCIONÁRIO'] = base['FUNCIONÁRIO'].str[:4]

    base['CARIMBO'] = pd.to_datetime(base['CARIMBO'], format='%d/%m/%Y')

    base_filtrada = base.query('CARIMBO >= @dataInicio and CARIMBO < @dataFim')

    # base_filtrada = base[base['CARIMBO'] == data]

    base_filtrada['Máquina'] = 'Montagem'         

    df_apontamento = base_filtrada.copy()
    
    base_filtrada['CONJUNTO'] = base_filtrada['CONJUNTO'].replace('-','_', regex=True)
    base_filtrada['CONJUNTO'] = base_filtrada['CONJUNTO'].replace('-','_', regex=True)
    base_filtrada['CONJUNTO'] = base_filtrada['CONJUNTO'].replace(' ','_', regex=True)
    base_filtrada['n_underscore'] = base_filtrada['CONJUNTO'].str.find('_')
    
    for i in range(len(base)+5):
        try:
            base_filtrada['CONJUNTO'][i] = base_filtrada['CONJUNTO'][i][:base_filtrada['n_underscore'][i]]
        except:
            pass

    base_filtrada = base_filtrada[base_filtrada['STATUS'].str.contains("OK ROB", na=False)]

    base_filtrada = base_filtrada[['CARIMBO','CONJUNTO','QUANTIDADE', 'Máquina']]
    df_apontamento = df_apontamento[['CARIMBO','CONJUNTO','QUANTIDADE', 'Máquina']]

    base_filtrada = base_filtrada.rename(columns={'CARIMBO':'Data','CONJUNTO':'Recurso','QUANTIDADE':'Produzido'})
    df_apontamento = df_apontamento.rename(columns={'CARIMBO':'Data','CONJUNTO':'Recurso','QUANTIDADE':'Produzido'})

    base_filtrada = base_filtrada[['Data','Recurso', 'Máquina','Produzido']]
    df_apontamento = df_apontamento[['Data','Recurso', 'Máquina','Produzido']]

    return(base_filtrada, df_apontamento, wks_montagem)

def planilha_pintura(filename, dataInicio, dataFim):

    #PINTURA#

    sheet = 'BANCO DE DADOS ÚNICO - PINTURA'
    worksheet1 = 'RQ PCP-005-003 (APONTAMENTO PINTURA)'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks_pintura = sh.worksheet(worksheet1)

    base = wks_pintura.get()
    base = pd.DataFrame(base)
    #base = base.iloc

    headers = wks_pintura.row_values(5)#[0:3]

    base = base.iloc[0:,0:24]

    base = base.set_axis(headers, axis=1, inplace=False)[5:]
    
    base = base[base['Carimbo'] != '']

    base['Carimbo'] = pd.to_datetime(base['Carimbo'], format='%d/%m/%Y')

    ########### Tratando planilhas ###########
    
    base_filtrada = base.query('Carimbo >= @dataInicio and Carimbo < @dataFim')

    #base_filtrada = base[base['Carimbo'] == data]

    #base_filtrada = base.loc[base['Carimbo'] == data]
    base_filtrada['Máquina'] = 'Pintura'

    df_apontamento = base_filtrada.copy()

    base_filtrada.columns

    base_filtrada = base_filtrada[base_filtrada['STATUS'].str.contains("OK ROB", na=False)]

    base_filtrada = base_filtrada[['Carimbo','CÓDIGO','Qtd', 'Máquina']]
    df_apontamento = df_apontamento[['Carimbo','CÓDIGO','Qtd', 'Máquina']]

    base_filtrada = base_filtrada.rename(columns={'Carimbo':'Data','CÓDIGO':'Recurso','Qtd':'Produzido'})
    df_apontamento = df_apontamento.rename(columns={'Carimbo':'Data','CÓDIGO':'Recurso','Qtd':'Produzido'})

    base_filtrada = base_filtrada[['Data','Recurso', 'Máquina','Produzido']]
    df_apontamento = df_apontamento[['Data','Recurso', 'Máquina','Produzido']]

    return(base_filtrada, df_apontamento, wks_pintura)

#### inicio ####

#data = data_ontem()

dataInicioInnovaro = '01/04/2023'
dataFimInnovaro = '18/04/2023'

dataInicio = '2023-04-01'
dataFim = '2023-04-18'

nav = acessar_innovaro()

try:
    while not WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/main/section/main/section[1]/form/label[1]/input'))):
        print('carrregando login')
except:
    print('carregou')

login(nav)

menu_innovaro(nav)

menu_apontamento(nav)

plan_serra, apontamentos_serra, wks_serra = planilha_serra(filename, dataInicio, dataFim)

plan_usinagem, apontamentos_usinagem, wks_usinagem = planilha_usinagem(filename, dataInicio, dataFim)

plan_corte, apontamentos_corte, wks_corte = planilha_corte(filename, dataInicio, dataFim)

preechendo_apontamentos(nav, dataInicioInnovaro, dataFimInnovaro)

plan_estamparia, apontamentos_estamparia, wks_estamparia = planilha_estamparia(filename, dataInicio, dataFim)

plan_montagem, apontamentos_montagem, wks_montagem = planilha_montagem(filename, dataInicio, dataFim)

plan_pintura, apontamentos_pintura, wks_pintura = planilha_pintura(filename, dataInicio, dataFim)

dfs_apontamentos = [apontamentos_serra,apontamentos_usinagem,apontamentos_corte,apontamentos_estamparia,apontamentos_montagem, apontamentos_pintura]
dfs = [plan_serra, plan_usinagem, plan_corte, plan_estamparia, plan_montagem, plan_pintura] # list of dataframes

plan_geral = pd.concat(dfs)
apontamentos_geral = pd.concat(dfs_apontamentos)

df = etl()

df = df.replace('Montagem Carretas', 'Montagem')
df = df.replace('Serras', 'Serra')
df = df.replace('Corte Guilhotina', 'Corte')

df['Produzido'] = df['Produzido'].replace(np.nan,0)
df['Produzido'] = df['Produzido'].astype(int)
df['Produzido'] = df['Produzido'].astype(str)
df['Produzido'] = df['Produzido'].replace(0,'')
df['Produzido'] = df['Produzido'].astype(int)

df = df.sort_values(by='Produzido', ascending=True)

apontamentos_geral['Produzido'] = pd.to_numeric(apontamentos_geral['Produzido'], errors='coerce')
apontamentos_geral = apontamentos_geral.dropna(subset=['Produzido'])

apontamentos_geral['Produzido'] = apontamentos_geral['Produzido'].astype(int)
plan_geral['Produzido'] = plan_geral['Produzido'].astype(int)

apontamentos_geral = apontamentos_geral.sort_values(by='Produzido', ascending=True)
plan_geral = plan_geral.sort_values(by='Produzido', ascending=True)

plan_geral['Produzido'] = plan_geral['Produzido'].astype(str)
apontamentos_geral['Produzido'] = apontamentos_geral['Produzido'].astype(str)

df['count'] = df.groupby('Recurso').cumcount() + 1
df['id'] = df['Recurso'] + df['count'].astype(str)
df['id'] = df['id'].replace(" "," - ",regex=True)
#df = df[df['Produzido'] != 0]

plan_geral['count'] = plan_geral.groupby('Recurso').cumcount() + 1
plan_geral['id'] = plan_geral['Recurso'] + " - " + plan_geral['count'].astype(str)
plan_geral = plan_geral.reset_index()
#plan_geral = plan_geral.rename(columns={'Recurso':'Recurso'})

apontamentos_geral['count'] = apontamentos_geral.groupby('Recurso').cumcount() + 1
apontamentos_geral['id'] = apontamentos_geral['Recurso'] + " - " + apontamentos_geral['count'].astype(str)

ids_not_in_df2 = plan_geral[~plan_geral['id'].isin(df['id'])].reset_index(drop=True)

ids_not_in_df2.to_csv("ids_not_in_innovaro.csv")


#nav.close()