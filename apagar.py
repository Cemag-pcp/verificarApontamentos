import pandas as pd
from datetime import datetime,timedelta 
import gspread
import warnings
import time

warnings.filterwarnings("ignore")

filename = "service_account.json"

def planilha_serra(filename):

    #SERRA#

    sheet = 'RQ PCP-001-000 (APONTAMENTO SERRA)'
    worksheet1 = 'USO DO PCP'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks_serra = sh.worksheet(worksheet1)

    return wks_serra

def planilha_usinagem(filename):

    #USINAGEM#

    sheet = 'RQ PCP-001-000 (APONTAMENTO SERRA)'
    worksheet1 = 'USINAGEM 2022'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks_usinagem = sh.worksheet(worksheet1)

    return wks_usinagem

def planilha_corte(filename):

    #CORTE#

    sheet = 'Banco de dados OP'
    worksheet1 = 'Finalizadas'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks_corte = sh.worksheet(worksheet1)

    return wks_corte

def planilha_estamparia(filename):

    #ESTAMPARIA#

    sheet = 'RQ PCP-003-001 (APONTAMENTO ESTAMPARIA) e RQ PCP-009-000 (SEQUENCIAMENTO ESTAMPARIA)'
    worksheet1 = 'APONTAMENTO PCP (RQ PCP 003 001)'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks_estamparia = sh.worksheet(worksheet1)

    return wks_estamparia

def planilha_montagem(filename):

    #MONTAGEM#

    sheet = 'RQ PCP-004-000 APONTAMENTO MONTAGEM M22'
    worksheet1 = 'APONTAMENTO'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks_montagem = sh.worksheet(worksheet1)

    return wks_montagem

def planilha_pintura(filename):

    #PINTURA#

    sheet = 'BANCO DE DADOS ÚNICO - PINTURA'
    worksheet1 = 'RQ PCP-005-003 (APONTAMENTO PINTURA)'

    sa = gspread.service_account(filename)
    sh = sa.open(sheet)

    wks_pintura = sh.worksheet(worksheet1)

    return wks_pintura

ids_not_in_df2 = pd.read_csv("ids_not_in_innovaro.csv")

wks_pintura = planilha_pintura(filename)
wks_montagem = planilha_montagem(filename)
wks_estamparia = planilha_estamparia(filename)
wks_corte = planilha_corte(filename)
wks_usinagem = planilha_usinagem(filename)
wks_serra = planilha_serra(filename)

for i in range(len(ids_not_in_df2)):
    
    try:

        if ids_not_in_df2['Máquina'][i] == 'Pintura':
            wks_pintura.update('O' + str(ids_not_in_df2['index'][i]+1), '')

        time.sleep(1)    
        if ids_not_in_df2['Máquina'][i] == 'Montagem':
            wks_montagem.update('I' + str(ids_not_in_df2['index'][i]+1), '')
        
        time.sleep(1)
        if ids_not_in_df2['Máquina'][i] == 'Estamparia':
            wks_estamparia.update('N' + str(ids_not_in_df2['index'][i]+1), '')
        
        time.sleep(1)
        if ids_not_in_df2['Máquina'][i] == 'Corte':
            wks_corte.update('L' + str(ids_not_in_df2['index'][i]+1), '')
        
        time.sleep(1)
        if ids_not_in_df2['Máquina'][i] == 'Usinagem':
            wks_usinagem.update('J' + str(ids_not_in_df2['index'][i]+1), '')
        
        time.sleep(1)
        if ids_not_in_df2['Máquina'][i] == 'Serra':
            wks_serra.update('S' + str(ids_not_in_df2['index'][i]+1), '')

    except:
        pass