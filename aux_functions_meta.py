import MetaTrader5 as mt5
import time 
import datetime
import pandas as pd
import xlwings as xw
from itertools import combinations

def iniciar():
    if not mt5.initialize(login=87851089, server="XPMT5-DEMO",password="senha"):
        print("initialize() failed, error code =",mt5.last_error())
        quit()






def convert_unix_timestamp(timestamp):
    formatted_time = datetime.datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')
    return formatted_time

def copy_columns(df,lista,string):
    for i in lista:
        df['{}'.format(i)] = df[string]
        
def lista_ibo():
    df = pd.read_excel("IBOVDia_29-08-23.xlsx")
    return list(df['Código'])+['BOVA11']


###
#43 strike
#31 bid
#34 ask
# 93 ticker
#83 basis
#18 vencimento
#30 tipo
def scrapp_opc():
    list_df_opc     = []
    for i in lista_ibo():
        aux = mt5.symbols_get("*{}*".format(i[0:4]))
        df = pd.DataFrame(aux)
        print(len(df),i)
        if i == "KLBN111234":
            xw.view(df)
        if len(df)!=0 and len(df[df[43] != 0])>3  :
            aux_1 = df[(df[43] != 0)&((df[31] != 0)|(df[34] != 0))&(df[18].isin([1705702535, 1705708799, 1708121735, 1708127999, 1710547199, 1713571199 ]))][[93,43,18,83,30,31,34,]]
            
            list_df_opc.append(aux_1)
    df_opc = pd.concat(list_df_opc)
    df_opc.columns = ['ticker','strike','vencimento','basis','tipo','bid','ask']
    df_opc.to_excel("opc_bolsa.xlsx")
    
#como estou trabalhando apenas com o ibov, esses ativos não são todos de negocio e sim os que estão na carteira hipoteteica
#formatted_time = datetime.datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d')p

def butterfly_oe(df):
    
    for ativo,grouped_df in df.groupby(['basis','vencimento','tipo']):
        teste = time.time()
        if len(grouped_df)>=3:
            aux = mt5.symbols_get("*{}*".format(ativo[0][0:4]))
            df_aux = pd.DataFrame(aux)
            grouped_df =  grouped_df.sort_values(by='strike')
            copy_columns(grouped_df,['k1', 'k2', 'k3'],"strike")
            copy_columns(df_aux,['ticker_x', 'ticker_y', 'ticker'],93)
            df = pd.DataFrame(list(combinations(grouped_df['strike'], 3)), columns=['k1', 'k2', 'k3'])
            
            df = df.merge(grouped_df[['ticker','k1']], on='k1', how='left')
            df = df.merge(grouped_df[['ticker','k2']], on='k2', how='left')
            df = df.merge(grouped_df[['ticker','k3']], on='k3', how='left')
            df['basis'] = ativo[0]
            df['vencimento'] = ativo[1]
            df['tipo'] = ativo[2]
            df['qtd_k1'] = (df['k3'] - df['k2']) / (df['k3'] - df['k1'])*200
            df['qtd_k2'] = 200
            df['qtd_k3'] = 200 - df['qtd_k1']
            df = df.merge(df_aux[['ticker_x',34]], on='ticker_x', how='left')
            df = df.merge(df_aux[['ticker_y',31]], on='ticker_y', how='left')
            df = df.merge(df_aux[['ticker',34]], on='ticker', how='left')            
            df = df[(df['34_x'] != 0)&(df['34_y'] != 0)&(df[31] != 0)]
            df = df[(((df['34_x']*df['qtd_k1'])+(df['34_y']*df['qtd_k3']))<df['qtd_k2']*df[31])]

            if len(df) != 0:
                print("ENTRADA")
                print(df)
                
            
iniciar()
start = time.time()
scrapp_opc()
print(time.time() - start, "tempo para scrapp opc")

start = time.time()
df = pd.read_excel("opc_bolsa.xlsx")

butterfly_oe(df)
print(time.time() - start,"tempo de analise total trincas")

#fitro de bid ask na propria planilha e filtro de tempo ai já crio o opc_ibov diferente
#outra pergunta, preciso criar o opc_ibov,
#posso fazer tudo em uma tacada em loops continuos, assim que vejo as opções já faço as combinações 