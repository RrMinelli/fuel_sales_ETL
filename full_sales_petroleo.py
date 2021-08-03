import xml.etree.ElementTree as et
import pandas as pd
import requests
import zipfile
import win32com.client as win32
import os, sys, os.path

from datetime import datetime
from requests.exceptions import ConnectionError

# Extração dos dados para XML 

def extract_xlsx_to_xmlE():
    #Responsavel por extrair e trasnformarar o dados em xlsx e converter em xml para manipulação dos dados

    try:
        os.stat('C:\\Users\\rmine\\fuel_sales_ETL\\raw')
    except:
        os.mkdir('C:\\Users\\rmine\\fuel_sales_ETL\\raw')
        print('LOG[INFO]: creat directory successfully.')


    raw_anp = requests.get('http://www.anp.gov.br/arquivos/dados-estatisticos/vendas-combustiveis/vendas-combustiveis-m3.xls', allow_redirects=True)
    open('raw/vendas-combustiveis-m3.xls', 'wb').write(raw_anp.content)
    print('LOG[INFO]: downloaded successfully')

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open('C:\\Users\\rmine\\fuel_sales_ETL\\raw\\vendas-combustiveis-m3.xls')    
    wb.SaveAs("C:\\Users\\rmine\\fuel_sales_ETL\\raw\\vendas-combustiveis-m3.xlsx", FileFormat=51)
    wb.Close()
    excel.Application.Quit()
    print('LOG[INFO]: Conversion xlsx successfully')

    with zipfile.ZipFile('C:\\Users\\rmine\\fuel_sales_ETL\\raw\\vendas-combustiveis-m3.xlsx', 'r') as zip_ref:
         zip_ref.extractall('raw/xml_extract')
    print('LOG[INFO]: Conversion xml successfully.')

xml_definition = 'C:\\Users\\rmine\\fuel_sales_ETL\\raw\\xml_extract\\xl\\pivotCache\\pivotCacheDefinition1.xml'
xml_records = 'C:\\Users\\rmine\\fuel_sales_ETL\\raw\\xml_extract\\xl\\pivotCache\\pivotCacheRecords1.xml'

# Extração de dados das tag XML para manipulação

def get_info_definitionP():
    ## Gera dicionario de tag Produto,UF,ANO

    xtree = et.parse(xml_definition)
    xroot = xtree.getroot()

    lista = []
    

    for element in xroot.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}s'):
        lista.append(element.attrib['v'])

    Index_produto = {}
    Index_UF= {}
    
    ## Lista de combustível
    for index, value in enumerate(lista[:8]):
        Index_produto[index] = value
    print('LOG[INFO]: Extracted fuel successfully.')
    ## Lista de Estados
    for index, value in enumerate(lista[13:]):
        Index_UF[index] = value
    print('LOG[INFO]: Extracted state successfully.')

    lista_ano = []

    for element in xroot.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}n'):
        lista_ano.append(element.attrib['v'])

    Idex_ano= {}
    
    ## Lista de Ano
    for index, value in enumerate(lista_ano):
        Idex_ano[index] = value
    print('LOG[INFO]: Extracted year successfully.')

    print('LOG[INFO]: Completed extraction of tag information')

    return Index_produto, Index_UF, Idex_ano

# Coleta de valores do XML, normalização dos dados e Transformação em Parquet

def get_info_generalP(Index_produto, Index_UF, Idex_ano):
    ## Responsavel por montar e normalizar as informaçoes fornecidas pelas tats e valores extraidos.

    try:
        os.stat('C:\\Users\\rmine\\fuel_sales_ETL\\structure')
    except:
        os.mkdir('C:\\Users\\rmine\\fuel_sales_ETL\\structure')
        print('LOG[INFO]: creat directory successfully.')

    xtree = et.parse(xml_records)
    xroot = xtree.getroot()

    list_info = []
    
    for value_tag in xroot.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}x'):
        list_info.append(value_tag.attrib['v'])
    
    data_info = {"product": [], "year": [], "uf": []}
    
    #normalização de Tags
    for i in range(0, len(list_info), 4):
        data_info['product'].append(Index_produto[int(list_info[i])])
        data_info['year'].append(Idex_ano[int(list_info[i + 1])])
        data_info['uf'].append(Index_UF[int(list_info[i + 3])])
    print('LOG[INFO]: Completed normalization tag information.')

    list_data = []

    for value_tag in xroot.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}n'):
        list_data.append(value_tag.attrib['v'])

    data_volu = {"1": [], "2": [], "3": [], "4": [], "5": [], "6": [], "7": [], "8": [], "9": [], "10": [], "11": [], "12": [], "total": []}           
    indicator = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]

    # Listar tags XML
    lista_tag = [elem.tag for elem in xroot.iter()]

    # Atribuir 0 para menes com valor nulo 
    listI = []
    for i in lista_tag:
        if '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}n' in i:
            listI.append('1')
        elif '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}m' in i:
            listI.append('0')
    
    for index_new, value_new in enumerate(listI):
        if value_new != '1':
            list_data.insert(index_new, 0)
    print('LOG[INFO]: Successfully complete transformation from null to 0 month.')

    # normalização de dados valor mes     
    for i in range(0, len(listI), 13):
        index_list = indicator.pop(0)
        indicator.insert(0, index_list)

        data_volu['1'].append(list_data[i + indicator[0]])
        data_volu['2'].append(list_data[i + indicator[1]])
        data_volu['3'].append(list_data[i + indicator[2]])
        data_volu['4'].append(list_data[i + indicator[3]])
        data_volu['5'].append(list_data[i + indicator[4]])
        data_volu['6'].append(list_data[i + indicator[5]])
        data_volu['7'].append(list_data[i + indicator[6]])
        data_volu['8'].append(list_data[i + indicator[7]])
        data_volu['9'].append(list_data[i + indicator[8]])
        data_volu['10'].append(list_data[i + indicator[9]])
        data_volu['11'].append(list_data[i + indicator[10]])
        data_volu['12'].append(list_data[i + indicator[11]])
        data_volu['total'].append(list_data[i + indicator[12]])
    print('LOG[INFO]: Successfully value month.')

    #criação da pandas dataframe
    df_pd_info = pd.DataFrame(data_info)
    df_pd_vol = pd.DataFrame(data_volu)
    df_file = pd.concat([df_pd_info, df_pd_vol], axis=1)
    print('LOG[INFO]: Successfully created dataframe pandas.')

    #Normalização de colunas pandas dataframe
    df_file.drop(columns=['total'], axis=1, inplace=True)
    df_fileN = df_file.melt(id_vars=['product', 'year', 'uf'],var_name='date', value_name='unit')
    
    df_fileN['year_month'] = df_fileN[['year','date']].apply(lambda x : '{}-{}'.format(x[0],x[1]), axis=1)
    df_fileN['volume'] = 'm3'
    df_fileN['created_at'] = datetime.now()
    print('LOG[INFO]: Successfully normalized dataframe pandas.')

    #Transdormação de conteudo para parquet
    df = df_fileN[df_fileN["product"] != 'ETANOL HIDRATADO (m3)']
    df_data = df[['year_month','uf','product','unit','volume','created_at']]

    schema ={"year_month": "datetime64", "uf": "string", "product": "string", "unit": "float64", "volume": "string", 'created_at': 'datetime64[ms]'}
    df_data = df_data.astype(schema)
    df_data.to_parquet('structure/derivado/result_derivado.parquet', engine='pyarrow', partition_cols = ['uf', 'product'], compression='snappy')
    print('LOG[INFO]: Dataframe transformation to parquet successfully.')

    return df_data

# Executar ETL derivado de Petróleo 
extract_xlsx_to_xmlE()
Index_produto, Index_UF, Idex_ano = get_info_definitionP
get_info_generalP(Index_produto, Index_UF, Idex_ano)