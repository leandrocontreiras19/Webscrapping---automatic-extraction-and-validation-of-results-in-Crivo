# -*- coding: utf-8 -*-
"""
Created on Thu Sep 30 14:45:00 2021

@author: F8082762
"""


import pandas as pd
import time
from datetime import datetime, date, timedelta
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook
import win32com.client as win32 


nav=webdriver.Chrome()

nav.get("http://ociwin0025/#/login") #website
time.sleep(10) #Rodar a próxima linha de código depois de 10 segundos
nav.find_element_by_xpath('/html/body/main/div[2]/div[2]/ng-view/div/div/div/div/form/input[1]').send_keys("*******") #login Encontra o elemento via HTML e coloca a matrícula do colaborador
nav.find_element_by_xpath('/html/body/main/div[2]/div[2]/ng-view/div/div/div/div/form/input[2]').send_keys("*******") #senha Encontra o elemento via HTML e coloca a senha do colaborador
acessar = nav.find_element_by_xpath('/html/body/main/div[2]/div[2]/ng-view/div/div/div/div/form/button')
time.sleep(5)
#Utilizar o teclado do computador e executar 
ActionChains(nav) \
    .key_down(Keys.CONTROL) \
    .click(acessar) \
    .key_up(Keys.CONTROL) \
    .perform()

time.sleep(5)
ut_crivo = nav.find_element_by_link_text('Utilização do Crivo') #Procura o elemento pelo nome
time.sleep(2)
ut_crivo.click()
time.sleep(3)

#Encontra o elemento de data, seleciona, deleta e seleciona coloca a data que queremos, podendo ser um range / 41 - 51
nav.find_element_by_xpath('/html/body/main/div[2]/div[2]/ng-view/div/div[2]/filtros/div/div/div[2]/div/fieldset/div[1]/div/input').click()
time.sleep(3)
ActionChains(nav).key_down(Keys.CONTROL).send_keys('A').key_up(Keys.CONTROL).perform()
time.sleep(2)
nav.find_element_by_xpath('/html/body/main/div[2]/div[2]/ng-view/div/div[2]/filtros/div/div/div[2]/div/fieldset/div[1]/div/input').send_keys(Keys.DELETE)
time.sleep(2)
sdate = date(2021,12,7)   # start date
edate = date(2021,12,7)
data_prog = pd.date_range(sdate,edate,freq='d')
data_prog = data_prog.strftime('%d-%m-%Y')
data_prog = data_prog.to_list()

'29/09/2021 00:00:00 - 29/09/2021 23:59:59'
#Criando um looping para pegar as informações nas datas propostas
for i in data_prog:
    time.sleep(2)
    nav.find_element_by_xpath('/html/body/main/div[2]/div[2]/ng-view/div/div[2]/filtros/div/div/div[2]/div/fieldset/div[1]/div/input').send_keys(i+ " 00:00:00 - "+i+" 23:59:59")
    time.sleep(2)
    ok = nav.find_element_by_xpath('/html/body/div[2]/div[1]/div/button[1]')
    time.sleep(2)
    ok.click
    ''' time.sleep(2)
    ActionChains(nav) \
    .key_down(Keys.CONTROL) \
    .click(ok) \
    .key_up(Keys.CONTROL) \
    .perform()'''
    time.sleep(15)
    nav.find_element_by_xpath('/html/body/main/div[2]/div[2]/ng-view/div/div[2]/filtros/div/div/div[2]/div/div[3]/div/button[1]').click()
    time.sleep(25)
    
    #Pegando o valor no Crivo e colocando em uma variável como numérico
    first_value = nav.find_element_by_xpath('/html/body/main/div[2]/div[2]/ng-view/div/div[2]/div[1]/div[1]/div/div/div/div[2]/div[2]/div/div[1]/div/div[6]/div').get_attribute('innerText')
    first_value = first_value.replace('.', '')
    first_value = int(first_value)
    print(first_value)
    
    #Pegando o valor no Crivo e colocando em uma variável como numérico
    second_value = nav.find_element_by_xpath('/html/body/main/div[2]/div[2]/ng-view/div/div[2]/div[1]/div[1]/div/div/div/div[2]/div[2]/div/div[2]/div/div[6]/div').get_attribute('innerText')
    second_value = second_value.replace('.', '')
    second_value = int(second_value)
    print(second_value)
    
    #Somando os valores 
    total_crivo = first_value + second_value
    print(total_crivo)
    
    #Colocando via pandas, dataframe, essas informações e tratando os valores
    base_diariofinal = pd.read_excel("*********",sheet_name = 'Estudo')
    base_diariofinal = base_diariofinal.dropna(axis=1, how="all", thresh=None, subset=None, inplace=False)
    base_diariofinal['data'] = base_diariofinal ['data'].apply(lambda x: x.strftime('%Y/%m/%d'))
    
    #mudar o i para o formato data para conseguir comparar
    i=i[6:]+"/"+i[3:5]+"/"+i[0:2]
    base_diariofinal.loc[base_diariofinal['data'] == i,'consulta_crivo']= total_crivo
    
    #Criando a coluna diferença percentual a partir das outras colunas
    
    base_diariofinal['dif_percentual'] = ((base_diariofinal['consulta_crivo']-base_diariofinal['consulta_puro'])/base_diariofinal['consulta_crivo'])*100
    base_diariofinal['data'] = pd.to_datetime(base_diariofinal['data'])
    
    #Utilizando o OpenPyXL para atualizar o arquivo Excel / 103 - 112
    path = "********" #caminho do arquivo
    book = load_workbook(path)
    del book['Estudo']
    writer = pd.ExcelWriter(path, engine = 'openpyxl')
    writer.book = book


    base_diariofinal.to_excel(writer, sheet_name = 'Estudo',index=False,header=True)

    writer.close()

    #Retorna a página anterior para poder executar o looping novamente
    time.sleep(3)
    nav.execute_script("window.history.go(-1)") #
    time.sleep(3)
    ut_crivo = nav.find_element_by_link_text('Utilização do Crivo')
    time.sleep(2)
    ut_crivo.click()
    time.sleep(3)
    nav.find_element_by_xpath('/html/body/main/div[2]/div[2]/ng-view/div/div[2]/filtros/div/div/div[2]/div/fieldset/div[1]/div/input').click()
    time.sleep(3)
    ActionChains(nav).key_down(Keys.CONTROL).send_keys('A').key_up(Keys.CONTROL).perform()
    time.sleep(2)
    nav.find_element_by_xpath('/html/body/main/div[2]/div[2]/ng-view/div/div[2]/filtros/div/div/div[2]/div/fieldset/div[1]/div/input').send_keys(Keys.DELETE)
    time.sleep(2)

#Fecha o website
nav.close() 

#Depois de terminar a atualização irá fazer a conta e verificar se teve outlier, caso sim, enviará um e-mail com a biblioteca WIN32


for k in data_prog:
    base_diariofinal= base_diariofinal.drop_duplicates(subset=['dif_percentual'],keep="first",inplace=False)
    k=k[6:]+"/"+k[3:5]+"/"+k[0:2]
    var = base_diariofinal.loc[base_diariofinal['data'] == k,'dif_percentual'].values
    print(var)
    if var>= 0.01:
        Template =  r""" 
                    <hl> Análise de Dados do Crivo para o BigData</hl>
                    <p> Foi encontrado nos dados do dia %s a diferença superior 0.01&#8274 ocorrendo erro.
                    <p> A Diferença Percentual do Crivo para o BigData Puro foi: %s&#8274 consultas 
                    <p>Att,<p>""" %(k,var[0])
        
        outlook = win32.Dispatch('outlook.application')
        mail= outlook.CreateItem(0)
        mail.To = "*************"
        mail.Subject = "Detectação de divergência Crivo com BigData (diária) - Envio Automático"
        mail.htmlBody = Template
        inspector = mail.GetInspector
        inspector.Display()
        doc=inspector.WordEditor
        selection=doc.Content
        
        mail.Send()
        
    else:
        continue

        
        

    