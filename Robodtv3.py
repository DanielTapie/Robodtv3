import os
import pandas as pd
import time

from playwright.sync_api import sync_playwright

nome_do_arquivo = "Kremilin.xlsx"
url_do_forms = "https://pt.surveymonkey.com/r/NB2FWN5"
df = pd.read_excel(nome_do_arquivo)

# Convertendo a coluna "Status" para o tipo de dados 'object'
df['Status'] = df['Status'].astype('object')

for index,row in df.iterrows():
    print ("Linha: " + str(index) + " E o nome da fera é " + row["Nome"] + " E seu CPF é: " + str(row["CPF"]))
    if row["Status"] == "ok":
        print ("Essa linha já foi preenchida c" + row ["Nome"])
        continue
    
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        
        page.goto (url_do_forms)
        
        time.sleep(2)
        
        page.fill('input[id="162745865"]', row["Nome"])
        
        page.fill('input[id="162745886"]', str(row["CPF"]))
            
        page.fill('input[id="162745895"]', row["Justificativa"])
        
        time.sleep(0)
        
        page.click ('#patas > main > article > section > form > div.survey-submit-actions.center-text.clearfix > button')
        
        browser.close()
     # Atualiza a coluna "Status" para "ok" após o envio dos dados
        df.at[index, 'Status'] = 'ok'
 # Salva as alterações no arquivo Excel   
df.to_excel(nome_do_arquivo, index=False) 

# Abre o arquivo Excel após o processamento
os.system(f'start excel.exe {nome_do_arquivo}')   
        
