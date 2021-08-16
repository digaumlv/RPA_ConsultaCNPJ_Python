#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import requests
import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def ConsultCNPJ():
    
    #Set Location Chromedriver or Any driver
    driver = webdriver.Chrome(executable_path=r"chromedriver.exe")
    driver.maximize_window()

    #Set Location xlsx Excel
    df = pd.read_excel('CNPJ.xlsx')

    for i in range(len(df)):
        
        #Try 3 times export#
        for j in range(2):
            
            #Set Varible empty
            DataAbertura = ""
            NomeEmpresarial = ""
            NomeFantaisa = ""
            AtividadeEconomicaPrincipal = ""
            AtividadeEconomicaSecundaria = ""
            NaturezaJuridica = ""
            Logradouro = ""
            Numero = ""
            Complemento = ""
            CEP = ""
            Bairro = ""
            Municipio = ""
            UF = ""
            Email = ""
            Telefone = ""
            EFR = ""
            SituacaoCadastral = ""
            DataSituacaoCadastral = ""
            MotivoSituacaoCadastral = ""
            SituacaoEspecial = ""
            DataSituacaoEspecial = ""
            
            try:
                # Navigate to the application home page
                driver.get("http://servicos.receita.fazenda.gov.br/Servicos/cnpjreva/cnpjreva_solicitacao.asp")

                #Click CaptchaSonoro
                clickCaptcha = driver.find_element_by_xpath('//*[@id="captchaSonoro"]')
                clickCaptcha.click()

                #getting the captcha image
                sfile = 'filename.png'
                with open(sfile, 'wb') as file:
                    file.write(driver.find_element_by_xpath('/html/body/div/div[1]/div/div/div/form/div[1]/div[2]/div/div[1]/div[1]/img').screenshot_as_png)
                    file.close()

                url = "https://2captcha.com/in.php"
                tokenApi = "<Token>" # token 2captcha here!!

                payload = {'key': "" + tokenApi + ""}
                print(payload)
                png = open('filename.png','rb')
                files = [
                  ('file', png )
                ]
                headers= {}

                #Get return api
                response = requests.request("POST", url, headers=headers, data = payload, files = files)

                tokenPost = str(response.text.encode('utf8')).split("|")[1]

                #close file
                while png.closed==False:
                    png.close()

                #Get result Broken captcha
                url = "https://2captcha.com/res.php?key="+tokenApi+"&action=get&id="+str(tokenPost)

                payload = {}
                headers= {}

                #Wait resolution of the captcha
                while(True):
                    response = requests.request("GET", url, headers=headers, data = payload)

                    if str(response.text.encode('utf8')).find("OK") >= 0:
                        break
                    else:
                        time.sleep(1)

                tokenGet = str(response.text.encode('utf8')).split("|")[1]

                #Set CNPJ in Input and Click
                cnpj = "'" + str(df["CNPJ"][i]) + "'"
                driver.execute_script('document.getElementById("cnpj").value =' + cnpj)

                driver.find_element_by_xpath('/html/body/div/div[1]/div/div/div/form/div[1]/div[2]/div/div[2]/input').send_keys(tokenGet)
                driver.find_element_by_xpath('//*[@id="frmConsulta"]/div[3]/div/button[1]').click()

                #Extract Data of Enterprise
                DataAbertura = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(3) > tbody > tr > td:nth-child(3) > font:nth-child(3) > b").innerText')
                NomeEmpresarial = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(5) > tbody > tr > td > font:nth-child(3)").innerText')
                NomeFantaisa = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(7) > tbody > tr > td:nth-child(1) > font:nth-child(3) > b").innerText')
                AtividadeEconomicaPrincipal = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(9) > tbody > tr > td > font:nth-child(3)").innerText')
                AtividadeEconomicaSecundaria = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(11) > tbody > tr > td > font:nth-child(3) > b").innerText')
                NaturezaJuridica = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(13) > tbody > tr > td > font:nth-child(3) > b").innerText')
                Logradouro = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(15) > tbody > tr > td:nth-child(1) > font:nth-child(3) > b").innerText')
                Numero = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(15) > tbody > tr > td:nth-child(3) > font:nth-child(3)").innerText')
                Complemento = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(15) > tbody > tr > td:nth-child(5) > font:nth-child(3) > b").innerText')
                CEP = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(17) > tbody > tr > td:nth-child(1) > font:nth-child(3) > b").innerText')
                Bairro = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(17) > tbody > tr > td:nth-child(3) > font:nth-child(3) > b").innerText')
                Municipio = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(17) > tbody > tr > td:nth-child(5) > font:nth-child(3) > b").innerText')
                UF = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(17) > tbody > tr > td:nth-child(7) > font:nth-child(3) > b").innerText')
                Email = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(19) > tbody > tr > td:nth-child(1) > font:nth-child(3) > b").innerText')
                Telefone = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(19) > tbody > tr > td:nth-child(3) > font:nth-child(3)").innerText')
                EFR = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(21) > tbody > tr > td > font:nth-child(3) > b").innerText')
                SituacaoCadastral = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(23) > tbody > tr > td:nth-child(1) > font:nth-child(3) > b").innerText')
                DataSituacaoCadastral = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(23) > tbody > tr > td:nth-child(3) > font:nth-child(3) > b").innerText')
                MotivoSituacaoCadastral = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(25) > tbody > tr > td > font:nth-child(3) > b").innerText')
                SituacaoEspecial = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(27) > tbody > tr > td:nth-child(1) > font:nth-child(3) > b").innerText')
                DataSituacaoEspecial = driver.execute_script('return document.querySelector("#principal > table:nth-child(1) > tbody > tr > td > table:nth-child(27) > tbody > tr > td:nth-child(3) > font:nth-child(3) > b").innerText')
                break
            except:
                pass
            finally:
                while png.closed==False:
                    png.close()
        #Error
        if j>2:
            continue
            
        #Append Dict Result
        retorno = []
        retorno.append({

            "CNPJ": str(df["CNPJ"][i]),
            "DataAbertura": DataAbertura,
            "NomeEmpresarial": NomeEmpresarial,
            "NomeFantaisa": NomeFantaisa,
            "AtividadeEconomicaPrincipal": AtividadeEconomicaPrincipal,
            "AtividadeEconomicaSecundaria": AtividadeEconomicaSecundaria,
            "NaturezaJuridica": NaturezaJuridica,
            "Logradouro": Logradouro,
            "Numero": Numero,
            "Complemento": Complemento,
            "CEP": CEP,
            "Bairro": Bairro,
            "Municipio": Municipio,
            "UF": UF,
            "Email": Email,
            "Telefone": Telefone,
            "EFR": EFR,
            "SituacaoCadastral": SituacaoCadastral,
            "DataSituacaoCadastral": DataSituacaoCadastral,
            "MotivoSituacaoCadastral": MotivoSituacaoCadastral,
            "SituacaoEspecial": SituacaoEspecial,
            "DataSituacaoEspecial": DataSituacaoEspecial
        })

        #Write Result in File
        columns = ["CNPJ","DataAbertura", "NomeEmpresarial", "NomeFantaisa", "AtividadeEconomicaPrincipal", "AtividadeEconomicaSecundaria", "NaturezaJuridica", "Logradouro", "Numero", "Complemento", "CEP", "Bairro", "Municipio", "UF", "Email", "Telefone", "EFR", "SituacaoCadastral", "DataSituacaoCadastral", "MotivoSituacaoCadastral", "SituacaoEspecial", "DataSituacaoEspecial"]
        dfretorno = pd.DataFrame(retorno, columns=columns)
        dfretorno.to_csv("retorno_"+str(df["CNPJ"][i].replace(".","").replace("-","").replace("/",""))+".csv",index=False)

    #Close browser
    driver.close()

    #Close,Delete Image Captcha if Exists
    while png.closed==False:
        png.close()
        
    if os.path.isfile("filename.png") == True:
        os.remove("filename.png")

if __name__ == '__main__':
    ConsultCNPJ()

