from pandas.core.reshape.concat import concat
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
import math
import openpyxl
import numpy as np
import time
import sys
import os
import unittest
import csv
import re
import easygui
#os.environ['MOZ_HEADLESS'] = '1'
#choix du fichier Excel
chemin = easygui.fileopenbox(msg=None, title='Selectionner votre fichier Excel', default='*.xlsx', filetypes='', multiple=False)
#Connexion au site
print("connexion au site")
driver = webdriver.Firefox()           
driver.get("https://store-vivalink.extranet.myopenip.fr/Qualification")
time.sleep(1)
elem = driver.find_element_by_id("login")
elem.send_keys("E-0000019941")
# time.sleep(0.1)
elem = driver.find_element_by_id("password")
elem.send_keys("SupLinq%#78")
# time.sleep(0.1)
elem.send_keys(Keys.RETURN)
print("connexion reussi")
time.sleep(.5)
driver.get("https://store-vivalink.extranet.myopenip.fr/Qualification")
#Read Excel file as a DataFrame
print("démarage du process")
data = pd.read_excel(chemin)
df = pd.DataFrame(data, columns = ['numéro'])
lst = []
cols = ['type', 'derniere_action', 'operateur_attributaire', 'operateur_exploitant', 'operateur_telecom']
def testnumero():
    numero = str(num)
    num1 = re.sub(r"\s+", "", numero)
    num2 = len(numero)
    if num1.startswith('0'):
            return num1
    elif num2 == 9:
            newnum = '0' + str(num1)
            return newnum
    else:
            return "0100000000"
for num in df.numéro:
       #numero = line.replace(" ", "")          
        
        #Vérifier que la zonne de recherche de numéro est presente
        try:
            element = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "tx_numberSearch")))
        finally:
                #fin
            x = True
        # recherche de la zonne de text et envoie du numéro
        numerouno = testnumero() 
        """ print(numerouno)
        numerouno = str(testnumero) """
        elem = driver.find_element_by_id("tx_numberSearch")
        elem.send_keys(numerouno)
        elem.send_keys(Keys.RETURN)
        # on vient désactiver le Loading ou waiting qui s'affiche
        driver.execute_script("document.getElementById('waiting-indicator').style.display = 'none';")
        #verification de la presence du premier element
        val = 60 # in seconds
        driver.implicitly_wait(val)
        try:
            element = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[5]/div[1]/div/div[2]/div/div[1]/div/p")))
        finally:
                #fin
            x = True
        # print("récupération des datas")
        typelign = driver.find_element_by_xpath('/html/body/div[1]/div[5]/div[1]/div/div[2]/div/div[1]/div/p').text
        lastact = driver.find_element_by_xpath('/html/body/div[1]/div[5]/div[1]/div/div[2]/div/div[2]/div[1]/p').text
        operatrib = driver.find_element_by_xpath('/html/body/div[1]/div[5]/div[1]/div/div[2]/div/div[3]/div[1]/p').text
        opeexploit = driver.find_element_by_xpath('/html/body/div[1]/div[5]/div[1]/div/div[2]/div/div[3]/div[2]/p').text
        opecom = driver.find_element_by_xpath('/html/body/div[1]/div[5]/div[1]/div/div[2]/div/div[3]/div[3]/p').text
        # on récupé le string de chaque résultat pour récupérer cequi nous interesse
        #type de ligne
        typelig = str(typelign)
        a = re.search("Type de ligne\n", typelig)
        typelign1 = typelig[:a.start()] + typelig[a.end():]
        # derniere action
        last = str(lastact)
        b = re.search("Dernière action\n", last)
        lastact1 = last[:b.start()] + last[b.end():]
        # operateur attributaire
        operatatri = str(operatrib)
        c = re.search("Opérateur Attributaire\n", operatatri)
        operatatrib1 = operatatri[:c.start()] + operatatri[c.end():]
        # operateur exploitant
        opeexploi = str(opeexploit)
        d = re.search("Opérateur Exploitant\n", opeexploi)
        opeexploit1 = opeexploi[:d.start()] + opeexploi[d.end():]
        # operateur exploitant
        opeco = str(opecom)
        e = re.search("Opérateur Commercial\n", opeco)
        opecom1 = opeco[:e.start()] + opeco[e.end():]
        #print("ajout des datas au fichier")
         # on stock le resultat dans le tableau vide lst            
        resultat = lst.append([typelign1, lastact1, operatatrib1, opeexploit1, opecom1])
        #on reactualise la page car autrement ça ne marche pas:(
        driver.get("https://store-vivalink.extranet.myopenip.fr/Qualification")
#on definie df1 avec lst et les collone
df1 = pd.DataFrame(lst, columns=cols)
#on agrege df1 et data
final = concat([data,df1], axis=1)
#on balance en excel
final.to_excel(chemin+"out.xlsx", index = False, header=True)
driver.close()
print("opération terminé")