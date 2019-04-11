import docx
import re
from splinter import Browser
import time
from mandarmail import main
import os

resultado = []
nroExptes = []
doc = ''
fullText = []

def definirLista():

    hoy = time.strftime('%d.%m.%Y')
    tareas = os.listdir('PATH')
    for i in tareas:
        if hoy in i:
            doc = docx.Document('PATH' + i)

            
    for para in doc.paragraphs:
        fullText.append(para.text)

    for i in fullText:
        match = re.search(r'\d+/\d+', i)
        if match:                      
            temp = match.group().split('/')
            tempDict = { "nro":temp[0],"ano":temp[1]}
            nroExptes.append(tempDict)
            
    getEstado()            
            
def getEstado():
    
    browser = Browser('chrome')
    browser.visit('http://scw.pjn.gov.ar/scw/home.seam')
    candadito = browser.find_by_tag('li')[3]
    candadito.click()
    browser.fill('username', 'USERNAME')
    browser.fill('password', 'PASSWORD')
    browser.find_by_id("kc-login").click()
    browser.find_by_css('.fa-list-ul').click()
    browser.find_by_tag('li')[4].click()

    for i in nroExptes:
        try:
            
            browser.find_by_css('.fa-share').click()
            time.sleep(2)
            browser.find_by_css('.form-control')[1].fill(i['nro'])
            browser.find_by_value('Consultar').click()
            time.sleep(2)
            numero = browser.find_by_css('.column')[0].value
            juzgado = browser.find_by_css('.column')[1].value
            autos = browser.find_by_css('.column')[2].value
            estado = browser.find_by_css('.column')[3].value
            mensaje = autos+": "+estado
            print mensaje
            resultado.append(mensaje)
        except Exception as e:
            print "Error con expte " + i['nro']
            
    browser.quit()        
    enviarMail()

def enviarMail():        
    resSeparado = '<br><br>'.join(resultado)

    paraEnviar = resSeparado.encode('ascii', 'ignore')

    main(paraEnviar,paraEnviar)
    quit()

definirLista()
