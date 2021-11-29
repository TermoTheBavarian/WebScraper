# encoding=utf8

import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from distutils.core import setup
from email.mime.multipart import MIMEMultipart
import os
import sys
import win32com.client as client
import pandas as pd

from unidecode import unidecode
from bs4 import BeautifulSoup
import requests


def french_translator(lista):
    mensaje = "".join(lista)
    letterXchange = {'à': 'a', 'â': 'a', 'ä': 'a', 'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
                     'î': 'i', 'ï': 'i', 'ô': 'o', 'ö': 'o', 'ù': 'u', 'û': 'u', 'ü': 'u', 'ç': 'c', "’": "´"}
    text = mensaje  # Replace it with the string in your code.
    for item in list(text):
        if item in letterXchange:
            text = text.replace(item, letterXchange.get(str(item)))
        else:
            pass

    return text


def get_BCEAO():
    html_txt = requests.get("https://www.bceao.int/fr/appels-offres/appels-offres-marches-publics-achats").text
    soup = BeautifulSoup(html_txt, "lxml")
    offers = soup.find_all("span", class_="descFile")

    num = 1
    user = os.getlogin()
    file_path= "C:/Users/"+ user +"/Desktop/BCEAO.txt"
    with open(file_path, "w", encoding='utf-8') as f:
        f.write("OFERTAS BCEAO \n")
        lista = []
        for offer in offers:
            texto = offer.find("span", class_="ttr").text

            if "climatisation" in texto:
                f.write(f"{num}.-{texto},\n")
                lista.append(f"{num}.-{texto},\n")
                num = num + 1
    return lista, file_path


print("ARCHIVO GUARDADO")


def correo(text,file_path):
     try:
         outlook = client.Dispatch("Outlook.Application")
         message = outlook.CreateItem(0)
         message.To = "german@ecosen.org"
         message.Subject = "OFERTAS BCEAO"
         message.Body = text
         message.Attachments.Add(file_path)
         message.Send()

         print("Mail enviado")

     except:
        print("Eror: Mail no enviado")




def main():
    lista,file_path = get_BCEAO()
    text = french_translator(lista)
    print(text)
    correo(text,file_path)



if __name__ == "__main__":
    main()
