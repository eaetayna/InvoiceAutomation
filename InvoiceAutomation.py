"""
Invoice Automation Script
Original: 2022
Updated:  2025
Author: Tayná Alves

This script automates sending invoice reminder emails to delinquent customers via Outlook, attaching the invoice and payment slip, and sending them individually to each customer.

"""

#!/usr/bin/env python
# coding: utf-8


get_ipython().system('pip install pywin32')



import win32com.client as cliente
import pandas as pd
import datetime as dt
import os
import cv2
import time
import pyautogui


# Reading the Excel File


tabela = pd.read_excel('21 - Inadimplentes.xlsx')
tabela.info ()


# Checking Today's Date


hoje = dt.datetime.now()
hoje = hoje.strftime("%d/%m/%y")
display(hoje)


# Collecting Data Only from Delinquent Customers


tabela_devedoresontem = tabela.loc[tabela['Cobrar']=='cobrar']
display (tabela_devedoresontem)
tabela_devedoresontem = tabela_devedoresontem.loc[tabela_devedoresontem['Vencto']<hoje] #when it's the weekend, change to <3
display (tabela_devedoresontem)


# Insert email signature image


imagem = cv2.imread("imagens/Assinatura.jpg")
cv2.imshow("Assinatura",imagem)


# Sending automatic emails via Outlook


outlook = cliente.Dispatch('Outlook.Application')


# Change here to your company sender e-mail

emissor = outlook.session.Accounts['tayna.alves@companyx.com.br']


mensagem = outlook.CreateItem(0x0)
mensagem.To = 'tayna.alves@companyz.com.br'
mensagem.Subject = 'COMPANYX - UNIDENTIFIED PAYMENTS'
mensagem.Body = """

Dear Sir/Madam, Good morning.

We would like to inform you that our system shows an outstanding balance regarding the invoice(s) mentioned below.

Please verify this information and let us know as soon as possible.

If the payment(s) have already been made, please send us the proof(s) of payment so that we can remove the invoice(s).

"""

mensagem._oleobj_.Invoke(*(64209,0,8,0,emissor))
mensagem.Save()
mensagem.Send()

dados = tabela_devedoresontem[['Titulo','Parcela','CNPJ','Cliente','Emissão','Vencto','Valor','E-mail']].values.tolist()


# Sending the email to all recipients


for dado in dados:
    NF = dado [0]
    parcela = dado [1]
    CNPJ = dado [2]
    clientte = dado [3]
    emissao = dado [4]
    vcto = dado [5]
    vcto = vcto.strftime("%d/%m/%y")
    valor = dado [6]
    destinatario = dado [7]
    mensagem = outlook.CreateItem(0x0)
    mensagem.Display() #retirar apos os testes
    mensagem.To = destinatario
    mensagem.Subject = 'COMPANYX - UNIDENTIFIED PAYMENTS'
    mensagem.HTMLBody = f"""

    <p> Dear Sir/Madam, good morning. <p>

    <p> We inform you that our system shows an outstanding balance for invoice <b>{NF}</b> due on <b> {vcto}</b>, in the amount of
    <b> R${valor:,.2f}</b>. <p>
    <p> Please verify and let us know as soon as possible. <p>

    <p> If the payment(s) have already been made, please send us the proof(s) of payment so that we can remove the invoice(s).<p>
   
   
   
    <p> <b>Company X LTDA</b> <p>
    <p> <b> Accounts Receivable/Collections Team</b> <p>      
    <p> <b> (11) 4999-9000 | Ramal: 336 </b> <p>
    <p> <b> www.companyx.com.br </b> <p>
    <p> <b> CNPJ: 96.966.966/0001-96 </b> <p>
   

    """
   

    mensagem._oleobj_.Invoke(*(64209,0,8,0,emissor))
    mensagem.Save()
    mensagem.Send()

pyautogui.alert('Invoices sent successfully!')


valor_brasil = 'R{valor:_.2f}'
valor_brasil = valor_brasil.replace('.',',').replace('_','.')
print(valor_brasil)