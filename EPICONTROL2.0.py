import xlrd
import smtplib
import datetime
from datetime import datetime

workbook = xlrd.open_workbook('dados.xlsx')

worksheet = workbook.sheet_by_index(0)

hoje = datetime.today()


for row_num in range(worksheet.nrows):
    if row_num == 0:
        continue
    row = worksheet.row_values(row_num)
    datemode = workbook.datemode
    data_recebimento = datetime(*xlrd.xldate_as_tuple(row[2],datemode))
    data_devolucao = datetime(*xlrd.xldate_as_tuple(row[3], datemode))
    nome = (row[0])
    epi = (row[1])

if data_devolucao <= hoje:
    print(data_devolucao)
    print("enviar email")
    print(nome)

    smtp = smtplib.SMTP_SSL('smtp.gmail.com', 465)

    smtp.login('seu email', 'sua senha')

    de = 'gatoajato8000@gmail.com'
    para = "%s"%nome
    msg = """De: Segurança do trabalho

    Subject: Teste envio de email EPI

    Por favor trocar seu EPI : %s.""" %(epi)

    smtp.sendmail(de, para, msg)

    smtp.quit()











