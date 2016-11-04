import xlrd
import smtplib
from datetime import datetime

workbook = xlrd.open_workbook('sua planilha excel com os dados.xlsx') #Define para workbook a função de abertura, designando uma planilha a ele.

worksheet = workbook.sheet_by_index(0) #Define qual woorkbook sera utilizado, qual aba. 

hoje = datetime.today() #designa a data de hoje a variavel hoje


# O for a seguir atribui os valores que estão na planilha para as variaveis. 
for row_num in range(worksheet.nrows): 
    if row_num == 0: 
        continue
        row = worksheet.row_values(row_num)
        datemode = workbook.datemode
        data_recebimento = datetime(*xlrd.xldate_as_tuple(row[2],datemode)) #a função datetime converte os valores da planilha para data normal xx/xx/xx, possibliitando sua comparação em python
        data_devolucao = datetime(*xlrd.xldate_as_tuple(row[3], datemode))
        nome = (row[0])
        epi = (row[1])

    if data_devolucao <= hoje: #faz a comparação de datas. 
        print(data_devolucao)
        print("enviar email")
        print(nome)

    '''
    Parte do código para enviar o email atraves de servidor SMTP, no caso utilizo o do gmail, pode se utilizar outro. 
    
    Exemplos no site :  http://www.codepianist.com/other/lista-de-servidores-de-email-smtp-pop/
    
    
    ''' 
    
    smtp = smtplib.SMTP_SSL('smtp.gmail.com', 465)

    smtp.login('seu email', 'sua senha')

    de = 'gatoajato8000@gmail.com'
    para = "%s"%nome
    msg = """De: Segurança do trabalho

    Subject: Teste envio de email EPI

    Por favor trocar seu EPI : %s.""" %(epi)

    smtp.sendmail(de, para, msg)

    smtp.quit()











