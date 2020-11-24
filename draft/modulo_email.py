import win32com.client
from openpyxl import load_workbook
import time


def planinformation():
    t = time.localtime()
    resposta = []
    month = input('\nQual é o numero do mês da planilha guia de medição? (ex:1-Janeiro, 2-Fevereiro, 3-Março, ...)?')
    review = input('Qual é a revisão da planilha guia de medição? (ex:1, 2, 3,...)?')
    if review.isdigit():
        review = int(review)
    else:
        print('Erro')

    if month.isdigit():
        month = int(month)
        if month == 1:
            period = '26/12/' + str(t[0] - 1) + ' a 25/' + str(month) + '/' + str(t[0])
        else:
            period = '26/' + str(month - 1) + '/' + str(t[0]) + ' a 25/' + str(month) + '/' + str(t[0])
        meses_ano = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio',
                 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        mes = meses_ano[month - 1]
        frase1 = mes + ' de ' + str(t[0])
        resposta.append(frase1)
        frase2 = '(Protocolo de envio - Planilha guia encaminhada no dia ' + str(t[2]) + '/' + str(t[1]) + '/' + \
                 str(t[0]) + ' as ' + str(t[3]) + ':' + str(t[4]) + ':' + str(t[5]) + ')'
        frase3 = period
        resposta.append(frase3)
        resposta.append('0'+str(review))
        resposta.append(frase2)
    return resposta

def filesname(namesdata):
    # wb = load_workbook(r'C:\Users\Julio\Desktop\arquivo_editado.xlsx')
    filename = namesdata[0]+ '- Planilha Guia_R'+namesdata[2]
    # wb.save(r'C:\Users\Julio\Desktop'+filename+'.xlsx')
    return filename


def preparedest():
    wb = load_workbook(r'C:\Users\Julio\Desktop\arquivo_editado.xlsx')
    ws = wb['Chaves']
    destiny = []
    copy = []
    cont = ws.max_row
    for w in range(2, cont+1):
        status = ws.cell(row=w, column=6).value
        if status == 'Sim':
            destiny.append(ws.cell(row=w, column=1).value)
        elif status == 'Copy':
            copy.append(ws.cell(row=w, column=1).value)
        else:
            pass
    finaldestiny = '; '.join(destiny)
    finalcopy = '; '.join(copy)
    return finaldestiny, finalcopy


def planformat(listdata):
    wb = load_workbook(r'C:\Users\Julio\Desktop\arquivo_editado.xlsx')
    ws = wb['Medição']
    fields = [2, 5, 8]
    for i, n in enumerate(fields):
        ws.cell(row=1, column=n).value = listdata[i]
    wb.save(r'C:\Users\Julio\Desktop\arquivo_editado.xlsx')


def sendemail(rev, dest):
    try:
        anexo = r'C:\Users\Julio\Desktop\arquivo_editado.xlsx'
        o = win32com.client.Dispatch("Outlook.Application")
        msg = o.CreateItem(0)
        msg.To = dest[0]
        msg.CC = dest[1]
        msg.BCC = ''
        msg.Subject = emailcontent(rev)
        msg.Body = emailcontent(rev)+ '\nHistórico de ajustes realizados - Revisão ' + str(rev[2]) + ':\n' + \
                   listar_revisao() + '\nAtenciosamente,\nEquipe de Gerenciamento Marítimo' \
                                           '\nLOEP/LOFF/GCI/CMAR\n'
        msg.Attachments.Add(anexo)
        index = anexo.rfind('/')
        extensao = anexo[index + 1:]
        lista_resposta = [extensao, '\nEmail enviado com sucesso!']
        msg.Send()
        return lista_resposta
    except:
        return 'Erro ao enviar email!!!'


def saudar():
    t = time.localtime()
    z = t[3]
    if z < 12:
        a = 'Prezados Gerentes e Fiscais de contrato, Bom dia!'
        return a
    elif z >= 18:
        b = 'Prezados Gerentes e Fiscais de contrato, Boa noite!'
        return b
    else:
        c = 'Prezados Gerentes e Fiscais de contrato, Boa tarde!'
        return c


def emailcontent(rev):
    corpo_email = saudar() + '\n\nSegue anexo a planilha guia de medição (Revisão ' + str(rev[2]) + ').\nQualquer ' \
                        'informação a acrescentar em alguma embarcação da frota, ' \
                        'que não constar na planilha, favor avisar para ajuste.' \
                        '\n\nATENÇÃO: PARA EMBARCAÇÕES QUE NÃO FORAM LIBERADAS ' \
                        'PARA MEDIÇÃO VER NA COLUNA (N - "Embarcação Liberada") DA ABA "MEDIÇÃO".' \
                        ' TODAS AS ALTERAÇÕES REALIZADAS ' \
                        'NA PLANILHA GUIA SÃO REGISTRADAS NA ABA "REVISÃO".' \
                        '\n\nAs embarcações "Não Liberadas" estão sendo analisadas ' \
                        'pela equipe de coordenação e controle das informações ' \
                        'referente a frota sob a gestão do CMAR, e em breve essas ' \
                        'embarcações poderão sofrer alteração quanto ao seu status. ' \
                        'Caso isso ocorra, uma nova revisão será emitida ' \
                        'com a finalidade de ajustar a planilha guia.'

    subject = 'CMAR - Planilha Guia de Medição - ' + rev[0] + ' - REVISÃO ' + rev[2]
    body = saudar()+ '\nHistórico de ajustes realizados - Revisão ' + rev[2] + ':\n' + \
                   str(listar_revisao())+ '\nAtenciosamente,\nEquipe de Gerenciamento Marítimo' \
                                           '\nLOEP/LOFF/GCI/CMAR\n'

    return corpo_email, subject, body

def listar_revisao():
    pass



x = planinformation()
print(x)
filesname(x)
planformat(x)
z = preparedest()
sendemail(x,z)

