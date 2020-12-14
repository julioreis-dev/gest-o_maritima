from tkinter import messagebox
import pandas as pd
import modulo_planilha as pl
import modulo_calculo as mc
import modulo_email as em
import modulo_backupfiles as mb
import time
from tkinter import *

origin = r'C:\Users\Julio\Desktop\teste\Projeto_planilha_Guia_Medição.xlsx'
destination = r'C:\Users\Julio\Desktop\teste\planguia_draft.xlsx'


def prl(pfile):
    df = pd.read_excel(pfile, sheet_name='PRL', skiprows=[0])
    df = df[['Regional', 'Centro de Trabalho', 'Objeto de custo', 'CP', 'Descrição',
             'Unnamed: 5', 'Pólo (Tem cessão)', 'Objeto de custo (Não tem cessão)',
             'Unnamed: 8', 'PRL Petro', 'PRL PBLog']]
    df = df.to_dict('index')
    return df


def porte(pfile):
    df1 = pd.read_excel(pfile, sheet_name='Porte')
    dictdados = {}
    df1 = df1.to_dict('list')
    lisvalor = df1['Classe Desmembrada']
    for valor in lisvalor:
        for pt in ('%s' % (df1['PORTE'][i]) for i, v in enumerate(lisvalor) if v == valor):
            dictdados[valor] = pt
    return dictdados


def contract(pfile):
    df2 = pd.read_excel(pfile, sheet_name='Info Contrato')
    dictdados1 = {}
    df2 = df2.to_dict('list')
    listvalor2 = df2['ICJ']
    for valor2 in listvalor2:
        for pt1 in (i for i, v in enumerate(listvalor2) if v == valor2):
            dictdados1[valor2] = df2['Equipamento'][pt1], df2['Embarcação'][pt1], df2['Gerente'][pt1], \
                                 df2['Fiscal'][pt1], df2['Cessão'][pt1]
    return dictdados1


def aplicationtime(func):
    def resultime(*args, **kwargs):
        t0 = time.time()
        func(*args, **kwargs)
        t1 = time.time()
        temp = t1 - t0
        print('Tempo de execução : {} sec.'.format(round(temp, 2)))
    return resultime


@aplicationtime
def opt1(orig, desti):
    try:
        pr = prl(orig)
        port = porte(orig)
        contrac = contract(orig)
        plan = pl.PlanInitial(orig, desti, pr, port, contrac)
        plan.agregdatamedic()
        plan.agregacion()
        plan.ajustar_celulas()
        plan.validcontract()
        plan.validport()
        plan.alocatedataship()
        plan.formatation()
        opt3(orig, desti)
        msg = 'Prévia da planilha realizada com sucesso!!!'
        cav.itemconfig(result, text=msg)
    except Exception as err:
        handleerror(err)


@aplicationtime
def opt2(origin, destination):
    x = origin
    y = destination
    print('Em construção!!!!')


@aplicationtime
def opt3(orig, dest):
    # window.after(0000, func=clean)
    calculate = mc.CalcPlanGui(orig, dest)
    calculate.calcdata()
    retorno = 'Calculo de rateio aplicado.\nPlanilha guia finalizada com sucesso!!!'
    cav.itemconfig(result, text=retorno)
    window.after(5000, func=clean)


@aplicationtime
def opt4(orig, dest):
    corr = em.Email(orig, dest)
    revdata = corr.planinformation()
    corr.planformat(revdata)
    namefiley = corr.filesname(revdata)
    tocc = corr.preparedest()
    corr.planformat(revdata)
    issue = corr.emailcontent(revdata)
    back = mb.Backup(orig, dest, namefiley)
    back.verificar_pasta()
    nameadress = back.mover_arquivo()
    back.preparar_email(nameadress)
    statusend = corr.sendemail(revdata, tocc, issue, nameadress)
    print(statusend)


def optgeneral(choose):
    opt_general = {
        1: opt1,
        2: opt2,
        3: opt3,
        4: opt4,
    }
    operationchoosen = opt_general[choose]
    return operationchoosen


def clean():
    cav.itemconfig(result, text='')


def handleerror(erro):
    messagebox.showerror(title='Tratamento de erro', message=f'Erro:\nErro de aplicação {erro}')


BLUE = '#04d8fb'
GREEN = '#a8dda8'
t = time.localtime()
window = Tk()
window.title(f'Contratos marítimos {t[0]} - CMAR')
window.minsize(width=750, height=400)
window.config(padx=5, pady=5, bg=GREEN)
window.iconbitmap('ship-icon-png-29.ico')

canvas = Canvas(width=245, height=110, bg=GREEN, highlightthickness=0)
petro = PhotoImage(file='logo.png')
canvas.create_image(125, 55, image=petro)
canvas.grid(row=0, column=1, padx=5, pady=5)

cav = Canvas(width=300, height=100, bg=GREEN, highlightthickness=0)
result = cav.create_text(150, 20, text='', font=('Ariel', 10, 'bold'))
cav.grid(row=7, column=1, padx=0, pady=0)

# Label
my_label = Label(text=f'SISTEMA DE GERENCIAMENTO DE CONTRATOS',
                 fg='black', bg=GREEN, font=('Arial', 10, 'bold'))
my_label.grid(row=1, column=1, padx=15, pady=15)

version_label = Label(text='Versão',
                      fg='black', bg=GREEN, font=('Arial', 12, 'bold'))
version_label.grid(row=2, column=1, padx=5, pady=5)

med_label = Label(text='Escolha o mês:',
                  fg='black', bg=GREEN, font=('Arial', 12, 'bold'))
med_label.grid(row=2, column=2, padx=5, pady=5)

med_label = Label(text='Aplicações disponíveis:',
                  fg='black', bg=GREEN, font=('Arial', 12, 'bold'))
med_label.grid(row=2, column=0, padx=5, pady=5)

# botão
previa_buttom = Button(text='1.Emitir prévia', font=('Arial', 12, 'bold'), width=12, height=1)
previa_buttom['command'] = lambda: opt1(origin, destination)
previa_buttom.grid(row=3, column=0)

inoper_buttom = Button(text='2.Indisp', font=('Arial', 12, 'bold'), width=12, height=1)
inoper_buttom.grid(row=4, column=0, padx=5, pady=5)

calcular_buttom = Button(text='3.Versão final', font=('Arial', 12, 'bold'), width=12, height=1)
calcular_buttom['command'] = lambda: opt3(origin, destination)
calcular_buttom.grid(row=5, column=0)

email_buttom = Button(text='4.Enviar email', font=('Arial', 12, 'bold'), width=12, height=1)
email_buttom['command'] = lambda: opt4(origin, destination)
email_buttom.grid(row=6, column=0, padx=5, pady=5)

opt = IntVar()
opt.set(t[1])
month = [("Janeiro", 1), ("Fevereiro", 2), ("Março", 3), ("Abril", 4), ("Maio", 5), ('Junho', 6)]
month2 = [('Julho', 7), ('Agosto', 8), ('Setembro', 9), ('Outubro', 10), ('Novembro', 11), ('Dezembro', 12)]
incr = 0
for mes, val in month:
    Radiobutton(window,
                text=mes,
                padx=20,
                variable=opt,
                bg=GREEN,
                font=('Arial', 10, 'bold'),
                # command=ShowChoice,
                value=val).place(x=500, y=210 + incr)
    incr += 25
incr2 = 0
for mes2, val2 in month2:
    Radiobutton(window,
                text=mes2,
                padx=20,
                variable=opt,
                bg=GREEN,
                font=('Arial', 10, 'bold'),
                # command=ShowChoice,
                value=val2).place(x=610, y=210 + incr2)
    incr2 += 25
version = Spinbox(window, from_=0, to=10, width=2, font=('Arial', 14, 'bold'), bg=GREEN)
version.grid(row=3, column=1)
window.mainloop()

# flag = True
# origin = r'C:\Users\Julio\Desktop\Projeto_planilha_Guia_Medição.xlsx'
# destination = r'C:\Users\Julio\Desktop\planguia_draft.xlsx'
# while flag:
#     answer = input('################################################'
#                    '\nTipos de opções disponíveis nesta aplicação:'
#                    '\nDigite 1 --> Emitir prévia da planilha guia.'
#                    '\nDigite 2 --> Computar indices de inoperâncias. (Em construção!!!)'
#                    '\nDigite 3 --> Preparar versão final da planilha guia.'
#                    '\nDigite 4 --> (Opcional) - Enviar planilha guia para os gerentes e fiscais de contrato.'
#                    '\nDigite 0 --> Sair.'
#                    '\n################################################'
#                    '\nPrezado usuário, escolha uma opção?')
#     if answer.isdigit():
#         answer = int(answer)
#         if answer == 1:
#             opt1(origin, destination)
#             opt2(origin, destination)
#             sleep(7)
#         elif answer == 2:
#             print('Aplicação em construção!!!!')
#             # opt2(destination)
#             sleep(7)
#         elif answer == 3:
#             opt2(origin, destination)
#             sleep(3)
#         elif answer == 4:
#             lastdest = r''
#         elif answer == 0:
#             print('Finalizando................')
#             flag = False
#     else:
#         print('\nPrezado usuário, tente novamente digitando um numero válido!!!\n')
# except Exception as err:
#     handleerror(err)
# erro = f'Erro:\nErro de aplicação {err}'
# cav = Canvas(width=300, height=100, bg='#a8dda8', highlightthickness=0)
# result = cav.create_text(150, 20, text='', font=('Ariel', 10, 'bold'))
# cav.grid(row=7, column=1, padx=0, pady=0)
# cav.itemconfig(result, text=erro)


# if __name__ == '__main__':
#     main()
