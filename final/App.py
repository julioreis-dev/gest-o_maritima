from tkinter import *
from tkinter import messagebox
import pandas as pd
import modulo_planilha as pl
import modulo_calculo as mc
import modulo_email as em
import modulo_backupfiles as mb
import time


origin = r'C:\Users\(chave)\Desktop\planguia\planilha_Guia_Medição.xlsx'
destination = r'C:\Users\ay4m\Desktop\planguia\planguia_draft.xlsx'


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
        pt = list(('%s' % (df1['PORTE'][i]) for i, v in enumerate(lisvalor) if v == valor))
        dictdados[valor] = pt[0]
        # for pt in ('%s' % (df1['PORTE'][i]) for i, v in enumerate(lisvalor) if v == valor):
        #     dictdados[valor] = pt
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
        clean()
        t0 = time.time()
        func(*args, **kwargs)
        t1 = time.time()
        temp = t1 - t0
        total_time = 'Tempo de execução : {} sec.'.format(round(temp, 2))
        canvas1.itemconfig(result_time, text=total_time)
        window.after(5000, func=clean)
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
        msg = 'Prévia da planilha guia realizada com sucesso!'
        canvas0.itemconfig(result, text=msg)
    except Exception as err:
        handleerror(err)


@aplicationtime
def opt2():
    # def opt2(origin, destination):
    # x = origin
    # y = destination
    mesgopt2 = 'Em construção!!!!'
    canvas0.itemconfig(result, text=mesgopt2)
    window.after(5000, func=clean)


@aplicationtime
def opt3(orig, dest):
    calculate = mc.CalcPlanGui(orig, dest)
    calculate.calcdata()
    calculate.sheetconfiguration()
    calculate.finalversion()
    calculate.sendproduct()
    retorno = 'Calculo da frota realizada com sucesso.\nPlanilha guia finalizada, pronta para o envio!!!'
    canvas0.itemconfig(result, text=retorno)
    window.after(5000, func=clean)


@aplicationtime
def opt4(orig, dest):
    version_number = optversion()
    month_number = optmonth()
    mail = em.Email(orig, dest, month_number, version_number)
    revdata = mail.planinformation()
    mail.planformat(revdata)
    namefiley = mail.filesname(revdata)
    tocc = mail.preparedest()
    mail.planformat(revdata)
    issue = mail.emailcontent(revdata)
    back = mb.Backup(orig, dest, namefiley)
    back.verificar_pasta()
    nameadress = back.mover_arquivo()
    back.preparar_email(nameadress)
    statusend = mail.sendemail(revdata, tocc, issue, nameadress)
    canvas0.itemconfig(result, text=statusend)
    window.after(5000, func=clean)


def clean():
    canvas0.itemconfig(result, text='')
    canvas1.itemconfig(result, text='')


def optmonth():
    return opt_radio.get()


def optversion():
    return version.get()


def handleerror(erro):
    messagebox.showerror(title='Mensagem de erro', message=f'Erro:\n{erro}')


BLACK = '#000000'
CINZA = '#B4B7BF'
BLUE = '#04d8fb'
GREEN = '#a8dda8'
t = time.localtime()
window = Tk()
window.title(f'SISTEMA DE GERENCIAMENTO DE MEDIÇÃO {t[0]} - LOEP/LOFF/GCI/CMAR')
window.minsize(width=730, height=350)
window.config(padx=15, pady=15, bg=GREEN)
window.iconbitmap('ship-icon-png-29.ico')

canvas = Canvas(width=245, height=110, bg=GREEN, highlightthickness=0)
petro = PhotoImage(file='logo.png')
canvas.create_image(125, 55, image=petro)
canvas.grid(row=0, column=1, padx=5, pady=5)

canvas0 = Canvas(width=300, height=100, bg=GREEN, highlightthickness=0)
result = canvas0.create_text(150, 20, text='', font=('Ariel', 10, 'bold'))
canvas0.grid(row=7, column=1, padx=0, pady=0)

canvas1 = Canvas(width=300, height=30, bg=GREEN, highlightthickness=0)
result_time = canvas1.create_text(150, 20, text='', font=('Ariel', 10, 'bold'))
canvas1.grid(row=8, column=1, padx=0, pady=0)

# Construção das Label da interface gráfica
my_label = Label(text=f'CONTRATOS MARÍTIMOS - CMAR',
                 fg='black', bg=GREEN, font=('Arial', 12, 'bold'))
my_label.grid(row=1, column=1, padx=5, pady=5)

version_label = Label(text='Revisão',
                      fg='black', bg=GREEN, font=('Arial', 12, 'bold'))
version_label.grid(row=2, column=1)

med_label = Label(text='Mês de medição:',
                  fg='black', bg=GREEN, font=('Arial', 12, 'bold'))
med_label.grid(row=2, column=2, padx=5, pady=5)

med_label = Label(text='Aplicações disponíveis:',
                  fg='black', bg=GREEN, font=('Arial', 12, 'bold'))
med_label.grid(row=2, column=0, padx=5, pady=5)

# Construção dos botões da interface gráfica
previa_buttom = Button(text='1. Emitir prévia', font=('Arial', 12, 'bold'),
                       width=12, height=1, anchor=NW, activebackground=CINZA, bd=5)
previa_buttom['command'] = lambda: opt1(origin, destination)
previa_buttom.grid(row=3, column=0)

inoper_buttom = Button(text='2. Inoperância', font=('Arial', 12, 'bold'),
                       width=12, height=1, anchor=NW, activebackground=CINZA, bd=5)
inoper_buttom['command'] = lambda: opt2()
inoper_buttom.grid(row=4, column=0, padx=5, pady=5)

calcular_buttom = Button(text='3. Calcular', font=('Arial', 12, 'bold'),
                         width=12, height=1, anchor=NW, activebackground=CINZA, bd=5)
calcular_buttom['command'] = lambda: opt3(origin, destination)
calcular_buttom.grid(row=5, column=0)

email_buttom = Button(text='4. Enviar email', font=('Arial', 12, 'bold'),
                      width=12, height=1, anchor=NW, activebackground=CINZA, bd=5)
email_buttom['command'] = lambda: opt4(origin, destination)
email_buttom.grid(row=6, column=0, padx=5, pady=5)

opt_radio = IntVar()
opt_radio.set(t[1])
month = [("Janeiro", 1), ("Fevereiro", 2), ("Março", 3), ("Abril", 4), ("Maio", 5), ('Junho', 6)]
month2 = [('Julho', 7), ('Agosto', 8), ('Setembro', 9), ('Outubro', 10), ('Novembro', 11), ('Dezembro', 12)]

# Construção dos radio buttom da interface gráfica
incr = 0
for mes, val in month:
    Radiobutton(window,
                text=mes,
                padx=20,
                variable=opt_radio,
                bg=GREEN,
                font=('Arial', 10, 'bold'),
                command=optmonth,
                value=val).place(x=450, y=190 + incr)
    incr += 25

incr2 = 0
for mes2, val2 in month2:
    Radiobutton(window,
                text=mes2,
                padx=20,
                variable=opt_radio,
                bg=GREEN,
                font=('Arial', 10, 'bold'),
                command=optmonth,
                value=val2).place(x=570, y=190 + incr2)
    incr2 += 25

# Construção do spinbox da interface gráfica
version = Spinbox(window, from_=0, to=9, width=2, font=('Arial', 14, 'bold'), bg=GREEN)
version.grid(row=3, column=1)
window.mainloop()

# if __name__ == '__main__':
#     main()
