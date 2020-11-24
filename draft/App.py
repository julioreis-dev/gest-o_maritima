import pandas as pd
import modulo_planilha as pl
import modulo_calculo as mc
import time
from tkinter import messagebox  # filedialog
from time import sleep


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


def aplicationtime(function):
    def resultime(*args, **kwargs):
        t0 = time.time()
        function(*args, **kwargs)
        t1 = time.time()
        temp = t1 - t0
        print('Tempo de execução : {} sec.'.format(round(temp, 2)))

    return resultime


def message():
    return messagebox.askyesno(title='Gerência de Contratos Marítimos - CMAR',
                               message='Prezado usuário gostaria de emitir a planilha guia?')


@aplicationtime
def opt1(origin, destination):
    pr = prl(origin)
    port = porte(origin)
    contrac = contract(origin)
    plan = pl.PlanInitial(origin, destination, pr, port, contrac)
    plan.agregdatamedic()
    plan.agregacion()
    plan.ajustar_celulas()
    plan.validcontract()
    plan.validport()
    plan.alocatedataship()
    plan.formatation()


def opt2(orig, dest):
    calculate = mc.CalcPlanGui(orig, dest)
    calculate.calcdata()
    print('Planilha guia finalizada com sucesso!!!')


def main():
    try:
        flag = True
        origin = r'C:\Users\Julio\Desktop\Projeto_planilha_Guia_Medição.xlsx'
        destination = r'C:\Users\Julio\Desktop\planguia_draft.xlsx'
        while flag:
            answer = input('################################################'
                           '\nTipos de opções disponíveis nesta aplicação:'
                           '\nDigite 1 --> Emitir prévia da planilha guia.'
                           '\nDigite 2 --> Computar indices de inoperâncias. (Em construção!!!)'
                           '\nDigite 3 --> Preparar versão final da planilha guia.'
                           '\nDigite 4 --> (Opcional) - Enviar planilha guia para os gerentes e fiscais de contrato.'
                           '\nDigite 0 --> Sair.'
                           '\n################################################'
                           '\nPrezado usuário, escolha uma opção?')
            if answer.isdigit():
                answer = int(answer)
                if answer == 1:
                    opt1(origin, destination)
                    opt2(origin, destination)
                    sleep(7)
                elif answer == 2:
                    print('Aplicação em construção!!!!')
                    # opt2(destination)
                    sleep(7)
                elif answer == 3:
                    opt2(origin, destination)
                    sleep(3)
                elif answer == 4:
                    lastdest = r''
                elif answer == 0:
                    print('Finalizando................')
                    flag = False
            else:
                print('\nPrezado usuário, tente novamente digitando um numero válido!!!\n')
    except Exception as err:
        print(f'Erro:\nErro de aplicação {err}')
        exit()


if __name__ == '__main__':
    main()
