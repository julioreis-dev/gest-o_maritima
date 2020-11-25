import pandas as pd
import modulo_planilha as pl
import modulo_calculo as mc
import modulo_email as em
import time
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


def aplicationtime(func):
    def resultime(*args, **kwargs):
        t0 = time.time()
        func(*args, **kwargs)
        t1 = time.time()
        temp = t1 - t0
        print('Tempo de execução : {} sec.'.format(round(temp, 2)))

    return resultime


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

@aplicationtime
def opt2(origin, destination):
    x = origin
    y = destination
    print('Em construção!!!!')

@aplicationtime
def opt3(orig, dest):
    calculate = mc.CalcPlanGui(orig, dest)
    calculate.calcdata()
    print('Planilha guia finalizada com sucesso!!!')

@aplicationtime
def opt4(orig, dest):
    corr = em.email(orig, dest)
    revdata = corr.planinformation()
    corr.planformat(revdata)
    y = corr.filesname(revdata)
    tocc = corr.preparedest()
    corr.planformat(revdata)
    issue = corr.emailcontent(revdata)
    statusend = corr.sendemail(revdata, tocc, issue)
    print(statusend)

def optgeneral(choose):
    opt = {
        1: opt1,
        2: opt2,
        3: opt3,
        4: opt4,
    }
    operationchoosen = opt[choose]
    return operationchoosen


def main():
    try:
        flag = True
        origin = r'C:\Users\ay4m\Desktop\planguia\planilha_Guia_Medição.xlsx'
        destination = r'C:\Users\ay4m\Desktop\planguia\memoria_calculo.xlsx'
        while flag:
            answer = input('################################################'
                           '\nTipos de opções disponíveis nesta aplicação:'
                           '\nDigite 1 --> Emitir prévia da planilha guia.'
                           '\nDigite 2 --> Computar indices de inoperâncias. (Em construção!!!)'
                           '\nDigite 3 --> Recalcular planilha.'
                           '\nDigite 4 --> (Opcional) - Enviar planilha guia para os gerentes e fiscais de contrato.'
                           '\nDigite 0 --> Sair.'
                           '\n################################################'
                           '\nPrezado usuário, escolha uma opção?')
            if answer.isdigit():
                answer = int(answer)
                if answer in range(1, 5):
                    operation = optgeneral(answer)
                    operation(origin, destination)
                    time.sleep(3)
                elif answer == 0:
                    print('Finalizando................')
                    time.sleep(1)
                    flag = False
                else:
                    print('\nPrezado usuário, tente uma opção válida!!!\n')
                    time.sleep(3)
            else:
                print('\nPrezado usuário, tente novamente digitando um numero válido!!!\n')
                time.sleep(3)
    except Exception as err:
        print(f'Erro:\nErro de aplicação {err}')
        exit()


if __name__ == '__main__':
    main()
