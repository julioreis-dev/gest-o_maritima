import pandas as pd
from openpyxl import load_workbook
import xlsxwriter
import xlrd
from openpyxl.styles import PatternFill, Border, Side, Protection
import time
from openpyxl.styles import Font, numbers
from openpyxl.styles import colors

'''Funções de apoio que são necessárias para a elaboração da planilha guia.'''


def openr(arquivo, aba):
    lista = []
    wb = load_workbook(filename=arquivo)
    aba_dados = wb[aba]
    lista.append(wb)
    lista.append(aba_dados)
    return lista


def closer(nome_arquivo, wb):
    wb.save(filename=nome_arquivo)
    wb.close()


def contador_Linhas(endereco, aba_planilha):
    planilha_atual = pd.read_excel(endereco, sheet_name=aba_planilha)
    valor = planilha_atual.shape[0]
    return valor + 1


def contador_Colunas(endereco, aba_planilha):
    planilha_atual = pd.read_excel(endereco, sheet_name=aba_planilha)
    valor = planilha_atual.shape[1]
    return valor + 1


def atribuir_equipamento(consulta_ICJ, dict_ICJ_equip):
    equipamento_extraido = dict_ICJ_equip[consulta_ICJ]
    return equipamento_extraido


def agregar_dados():
    arquivo_excel = r'C:\Users\ay4m\Desktop\Python\projetos\Projeto_planilha_Guia_Medição_2020.xlsx'
    aba = pd.read_excel(arquivo_excel, sheet_name='Base Dados')
    resumo = pd.DataFrame(aba.groupby(['Embarcação', 'Regional1', 'Equipamento', 'Regional', 'Tipo'])['Dias'].sum())
    destino = pd.ExcelWriter('Projeto_planilha_Guia_Medição_2020_1.xlsx')
    resumo.to_excel(destino, 'Base Dados', index=True)
    destino.save()


def relacionar_porte():
    wb = load_workbook(filename='Projeto_planilha_Guia_Medição_2020.xlsx')
    aba_porte = wb['Porte']
    dict_porte = {}
    numero_linha = aba_porte.max_row
    for n in range(2, numero_linha+1):
        tipo = aba_porte.cell(row=n, column=1).value
        dict_porte[tipo] = aba_porte.cell(row=n, column=2).value
        dict_porte.copy()
    return dict_porte


def configurar_celula(aba):
    coluna_zeros = [8, 9, 11, 14]
    coluna_percentual = [10, 13]
    numero_linha = aba.max_row
    for linha in range(2, numero_linha+1):
        for n in coluna_zeros:
            aba.cell(row=linha, column=n).value = 0
            dias = aba.cell(row=linha, column=n)
            dias.number_format = '0.000'

    for linha1 in range(2, numero_linha+1):
        for t in coluna_percentual:
            percentual = aba.cell(row=linha1, column=t)
            percentual.number_format = '0.000%'


def mostrar_desempenho(numero, processo):
    lista_realizacoes = ['Importação e consolidação realizada com sucesso!!!', 'Centros de custos alocados com sucesso!!!', 'Prévia da planilha guia realizada com sucesso!!!']
    tempo = 0.5
    for i in range(1, numero+1):
        time.sleep(tempo)
        print('Processo {}/{} realizado com sucesso!!!'.format(i, numero))
        tempo = tempo + 0.5
    time.sleep(1)
    print(lista_realizacoes[processo])