from openpyxl import load_workbook
import Planguia_funcoes
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment, Protection
from openpyxl.styles import Font, numbers
import requests
import json


class EstimativaCustos:
    def __init__(self, arquivo, aba_leitura, aba_destino, cambio):
        self.arquivo = arquivo
        self.aba_leitura = aba_leitura
        self.aba_destino = aba_destino
        self.cambio = cambio

    def listar_taxadiaria(self):
        w = Planguia_funcoes.openr(self.arquivo, self.aba_leitura)
        dict_taxas = {}
        lista_dados = []
        numero_linhas = w[1].max_row
        for linhas in range(4, numero_linhas + 1):
            equipamento = w[1].cell(row=linhas, column=1).value
            nome = w[1].cell(row=linhas, column=2).value
            tipo = w[1].cell(row=linhas, column=3).value
            dolar = float(w[1].cell(row=linhas, column=4).value)
            real = float(w[1].cell(row=linhas, column=5).value)
            reajuste = float(w[1].cell(row=linhas, column=6).value)
            valor_taxa = round(((real * reajuste) / self.cambio) + dolar, 2)
            lista_dados.append(nome)
            lista_dados.append(tipo)
            lista_dados.append(valor_taxa)
            dict_taxas[equipamento] = lista_dados
            dict_taxas.copy()
            lista_dados = []
        return dict_taxas

    def listar_estimativa(self):
        w1 = Planguia_funcoes.openr(self.arquivo, 'Previa')
        lista_geral = []
        lista_embarc = []
        numero_linha = w1[1].max_row
        for linha in range(2, numero_linha + 1):
            equip_previa = w1[1].cell(row=linha, column=3).value
            embarc_previa = w1[1].cell(row=linha, column=1).value
            dias_petro = round(w1[1].cell(row=linha, column=11).value, 3)
            dias_pblog = round(w1[1].cell(row=linha, column=14).value, 3)
            lista_embarc.append(equip_previa)
            lista_embarc.append(embarc_previa)
            lista_embarc.append(dias_petro)
            lista_embarc.append(dias_pblog)
            lista_geral.append(lista_embarc)
            lista_embarc = []
        return lista_geral

    def catalogar_dados(self):
        planilha_estimativa = self.listar_estimativa()
        dados_taxa = self.listar_taxadiaria()
        planilha = Planguia_funcoes.openr(self.arquivo, self.aba_destino)
        linha = 3
        for listas in range(0, len(planilha_estimativa)):
            dados = planilha_estimativa[listas]
            planilha[1].cell(row=linha, column=1).value = dados[0]
            equipamento1 = planilha[1].cell(row=linha, column=1).value
            planilha[1].cell(row=linha, column=2).value = dados[1].upper()
            lista_resultados = dados_taxa[equipamento1]
            planilha[1].cell(row=linha, column=3).value = lista_resultados[1]
            taxa_diaria = lista_resultados[2]
            indice_petro = dados[2]
            indice_pblog = dados[3]
            valor_petro = round(taxa_diaria * indice_petro, 2)
            planilha[1].cell(row=linha, column=4).value = Planguia_funcoes.converter_moeda(valor_petro)
            valor_pblog = round(taxa_diaria * indice_pblog, 2)
            planilha[1].cell(row=linha, column=5).value = Planguia_funcoes.converter_moeda(valor_pblog)
            somatorio = round(valor_petro + valor_pblog, 2)
            planilha[1].cell(row=linha, column=6).value = Planguia_funcoes.converter_moeda(somatorio)
            linha = linha + 1
        Planguia_funcoes.closer('Projeto_planilha_Guia_Medição.xlsx', planilha[0])

    def formatar(self):
        planilha = Planguia_funcoes.openr(self.arquivo, self.aba_destino)
        planilha[1].cell(row=1, column=1).value = 'Cambio: ' + str(self.cambio)
        cabecalho = ['Equipamento', 'Embarcação', 'Tipo', 'Petrobras', 'PBLOG', 'Total']


        for dados in range(1, len(cabecalho) + 1):
            planilha[1].cell(row=2, column=dados).value = cabecalho[dados - 1]
            cores_verdec = PatternFill(fill_type='solid', start_color='c6ffb3', end_color='c6ffb3')
            planilha[1].cell(row=2, column=dados).fill = cores_verdec
            planilha[1].cell(row=2, column=dados).alignment = Alignment(horizontal='center')
            planilha[1].cell(row=2, column=dados).font = Font(bold=True)

        planilha[1].cell(row=1, column=1).font = Font(bold=True)
        planilha[1].cell(row=1, column=1).alignment = Alignment(horizontal='center')

        # numero_linha = planilha[1].max_row
        # regiao1 = planilha[1]['A2':'C{}'.format(2)]
        # thin = Side(border_style="thin", color="000000")
        # regiao1.border = Border(top=thin, left=thin, right=thin, bottom=thin)




        Planguia_funcoes.closer('Projeto_planilha_Guia_Medição.xlsx', planilha[0])


def iniciar_4():
    valor_cambio = Planguia_funcoes.importar_cambio()
    custo = EstimativaCustos('Projeto_planilha_Guia_Medição.xlsx', 'Taxa Diária', 'Estimativa Custo', valor_cambio)
    custo.catalogar_dados()
    custo.formatar()
