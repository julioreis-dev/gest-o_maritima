from openpyxl import load_workbook
import Planguia_funcoes
import requests
import json


class EstimativaCustos:
    def __init__(self, arquivo, aba_leitura, aba_destino):
        self.arquivo = arquivo
        self.aba_leitura = aba_leitura
        self.aba_destino = aba_destino

    @staticmethod
    def importar_cambio():
        request = requests.get('https://economia.awesomeapi.com.br/json/all')
        cotacao = json.loads(request.text)
        valor_compra = (cotacao['USD']['high'])
        return float(valor_compra)

    def listar_taxadiaria(self):
        w = Planguia_funcoes.openr('Projeto_planilha_Guia_Medição.xlsx', 'Taxa Diária')
        cambio = self.importar_cambio()
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
            valor_taxa = round(((real * reajuste) / cambio) + dolar, 2)
            lista_dados.append(nome)
            lista_dados.append(tipo)
            lista_dados.append(valor_taxa)
            dict_taxas[equipamento] = lista_dados
            dict_taxas.copy()
            lista_dados = []
        return dict_taxas


    def listar_estimativa(self):
        

z = EstimativaCustos('Projeto_planilha_Guia_Medição.xlsx', 'Taxa Diária', 'Estimativa Custo')
t = z.importar_cambio()
j = z.listar_taxadiaria()
print(t)
print(j)
