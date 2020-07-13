import Planguia_funcoes
import pandas as pd
import xlsxwriter
import time
from openpyxl import load_workbook

'''Esta classe molda a planilha inicial de dados referentes a movimentação da frota maritima.'''


class PlanilhaInicial:
    def __init__(self, arquivo, aba, aba1):
        self.arquivo = arquivo
        self.aba = aba
        self.aba1 = aba1

    def documentar_equipamento(self):
        """ Prepara um dicionário com ICJ e número de equipamento."""
        dados_arquivo = Planguia_funcoes.openr(self.arquivo, self.aba)
        contador_linha = dados_arquivo[1].max_row
        dict_equipamento = {}
        for n in range(2, contador_linha + 1):
            numero_icj = dados_arquivo[1].cell(row=n, column=1).value
            dict_equipamento[numero_icj] = dados_arquivo[1].cell(row=n, column=2).value
            dict_equipamento.copy()
        return dict_equipamento

    def preparar_tupla(self):
        """Prepara uma Tupla com os dados relevantes de medição da aba com as informações recebidos."""
        dados_arquivo = Planguia_funcoes.openr(self.arquivo, self.aba1)
        contador = dados_arquivo[1].max_row
        lista_dados = []
        for linha in range(2, contador + 1):
            embarcacao = dados_arquivo[1].cell(row=linha, column=1).value
            icj = dados_arquivo[1].cell(row=linha, column=2).value
            tipo = dados_arquivo[1].cell(row=linha, column=3).value
            regional = dados_arquivo[1].cell(row=linha, column=4).value
            regional_1 = dados_arquivo[1].cell(row=linha, column=9).value
            dias_operacao = dados_arquivo[1].cell(row=linha, column=21).value
            tupla_dados = (embarcacao, icj, tipo, regional, regional_1, round(dias_operacao, 3))
            lista_dados.append(tupla_dados)
        return lista_dados

    def organizar(self):
        """Organiza a lista de dados de medição em ordem alfabética e escreve na planilha destino."""
        dados_arquivo = Planguia_funcoes.openr(self.arquivo, self.aba1)
        dados_arquivo[0].create_sheet('Previa')
        nova_aba = dados_arquivo[0].create_sheet('Base Dados')
        organizado = sorted(self.preparar_tupla())
        dict_dados = self.documentar_equipamento()
        linha_preenchida = 2
        for t in organizado:
            nova_aba.cell(row=linha_preenchida, column=2).value = t[0]
            nova_aba.cell(row=linha_preenchida, column=3).value = t[1]
            nova_aba.cell(row=linha_preenchida, column=4).value = t[2]
            nova_aba.cell(row=linha_preenchida, column=5).value = t[3]
            nova_aba.cell(row=linha_preenchida, column=6).value = t[4]
            nova_aba.cell(row=linha_preenchida, column=7).value = t[5]
            icj_extraido = t[1]
            equip = Planguia_funcoes.atribuir_equipamento(icj_extraido, dict_dados)
            nova_aba.cell(row=linha_preenchida, column=1).value = equip
            linha_preenchida = linha_preenchida + 1

        cabecalho = ['Equipamento', 'Embarcação', 'ICJ', 'Tipo', 'Regional', 'Regional1', 'Dias']
        for v1 in range(1, len(cabecalho) + 1):
            nova_aba.cell(row=1, column=v1).value = cabecalho[v1 - 1]
        Planguia_funcoes.closer('Projeto_planilha_Guia_Medição_2020.xlsx', dados_arquivo[0])

    @staticmethod
    def agregar_valores():
        Planguia_funcoes.agregar_dados()
        arquivo_excel = r'C:\Users\ay4m\Desktop\Python\projetos\Projeto_planilha_Guia_Medição_2020_1.xlsx'
        planilha = pd.read_excel(arquivo_excel, sheet_name='Base Dados')

        planilha_nova = planilha[['Embarcação', 'Regional1', 'Equipamento', 'Regional', 'Tipo', 'Dias']]
        destino = pd.ExcelWriter('Projeto_planilha_Guia_Medição_2020_1.xlsx')
        planilha_nova.to_excel(destino, 'Base Dados', index=False)
        destino.save()

    @staticmethod
    def ajustar_celulas():
        wb = Planguia_funcoes.openr('Projeto_planilha_Guia_Medição_2020_1.xlsx', 'Base Dados')
        wb2 = Planguia_funcoes.openr('Projeto_planilha_Guia_Medição_2020.xlsx', 'Previa')
        contar_linha = wb[1].max_row
        for n in range(2, contar_linha + 1):
            for t in range(1, 4):
                atual = wb[1].cell(row=n, column=t).value
                if atual is None:
                    anterior = wb[1].cell(row=n - 1, column=t).value
                    wb[1].cell(row=n, column=t).value = anterior

        for linha in range(1, contar_linha + 1):
            for coluna in range(1, 7):
                wb2[1].cell(row=linha, column=coluna).value = wb[1].cell(row=linha, column=coluna).value
        Planguia_funcoes.closer('Projeto_planilha_Guia_Medição_2020.xlsx', wb2[0])
        wb[0].close()


planguia = PlanilhaInicial('Planilha Guia_dados.xlsx', 'Info Contrato', 'GOPI')
lista_realizacoes = []
planguia.organizar()
planguia.agregar_valores()
planguia.ajustar_celulas()
Planguia_funcoes.mostrar_desempenho(3, 0)
