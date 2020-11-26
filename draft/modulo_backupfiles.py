import os
import time
import shutil
from openpyxl import load_workbook
from modulo_calculo import CalcPlanGui

class Backup(CalcPlanGui):
    def __init__(self, pathorigin, pathdest, finalname):
        super().__init__(pathorigin, pathdest)
        self. orig = r'C:\Users\ay4m\Desktop\planguia'
        self. dest = r'C:\Users\ay4m\Desktop'
        self.namefile = 'arquivo_editado.xlsx'
        self.finalname = finalname
        # self.adress = adress

    def verificar_pasta(self):
        atual = time.localtime()
        pasta_principal = os.path.join(self.dest, str(atual[0]))
        if not os.path.isdir(pasta_principal):
            os.mkdir(pasta_principal)
        nome_pasta = ['1-Janeiro', '2-Fevereiro', '3-Março', '4-Abril', '5-Maio', '6-Junho', '7-Julho', '8-Agosto',
                      '9-Setembro', '10-Outubro', '11-Novembro', '12-Dezembro']
        for month in nome_pasta:
            adress = os.path.join(pasta_principal, month)
            if not os.path.isdir(adress):
                os.mkdir(adress)
        mes_atual = atual[1]
        pasta_atual = nome_pasta[mes_atual - 1]
        endereco = os.path.join(pasta_principal, pasta_atual)
        return endereco


    def mover_arquivo(self):
        endereco_final = self.verificar_pasta()
        # origem = os.path.abspath('.')
        arquivo_fonte = os.path.join(self.orig, self.namefile)
        arquivo_mover = os.path.join(endereco_final, self.finalname)
        if os.path.exists(arquivo_mover):
            os.remove(arquivo_mover)
        shutil.copyfile(arquivo_fonte, arquivo_mover)
        # os.rename(os.path.join(self.orig, self.namefile), os.path.join(endereco_final, self.finalname))
        return arquivo_mover

    @staticmethod
    def preparar_email(adress):
        wb = load_workbook(adress)
        listsheet = ['Porte', 'Taxa Diária', 'Previa', 'Info Contrato - Pblog', 'Chaves']
        for sheet in listsheet:
            aba = wb.get_sheet_by_name(sheet)
            wb.remove_sheet(aba)
        # aba_taxa = wb.get_sheet_by_name('Taxa Diária')
        # wb.remove_sheet(aba_taxa)
        # aba_previa = wb.get_sheet_by_name('Previa')
        # wb.remove_sheet(aba_previa)
        # aba_info_pblog = wb.get_sheet_by_name('Info Contrato - Pblog')
        # wb.remove_sheet(aba_info_pblog)
        # aba_chaves = wb.get_sheet_by_name('Chaves')
        # wb.remove_sheet(aba_chaves)
        wb.save(adress)

# destiny = r'C:\Users\ay4m\Desktop'
# # finalsheet = verificar_pasta(destiny)
# # mover_arquivo(finalsheet, 'julio.xlsx')
# sheetadress = r'C:\Users\ay4m\Desktop\planguia'
# namefile = 'arquivo_editado.xlsx'
# mover_arquivo(sheetadress, destiny, namefile)
# x=r'C:\Users\ay4m\Desktop\planguia\arquivo_editado.xlsx'