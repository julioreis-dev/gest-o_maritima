import pandas as pd
import Fplanguia as fp
import mergedatas as mg


class PlanMonitor:
    def __init__(self, df, dolar):
        self.df = df
        self.dolar = dolar

    def formatacao(self):
        df_new = self.df.fillna(0)
        listcolumns = ['EMBARCAÇÃO\nPARCELA BRL\n(COM REAJUSTE)', 'EMBARCAÇÃO\nPARCELA BRL\n(COM REAJUSTE).1']
        df_new['Valor Petrobras $'] = df_new[listcolumns[0]] / self.dolar
        df_new['Valor Petrobras $'] = df_new[['EMBARCAÇÃO\nPARCELA USD', 'Valor Petrobras $']].sum(axis=1).round(2)
        df_new['Valor PBLOG $'] = df_new[listcolumns[1]] / self.dolar
        df_new['Valor PBLOG $'] = df_new[['EMBARCAÇÃO PARCELA USD', 'Valor PBLOG $']].sum(axis=1).round(2)
        df_new['Total $'] = df_new[['Valor Petrobras $', 'Valor PBLOG $']].sum(axis=1).round(2)
        df_new = df_new[['Embarcação', 'Valor Petrobras $', 'Valor PBLOG $', 'Total $']]
        df = df_new.copy()
        df['Embarcação'] = df_new.Embarcação.str.upper()
        df.to_excel(r'C:\Users\Julio\Desktop\teste\analise_monitoramento.xlsx', index=False)


if __name__ == '__main__':
    cotacao = float(input('Digite a cotação do Dolar: '))
    listsheets = ['GTLQ - MC', 'Planilha_guia']
    for n in range(0, 2):
        frame = pd.read_excel(r'C:\Users\Julio\Desktop\teste\analise_CMAR.xlsx', sheet_name=listsheets[n], skiprows=[0])
        if n == 0:
            plan = PlanMonitor(frame, cotacao)
            plan.formatacao()
        else:
            plan = fp.PlanGuia(frame)
            plan.formatacao()
    mg.merge()
