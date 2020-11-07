import pandas as pd


class PlanGuia:
    def __init__(self, df):
        self.df = df

    def formatacao(self):
        df_new = self.df.fillna(0)
        listdel = ['PRL Petro', 'Medir Petro', 'PRL PBLOG', 'Medir PBLOG', 'Objeto de custo',
                   'Liberada', 'Observações']
        df_new = df_new.drop(listdel, axis=1)
        df_final = df_new[['Equipamento', 'Embarcação', 'Tipo', 'Porte', 'Regional', 'Regional1', 'Dias', 'Indisp',
                           'Medir', 'Centro de Custo']]
        listcolumn = ['Dias', 'Indisp', 'Medir']
        for n in listcolumn:
            df_final[n] = df_final[n].round(3)
        df_final['Embarcação'] = df_final['Embarcação'].str.upper()
        df_final.to_excel(r'C:\Users\Julio\Desktop\teste\analise_planguia.xlsx', index=False)
