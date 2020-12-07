class PlanGuia:
    def __init__(self, df):
        self.df = df

    def formatacao(self):
        df_new = self.df.fillna(0)
        listdel = ['PRL Petro', 'Medir Petro', 'PRL PBLOG', 'Medir PBLOG', 'Objeto de Custo',
                   'Autorizado', 'Observações', 'Cessão', 'Gerente', 'Fiscal']
        df_new = df_new.drop(listdel, axis=1)
        df_new = df_new[['Equipamento', 'Embarcação', 'Classe', 'Porte', 'Regional', 'Regional CMAR',
                         'Dias Medir', 'Indisp', 'Medir', 'Centro de Custo']]
        listcolumn = ['Dias Medir', 'Indisp', 'Medir']
        df = df_new.copy()
        for n in listcolumn:
            df[n] = df_new[n].round(3)
        df['Embarcação'] = df_new.Embarcação.str.upper()
        df.to_excel(r'C:\Users\Julio\Desktop\teste\analise_planguia.xlsx', index=False)
