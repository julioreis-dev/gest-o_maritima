import pandas as pd

def merge():
    df1 = pd.read_excel(r'C:\Users\Julio\Desktop\teste\analise_monitoramento.xlsx')
    df2 = pd.read_excel(r'C:\Users\Julio\Desktop\teste\analise_planguia.xlsx')
    outputfile = r'C:\Users\Julio\Desktop\teste\Relatório_CMAR.xlsx'
    df3 = pd.merge(df1, df2, on='Embarcação', how='left')
    df_final = df3[['Equipamento', 'Embarcação', 'Classe', 'Porte', 'Regional', 'Regional CMAR', 'Dias Medir',
                    'Indisp', 'Medir', 'Centro de Custo', 'Valor Petrobras $', 'Valor PBLOG $', 'Total $']]
    df_final = df_final.drop_duplicates('Equipamento', keep='first')
    # df_final['Taxa Diária'] = (df_final['Total $'] / df_final['Medir']).round(2)
    df_final.to_excel(outputfile, index=False)
