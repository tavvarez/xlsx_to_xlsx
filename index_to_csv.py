import pandas as pd

def convert_excel_to_csv(excel, csv, colunas):
    df = pd.read_excel(excel)

    df_selecionado = df[colunas]

    df_selecionado.to_csv(csv, index = False)

excel = 'C:/Users/Bialog-006/desktop/PlanilhasBOT/SUZANO_MUCURI_BA.xlsx'
csv = 'C:/Users/Bialog-006/desktop/PlanilhasBOT/Planilha.csv'
colunas_desejadas = ['ORIGEM', 'UF', 'DESTINO', 'VEICULO', 'FRETE']

convert_excel_to_csv(excel, csv, colunas_desejadas)