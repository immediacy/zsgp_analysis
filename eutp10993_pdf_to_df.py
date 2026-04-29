import tabula

pdf_path = (r"G:\Мой диск\Tasks"
            r"\EUTP-109993 Формирование отчета комиссионера для Агрегаторов (Домклик)"
            r"\Акт_отчет.pdf")

df = tabula.read_pdf(pdf_path, encoding='cp1252', pages='all')

df_0 = df[0]
df_1 = df[1]
df_2 = df[2]
df_3 = df[3]
df_4 = df[4]

df_0 = df_0.rename(columns={x: df_0.loc[0, x] for x in df_0.columns})
df_0.drop([0, 1], axis=0, inplace=True)
df_0.reset_index(drop=True, inplace=True)
df_0.to_excel(r"G:\Мой диск\Tasks"
              r"\EUTP-109993 Формирование отчета комиссионера для Агрегаторов (Домклик)"
              r"\first.xlsx")
print(df_0[1])

