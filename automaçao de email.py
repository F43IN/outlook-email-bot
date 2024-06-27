import pandas as pd
import win32com.client as win32

file_path = 'path/for/your/sheets.xlsx'
sheet_name = 'sheet_name'

df = pd.read_excel(file_path, sheet_name=sheet_name)

outlook = win32.Dispatch('outlook.application')

for index, row in df.iterrows():
    email_address = row['enter_column_name']

    mail = outlook.CreateItem(0)
    mail.To = email_address
    mail.Subject = 'Assunto do Email'
    mail.HTMLBody = """Corpo do email com tags HTML"""

    # Adicione anexos se necessário
    # attachment = 'way/for/your/attachment'
    # mail.Attachments.Add(attachment)

    mail.Send()

print('Emails enviados com sucesso!')
