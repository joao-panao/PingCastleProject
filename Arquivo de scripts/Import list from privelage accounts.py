import openpyxl
from bs4 import BeautifulSoup

with open(r'C:\Users\joao.panao\OneDrive - RIS 2048\Documentos\GitHub\PingCastleProject\Ping Castle Reports\ad_hc_cmas.local.html', encoding="utf8") as file:
    soup = BeautifulSoup(file, "html.parser")

priv_section = soup.find('div', id='sectionPrivilegedAccounts')
priv_accounts = priv_section.find_all('button')

wb = openpyxl.Workbook()
ws = wb.active

for idx, account in enumerate(priv_accounts):
    ws.cell(row=idx+1, column=1, value=account.text)

wb.save(r'C:\Users\joao.panao\OneDrive - RIS 2048\Documentos\GitHub\PingCastleProject\Extracted Tables\Privileged Accounts.xlsx')
