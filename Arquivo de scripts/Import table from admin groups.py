import os
import pandas as pd
from bs4 import BeautifulSoup

# specify the file path
file_path = r"C:\Users\joao.panao\OneDrive - RIS 2048\Documentos\GitHub\PingCastleProject\Ping Castle Reports\ad_hc_cmas.local.html"

# open the file and create a BeautifulSoup object
with open(file_path, encoding='utf-8') as file:
    soup = BeautifulSoup(file, 'html.parser')

# find the comment with text ' SubSection Groups end '
comment = soup.find(text=' SubSection Groups end ')

# find the first table after the comment
table = comment.find_next('table')

# convert the table to a dataframe
df = pd.read_html(str(table))[0]

# create a file path to save the excel file
save_path = r"C:\Users\joao.panao\OneDrive - RIS 2048\Documentos\GitHub\PingCastleProject\Extracted Tables\admin groups.xlsx"

# save the dataframe as an excel file
df.to_excel(save_path, index=False)
