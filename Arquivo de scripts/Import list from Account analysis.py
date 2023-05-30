import os
import pandas as pd
from bs4 import BeautifulSoup

# specify the file path
file_path = r"C:\Users\joao.panao\OneDrive - RIS 2048\Documentos\GitHub\PingCastleProject\Ping Castle Reports\ad_hc_cmas.local.html"

# open the file and create a BeautifulSoup object
with open(file_path, encoding='utf-8') as file:
    soup = BeautifulSoup(file, 'html.parser')

# find the div with id 'usersaccordion'
usersaccordion_div = soup.find('div', {'id': 'usersaccordion'})

# find all the button tags in the usersaccordion div
button_tags = usersaccordion_div.find_all('button')

# get the text from each button tag
button_text = [button.get_text(strip=True) for button in button_tags]

# find all the i tags in the usersaccordion div
i_tags = usersaccordion_div.find_all('i')

# get the numbers from each i tag
i_text = [i.get_text(strip=True)[1:-1] for i in i_tags]

# create a dataframe with two columns: Button Text and I Text
df = pd.DataFrame({'Button Text': button_text, 'I Text': i_text})

# create a file path to save the excel file
save_path = r"C:\Users\joao.panao\OneDrive - RIS 2048\Documentos\GitHub\PingCastleProject\Extracted Tables\Account analysis.xlsx"

# save the dataframe as an excel file
df.to_excel(save_path, index=False, header=False)
