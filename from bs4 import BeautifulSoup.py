from bs4 import BeautifulSoup
import pandas as pd

with open('C:/Users/joao.panao/OneDrive - RIS 2048/Ambiente de Trabalho/Test/ad_hc_cmas.local.html', encoding='utf-8') as f:
    soup = BeautifulSoup(f, 'html.parser')

button_texts = [button.text.strip() for button in soup.find_all('button', class_='btn-link')]

df = pd.DataFrame(button_texts, columns=['Button Text'])
df.to_excel('C:/Users/joao.panao/OneDrive - RIS 2048/Ambiente de Trabalho/output.xlsx', index=False)
