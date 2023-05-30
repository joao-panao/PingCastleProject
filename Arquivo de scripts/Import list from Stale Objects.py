from bs4 import BeautifulSoup
import openpyxl

# Read the HTML file
with open("C:/Users/joao.panao/OneDrive - RIS 2048/Documentos/GitHub/PingCastleProject/Ping Castle Reports/ad_hc_cmas.local.html", encoding="utf8") as f:
    html = f.read()

# Parse the HTML using BeautifulSoup
soup = BeautifulSoup(html, "html.parser")

# Find the div with ID="rulesStaleObjects"
div = soup.find("div", {"id": "rulesStaleObjects"})

# Get the text of the buttons inside the div
button_texts = [button.get_text(strip=True) for button in div.find_all("button")]

# Create a new Excel workbook and worksheet
wb = openpyxl.Workbook()
ws = wb.active

# Write the button texts to the worksheet
for i, text in enumerate(button_texts):
    ws.cell(row=i+1, column=1, value=text)

# Save the workbook
wb.save("C:/Users/joao.panao/OneDrive - RIS 2048/Documentos/GitHub/PingCastleProject/Extracted Tables/Stale Objects.xlsx")
