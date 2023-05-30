import tkinter as tk
from tkinter import filedialog
from bs4 import BeautifulSoup
import cairosvg
from PIL import Image

# Create a Tkinter window to select the HTML file
root = tk.Tk()
root.withdraw()

# Prompt the user to select an HTML file using a file dialog
html_file_path = filedialog.askopenfilename(title="Select HTML file", filetypes=[("HTML files", "*.html")])

# Open the HTML file and parse it with Beautiful Soup
with open(html_file_path, 'r', encoding='utf-8') as f:
    soup = BeautifulSoup(f.read(), 'html.parser')

# Find the SVG element in the HTML code
svg_element = soup.find('svg')

# Convert the SVG to a PNG
png_data = cairosvg.svg2png(bytestring=str(svg_element))

# Convert the PNG to a JPEG
image = Image.open(io.BytesIO(png_data))
jpeg_data = io.BytesIO()
image.save(jpeg_data, format='JPEG')

# Prompt the user to select a save location and filename for the JPEG file
jpeg_file_path = filedialog.asksaveasfilename(title="Save JPEG file", defaultextension=".jpg", filetypes=[("JPEG files", "*.jpg")])

# Save the JPEG data to a file
with open(jpeg_file_path, 'wb') as f:
    f.write(jpeg_data.getvalue())
