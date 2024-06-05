import pandas as pd
import qrcode
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os

# Steg 1: Läs Excel-filen för att få länkarna
file_path = 'lankar.xlsx'  # Byt ut med din Excel-fils namn
sheet_name = 'Sheet1'  # Byt ut med rätt bladnamn om nödvändigt

print(f"Läser in Excel-filen från: {file_path}")

# Försöker läsa in Excel-filen och specificerar bladnamnet
try:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    print("Excel-filen inläst. Här är data:\n", df)
except Exception as e:
    print(f"Kunde inte läsa Excel-filen: {e}")
    exit()

# Kontrollera om DataFrame är tom
if df.empty:
    print("DataFrame är tom. Kontrollera att Excel-filen innehåller data.")
    exit()

# Kontrollera att rätt kolumn finns
if 'Länk' not in df.columns:
    print("Kolumnen 'Länk' finns inte i Excel-filen. Kontrollera filstrukturen.")
    exit()

# Skapa en mapp för att spara PDF-filerna
output_dir = 'qr_codes_pdfs'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
print(f"PDF-filer kommer att sparas i: {output_dir}")

# Steg 2: Generera QR-koder för varje länk
for index, row in df.iterrows():
    link = row['Länk']  # Byt ut med rätt kolumnnamn om nödvändigt
    print(f"Genererar QR-kod för länk: {link}")
    qr = qrcode.make(link)

    # Spara QR-koden som en tillfällig bildfil
    qr_image_path = os.path.join(output_dir, f'qr_code_{index}.png')
    qr.save(qr_image_path)
    print(f"QR-kod sparad som bild: {qr_image_path}")

    # Skapa en PDF-fil och lägg till QR-koden
    pdf_path = os.path.join(output_dir, f'qr_code_{index}.pdf')
    print(f"Skapar PDF-fil: {pdf_path}")
    c = canvas.Canvas(pdf_path, pagesize=letter)
    c.drawImage(qr_image_path, 100, 500, width=200, height=200)
    c.drawString(100, 480, link)  # Lägg till länken som text under QR-koden
    c.save()
    print(f"PDF-fil sparad: {pdf_path}")

    # Ta bort den tillfälliga bildfilen
    os.remove(qr_image_path)
    print(f"Tillfällig bildfil borttagen: {qr_image_path}")

print("QR-koder genererade och sparade som PDF-filer.")
