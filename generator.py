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

# Steg 2: Generera QR-koder och lägg upp till 4 på varje PDF-sida
c = None
pdf_index = 0
qr_count = 0
positions = [(100, 500), (350, 500), (100, 250), (350, 250)]  # Positioner för QR-koder på sidan
start_index = 1  # Håller reda på startindex för varje PDF-fil, börjar med 1 istället för 0

for index, row in df.iterrows():
    link = row['Länk']  # Byt ut med rätt kolumnnamn om nödvändigt
    print(f"Genererar QR-kod för länk: {link}")
    qr = qrcode.make(link)

    # Spara QR-koden som en tillfällig bildfil
    qr_image_path = os.path.join(output_dir, f'qr_code_{index + 1}.png')
    qr.save(qr_image_path)
    print(f"QR-kod sparad som bild: {qr_image_path}")

    if qr_count % 4 == 0:
        if c:
            pdf_path = os.path.join(output_dir, f'qr_codes_{start_index}-{index}.pdf')
            c.save()
            print(f"PDF-fil sparad: {pdf_path}")
        pdf_index += 1
        start_index = index + 1  # Uppdatera startindex till nästa QR-kod, justerat till 1-baserad
        pdf_path = os.path.join(output_dir, f'qr_codes_{start_index}-{start_index+3}.pdf')
        c = canvas.Canvas(pdf_path, pagesize=letter)
        print(f"Skapar ny PDF-fil: {pdf_path}")

    pos_x, pos_y = positions[qr_count % 4]
    c.drawImage(qr_image_path, pos_x, pos_y, width=200, height=200)
    c.drawString(pos_x, pos_y - 20, link)  # Lägg till länken som text under QR-koden
    qr_count += 1

    # Ta bort den tillfälliga bildfilen
    os.remove(qr_image_path)
    print(f"Tillfällig bildfil borttagen: {qr_image_path}")

# Spara sista PDF-filen med korrekt namn
if c:
    pdf_path = os.path.join(output_dir, f'qr_codes_{start_index}-{index + 1}.pdf')
    c.save()
    print(f"Sista PDF-fil sparad: {pdf_path}")

print("QR-koder genererade och sparade som PDF-filer.")
