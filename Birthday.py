import os
import sys
import pandas as pd
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# === LOCK FILE to prevent multiple runs ===
LOCK_FILE = "script.lock"

def create_lock():
    if os.path.exists(LOCK_FILE):
        print("Script is already running or was not closed properly.".encode(sys.stdout.encoding or 'utf-8', errors='replace').decode())
        sys.exit()
    with open(LOCK_FILE, "w") as f:
        f.write("locked")

def remove_lock():
    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)

# ==== CONFIGURATION ====
TEMPLATE_PATH = "birthday-template.png"
CARD_FOLDER = "birthday_cards"
FONT_PATH = "arial.ttf"
BOLD_FONT_PATH = "arialbd.ttf"
FONT_SIZE_NAME = 40
FONT_SIZE_HR = 36  

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "shrishtianand26@gmail.com"
SENDER_PASSWORD = "hxgw pnrl auvq zwaw" 

# ==== Ensure card folder exists ====
if not os.path.exists(CARD_FOLDER):
    os.makedirs(CARD_FOLDER)


def create_birthday_card(name, hr_name, save_path):
    img = Image.open(TEMPLATE_PATH)
    draw = ImageDraw.Draw(img)

    font_name = ImageFont.truetype(FONT_PATH, FONT_SIZE_NAME)
    font_hr = ImageFont.truetype(BOLD_FONT_PATH, FONT_SIZE_HR)

    # Centered Name
    bbox = draw.textbbox((0, 0), name, font=font_name)
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]
    width, height = img.size
    x_name = (width - w) // 2
    y_name = int(height * 0.41)
    draw.text((x_name, y_name), name, font=font_name, fill="black")

    # HR Message
    x_hr = 50
    y_hr = height - 100
    hr_text = f"Best wishes,\n{hr_name} (HR Team)"
    draw.multiline_text((x_hr, y_hr), hr_text, font=font_hr, fill="black")
    img.save(save_path)

    #Automate Email
def send_birthday_email(name, recipient_email, card_path):
    from_address = "shrishtianand26@gmail.com"      #Senders mail id
    app_password = "Your App password" #16 digit app_password     

    print(f"[EMAIL] Preparing email for {name} ({recipient_email})...")

    # Create the message
    msg = MIMEMultipart()
    msg['From'] = from_address
    msg['To'] = recipient_email
    msg['Subject'] = f"🎉 Happy Birthday {name}!"

    body = f"""
    <p>Dear {name},</p>
    <p>Wishing you a day filled with happiness and a year filled with joy. 🎂🎁</p>
    <p>Have a fantastic birthday!</p>
    <p>Best regards,<br>Your Birthday Bot 🤖</p>
    """
    msg.attach(MIMEText(body, 'html'))

    # Attach the card
    if os.path.exists(card_path):
        with open(card_path, "rb") as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(card_path)}"')
        msg.attach(part)
    else:
        print(f"[WARNING] Card file not found: {card_path}")

    try:
        print("[EMAIL] Connecting to Gmail SMTP server...")
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.ehlo()
        server.starttls()
        server.login(from_address, app_password)
        server.send_message(msg)
        server.quit()
        print(f"[EMAIL] Email successfully sent to {name} at {recipient_email}")
    except Exception as e:
        print(f"[ERROR] Failed to send email to {name}: {e}")

def main():
    print(" Script started!")
    today_str = datetime.today().strftime('%d-%m')
    print(f" Today's date: {today_str}")

    df = pd.read_excel("birthdays.xlsx")
    df["DOB"] = pd.to_datetime(df["DOB"], errors="coerce")

    found_birthday = False

    for _, row in df.iterrows():
        dob_value = row["DOB"]
        if pd.isna(dob_value):
            print(f"[WARNING] Skipping {row.get('NAME', '[Unknown]')} - DOB is invalid or empty.")
            continue

        dob = dob_value.strftime('%d-%m')

        if dob == today_str:
            name = row["NAME"]
            email = row["EMAIL"]
            hr_name = row["HR"]

            print(f"[BIRTHDAY] It's {name}'s birthday today!")
            
            found_birthday = True
            card_filename = f"birthday_card_{name.replace(' ', '_').lower()}.png"
            card_path = os.path.join(CARD_FOLDER, card_filename)

            create_birthday_card(name, hr_name, card_path)

            if os.path.exists(card_path):
                print(f"[CARD] Card created at {card_path}")

                send_birthday_email(name, email, card_path)
            else:
                print(f"⚠️ Card not found for {name}. Skipping email.")

    if not found_birthday:
        print("📭 No birthdays today.")

if __name__ == "__main__":
    create_lock()
    try:
        main() 
    finally:
        remove_lock()


