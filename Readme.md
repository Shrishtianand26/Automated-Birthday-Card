# 🎉 Automated Birthday Wishing System

This Python-based project was developed during an internship at **Indian Oil Corporation Ltd.** It automates the process of wishing employees on their birthdays by generating personalized greeting cards and emailing them automatically.

## 🚀 Features

- Reads employee data from Excel
- Generates birthday cards using Pillow
- Sends emails with the cards attached
- Scheduled to run daily via Task Scheduler
- Prevents duplicate runs using a lock file

## 🛠️ Technologies

- Python 3.13  
- pandas  
- Pillow (PIL)  
- smtplib & email.mime  
- Windows Task Scheduler

## 📝 Setup

1. Install requirements:

   ```bash
   pip install pandas pillow

2. Update your data.xlsx file with columns:
NAME, DOB (DD-MM), EMAIL

3. Update your email credentials in the script:
server.login("your-email@example.com", "app-password")

4. Set up a Task Scheduler job:
Program/script: python
Arguments: birthday.py
Start in: folder path where script is saved

📁 Output
Logs: script_log.txt
Cards: saved in birthday_cards/
Email: with personalized card attached