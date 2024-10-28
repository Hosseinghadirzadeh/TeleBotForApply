import os
import pandas as pd
from telegram import Update
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters, CallbackContext
import smtplib
from email.mime.text import MIMEText

# Email configuration (replace with your email credentials)
EMAIL_ADDRESS = "hosseinghadirzadeh@khu.ac.ir"
EMAIL_PASSWORD = "Hh@32827052"

def start(update: Update, context: CallbackContext):
    update.message.reply_text("Hi! Send me an Excel file with columns 'mail', 'Name', 'University', and 'Research & Interest'.")

def handle_excel(update: Update, context: CallbackContext):
    file = update.message.document.get_file()
    file.download('email_data.xlsx')
    update.message.reply_text("File received. Now send me the text for the email body, using placeholders like [name], {university}, and {interest}.")

    context.user_data['excel_received'] = True

def handle_email_text(update: Update, context: CallbackContext):
    if context.user_data.get('excel_received', False):
        email_body_template = update.message.text
        context.user_data['email_body_template'] = email_body_template
        update.message.reply_text("Email template received. Now specify the columns for 'university' and 'research interests' separated by a comma (e.g., 'University, Research & Interest').")
    else:
        update.message.reply_text("Please send me the Excel file first.")

def handle_columns(update: Update, context: CallbackContext):
    if 'email_body_template' not in context.user_data:
        update.message.reply_text("Please provide the email body template first.")
        return

    columns = update.message.text.split(',')
    if len(columns) != 2:
        update.message.reply_text("Please specify exactly two columns: one for 'university' and one for 'research interests'.")
        return

    university_col = columns[0].strip()
    interest_col = columns[1].strip()

    context.user_data['university_col'] = university_col
    context.user_data['interest_col'] = interest_col

    try:
        # Load the Excel file
        df = pd.read_excel('email_data.xlsx')

        # Check if the required columns exist
        if university_col not in df.columns or interest_col not in df.columns or 'mail' not in df.columns or 'Name' not in df.columns:
            update.message.reply_text(f"One or more columns '{university_col}', '{interest_col}', 'mail', or 'Name' not found in the Excel file.")
            return

        # Extract email addresses and relevant info
        for _, row in df.iterrows():
            name = row['Name']
            university = row[university_col]
            interest = row[interest_col]
            email_address = row['mail']

            # Generate the personalized email body
            email_body = context.user_data['email_body_template'].replace("[name]", name).format(
                university=university,
                interest=interest
            )

            # Send the email
            send_email(email_address, email_body)

        update.message.reply_text("Emails have been sent successfully!")

    except Exception as e:
        update.message.reply_text(f"An error occurred: {str(e)}")

def send_email(to_email, email_body):
    msg = MIMEText(email_body)
    msg['Subject'] = "Research Collaboration Opportunity"
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_email

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, to_email, msg.as_string())
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {str(e)}")

def main():
    updater = Updater("7747359291:AAEuGhkE-RGN8Xli33w_fFhtP71miFpKFkw", use_context=True)
    dispatcher = updater.dispatcher

    dispatcher.add_handler(CommandHandler('start', start))
    dispatcher.add_handler(MessageHandler(Filters.document.mime_type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), handle_excel))
    dispatcher.add_handler(MessageHandler(Filters.text & ~Filters.command, handle_email_text))
    dispatcher.add_handler(MessageHandler(Filters.text & ~Filters.command, handle_columns))

    updater.start_polling()
    updater.idle()

if __name__ == '__main__':
    main()
