import smtplib, ssl, csv, argparse
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

SMTP_PORT = 587
SMTP_SERVER = "smtp.gmail.com"

def send_email(sender_email, receiver_email, password, email_data):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = email_data['subject']

        body = email_data['body']
        msg.attach(MIMEText(body, 'plain'))

        context = ssl.create_default_context()
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls(context=context)
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        
        return True
    except Exception as e:
        print(f'Failed to send email: {e}')
        return False

def read_csv(path) -> list[dict]:
    try:
        result_data = []
        with open(path, newline='', encoding='latin-1') as csvfile:
            csv_data = csv.reader(csvfile, delimiter=';',)
            for row in csv_data:
                result_data.append({
                    "subject": row[0],
                    "body": (';').join(row)
                })
        return result_data
    except Exception as e:
        print(f'Failed to read csv file: {e}')
        return None

if __name__ == '__main__':
    parser = argparse.ArgumentParser(prog='Excel Automation to Trello Board', 
        description='Automation created to send csv files from excel worksheet to trello using email system.')
    
    parser.add_argument('-s', '--sender', type=str, required=True, help='Email which its going to be used to send the emails')
    parser.add_argument('-p', '--password', type=str, required=True, help='Sender email password')
    parser.add_argument('-r', '--receiver', type=str, required=True, help='Trello Email which will receive the emails')
    parser.add_argument('-f', '--file', type=str, required=True, help='Path of the csv file')

    args = parser.parse_args()

    print('Reading csv...')
    csv_data = read_csv(args.file)

    if csv_data is None:
        input(f'Failed to load {args.file}.\nNot sending email!\n\nPress any key to continue...')
        exit(-1)
    
    print(f'Sending email from {args.sender} to {args.receiver}...')
    print('Records found:')

    for item in csv_data:
        print(f'{item}\n')
        success = send_email(args.sender, args.receiver, args.password, item)
        if not success:
            input(f'Failed to send email.\nNot sending anymore!\n\nPress any key to continue...')
            exit(-1)
    
    input('Press any key to continue...')
    exit(0)