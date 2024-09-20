import smtplib, ssl, argparse, openpyxl
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

def read_worksheet(workbook_path, worksheet_name, selection):
    try:
        wb = openpyxl.load_workbook(workbook_path, data_only=True)
        worksheet = wb[worksheet_name]
        firstCell, lastCell = selection.split(':')

        if not firstCell or not lastCell:
            raise Exception(f'Selection value {selection} is not valid!') 

        range = worksheet[firstCell:lastCell]

        result_data = []

        for row in range:
            subject = ''
            body = ''
            for index, cell in enumerate(row):
                cell_value = str(cell.value).replace('\n', ' ')
                
                if index == 0:
                    subject = cell_value

                body += cell_value + (' | ' if index < len(row)-1 else '')

            result_data.append({
                "subject": subject,
                "body": body
            })
        
        return result_data
    except Exception as e:
        print(f'Failed to read workbook: {e}')
        return None

def validate_args(args):
    if not args.email:
        print('E-mail is required.')
        return False
    
    if not args.email.__contains__('@'):
        print('E-mail is not valid.')
        return False
    
    if not args.password:
        print('Password is required.')
        return False
    
    if not args.trello_email:
        print('Trello e-mail is required.')
        return False
    
    if not args.trello_email.__contains__('@'):
        print('Trello e-mail is not valid.')
        return False
    
    if not args.workbook:
        print('Workbook file path is required.')
        return False
    
    if not args.worksheet:
        print('Worksheet name is required. Example: Sheet1')
        return False
    
    if not args.selection:
        print('Worksheet selection is required. Example: A1:F10')
        return False
    
    return True

def print_press_to_continue(exit_code, message = ''):
    print(message)
    input('\n\nPress any key to continue...')
    exit(exit_code)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(prog='Excel Automation to Trello Board', 
        description='Automation created to send csv files from excel worksheet to trello using email system.')
    
    parser.add_argument('-e', '--email', type=str, help='Email which its going to be used to send the emails')
    parser.add_argument('-p', '--password', type=str, help='Sender email password')
    parser.add_argument('-t', '--trello_email', type=str, help='Trello Email which will receive the emails')

    parser.add_argument('-b', '--workbook', type=str, help='Workbook path (Ex.: ~/home/Book1.xlsx)')
    parser.add_argument('-w', '--worksheet', type=str, help='Workshet name (Ex.: Sheet1)')
    parser.add_argument('-s', '--selection', type=str, help='Worksheet selection (Ex.: A1:F10)')

    args = parser.parse_args()
    print('Arguments:', args)

    validated = validate_args(args)

    if not validated:
        print_press_to_continue(1, f'Args are not valid. Please verify!')

    print('Reading worksheet...')
    selection_data = read_worksheet(args.workbook, args.worksheet, args.selection)

    if selection_data is None:
        print_press_to_continue(1, f'Failed to load {args.workbook}.\nNot sending email!')
    
    print(f'Sending email from {args.email} to {args.trello_email}...')
    print('Records found:')

    for item in selection_data:
        print(f'{item}\n')
        success = send_email(args.email, args.trello_email, args.password, item)

        if not success:
            print_press_to_continue(1, f'Failed to send one email.\nNot sending anymore!')
    
    print_press_to_continue(0)