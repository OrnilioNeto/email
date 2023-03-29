from imapclient import IMAPClient
import email
import os
import re
from datetime import datetime, date
import quopri


# Configurações de login
HOST = 'imap.gmail.com'
USERNAME = 'orniliooficial@gmail.com'
PASSWORD = ''

# Diretório de download
download_dir = 'downloads'

def sanitize_foldername(foldername):
    """
    Remove caracteres inválidos do nome da pasta.
    """
    return re.sub(r'[<>:"/\|?*]', '', foldername).strip()


def sanitize_filename(filename):
    """
    Remove caracteres inválidos do nome do arquivo.
    """
    return re.sub(r'[<>:"/\|?*]', '', filename.replace(' ', '_')).strip()


# Abrir conexão com servidor gmail
with IMAPClient(HOST) as server:
    server.login(USERNAME, PASSWORD)
    server.select_folder('INBOX', readonly=False)
    # Pesquisa
    messages = server.search([u'SINCE', date(datetime.now().year, datetime.now().month, datetime.now().day)])
    
    for uid, message_data in server.fetch(messages,'RFC822').items():
        email_message = email.message_from_bytes(message_data[b'RFC822'])
        print(f'id: {uid}')
        print(f'FROM: {email_message.get("From")}')
        print(f'TÓPICO: {email_message.get("Subject")}')
        # Extrair nome do remetente
        sender_email = email_message.get('From')
        sender_name = re.search(r'"?([^"]+)"?\s?<', str(sender_email)).group(1) if '<' in str(sender_email) else str(sender_email)


        # Criar pasta com nome do remetente, se não existir
        sender_folder = os.path.join(download_dir, sanitize_foldername(sender_name.replace(' ', '_')))


        if not os.path.exists(sender_folder):
            os.makedirs(sender_folder)

        # Baixar anexos, se houverem
        for part in email_message.walk():
            if part.get_content_disposition() == 'attachment':
                filename = part.get_filename()
                filepath = os.path.abspath(os.path.join(sender_folder, sanitize_filename(quopri.decodestring(filename).decode('utf-8'))))



                # Salvar arquivo no diretório correspondente
                with open(filepath, 'wb') as f:
                    f.write(part.get_payload(decode=True))



        # Marcar email como lido
        server.add_flags(uid, [b'Seen'])
