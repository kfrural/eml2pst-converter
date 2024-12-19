import os
import win32com.client
from email import parser

def convert_eml_to_pst(eml_file_path):
    outlook = win32com.client.Dispatch('Outlook.Application')
    
    pst_file = os.path.join(os.getcwd(), 'output.pst')
    mailbox = outlook.GetNamespace('MAPI')
    
    inbox = mailbox.GetDefaultFolder(6)
    folder = inbox.Folders.Add('Converted Emails')
    
    with open(eml_file_path, 'r') as eml_file:
        email_message = parser.Parser().parsestr(eml_file.read())
        
        message = folder.Items.Add("IPM.Note")
        for header in ['Subject', 'From', 'To']:
            if header in email_message:
                setattr(message, header, email_message[header])
        
        body = email_message.get_payload()
        if isinstance(body, list): 
            body = ''.join([part.get_payload(decode=True).decode('utf-8') for part in body])
        else:
            body = body.decode('utf-8') if isinstance(body, bytes) else body
        
        message.Body = body
        message.Save()

def convert_emls_to_pst(eml_folder_path, max_files=1000):
    count = 0
    for filename in os.listdir(eml_folder_path):
        if filename.endswith('.eml'):
            eml_file_path = os.path.join(eml_folder_path, filename)
            try:
                convert_eml_to_pst(eml_file_path)
                count += 1
                print(f"Convertido: {count}/{max_files}")
                
                if count >= max_files:
                    break
            except Exception as e:
                print(f"Erro ao converter {filename}: {str(e)}")
    
    print("Conversão concluída.")

if __name__ == "__main__":
    eml_folder_path = r"C:\caminho\para\seus\arquivos\eml"
    convert_emls_to_pst(eml_folder_path)
