import win32com.client
from datetime import date

#---------------------
data_de_hoje = date.today()
outlook = win32com.client.Dispatch("Outlook.Application") # apenas liga o python ao outlook 
print("1 - conectado ao outlook") 
#---------------------

mail = outlook.CreateItem(0) #aqui ele segue a arvore, sendo outlook, ou outloo. aberto e crate item, que cria um item, o 0 significa o tipo de itrem, e 0 nesse caso significa email, entao a variavel email tem como resultado a criação de um email dentro do outlook 
print("2 - mail criado") 

mail.subject = "assunto teste"
assunto_do_email = mail.Subject
print(f"3 - o assunto do email é: {assunto_do_email}")


mail.Body = f" Bom dia, seguem anexos os mapas de faturamento de julho, atualizados até {data_de_hoje}, (Lembrando que as metas individuais não estão atualizadas), att Murilo MoraeS"
corpo_do_email = mail.body
print(f"4 - o corpo do email é: {corpo_do_email}")

mail.To = 



#   Application
#   │
#   └── GetNamespace("MAPI")  → Namespace
#              │
#              ├── Folders (coleção de todas as contas/pastas raiz)
#              │
#              └── GetDefaultFolder(N) → um MAPIFolder específico
#                        │
#                        ├── Items (coleção de e-mails/itens dessa pasta)
#                        │        └── um item específico → MailItem, AppointmentItem, ContactItem...
#                        │
#                        └── Folders (subpastas dentro dessa pasta)