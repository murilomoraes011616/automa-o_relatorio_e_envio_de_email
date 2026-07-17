import win32com.client
from datetime import date

#---------------------
data_de_hoje = date.today()
outlook = win32com.client.Dispatch("Outlook.Application") # apenas liga o python ao outlook 
print("1 - conectado ao outlook") 
#---------------------


print("--------")
mail = outlook.CreateItem(0) #aqui ele segue a arvore, sendo outlook, ou outloo. aberto e crate item, que cria um item, o 0 significa o tipo de itrem, e 0 nesse caso significa email, entao a variavel email tem como resultado a criação de um email dentro do outlook 
print("2 - mail criado") 
print("--------")


print("--------")
mail.subject = "assunto teste" # feito para definir o assunto do email
assunto_do_email = mail.Subject
print(f"3 - o assunto do email é: {assunto_do_email}")
print("--------")


print("--------")
mail.Body = f" Bom dia, seguem anexos os mapas de faturamento de julho, atualizados até {data_de_hoje}, (Lembrando que as metas individuais não estão atualizadas), att Murilo MoraeS"
corpo_do_email = mail.body # escrita do corpo do email.
print(f"4 - o corpo do email é: {corpo_do_email}")
print("--------")


print("--------")
lista_emails = [
    "cristiane@greentech.log.br",
    "rafael.gomes@greentech.log.br",
    "paulo.chequetti@greentech.log.br",
    "ariel.rosenblatt@greentech.log.br",
    "felipe.andriolo@greentech.log.br",
    "marcelo.mota@greentech.log.br",
]
mail.To = ";".join(lista_emails) #lista de destinatarios do email
destinatarios = mail.to 
print(f"5 - os destinatarios dos email são: {destinatarios}")
print("--------")












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