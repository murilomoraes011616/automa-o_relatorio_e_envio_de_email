import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application")
print("1 - conectado ao outlook")

mail = outlook.CreateItem(0)
print("2 - mail criado")

print(type(mail))
print(mail.Subject)

print("3"), 