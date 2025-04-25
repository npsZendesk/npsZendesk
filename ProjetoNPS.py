# Importar as bibliotecas necessárias
import openpyxl
from zenpy import Zenpy
# CustomField é uma função para preencher campos de tickets do Zendesk
from zenpy.lib.api_objects import CustomField
from zenpy.lib.api_objects import Ticket, User
 
# Fazer login no Zendesk
creds = {
    'email': 'hellen.araujo@conexasaude.com.br',
    'token': 'rYjykv69LkLxcsLDOVpPkdHPCGZXLaG6NziGQ6rQ',
    'subdomain': 'conexasaude3465'
}
 
zenpy_client = Zenpy(**creds)
 
# Abrir o arquivo Excel
workbook = openpyxl.load_workbook('Teste Projeto NPS.xlsx')
worksheet = workbook['Planilha1']
 
# Percorrer as linhas da planilha
for row in worksheet.iter_rows(min_row=2):
    # Obter os valores das células
    requester_name = row[0].value
    requester_email = row[1].value
    title = row[2].value
    description = row[3].value
    user_type = row[4].value
   
    # Criar um ticket no Zendesk
    ticket_audit = zenpy_client.tickets.create(
        Ticket(
            subject= title,
            description= description,
            requester=
                User(
                    name= requester_name,
                    email= requester_email
                ),
            custom_fields=[
                CustomField(
                    id=31483088708887, value= user_type)
            ]
        )
    )
 
    # Imprimir o número do ticket criado
    # print({ticket_audit})
    # Inserir ID do ticket na planilha.