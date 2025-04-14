# Importar as bibliotecas necessárias
import openpyxl
from zenpy import Zenpy
from zenpy.lib.api_objects import CustomField

# Fazer login no Zendesk
from zenpy.lib.api_objects import Ticket, User
creds = {
    'email': 'hellen.araujo@conexasaude.com.br',
    'token': 'rYjykv69LkLxcsLDOVpPkdHPCGZXLaG6NziGQ6rQ',
    'subdomain': 'conexasaude3465'
}

zenpy_client = Zenpy(**creds)
 
# Abrir o arquivo Excel
workbook = openpyxl.load_workbook('nps-zendesk-v1.xlsx')
worksheet = workbook['nps']
 
# Percorrer as linhas da planilha
for row in worksheet.iter_rows(min_row=2):
    # Obter os valores das células
    requester_name = row[0].value
    requester_email = row[1].value
    title = row[2].value
    description = row[3].value
    
    # Criar um ticket no Zendesk
    ticket_audit = zenpy_client.tickets.create(
        Ticket(
            subject="Teste - npsZendesk",
            description="Teste - npsZendesk",
            requester=
                User(
                    name="Hellen Araujo",
                    email="hellenaraujo703@gmail.com"
                ),
            custom_fields=[
                #CustomField(
                    #id=, value=)
            ]
        )
    )
 
    # Imprimir o número do ticket criado
    print(f'Ticket {ticket_audit.id} criado.')
    # Inserir ID do ticket na planilha.
 