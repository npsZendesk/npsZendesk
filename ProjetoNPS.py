# Importar as bibliotecas necessárias
import openpyxl
from zenpy import Zenpy
# CustomField é uma função para preencher campos de tickets do Zendesk
from zenpy.lib.api_objects import CustomField
from zenpy.lib.api_objects import Ticket, User
 
# Fazer login no Zendesk
creds = {
    'email': 'cto@conexasaude.com.br',
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
    user_catgr = row[4].value
    user_type = row[5].value
    clinic = row[6].value
    note = row[7].value
    login = row[8].value
    comentary = row[9].value
    contact_reason = row[10].value
    bu = row[11].value
    cs = row[12].value
   
    # Criar um ticket no Zendesk
    ticket_audit = zenpy_client.tickets.create(
    # Cria e preenche o Assunto, Descrição e Solicitante do ticket
        Ticket(
            subject= title,
            description= description,
            requester= User(
                    name= requester_name,
                    email= requester_email
                ),
        # Preenche os campos do formulário do ticket
            custom_fields=[
                CustomField(
                    id=31483088708887, value= user_catgr),
                CustomField(
                    id=1500004336442, value= user_type),
                CustomField(
                    id=31782165825303, value= clinic),
                CustomField(
                    id=31723268674583, value= note),
                CustomField(
                    id=31723260609047, value= login),
                CustomField(
                    id=31782250293527, value= comentary),
                CustomField(
                    id=31782338291095, value= contact_reason),
                CustomField(
                    id=31967660376215, value= bu),
                CustomField(
                    id=31967742138007, value= cs)
            ]
        )
    )
 
    # Imprimir o número do ticket criado
    # print(ticket_audit.text)
    # Inserir ID do ticket na planilha.