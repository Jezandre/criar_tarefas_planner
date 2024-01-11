from msal import ConfidentialClientApplication
import requests
import json
import datetime
import configparser
import pyodbc
import pandas as pd
from datetime import datetime, timedelta

config = configparser.ConfigParser()

# Variavel que le o arquivo de configuracao
config = configparser.ConfigParser()
config.read('<<Caminho do arquivo .ini>>')

BD_USER = str(config['SQL_SERVER']['BD_USER'])
BD_PASS = str(config['SQL_SERVER']['BD_PASS'])
BD_HOST = str(config['SQL_SERVER']['BD_HOST'])
BD_BD   = str(config['SQL_SERVER']['BD_BD'])
ID_CLIENT = str(config['MICROSOFT']['CLIENT_ID'])
SECRET_CLIENT = str(config['MICROSOFT']['CLIENT_SECRET'])
AUTHORITY = str(config['MICROSOFT']['AUTHORITY'])

#Função para conectar na Api do graf - Geral
def conectaApi():
    #Seu ID e segredo do cliente Microsoft 365
    client_id = ID_CLIENT
    client_secret = SECRET_CLIENT

    #Parâmetros de autenticação
    authority = AUTHORITY
    scopes = ['https://graph.microsoft.com/.default']

    #Criar uma aplicação de cliente confidencial
    app = ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret
    )

    # Obter um token
    token_response = app.acquire_token_for_client(scopes=scopes)
    access_token = token_response['access_token']
    
    #Verificar se o Token está funcionando
    if 'access_token' in token_response:
        access_token = token_response['access_token']
    else:
        print("Não foi possível obter o token de acesso.")
    
    #Criar a variável Headers para autenticação
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json'
    }

    return headers

#Função para criar tarefa - Geral
def criarTarefa(assignments, plan_id, bucket_id, label_id, dataVencimento, dataInicio, titulo, descricao, prioridade, percentualCompleto, checklist):
    
    #Url da API para criação de tarefas
    base_url = 'https://graph.microsoft.com/v1.0/planner/tasks/'

    #Parametros json para inserir dados em uma tarefa
    task_data = {
        "planId": plan_id,
        "bucketId": bucket_id,
        "dueDateTime": dataVencimento,
        "startDateTime": dataInicio,
        "title": titulo,
        "priority": prioridade,
        "percentComplete": percentualCompleto,
        "appliedCategories": {"@odata.type": "microsoft.graph.plannerAppliedCategories",
            "category"+str(label_id): True,
            "category25": True
            },
        "status": 5,
        "assignments": assignments,
        "details": {
            "description": descricao,
            "previewType": "description",
            "checklist": checklist
        }
    }
    
    #Variável para a criação da tarefa
    response_post = requests.post(base_url, headers=conectaApi(), json=task_data)

    #Variável para obter a id da tarefa criada para atribuir em outra etapa
    task_id  = response_post.json()["id"]    

    # print(response_post.json())
    print("Tarefa criada com sucesso")

    return task_id

#Função para atualizar tarefa - OK
def atualizarTarefa():

    #Query que lista as tarefas criadas no planner
    querySqlUpdate = """SELECT
                        <<COLUNAS>>
                    FROM
                        <<TABELA>>
                    WHERE
                        <<COLUNA>> <> 'None'
                    """
    resultado = executaQuery(querySqlUpdate)
    

    #Iteração para verificar as alterações realizadas no planner
    for i, linha in resultado.iterrows():

        #Atualizar o bucket conforme o status da Não conformidade

        plan_id, bucket_id, label_id = bucketSetor(linha["NOMESTATUS"], linha["IDSIGNIFICANCIANAOCONFORMIDADE"])
        
        #Dados para atualização
        task_id = str(linha["NK_TASK_PLANNER"])+'?'
        andamento = andamentoPendencia(str(linha["NOMESTATUS"]))

        #Url base para atualizar a tarefa criada
        base_url_update = f'https://graph.microsoft.com/v1.0/planner/tasks/{task_id}/details'

        #Pegar o Headers
        headers_update = conectaApi()

        #Pegar mais informações do headers
        response_get = requests.get(base_url_update, headers=headers_update)

        #Pegar etag para avaliar se uma task existe
        etag = response_get.headers.get('ETag', '')
        headers_update['If-Match'] = etag

        #json para atualizar as informações da tarefa
        task_data_update = {
            "bucketId": bucket_id,
            "percentComplete": andamento,
            "appliedCategories": {"@odata.type": "microsoft.graph.plannerAppliedCategories",
                "category"+str(label_id): True,
                "category25": True
            },
        }

        #Executor da atualização
        response_patch = requests.patch(base_url_update, headers=headers_update, json=task_data_update)
    print("Todas as tarefas foram atualizadas")
        
#Função para conectar no banco de dados - Geral
def conexao():

    #Criar a string de conexão com o sql server Windows
    # conn_str = 'DRIVER={SQL Server};SERVER=' + BD_HOST + ';DATABASE=' + BD_BD + ';UID=' + BD_USER + ';PWD=' + BD_PASS + ';Trusted_Connection=no;'
    
    # Criar a string de conexão com o sql server Ubuntu
    conn_str = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + BD_HOST + ';DATABASE=' + BD_BD + ';UID=' + BD_USER + ';PWD=' + BD_PASS + ';Trusted_Connection=no;'
    
    #Conexao com o banco de dados
    cnxn = pyodbc.connect(conn_str)
    cursor = cnxn.cursor()
    return cnxn, cursor

#Função para executar a query no banco de dados - Geral
def executaQuery(querySQL):

    #Função de executor
    cnxn, cursor = conexao()
    resultado = pd.read_sql(querySQL, cnxn)
    print(resultado)

    return resultado

#Função para converter data em timestamp - Geral
def converterData(data_converte):    

    #Função para converter os dados em timestamp
    data_formatada = data_converte.strftime('%Y-%m-%dT%H:%M:%S.%fZ')
    
    return data_formatada

#Função para criar o checklist da tarefa - Modificar
def listaDeChecagem(listaDeChecagem):
    
    #Criar dicionario dos checklists
    checklist = {}

    #Função se para determinar quantos checklists serão marcados no inicio da tarefa
    for i, valor in enumerate(listaDeChecagem):        
            item = {
                    "@odata.type": "microsoft.graph.plannerChecklistItem",
                    "title": valor,
                    "isChecked": False,
                }
            checklist[i] = item        

    return checklist

#Função para definir os usuários atribuidos na tarefa - main
def listaUsuarios(usuario_id, atribuidor_id):
    assignments = {}
    coordenador_nuq = queryCoordenadoresNUQ()
    if usuario_id != None:
        item = {
            "@odata.type": "#microsoft.graph.plannerAssignment",
            "orderHint": " !",
        }
        assignments[usuario_id] = item
        item = {
            "@odata.type": "#microsoft.graph.plannerAssignment",
            "orderHint": " !",
        }
        assignments[atribuidor_id] = item
        if coordenador_nuq.empty:
            pass
        else:
            for i, linha in coordenador_nuq.iterrows():
                item = {
                "@odata.type": "#microsoft.graph.plannerAssignment",
                "orderHint": " !",
                }
                assignments[linha['userId']] = item
            print(assignments)
    else:
        if coordenador_nuq.empty:
            pass
        else:
            for i, linha in coordenador_nuq.iterrows():
                item = {
                "@odata.type": "#microsoft.graph.plannerAssignment",
                "orderHint": " !",
                }
                assignments[linha['userId']] = item
        print(assignments)
    
    return assignments        

#Função para direcionar os buckets que serão inseridas as tarefas - OK
def bucketSetor(nomeStatus, impacto):

    #Planner pré-definido RENP
    plan_id = "<<Nome planner>>"
    #Buckets serão selecionados conforme as informações no Fluxis
    # bucket_id = 'nzzvxaY7QE6Ss9gGsxUBy2UAMy5c'


    if nomeStatus == 'Aberto' and impacto is None:
        bucket_id = '<<Nome bucket>>'
        label_id =  7
   
    elif nomeStatus == 'Aberto':
        bucket_id = '<<Nome bucket>>'
        label_id =  4    
    else:
         bucket_id = '<<Nome bucket>>'
         label_id =  1
    return plan_id, bucket_id, label_id

#Função da para determinar a query principal - OK
def queryPrincipal():

    #Query que será utilizada para atribuir os dados na tarefa
    querySqlMain = """SELECT * FROM <<TABELA>> WHERE <<COLUNA>> <> 1"""
    resultado = executaQuery(querySqlMain)
    return resultado

#Função para definir os coordenadores que serão atribuidos nas tarefas - main
def queryCoordenadoresNUQ():

    #Query utilizada para determinar os líderes que serão atribuidos na tarefa no inicio da criação
    querySqlMain = """SELECT * FROM <<TABELA>>"""
    resultado = executaQuery(querySqlMain)
    return resultado

#Função para atualizar as informações no banco de dados de maneira que não tenha tarefa criada de maneira duplicada - Modificar
def querySqlUpdate(pendencia_id, task_id):

    #Query para atualizar a tabela de verificação se a tarefa já existe
    querySqlUpdate = f"""UPDATE <<TABELA>>
                        SET <<COLUNA>> = 1,
                        <<COLUNA_TASK>> = '{task_id}'
                        WHERE <<COLUNA_IDENTIFICACAO>> = {pendencia_id};"""

    #Executores da Query
    cnxn, cursor = conexao()
    cursor.execute(querySqlUpdate)
    cursor.commit()

#Função para definir o andamento da tarefa - main
def andamentoPendencia(status):
    #Determinar o andamento da pendencia
    if status == "Aberto":
        percentualCompleto = 0
    elif status == "Em andamento":
        percentualCompleto = 1
    else:
        percentualCompleto = 100
    return percentualCompleto

#Função para determinar a prioridade da tarefa - main
def prioridadePendencia(prioridadeFluxis):
    #Determinar a prioridade da pendencia
    if prioridadeFluxis == "Média":            
        prioridade = 2
    elif prioridadeFluxis == "Alta":
        prioridade = 1
    else:
        prioridadeFluxis = 5
    return prioridade

# Função para somar apenas dias úteis à data de criação
def somarDiasUteis(diaInput):
    if diaInput.weekday() == 4:
        dataVencimento = converterData(diaInput+timedelta(days=3))
    elif diaInput.weekday() == 5:
        dataVencimento = converterData(diaInput+timedelta(days=2))
    elif diaInput.weekday() == 6:
        dataVencimento = converterData(diaInput+timedelta(days=1))
    else:
         dataVencimento = converterData(diaInput+timedelta(days=1))
    return dataVencimento

#Função para realizar todo o processo - Main
def main():
        
    # plan_id = "ev4x0-qiRUSSK0OOsyf08mUAChwX"
    # bucket_id = "CWVemErT0U-nocIfHWpgg2UANFZI"

    # Executar query do for
    resultado = queryPrincipal()    

    # executaQuery(querySQL)
    if resultado.empty:
        print("Não há tarefas para serem criadas!")
    else:
        for indice, linha in resultado.iterrows():        

            #Usuário Ofice365
            usuario_id = linha["<<COLUNA_USERID>>"]
            atribuidor_id = linha["<<COLUNA_ATRIBUIDOR_NC>>"]

            #id da NC
            nc_id = linha["<<COLUNA_ID_TAREFA>>"]

            #Status da pendencia no Fluxis
            percentualCompleto = andamentoPendencia(linha["<<COLUNA_NOMESTATUS>>"])
            

            #Prioridade do Fluxis
            prioridade = 5

            # Data de registro
            dataRegistro = linha["<<COLUNA_DATAABERTURA>>"]
            # datalimite =  dataRegistro+timedelta(days=4)

            # Somar 1 dia útil          
            dataVencimento = (somarDiasUteis(dataRegistro))

            # Data de registro da pendencia
            dataInicio = str(converterData(dataRegistro))                    

            #Título da tarefa
            titulo = f'NC {nc_id} - {linha["<<COLUNA_SETOR>>"]}'

            #Descrição da tarefa
            descricao = f'ID: {nc_id}\n\nResponsável: {linha["<<COLUNA_RESPONSAVEL>>"]}\n\n{linha["<<COLUNA_TITULO>>"]}\n\n{linha["<<COLUNA_DESCRICAO>>"]}\n\n{linha["<<COLUNA_DESCRICAOSIGNIFICANCIA>>"]}'
            
            #Lista de ítens do checklist da tarefa, descever aqui o passo a passo
            lista =["Atividade 01",
                    "Atividade 02",
                    "Atividade 03",
                    "Atividade 04",
                    "Atividade 05",
                    ]
            
            #Cria o dicionario da lista de checkagem
            checklist = listaDeChecagem(lista)

            # #Cria o dicionario de usuários atribuídos
            assignments = listaUsuarios(usuario_id, atribuidor_id)

            # #Escolher plano e bucklet conforme o setor da pendencia
            nomeSatus = linha["NOMESTATUS"]
            idSignificancia = linha["IDSIGNIFICANCIANAOCONFORMIDADE"]
            plan_id, bucket_id, label_id = bucketSetor(nomeSatus, idSignificancia)

            #Cria a tarefa conforme valores definidos na main
            task_id = criarTarefa(assignments, plan_id, bucket_id, label_id, dataVencimento, dataInicio, titulo, descricao, prioridade, percentualCompleto, checklist)

            # #Atualiza dados no banco, para não haver criação de tarefas duplicadas
            update_tabela = querySqlUpdate(nc_id, str(task_id))

    #Atualiza status da tarefa conforme o Fluxis
    atualizarTarefa()

if __name__ == "__main__":
    main()
