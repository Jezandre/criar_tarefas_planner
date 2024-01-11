# Utilizar a ferramenta Graph para criar tarefas no Microsoft Planner atrvés de Python e SQL
 ## Objetivo
 Automatizar processos é um ponto importante para a garantia da qualidade e aumento de produtividade. O objetivo deste artigo é ajudar a como utilizar informações de um banco de dados para criar tarefas automaticamente no Microsoft Planner.

 ## Introdução

 Hoje em dia com a tecnologia é quase impossível trabalhar com a qualidade de processos sem pensar em automatização. A automação permite garantir que um processo seja executado de maneira uniforme e ágil, garantindo assim melhor performance e evitando erros. 
 
 Em uma empresa normalmente todos os processos são armazenados em sistemas ERP, o que ajuda muito na gestão de todo processo produtivo. A utilização de ferramentas que permitem uma gestão ágil aliados, garantem que tarefas específicas de um processo sejam realizadas dentro do prazo e de maneira coeza.

 O MS Planner, é uma ferramenta da microsoft que possuí esta possibilidade de gestão de tarefas. Aliando esta ferramenta com o banco de dados de um ERP, é possível estabelecer a comunicação e a criação de tarefas de maneira automática, diminuindo assim o tempo de resposta para atuação em um processo.

 Ao longo do desenvolvimento desta aplicação, encontrei pouco material que mostrasse como isso poderia ser feito. Por isso achei válido escrever este artigo sobre o tema. Tem muito conteúdo sobre como obter as tarefas através da API da microsoft para criar dashboards de gerenciamento. Esses conteúdos geralmente utilizam GET. 

 A idéia é basicamente utiizar as informações registradas em um banco de dados para criação e atualização de tarefas automaticamente no MS Planner utilizando a ferramenta GRAPH com Python. Explorando assim as opções GET, PATCH e POST da API. 

 Para realizar esse processo você precisará ter acesso ás credenciais de Client_id, Secret_ID e Tenant_ID API Graph da Microsoft. Além disso você precisará criar uma aplicação do GRAPH na Azure e ter acesso ao banco de dados ou a API do sistema que você utilizará para criar ou atualizar as tarefas. Outro detalhe é atribuir a permissão **Group.ReadWrite.All** no portal da Azure, para que as opções de PATCH e POST possam ser utilizadas

 ## Requisitos

 As ferramentas que utilizei foram
 
 * Python 3.10.11
 * Bibliotecas
  * msal
  * requests
  * json
  * datetime
  * configparser
  * pyodbc
  * pandas
  * datetime
 * Uma conta administradora no portal Azure
 * Credenciais da API Microsoft Graph
 * Servidor Linux UBUNTU
 * Banco de dados

## Credencias:

O primeiro passo é criar um arquivo ini onde possa inserir as credenciais de acessos tanto ao banco de dados como a API do Graph
para importar o arquivos basta

```python
config = configparser.ConfigParser()

config = configparser.ConfigParser()
config.read('infoDB.ini')

BD_USER = str(config['SQL_SERVER']['BD_USER'])
BD_PASS = str(config['SQL_SERVER']['BD_PASS'])
BD_HOST = str(config['SQL_SERVER']['BD_HOST'])
BD_BD   = str(config['SQL_SERVER']['BD_BD'])
ID_CLIENT = str(config['MICROSOFT']['CLIENT_ID'])
SECRET_CLIENT = str(config['MICROSOFT']['CLIENT_SECRET'])
AUTHORITY = str(config['MICROSOFT']['AUTHORITY'])

```
 
## Conectar a a API

Em seguida é necessário criar uma conexão com a API utilizando as credenciais fornecidas. Para isso iremos criar a função conecta API. No site da [Microsoft](https://developer.microsoft.com/en-us/graph/graph-explorer)  você pode conferir como obter as informações para criar esta conexão. Para conectar na API voce precisará criar a aplicação dentro do portal azure, e a partir daí obter o client_id, secrete_id e o tenant_id. Com essas informações através do código abaixo você obterá o headers que será importante para os próximos passos.

```python
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
```
## Conectar ao banco de dados

Abaixo o código para a criação da string de conexão ao banco de dados. No caso o banco utilizado foi SQLserver então as configurações tanto para windows quanto para Linux ubuntu são as predeterminadas

```python
def conexao():

    #Criar a string de conexão com o sql server Windows
    # conn_str = 'DRIVER={SQL Server};SERVER=' + BD_HOST + ';DATABASE=' + BD_BD + ';UID=' + BD_USER + ';PWD=' + BD_PASS + ';Trusted_Connection=no;'
    
    # Criar a string de conexão com o sql server Ubuntu
    conn_str = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + BD_HOST + ';DATABASE=' + BD_BD + ';UID=' + BD_USER + ';PWD=' + BD_PASS + ';Trusted_Connection=no;'
    
    #Conexao com o banco de dados
    cnxn = pyodbc.connect(conn_str)
    cursor = cnxn.cursor()
    return cnxn, cursor

```

## Criar e executar Querys
A função abaixo é para que possamos executar querys no banco de dados e atribuir a um dataframe do pandas, assim fica bem mais fácil de obter as informações que precisamos para criar as tarefas. Lembrando que o usuário que você estiver utilizando precisa ter as permissões necessárias para executar os comandos.

```python
#Função para executar a query no banco de dados - Geral
def executaQuery(querySQL):

    #Função de executor
    cnxn, cursor = conexao()
    resultado = pd.read_sql(querySQL, cnxn)
    print(resultado)

    return resultado
```
Você pode definir como query principal a query que trará todas as informações necessárias para sua tarefa, como descrição, atribuições, lista de checkagem, data de inicio, data de fim. Lembrando que, para fazer as atribuições é preciso relacionar a base de dados do sistema erp com os usuarios microsoft. Sugiro que estas informações sejam parametrizadas. No caso como utilizamos office365, fiz isso utilizando o e-mail dos usuários que é igual ao e-mail do sistema ERP.

```python
def queryPrincipal():

    #Query que será utilizada para atribuir os dados na tarefa
    querySqlMain = """SELECT * FROM <<Tabela>> WHERE TASK_CRIADA <> 1"""
    resultado = executaQuery(querySqlMain)
    return resultado


```
## Controlar tarefas criadas
Para que o programa rode e não crie milhares de tarefas é preciso apontar de alguma maneira que a tarefa foi criada. Para isso utilizei a query a seguir, de maneira que ele atribua tanto id da task criada no planner, quanto que ela informe que a tarefa foi criada. A consulta principal precisa ter essa informação.

```python
#Função para atualizar as informações no banco de dados de maneira que não tenha tarefa criada de maneira duplicada - Modificar
def querySqlUpdate(pendencia_id, task_id):

    #Query para atualizar a tabela de verificação se a tarefa já existe
    querySqlUpdate = f"""UPDATE <<Tabela>>
                        SET TASK_CRIADA = 1,
                        NK_TASK_PLANNER = '{task_id}'
                        WHERE id = {id};"""

    #Executores da Query
    cnxn, cursor = conexao()
    cursor.execute(querySqlUpdate)
    cursor.commit()
```

## Criar uma tarefa

Para criar uma tarefa no planner você precisará da autorização do administrador Group.ReadWrite.All ao Microsft Graph.
Nesse [site](https://learn.microsoft.com/pt-br/graph/api/resources/plannertask?view=graph-rest-1.0) você encontra toda a documentação referente a como construir um arquivo json para a criação da tarefa. 

As opções que eu necessitava utilizei da maneira descrita no código. Utilizei de variáveis em alguns campos para inserir algumas lógicas nos detalhes das tarefas. 

O return é proposital, pois pretendo inserir no banco de dados a id da tarefa para que ela possa ser atualizada conforme as informações do banco de dados.

```python
#Função para criar tarefa - Geral
def criarTarefa(assignments, plan_id, bucket_id, dataVencimento, dataInicio, titulo, descricao, prioridade, percentualCompleto, checklist):
    
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
            "category4": True,
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

```

## Atualizar tarefa
Aqui é importante ter o taskid salvo em uma tabela que se relacione com a tabela da query principal, pois é através destas tabelas que controlaremos as atualizações e a quantidade de tarefas criadas.

A atualização da tarefa ocorrerá sempre que algum atributo for alterado. No caso defini apenas o atributo de andamento. 

Como estamos conectando um banco de dados e não diretamente a API do sistema ERP, as alterações que vem do ERP é que vão interferir na tarefa e as alterações da tarefa não interferirão no banco do ERP. Então coloquei apenas o andamento. Dessa forma se a demanda tiver concluida no ERP a tarefa será concluida junto. Claro que se for necessário atualizar outros campos como por exemplo, bucket, plano, datas é possível fazer isso por aqui. Basta configurar o json de acordo com as necessidades.

Observe que para atualizar a tarefa o headers é diferente. Isso porque a API analisa se a tarefa existe ou não antes de executar o comando. Isso é interessante, pois, se a tarefa não existir nada vai acontecer.

```python
#Função para atualizar tarefa - OK
def atualizarTarefa(bucket_id):

    #Query que lista as tarefas criadas no planner
    querySqlUpdate = """<<Utilize sua query aqui>>
                    """
    resultado = executaQuery(querySqlUpdate)

    #Iteração para verificar as alterações realizadas no planner
    for i, linha in resultado.iterrows():
        
        #Dados para atualização
        task_id = str(linha["task_id"])+'?'
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
            "percentComplete": andamento
        }

        #Executor da atualização
        response_patch = requests.patch(base_url_update, headers=headers_update, json=task_data_update)
    print("Todas as tarefas foram atualizadas")
```

## Definir variávels e função main()

Cada campo do json pode ser utilizada um variável para definir atribuições, checklist, label e por aí vai. Abaixo vou descrever algumas que utilizei. É claro, a utilização de cada de cada variável dependerá da aplicação.
Você pode criar uma função com uma lógica específica para cada campo, mas acredito que as que irei informar abaixo já devem vão produzir alguns insights.

Primeiramente utilizei a função queryPrincipal para obter o dataframe em que pretendo iterar e fazer as atualizações. Em seguida eu avalio se o dataframe está vazio, se ele estiver vazio é pq não existem tarefas a serem criadas. Caso contrário ele vai iniciar a iteração na tarefa no qual vai caminhar linha por linha e criar tarefa para cada identificação com as informações contidas na query principal.

```python
      
    # Executar query do for
    resultado = queryPrincipal()    

    # executaQuery(querySQL)
    if resultado.empty:
        print("Não há tarefas para serem criadas!")
    else:
        for indice, linha in resultado.iterrows():        
```
Como disse, as variáveis dependem muito de como você pretende criar suas tarefas e quais informações você tem disponíveis para isso. Abaixo defini algumas importantes para o preenchimento da tarefa em questão. Nessa tarefa em especifico, eu preciso do usuário que será atribuido, atribuidor que é quem designa a tarefa e numero de identificação. Todas essas informações vieram da query principal sem a necessidade de criar uma função para elas 

```python
            #Usuário Ofice365
            usuario_id = linha["USERID"]
            atribuidor_id = linha["NK_ATRIBUIDOR_NC"]

            #id da NC
            nc_id = linha["IDNAOCONFORMIDADE"]
```
Para o andamento eu utilizei uma função que através do status vinda do banco de dados ela define como está o andamento da tarefa.

```python
         def andamentoPendencia(status):
             #Determinar o andamento da pendencia
             if status == "Aberto":
                 percentualCompleto = 0
             elif status == "Em andamento":
                 percentualCompleto = 1
             else:
                 percentualCompleto = 100
             return percentualCompleto

            #Status da pendencia no Fluxis
            percentualCompleto = andamentoPendencia(linha["NOMESTATUS"])            
```
A priotridade também pode se utilizar uma função baseada em algum campo da query principal, mas nesse caso não utilizei pois todas as tarefas no caso possuem o mesmo nível de prioridade.

```python
            #Prioridade do Fluxis
            prioridade = 5
```
As datas no MS planner precisam estar em datastamp. Deixei já determinado o prazo somado, essa informação depende muito do negócio. Como são tarefas protocolares em geral o prazo é fixo.

```python
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

            # Data de registro
            dataRegistro = linha["DATAABERTURA"]
            # datalimite =  dataRegistro+timedelta(days=4)

            # Somar 1 dia útil          
            dataVencimento = (somarDiasUteis(dataRegistro))

            # Data de registro da pendencia
            dataInicio = str(converterData(dataRegistro))
```
As variáveis abaixo serão que escreverão na nossa tarefas os detalhes que precisamos. 

```python

            #Título da tarefa
            titulo = f'NC {nc_id} - {linha["SETOR"]}'

            #Descrição da tarefa
            descricao = f'ID Pendência: {nc_id}\n\nResponsável: {linha["RESPONSAVEL"]}\n\n{linha["TITULO"]}\n\n{linha["DESCRICAONAOCONFORMIDADE"]}\n\n{linha["DESCRICAOSIGNIFICANCIA"]}'
            
    
```

Notem que o checklist eu passei uma lista predeterminas com os passos da tarefa que podem ser alterados conforme a utilização.

```python

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

           #Lista de ítens do checklist da tarefa, descever aqui o passo a passo
            lista =["Atividade 01",
                    "Atividade 02",
                    "Atividade 03",
                    "Atividade 04",
                    "Atividade 05"]

            
            #Cria o dicionario da lista de checkagem
            checklist = listaDeChecagem(lista)
```
Abaixo precisamos criar a string que irá atribuir os usuários na tarefa. A função é baseada em uma query que retorna os usuários que serão atribuídos, então é preciso avaliar se existem os usuários antes de atribuir.

```python

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
            # #Cria o dicionario de usuários atribuídos
            assignments = listaUsuarios(usuario_id, atribuidor_id)
```

Aqui podemos utilizar a lógica da metodologia ágil, que conforme o andamento do processo no ERP a tarefa anda pelos buckets pré-setados. Você pode criar uma função que atualize a tarefa conforme status ou andamento da tarefa. Além disso é possível atribuir um label a tarefa conforme o andamento. Um coisa importante, sempre identifique as tarefas que são criadas automáticamente, isso ajudará muito na identificação da tarefa.

```python
            # #Escolher plano e bucklet conforme o setor da pendencia
            nomeSatus = linha["NOMESTATUS"]
            idSignificancia = linha["IDSIGNIFICANCIANAOCONFORMIDADE"]
            def bucketSetor(nomeStatus, impacto):
            
                #Planner pré-definido RENP
                plan_id = "<<DEFINA O PLANNER AQUI>>"
            
                
                if nomeStatus == 'Aberto' and impacto is None:
                    bucket_id = '<<DEFINA O BUCKET>>'
                    label_id =  <<Nº DO LABEL(PROCURE NA DOCUMENTAÇÃO)>>                
                elif nomeStatus == 'Aberto':
                    bucket_id = '<<DEFINA O BUCKET>>'
                    label_id =  <<Nº DO LABEL(PROCURE NA DOCUMENTAÇÃO)>>               
                else:
                     bucket_id = '<<DEFINA O BUCKET>>'
                     label_id =  <<Nº DO LABEL(PROCURE NA DOCUMENTAÇÃO)>>  
                return plan_id, bucket_id, label_id
            plan_id, bucket_id, label_id = bucketSetor(nomeSatus, idSignificancia)
```
Por fim as funções que irão criar e atualizar a tarefa e a função que vai atualizar no banco de dados as informções para que a tarefa não seja criada várias vezes.

```python

            #Cria a tarefa conforme valores definidos na main
            task_id = criarTarefa(assignments, plan_id, bucket_id, label_id, dataVencimento, dataInicio, titulo, descricao, prioridade, percentualCompleto, checklist)

            # #Atualiza dados no banco, para não haver criação de tarefas duplicadas
            update_tabela = querySqlUpdate(nc_id, str(task_id))

    #Atualiza status da tarefa conforme o Fluxis
    atualizarTarefa()

```

 ## Execução

 O código foi implementado em um servidor linux e roda utilizando um arquivo shell de 10 em 10 minutos no crontab

 ## Conclusão
O processo trata se de uma otimização de tempo dos processos de qualidade e exige conhecimentos amplos, como SQL, Python, API, GRAPH, MSPlanner e muita curiosidade. Tal processo pode ser aplicado para otimizar tarefas de Kanban da empresa e garantir que o andamento das tarefas seguem o andamento da qualidade. A automação de processos é importante para a diminuição de erros e padronização de resultados. Se achou interessante me procure pelo linkedin estarei disposto a ajudá-lo.
 
