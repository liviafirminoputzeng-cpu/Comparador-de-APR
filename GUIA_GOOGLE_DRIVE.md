# Ligar o acompanhamento de BowTies à pasta da PUTZ

Pasta de destino: https://drive.google.com/drive/folders/1ichy39_gjGtfzALO66pe275YxY-HdRSh?usp=drive_link

O ID dessa pasta já está em `drive_acompanhamento.py`. O botão cria um `.xlsx`
por projeto e par de APRs. Na pasta, os analistas com permissão de **Editor**
abrem o mesmo arquivo para atualizar o acompanhamento online. O comparador de
APRs continua independente, inclusive para visitantes sem login.

## 1. Descubra o tipo da pasta

Abra o link com a conta da PUTZ e veja o caminho mostrado no Google Drive:

- **Drives compartilhados**: use a opção A, conta de serviço.
- **Meu Drive** ou **Compartilhados comigo**: use a opção B, autorização de uma
  conta humana da PUTZ que seja proprietária ou editora da pasta.

Confirme também que os dois analistas e o coordenador conseguem **Editar** a
pasta. O site não muda permissões de compartilhamento automaticamente.

## 2. Prepare o Google Cloud

Peça ajuda ao administrador do Google Workspace da PUTZ se a equipe não puder
criar projetos ou clientes OAuth. Em https://console.cloud.google.com/ selecione
ou crie um projeto e ative a **Google Drive API** em *APIs e serviços > Biblioteca*.
Em **Google Auth Platform**, configure a tela de consentimento para usuários
**Internos**, se o domínio da PUTZ permitir; em *Testing* só os usuários de teste
conseguem entrar. Crie uma credencial OAuth **Aplicativo da Web** para o login
no Streamlit. Em *URIs de redirecionamento autorizados* inclua exatamente:

`https://comparador-de-apr-bkcbgj6tuz9ggkqmavy7yb.streamlit.app/oauth2callback`

Copie o **client ID** e o **client secret** desse cliente para o bloco `[auth]`
dos Secrets. Esse login identifica cada pessoa; ele **não** concede ao site
acesso de gravação ao Drive. Escolha A ou B abaixo para isso.

### Opção A: pasta em Drives compartilhados

1. Em *IAM e administrador > Contas de serviço*, crie uma conta de serviço no
   mesmo projeto e gere uma chave **JSON**. Guarde a chave fora do GitHub.
2. No Google Drive, adicione o `client_email` da chave como membro do Drive
   compartilhado com permissão **Colaborador** (ou função que permita adicionar
   arquivos); se o administrador usar acesso por pasta, confirme que a conta
   consegue criar arquivos na pasta exata do link.
3. No Streamlit, copie todos os campos do JSON para `[drive_service_account]`
   dos Secrets. Em TOML, preserve as quebras de linha da chave privada usando
   `\n` dentro de aspas duplas; confira o exemplo abaixo.

Uma conta de serviço não possui cota própria de armazenamento no Meu Drive;
por isso, use esta opção apenas em **Drives compartilhados**.

### Opção B: pasta em Meu Drive

1. Em Google Cloud, crie **outro cliente OAuth**, do tipo **Aplicativo para
   computador**. Use uma conta da PUTZ que tenha permissão **Editor** na pasta.
2. Baixe o JSON desse cliente para o computador do administrador e renomeie
   para `credenciais_drive.json` dentro da pasta do programa. Esse arquivo está
   no `.gitignore` e **nunca deve ir ao GitHub**.
3. No terminal do VS Code, com o ambiente virtual ativo, execute
   `python autorizar_drive.py`. Uma janela do Google abrirá para a conta da PUTZ.
   Autorize o escopo da API Drive e copie o **refresh token** exibido no terminal.
4. Coloque o `client_id`, o `client_secret` desse cliente **Desktop** e o token
   no bloco `[drive_oauth]` dos Secrets. O token dá acesso aos arquivos da conta:
   trate-o como senha. Revogue-o no Google se houver exposição.

Algumas organizações precisam aprovar o escopo Drive no painel do administrador.
O cliente OAuth de login do site e o cliente Desktop são credenciais diferentes.

## 3. Configure os Secrets no Streamlit Community Cloud

Abra o aplicativo publicado > **Manage app** > **Settings** > **Secrets**.
Insira o exemplo a seguir com valores reais. Os e-mails são os usuários que
podem entrar na área; os criadores também podem gerar novas planilhas. Coloque
essas duas listas **antes** de qualquer bloco `[seção]`, como no exemplo.

```toml
usuarios_equipe = ["analista1@putz.com.br", "analista2@putz.com.br"]
usuarios_criadores = ["coordenador@putz.com.br"]

[auth]
redirect_uri = "https://comparador-de-apr-bkcbgj6tuz9ggkqmavy7yb.streamlit.app/oauth2callback"
cookie_secret = "UMA-SEQUENCIA-ALEATORIA-LONGA-E-SECRETA"
client_id = "CLIENT_ID_DO_OAUTH_WEB"
client_secret = "CLIENT_SECRET_DO_OAUTH_WEB"
server_metadata_url = "https://accounts.google.com/.well-known/openid-configuration"

# Use APENAS UM dos blocos abaixo.
# Se a pasta estiver em Drives compartilhados:
[drive_service_account]
type = "service_account"
project_id = "ID_DO_PROJETO_GOOGLE"
private_key_id = "ID_DA_CHAVE"
private_key = "-----BEGIN PRIVATE KEY-----\nCOLE_A_CHAVE\n-----END PRIVATE KEY-----\n"
client_email = "NOME_DA_CONTA@PROJETO.iam.gserviceaccount.com"
client_id = "ID_NUMERICO_DA_CONTA"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "URL_DO_CERTIFICADO_NO_JSON"

# Se estiver em Meu Drive, REMOVA TODO o bloco acima e use:
# [drive_oauth]
# client_id = "CLIENT_ID_DO_OAUTH_DESKTOP"
# client_secret = "CLIENT_SECRET_DO_OAUTH_DESKTOP"
# refresh_token = "TOKEN_OBTIDO_COM_AUTORIZAR_DRIVE_PY"
```

O comentário da opção B ilustra os campos; remova o `#` das quatro linhas
quando for usá-la. Não envie este conteúdo a terceiros nem o inclua no GitHub.
Para testar localmente, crie `.streamlit/secrets.toml` com a mesma estrutura e
troque o redirecionamento do `[auth]` para
`http://localhost:8501/oauth2callback`; inclua **também** esse URI no cliente
OAuth Web. O código da aplicação não precisa ser modificado.

## 4. Atualize e confira

Copie `app.py`, `requirements.txt`, `acompanhamento_drive_ui.py`,
`drive_acompanhamento.py`, `gerar_acompanhamento.py`, `autorizar_drive.py`,
`README.md` e a pasta inteira `assets` para o repositório que publica o site.
Mantenha os outros arquivos Python do pacote, pois o comparador precisa deles.
Faça commit e aguarde a nova versão do Streamlit carregar.

1. Abra **Comparador de APR** e teste a comparação antiga, sem alterar o fluxo.
2. Envie APR antiga e atual, clique em **Extrair** e **Comparar os cenários**.
3. Mude para **Acompanhamento de BowTies**, entre com Google usando e-mail
   permitido, digite o projeto e clique em **Criar planilha de acompanhamento**.
4. Clique em **Abrir esta planilha online** e confira as abas `Painel`,
   `Controle`, `Equipe`; confirme candidatos, nós, cenários e cores com a APR.
5. Um analista deve editar uma célula amarela e outro deve abrir o **mesmo link**
   para ver a edição. Clique de novo em **Criar**: o site deve abrir a mesma
   planilha, sem apagar os dados lançados.

Os campos de responsável, novo/correção, datas e progresso começam vazios ou
“A definir”, porque a APR não prova se já existe BowTie anterior nem quem fará
o trabalho. Faça essa classificação após conferir a comparação. A edição de
arquivos `.xlsx` no Google Sheets costuma salvar no próprio formato Office;
confira no primeiro arquivo se gráfico, validações e estilos do modelo aparecem
corretamente no navegador da equipe. Não converta o arquivo para uma segunda
planilha Google sem decidir trocar o formato oficial do acompanhamento.
