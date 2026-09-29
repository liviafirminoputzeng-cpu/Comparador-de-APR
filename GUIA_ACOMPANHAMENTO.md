# Acompanhamento compartilhado de BowTies

O site apresenta dois acessos: **Comparador de APR** (fluxo original) e
**Acompanhamento de BowTies** (projetos, equipe, painel e histórico).
O acompanhamento não gera Excel para baixar. Uma área não modifica os dados
da outra. Os PDFs de importação são processados na sessão, sem serem gravados
no banco. O banco guarda somente os campos dos cenários e as atualizações.

## Antes de publicar

Esta versão precisa de uma instância PostgreSQL compartilhada com conexão
TLS e de um aplicativo de login Microsoft Entra ID. A hospedagem do site no
Streamlit Community Cloud, sozinha, não fornece armazenamento persistente.
Até configurar estes serviços, a área de acompanhamento mostra uma mensagem
de configuração pendente e o comparador continua funcionando.

### 1. Banco de dados

1. Crie um banco PostgreSQL gerenciado autorizado pela empresa, com backup.
2. Abra o editor SQL desse serviço, ou acesse-o com um cliente PostgreSQL.
3. Execute **uma vez** todo o arquivo `schema.sql` no banco escolhido.
4. Guarde a URL de conexão no formato
   `postgresql://USUARIO:SENHA@HOST:5432/NOME_DO_BANCO`.
   A conexão do aplicativo exige TLS (`sslmode=require`). Se a senha contiver
   `@`, `:`, `/` ou `#`, codifique esses caracteres na URL.
5. Nunca coloque a URL, a senha ou o arquivo real `secrets.toml` no GitHub.

Os projetos, membros, BowTies e eventos de histórico são tabelas distintas.
O mesmo ID de cenário pode aparecer em **projetos diferentes**, mas não pode
ser cadastrado duas vezes no mesmo projeto. Importações repetidas ignoram
IDs já presentes, preservando o acompanhamento da equipe.

### 2. Login Microsoft

1. Peça ao administrador do Microsoft 365 para registrar um aplicativo Web
   no Microsoft Entra ID, restrito ao diretório da empresa.
2. Configure o URI de redirecionamento do aplicativo como
   `https://SEU-SITE.streamlit.app/oauth2callback`.
3. Guarde o **Tenant ID**, o **Application/Client ID** e o **Client Secret**.
4. O e-mail enviado pelo provedor de login precisa coincidir com o e-mail
   cadastrado para o coordenador ou analista no projeto.

O e-mail listado em `tracker_admin_emails` pode criar projetos e coordenar
qualquer projeto. Cadastre pelo menos uma conta do coordenador. Coordenadores
podem dar acesso a outras pessoas dentro do projeto. Analistas só acessam
projetos nos quais foram cadastrados e só alteram os próprios BowTies.

### 3. Configurar Secrets no Streamlit Community Cloud

Abra o aplicativo publicado → **Settings** → **Secrets** e insira os valores
correspondentes ao exemplo `.streamlit/secrets.toml.example`. Use o endereço
real do site publicado no `redirect_uri`. Atualize também o URI na configuração
Microsoft Entra. Salve as configurações e reinicie o aplicativo.

Para testes locais no VS Code, crie `.streamlit/secrets.toml` com os valores
locais e o endereço `http://localhost:8501/oauth2callback`. Esse arquivo
está no `.gitignore`; confira que não foi adicionado ao GitHub.

### 4. Atualizar o site

Copie **todos** os arquivos deste pacote para a pasta do projeto que publica
o comparador, preservando a pasta `assets`. Atualize `requirements.txt`;
ele instala `psycopg[binary]` para acesso ao PostgreSQL. Faça commit e envie
as mudanças à branch do GitHub vinculada ao Streamlit. Se você usou a
interface do GitHub para upload, envie também os arquivos novos
`acompanhamento_ui.py`, `acompanhamento_db.py`, `schema.sql` e os arquivos
na pasta `.streamlit` — somente o exemplo, nunca o segredo real.

### 5. Primeiro uso

1. Abra **Acompanhamento de BowTies**, clique em **Entrar** e acesse com o
   e-mail definido em `tracker_admin_emails`.
2. Crie um projeto, por exemplo, uma instalação e uma revisão da APR.
3. Cadastre o coordenador e os dois analistas com os e-mails corporativos.
4. Na seção **Equipe e cadastro**, carregue a APR atualizada em PDF, confira
   a lista de candidatos e clique em **Cadastrar estes candidatos**. Também
   é possível importar a planilha inicial `Acompanhamento_BowTies_APR_RevK.xlsx`
   pela aba **Planilha modelo (Excel)**, desde que o andamento e as atribuições
   nela ainda estejam em branco. Não importe ambos se não for necessário;
   os IDs repetidos serão ignorados.
5. Em **Atualizar BowTie**, defina trabalho **Novo** ou **Correção**,
   responsável, prioridade e prazo. Cada analista registra o andamento,
   percentual, datas e impedimentos dos seus itens. O coordenador confere e
   marca **Concluído** com a data da validação.
6. O painel aberto por outra pessoa consulta o banco automaticamente a cada
   **20 segundos**. Em uma edição simultânea do mesmo item, o segundo
   salvamento é recusado e pede atualização da página. O histórico registra
   autor, horário e os campos anteriores e novos.

Para outro trabalho, crie **outro projeto**, cadastre a equipe correspondente
e importe a APR desse projeto. Os dados ficam separados.

## Verificação antes de liberar à equipe

Faça login com uma conta de coordenador e com uma conta de analista em dois
navegadores. Atribua um item, altere a situação com o analista e confirme que
o outro painel reflete a mudança em até 20 segundos. Confira também se uma
conta não cadastrada não vê o projeto e se o comparador ainda gera os Excel
e BowTies normalmente.

## Limites desta versão

- A área nova **não disponibiliza exportação Excel**.
- A importação da planilha modelo aceita o cadastro inicial sem andamento.
  Para planilhas já editadas, ela recusa a importação para evitar perder
  responsável, status ou prazo.
- Os horários do histórico são gravados com fuso horário pelo PostgreSQL.
- Não publique uma URL de banco com privilégios administrativos mais amplos
  que os necessários para as tabelas de acompanhamento.
