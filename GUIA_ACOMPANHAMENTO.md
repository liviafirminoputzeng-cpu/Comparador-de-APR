# Acompanhamento de BowTies: Excel e links compartilhados

## Publicar esta versão

1. Extraia o ZIP e envie **seu conteúdo** ao repositório GitHub do site,
   substituindo `app.py` e `requirements.txt`. Inclua os novos arquivos
   `acompanhamento_manual_ui.py`, `gerar_acompanhamento.py`,
   `assets/modelo_acompanhamento.xlsx`. **Não substitua o arquivo
   `links_acompanhamento.json` já existente no GitHub**, pois ele guarda os
   links dos projetos anteriores. Se ainda não existir, o primeiro cadastro
   no site produzirá esse arquivo.
2. Espere o Streamlit atualizar. No topo do site, selecione
   **Acompanhamento de BowTies**. Confirme as duas abas:
   **Gerar planilha Excel** e **Planilhas compartilhadas**.
3. Não é necessário ativar a Google Drive API, criar um projeto no Google
   Cloud nem inserir senhas no Streamlit para usar este fluxo.

## Gerar a planilha de um projeto

1. Na área **Comparador de APR**, envie os PDFs antigo e atualizado e clique
   em **Extrair cenários das duas APRs**, depois em **Comparar os cenários**.
2. No topo, mude para **Acompanhamento de BowTies → Gerar planilha Excel**.
   Confira os nomes das APRs, informe o nome do projeto e clique em
   **Baixar planilha de acompanhamento (.xlsx)**.
3. Abra o Excel e confira as abas `Painel`, `Controle` e `Equipe`. Nós,
   cenários, sistemas, eventos topo e páginas vêm da APR atualizada. O
   responsável, tipo de trabalho, prazo e progresso ficam para a equipe.
   Na aba `Controle`, classifique `Trabalho` como `Novo` ou `Correção`;
   para cada `Novo`, escolha a etapa em `Planilha produzida` (`A fazer`,
   `Em produção` ou `Produzida`) e o `Responsável pela produção`.
   Cadastre os nomes dos dois estagiários em `Equipe`, células `C25:C26`.
   O bloco final do `Painel` mostra as planilhas novas por etapa e por
   estagiário, além das que faltam atribuir ou classificar.

## Compartilhar a planilha no Drive

1. Abra a pasta da equipe:
   https://drive.google.com/drive/folders/1ichy39_gjGtfzALO66pe275YxY-HdRSh?usp=drive_link
2. Clique em **Novo → Upload de arquivo** e escolha o `.xlsx` baixado.
3. No Drive, configure o acesso dos dois analistas e do coordenador como
   **Editor**. Abra o Excel no navegador e confirme que ele pode ser editado
   sem convertê-lo para outro arquivo.
4. Clique em **Compartilhar → Copiar link** do **arquivo Excel**, não da pasta.

## Colocar o link na aba do site

1. No site, abra **Acompanhamento de BowTies → Planilhas compartilhadas**.
2. Preencha **Projeto da planilha compartilhada** e cole o link do arquivo.
3. Clique em **Preparar inclusão do link**. O link é mostrado para teste na
   sua sessão; ainda não está publicado para a equipe.
4. Você pode repetir os passos 2 e 3 para preparar **vários projetos** nesta
   sessão. Os novos projetos entram no fim da lista; trocar o link de um
   projeto existente mantém sua posição.
5. Depois do último projeto, clique em **Baixar catálogo atualizado para
   publicar no site (.json)**. Ele contém os projetos anteriores e os novos.
6. No GitHub, abra o mesmo repositório do site e substitua o arquivo
   `links_acompanhamento.json` pelo arquivo recém-baixado. Faça **Commit
   changes**. Depois da atualização do Streamlit, os links publicados
   aparecerão para todos os visitantes da aba.

## Mudar a ordem ou excluir um link

1. Abra **Acompanhamento de BowTies → Planilhas compartilhadas**.
2. Use **↑** ou **↓** ao lado do projeto para movê-lo uma posição. Use
   **Excluir** para retirá-lo da lista. Esses botões preparam uma prévia
   apenas na sua sessão; o arquivo do Drive não é apagado.
3. Confira a lista. Se mudar de ideia, clique em **Desfazer alterações desta
   sessão**. Você também pode cadastrar novos links antes de publicar.
4. Clique em **Baixar catálogo atualizado para publicar no site (.json)**.
   No GitHub, substitua o arquivo `links_acompanhamento.json` da página
   principal do repositório e faça **Commit changes**. Só então toda a
   equipe verá a nova ordem e as exclusões.

Para adicionar projetos em outro dia, **abra o site novamente depois que o
GitHub publicar o catálogo anterior**; assim, o próximo download já incluirá
todos os projetos anteriores. Se usar o mesmo nome de projeto, o link anterior
será substituído no novo catálogo. Não edite o Excel do modelo dentro do GitHub:
o acompanhamento da equipe acontece no arquivo compartilhado do Drive.

O site é público, portanto a lista de nomes e URLs no catálogo também poderá
ser vista por visitantes e no repositório, caso ele seja público. O acesso ao
conteúdo dos Excel depende das permissões configuradas no Google Drive.
