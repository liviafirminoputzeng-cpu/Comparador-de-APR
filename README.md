# Comparador de APR e BowTies por cenário

## Acompanhamento de BowTies sem conexão à API do Drive

No topo, selecione **Comparador de APR**, **Acompanhamento de BowTies** ou
**Planilha de carga**.
Após extrair e comparar as duas APRs, a aba **Gerar planilha Excel** oferece
um download `.xlsx` por projeto com os candidatos da APR atualizada. A planilha
copia as abas `Painel`, `Controle` e `Equipe` do modelo enviado e mantém suas
cores, validações, fórmulas e gráfico. O site não define responsável, datas
ou se um BowTie existente precisará de correção; a equipe preencherá isso.

O coordenador envia o Excel à
[pasta compartilhada da PUTZ](https://drive.google.com/drive/folders/1ichy39_gjGtfzALO66pe275YxY-HdRSh?usp=drive_link)
manualmente e copia o link do arquivo. A aba `Controle` inclui a etapa de
produção da planilha e a atribuição a um dos dois estagiários; o `Painel`
resume a produção dos itens classificados como `Novo`. Na aba **Planilhas compartilhadas**,
pode colar nome do projeto e o **link da planilha, nunca o da pasta**. É
possível preparar vários projetos na mesma sessão, mover sua posição com
**↑/↓** e excluí-los da lista preparada. A tela gera o arquivo
`links_acompanhamento.json` na ordem escolhida; para que todos vejam as
inclusões, mudanças de ordem ou exclusões,
substitua esse arquivo no repositório GitHub e faça o commit. O site não
grava dados persistentes no servidor. Consulte
[`GUIA_ACOMPANHAMENTO.md`](GUIA_ACOMPANHAMENTO.md) para o passo a passo.
Não sobrescreva o catálogo existente ao atualizar o código do programa.

Não é necessário criar projeto Google Cloud, configurar OAuth nem usar Secrets
do Streamlit para esse fluxo. O Drive continua sendo necessário se a equipe
quiser editar o mesmo Excel online; o GitHub já usado pelo site guarda a lista
de links. Se o repositório for público, nomes e URLs do catálogo também serão
públicos, mesmo que os arquivos do Drive mantenham acesso restrito.

## O que o site entrega

- Um Excel com a comparação entre as duas versões da APR e os candidatos a BowTie.
- Um BowTie novo por cenário da APR atualizada, em `.xlsx`, no desenho do arquivo `NÓ_1.xlsx`. Você pode baixar um arquivo ou todos em ZIP.
- A atualização individual de um BowTie `.xlsm` existente, mantendo o desenho e os macros desse arquivo enviado. Para isso, selecione o cenário correspondente e envie o `.xlsm` no campo opcional.
- Na área **Planilha de carga**, envie um BowTie `.xlsx` ou `.xlsm` para baixar
  um Excel com as abas `Def` e `Data`. A aba `Def` recebe o nó, evento topo,
  ameaças e consequências da aba `BT`; `Data` reúne as linhas preenchidas das
  abas de barreiras `B1`, `B2` etc. Os estilos, cabeçalhos, larguras e filtros
  seguem o modelo enviado. Os registros de exemplo do modelo são removidos.
  Se `SITE_CODE`, `AREA_CODE` ou `ID` não estiverem no BowTie, o site avisa e
  deixa as células correspondentes vazias. O Excel gerado não inclui macros;
  o BowTie de entrada permanece intacto.
- Para processar vários BowTies, entre em **Planilha de carga → Vários BowTies
  em ZIP** e envie um arquivo `.zip` contendo os `.xlsx` e/ou `.xlsm`. O site
  disponibiliza um ZIP com uma planilha de carga por BowTie, mantendo nomes
  distintos mesmo que dois arquivos em pastas diferentes tenham o mesmo nome.
  Pastas e arquivos que não são Excel são ignorados. Se um dos BowTies Excel
  estiver inválido, o site mostra qual é e não libera um lote incompleto;
  corrija-o ou retire-o do ZIP e envie novamente. Os avisos por BowTie podem
  ser conferidos antes do download.

Na aba **BT**, entram apenas o evento topo, as ameaças e as consequências. As colunas B1 a B10 mantêm suas cores e ficam livres de textos das salvaguardas. Cada salvaguarda identificada na APR vai para a aba de barreiras correspondente; os modos de detecção não entram nos BowTies. Nos itens **SM2, SM7, SM8 e SM10**, o código vai para a coluna **TAG** e o restante do texto vai para a descrição. Os demais campos de TAG só são preenchidos se houver evidência na APR. Todo o texto visível nas planilhas geradas é convertido para letras maiúsculas. As tabelas terminam na última linha preenchida, respeitando as linhas de cabeçalho e as mesclagens necessárias. As cores, cabeçalhos, larguras e estilos vêm do modelo original; novas ameaças e consequências repetem a formatação das anteriores. Nos BowTies novos, classificações que precisam de confirmação vão para a aba **A_conferir**. Informações de SAP, FLOC e outras fontes não enviadas com a APR não são preenchidas.

Nas abas de barreiras, a coluna **ID** de cada linha preenchida usa `=BT!$W$5`. O mesmo vale para a coluna ID da aba **BT - Todos** quando ela existe. Em um arquivo novo, `BT!W5` recebe o identificador do cenário extraído da APR; ao atualizar um `.xlsm`, um ID já registrado em `BT!W5` é mantido. Quando texto e código correspondem, aplicam-se as associações fornecidas pela equipe: SP19/procedimento de execução, SP13/SPDA e SP45/drenagem da área de descarregamento em B1; procedimento operacional de emergência em B8; SM1/brigada em B9; SM4/dique, SM10/kit de proteção ambiental e SM8/Diphoterine em B6. A classificação dessas associações é uma regra da equipe e continua sujeita a conferência técnica.

O novo arquivo `.xlsx` não tem macros porque o modelo `NÓ_1.xlsx` também não tem. Os macros de um `.xlsm` que você enviar para atualização são mantidos. A atualização preserva outros dados já existentes no BowTie enviado para sua conferência; esses dados anteriores não são confirmados pela APR. Confira as classes das barreiras e o conteúdo dos arquivos antes de usá-los.

## Usar no site

1. Envie os PDFs da APR antiga e da APR atualizada, nessa ordem.
2. Clique em **Extrair cenários das duas APRs** e depois em **Comparar os cenários**.
3. Na aba **Planilha de comparação**, confira os resultados e clique em **Gerar e baixar comparação em Excel (.xlsx)**. O download começa com um clique.
4. Na aba **BowTies por cenário**, selecione um cenário e clique em **Gerar e baixar este BowTie (Excel)**.
5. Se você já tem um BowTie `.xlsm` desse cenário, envie esse arquivo no campo opcional antes de baixar a atualização: a versão atualizada será `.xlsm` e manterá as macros.
6. Clique em **Gerar e baixar todos os BowTies (.zip)** para baixar um `.xlsx` novo por cenário disponível. Para incluir cenários que não são candidatos, marque a opção antes do download. O ZIP é preparado na tela após a comparação, então arquivos grandes podem exigir alguns instantes antes de o botão aparecer.

## Rodar no VS Code, Windows

Abra esta pasta no VS Code. No terminal PowerShell, execute:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Se `.venv` já existe, execute apenas os três últimos comandos. Se o site já está rodando, interrompa-o no terminal com `Ctrl+C` antes de iniciá-lo novamente.

## Atualizar o site publicado

Copie **todos os arquivos e a pasta `assets`** deste pacote para o repositório
usado pelo Streamlit. Preserve os caminhos `assets/modelo_no1.xlsx`,
`assets/modelo_no.xlsm`, `assets/modelo_resultado_apr.xlsx` e
`assets/putz_identidade.webp`, além do novo
`assets/modelo_planilha_carga.xlsx`. Envie os arquivos atualizados para a mesma
branch configurada para executar `app.py`. **Preserve no GitHub o arquivo
`links_acompanhamento.json` que já reúne seus projetos.** Ele não está neste
pacote de atualização para evitar a perda de links cadastrados. Não envie
`.venv`, PDFs de clientes ou arquivos Excel de teste ao repositório.

Ao trocar a versão anterior com API do Google, os arquivos antigos
`acompanhamento_drive_ui.py`, `drive_acompanhamento.py`, `autorizar_drive.py`
e `GUIA_GOOGLE_DRIVE.md` não são mais utilizados e podem ser retirados do
repositório. O novo `links_acompanhamento.json` deve ficar ao lado de `app.py`.
