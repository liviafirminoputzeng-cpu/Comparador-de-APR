# Comparador de APR e BowTies por cenário

## O que o site entrega

- Um Excel com a comparação entre as duas versões da APR e os candidatos a BowTie.
- Um BowTie novo por cenário da APR atualizada, em `.xlsx`, no desenho do arquivo `NÓ_1.xlsx`. Você pode baixar um arquivo ou todos em ZIP.
- A atualização individual de um BowTie `.xlsm` existente, mantendo o desenho e os macros desse arquivo enviado. Para isso, selecione o cenário correspondente e envie o `.xlsm` no campo opcional.

Na aba **BT**, entram apenas o evento topo, as ameaças e as consequências. As colunas B1 a B10 mantêm suas cores e ficam livres de textos das salvaguardas. Cada salvaguarda identificada na APR vai para a aba de barreiras correspondente; os modos de detecção não entram nos BowTies. Nos itens **SM2, SM7, SM8 e SM10**, o código vai para a coluna **TAG** e o restante do texto vai para a descrição. Os demais campos de TAG só são preenchidos se houver evidência na APR. Todo o texto visível nas planilhas geradas é convertido para letras maiúsculas. As tabelas terminam na última linha preenchida, respeitando as linhas de cabeçalho e as mesclagens necessárias. As cores, cabeçalhos, larguras e estilos vêm do modelo original; novas ameaças e consequências repetem a formatação das anteriores. Nos BowTies novos, classificações que precisam de confirmação vão para a aba **A_conferir**. Informações de SAP, FLOC e outras fontes não enviadas com a APR não são preenchidas.

O novo arquivo `.xlsx` não tem macros porque o modelo `NÓ_1.xlsx` também não tem. Os macros de um `.xlsm` que você enviar para atualização são mantidos. A atualização preserva outros dados já existentes no BowTie enviado para sua conferência; esses dados anteriores não são confirmados pela APR. Confira as classes das barreiras e o conteúdo dos arquivos antes de usá-los.

## Usar no site

1. Envie os PDFs da APR antiga e da APR atualizada, nessa ordem.
2. Clique em **Extrair cenários das duas APRs** e depois em **Comparar os cenários**.
3. Baixe a planilha de comparação em Excel, se desejar.
4. Em **Gerar ou atualizar BowTies por cenário**, escolha um cenário e clique em **Preparar BowTie selecionado**. Baixe o `.xlsx` preparado.
5. Se você já tem um BowTie `.xlsm` desse cenário, envie esse arquivo no campo opcional antes de preparar o BowTie: a versão atualizada será `.xlsm`.
6. Clique em **Preparar todos os BowTies (ZIP)** para criar um `.xlsx` novo por cenário disponível. Se desejar incluir também cenários que não são candidatos, marque a opção correspondente antes de gerar.

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

Copie **todos os arquivos e a pasta `assets`** deste pacote para o repositório usado pelo Streamlit. Preserve os caminhos `assets/modelo_no1.xlsx`, `assets/modelo_no.xlsm`, `assets/modelo_resultado_apr.xlsx` e `assets/logo_putz.jpg`. Envie os arquivos atualizados para a mesma branch configurada para executar `app.py`. Não envie `.venv`, PDFs de clientes ou arquivos Excel de teste ao repositório.
