import hashlib
import json
import re
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from bowtie_preview import altura_previa_bowtie, montar_html_bowtie
from comparacao import (
    candidato_bowtie,
    comparar_revisoes,
    inventariar_salvaguardas_bowtie,
)
from extracao import extrair_cenarios_pdf
from exportacao import gerar_excel, nome_arquivo_excel


st.set_page_config(
    page_title="Comparador de APR",
    page_icon="📋",
    layout="wide",
)

CAMINHO_LOGO = Path(__file__).resolve().parent / "assets" / "logo_putz.jpg"

coluna_logo, coluna_titulo = st.columns([1, 5], vertical_alignment="center")
with coluna_logo:
    if CAMINHO_LOGO.exists():
        st.image(str(CAMINHO_LOGO), width=170)
with coluna_titulo:
    st.title("Comparador de revisões de APR")
    st.write(
        "Envie a versão antiga e a versão atualizada para extrair e comparar "
        "os cenários das duas APRs."
    )

st.info(
    "O resultado será disponibilizado em um único arquivo Excel (.xlsx), "
    "com todas as planilhas da análise organizadas por nó."
)

coluna_antiga, coluna_atualizada = st.columns(2)

with coluna_antiga:
    apr_antiga = st.file_uploader(
        "1. Selecione a APR antiga",
        type=["pdf"],
        key="apr_antiga",
    )

with coluna_atualizada:
    apr_atualizada = st.file_uploader(
        "2. Selecione a APR atualizada",
        type=["pdf"],
        key="apr_atualizada",
    )


def extrair_documento(arquivo):
    resultado = extrair_cenarios_pdf(arquivo.getvalue())
    resultado["nome_arquivo"] = arquivo.name
    return resultado


def assinatura_pdf(arquivo):
    return hashlib.sha256(arquivo.getvalue()).hexdigest()


def _texto_assinatura(valor):
    return re.sub(r"\s+", " ", str(valor or "")).strip().casefold()


def assinatura_conteudo(resultado):
    campos = (
        "Sistema",
        "Trecho de análise",
        "Perigo",
        "Causas",
        "Possíveis efeitos",
        "Modo de detecção",
        "Salvaguardas preventivas (SP)",
        "Salvaguardas mitigadoras (SM)",
        "F",
        "Sev S",
        "Sev P",
        "Sev M",
        "Sev I",
        "Risco S",
        "Risco P",
        "Risco M",
        "Risco I",
    )
    conteudo = [
        [_texto_assinatura(cenario.get(campo, "")) for campo in campos]
        for cenario in resultado["cenarios"]
    ]
    serializado = json.dumps(conteudo, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(serializado.encode("utf-8")).hexdigest()


def documentos_extraidos_identicos(antiga, atualizada):
    return assinatura_conteudo(antiga) == assinatura_conteudo(atualizada)


def _numero(valor):
    achado = re.search(r"\d+", str(valor or ""))
    return int(achado.group()) if achado else 10**9


def _ordenar_cenarios(cenarios):
    return sorted(
        cenarios,
        key=lambda cenario: (
            _numero(cenario.get("Nó")),
            _numero(cenario.get("Cenário")),
            str(cenario.get("ID", "")),
        ),
    )


STATUS_EXIBICAO = {
    "MANTIDO / ALTERADO": "NOME DO DESVIO MANTIDO",
    "RENOMEADO / ALTERADO": "NOME DO DESVIO ALTERADO",
    "REMOVIDO": "DESVIO REMOVIDO",
    "NOVO": "DESVIO NOVO",
}


arquivos_enviados = apr_antiga is not None and apr_atualizada is not None
pdfs_identicos = (
    arquivos_enviados
    and assinatura_pdf(apr_antiga) == assinatura_pdf(apr_atualizada)
)

if pdfs_identicos:
    st.error("Documento repetido. Por favor, verifique os documentos.")
    for chave in (
        "resultado_antiga",
        "resultado_atualizada",
        "resultado_comparacao",
        "arquivo_excel",
        "nome_excel",
    ):
        st.session_state.pop(chave, None)

if st.button(
    "Extrair cenários das duas APRs",
    type="primary",
    disabled=not arquivos_enviados or pdfs_identicos,
):
    try:
        with st.spinner(
            "Extraindo as tabelas. Em documentos grandes isso pode levar alguns minutos..."
        ):
            resultado_antiga = extrair_documento(apr_antiga)
            resultado_atualizada = extrair_documento(apr_atualizada)

        if documentos_extraidos_identicos(resultado_antiga, resultado_atualizada):
            st.error("Documento repetido. Por favor, verifique os documentos.")
            st.stop()

        st.session_state["resultado_antiga"] = resultado_antiga
        st.session_state["resultado_atualizada"] = resultado_atualizada
        st.session_state.pop("resultado_comparacao", None)
        st.session_state.pop("arquivo_excel", None)
        st.session_state.pop("nome_excel", None)
        st.success("Extração concluída.")

    except Exception as erro:
        st.error("Não foi possível extrair os cenários dos documentos.")
        st.code(str(erro))


if (
    "resultado_antiga" in st.session_state
    and "resultado_atualizada" in st.session_state
):
    antiga = st.session_state["resultado_antiga"]
    atualizada = st.session_state["resultado_atualizada"]

    st.divider()
    st.subheader("Resultado da extração")

    metrica_1, metrica_2, metrica_3, metrica_4 = st.columns(4)
    metrica_1.metric("Cenários da APR antiga", len(antiga["cenarios"]))
    metrica_2.metric("Cenários da APR atualizada", len(atualizada["cenarios"]))
    metrica_3.metric("Páginas da APR antiga", antiga["total_paginas"])
    metrica_4.metric("Páginas da APR atualizada", atualizada["total_paginas"])

    dados_antiga, dados_atualizada = st.columns(2)

    with dados_antiga:
        st.write(f"**Relatório antigo:** {antiga['relatorio'] or 'Não identificado'}")
        st.write(f"**Aprovação:** {antiga['aprovacao'] or 'Não identificada'}")

    with dados_atualizada:
        st.write(
            f"**Relatório atualizado:** "
            f"{atualizada['relatorio'] or 'Não identificado'}"
        )
        st.write(
            f"**Aprovação:** "
            f"{atualizada['aprovacao'] or 'Não identificada'}"
        )

    colunas_visualizacao = [
        "ID",
        "Nó",
        "Sistema",
        "Cenário",
        "Perigo",
        "Causas",
        "Possíveis efeitos",
        "Modo de detecção",
        "Salvaguardas preventivas (SP)",
        "Salvaguardas mitigadoras (SM)",
        "F",
        "Sev S",
        "Sev P",
        "Sev M",
        "Sev I",
        "Risco S",
        "Risco P",
        "Risco M",
        "Risco I",
        "Observações / Recomendações",
        "Página PDF",
    ]

    tabela_antiga = pd.DataFrame(_ordenar_cenarios(antiga["cenarios"]))[
        colunas_visualizacao
    ]
    tabela_atualizada = pd.DataFrame(_ordenar_cenarios(atualizada["cenarios"]))[
        colunas_visualizacao
    ]

    aba_antiga, aba_atualizada = st.tabs(
        ["Cenários da APR antiga", "Cenários da APR atualizada"]
    )

    with aba_antiga:
        st.dataframe(
            tabela_antiga,
            use_container_width=True,
            hide_index=True,
            height=500,
        )

    with aba_atualizada:
        st.dataframe(
            tabela_atualizada,
            use_container_width=True,
            hide_index=True,
            height=500,
        )

    ids_antiga = [cenario["ID"] for cenario in antiga["cenarios"]]
    ids_atualizada = [cenario["ID"] for cenario in atualizada["cenarios"]]

    if len(ids_antiga) != len(set(ids_antiga)):
        st.warning("A APR antiga possui IDs de cenário repetidos.")

    if len(ids_atualizada) != len(set(ids_atualizada)):
        st.warning("A APR atualizada possui IDs de cenário repetidos.")

    st.divider()
    st.subheader("Comparação entre as revisões")
    st.write(
        "O sistema prioriza o nome do desvio/perigo para relacionar os cenários. "
        "Sistema, trecho, causas, efeitos e equipamentos são usados como apoio. "
        "O número do nó ou do cenário não determina a correspondência."
    )

    if st.button("Comparar os cenários", type="primary"):
        try:
            with st.spinner("Comparando os cenários e identificando alterações..."):
                resultado_comparacao = comparar_revisoes(
                    antiga["cenarios"], atualizada["cenarios"]
                )
            st.session_state["resultado_comparacao"] = resultado_comparacao
            st.session_state.pop("arquivo_excel", None)
            st.session_state.pop("nome_excel", None)
            st.success("Comparação concluída.")
        except Exception as erro:
            st.error("Não foi possível comparar os cenários.")
            st.code(str(erro))

    if "resultado_comparacao" in st.session_state:
        comparacao = st.session_state["resultado_comparacao"]
        resumo = comparacao["resumo"]

        resumo_1, resumo_2, resumo_3, resumo_4 = st.columns(4)
        resumo_1.metric(
            "Nome do desvio mantido",
            resumo.get("MANTIDO / ALTERADO", 0),
            help=(
                "Quantidade de cenários encontrados nas duas APRs com o mesmo "
                "nome de desvio. Outros campos podem ter sido alterados."
            ),
        )
        resumo_2.metric(
            "Nome do desvio alterado",
            resumo.get("RENOMEADO / ALTERADO", 0),
            help=(
                "Quantidade de cenários correspondentes cujo nome do desvio "
                "foi modificado entre as revisões."
            ),
        )
        resumo_3.metric(
            "Desvios removidos da APR atual",
            resumo.get("REMOVIDO", 0),
        )
        resumo_4.metric(
            "Desvios novos na APR atual",
            resumo.get("NOVO", 0),
        )

        st.caption(
            "“Nome do desvio mantido” significa que o cenário foi localizado nas "
            "duas APRs com o mesmo nome; causas, efeitos, frequência, risco ou "
            "salvaguardas ainda podem ter mudado. “Nome do desvio alterado” "
            "significa que o cenário foi relacionado pelo conteúdo, mas recebeu "
            "outro nome na revisão atual."
        )

        bowtie_1, bowtie_2, bowtie_3, bowtie_4 = st.columns(4)
        bowtie_1.metric(
            "Candidatos Bow Tie na APR antiga",
            resumo.get("CANDIDATOS BOW TIE ANTIGA", 0),
        )
        bowtie_2.metric(
            "Candidatos Bow Tie na APR atualizada",
            resumo.get("CANDIDATOS BOW TIE ATUALIZADA", 0),
        )
        bowtie_3.metric(
            "Novos candidatos",
            resumo.get("NOVOS CANDIDATOS BOW TIE", 0),
        )
        bowtie_4.metric(
            "Deixaram de ser candidatos",
            resumo.get("DEIXARAM DE SER CANDIDATOS BOW TIE", 0),
        )

        st.caption(
            "Critério aplicado: cenário com severidade IV ou V em Segurança, "
            "Patrimônio ou Meio Ambiente. Para Gestão Dinâmica de Barreiras, "
            "o Guia Bow Tie rev.7 também inclui Imagem IV/V."
        )

        st.info(
            "Correspondências com confiança baixa e cenários sem par devem "
            "ser conferidos antes da emissão do Excel final."
        )

        (
            aba_correspondencias,
            aba_comparativo,
            aba_alteracoes,
            aba_bowtie,
            aba_previa_bowtie,
        ) = st.tabs(
            [
                "Correspondências",
                "Comparativo",
                "Alterações detectadas",
                "Candidatos Bow Tie",
                "Prévia Bow Tie",
            ]
        )

        with aba_correspondencias:
            tabela_correspondencias = pd.DataFrame(
                comparacao["correspondencias"]
            )
            tabela_correspondencias["Status"] = tabela_correspondencias[
                "Status"
            ].replace(STATUS_EXIBICAO)
            status_disponiveis = tabela_correspondencias["Status"].unique().tolist()
            status_selecionados = st.multiselect(
                "Filtrar por status",
                status_disponiveis,
                default=status_disponiveis,
            )
            tabela_filtrada = tabela_correspondencias[
                tabela_correspondencias["Status"].isin(status_selecionados)
            ]
            st.dataframe(
                tabela_filtrada,
                use_container_width=True,
                hide_index=True,
                height=520,
            )

        with aba_comparativo:
            tabela_comparativo = pd.DataFrame(comparacao["comparativo"])
            tabela_comparativo["Status"] = tabela_comparativo["Status"].replace(
                STATUS_EXIBICAO
            )
            st.dataframe(
                tabela_comparativo,
                use_container_width=True,
                hide_index=True,
                height=520,
            )

        with aba_alteracoes:
            tabela_alteracoes = pd.DataFrame(comparacao["alteracoes"])
            tipos_disponiveis = sorted(
                tabela_alteracoes["Tipo de alteração"].unique().tolist()
            )
            tipos_selecionados = st.multiselect(
                "Filtrar por tipo de alteração",
                tipos_disponiveis,
                default=tipos_disponiveis,
            )
            alteracoes_filtradas = tabela_alteracoes[
                tabela_alteracoes["Tipo de alteração"].isin(tipos_selecionados)
            ]
            st.dataframe(
                alteracoes_filtradas,
                use_container_width=True,
                hide_index=True,
                height=520,
            )

        with aba_bowtie:
            tabela_bowtie = pd.DataFrame(comparacao["evolucao_bowtie"])
            situacoes_disponiveis = tabela_bowtie[
                "Situação Bow Tie"
            ].unique().tolist()
            situacoes_padrao = [
                situacao
                for situacao in situacoes_disponiveis
                if situacao != "NÃO CANDIDATO"
            ]
            situacoes_selecionadas = st.multiselect(
                "Filtrar por situação Bow Tie",
                situacoes_disponiveis,
                default=situacoes_padrao,
            )
            bowtie_filtrado = tabela_bowtie[
                tabela_bowtie["Situação Bow Tie"].isin(situacoes_selecionadas)
            ]
            st.dataframe(
                bowtie_filtrado,
                use_container_width=True,
                hide_index=True,
                height=520,
            )

        with aba_previa_bowtie:
            st.write(
                "Selecione um nó e, em seguida, o desvio que será usado como "
                "evento topo da prévia."
            )
            mostrar_todos = st.checkbox(
                "Mostrar também desvios não candidatos",
                value=False,
                key="mostrar_todos_previa_bowtie",
                help=(
                    "Desmarcado: apresenta apenas cenários com severidade IV ou V "
                    "em Segurança, Patrimônio ou Meio Ambiente."
                ),
            )

            cenarios_previa = _ordenar_cenarios(atualizada["cenarios"])
            if not mostrar_todos:
                cenarios_previa = [
                    cenario
                    for cenario in cenarios_previa
                    if candidato_bowtie(cenario)
                ]

            if not cenarios_previa:
                st.warning(
                    "Nenhum cenário atende ao critério de candidato a Bow Tie "
                    "na APR atualizada."
                )
            else:
                nos_disponiveis = sorted(
                    {cenario["Nó"] for cenario in cenarios_previa},
                    key=_numero,
                )
                sistemas_por_no = {}
                for no in nos_disponiveis:
                    sistemas = list(
                        dict.fromkeys(
                            str(cenario.get("Sistema", "") or "")
                            for cenario in cenarios_previa
                            if cenario["Nó"] == no
                        )
                    )
                    sistemas_por_no[no] = " / ".join(
                        sistema for sistema in sistemas if sistema
                    )

                coluna_no, coluna_desvio = st.columns([1, 2])
                with coluna_no:
                    if st.session_state.get("no_previa_bowtie") not in (
                        None,
                        *nos_disponiveis,
                    ):
                        st.session_state.pop("no_previa_bowtie", None)
                    no_selecionado = st.selectbox(
                        "Nó",
                        nos_disponiveis,
                        format_func=lambda no: (
                            f"Nó {no} — {sistemas_por_no.get(no, '')}"
                        ),
                        key="no_previa_bowtie",
                    )

                cenarios_do_no = [
                    cenario
                    for cenario in cenarios_previa
                    if cenario["Nó"] == no_selecionado
                ]
                cenarios_por_id = {
                    cenario["ID"]: cenario for cenario in cenarios_do_no
                }

                with coluna_desvio:
                    id_selecionado = st.selectbox(
                        "Desvio / evento topo",
                        list(cenarios_por_id),
                        format_func=lambda identificador: (
                            f"{identificador} — "
                            f"{str(cenarios_por_id[identificador].get('Perigo', '')).replace(chr(10), ' ')}"
                        ),
                        key=f"desvio_previa_bowtie_{no_selecionado}",
                    )

                cenario_selecionado = cenarios_por_id[id_selecionado]
                inventario_cenario = inventariar_salvaguardas_bowtie(
                    [cenario_selecionado],
                    "APR atualizada",
                )
                html_previa = montar_html_bowtie(
                    cenario_selecionado,
                    inventario_cenario,
                )
                altura_previa = altura_previa_bowtie(
                    cenario_selecionado,
                    inventario_cenario,
                )
                components.html(
                    html_previa,
                    height=altura_previa,
                    scrolling=True,
                )

        st.divider()
        st.subheader("Gerar arquivo Excel")
        st.write(
            "O arquivo reúne o resumo, as duas APRs estruturadas, o comparativo, "
            "as alterações e uma aba exclusiva com os candidatos Bow Tie. Todas "
            "as planilhas são entregues no formato Excel (.xlsx)."
        )

        if st.button("Preparar Excel para download"):
            try:
                with st.spinner("Montando e formatando o arquivo Excel..."):
                    arquivo_excel = gerar_excel(antiga, atualizada, comparacao)
                st.session_state["arquivo_excel"] = arquivo_excel
                st.session_state["nome_excel"] = nome_arquivo_excel(
                    antiga, atualizada
                )
                st.success("Arquivo Excel preparado.")
            except Exception as erro:
                st.error("Não foi possível montar o arquivo Excel.")
                st.code(str(erro))

        if "arquivo_excel" in st.session_state:
            st.download_button(
                "Baixar arquivo Excel (.xlsx)",
                data=st.session_state["arquivo_excel"],
                file_name=st.session_state["nome_excel"],
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                type="primary",
            )
