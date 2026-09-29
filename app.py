import base64
import hashlib
import io
import json
import re
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from bowtie_preview import altura_previa_bowtie, montar_html_bowtie
from bowtie_excel import gerar_bowtie, nome_bowtie
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

area = st.radio(
    "Área do site",
    ["Comparador de APR", "Acompanhamento de BowTies"],
    horizontal=True,
    label_visibility="collapsed",
    key="area_site",
)
if area == "Acompanhamento de BowTies":
    from acompanhamento_drive_ui import mostrar_acompanhamento

    mostrar_acompanhamento()
    st.stop()

CAMINHO_CAPA = Path(__file__).resolve().parent / "assets" / "putz_identidade.webp"
capa = (
    f"data:image/webp;base64,{base64.b64encode(CAMINHO_CAPA.read_bytes()).decode('ascii')}"
    if CAMINHO_CAPA.exists() else ""
)
st.markdown(
    """
    <style>
    .stApp {background: #f3f5fc; color: #101632;}
    .block-container {max-width: 1500px; padding-top: 1.4rem;}
    .putz-hero {
        height: 260px; border-radius: 18px; overflow: hidden;
        background-color: #0a0a22;
        background-image: linear-gradient(90deg, rgba(7,8,30,.94) 0%,
          rgba(7,8,30,.85) 32%, rgba(7,8,30,.03) 63%), url('""" + capa + """');
        background-size: cover; background-position: center 52%;
        padding: 2.3rem 3rem; display: flex; flex-direction: column;
        justify-content: center; box-shadow: 0 8px 28px #151a3d22;
    }
    .putz-hero h1 {color: #fff; font-size: clamp(1.6rem, 3vw, 2.65rem);
        max-width: 520px; margin: .4rem 0; line-height: 1.15;}
    .putz-hero p {color: #e0eaff; max-width: 440px; margin: .45rem 0;
        font-size: 1.02rem;}
    .putz-eyebrow {color: #48c8ff; font-weight: 750; letter-spacing: .16em;
        font-size: .78rem;}
    .area-titulo {border-radius: 12px; padding: .9rem 1.2rem; margin: 1.2rem 0;
        color: white; font-size: 1.17rem; font-weight: 700;}
    .area-titulo.comparacao {background: linear-gradient(90deg,#065b98,#0b9ce6);}
    .area-titulo.bowtie {background: linear-gradient(90deg,#301367,#6738d2);}
    .stTabs [data-baseweb="tab-list"] {gap: .45rem;}
    .stTabs [data-baseweb="tab"] {border-radius: 9px 9px 0 0;
        background: #e9edf9; padding: .35rem .9rem;}
    .stTabs [aria-selected="true"] {color: #2133a0; font-weight: 700;
        background: #fff;}
    div[data-testid="stMetric"] {background: #fff; padding: .9rem;
        border: 1px solid #e0e6f5; border-radius: 12px;}
    @media(max-width: 700px) {
        .putz-hero {height: 220px; padding: 1.2rem;
            background-position: 66% 50%;}
        .putz-hero h1 {max-width: 270px;}
        .putz-hero p {max-width: 230px; font-size: .9rem;}
    }
    </style>
    <div class="putz-hero">
      <span class="putz-eyebrow">PUTZ TECNOLOGIA</span>
      <h1>Comparador de revisões de APR</h1>
      <p>Compare as revisões, encontre as mudanças e gere os BowTies dos cenários.</p>
    </div>
    """,
    unsafe_allow_html=True,
)
st.caption("Envie as duas APRs em PDF para começar. A comparação e os BowTies ficam em áreas separadas.")

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


def _cor_linha_alteracao(linha):
    tipo = str(linha.get("Tipo de alteração", "") or "").casefold()
    if "incluíd" in tipo or "novo" in tipo:
        cor = "#C6EFCE"
    elif "excluíd" in tipo or "removido" in tipo:
        cor = "#FFC7CE"
    else:
        cor = "#FFEB9C"
    return [f"background-color: {cor}"] * len(linha)


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
        "assinatura_excel",
        "arquivo_bowtie_individual",
        "arquivo_bowtie_zip",
        "assinatura_bowtie_individual",
        "assinatura_bowtie_zip",
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
        st.session_state.pop("assinatura_excel", None)
        st.session_state.pop("arquivo_bowtie_individual", None)
        st.session_state.pop("arquivo_bowtie_zip", None)
        st.session_state.pop("assinatura_bowtie_individual", None)
        st.session_state.pop("assinatura_bowtie_zip", None)
        st.session_state["assinaturas_apr_extraidas"] = (
            assinatura_pdf(apr_antiga), assinatura_pdf(apr_atualizada)
        )
        st.success("Extração concluída.")

    except Exception as erro:
        st.error("Não foi possível extrair os cenários dos documentos.")
        st.code(str(erro))


if (
    "resultado_antiga" in st.session_state
    and "resultado_atualizada" in st.session_state
    and arquivos_enviados
    and st.session_state.get("assinaturas_apr_extraidas")
    == (assinatura_pdf(apr_antiga), assinatura_pdf(apr_atualizada))
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

    with st.expander("Conferir cenários extraídos dos PDFs", expanded=False):
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
            st.session_state.pop("assinatura_excel", None)
            st.session_state.pop("arquivo_bowtie_individual", None)
            st.session_state.pop("arquivo_bowtie_zip", None)
            st.session_state.pop("assinatura_bowtie_individual", None)
            st.session_state.pop("assinatura_bowtie_zip", None)
            st.success("Comparação concluída.")
        except Exception as erro:
            st.error("Não foi possível comparar os cenários.")
            st.code(str(erro))

    if "resultado_comparacao" in st.session_state:
        comparacao = st.session_state["resultado_comparacao"]
        resumo = comparacao["resumo"]

        aba_comparacao, aba_bowties = st.tabs(
            ["📊 Planilha de comparação", "🟣 BowTies por cenário"]
        )

        with aba_comparacao:
            st.markdown('<div class="area-titulo comparacao">Comparação entre as APRs</div>', unsafe_allow_html=True)
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

            aba_correspondencias, aba_comparativo, aba_alteracoes = st.tabs(
                ["Correspondências", "Comparativo", "Alterações detectadas"]
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
                    alteracoes_filtradas.style.apply(_cor_linha_alteracao, axis=1),
                    use_container_width=True,
                    hide_index=True,
                    height=520,
                )

            st.divider()
            st.subheader("Gerar arquivo Excel")
            st.write(
                "O arquivo reúne o resumo, as duas APRs estruturadas, o comparativo, "
                "as alterações e as abas de enquadramento dos candidatos Bow Tie, "
                "seguindo o modelo oficial. Todas as planilhas são entregues no "
                "formato Excel (.xlsx)."
            )

            assinatura_excel = (
                assinatura_pdf(apr_antiga), assinatura_pdf(apr_atualizada)
            )
            if st.session_state.get("assinatura_excel") != assinatura_excel:
                try:
                    with st.spinner("Montando e formatando o arquivo Excel..."):
                        arquivo_excel = gerar_excel(antiga, atualizada, comparacao)
                    st.session_state["arquivo_excel"] = arquivo_excel
                    st.session_state["nome_excel"] = nome_arquivo_excel(
                        antiga, atualizada
                    )
                    st.session_state["assinatura_excel"] = assinatura_excel
                except Exception as erro:
                    st.session_state.pop("arquivo_excel", None)
                    st.error("Não foi possível montar o arquivo Excel.")
                    st.code(str(erro))

            if ("arquivo_excel" in st.session_state
                    and st.session_state.get("assinatura_excel") == assinatura_excel):
                st.download_button(
                    "Gerar e baixar comparação em Excel (.xlsx)",
                    data=st.session_state["arquivo_excel"],
                    file_name=st.session_state["nome_excel"],
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    type="primary",
                )


        with aba_bowties:
            st.markdown('<div class="area-titulo bowtie">BowTies por cenário</div>', unsafe_allow_html=True)
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
                f"Resultados marcados como REVISAR: "
                f"{resumo.get('CORRESPONDÊNCIAS A REVISAR', 0)}. "
                "Eles serão mantidos no Excel, com o motivo da revisão indicado."
            )

            aba_bowtie, aba_previa_bowtie = st.tabs(
                ["Candidatos BowTie", "Prévia BowTie"]
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
            st.subheader("Gerar ou atualizar BowTies por cenário")
            st.write(
                "Cada cenário da APR atualizada gera um arquivo .xlsx no layout "
                "do modelo NÓ_1. Ao atualizar um BowTie .xlsm existente, o arquivo "
                "gerado continua .xlsm e conserva suas macros."
            )
            incluir_todos = st.checkbox(
                "Incluir também cenários que não são candidatos a BowTie",
                value=False,
                key="incluir_todos_bowties",
            )
            cenarios_bowtie = [
                c for c in _ordenar_cenarios(atualizada["cenarios"])
                if incluir_todos or candidato_bowtie(c)
            ]
            if not cenarios_bowtie:
                st.warning("A APR atualizada não contém cenários para esta seleção.")
            else:
                opcoes = list(range(len(cenarios_bowtie)))
                indice = st.selectbox(
                    "Escolha o cenário para baixar ou atualizar",
                    opcoes,
                    format_func=lambda i: (
                        f"Nó {cenarios_bowtie[i].get('Nó', '')} | "
                        f"{cenarios_bowtie[i].get('ID', '')} | "
                        f"{str(cenarios_bowtie[i].get('Perigo', '')).replace(chr(10), ' ')}"
                    ),
                    key="cenario_para_excel_bowtie",
                )
                cenario_escolhido = cenarios_bowtie[indice]
                st.caption(
                    f"{len(cenarios_bowtie)} cenário(s) disponíveis. "
                    "Modos de detecção não entram no BowTie. "
                    "Os códigos SM2, SM7, SM8 e SM10 vão para TAG. "
                    "Dados técnicos que não aparecem nas APRs permanecem em branco. "
                    "Confira a classificação automática das barreiras antes de usar o arquivo."
                )
                bowtie_existente = st.file_uploader(
                    "Atualizar um BowTie com macros (opcional): envie o .xlsm do cenário selecionado",
                    type=["xlsm"],
                    key="bowtie_existente",
                    help="Ao enviar um arquivo, os dados preenchidos anteriormente são preservados. "
                         "O evento topo e o sistema devem corresponder ao cenário nas APRs.",
                )

                antigo_por_id = {c["ID"]: c for c in antiga["cenarios"]}
                pares = [
                    p for p in comparacao["correspondencias"]
                    if p.get("ID atualizada") == cenario_escolhido["ID"]
                    and p.get("ID antiga") in antigo_por_id
                ]
                par = pares[0] if len(pares) == 1 else None
                if bowtie_existente is not None:
                    if not par:
                        st.warning(
                            "Não foi localizado um único cenário correspondente na APR antiga. "
                            "A atualização deste BowTie precisa de conferência humana."
                        )
                    elif par.get("Validação") == "REVISAR":
                        st.warning(
                            "Correspondência marcada como REVISAR: "
                            + str(par.get("Motivo da revisão") or "confira o par de cenários")
                        )

                dados_existentes = bowtie_existente.getvalue() if bowtie_existente else None
                assinatura_selecao = (
                    assinatura_pdf(apr_antiga), assinatura_pdf(apr_atualizada),
                    cenario_escolhido["ID"],
                    hashlib.sha256(dados_existentes).hexdigest() if dados_existentes else "",
                )
                if st.session_state.get("assinatura_bowtie_individual") != assinatura_selecao:
                    try:
                        with st.spinner("Preenchendo o modelo Excel com a APR atualizada..."):
                            arquivo, alertas = gerar_bowtie(
                                cenario_escolhido,
                                atualizada,
                                modelo=dados_existentes,
                                cenario_antigo=antigo_por_id[par["ID antiga"]]
                                if bowtie_existente is not None and par else None,
                            )
                        st.session_state["arquivo_bowtie_individual"] = arquivo
                        st.session_state["nome_bowtie_individual"] = nome_bowtie(
                            cenario_escolhido, atualizado=bowtie_existente is not None
                        )
                        st.session_state["alertas_bowtie_individual"] = alertas
                        st.session_state.pop("erro_bowtie_individual", None)
                    except (ValueError, OSError, zipfile.BadZipFile) as erro:
                        st.session_state.pop("arquivo_bowtie_individual", None)
                        st.session_state["erro_bowtie_individual"] = str(erro)
                    st.session_state["assinatura_bowtie_individual"] = assinatura_selecao

                if st.session_state.get("erro_bowtie_individual"):
                    st.error(st.session_state["erro_bowtie_individual"])
                if (st.session_state.get("arquivo_bowtie_individual") is not None
                        and st.session_state.get("assinatura_bowtie_individual") == assinatura_selecao):
                    for aviso in st.session_state.get("alertas_bowtie_individual", []):
                        st.warning(aviso)
                    st.download_button(
                        "Gerar e baixar este BowTie (Excel)",
                        data=st.session_state["arquivo_bowtie_individual"],
                        file_name=st.session_state["nome_bowtie_individual"],
                        mime=("application/vnd.ms-excel.sheet.macroEnabled.12"
                              if bowtie_existente is not None
                              else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                        key="download_bowtie_individual",
                    )

                st.caption(
                    "O download de todos gera novos BowTies .xlsx a partir da APR atualizada. "
                    "Para atualizar um BowTie já preenchido, use o envio individual acima."
                )
                assinatura_zip = (assinatura_pdf(apr_atualizada), incluir_todos)
                if st.session_state.get("assinatura_bowtie_zip") != assinatura_zip:
                    try:
                        with st.spinner("Gerando um arquivo Excel para cada cenário..."):
                            memoria = io.BytesIO()
                            revisoes = []
                            with zipfile.ZipFile(memoria, "w", compression=zipfile.ZIP_DEFLATED) as pacote:
                                nomes_usados = set()
                                for c in cenarios_bowtie:
                                    arquivo, avisos = gerar_bowtie(c, atualizada)
                                    nome = nome_bowtie(c)
                                    if nome in nomes_usados:
                                        raise ValueError(
                                            f"Há mais de um cenário com o nome {nome}. "
                                            "Confira os identificadores antes de baixar todos."
                                        )
                                    nomes_usados.add(nome)
                                    pacote.writestr(nome, arquivo)
                                    if avisos:
                                        revisoes.append(f"{nome}: " + " | ".join(avisos))
                            st.session_state["arquivo_bowtie_zip"] = memoria.getvalue()
                            st.session_state["revisoes_bowtie_zip"] = revisoes
                        st.session_state.pop("erro_bowtie_zip", None)
                    except (ValueError, OSError, zipfile.BadZipFile) as erro:
                        st.session_state.pop("arquivo_bowtie_zip", None)
                        st.session_state["erro_bowtie_zip"] = str(erro)
                    st.session_state["assinatura_bowtie_zip"] = assinatura_zip

                if st.session_state.get("erro_bowtie_zip"):
                    st.error(st.session_state["erro_bowtie_zip"])
                if (st.session_state.get("arquivo_bowtie_zip") is not None
                        and st.session_state.get("assinatura_bowtie_zip")
                        == assinatura_zip):
                    st.download_button(
                        "Gerar e baixar todos os BowTies (.zip)",
                        data=st.session_state["arquivo_bowtie_zip"],
                        file_name="BowTies_APR_atualizada.zip",
                        mime="application/zip",
                        key="download_bowties_zip",
                    )
                    with st.expander("Avisos para conferência dos BowTies gerados"):
                        for aviso in st.session_state.get("revisoes_bowtie_zip", []):
                            st.write(aviso)
