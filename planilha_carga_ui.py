"""Área independente para gerar a planilha de carga de um BowTie enviado."""

import base64
import re
from pathlib import Path

import streamlit as st

from planilha_carga import MODELO, gerar_planilha_carga, gerar_planilhas_carga_zip


IMAGEM = Path(__file__).resolve().parent / "assets" / "putz_identidade.webp"


@st.cache_data(show_spinner=False, max_entries=2)
def _lote_processado(conteudo):
    return gerar_planilhas_carga_zip(conteudo)


def mostrar_planilha_carga():
    capa = ("data:image/webp;base64," + base64.b64encode(IMAGEM.read_bytes()).decode()) \
        if IMAGEM.exists() else ""
    st.markdown("""<style>
    .stApp {background: #f3f5fc; color: #101632;}
    .block-container {max-width: 1500px; padding-top: 1.4rem;}
    .area-carga {padding: 1.6rem 2rem; border-radius: 16px;
      background: linear-gradient(90deg,rgba(7,8,30,.95),rgba(7,8,30,.75),rgba(7,8,30,.15)),
      url('""" + capa + """') center/cover #101936; margin: 1rem 0;}
    .area-carga h1 {color:#fff; margin:.2rem 0;}
    .area-carga p {color:#e0eaff; margin:0;}
    </style><div class="area-carga"><h1>Planilha de carga</h1>
    <p>Envie um BowTie e gere um Excel com as abas Def e Data no formato dos modelos.</p>
    </div>""", unsafe_allow_html=True)
    if not MODELO.is_file():
        st.error("Falta o modelo de carga no site: assets/modelo_planilha_carga.xlsx. "
                 "Envie esse arquivo para a pasta assets no GitHub e aguarde a atualização do Streamlit.")
        return
    individual, lote = st.tabs(["Um BowTie", "Vários BowTies em ZIP"])
    with individual:
        arquivo = st.file_uploader("Selecione um BowTie em Excel", type=["xlsx", "xlsm"],
                                   key="bowtie_para_carga")
        if arquivo is not None:
            try:
                dados, resumo, avisos = gerar_planilha_carga(arquivo.getvalue())
            except ValueError as erro:
                st.error(str(erro))
            except Exception as erro:
                st.error("Não foi possível gerar a planilha de carga. "
                         f"Falha técnica: {type(erro).__name__}: {erro}")
            else:
                st.write(f"**Def:** {resumo['def']} registros · **Data:** {resumo['data']} registros")
                for aviso in avisos:
                    st.warning(aviso)
                nome = re.sub(r"[^A-Za-z0-9_-]", "_", Path(arquivo.name).stem).strip("_")[:90]
                st.download_button(
                    "Baixar planilha de carga (.xlsx)", data=dados,
                    file_name=f"Planilha_de_carga_{nome or 'BowTie'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary", key="baixar_carga_individual",
                )
        else:
            st.info("Envie um arquivo com a aba BT e, para as barreiras, abas B1, B2 etc.")
    with lote:
        pacote = st.file_uploader("Selecione um ZIP com BowTies .xlsx ou .xlsm",
                                  type=["zip"], key="bowties_zip_para_carga")
        if pacote is not None:
            try:
                with st.spinner("Preparando uma planilha de carga para cada BowTie..."):
                    dados_zip, resultados = _lote_processado(pacote.getvalue())
            except ValueError as erro:
                st.error(str(erro))
            except Exception as erro:
                st.error("Não foi possível processar o ZIP. "
                         f"Falha técnica: {type(erro).__name__}: {erro}")
            else:
                st.success(f"{len(resultados)} planilha(s) de carga preparada(s).")
                st.dataframe([
                    {"BowTie": item["origem"], "Planilha de carga": item["arquivo"],
                     "Registros Def": item["def"], "Registros Data": item["data"]}
                    for item in resultados
                ], hide_index=True, use_container_width=True)
                pendencias = [(item["origem"], aviso) for item in resultados for aviso in item["avisos"]]
                if pendencias:
                    with st.expander(f"Avisos para conferência ({len(pendencias)})"):
                        for origem, aviso in pendencias:
                            st.warning(f"{origem}: {aviso}")
                st.download_button(
                    "Baixar todas as planilhas de carga (.zip)", data=dados_zip,
                    file_name="Planilhas_de_carga_BowTies.zip", mime="application/zip",
                    type="primary", key="baixar_carga_lote",
                )
        else:
            st.info("Inclua os BowTies no ZIP. Se algum arquivo Excel estiver inválido, "
                    "o site indicará qual precisa ser corrigido antes do download do lote.")
    st.caption("Os arquivos de exemplo fornecem somente o formato. Campos ausentes no BowTie ficam vazios no Excel gerado.")
