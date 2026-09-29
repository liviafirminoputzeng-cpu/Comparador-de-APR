"""Área separada para criar e abrir planilhas compartilhadas no Google Drive."""

import base64
import hashlib
import re
import unicodedata
from pathlib import Path

import streamlit as st
from googleapiclient.errors import HttpError

from drive_acompanhamento import (
    cliente_drive, criar_ou_localizar, listar_planilhas, verificar_pasta,
)
from gerar_acompanhamento import candidatos_da_apr, gerar_planilha_acompanhamento


IMAGEM = Path(__file__).resolve().parent / "assets" / "putz_identidade.webp"


def _nome_seguro(texto):
    simples = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9_-]+", "_", simples).strip("_")[:70] or "Projeto"


def _avisar_erro(erro):
    if isinstance(erro, HttpError):
        codigo = getattr(erro.resp, "status", 0)
        if codigo in (403, 404):
            st.error("A conta configurada não consegue acessar ou gravar na pasta do Drive. "
                     "Confirme a permissão de editor e o tipo de pasta no guia.")
        else:
            st.error(f"O Google Drive recusou a operação (HTTP {codigo}).")
    elif isinstance(erro, ValueError):
        st.warning(str(erro))
    else:
        st.error("Não foi possível consultar ou criar a planilha no Google Drive.")


def mostrar_acompanhamento():
    capa = ("data:image/webp;base64," + base64.b64encode(IMAGEM.read_bytes()).decode()) \
        if IMAGEM.exists() else ""
    st.markdown("""<style>
    .stApp {background: #f3f5fc; color: #101632;}
    .block-container {max-width: 1500px; padding-top: 1.4rem;}
    .area-drive {padding: 1.6rem 2rem; border-radius: 16px;
      background: linear-gradient(90deg,rgba(7,8,30,.95),rgba(7,8,30,.75),rgba(7,8,30,.15)),
      url('""" + capa + """') center/cover #101936; margin: 1rem 0;}
    .area-drive h1 {color:#fff; margin:.2rem 0;}
    .area-drive p {color:#e0eaff; margin:0;}
    </style><div class="area-drive"><h1>Acompanhamento de BowTies</h1>
    <p>Crie a planilha do projeto e abra a cópia compartilhada no Google Drive da PUTZ.</p>
    </div>""", unsafe_allow_html=True)
    try:
        cfg = dict(st.secrets)
    except FileNotFoundError:
        cfg = {}
    if "auth" not in cfg or not any(k in cfg for k in ("drive_service_account", "drive_oauth")):
        st.info("O acesso ao Google Drive ainda não foi configurado no site. "
                "Consulte GUIA_GOOGLE_DRIVE.md. O Comparador de APR continua funcionando.")
        return
    if not st.user.is_logged_in:
        st.button("Entrar com a conta Google da equipe", on_click=st.login, type="primary")
        return
    email = str(st.user.get("email") or "").strip().casefold()
    equipe = {str(x).strip().casefold() for x in cfg.get("usuarios_equipe", [])}
    criadores = {str(x).strip().casefold() for x in cfg.get("usuarios_criadores", [])}
    if not email or st.user.get("email_verified") is False or email not in equipe | criadores:
        st.error("Sua conta Google ainda não foi autorizada para esta área.")
        return
    st.caption(f"Conta conectada: {email}")
    st.button("Sair da área de acompanhamento", on_click=st.logout)
    try:
        drive = cliente_drive(cfg)
        pasta = verificar_pasta(drive)
        if "drive_service_account" in cfg and not pasta.get("driveId"):
            st.error("Esta pasta está em Meu Drive. Use a configuração drive_oauth.")
            return
        arquivos = listar_planilhas(drive)
    except Exception as erro:
        _avisar_erro(erro)
        return
    st.subheader("Planilhas de acompanhamento existentes")
    if not arquivos:
        st.info("Ainda não há planilhas de acompanhamento nesta pasta do Drive.")
    for arquivo in arquivos:
        with st.container(border=True):
            titulo = arquivo.get("description") or arquivo["name"]
            coluna_titulo, coluna_abrir = st.columns([3, 1], vertical_alignment="center")
            coluna_titulo.write(f"**{titulo}**")
            coluna_titulo.caption(arquivo["name"])
            url = arquivo.get("webViewLink") or f'https://drive.google.com/file/d/{arquivo["id"]}/view'
            coluna_abrir.link_button("Abrir planilha online ↗", url, use_container_width=True)
    st.divider()
    st.subheader("Criar planilha a partir das APRs")
    if email not in criadores:
        st.caption("A criação de projetos está disponível para os coordenadores autorizados.")
        return
    if not pasta.get("capabilities", {}).get("canAddChildren", False):
        st.error("A conta configurada pode consultar a pasta, mas não criar planilhas nela. "
                 "Peça permissão de editor no Google Drive.")
        return
    antiga = st.session_state.get("resultado_antiga")
    atualizada = st.session_state.get("resultado_atualizada")
    comparacao = st.session_state.get("resultado_comparacao")
    assinaturas = st.session_state.get("assinaturas_apr_extraidas")
    if not (antiga and atualizada and comparacao and assinaturas):
        st.info("Na área Comparador de APR, envie as duas versões, faça a extração "
                "e clique em Comparar os cenários. Depois volte aqui para criar a planilha.")
        return
    st.caption(
        "Comparação carregada nesta sessão: "
        f"{antiga.get('nome_arquivo', 'APR antiga')} → "
        f"{atualizada.get('nome_arquivo', 'APR atualizada')}. "
        "Confira esses nomes antes de criar a planilha."
    )
    try:
        quantidade = len(candidatos_da_apr(atualizada["cenarios"]))
    except ValueError as erro:
        _avisar_erro(erro)
        return
    st.write(f"**{quantidade} candidatos a BowTie** encontrados na APR atualizada.")
    st.caption("ID, nó, cenário, sistema, evento topo e página vêm da APR atualizada. "
               "Responsável, tipo de trabalho (novo/correção), datas e andamento ficarão "
               "para a equipe preencher, após conferir a comparação.")
    projeto = st.text_input("Nome do projeto", value=atualizada.get("relatorio") or "",
                            placeholder="Ex.: UTE Vale do Açu · APR RevK")
    if st.button("Criar planilha de acompanhamento", type="primary", disabled=not projeto.strip()):
        try:
            # O mesmo projeto e o mesmo par de PDFs apontam para uma única cópia.
            origem = "|".join((projeto.strip().casefold(), *assinaturas)).encode("utf-8")
            identificador = hashlib.sha256(origem).hexdigest()[:12]
            nome = f"Acompanhamento_BowTies_{_nome_seguro(projeto)}_{identificador}.xlsx"
            existentes = [a for a in arquivos if a["name"] == nome]
            if existentes:
                arquivo, criado = existentes[0], False
            else:
                with st.spinner("Criando o modelo e salvando no Google Drive..."):
                    conteudo, _ = gerar_planilha_acompanhamento(atualizada, projeto.strip())
                    arquivo, criado = criar_ou_localizar(drive, nome, conteudo, projeto.strip())
            url = arquivo.get("webViewLink") or f'https://drive.google.com/file/d/{arquivo["id"]}/view'
            st.success("Planilha criada no Google Drive." if criado else
                       "Esta planilha já existia; seus dados foram preservados.")
            st.link_button("Abrir esta planilha online ↗", url, type="primary")
        except Exception as erro:
            _avisar_erro(erro)
