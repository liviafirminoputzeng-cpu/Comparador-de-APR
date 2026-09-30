"""Gera Excel local e organiza links sem credenciais ou gravação no servidor."""

import base64
import hashlib
import json
import re
import unicodedata
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import streamlit as st

from gerar_acompanhamento import candidatos_da_apr, gerar_planilha_acompanhamento


RAIZ = Path(__file__).resolve().parent
IMAGEM = RAIZ / "assets" / "putz_identidade.webp"
CATALOGO = RAIZ / "links_acompanhamento.json"


def _nome_seguro(texto):
    simples = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9_-]+", "_", simples).strip("_")[:70] or "Projeto"


def _link_planilha(texto):
    """Aceita somente links HTTPS de planilhas/arquivos do Google Drive."""
    url = str(texto or "").strip()
    partes = urlparse(url)
    if partes.scheme != "https" or partes.hostname not in {"drive.google.com", "docs.google.com"}:
        raise ValueError("Cole um link HTTPS do Google Drive ou do Planilhas Google.")
    arquivo_drive = partes.hostname == "drive.google.com" and (
        partes.path.startswith("/file/d/") or
        (partes.path == "/open" and bool(parse_qs(partes.query).get("id")))
    )
    arquivo_planilhas = partes.hostname == "docs.google.com" and partes.path.startswith(
        "/spreadsheets/d/"
    )
    if not (arquivo_drive or arquivo_planilhas):
        raise ValueError("Cole o link do arquivo Excel, não o link da pasta do Drive.")
    return url


def carregar_catalogo():
    """Lê apenas os links publicados no repositório do site."""
    if not CATALOGO.is_file():
        return []
    dados = json.loads(CATALOGO.read_text(encoding="utf-8"))
    projetos = dados.get("projetos", [])
    if not isinstance(projetos, list):
        raise ValueError("O catálogo de links precisa conter uma lista 'projetos'.")
    return [
        {"projeto": str(item["projeto"]), "url": _link_planilha(item["url"])}
        for item in projetos
        if isinstance(item, dict) and item.get("projeto") and item.get("url")
    ]


def preparar_catalogo(projetos, projeto, url):
    """Produz bytes JSON, sem gravar no disco do Streamlit."""
    nome = " ".join(str(projeto).split())
    if not nome:
        raise ValueError("Informe o nome do projeto antes de registrar o link.")
    if len(nome) > 120:
        raise ValueError("Use um nome de projeto com até 120 caracteres.")
    url = _link_planilha(url)
    atualizados = [item.copy() for item in projetos]
    for indice, item in enumerate(atualizados):
        if item["projeto"].casefold() == nome.casefold():
            atualizados[indice] = {"projeto": nome, "url": url}
            break
    else:
        atualizados.append({"projeto": nome, "url": url})
    return serializar_catalogo(atualizados)


def serializar_catalogo(projetos):
    """Preserva a ordem exibida no site ao preparar o catálogo do GitHub."""
    return json.dumps({"projetos": projetos}, ensure_ascii=False, indent=2).encode("utf-8") + b"\n"


def mover_link(projetos, indice, deslocamento):
    """Retorna uma nova lista com um link movido uma posição."""
    destino = indice + deslocamento
    if not 0 <= indice < len(projetos) or not 0 <= destino < len(projetos):
        raise ValueError("Não é possível mover o link além dos limites da lista.")
    atualizados = [item.copy() for item in projetos]
    atualizados[indice], atualizados[destino] = atualizados[destino], atualizados[indice]
    return atualizados


def excluir_link(projetos, indice):
    """Remove um link apenas da prévia desta sessão."""
    if not 0 <= indice < len(projetos):
        raise ValueError("Não foi possível encontrar o link selecionado.")
    return [item.copy() for posicao, item in enumerate(projetos) if posicao != indice]


def _guardar_previa(publicados, pendentes):
    st.session_state["links_pendentes"] = pendentes
    if pendentes == publicados:
        st.session_state.pop("catalogo_preparado", None)
    else:
        st.session_state["catalogo_preparado"] = serializar_catalogo(pendentes)


def mostrar_acompanhamento():
    capa = ("data:image/webp;base64," + base64.b64encode(IMAGEM.read_bytes()).decode()) \
        if IMAGEM.exists() else ""
    st.markdown("""<style>
    .stApp {background: #f3f5fc; color: #101632;}
    .block-container {max-width: 1500px; padding-top: 1.4rem;}
    .area-acompanhamento {padding: 1.6rem 2rem; border-radius: 16px;
      background: linear-gradient(90deg,rgba(7,8,30,.95),rgba(7,8,30,.75),rgba(7,8,30,.15)),
      url('""" + capa + """') center/cover #101936; margin: 1rem 0;}
    .area-acompanhamento h1 {color:#fff; margin:.2rem 0;}
    .area-acompanhamento p {color:#e0eaff; margin:0;}
    </style><div class="area-acompanhamento"><h1>Acompanhamento de BowTies</h1>
    <p>Gere o Excel do projeto e consulte os links das planilhas compartilhadas.</p>
    </div>""", unsafe_allow_html=True)

    aba_excel, aba_links = st.tabs(["📊 Gerar planilha Excel", "🔗 Planilhas compartilhadas"])
    with aba_excel:
        antiga = st.session_state.get("resultado_antiga")
        atualizada = st.session_state.get("resultado_atualizada")
        comparacao = st.session_state.get("resultado_comparacao")
        assinaturas = st.session_state.get("assinaturas_apr_extraidas")
        st.subheader("Planilha de acompanhamento do projeto")
        if not (antiga and atualizada and comparacao and assinaturas):
            st.info("Em Comparador de APR, envie as duas versões, clique em Extrair cenários "
                    "e em Comparar os cenários. Depois volte aqui para baixar a planilha.")
        else:
            st.caption("APRs comparadas nesta sessão: "
                       f"{antiga.get('nome_arquivo', 'APR antiga')} → "
                       f"{atualizada.get('nome_arquivo', 'APR atualizada')}.")
            projeto = st.text_input("Nome do projeto", value=atualizada.get("relatorio") or "",
                                    placeholder="Ex.: UTE Vale do Açu · APR RevK")
            if projeto.strip():
                try:
                    conteudo, quantidade = gerar_planilha_acompanhamento(atualizada, projeto.strip())
                    identificador = hashlib.sha256(
                        "|".join((projeto.strip().casefold(), *assinaturas)).encode("utf-8")
                    ).hexdigest()[:12]
                    nome = f"Acompanhamento_BowTies_{_nome_seguro(projeto)}_{identificador}.xlsx"
                    st.write(f"**{quantidade} candidatos a BowTie** da APR atualizada, "
                             "ordenados por nó e cenário.")
                    st.download_button("Baixar planilha de acompanhamento (.xlsx)",
                                       data=conteudo, file_name=nome,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       type="primary")
                    st.caption("Após baixar, envie o Excel à pasta compartilhada da equipe no Drive. "
                               "Em seguida, copie o link do arquivo e registre-o na aba ao lado. "
                               "Na aba Controle, classifique os novos BowTies, a etapa da planilha "
                               "e o estagiário responsável pela produção.")
                except ValueError as erro:
                    st.warning(str(erro))
                except Exception:
                    st.error("Não foi possível gerar o Excel de acompanhamento. Confira a extração das APRs.")

    with aba_links:
        st.subheader("Links das planilhas no Google Drive")
        try:
            publicados = carregar_catalogo()
        except (ValueError, OSError, json.JSONDecodeError, KeyError) as erro:
            st.error("O arquivo links_acompanhamento.json está inválido. "
                     "Corrija-o no repositório antes de cadastrar outro link.")
            return
        # Uma publicação nova passa a ser a base da sessão; um rascunho antigo
        # nunca deve sobrescrever links que chegaram depois pelo GitHub.
        assinatura_publicada = hashlib.sha256(
            CATALOGO.read_bytes() if CATALOGO.is_file() else b""
        ).hexdigest()
        if st.session_state.get("catalogo_publicado_hash") != assinatura_publicada:
            for chave in ("catalogo_preparado", "links_pendentes",
                          "projeto_preparado", "url_preparada"):
                st.session_state.pop(chave, None)
            st.session_state["catalogo_publicado_hash"] = assinatura_publicada
        if "links_pendentes" not in st.session_state:
            st.session_state["links_pendentes"] = [item.copy() for item in publicados]
        pendentes = st.session_state["links_pendentes"]
        if pendentes != publicados:
            st.info("Prévia nesta sessão: a nova ordem e as exclusões só aparecerão para "
                    "toda a equipe depois que você publicar o JSON no GitHub.")
        if pendentes:
            for indice, item in enumerate(pendentes):
                with st.container(border=True):
                    coluna_nome, coluna_link, subir, descer, excluir = st.columns(
                        [3, 1.4, .55, .55, .8], vertical_alignment="center"
                    )
                    coluna_nome.write(f"**{item['projeto']}**")
                    coluna_link.link_button("Abrir planilha ↗", item["url"], use_container_width=True)
                    if subir.button("↑", key=f"link_subir_{indice}",
                                    help=f"Mover {item['projeto']} para cima",
                                    disabled=indice == 0, use_container_width=True):
                        _guardar_previa(publicados, mover_link(pendentes, indice, -1))
                        st.rerun()
                    if descer.button("↓", key=f"link_descer_{indice}",
                                     help=f"Mover {item['projeto']} para baixo",
                                     disabled=indice == len(pendentes) - 1,
                                     use_container_width=True):
                        _guardar_previa(publicados, mover_link(pendentes, indice, 1))
                        st.rerun()
                    if excluir.button("Excluir", key=f"link_excluir_{indice}",
                                      help=f"Retirar {item['projeto']} da lista preparada",
                                      use_container_width=True):
                        _guardar_previa(publicados, excluir_link(pendentes, indice))
                        st.rerun()
        else:
            st.info("Nenhum link nesta lista. Se você retirou o último link, baixe e "
                    "publique o catálogo atualizado para removê-lo do site da equipe.")

        if pendentes != publicados and st.button("Desfazer alterações desta sessão"):
            _guardar_previa(publicados, [item.copy() for item in publicados])
            st.rerun()

        st.markdown("#### Colocar o link de uma planilha")
        st.caption("Depois de enviar o Excel ao Drive, abra o arquivo, copie seu link "
                   "e cole abaixo. O acesso de Editor é configurado no próprio Drive.")
        with st.form("formulario_link_acompanhamento"):
            projeto_link = st.text_input("Projeto da planilha compartilhada")
            url_link = st.text_input("Link do arquivo no Google Drive")
            preparar = st.form_submit_button("Preparar inclusão do link")
        if preparar:
            try:
                novo_catalogo = preparar_catalogo(pendentes, projeto_link, url_link)
                _guardar_previa(publicados, json.loads(novo_catalogo)["projetos"])
                st.rerun()
            except ValueError as erro:
                st.error(str(erro))
        if st.session_state.get("catalogo_preparado"):
            st.success("Alterações preparadas. Confira a lista acima antes de publicar.")
            st.download_button("Baixar catálogo atualizado para publicar no site (.json)",
                               st.session_state["catalogo_preparado"],
                               file_name="links_acompanhamento.json", mime="application/json")
            st.info("O arquivo baixado contém exatamente os links da prévia, na mesma ordem. "
                    "Substitua links_acompanhamento.json na página principal do "
                    "repositório GitHub e faça Commit changes. Até lá, ninguém mais verá "
                    "a nova ordem ou as exclusões.")
        st.caption("Se o repositório ou site for público, nomes de projetos e links "
                   "publicados também ficarão visíveis. Mantenha os arquivos do Drive "
                   "restritos à equipe.")
