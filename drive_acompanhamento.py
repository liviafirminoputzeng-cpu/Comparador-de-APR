"""Envio e consulta de planilhas de acompanhamento na pasta autorizada do Drive."""

import io

from google.oauth2 import service_account
from google.oauth2.credentials import Credentials as UserCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload


ESCOPO = "https://www.googleapis.com/auth/drive"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PASTA_PADRAO = "1ichy39_gjGtfzALO66pe275YxY-HdRSh"


def cliente_drive(configuracao):
    """Conta de serviço para drive compartilhado ou conta OAuth para Meu Drive."""
    if "drive_service_account" in configuracao:
        info = dict(configuracao["drive_service_account"])
        credenciais = service_account.Credentials.from_service_account_info(
            info, scopes=[ESCOPO]
        )
    elif "drive_oauth" in configuracao:
        dados = configuracao["drive_oauth"]
        credenciais = UserCredentials(
            token=None,
            refresh_token=dados["refresh_token"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=dados["client_id"],
            client_secret=dados["client_secret"],
            scopes=[ESCOPO],
        )
    else:
        raise ValueError("Configure drive_service_account ou drive_oauth nos Secrets.")
    return build("drive", "v3", credentials=credenciais, cache_discovery=False)


def verificar_pasta(cliente, pasta_id=PASTA_PADRAO):
    pasta = cliente.files().get(
        fileId=pasta_id,
        fields="id,name,mimeType,driveId,capabilities(canAddChildren)",
        supportsAllDrives=True,
    ).execute()
    if pasta.get("mimeType") != "application/vnd.google-apps.folder":
        raise ValueError("O ID configurado não corresponde a uma pasta do Google Drive.")
    return pasta


def listar_planilhas(cliente, pasta_id=PASTA_PADRAO):
    arquivos = []
    pagina = None
    while True:
        resposta = cliente.files().list(
            q=(f"'{pasta_id}' in parents and trashed = false "
               f"and mimeType = '{MIME_XLSX}' "
               "and name contains 'Acompanhamento_BowTies_'"),
            fields="nextPageToken, files(id,name,webViewLink,createdTime,description)",
            pageSize=100, pageToken=pagina, supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        arquivos.extend(resposta.get("files", []))
        pagina = resposta.get("nextPageToken")
        if not pagina:
            break
    return sorted(arquivos, key=lambda a: a.get("createdTime", ""), reverse=True)


def criar_ou_localizar(cliente, nome, conteudo, projeto, pasta_id=PASTA_PADRAO):
    existentes = [f for f in listar_planilhas(cliente, pasta_id) if f["name"] == nome]
    if existentes:
        return existentes[0], False
    metadados = {
        "name": nome, "parents": [pasta_id], "mimeType": MIME_XLSX,
        "description": f"Acompanhamento de BowTies · {projeto}",
    }
    arquivo = cliente.files().create(
        body=metadados,
        media_body=MediaIoBaseUpload(io.BytesIO(conteudo), mimetype=MIME_XLSX,
                                     resumable=False),
        fields="id,name,webViewLink,createdTime,description",
        supportsAllDrives=True,
    ).execute()
    return arquivo, True
