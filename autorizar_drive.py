"""Executar SOMENTE no computador do administrador para obter um refresh token.

Requer credenciais_drive.json (OAuth Desktop) no mesmo diretório. Não envie
esse arquivo nem o token para o GitHub ou para outras pessoas.
"""

from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow


ARQUIVO = Path(__file__).resolve().parent / "credenciais_drive.json"
ESCOPO = ["https://www.googleapis.com/auth/drive"]


if __name__ == "__main__":
    if not ARQUIVO.exists():
        raise SystemExit("Coloque credenciais_drive.json nesta pasta antes de executar.")
    fluxo = InstalledAppFlow.from_client_secrets_file(str(ARQUIVO), ESCOPO)
    credenciais = fluxo.run_local_server(port=0, access_type="offline", prompt="consent")
    if not credenciais.refresh_token:
        raise SystemExit("O Google não forneceu refresh token; revogue a autorização e tente novamente.")
    print("\nCole este valor SOMENTE em [drive_oauth] nos Secrets do Streamlit:")
    print(credenciais.refresh_token)
    print("\nNão compartilhe nem publique esse valor.")
