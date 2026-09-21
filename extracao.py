import io
import re
from typing import BinaryIO

import pdfplumber


def _limpar(valor):
    if valor is None:
        return ""
    texto = str(valor).replace("\r", "\n")
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r" *\n *", "\n", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    return texto.strip()


def _juntar(anterior, novo):
    anterior = _limpar(anterior)
    novo = _limpar(novo)
    if not novo:
        return anterior
    if not anterior:
        return novo
    return f"{anterior}\n{novo}"


def _valor_linha(linha, campo):
    """Converte os dois formatos de grade observados nas APRs."""
    if len(linha) == 18:
        indices = {
            "perigo": 0,
            "causas": 1,
            "efeitos": 2,
            "protecao": 4,
            "r": 5,
            "frequencia": 6,
            "sev_s": 7,
            "sev_p": 8,
            "sev_m": 9,
            "sev_i": 10,
            "risco_s": 12,
            "risco_p": 13,
            "risco_m": 14,
            "risco_i": 15,
            "observacoes": 16,
            "cenario": 17,
        }
    elif len(linha) == 17:
        indices = {
            "perigo": 0,
            "causas": 1,
            "efeitos": 2,
            "protecao": 3,
            "r": 5,
            "frequencia": 6,
            "sev_s": 7,
            "sev_p": 8,
            "sev_m": 9,
            "sev_i": 10,
            "risco_s": 11,
            "risco_p": 12,
            "risco_m": 13,
            "risco_i": 14,
            "observacoes": 15,
            "cenario": 16,
        }
    else:
        return ""

    return _limpar(linha[indices[campo]])


def _separar_protecoes(texto):
    texto = _limpar(texto)
    if not texto:
        return "", "", "", ""

    # Alguns PDFs trazem aspas soltas antes do código da salvaguarda
    # (por exemplo: 'SP61). Elas não fazem parte do conteúdo e impedem
    # a identificação correta do início do item.
    texto = re.sub(
        r"(?im)^\s*['\"“”‘’]+\s*(?=(?:SP|SM)\s*\d+\b)",
        "",
        texto,
    )

    padrao = re.compile(
        r"(?im)^\s*(MODO\s+DE\s+DETEC[CÇ][AÃ]O|"
        r"SALVAGUARDAS?\s+PREVENTIVAS?|"
        r"SALVAGUARDAS?\s+MITIGADORAS?|"
        r"SALVAGUARDAS?)\s*$"
    )
    marcadores = list(padrao.finditer(texto))

    if not marcadores:
        return "", "", "", texto

    secoes = {
        "deteccao": "",
        "preventivas": "",
        "mitigadoras": "",
        "sem_tipo": _limpar(texto[: marcadores[0].start()]),
    }

    for posicao, marcador in enumerate(marcadores):
        inicio = marcador.end()
        fim = (
            marcadores[posicao + 1].start()
            if posicao + 1 < len(marcadores)
            else len(texto)
        )
        conteudo = _limpar(texto[inicio:fim])
        titulo = marcador.group(1).upper()

        if "DETEC" in titulo:
            chave = "deteccao"
        elif "PREVENTIVA" in titulo:
            chave = "preventivas"
        elif "MITIGADORA" in titulo:
            chave = "mitigadoras"
        else:
            # Há cenários, como N6.1 e N6.2 da Rev. K, em que a APR usa
            # apenas o título genérico "SALVAGUARDA". Não é seguro
            # presumir que o item seja preventivo ou mitigador.
            chave = "sem_tipo"

        secoes[chave] = _juntar(secoes[chave], conteudo)

    return (
        secoes["deteccao"],
        secoes["preventivas"],
        secoes["mitigadoras"],
        secoes["sem_tipo"],
    )


def extrair_cenarios_pdf(arquivo: bytes | BinaryIO):
    if isinstance(arquivo, bytes):
        arquivo = io.BytesIO(arquivo)

    cenarios = []
    atual = None
    no_atual = None
    sistema_atual = ""
    trecho_atual = ""
    relatorio = ""
    aprovacao = ""

    with pdfplumber.open(arquivo) as pdf:
        total_paginas = len(pdf.pages)

        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            texto_pagina = pagina.extract_text() or ""

            resultado_no = re.search(
                r"N[oó]\s+(\d+)\s*-\s*(.*?)\s+Situa[cç][aã]o\s+do\s+Estudo",
                texto_pagina,
                flags=re.IGNORECASE | re.DOTALL,
            )
            if resultado_no:
                no_atual = int(resultado_no.group(1))
                sistema_atual = _limpar(resultado_no.group(2)).replace("\n", " ")

            tabelas = pagina.find_tables()
            for tabela in tabelas:
                dados = tabela.extract()

                for linha in dados:
                    texto_linha = " ".join(
                        _limpar(celula).replace("\n", " ")
                        for celula in linha
                        if _limpar(celula)
                    )

                    if "Relatório:" in texto_linha:
                        achado = re.search(
                            r"Relat[oó]rio:\s*(.*?)(?:Aprova[cç][aã]o:|$)",
                            texto_linha,
                            flags=re.IGNORECASE,
                        )
                        if achado:
                            relatorio = _limpar(achado.group(1))

                    if "Aprovação:" in texto_linha:
                        achado = re.search(
                            r"Aprova[cç][aã]o:\s*(\d{2}/\d{2}/\d{4})",
                            texto_linha,
                            flags=re.IGNORECASE,
                        )
                        if achado:
                            aprovacao = achado.group(1)

                    if "Trecho de Análise:" in texto_linha:
                        trecho_atual = re.sub(
                            r"^.*?Trecho\s+de\s+An[aá]lise:\s*",
                            "",
                            texto_linha,
                            flags=re.IGNORECASE,
                        )
                        trecho_atual = re.split(
                            r"Documentos\s+de\s+Refer[eê]ncia:",
                            trecho_atual,
                            maxsplit=1,
                            flags=re.IGNORECASE,
                        )[0].strip()

                    if len(linha) not in (17, 18):
                        continue

                    texto_cabecalho = texto_linha.casefold()
                    marcadores_cabecalho = (
                        "segurança, meio ambiente e saúde",
                        "data de ",
                        "atividade:",
                        "processo:",
                        "análise preliminar de riscos",
                        "situação do estudo:",
                        "trecho de análise:",
                        "documentos de referência:",
                        "categoria de severidade",
                        "categoria de risco",
                    )
                    if (
                        _valor_linha(linha, "perigo").casefold() == "perigo"
                        or any(
                            marcador in texto_cabecalho
                            for marcador in marcadores_cabecalho
                        )
                    ):
                        continue

                    numero_cenario = _valor_linha(linha, "cenario")
                    inicia_cenario = numero_cenario.isdigit()

                    if inicia_cenario:
                        atual = {
                            "ID": f"N{no_atual}.{numero_cenario}",
                            "Nó": no_atual,
                            "Sistema": sistema_atual,
                            "Trecho de análise": trecho_atual,
                            "Cenário": int(numero_cenario),
                            "Perigo": _valor_linha(linha, "perigo"),
                            "Causas": _valor_linha(linha, "causas"),
                            "Possíveis efeitos": _valor_linha(linha, "efeitos"),
                            "Proteções (texto original)": _valor_linha(
                                linha, "protecao"
                            ),
                            "F": _valor_linha(linha, "frequencia"),
                            "Sev S": _valor_linha(linha, "sev_s"),
                            "Sev P": _valor_linha(linha, "sev_p"),
                            "Sev M": _valor_linha(linha, "sev_m"),
                            "Sev I": _valor_linha(linha, "sev_i"),
                            "Risco S": _valor_linha(linha, "risco_s"),
                            "Risco P": _valor_linha(linha, "risco_p"),
                            "Risco M": _valor_linha(linha, "risco_m"),
                            "Risco I": _valor_linha(linha, "risco_i"),
                            "Observações / Recomendações": _valor_linha(
                                linha, "observacoes"
                            ),
                            "Página PDF": numero_pagina,
                            "Relatório": relatorio,
                            "Aprovação": aprovacao,
                        }
                        cenarios.append(atual)
                        continue

                    if atual is None:
                        continue

                    r = _valor_linha(linha, "r").upper()
                    if r in {"S", "C"}:
                        continue

                    campos_continuacao = {
                        "Perigo": "perigo",
                        "Causas": "causas",
                        "Possíveis efeitos": "efeitos",
                        "Proteções (texto original)": "protecao",
                        "Observações / Recomendações": "observacoes",
                    }
                    for destino, origem in campos_continuacao.items():
                        atual[destino] = _juntar(
                            atual[destino], _valor_linha(linha, origem)
                        )

    for cenario in cenarios:
        deteccao, preventivas, mitigadoras, sem_tipo = _separar_protecoes(
            cenario["Proteções (texto original)"]
        )
        cenario["Modo de detecção"] = deteccao
        cenario["Salvaguardas preventivas (SP)"] = preventivas
        cenario["Salvaguardas mitigadoras (SM)"] = mitigadoras
        cenario["Salvaguardas sem tipo"] = sem_tipo

    return {
        "cenarios": cenarios,
        "total_paginas": total_paginas,
        "relatorio": relatorio,
        "aprovacao": aprovacao,
    }
