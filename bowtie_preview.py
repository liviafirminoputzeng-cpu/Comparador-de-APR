import html
import re
from collections import defaultdict


ORDEM_SEVERIDADE = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5}


def _texto(valor):
    return html.escape(str(valor or "").strip())


def _itens_texto(texto):
    linhas = [linha.strip() for linha in str(texto or "").splitlines() if linha.strip()]
    if not linhas:
        return []

    itens = []
    inicio = re.compile(r"^(?:[-•]\s*|(?:SP|SM)\s*\d+\b)", re.IGNORECASE)
    for linha in linhas:
        if inicio.search(linha) or not itens:
            itens.append(re.sub(r"^[-•]\s*", "", linha).strip())
        else:
            itens[-1] = f"{itens[-1]} {linha}".strip()
    return list(dict.fromkeys(item for item in itens if item))


def _candidato_bowtie(cenario):
    return any(
        ORDEM_SEVERIDADE.get(str(cenario.get(campo, "")).strip().upper(), 0) >= 4
        for campo in ("Sev S", "Sev P", "Sev M")
    )


def _lista_html(itens, mensagem_vazia):
    if not itens:
        return f'<p class="vazio">{_texto(mensagem_vazia)}</p>'
    return "<ul>" + "".join(f"<li>{_texto(item)}</li>" for item in itens) + "</ul>"


def _barreiras_html(grupos, codigos, mensagem_vazia):
    blocos = []
    for codigo in codigos:
        for descricao in grupos.get(codigo, []):
            blocos.append(
                '<div class="barreira">'
                f'<span class="codigo">{_texto(codigo)}</span>'
                f'<span>{_texto(descricao)}</span>'
                "</div>"
            )
    if not blocos:
        return f'<p class="vazio">{_texto(mensagem_vazia)}</p>'
    return "".join(blocos)


def preparar_dados_bowtie(cenario, inventario_cenario):
    grupos = defaultdict(list)
    alertas = []
    for item in inventario_cenario:
        codigo = str(item.get("Código Bow Tie", "AV") or "AV")
        descricao = str(item.get("Descrição", "") or "").strip()
        if descricao and descricao not in grupos[codigo]:
            grupos[codigo].append(descricao)
        alerta = str(item.get("Alerta", "") or "").strip()
        if alerta:
            alertas.append(alerta)

    candidato = _candidato_bowtie(cenario)
    lacunas = list(dict.fromkeys(alertas))
    if candidato and not grupos.get("B1"):
        lacunas.append("B1: confirmar os elementos da contenção primária do trecho.")
    consequencias_texto = str(cenario.get("Possíveis efeitos", "") or "").lower()
    if (
        candidato
        and ("incênd" in consequencias_texto or "explos" in consequencias_texto)
        and not grupos.get("B7")
    ):
        lacunas.append("B7: confirmar os controles de fontes de ignição.")
    if candidato and not grupos.get("B10"):
        lacunas.append("B10: confirmar evacuação, resgate e abandono no PRE.")

    controles = grupos.get("CD", [])
    nao_barreiras = grupos.get("NB", [])
    if nao_barreiras:
        controles = controles + [f"Não barreira: {item}" for item in nao_barreiras]

    return {
        "id": cenario.get("ID", ""),
        "no": cenario.get("Nó", ""),
        "sistema": cenario.get("Sistema", ""),
        "evento_topo": cenario.get("Perigo", ""),
        "ameacas": _itens_texto(cenario.get("Causas", "")),
        "consequencias": _itens_texto(cenario.get("Possíveis efeitos", "")),
        "grupos": grupos,
        "controles": controles,
        "a_avaliar": grupos.get("AV", []),
        "lacunas": list(dict.fromkeys(lacunas)),
        "candidato": candidato,
        "frequencia": cenario.get("F", ""),
        "severidade": "/".join(
            str(cenario.get(f"Sev {dimensao}", "") or "") for dimensao in "SPMI"
        ),
        "risco": "/".join(
            str(cenario.get(f"Risco {dimensao}", "") or "") for dimensao in "SPMI"
        ),
    }


def montar_html_bowtie(cenario, inventario_cenario):
    dados = preparar_dados_bowtie(cenario, inventario_cenario)
    candidato = "SIM" if dados["candidato"] else "NÃO"
    classe_candidato = "sim" if dados["candidato"] else "nao"
    controles = _lista_html(
        dados["controles"], "Nenhum controle de degradação classificado."
    )
    avaliar = _lista_html(
        dados["a_avaliar"], "Nenhum item pendente de classificação automática."
    )
    lacunas = _lista_html(
        dados["lacunas"], "Nenhuma lacuna automática identificada."
    )

    return f"""
    <style>
      * {{ box-sizing: border-box; }}
      body {{ margin: 0; font-family: Arial, sans-serif; color: #172033; background: transparent; }}
      .painel {{ border: 1px solid #d9e1ec; border-radius: 12px; background: #f7f9fc; padding: 18px; }}
      .cabecalho {{ display: flex; justify-content: space-between; gap: 14px; align-items: flex-start; margin-bottom: 16px; }}
      .titulo {{ font-size: 18px; font-weight: 700; color: #17365d; margin-bottom: 5px; }}
      .meta {{ color: #5c667a; font-size: 13px; }}
      .selo {{ padding: 9px 13px; border-radius: 8px; font-weight: 700; white-space: nowrap; }}
      .selo.sim {{ background: #fce4c5; color: #8a4b00; }}
      .selo.nao {{ background: #e7e9ed; color: #4b5563; }}
      .bowtie {{ display: grid; grid-template-columns: 1.15fr 1fr .9fr 1fr 1.15fr; gap: 10px; align-items: stretch; }}
      .coluna {{ border-radius: 9px; padding: 12px; min-width: 0; }}
      .coluna h3 {{ font-size: 14px; margin: 0 0 10px; }}
      .ameacas, .consequencias {{ background: #fde8e7; color: #842f34; border: 1px solid #f3c4c2; }}
      .preventivas {{ background: #e5f0fc; color: #174f86; border: 1px solid #bed6ef; }}
      .mitigadoras {{ background: #e7f4ea; color: #265e3a; border: 1px solid #c4dfca; }}
      .evento {{ align-self: center; aspect-ratio: 1; border-radius: 50%; padding: 18px; display: grid; place-content: center; text-align: center; background: #fce4c5; color: #804800; border: 2px solid #e5a34e; }}
      .evento strong {{ display: block; margin-bottom: 7px; }}
      ul {{ margin: 0; padding-left: 18px; }}
      li {{ margin-bottom: 7px; font-size: 13px; line-height: 1.3; }}
      .barreira {{ display: flex; gap: 7px; align-items: flex-start; padding: 7px 0; border-bottom: 1px solid rgba(23,54,93,.13); font-size: 13px; line-height: 1.3; }}
      .barreira:last-child {{ border-bottom: 0; }}
      .codigo {{ flex: 0 0 auto; font-weight: 700; border-radius: 5px; padding: 2px 5px; background: rgba(255,255,255,.62); }}
      .vazio {{ margin: 0; color: #667085; font-size: 13px; font-style: italic; }}
      .apoio {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px; margin-top: 12px; }}
      .apoio section {{ border: 1px solid #d9e1ec; border-radius: 9px; background: #fff; padding: 12px; }}
      .apoio h3 {{ margin: 0 0 8px; font-size: 14px; color: #17365d; }}
      .rodape {{ margin-top: 12px; color: #667085; font-size: 12px; }}
      @media (max-width: 900px) {{
        .bowtie {{ grid-template-columns: 1fr; }}
        .evento {{ aspect-ratio: auto; border-radius: 9px; min-height: 120px; }}
        .apoio {{ grid-template-columns: 1fr; }}
        .cabecalho {{ flex-direction: column; }}
      }}
    </style>
    <div class="painel">
      <div class="cabecalho">
        <div>
          <div class="titulo">Nó {_texto(dados['no'])} — {_texto(dados['sistema'])}</div>
          <div class="meta">Cenário {_texto(dados['id'])} · F: {_texto(dados['frequencia'])} · Severidade S/P/M/I: {_texto(dados['severidade'])} · Risco S/P/M/I: {_texto(dados['risco'])}</div>
        </div>
        <div class="selo {classe_candidato}">Candidato Bow Tie: {candidato}</div>
      </div>
      <div class="bowtie">
        <section class="coluna ameacas"><h3>Ameaças</h3>{_lista_html(dados['ameacas'], 'Nenhuma ameaça extraída.')}</section>
        <section class="coluna preventivas"><h3>Barreiras preventivas</h3>{_barreiras_html(dados['grupos'], ('B1','B2','B3','B4','B5'), 'Nenhuma barreira preventiva classificada.')}</section>
        <section class="evento"><div><strong>Evento topo</strong>{_texto(dados['evento_topo'])}</div></section>
        <section class="coluna mitigadoras"><h3>Barreiras mitigadoras</h3>{_barreiras_html(dados['grupos'], ('B6','B7','B8','B9','B10'), 'Nenhuma barreira mitigadora classificada.')}</section>
        <section class="coluna consequencias"><h3>Consequências</h3>{_lista_html(dados['consequencias'], 'Nenhuma consequência extraída.')}</section>
      </div>
      <div class="apoio">
        <section><h3>Controles de degradação / não barreiras</h3>{controles}</section>
        <section><h3>Itens a avaliar</h3>{avaliar}</section>
        <section><h3>Lacunas para validação</h3>{lacunas}</section>
      </div>
      <div class="rodape">Prévia automática baseada na APR. O enquadramento das barreiras e as lacunas devem ser validados pela equipe técnica.</div>
    </div>
    """


def altura_previa_bowtie(cenario, inventario_cenario):
    dados = preparar_dados_bowtie(cenario, inventario_cenario)
    maior_coluna = max(
        len(dados["ameacas"]),
        len(dados["consequencias"]),
        sum(len(dados["grupos"].get(codigo, [])) for codigo in ("B1", "B2", "B3", "B4", "B5")),
        sum(len(dados["grupos"].get(codigo, [])) for codigo in ("B6", "B7", "B8", "B9", "B10")),
    )
    maior_apoio = max(
        len(dados["controles"]), len(dados["a_avaliar"]), len(dados["lacunas"])
    )
    return min(1050, max(620, 470 + 34 * maior_coluna + 22 * maior_apoio))
