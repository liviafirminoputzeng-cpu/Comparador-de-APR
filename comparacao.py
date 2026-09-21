import re
import unicodedata
from collections import Counter, defaultdict
from difflib import SequenceMatcher
from functools import lru_cache


@lru_cache(maxsize=50000)
def _normalizar(texto):
    texto = unicodedata.normalize("NFKD", texto or "")
    texto = texto.encode("ascii", "ignore").decode("ascii").lower()
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def _chave_id(valor):
    """Ordena IDs como N2.9 antes de N10.1, sem depender deles no pareamento."""
    achado = re.search(r"N\s*(\d+)\s*[.]\s*(\d+)", str(valor or ""), re.I)
    if achado:
        return int(achado.group(1)), int(achado.group(2)), str(valor or "")
    numeros = [int(numero) for numero in re.findall(r"\d+", str(valor or ""))]
    return (
        numeros[0] if numeros else 10**9,
        numeros[1] if len(numeros) > 1 else 10**9,
        str(valor or ""),
    )


@lru_cache(maxsize=50000)
def _similaridade_texto(texto_a, texto_b):
    a = _normalizar(texto_a)
    b = _normalizar(texto_b)
    if not a or not b:
        return 0.0

    sequencia = SequenceMatcher(None, a, b, autojunk=False).ratio()
    tokens_a = set(a.split())
    tokens_b = set(b.split())
    jaccard = len(tokens_a & tokens_b) / len(tokens_a | tokens_b)
    return max(sequencia, jaccard)


PALAVRAS_GENERICAS = {
    "pequeno",
    "pequena",
    "grande",
    "vazamento",
    "liberacao",
    "produtos",
    "produto",
    "quimicos",
    "quimico",
    "perda",
    "sistema",
    "causas",
    "naturais",
    "agua",
    "gas",
    "linha",
    "pressurizado",
    "pressurizada",
    "integridade",
    "estrutura",
}


def _termos_especificos(texto):
    return {
        termo
        for termo in _normalizar(texto).split()
        if len(termo) > 3 and termo not in PALAVRAS_GENERICAS
    }


def _similaridade_codigos(texto_a, texto_b):
    padrao = r"\b[a-z]{1,8}[- ]?\d+[a-z0-9/-]*\b"
    codigos_a = set(re.findall(padrao, _normalizar(texto_a)))
    codigos_b = set(re.findall(padrao, _normalizar(texto_b)))
    if not codigos_a or not codigos_b:
        return 0.0
    return len(codigos_a & codigos_b) / len(codigos_a | codigos_b)


def _texto_completo(cenario):
    campos = (
        "Perigo",
        "Causas",
        "Possíveis efeitos",
        "Modo de detecção",
        "Salvaguardas preventivas (SP)",
        "Salvaguardas mitigadoras (SM)",
    )
    return " ".join(str(cenario.get(campo, "")) for campo in campos)


def _pontuar(antigo, atualizado):
    sistema = _similaridade_texto(antigo["Sistema"], atualizado["Sistema"])
    trecho = _similaridade_texto(
        antigo["Trecho de análise"], atualizado["Trecho de análise"]
    )
    perigo = _similaridade_texto(antigo["Perigo"], atualizado["Perigo"])
    causas = _similaridade_texto(antigo["Causas"], atualizado["Causas"])
    efeitos = _similaridade_texto(
        antigo["Possíveis efeitos"], atualizado["Possíveis efeitos"]
    )
    codigos = _similaridade_codigos(
        _texto_completo(antigo), _texto_completo(atualizado)
    )

    # O nome do desvio/perigo é a referência principal. Nó, número do
    # cenário e ID não participam da pontuação, pois podem mudar quando um
    # nó é incluído, removido ou renumerado entre revisões.
    pontuacao = (
        0.08 * sistema
        + 0.03 * trecho
        + 0.58 * perigo
        + 0.18 * causas
        + 0.10 * efeitos
        + 0.03 * codigos
    )

    if _normalizar(antigo["Perigo"]) == _normalizar(atualizado["Perigo"]):
        pontuacao += 0.08

    perigo_antigo = _normalizar(antigo["Perigo"])
    perigo_atualizado = _normalizar(atualizado["Perigo"])
    pequeno_antigo = "pequen" in perigo_antigo
    pequeno_atualizado = "pequen" in perigo_atualizado
    grande_antigo = "grand" in perigo_antigo
    grande_atualizado = "grand" in perigo_atualizado

    if (pequeno_antigo and pequeno_atualizado) or (
        grande_antigo and grande_atualizado
    ):
        pontuacao += 0.08
    if (pequeno_antigo and grande_atualizado) or (
        grande_antigo and pequeno_atualizado
    ):
        pontuacao -= 0.20

    # Quando a revisão antiga não informa o porte do vazamento, uma nova
    # descrição explicitamente "grande" não deve ganhar só por acrescentar
    # essa palavra. Isso ajuda a preservar o desvio equivalente pelo nome e
    # pelo conteúdo, sem recorrer ao número do cenário.
    if not pequeno_antigo and not grande_antigo and grande_atualizado:
        pontuacao -= 0.04

    return max(0.0, min(1.0, pontuacao))


def _candidato(
    antigo,
    atualizado,
):
    sistema_antigo = _normalizar(antigo["Sistema"])
    sistema_atualizado = _normalizar(atualizado["Sistema"])
    mesmo_sistema = sistema_antigo == sistema_atualizado
    perigo = _similaridade_texto(antigo["Perigo"], atualizado["Perigo"])
    causas = _similaridade_texto(antigo["Causas"], atualizado["Causas"])
    efeitos = _similaridade_texto(
        antigo["Possíveis efeitos"], atualizado["Possíveis efeitos"]
    )

    nome_igual = _normalizar(antigo["Perigo"]) == _normalizar(
        atualizado["Perigo"]
    )
    termos_comuns = _termos_especificos(antigo["Perigo"]) & _termos_especificos(
        atualizado["Perigo"]
    )

    if mesmo_sistema:
        if nome_igual:
            return True
        if perigo >= 0.52 and termos_comuns:
            return True
        # Permite uma renomeação real quando causas e consequências continuam
        # muito próximas. Um limite mínimo para o nome evita relacionar, por
        # exemplo, "pressão elevada" com "vazamento de lodo" apenas porque os
        # textos auxiliares são iguais.
        if perigo >= 0.55 and causas >= 0.80 and efeitos >= 0.70:
            return True
        if perigo >= 0.50 and causas >= 0.68 and efeitos >= 0.80:
            return True
        # Exceção conservadora para uma troca ampla de nome: o trecho, as
        # causas e as consequências precisam permanecer praticamente iguais.
        trecho = _similaridade_texto(
            antigo["Trecho de análise"], atualizado["Trecho de análise"]
        )
        return trecho >= 0.93 and causas >= 0.80 and efeitos >= 0.80

    migracao_especifica = bool(
        termos_comuns & {"diesel", "combustivel", "hidrogenio", "sulfurico"}
    )
    return perigo >= 0.80 and migracao_especifica and (
        causas >= 0.25 or efeitos >= 0.35
    )


def _confianca(pontuacao):
    if pontuacao >= 0.85:
        return "Alta"
    if pontuacao >= 0.70:
        return "Média"
    return "Baixa"


def _encontrar_correspondencias(antigos, atualizados):
    possibilidades = []
    for antigo in antigos:
        for atualizado in atualizados:
            if not _candidato(antigo, atualizado):
                continue
            pontuacao = _pontuar(antigo, atualizado)
            mesma_area = (
                _normalizar(antigo["Sistema"])
                == _normalizar(atualizado["Sistema"])
            )
            renomeacao_forte = (
                mesma_area
                and _similaridade_texto(
                    antigo["Trecho de análise"], atualizado["Trecho de análise"]
                )
                >= 0.93
                and _similaridade_texto(antigo["Causas"], atualizado["Causas"])
                >= 0.80
                and _similaridade_texto(
                    antigo["Possíveis efeitos"], atualizado["Possíveis efeitos"]
                )
                >= 0.80
            )
            if pontuacao >= 0.60 or renomeacao_forte:
                possibilidades.append((pontuacao, antigo, atualizado))

    usados_antigos = set()
    usados_atualizados = set()
    pares = []

    for pontuacao, antigo, atualizado in sorted(
        possibilidades, key=lambda item: item[0], reverse=True
    ):
        if antigo["ID"] in usados_antigos or atualizado["ID"] in usados_atualizados:
            continue
        nome_igual = _normalizar(antigo["Perigo"]) == _normalizar(
            atualizado["Perigo"]
        )
        origem = (
            "Nome do desvio igual"
            if nome_igual
            else "Nome do desvio semelhante / renomeado"
        )
        pares.append((antigo, atualizado, pontuacao, origem))
        usados_antigos.add(antigo["ID"])
        usados_atualizados.add(atualizado["ID"])

    # Permite unificação: dois cenários antigos podem convergir em um atual.
    for antigo in antigos:
        if antigo["ID"] in usados_antigos:
            continue
        opcoes = []
        for atualizado in atualizados:
            if atualizado["ID"] not in usados_atualizados:
                continue
            if _normalizar(antigo["Sistema"]) != _normalizar(atualizado["Sistema"]):
                continue
            perigo = _similaridade_texto(antigo["Perigo"], atualizado["Perigo"])
            termos = _termos_especificos(antigo["Perigo"]) & _termos_especificos(
                atualizado["Perigo"]
            )
            pontuacao = _pontuar(antigo, atualizado)
            if pontuacao >= 0.73 and perigo >= 0.65 and termos:
                opcoes.append((pontuacao, atualizado))

        if opcoes:
            pontuacao, atualizado = max(opcoes, key=lambda item: item[0])
            pares.append((antigo, atualizado, pontuacao, "Possível unificação"))
            usados_antigos.add(antigo["ID"])

    removidos = [c for c in antigos if c["ID"] not in usados_antigos]
    novos = [c for c in atualizados if c["ID"] not in usados_atualizados]
    return pares, removidos, novos


def _itens(texto):
    linhas = [linha.strip() for linha in (texto or "").splitlines() if linha.strip()]
    if not linhas:
        return []

    itens = []
    inicio_item = re.compile(r"^(?:[-•]\s*|(?:SP|SM)\s*\d+\b)", re.IGNORECASE)
    for linha in linhas:
        if inicio_item.search(linha) or not itens:
            itens.append(re.sub(r"^[-•]\s*", "", linha).strip())
        else:
            itens[-1] = f"{itens[-1]} {linha}".strip()
    return itens


def _diferenciar_itens(texto_antigo, texto_atualizado):
    antigos = _itens(texto_antigo)
    atualizados = _itens(texto_atualizado)
    usados_antigos = set()
    usados_atualizados = set()
    alterados = []

    for i, antigo in enumerate(antigos):
        for j, atualizado in enumerate(atualizados):
            if j in usados_atualizados:
                continue
            if _normalizar(antigo) == _normalizar(atualizado):
                usados_antigos.add(i)
                usados_atualizados.add(j)
                break

    # Itens com o mesmo código SP/SM são a mesma salvaguarda, ainda que
    # sua descrição tenha sido ampliada entre as revisões.
    def codigo(item):
        achado = re.match(r"^(SP|SM)\s*(\d+)\b", item, flags=re.IGNORECASE)
        return f"{achado.group(1).upper()}{achado.group(2)}" if achado else ""

    for i, antigo in enumerate(antigos):
        if i in usados_antigos or not codigo(antigo):
            continue
        for j, atualizado in enumerate(atualizados):
            if j in usados_atualizados:
                continue
            if codigo(antigo) == codigo(atualizado):
                usados_antigos.add(i)
                usados_atualizados.add(j)
                alterados.append((antigo, atualizado, _similaridade_texto(antigo, atualizado)))
                break

    candidatos = []
    for i, antigo in enumerate(antigos):
        if i in usados_antigos:
            continue
        for j, atualizado in enumerate(atualizados):
            if j in usados_atualizados:
                continue
            similaridade = _similaridade_texto(antigo, atualizado)
            if similaridade >= 0.60:
                candidatos.append((similaridade, i, j))

    for similaridade, i, j in sorted(candidatos, reverse=True):
        if i in usados_antigos or j in usados_atualizados:
            continue
        usados_antigos.add(i)
        usados_atualizados.add(j)
        alterados.append((antigos[i], atualizados[j], similaridade))

    excluidos = [item for i, item in enumerate(antigos) if i not in usados_antigos]
    incluidos = [item for j, item in enumerate(atualizados) if j not in usados_atualizados]
    return incluidos, excluidos, alterados


def _adicionar_alteracoes_lista(
    alteracoes,
    antigo,
    atualizado,
    campo,
    texto_antigo,
    texto_atualizado,
    secao=None,
    tratar_texto_alterado=True,
):
    incluidos, excluidos, alterados = _diferenciar_itens(
        texto_antigo, texto_atualizado
    )
    base = {
        "ID antiga": antigo["ID"],
        "ID atualizada": atualizado["ID"],
        "Sistema": atualizado["Sistema"],
        "Perigo": atualizado["Perigo"].replace("\n", " "),
        "Campo": campo,
    }

    if secao:
        rotulo_incluido = f"Salvaguarda incluída ({secao})"
        rotulo_excluido = f"Salvaguarda excluída ({secao})"
        rotulo_alterado = f"Salvaguarda alterada ({secao})"
    else:
        terminacao = "o" if campo == "Efeito" else "a"
        rotulo_incluido = f"{campo} incluíd{terminacao}"
        rotulo_excluido = f"{campo} excluíd{terminacao}"
        rotulo_alterado = f"{campo} alterad{terminacao}"

    for item in incluidos:
        alteracoes.append(
            {
                **base,
                "Tipo de alteração": rotulo_incluido,
                "Antes": "",
                "Depois": item,
                "Detalhe / comentário": "",
            }
        )
    for item in excluidos:
        alteracoes.append(
            {
                **base,
                "Tipo de alteração": rotulo_excluido,
                "Antes": item,
                "Depois": "",
                "Detalhe / comentário": "",
            }
        )
    for antes, depois, similaridade in alterados:
        if tratar_texto_alterado:
            alteracoes.append(
                {
                    **base,
                    "Tipo de alteração": rotulo_alterado,
                    "Antes": antes,
                    "Depois": depois,
                    "Detalhe / comentário": f"similaridade {similaridade:.0%}",
                }
            )
        else:
            # O modelo registra mudanças de causa/efeito como a exclusão do
            # texto anterior e a inclusão do novo texto, sem a categoria
            # intermediária "alterada".
            alteracoes.extend(
                [
                    {
                        **base,
                        "Tipo de alteração": rotulo_excluido,
                        "Antes": antes,
                        "Depois": "",
                        "Detalhe / comentário": "",
                    },
                    {
                        **base,
                        "Tipo de alteração": rotulo_incluido,
                        "Antes": "",
                        "Depois": depois,
                        "Detalhe / comentário": "",
                    },
                ]
            )


ORDEM_SEVERIDADE = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5}
DIMENSOES = (
    ("S", "Segurança", "Sev S"),
    ("P", "Patrimônio", "Sev P"),
    ("M", "Meio Ambiente", "Sev M"),
    ("I", "Imagem", "Sev I"),
)


def _severidade_numero(valor):
    return ORDEM_SEVERIDADE.get(str(valor or "").strip().upper(), 0)


def candidato_bowtie(cenario):
    """Regra do Guia: severidade IV/V em S, P ou M."""
    return any(
        _severidade_numero(cenario.get(campo)) >= 4
        for campo in ("Sev S", "Sev P", "Sev M")
    )


def candidato_bowtie_gestao_dinamica(cenario):
    """Observação 9.1 do Guia: na GDB, Imagem IV/V também é considerada."""
    return any(
        _severidade_numero(cenario.get(campo)) >= 4
        for campo in ("Sev S", "Sev P", "Sev M", "Sev I")
    )


BARREIRAS_BOWTIE = {
    "B1": ("Contenção Primária", "Preventiva", "Guia Bow Tie rev.7, páginas 9 e 18"),
    "B2": ("Controle Básico do Processo", "Preventiva", "Guia Bow Tie rev.7, páginas 9 e 18"),
    "B3": ("Alarmes Críticos e Intervenção Humana", "Preventiva", "Guia Bow Tie rev.7, páginas 9, 10 e 18"),
    "B4": ("Intertravamentos de Segurança", "Preventiva", "Guia Bow Tie rev.7, páginas 9 e 18"),
    "B5": ("Sistemas de Alívio", "Preventiva", "Guia Bow Tie rev.7, páginas 9 e 18"),
    "B6": ("Proteção Pós-Liberação - ESD", "Mitigadora", "Guia Bow Tie rev.7, páginas 10 e 16"),
    "B7": ("Controle de Fonte de Ignição", "Mitigadora", "Guia Bow Tie rev.7, páginas 10 e 16"),
    "B8": ("Resposta à Emergência da Operação", "Mitigadora", "Guia Bow Tie rev.7, páginas 10 e 17"),
    "B9": ("Resposta à Emergência da Brigada", "Mitigadora", "Guia Bow Tie rev.7, páginas 10 e 17"),
    "B10": ("Evacuação, Resgate e Abandono", "Mitigadora", "Guia Bow Tie rev.7, páginas 10 e 17"),
    "CD": ("Controle de degradação", "Controle de degradação", "Guia Bow Tie rev.7, páginas 19 a 23"),
    "NB": ("Não se enquadra como barreira", "Não barreira", "Guia Bow Tie rev.7, páginas 7 e 8"),
    "AV": ("A avaliar pelo validador", "A avaliar", "Guia Bow Tie rev.7, páginas 7 a 23"),
}


def _contem_termo(texto, *termos):
    return any(_normalizar(termo) in texto for termo in termos)


def classificar_salvaguarda_bowtie(descricao, secao_apr="", cenario=None):
    """Sugere o enquadramento usando as definições e exemplos do Guia rev.7.

    O resultado é deliberadamente conservador. Quando o texto não evidencia
    função, momento de atuação ou ação automática/manual, retorna AV.
    """
    cenario = cenario or {}
    texto = _normalizar(descricao)
    secao = _normalizar(secao_apr)
    codigo = "AV"
    confianca = "Baixa"
    justificativa = (
        "O texto não evidencia, sozinho, a função e o momento de atuação "
        "necessários para classificar a barreira."
    )

    # B10: últimas ações do PRE e proteção de pessoas expostas.
    if _contem_termo(
        texto,
        "evacuação",
        "abandono",
        "resgate",
        "resposta a emergência na comunidade",
        "ponto de encontro",
        "rota de fuga",
    ):
        codigo, confianca = "B10", "Alta"
        justificativa = (
            "Evacuação, resgate ou abandono: últimas ações do Plano de "
            "Resposta a Emergências."
        )

    # B9: primeiras ações e recursos da brigada, conforme o PRE.
    elif _contem_termo(
        texto,
        "brigada",
        "plano de resposta a emergencia",
        "procedimento operacional de resposta a emergencia",
        "sistema fixo de diluvio",
        "hidrante",
        "canhao monitor",
        "combate a incendio",
    ):
        codigo, confianca = "B9", "Alta"
        justificativa = (
            "Primeiras ações, equipamentos ou insumos da brigada previstos "
            "no Plano de Resposta a Emergências."
        )

    # B5 deve preceder as regras de instrumentos/automação.
    elif _contem_termo(
        texto,
        "psv",
        "valvula de alivio",
        "válvula de alívio",
        "valvulas de alivio",
        "válvulas de alívio",
        "disco de ruptura",
        "sistema de alivio",
        "sistema de alívio",
        "flare",
        "corta chama",
    ):
        codigo, confianca = "B5", "Alta"
        justificativa = (
            "Dispositivo mecânico de alívio ou sistema de queima que protege "
            "contra pressão excessiva."
        )

    # B4: sensor/lógica/elemento final com atuação automática antes do topo.
    elif _contem_termo(
        texto,
        "intertravamento",
        "intertravado",
        "trip",
        "shutdown",
        "desligamento automatico",
        "parada automatica",
        "fechamento automatico",
        "fechando automaticamente",
        "acionamento automatico",
        "corte automatico",
    ):
        codigo, confianca = "B4", "Alta"
        justificativa = (
            "Ação automática de segurança disparada por condição anormal, "
            "com sensor, lógica e elemento final."
        )

    # B7: requisitos de integridade para não se tornar fonte de ignição.
    elif _contem_termo(
        texto,
        "spda",
        "aterramento",
        "equipamento ex",
        "painel ex",
        "luminaria ex",
        "fonte de ignicao",
        "purga",
        "exaustor",
        "ventilacao",
    ):
        codigo, confianca = "B7", "Média"
        justificativa = (
            "Elemento associado ao controle de fontes de ignição. O Guia "
            "classifica B7 como mitigadora."
        )

    # B6: elementos passivos ou automáticos que mitigam após a liberação.
    elif _contem_termo(
        texto,
        "dique",
        "bacia de contencao",
        "bacia de contenção",
        "parede corta fogo",
        "enclausuramento",
        "sistema de drenagem",
        "protecao pos liberacao",
        "proteção pós liberação",
        "sistema de co2",
        "valvula vent",
        "válvula vent",
        "esd automatico",
        "esd automático",
        "area contida",
        "área contida",
        "chuveiro lava olhos",
        "chuveiro e lava olhos",
        "diphoterine",
        "kit de protecao ambiental",
        "kit de proteção ambiental",
        "protecao de acrilico",
        "proteção de acrílico",
        "sistema de protecao contra incendio",
        "sistema de proteção contra incêndio",
    ):
        codigo, confianca = "B6", "Alta"
        justificativa = (
            "Elemento passivo ou automático que reduz inventário ou mitiga "
            "as consequências após a perda de contenção."
        )

    # B2: malha automática usada no controle normal do processo.
    elif _contem_termo(
        texto,
        "controle basico",
        "controle básico",
        "valvula controladora",
        "válvula controladora",
        "malha de controle",
        "controle do nivel",
        "controle de nivel",
        "controle de pressão",
        "controle de pressao",
    ):
        codigo, confianca = "B2", "Média"
        justificativa = (
            "Malha automática que mantém variável de processo na faixa "
            "normal de operação."
        )

    # B3: anúncio de condição anormal que exige resposta humana predefinida.
    elif _contem_termo(
        texto,
        "alarme",
        "transmissor",
        "pressostato",
        "sensor",
        "detector",
    ) or re.search(
        r"\b(?:PIT|LIT|TIT|FIT|PT|LT|TT|FT|PAHH?|PALL?|LAHH?|LALL?|TAHH?|TALL?|LSHH?|LSLL?|VAHH?|VALL?)\b",
        str(descricao or "").upper(),
    ):
        codigo, confianca = "B3", "Média"
        justificativa = (
            "Instrumentação/alarme que anuncia condição anormal e requer "
            "intervenção humana predefinida. Confirmar procedimento associado."
        )

    # B8: manobra operacional local/remota após o evento topo.
    elif _contem_termo(
        texto,
        "caracteristica organoleptica",
        "características organolépticas",
        "botoeira esd",
        "manobra operacional",
        "atuacao manual",
        "atuação manual",
        "operador",
        "ruido",
        "ruído",
        "visual",
        "audicao",
        "audição",
        "olfato",
    ):
        codigo, confianca = "B8", "Média"
        justificativa = (
            "Detecção ou ação manual da operação para limitar consequências "
            "após o evento topo. Confirmar a manobra prevista."
        )

    # Controles que sustentam o desempenho das barreiras, sem agir no cenário.
    elif _contem_termo(
        texto,
        "plano de manutencao",
        "plano de manutenção",
        "plano de inspecao",
        "plano de inspeção",
        "manutencao preventiva",
        "manutenção preventiva",
        "calibracao",
        "calibração",
        "gestao de mudanca",
        "gestão de mudança",
        "matriz de treinamento",
        "simulado de emergencia",
        "simulado de emergência",
        "check list",
        "checklist",
        "procedimento",
        "fispq",
    ) or re.search(r"\bPE-\s*\d", str(descricao or "").upper()):
        codigo, confianca = "CD", "Alta"
        justificativa = (
            "Processo de gestão que mantém a confiabilidade/disponibilidade "
            "da barreira; atua como controle de degradação."
        )

    elif _contem_termo(
        texto,
        "controle de acesso",
        "sinalizacao",
        "sinalização",
        "redundancia",
        "redundância",
        "conservacao de areas",
        "conservação de áreas",
        "aceiro",
        "dragagem",
    ):
        codigo, confianca = "NB", "Alta"
        justificativa = (
            "Medida organizacional ou de disponibilidade sem função autônoma "
            "suficiente para evitar o evento topo ou mitigar a consequência."
        )

    elif _contem_termo(
        texto,
        "contencao primaria",
        "contenção primária",
        "material adequado",
        "revestimento adequado",
        "dupla parede",
    ):
        codigo, confianca = "B1", "Média"
        justificativa = (
            "Premissa de projeto associada à integridade da contenção primária."
        )

    nome, tipo, base_normativa = BARREIRAS_BOWTIE[codigo]
    alerta = ""
    if tipo == "Mitigadora" and "preventiva" in secao:
        alerta = (
            "A APR apresenta o item como preventivo, mas a função sugerida "
            "atua após o evento topo. Confirmar com o validador."
        )
    elif tipo == "Preventiva" and "mitigadora" in secao:
        alerta = (
            "A APR apresenta o item como mitigador, mas a função sugerida "
            "atua antes do evento topo. Confirmar com o validador."
        )
    elif codigo == "AV":
        alerta = "Classificação inconclusiva: validar função e momento de atuação."

    return {
        "Código Bow Tie": codigo,
        "Nome da barreira": nome,
        "Tipo Bow Tie": tipo,
        "Confiança": confianca,
        "Justificativa": justificativa,
        "Alerta": alerta,
        "Base normativa": base_normativa,
    }


def inventariar_salvaguardas_bowtie(cenarios, revisao):
    linhas = []
    secoes = [
        ("Detecção", "Modo de detecção"),
        ("Preventiva", "Salvaguardas preventivas (SP)"),
        ("Mitigadora", "Salvaguardas mitigadoras (SM)"),
        ("Sem tipo", "Salvaguardas sem tipo"),
    ]
    for cenario in cenarios:
        for secao, campo in secoes:
            for item in _itens(cenario.get(campo, "")):
                classificacao = classificar_salvaguarda_bowtie(
                    item,
                    secao_apr=secao,
                    cenario=cenario,
                )
                linhas.append(
                    {
                        "Revisão": revisao,
                        "ID": cenario["ID"],
                        "Nó": cenario["Nó"],
                        "Sistema": cenario["Sistema"],
                        "Perigo": cenario["Perigo"],
                        "Seção na APR": secao,
                        "Código": _codigo_item(item),
                        "Descrição": item,
                        **classificacao,
                        "Candidato Bow Tie mínimo": (
                            "SIM" if candidato_bowtie(cenario) else "Não"
                        ),
                        "Candidato GDB (inclui Imagem)": (
                            "SIM"
                            if candidato_bowtie_gestao_dinamica(cenario)
                            else "Não"
                        ),
                        "Página PDF": cenario.get("Página PDF", ""),
                    }
                )
    return linhas


def _codigo_item(item):
    achado = re.match(r"^[\s'\"“”‘’]*(SP|SM)\s*0*(\d+)\b", item, flags=re.IGNORECASE)
    if not achado:
        return ""
    return f"{achado.group(1).upper()}{achado.group(2)}"


def _severidades(cenario):
    return "/".join(str(cenario.get(campo, "") or "") for _, _, campo in DIMENSOES)


def _riscos(cenario):
    return "/".join(
        str(cenario.get(campo, "") or "")
        for campo in ("Risco S", "Risco P", "Risco M", "Risco I")
    )


def _dimensoes_criticas(cenario):
    return ", ".join(
        f"{nome} ({sigla}) = {cenario.get(campo)}"
        for sigla, nome, campo in DIMENSOES[:3]
        if _severidade_numero(cenario.get(campo)) >= 4
    )


def _analisar_evolucao_bowtie(
    cenarios_antigos,
    cenarios_atualizados,
    correspondencias,
):
    antigos = {cenario["ID"]: cenario for cenario in cenarios_antigos}
    atualizados = {cenario["ID"]: cenario for cenario in cenarios_atualizados}
    linhas = []

    for correspondencia in correspondencias:
        antigo = antigos.get(correspondencia["ID antiga"])
        atualizado = atualizados.get(correspondencia["ID atualizada"])
        candidato_antigo = candidato_bowtie(antigo) if antigo else False
        candidato_atual = candidato_bowtie(atualizado) if atualizado else False
        candidato_gdb_antigo = (
            candidato_bowtie_gestao_dinamica(antigo) if antigo else False
        )
        candidato_gdb_atual = (
            candidato_bowtie_gestao_dinamica(atualizado) if atualizado else False
        )

        if atualizado and candidato_atual and candidato_antigo:
            situacao = "PERMANECE CANDIDATO"
        elif atualizado and candidato_atual:
            situacao = "NOVO CANDIDATO"
        elif antigo and candidato_antigo and atualizado:
            situacao = "DEIXOU DE SER CANDIDATO"
        elif antigo and candidato_antigo:
            situacao = "CANDIDATO REMOVIDO"
        else:
            situacao = "NÃO CANDIDATO"

        referencia = atualizado or antigo
        linhas.append(
            {
                "Situação Bow Tie": situacao,
                "ID antiga": antigo["ID"] if antigo else "",
                "ID atualizada": atualizado["ID"] if atualizado else "",
                "Nó": referencia.get("Nó", ""),
                "Sistema": referencia.get("Sistema", ""),
                "Perigo": referencia.get("Perigo", "").replace("\n", " "),
                "Candidato na APR antiga": "SIM" if candidato_antigo else "Não",
                "Candidato na APR atualizada": "SIM" if candidato_atual else "Não",
                "Candidato GDB na APR antiga": (
                    "SIM" if candidato_gdb_antigo else "Não"
                ),
                "Candidato GDB na APR atualizada": (
                    "SIM" if candidato_gdb_atual else "Não"
                ),
                "Severidade antiga S/P/M/I": _severidades(antigo) if antigo else "",
                "Severidade atualizada S/P/M/I": (
                    _severidades(atualizado) if atualizado else ""
                ),
                "Dimensões críticas atuais": (
                    _dimensoes_criticas(atualizado) if atualizado else ""
                ),
                "Critério": (
                    "Severidade IV ou V em Segurança, Patrimônio ou Meio Ambiente"
                    if candidato_atual
                    else "Não atende ao critério IV/V em S, P ou M"
                ),
                "Status do cenário": correspondencia["Status"],
            }
        )

    ordem_situacao = {
        "NOVO CANDIDATO": 0,
        "PERMANECE CANDIDATO": 1,
        "DEIXOU DE SER CANDIDATO": 2,
        "CANDIDATO REMOVIDO": 3,
        "NÃO CANDIDATO": 4,
    }
    linhas.sort(
        key=lambda item: (
            _chave_id(item["ID atualizada"] or item["ID antiga"]),
            ordem_situacao[item["Situação Bow Tie"]],
        )
    )
    return linhas


def comparar_revisoes(cenarios_antigos, cenarios_atualizados):
    pares, removidos, novos = _encontrar_correspondencias(
        cenarios_antigos, cenarios_atualizados
    )
    correspondencias = []
    comparativo = []
    alteracoes = []

    campos_lista = [
        ("Causa", "Causas", None, False),
        ("Efeito", "Possíveis efeitos", None, False),
        ("Salvaguarda", "Modo de detecção", "Detecção", True),
        (
            "Salvaguarda",
            "Salvaguardas preventivas (SP)",
            "Preventiva",
            True,
        ),
        (
            "Salvaguarda",
            "Salvaguardas mitigadoras (SM)",
            "Mitigadora",
            True,
        ),
        ("Salvaguarda", "Salvaguardas sem tipo", "Sem tipo", True),
    ]

    for antigo, atualizado, pontuacao, origem in pares:
        perigo_igual = _normalizar(antigo["Perigo"]) == _normalizar(
            atualizado["Perigo"]
        )
        status = "MANTIDO / ALTERADO" if perigo_igual else "RENOMEADO / ALTERADO"
        correspondencias.append(
            {
                "Status": status,
                "ID antiga": antigo["ID"],
                "ID atualizada": atualizado["ID"],
                "Perigo antigo": antigo["Perigo"].replace("\n", " "),
                "Perigo atualizado": atualizado["Perigo"].replace("\n", " "),
                "Correspondência": origem,
                "Confiança": _confianca(pontuacao),
                "Similaridade": round(pontuacao * 100, 1),
            }
        )

        if not perigo_igual:
            alteracoes.append(
                {
                    "ID antiga": antigo["ID"],
                    "ID atualizada": atualizado["ID"],
                    "Sistema": atualizado["Sistema"],
                    "Perigo": (
                        f'{antigo["Perigo"].replace(chr(10), " ")} → '
                        f'{atualizado["Perigo"].replace(chr(10), " ")}'
                    ),
                    "Campo": "Perigo",
                    "Tipo de alteração": "Cenário renomeado",
                    "Antes": antigo["Perigo"].replace("\n", " "),
                    "Depois": atualizado["Perigo"].replace("\n", " "),
                    "Detalhe / comentário": "Perigo renomeado entre as revisões.",
                }
            )

        for nome, chave, secao, tratar_alterado in campos_lista:
            _adicionar_alteracoes_lista(
                alteracoes,
                antigo,
                atualizado,
                nome,
                antigo.get(chave, ""),
                atualizado.get(chave, ""),
                secao=secao,
                tratar_texto_alterado=tratar_alterado,
            )

        base_direta = {
            "ID antiga": antigo["ID"],
            "ID atualizada": atualizado["ID"],
            "Sistema": atualizado["Sistema"],
            "Perigo": atualizado["Perigo"].replace("\n", " "),
        }

        if _normalizar(str(antigo.get("F", ""))) != _normalizar(
            str(atualizado.get("F", ""))
        ):
            alteracoes.append(
                {
                    **base_direta,
                    "Campo": "F",
                    "Tipo de alteração": "Frequência alterada",
                    "Antes": antigo.get("F", ""),
                    "Depois": atualizado.get("F", ""),
                    "Detalhe / comentário": "",
                }
            )

        severidades_antes = _severidades(antigo)
        severidades_depois = _severidades(atualizado)
        if severidades_antes != severidades_depois:
            detalhes = []
            for sigla, nome, campo in DIMENSOES:
                antes = str(antigo.get(campo, "") or "")
                depois = str(atualizado.get(campo, "") or "")
                if antes != depois:
                    detalhes.append(f"{nome} ({sigla}): {antes}→{depois}")
            alteracoes.append(
                {
                    **base_direta,
                    "Campo": "Severidade",
                    "Tipo de alteração": "Severidade alterada",
                    "Antes": severidades_antes,
                    "Depois": severidades_depois,
                    "Detalhe / comentário": "; ".join(detalhes),
                }
            )

        riscos_antes = _riscos(antigo)
        riscos_depois = _riscos(atualizado)
        if riscos_antes != riscos_depois:
            alteracoes.append(
                {
                    **base_direta,
                    "Campo": "Categoria de risco",
                    "Tipo de alteração": "Categoria de risco alterada",
                    "Antes": riscos_antes,
                    "Depois": riscos_depois,
                    "Detalhe / comentário": (
                        "T=Tolerável, M=Moderado, NT=Não tolerável"
                    ),
                }
            )

        observacoes_antes = antigo.get("Observações / Recomendações", "") or ""
        observacoes_depois = atualizado.get("Observações / Recomendações", "") or ""
        if _normalizar(str(observacoes_antes)) != _normalizar(
            str(observacoes_depois)
        ):
            alteracoes.append(
                {
                    **base_direta,
                    "Campo": "Observações / Recomendações",
                    "Tipo de alteração": "Observação/recomendação alterada",
                    "Antes": observacoes_antes,
                    "Depois": observacoes_depois,
                    "Detalhe / comentário": "",
                }
            )

        comparativo.append(
            {
                "Status": status,
                "ID antiga": antigo["ID"],
                "ID atualizada": atualizado["ID"],
                "Nó antigo": antigo["Nó"],
                "Nó atualizado": atualizado["Nó"],
                "Sistema": atualizado["Sistema"],
                "Perigo antigo": antigo["Perigo"].replace("\n", " "),
                "Perigo atualizado": atualizado["Perigo"].replace("\n", " "),
                "F antiga": antigo["F"],
                "F atualizada": atualizado["F"],
                "Severidade antiga": "/".join(
                    str(antigo[c]) for c in ("Sev S", "Sev P", "Sev M", "Sev I")
                ),
                "Severidade atualizada": "/".join(
                    str(atualizado[c])
                    for c in ("Sev S", "Sev P", "Sev M", "Sev I")
                ),
                "Risco antigo": "/".join(
                    str(antigo[c])
                    for c in ("Risco S", "Risco P", "Risco M", "Risco I")
                ),
                "Risco atualizado": "/".join(
                    str(atualizado[c])
                    for c in ("Risco S", "Risco P", "Risco M", "Risco I")
                ),
                "Confiança": _confianca(pontuacao),
                "Similaridade": round(pontuacao * 100, 1),
                "Candidato Bow Tie antigo": (
                    "SIM" if candidato_bowtie(antigo) else "Não"
                ),
                "Candidato Bow Tie atualizado": (
                    "SIM" if candidato_bowtie(atualizado) else "Não"
                ),
            }
        )

    for cenario in removidos:
        correspondencias.append(
            {
                "Status": "REMOVIDO",
                "ID antiga": cenario["ID"],
                "ID atualizada": "",
                "Perigo antigo": cenario["Perigo"].replace("\n", " "),
                "Perigo atualizado": "",
                "Correspondência": "Sem par automático",
                "Confiança": "Revisar",
                "Similaridade": "",
            }
        )
        alteracoes.append(
            {
                "ID antiga": cenario["ID"],
                "ID atualizada": "",
                "Sistema": cenario["Sistema"],
                "Perigo": cenario["Perigo"].replace("\n", " "),
                "Campo": "Cenário",
                "Tipo de alteração": "Cenário removido",
                "Antes": cenario["Perigo"].replace("\n", " "),
                "Depois": "",
                "Detalhe / comentário": "",
            }
        )

    for cenario in novos:
        correspondencias.append(
            {
                "Status": "NOVO",
                "ID antiga": "",
                "ID atualizada": cenario["ID"],
                "Perigo antigo": "",
                "Perigo atualizado": cenario["Perigo"].replace("\n", " "),
                "Correspondência": "Sem par automático",
                "Confiança": "Revisar",
                "Similaridade": "",
            }
        )
        alteracoes.append(
            {
                "ID antiga": "",
                "ID atualizada": cenario["ID"],
                "Sistema": cenario["Sistema"],
                "Perigo": cenario["Perigo"].replace("\n", " "),
                "Campo": "Cenário",
                "Tipo de alteração": "Cenário novo",
                "Antes": "",
                "Depois": cenario["Perigo"].replace("\n", " "),
                "Detalhe / comentário": "",
            }
        )

    ordem_status = {
        "MANTIDO / ALTERADO": 0,
        "RENOMEADO / ALTERADO": 1,
        "REMOVIDO": 2,
        "NOVO": 3,
    }
    correspondencias.sort(
        key=lambda item: (
            _chave_id(item["ID atualizada"] or item["ID antiga"]),
            ordem_status[item["Status"]],
        )
    )
    comparativo.sort(
        key=lambda item: _chave_id(item["ID atualizada"] or item["ID antiga"])
    )
    alteracoes.sort(
        key=lambda item: (
            _chave_id(item["ID atualizada"] or item["ID antiga"]),
            item.get("Tipo de alteração", ""),
        )
    )
    resumo = Counter(item["Status"] for item in correspondencias)
    evolucao_bowtie = _analisar_evolucao_bowtie(
        cenarios_antigos,
        cenarios_atualizados,
        correspondencias,
    )
    resumo["CANDIDATOS BOW TIE ANTIGA"] = sum(
        candidato_bowtie(cenario) for cenario in cenarios_antigos
    )
    resumo["CANDIDATOS BOW TIE ATUALIZADA"] = sum(
        candidato_bowtie(cenario) for cenario in cenarios_atualizados
    )
    resumo["CANDIDATOS GDB ANTIGA"] = sum(
        candidato_bowtie_gestao_dinamica(cenario)
        for cenario in cenarios_antigos
    )
    resumo["CANDIDATOS GDB ATUALIZADA"] = sum(
        candidato_bowtie_gestao_dinamica(cenario)
        for cenario in cenarios_atualizados
    )
    resumo["NOVOS CANDIDATOS BOW TIE"] = sum(
        item["Situação Bow Tie"] == "NOVO CANDIDATO" for item in evolucao_bowtie
    )
    resumo["DEIXARAM DE SER CANDIDATOS BOW TIE"] = sum(
        item["Situação Bow Tie"] in {
            "DEIXOU DE SER CANDIDATO",
            "CANDIDATO REMOVIDO",
        }
        for item in evolucao_bowtie
    )
    return {
        "correspondencias": correspondencias,
        "comparativo": comparativo,
        "alteracoes": alteracoes,
        "evolucao_bowtie": evolucao_bowtie,
        "resumo": dict(resumo),
    }
