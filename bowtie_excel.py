"""Preenche o modelo XLSM com dados da APR, preservando os demais arquivos do ZIP.

O arquivo Excel é um pacote OOXML. Ao alterar somente o XML das abas BT/B1..B10,
as macros, vínculos externos e demais componentes existentes permanecem intactos.
"""

import io
import re
import unicodedata
from collections import defaultdict
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from lxml import etree

from comparacao import inventariar_salvaguardas_bowtie


MODELO = Path(__file__).resolve().parent / "assets" / "modelo_no.xlsm"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PCK = "http://schemas.openxmlformats.org/package/2006/relationships"
X = {"x": NS}
COLUNAS = dict(zip((f"B{i}" for i in range(1, 11)), "DEFGHJKLMN"))
COLUNAS_EXTENSAS = dict(zip((f"B{i}" for i in range(1, 11)),
                             ("AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ")))
LIMITE_ARQUIVO = 20 * 1024 * 1024


def _normalizar(texto):
    texto = unicodedata.normalize("NFKD", str(texto or ""))
    return re.sub(r"\s+", " ", texto.encode("ascii", "ignore").decode("ascii").casefold()).strip()


def _texto(texto):
    return re.sub(r"[ \t]+", " ", str(texto or "").replace("\r", "\n")).strip()


def _itens(texto):
    """Separa entradas da APR sem alterar suas descrições."""
    partes = []
    inicio = re.compile(r"^(?:[-•]\s*|(?:SP|SM)\s*\d+\b)", re.I)
    for linha in _texto(texto).splitlines():
        linha = linha.strip()
        if not linha:
            continue
        if inicio.search(linha) or not partes:
            partes.append(linha)
        else:
            partes[-1] += " " + linha
    return list(dict.fromkeys(partes))


def _codigo(texto):
    achado = re.match(r"^[\s'\"“”‘’]*(SP|SM)\s*0*(\d+)\b", str(texto), re.I)
    return f"{achado.group(1).upper()}{achado.group(2)}" if achado else ""


def _tag_e_descricao(descricao):
    """Transfere apenas os quatro códigos SM indicados pela APR para TAG."""
    codigo = _codigo(descricao)
    if codigo not in {"SM2", "SM7", "SM8", "SM10"}:
        return "", descricao
    texto = re.sub(
        r"^[\s'\"“”‘’]*SM\s*0*(?:2|7|8|10)\b\s*[-–—:;.]?\s*",
        "", descricao, count=1, flags=re.IGNORECASE,
    )
    return codigo, texto


def _grupo_itens(cenario):
    agrupados = defaultdict(list)
    revisar = []
    for item in inventariar_salvaguardas_bowtie([cenario], "APR atualizada"):
        if item["Seção na APR"] == "Detecção":
            continue
        codigo = item["Código Bow Tie"]
        descricao = _texto(item["Descrição"]).replace("\n", " ")
        if codigo in COLUNAS and item.get("Confiança") == "Alta" and not item.get("Alerta"):
            if descricao not in agrupados[codigo]:
                agrupados[codigo].append(descricao)
        else:
            # Uma inferência incerta jamais é registrada como barreira confirmada.
            revisar.append(descricao)
    return agrupados, list(dict.fromkeys(revisar))


def _abrir_pacote(dados):
    if not isinstance(dados, bytes) or len(dados) > LIMITE_ARQUIVO:
        raise ValueError("O arquivo BowTie deve ser um .xlsm de até 20 MB.")
    try:
        z = ZipFile(io.BytesIO(dados))
        nomes = set(z.namelist())
        obrigatorios = {"xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/vbaProject.bin"}
        if not obrigatorios.issubset(nomes) or sum(i.file_size for i in z.infolist()) > 120 * 1024 * 1024:
            raise ValueError("Envie um BowTie .xlsm compatível com o modelo, com macros.")
        if any(i.file_size > 25 * 1024 * 1024 for i in z.infolist()):
            raise ValueError("Uma parte do arquivo BowTie excede o tamanho permitido.")
        wb = etree.fromstring(z.read("xl/workbook.xml"), etree.XMLParser(resolve_entities=False, no_network=True))
        rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"), etree.XMLParser(resolve_entities=False, no_network=True))
        caminhos = {x.get("Id"): x.get("Target") for x in rels.findall(f"{{{PCK}}}Relationship")}
        abas = {}
        for aba in wb.findall(f".//{{{NS}}}sheet"):
            path = caminhos.get(aba.get(f"{{{REL}}}id"), "")
            path = path.lstrip("/") if path.startswith("/") else "xl/" + path
            abas[aba.get("name")] = path
        if any(abas.get(nome) not in nomes for nome in ("BT", *(f"B{i}" for i in range(1, 11)), "BT - Todos")):
            raise ValueError("O BowTie enviado não tem as abas BT, B1 a B10 e BT - Todos do modelo.")
        return z, abas
    except (BadZipFile, KeyError, etree.XMLSyntaxError) as exc:
        raise ValueError("O arquivo enviado não é um BowTie .xlsm válido.") from exc


def _arvore(z, path):
    return etree.fromstring(z.read(path), etree.XMLParser(resolve_entities=False, no_network=True))


def _celula(raiz, endereco, criar=False):
    row_n = int(re.search(r"\d+", endereco).group())
    dados = raiz.find(f"{{{NS}}}sheetData")
    linha = dados.find(f"{{{NS}}}row[@r='{row_n}']")
    if linha is None and criar:
        linha = etree.Element(f"{{{NS}}}row", r=str(row_n))
        seguinte = next((x for x in dados if int(x.get("r", 0)) > row_n), None)
        if seguinte is None:
            dados.append(linha)
        else:
            dados.insert(dados.index(seguinte), linha)
    if linha is None:
        return None
    cel = linha.find(f"{{{NS}}}c[@r='{endereco}']")
    if cel is None and criar:
        cel = etree.Element(f"{{{NS}}}c", r=endereco)
        def numero_coluna(ref):
            n = 0
            for ch in re.match(r"[A-Z]+", ref).group():
                n = n * 26 + ord(ch) - 64
            return n
        pos = numero_coluna(endereco)
        seguinte = next((x for x in linha if x.tag == f"{{{NS}}}c" and numero_coluna(x.get("r")) > pos), None)
        if seguinte is None:
            linha.append(cel)
        else:
            linha.insert(linha.index(seguinte), cel)
    return cel


def _valor(raiz, endereco, compartilhadas=()):
    c = _celula(raiz, endereco)
    if c is None:
        return ""
    if c.get("t") == "inlineStr":
        return "".join(c.itertext())
    v = c.find(f"{{{NS}}}v")
    if v is None:
        return ""
    if c.get("t") == "s":
        return compartilhadas[int(v.text)]
    return v.text or ""


def _escrever(raiz, endereco, texto):
    valor = _texto(texto)
    if len(valor) > 32767:
        raise ValueError(f"O conteúdo da APR excede o limite de uma célula Excel ({endereco}).")
    c = _celula(raiz, endereco, criar=True)
    for child in list(c):
        if child.tag in {f"{{{NS}}}v", f"{{{NS}}}is", f"{{{NS}}}f"}:
            c.remove(child)
    if valor:
        c.set("t", "inlineStr")
        it = etree.SubElement(c, f"{{{NS}}}is")
        t = etree.SubElement(it, f"{{{NS}}}t")
        t.text = valor
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    else:
        c.attrib.pop("t", None)


def _shared(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    raiz = _arvore(z, "xl/sharedStrings.xml")
    return ["".join(si.itertext()) for si in raiz.findall(f"{{{NS}}}si")]


def _maiusculas_aba(raiz):
    """Altera só o texto visível; fórmulas, estilos e tipos numéricos ficam intactos."""
    for celula in raiz.findall(f".//{{{NS}}}sheetData/{{{NS}}}row/{{{NS}}}c"):
        if celula.find(f"{{{NS}}}f") is not None:
            continue
        if celula.get("t") == "inlineStr":
            for texto in celula.findall(f".//{{{NS}}}t"):
                if texto.text:
                    texto.text = texto.text.upper()
        elif celula.get("t") == "str":
            valor = celula.find(f"{{{NS}}}v")
            if valor is not None and valor.text:
                valor.text = valor.text.upper()


def _maiusculas_compartilhadas(z, alteradas):
    caminho = "xl/sharedStrings.xml"
    if caminho not in z.namelist():
        return
    raiz = _arvore(z, caminho)
    for texto in raiz.findall(f".//{{{NS}}}si//{{{NS}}}t"):
        if texto.text:
            texto.text = texto.text.upper()
    alteradas[caminho] = etree.tostring(
        raiz, encoding="UTF-8", xml_declaration=True, standalone=True,
    )


def _grid(raiz, coluna, itens):
    """Quatro linhas do quadro: excedentes permanecem completos na última."""
    for pos in range(4):
        if pos < 3:
            conteudo = itens[pos] if pos < len(itens) else ""
        else:
            conteudo = "\n".join(itens[3:])
        _escrever(raiz, f"{coluna}{7 + pos}", conteudo)


def _preencher_bt_existente(bt, ameaças, efeitos, topo):
    """Mantém estilos e cores da BT enviada; estende o quadro quando preciso."""
    ultima = max(7, 6 + len(ameaças), 6 + len(efeitos))
    linhas = bt.findall(f".//{{{NS}}}sheetData/{{{NS}}}row")
    ate_limpar = max(ultima, *(int(linha.get("r")) for linha in linhas))
    for linha in range(7, ate_limpar + 1):
        for coluna in ("C", "O", *"DEFGHJKLMN"):
            endereco = f"{coluna}{linha}"
            if linha > 10 and _celula(bt, endereco) is None:
                anterior = _celula(bt, f"{coluna}{linha - 1}")
                celula = _celula(bt, endereco, criar=True)
                if anterior is not None and anterior.get("s"):
                    celula.set("s", anterior.get("s"))
            _escrever(bt, endereco, "")
    for pos, item in enumerate(ameaças):
        linha = 7 + pos
        if pos:
            # C7 traz a cor padrão; as próximas ameaças repetem o estilo.
            origem = _celula(bt, f"C{linha - 1}")
            atual = _celula(bt, f"C{linha}", criar=True)
            if origem is not None and origem.get("s"):
                atual.set("s", origem.get("s"))
        _escrever(bt, f"C{linha}", item)
    for pos, item in enumerate(efeitos):
        linha = 7 + pos
        if pos:
            origem = _celula(bt, f"O{linha - 1}")
            atual = _celula(bt, f"O{linha}", criar=True)
            if origem is not None and origem.get("s"):
                atual.set("s", origem.get("s"))
        _escrever(bt, f"O{linha}", item)
    _escrever(bt, "I7", topo)
    for conjunto in bt.findall(f"{{{NS}}}mergeCells"):
        for m in list(conjunto):
            if m.get("ref", "").startswith("I7:I"):
                if ultima == 7:
                    conjunto.remove(m)
                else:
                    m.set("ref", f"I7:I{ultima}")
        conjunto.set("count", str(len(conjunto)))
    for linha in range(7, ultima + 1):
        causa = _valor(bt, f"C{linha}")
        efeito = _valor(bt, f"O{linha}")
        row = bt.find(f".//{{{NS}}}sheetData/{{{NS}}}row[@r='{linha}']")
        if row is not None:
            altura = min(240, 14.4 * max(1, (len(causa) + 61) // 62, (len(efeito) + 23) // 24))
            row.set("ht", str(max(float(row.get("ht", "14.4")), altura)))
            row.set("customHeight", "1")


def _validar_anterior(bt, strings, cenario_antigo, cenario_atual):
    if cenario_antigo is None:
        raise ValueError("Não foi encontrada correspondência com a APR antiga; confira o cenário antes de atualizar.")
    original = _normalizar(_valor(bt, "Z6", strings))
    esperados = {_normalizar(cenario_antigo["Perigo"]), _normalizar(cenario_atual["Perigo"])}
    if not original or original not in esperados:
        raise ValueError("O evento topo do BowTie enviado não coincide com o desvio deste cenário nas APRs. Selecione o BowTie correto.")
    cabecalho = _normalizar(_valor(bt, "C4", strings))
    if _normalizar(cenario_antigo.get("Sistema", "")) not in cabecalho and _normalizar(cenario_atual.get("Sistema", "")) not in cabecalho:
        raise ValueError("O sistema no BowTie existente não coincide com o sistema das APRs.")
    titulo = _normalizar(_valor(bt, "C2", strings))
    achado = re.search(r"\bcenario\s+n?\s*(\d+\s*[.]\s*\d+)\b", titulo)
    if achado:
        identificado = re.sub(r"\s+", "", achado.group(1))
        conhecidos = {
            re.sub(r"[^\d.]", "", str(cenario_antigo.get("ID", ""))),
            re.sub(r"[^\d.]", "", str(cenario_atual.get("ID", ""))),
        }
        if identificado not in conhecidos:
            raise ValueError("O identificador no título do BowTie não corresponde a este cenário nas APRs. Verifique o arquivo enviado.")


def _preencher_barreiras(raiz, descricoes, strings, anterior, referencia, pagina, revisar, grupo):
    """Mantém metadados preenchidos e retira do conjunto ativo itens removidos."""
    guardadas = set()
    referencia = " ".join(x for x in (referencia, f"página {pagina}" if pagina else "") if x)
    referencias_mescladas = []
    for cel in raiz.findall(f".//{{{NS}}}mergeCell"):
        ref = cel.get("ref", "")
        achado = re.fullmatch(r"L(\d+):L(\d+)", ref)
        if achado:
            referencias_mescladas.append((int(achado.group(1)), int(achado.group(2))))

    def referencia_visivel(linha):
        return all(linha == inicio or not inicio <= linha <= fim
                   for inicio, fim in referencias_mescladas)

    por_codigo = {_codigo(x): x for x in descricoes if _codigo(x)}
    for linha in range(2, 1001):
        endereco = f"C{linha}"
        texto = _valor(raiz, endereco, strings)
        if not texto:
            continue
        tag_existente = _valor(raiz, f"E{linha}", strings)
        registro = (
            f"{tag_existente} - {texto}"
            if tag_existente in {"SM2", "SM7", "SM8", "SM10"}
            and not _codigo(texto) else texto
        )
        chave = _normalizar(registro)
        codigo = _codigo(registro)
        novo = next((x for x in descricoes if _normalizar(x) == chave), None)
        if not novo and codigo and codigo in por_codigo and codigo not in guardadas:
            novo = por_codigo[codigo]
        if novo:
            guardadas.add(_codigo(novo) or _normalizar(novo))
            tag, descricao = _tag_e_descricao(novo)
            _escrever(raiz, endereco, descricao)
            if tag:
                if tag_existente and tag_existente != tag:
                    revisar.append(
                        f"{grupo}, linha {linha}: TAG anterior '{tag_existente}' substituída por '{tag}' conforme a APR."
                    )
                _escrever(raiz, f"E{linha}", tag)
            if referencia and referencia_visivel(linha) and not _valor(raiz, f"L{linha}", strings):
                _escrever(raiz, f"L{linha}", referencia)
        elif anterior is not None and any(_normalizar(x) == chave for x in anterior):
            revisar.append(
                f"{grupo}, linha {linha}: retirado da APR atualizada: {texto}. "
                "Os demais campos da linha foram preservados para conferência."
            )
            _escrever(raiz, endereco, "")
            if tag_existente in {"SM2", "SM7", "SM8", "SM10"}:
                _escrever(raiz, f"E{linha}", "")
        # Conteúdo preenchido manualmente sem vínculo verificável permanece intocado.

    for novo in descricoes:
        chave = _codigo(novo) or _normalizar(novo)
        if chave in guardadas:
            continue
        disponivel = next((linha for linha in range(2, 1001)
                           if all(not _valor(raiz, f"{coluna}{linha}", strings)
                                  for coluna in "ABCDEFGHIJKL")), None)
        if disponivel is None:
            raise ValueError("A aba de barreiras do modelo não tem linhas livres.")
        tag, descricao = _tag_e_descricao(novo)
        _escrever(raiz, f"C{disponivel}", descricao)
        if tag:
            _escrever(raiz, f"E{disponivel}", tag)
        if referencia and referencia_visivel(disponivel):
            _escrever(raiz, f"L{disponivel}", referencia)
        guardadas.add(chave)


def _retirar_modos_deteccao(raiz, strings, cenarios, revisar, grupo):
    """Retira registros de detecção identificados pelos textos das APRs."""
    nomes = {
        _normalizar(item)
        for cenario in cenarios if cenario is not None
        for item in _itens(cenario.get("Modo de detecção", ""))
    }
    if not nomes:
        return
    for linha in range(2, 1001):
        texto = _valor(raiz, f"C{linha}", strings)
        if texto and _normalizar(texto) in nomes:
            for coluna in "ABCDEFGHIJKL":
                _escrever(raiz, f"{coluna}{linha}", "")
            revisar.append(f"{grupo}, linha {linha}: modo de detecção removido do BowTie.")


def _preencher_consolidacao(raiz, folhas, strings):
    """Lista barreiras ativas sem as 550 fórmulas matriciais voláteis do modelo."""
    itens = []
    total_anterior = 0
    for pos, codigo in enumerate(COLUNAS, start=2):
        grupo = []
        folha = folhas[codigo]
        for linha in range(2, 1001):
            if _valor(folha, f"C{linha}", strings):
                grupo.append(linha)
        _escrever(raiz, f"A{pos}", str(len(grupo)))
        _escrever(raiz, f"B{pos}", str(total_anterior + 1))
        _escrever(raiz, f"C{pos}", codigo)
        for ordem, linha in enumerate(grupo, start=1):
            itens.append((codigo, ordem, folha, linha))
        total_anterior += len(grupo)
    _escrever(raiz, "A12", str(total_anterior))

    # A aba BT - Todos do modelo tem 50 linhas prontas. Acrescentar mais se
    # necessário evita qualquer perda de informação da APR.
    ultima = max(12, len(itens) + 1)
    linhas_originais = raiz.findall(f".//{{{NS}}}sheetData/{{{NS}}}row")
    ate_limpar = max(ultima, *(int(linha.get("r")) for linha in linhas_originais))
    for pos in range(2, ate_limpar + 1):
        registro = itens[pos - 2] if pos - 2 < len(itens) else None
        _escrever(raiz, f"E{pos}", str(pos - 1) if registro else "")
        _escrever(raiz, f"F{pos}", registro[0] if registro else "")
        _escrever(raiz, f"G{pos}", str(registro[1]) if registro else "")
        for deslocamento, destino in enumerate("HIJKLMNOPQR"):
            if registro:
                origem = "ABCDEFGHIJK"[deslocamento]
                valor = _valor(registro[2], f"{origem}{registro[3]}", strings)
            else:
                valor = ""
            _escrever(raiz, f"{destino}{pos}", valor)
    if ultima > 51:
        dim = raiz.find(f"{{{NS}}}dimension")
        if dim is not None:
            dim.set("ref", f"A1:R{ultima}")


def _retirar_vinculo_sap(z, livro, folhas, alteradas):
    """Retira a referência demonstrativa a um SAP externo, se não for usada."""
    if "xl/externalLinks/_rels/externalLink1.xml.rels" not in z.namelist():
        return set()
    destino = z.read("xl/externalLinks/_rels/externalLink1.xml.rels").decode("utf-8")
    if "SAP%20-%20UTE%20VLA" not in destino:
        return set()
    # Conserva o vínculo se houver outra fórmula que o utilize.
    if any("[1]" in (f.text or "")
           for folha in folhas.values()
           for f in folha.findall(f".//{{{NS}}}f")):
        return set()

    caminhos = {"xl/externalLinks/externalLink1.xml",
                "xl/externalLinks/_rels/externalLink1.xml.rels"}
    rel_path = "xl/_rels/workbook.xml.rels"
    rels = _arvore(z, rel_path)
    ids = set()
    for rel in list(rels):
        if rel.get("Type", "").endswith("/externalLink") and rel.get("Target", "").lstrip("/").endswith("externalLinks/externalLink1.xml"):
            ids.add(rel.get("Id"))
            rels.remove(rel)
    for bloco in list(livro.findall(f"{{{NS}}}externalReferences")):
        for referencia in list(bloco):
            if referencia.get(f"{{{REL}}}id") in ids:
                bloco.remove(referencia)
        if len(bloco) == 0:
            livro.remove(bloco)
    alteradas[rel_path] = etree.tostring(rels, encoding="UTF-8", xml_declaration=True, standalone=True)

    tipos = _arvore(z, "[Content_Types].xml")
    for el in list(tipos):
        if el.get("PartName") == "/xl/externalLinks/externalLink1.xml":
            tipos.remove(el)
    alteradas["[Content_Types].xml"] = etree.tostring(tipos, encoding="UTF-8", xml_declaration=True, standalone=True)
    return caminhos


def _retirar_indice_formulas(z, alteradas):
    """O índice original referencia centenas de fórmulas removidas das abas."""
    caminho = "xl/calcChain.xml"
    if caminho not in z.namelist():
        return set()
    rel_path = "xl/_rels/workbook.xml.rels"
    rels = etree.fromstring(alteradas.get(rel_path, z.read(rel_path)))
    for rel in list(rels):
        if rel.get("Type", "").endswith("/calcChain"):
            rels.remove(rel)
    alteradas[rel_path] = etree.tostring(rels, encoding="UTF-8", xml_declaration=True, standalone=True)
    tipos = etree.fromstring(alteradas.get("[Content_Types].xml", z.read("[Content_Types].xml")))
    for el in list(tipos):
        if el.get("PartName") == "/xl/calcChain.xml":
            tipos.remove(el)
    alteradas["[Content_Types].xml"] = etree.tostring(tipos, encoding="UTF-8", xml_declaration=True, standalone=True)
    return {caminho}


def _enxugar_linhas_vazias(raiz, minimo, ultima_coluna):
    """Remove linhas finais sem dados e reduz as mesclagens correspondentes."""
    dados = raiz.find(f"{{{NS}}}sheetData")
    ultima = minimo
    for linha in dados.findall(f"{{{NS}}}row"):
        if any(c.find(f"{{{NS}}}v") is not None
               or c.find(f"{{{NS}}}is") is not None
               or c.find(f"{{{NS}}}f") is not None
               for c in linha if c.tag == f"{{{NS}}}c"):
            ultima = max(ultima, int(linha.get("r")))
    for linha in list(dados):
        if linha.tag == f"{{{NS}}}row" and int(linha.get("r")) > ultima:
            dados.remove(linha)
    for grupo in list(raiz.findall(f"{{{NS}}}mergeCells")):
        for mescla in list(grupo):
            partes = re.fullmatch(
                r"([A-Z]+)(\d+):([A-Z]+)(\d+)", mescla.get("ref", "")
            )
            if not partes:
                continue
            inicial, linha_inicial, final, linha_final = partes.groups()
            inicio = int(linha_inicial)
            if inicio > ultima or (inicio == ultima and int(linha_final) > ultima):
                grupo.remove(mescla)
            elif int(linha_final) > ultima:
                mescla.set("ref", f"{inicial}{inicio}:{final}{ultima}")
        if not len(grupo):
            raiz.remove(grupo)
        else:
            grupo.set("count", str(len(grupo)))
    dim = raiz.find(f"{{{NS}}}dimension")
    if dim is not None:
        dim.set("ref", f"A1:{ultima_coluna}{ultima}")


def gerar_bowtie(cenario, apr_atualizada, modelo=None, cenario_antigo=None):
    """Um XLSX novo no design NÓ_1; atualizações conservam o XLSM enviado."""
    if modelo is None:
        from bowtie_modelo_no1 import gerar_bowtie_novo
        return gerar_bowtie_novo(cenario, apr_atualizada)
    dados = modelo if modelo is not None else MODELO.read_bytes()
    z, abas = _abrir_pacote(dados)
    strings = _shared(z)
    alteradas = {}
    bt = _arvore(z, abas["BT"])
    if modelo is not None:
        _validar_anterior(bt, strings, cenario_antigo, cenario)

    grupos, inconclusivas = _grupo_itens(cenario)
    causas = _itens(cenario.get("Causas", ""))
    efeitos = _itens(cenario.get("Possíveis efeitos", ""))
    revisar = []
    if inconclusivas:
        revisar.append(f"{len(inconclusivas)} salvaguarda(s) sem classificação inequívoca: conferir na APR atualizada.")

    no = _texto(cenario.get("Nó", ""))
    sistema = _texto(cenario.get("Sistema", "")).replace("\n", " ")
    perigo = _texto(cenario.get("Perigo", "")).replace("\n", " ")
    relatorio = _texto(apr_atualizada.get("relatorio", ""))
    _escrever(bt, "C2", " | ".join(x for x in (f"NÓ {no} - {sistema}", f"CENÁRIO {cenario.get('ID', '')}", relatorio) if x))
    _escrever(bt, "C4", f"NÓ {no} - {sistema}")
    _escrever(bt, "Z6", perigo)
    _preencher_bt_existente(bt, causas, efeitos, perigo)
    for ref in ("V5", "V6", "W5", "W6", "Z5", "Z7", "Z8", "Z9", "Z10"):
        _escrever(bt, ref, "")
    antigos = defaultdict(list)
    if cenario_antigo is not None:
        antigos, _ = _grupo_itens(cenario_antigo)
    folhas_barreiras = {}
    for codigo in COLUNAS:
        path = abas[codigo]
        folha = _arvore(z, path)
        _retirar_modos_deteccao(
            folha, strings, (cenario_antigo, cenario), revisar, codigo,
        )
        if codigo == "B8":
            # O modelo traz duas fórmulas de demonstração ligadas a um SAP
            # externo. Essa fonte não está nas APRs enviadas.
            for ref in ("F2", "I2"):
                cel = _celula(folha, ref)
                formula = cel.find(f"{{{NS}}}f") if cel is not None else None
                if formula is not None and "SAP UTE VLA" in (formula.text or ""):
                    _escrever(folha, ref, "")
            if modelo is None:
                # Também eram fórmulas demonstrativas de SITE_CODE/ID. Esses
                # valores não aparecem na APR e ficam vazios no arquivo novo.
                for ref in ("A2", "B2"):
                    _escrever(folha, ref, "")
        _preencher_barreiras(folha, grupos[codigo], strings,
                            antigos[codigo] if cenario_antigo is not None else None,
                            relatorio, cenario.get("Página PDF", ""), revisar, codigo)
        _enxugar_linhas_vazias(folha, 2, "L")
        folhas_barreiras[codigo] = folha
        alteradas[path] = etree.tostring(folha, encoding="UTF-8", xml_declaration=True, standalone=True)
    resumo = _arvore(z, abas["BT - Todos"])
    _preencher_consolidacao(resumo, folhas_barreiras, strings)
    _enxugar_linhas_vazias(resumo, 12, "R")
    alteradas[abas["BT - Todos"]] = etree.tostring(resumo, encoding="UTF-8", xml_declaration=True, standalone=True)
    _enxugar_linhas_vazias(bt, 7, "AJ")
    alteradas[abas["BT"]] = etree.tostring(bt, encoding="UTF-8", xml_declaration=True, standalone=True)

    # A consolidação foi calculada a partir dos itens da APR; não há necessidade
    # de recalcular o modelo inteiro ao abrir no Excel.
    livro = _arvore(z, "xl/workbook.xml")
    calc = livro.find(f"{{{NS}}}calcPr")
    if calc is None:
        calc = etree.SubElement(livro, f"{{{NS}}}calcPr")
    calc.set("fullCalcOnLoad", "0")
    calc.set("forceFullCalc", "0")
    puladas = _retirar_vinculo_sap(z, livro, folhas_barreiras, alteradas)
    puladas |= _retirar_indice_formulas(z, alteradas)
    alteradas["xl/workbook.xml"] = etree.tostring(livro, encoding="UTF-8", xml_declaration=True, standalone=True)

    # Inclui as abas não preenchidas diretamente para que os cabeçalhos e
    # qualquer texto preexistente no arquivo também fiquem em maiúsculas.
    for path in abas.values():
        folha = etree.fromstring(alteradas[path]) if path in alteradas else _arvore(z, path)
        _maiusculas_aba(folha)
        alteradas[path] = etree.tostring(
            folha, encoding="UTF-8", xml_declaration=True, standalone=True,
        )
    _maiusculas_compartilhadas(z, alteradas)

    saida = io.BytesIO()
    with ZipFile(saida, "w") as destino:
        for info in z.infolist():
            if info.filename in puladas:
                continue
            destino.writestr(info, alteradas.get(info.filename, z.read(info.filename)))
    z.close()
    return saida.getvalue(), list(dict.fromkeys(revisar))


def nome_bowtie(cenario, atualizado=False):
    no = re.sub(r"[^a-zA-Z0-9_-]", "", str(cenario.get("Nó", "")))
    id_ = re.sub(r"[^a-zA-Z0-9_.-]", "", str(cenario.get("ID", "")))
    extensao = "xlsm" if atualizado else "xlsx"
    return f"BowTie_No{no}_{id_}{'_atualizado' if atualizado else ''}.{extensao}"
