"""Gera BowTies novos a partir da aparência do modelo NÓ_1.xlsx.

Os valores demonstrativos do modelo são apagados. Somente os PDFs enviados
alimentam o arquivo final; os estilos, larguras, cores e abas permanecem.
"""

import copy
import io
import math
import re
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from bowtie_excel import (
    NS, PCK, REL, _arvore, _celula, _escrever, _grupo_itens, _itens, _texto,
    _valor, _maiusculas_aba, _maiusculas_compartilhadas, _tag_e_descricao,
)


MODELO = Path(__file__).resolve().parent / "assets" / "modelo_no1.xlsx"
ORIGEM = "http://schemas.openxmlformats.org/package/2006/content-types"


def _limpar_conteudo(raiz, inicio_linha):
    for linha in raiz.findall(f".//{{{NS}}}sheetData/{{{NS}}}row"):
        if int(linha.get("r")) < inicio_linha:
            continue
        for cel in list(linha):
            if cel.tag == f"{{{NS}}}c" and any(
                x.tag in (f"{{{NS}}}v", f"{{{NS}}}is", f"{{{NS}}}f") for x in cel
            ):
                _escrever(raiz, cel.get("r"), "")


def _novo_estilo_quebra(estilos, indice, cache):
    if indice in cache:
        return cache[indice]
    xfs = estilos.find(f"{{{NS}}}cellXfs")
    base = xfs[int(indice)]
    alignment = base.find(f"{{{NS}}}alignment")
    if alignment is not None and alignment.get("wrapText") == "1":
        cache[indice] = indice
        return indice
    novo = copy.deepcopy(base)
    novo.set("applyAlignment", "1")
    a = novo.find(f"{{{NS}}}alignment")
    if a is None:
        a = etree.Element(f"{{{NS}}}alignment")
        outros = [x for x in novo if x.tag in (f"{{{NS}}}protection", f"{{{NS}}}extLst")]
        if outros:
            novo.insert(novo.index(outros[0]), a)
        else:
            novo.append(a)
    a.set("wrapText", "1")
    numero = str(len(xfs))
    xfs.append(novo)
    xfs.set("count", str(len(xfs)))
    cache[indice] = numero
    return numero


def _copiar_cor_da_linha(raiz, celula, origem):
    destino = _celula(raiz, celula, criar=True)
    anterior = _celula(raiz, origem)
    if anterior is not None and anterior.get("s") is not None:
        destino.set("s", anterior.get("s"))


def _linhas_da_bt(bt, ameaças, consequencias):
    dados = bt.find(f"{{{NS}}}sheetData")
    ultima = max(24, 6 + len(ameaças), 6 + len(consequencias))
    if ultima > 24:
        base = dados.find(f"{{{NS}}}row[@r='24']")
        for numero in range(25, ultima + 1):
            copia = copy.deepcopy(base)
            copia.set("r", str(numero))
            for cel in copia:
                if cel.tag == f"{{{NS}}}c":
                    col = re.match(r"[A-Z]+", cel.get("r")).group()
                    cel.set("r", f"{col}{numero}")
                    for child in list(cel):
                        if child.tag in (f"{{{NS}}}v", f"{{{NS}}}is", f"{{{NS}}}f"):
                            cel.remove(child)
            dados.append(copia)
        for m in bt.findall(f".//{{{NS}}}mergeCell"):
            if m.get("ref") == "I7:I24":
                m.set("ref", f"I7:I{ultima}")
        dim = bt.find(f"{{{NS}}}dimension")
        if dim is not None:
            dim.set("ref", f"A1:P{ultima}")

    for i, item in enumerate(ameaças):
        row = 7 + i
        if row > 7:
            _copiar_cor_da_linha(bt, f"C{row}", f"C{row - 1}")
        _escrever(bt, f"B{row}", f"A{i + 1}")
        _escrever(bt, f"C{row}", item)
    for i, item in enumerate(consequencias):
        row = 7 + i
        if row > 7:
            _copiar_cor_da_linha(bt, f"O{row}", f"O{row - 1}")
        _escrever(bt, f"O{row}", item)
        _escrever(bt, f"P{row}", f"C{i + 1}")
    for numero in range(7, ultima + 1):
        linha = dados.find(f"{{{NS}}}row[@r='{numero}']")
        if linha is None:
            continue
        causa = _valor(bt, f"C{numero}")
        efeito = _valor(bt, f"O{numero}")
        linhas = max(1, math.ceil(len(causa) / 43), math.ceil(len(efeito) / 22))
        altura = max(float(linha.get("ht", "14.4")), min(240.0, 14.4 * linhas))
        linha.set("ht", str(round(altura, 1)))
        linha.set("customHeight", "1")
    return ultima


def _ajustar_tamanho_tabela(raiz, ultima_linha):
    """Remove linhas vazias finais e reduz as mesclagens à área com dados."""
    dados = raiz.find(f"{{{NS}}}sheetData")
    for linha in list(dados):
        if linha.tag == f"{{{NS}}}row" and int(linha.get("r")) > ultima_linha:
            dados.remove(linha)
    for conjunto in list(raiz.findall(f"{{{NS}}}mergeCells")):
        for celula in list(conjunto):
            partes = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", celula.get("ref", ""))
            if not partes:
                continue
            col_inicio, linha_inicio, col_fim, linha_fim = partes.groups()
            inicio = int(linha_inicio)
            fim = int(linha_fim)
            if inicio > ultima_linha or (fim > ultima_linha and inicio == ultima_linha):
                conjunto.remove(celula)
            elif fim > ultima_linha:
                celula.set("ref", f"{col_inicio}{inicio}:{col_fim}{ultima_linha}")
        if not len(conjunto):
            raiz.remove(conjunto)
        else:
            conjunto.set("count", str(len(conjunto)))
    dim = raiz.find(f"{{{NS}}}dimension")
    if dim is not None:
        achado = re.search(r":([A-Z]+)\d+$", dim.get("ref", ""))
        if achado:
            dim.set("ref", f"A1:{achado.group(1)}{ultima_linha}")


def _fora_de_mescla(raiz, endereco):
    col = re.match(r"[A-Z]+", endereco).group()
    linha = int(re.search(r"\d+", endereco).group())
    for mescla in raiz.findall(f".//{{{NS}}}mergeCell"):
        m = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", mescla.get("ref", ""))
        if not m:
            continue
        ci, ri, cf, rf = m.groups()
        if ci <= col <= cf and int(ri) <= linha <= int(rf):
            return endereco == f"{ci}{ri}"
    return True


def _proxima_linha(raiz, inicio, usados):
    for num in range(max(2, inicio), 1001):
        if num in usados:
            continue
        if all(_fora_de_mescla(raiz, f"{col}{num}") for col in ("C", "E")):
            return num
    raise ValueError("Não há linhas disponíveis para registrar todas as salvaguardas da APR.")


def _inserir_barreiras(raiz, codigo, itens, referencia, estilos, cache):
    """Escreve somente informações literalmente presentes na APR."""
    usados = set()
    for descricao_original in itens:
        tag, descricao = _tag_e_descricao(descricao_original)
        linha = _proxima_linha(raiz, 2, usados)
        usados.add(linha)
        _escrever(raiz, f"C{linha}", codigo)
        _copiar_cor_da_linha(raiz, f"C{linha}", "C2")
        if tag:
            _escrever(raiz, f"D{linha}", tag)
            _copiar_cor_da_linha(raiz, f"D{linha}", "D2")
        # DESCRIPTION vem da APR. Os quatro códigos selecionados saem deste
        # campo e passam à coluna TAG; outros dados não são presumidos.
        _escrever(raiz, f"E{linha}", descricao)
        c = _celula(raiz, f"E{linha}")
        if c.get("s") is None:
            base = _celula(raiz, "E2")
            if base is not None and base.get("s"):
                c.set("s", base.get("s"))
        if c.get("s"):
            c.set("s", _novo_estilo_quebra(estilos, c.get("s"), cache))
        if referencia and _fora_de_mescla(raiz, f"I{linha}"):
            _escrever(raiz, f"I{linha}", referencia)
            if linha > 2:
                _copiar_cor_da_linha(raiz, f"I{linha}", "I2")
        row = raiz.find(f".//{{{NS}}}sheetData/{{{NS}}}row[@r='{linha}']")
        if row is not None:
            atual = float(row.get("ht", "14.4"))
            row.set("ht", str(min(180, max(atual, 14.4 * (1 + len(descricao) // 38)))))
            row.set("customHeight", "1")


def _criar_aba_extra(z, mapas, alteradas, nome, id_no, sheet_id, rel_id):
    """Copia o mesmo cabeçalho e estilos de B3 para uma aba necessária."""
    copia = _arvore(z, mapas["B3"])
    _limpar_conteudo(copia, 2)
    caminho = f"xl/worksheets/sheet{sheet_id}.xml"
    alteradas[caminho] = copia
    mapas[nome] = caminho
    sheets = id_no.find(f"{{{NS}}}sheets")
    aba = etree.Element(f"{{{NS}}}sheet", name=nome, sheetId=str(sheet_id))
    aba.set(f"{{{REL}}}id", f"rId{rel_id}")
    sheets.append(aba)
    return copia, caminho


def gerar_bowtie_novo(cenario, apr):
    """Retorna um XLSX por cenário usando o design do NÓ_1.xlsx."""
    with ZipFile(MODELO) as z:
        wb = _arvore(z, "xl/workbook.xml")
        rels = _arvore(z, "xl/_rels/workbook.xml.rels")
        estilos = _arvore(z, "xl/styles.xml")
        tipos = _arvore(z, "[Content_Types].xml")
        caminhos = {r.get("Id"): r.get("Target") for r in rels.findall(f"{{{PCK}}}Relationship")}
        mapas = {}
        for aba in wb.findall(f".//{{{NS}}}sheet"):
            alvo = caminhos[aba.get(f"{{{REL}}}id")]
            mapas[aba.get("name")] = alvo.lstrip("/") if alvo.startswith("/") else "xl/" + alvo
        alteradas = {}
        grupos, pendentes = _grupo_itens(cenario)
        novos = []
        if grupos.get("B2"):
            novos.append("B2")
        if pendentes:
            novos.append("A_conferir")
        usadas = set(z.namelist())
        maior_id = max(int(x.get("sheetId")) for x in wb.findall(f".//{{{NS}}}sheet"))
        maior_rel = max(int(m.group()) for r in rels if (m := re.search(r"\d+$", r.get("Id", ""))))
        for nome in novos:
            while f"xl/worksheets/sheet{maior_id + 1}.xml" in usadas:
                maior_id += 1
            maior_id += 1
            maior_rel += 1
            folha, caminho = _criar_aba_extra(z, mapas, alteradas, nome, wb, maior_id, maior_rel)
            rel = etree.SubElement(rels, f"{{{PCK}}}Relationship")
            rel.set("Id", f"rId{maior_rel}")
            rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
            rel.set("Target", f"worksheets/sheet{maior_id}.xml")
            ct = etree.SubElement(tipos, f"{{{ORIGEM}}}Override")
            ct.set("PartName", "/" + caminho)
            ct.set("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
            usadas.add(caminho)

        bt = _arvore(z, mapas["BT"])
        _limpar_conteudo(bt, 7)
        no = _texto(cenario.get("Nó", ""))
        sistema = _texto(cenario.get("Sistema", "")).replace("\n", " ")
        relatorio = _texto(apr.get("relatorio", ""))
        _escrever(bt, "C1", f"CASO: SISTEMA APR - NÓ {no} - {sistema}")
        _escrever(bt, "C2", f"CENÁRIO {cenario.get('ID', '')} - APR - {relatorio}")
        _escrever(bt, "C4", f"NÓ {no} - {sistema}")
        ameaças = _itens(cenario.get("Causas", ""))
        consequencias = _itens(cenario.get("Possíveis efeitos", ""))
        _linhas_da_bt(bt, ameaças, consequencias)
        _escrever(bt, "I7", _texto(cenario.get("Perigo", "")).replace("\n", " "))
        _ajustar_tamanho_tabela(bt, max(7, 6 + len(ameaças), 6 + len(consequencias)))
        alteradas[mapas["BT"]] = bt

        pagina = cenario.get("Página PDF", "")
        fonte = " ".join(x for x in (relatorio, f"página {pagina}" if pagina else "") if x)
        cache_estilos = {}
        for nome in list(mapas):
            if nome == "BT":
                continue
            folha = alteradas.get(mapas[nome])
            if folha is None:
                folha = _arvore(z, mapas[nome])
            _limpar_conteudo(folha, 2)
            itens = pendentes if nome == "A_conferir" else grupos.get(nome, [])
            if itens:
                _inserir_barreiras(folha, "REVISAR" if nome == "A_conferir" else nome,
                                   itens, fonte, estilos, cache_estilos)
            preenchidas = [
                int(linha.get("r"))
                for linha in folha.findall(f".//{{{NS}}}sheetData/{{{NS}}}row")
                if any(c.tag == f"{{{NS}}}c" and (
                    c.find(f"{{{NS}}}v") is not None or c.find(f"{{{NS}}}is") is not None
                ) for c in linha)
            ]
            _ajustar_tamanho_tabela(folha, max(preenchidas, default=1))
            alteradas[mapas[nome]] = folha

        for folha in alteradas.values():
            _maiusculas_aba(folha)
        _maiusculas_compartilhadas(z, alteradas)

        alteradas["xl/workbook.xml"] = wb
        alteradas["xl/_rels/workbook.xml.rels"] = rels
        alteradas["[Content_Types].xml"] = tipos
        alteradas["xl/styles.xml"] = estilos
        memoria = io.BytesIO()
        with ZipFile(memoria, "w") as saida:
            for info in z.infolist():
                parte = alteradas.pop(info.filename, None)
                saida.writestr(info, (
                    etree.tostring(parte, encoding="UTF-8", xml_declaration=True, standalone=True)
                    if isinstance(parte, etree._Element) else parte
                    if parte is not None else z.read(info.filename)
                ))
            for caminho, arvore in alteradas.items():
                saida.writestr(caminho, etree.tostring(arvore, encoding="UTF-8", xml_declaration=True, standalone=True))
    avisos = []
    if pendentes:
        avisos.append(f"{len(pendentes)} salvaguarda(s) da APR na aba A_conferir: classificação BowTie a validar.")
    return memoria.getvalue(), avisos
