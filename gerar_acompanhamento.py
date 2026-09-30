"""Preenche uma cópia do modelo de acompanhamento a partir da APR atualizada."""

import io
import re
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook.properties import CalcProperties

from comparacao import candidato_bowtie


MODELO = Path(__file__).resolve().parent / "assets" / "modelo_acompanhamento.xlsx"
LIMITE_MODELO = 200  # O painel do modelo soma as linhas 9 a 208.


def _ordem(valor):
    partes = re.findall(r"\d+", str(valor or ""))
    return tuple(int(p) for p in partes) if partes else (10**9,)


def candidatos_da_apr(cenarios):
    candidatos = [c for c in cenarios if candidato_bowtie(c)]
    candidatos.sort(key=lambda c: (_ordem(c.get("Nó")), _ordem(c.get("Cenário")),
                                   str(c.get("ID", ""))))
    ids = [str(c.get("ID") or "").strip() for c in candidatos]
    if any(not identificador for identificador in ids) or len(set(ids)) != len(ids):
        raise ValueError("A APR tem candidatos com IDs ausentes ou repetidos. Confira a extração.")
    if len(candidatos) > LIMITE_MODELO:
        raise ValueError(f"O modelo comporta até {LIMITE_MODELO} candidatos por projeto.")
    return candidatos


def gerar_planilha_acompanhamento(apr_atualizada, projeto):
    """Retorna bytes .xlsx. Não presume responsável, trabalho ou progresso."""
    cenarios = candidatos_da_apr(apr_atualizada["cenarios"])
    if not cenarios:
        raise ValueError("A APR atualizada não contém candidatos a BowTie para acompanhar.")
    wb = load_workbook(MODELO)
    try:
        controle = wb["Controle"]
        painel = wb["Painel"]
        # Mantém cabeçalhos, cores, validações, fórmulas e o gráfico do modelo.
        for linha in range(9, 209):
            for col in range(1, 21):
                controle.cell(linha, col).value = None
        estilo = [copy(controle.cell(9, col)._style) for col in range(1, 21)]
        for indice, cenario in enumerate(cenarios, start=9):
            dados = [
                str(cenario.get("ID", "")).strip(),
                cenario.get("Nó"), cenario.get("Cenário"),
                " ".join(str(cenario.get("Sistema") or "").split()),
                " ".join(str(cenario.get("Perigo") or "").split()),
            ]
            if not all(str(x or "").strip() for x in dados):
                raise ValueError(f'Dados incompletos no cenário {cenario.get("ID", "(sem ID)")}.')
            for col, valor in enumerate(dados, start=1):
                celula = controle.cell(indice, col)
                celula._style = copy(estilo[col - 1])
                celula.value = valor
            controle.cell(indice, 10).value = "A definir"
            pagina = cenario.get("Página PDF")
            controle.cell(indice, 20).value = pagina if pagina else None
        controle.tables["BowTiesRevK"].ref = f"A8:T{8 + len(cenarios)}"
        relatorio = str(apr_atualizada.get("relatorio") or "Relatório não identificado")
        fonte = str(apr_atualizada.get("nome_arquivo") or "APR atualizada")
        painel["B3"] = f"{relatorio} · {len(cenarios)} candidatos iniciais · {projeto}"
        controle["A3"] = (
            "Amarelo: preencher pela equipe. As células cinza e azul vêm da APR atualizada. "
            "Defina produção, responsáveis e andamento nas colunas amarelas."
        )
        controle["A5"] = f"Fonte: {fonte} · {relatorio}. IDs, nós, cenários, sistemas, eventos topo e páginas vêm da APR."
        controle["A6"] = (
            f"Os {len(cenarios)} itens começam como “A definir”. Classifique o trabalho como Novo ou Correção. "
            "Para os novos, acompanhe a produção da planilha e atribua um estagiário."
        )
        wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
        saida = io.BytesIO()
        wb.save(saida)
        return saida.getvalue(), len(cenarios)
    finally:
        wb.close()
