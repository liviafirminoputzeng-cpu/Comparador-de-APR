"""Testes das regras de acesso e da proteção contra edições simultâneas."""

import copy
import json
import sys
import types
import unittest
from contextlib import contextmanager
from datetime import date
from pathlib import Path
from unittest.mock import patch


psycopg = types.ModuleType("psycopg")
rows = types.ModuleType("psycopg.rows")
rows.dict_row = object()
types_mod = types.ModuleType("psycopg.types")
json_mod = types.ModuleType("psycopg.types.json")


def Jsonb(value):
    json.dumps(value)  # replica a exigência de serialização do driver
    return value


json_mod.Jsonb = Jsonb
sys.modules.update({"psycopg": psycopg, "psycopg.rows": rows,
                    "psycopg.types": types_mod, "psycopg.types.json": json_mod})
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import acompanhamento_db as db  # noqa: E402


class Cursor:
    def __init__(self, row=None):
        self.row = row

    def fetchone(self):
        return self.row


class BancoFalso:
    def __init__(self):
        self.item = dict(id="i1", projeto_id="p1", id_cenario="N1.2",
                         responsavel_email="analista@empresa.com", situacao="A fazer",
                         versao=1, progresso=None, proxima_acao="", inicio_real=None,
                         entrega_inicial=None, conclusao_validada=None, trabalho="Novo",
                         prioridade=None, prazo=None, inicio_previsto=None)
        self.historico = []

    def execute(self, sql, params):
        if "SELECT papel FROM bt_membros" in sql:
            projeto, email = params
            return Cursor({"papel": "Analista"} if
                          projeto == "p1" and email == "analista@empresa.com" else None)
        if "SELECT * FROM bt_itens" in sql:
            return Cursor(copy.deepcopy(self.item) if params == ("i1", "p1") else None)
        if sql.startswith("UPDATE bt_itens"):
            for field, value in zip(sql.split("SET ")[1].split(", atualizado_por")[0].split(", "),
                                    params[:-3]):
                self.item[field.split("=")[0]] = value
            self.item["versao"] += 1
            return Cursor()
        if "INSERT INTO bt_historico" in sql:
            self.historico.append(params)
            return Cursor()
        raise AssertionError(sql)


class RegrasDeAcompanhamento(unittest.TestCase):
    def setUp(self):
        self.banco = BancoFalso()

        @contextmanager
        def conectar(_):
            yield self.banco

        self.patcher = patch.object(db, "conectar", conectar)
        self.patcher.start()

    def tearDown(self):
        self.patcher.stop()

    def test_analista_nao_conclui_nem_altera_distribuicao(self):
        with self.assertRaises(db.SemPermissao):
            db.atualizar_item("url", "p1", "i1", "analista@empresa.com", set(), 1,
                             {"situacao": "Concluído"})
        with self.assertRaises(db.SemPermissao):
            db.atualizar_item("url", "p1", "i1", "analista@empresa.com", set(), 1,
                             {"prazo": date(2026, 10, 1)})

    def test_outro_projeto_e_outro_analista_sao_bloqueados(self):
        with self.assertRaises(db.SemPermissao):
            db.atualizar_item("url", "p2", "i1", "analista@empresa.com", set(), 1,
                             {"situacao": "Em desenvolvimento"})
        with self.assertRaises(db.SemPermissao):
            db.atualizar_item("url", "p1", "i1", "outro@empresa.com", set(), 1,
                             {"situacao": "Em desenvolvimento"})

    def test_conflito_apos_primeiro_salvamento_e_auditoria(self):
        db.atualizar_item("url", "p1", "i1", "analista@empresa.com", set(), 1,
                         {"situacao": "Em desenvolvimento", "inicio_real": date(2026, 9, 29)})
        self.assertEqual(self.banco.item["versao"], 2)
        self.assertEqual(self.banco.historico[0][-1]["inicio_real"], "2026-09-29")
        with self.assertRaises(db.EdicaoConcorrente):
            db.atualizar_item("url", "p1", "i1", "analista@empresa.com", set(), 1,
                             {"situacao": "Em revisão"})

    def test_coordenador_so_conclui_com_data(self):
        with self.assertRaises(ValueError):
            db.atualizar_item("url", "p1", "i1", "coord@empresa.com",
                             {"coord@empresa.com"}, 1, {"situacao": "Concluído"})
        db.atualizar_item("url", "p1", "i1", "coord@empresa.com",
                         {"coord@empresa.com"}, 1,
                         {"situacao": "Concluído", "conclusao_validada": date(2026, 9, 29)})
        self.assertEqual(self.banco.item["situacao"], "Concluído")


if __name__ == "__main__":
    unittest.main()
