"""Persistência compartilhada do acompanhamento; nenhuma tabela do comparador é alterada."""

from contextlib import contextmanager
from datetime import date, datetime

import psycopg
from psycopg.rows import dict_row
from psycopg.types.json import Jsonb


SITUACOES = (
    "A definir", "A fazer", "Em desenvolvimento", "Em revisão",
    "Correção solicitada", "Bloqueado", "Concluído",
)
COORD_FIELDS = (
    "trabalho", "responsavel_email", "situacao", "prioridade", "inicio_previsto",
    "prazo", "inicio_real", "entrega_inicial", "conclusao_validada",
    "progresso", "proxima_acao",
)
ANALYST_FIELDS = (
    "situacao", "inicio_real", "entrega_inicial", "progresso", "proxima_acao",
)


class SemPermissao(Exception):
    pass


class EdicaoConcorrente(Exception):
    pass


def _auditavel(valor):
    if isinstance(valor, (date, datetime)):
        return valor.isoformat()
    return valor


@contextmanager
def conectar(url):
    # A senha permanece exclusivamente nos Secrets do servidor Streamlit.
    with psycopg.connect(url, sslmode="require", row_factory=dict_row) as conn:
        yield conn


def email_normalizado(email):
    return str(email or "").strip().lower()


def _acesso(conn, projeto_id, email, administradores):
    if email in administradores:
        return "Coordenador"
    linha = conn.execute(
        "SELECT papel FROM bt_membros WHERE projeto_id=%s AND email=%s",
        (projeto_id, email),
    ).fetchone()
    if linha is None:
        raise SemPermissao("Você não tem acesso a este projeto.")
    return linha["papel"]


def _coordenador(conn, projeto_id, email, administradores):
    if _acesso(conn, projeto_id, email, administradores) != "Coordenador":
        raise SemPermissao("Esta ação é reservada ao coordenador do projeto.")


def listar_projetos(url, email, administradores):
    with conectar(url) as conn:
        if email in administradores:
            return conn.execute("SELECT * FROM bt_projetos ORDER BY nome").fetchall()
        return conn.execute(
            "SELECT p.* FROM bt_projetos p JOIN bt_membros m ON m.projeto_id=p.id "
            "WHERE m.email=%s ORDER BY p.nome", (email,),
        ).fetchall()


def criar_projeto(url, email, administradores, nome, descricao):
    if email not in administradores:
        raise SemPermissao("Somente um administrador pode criar projetos.")
    nome = nome.strip()
    if not 3 <= len(nome) <= 160:
        raise ValueError("Informe um nome de projeto com 3 a 160 caracteres.")
    with conectar(url) as conn:
        p = conn.execute(
            "INSERT INTO bt_projetos(nome, descricao, criado_por) VALUES (%s,%s,%s) RETURNING id",
            (nome, descricao.strip(), email),
        ).fetchone()
        conn.execute(
            "INSERT INTO bt_membros(projeto_id,email,nome,papel) VALUES (%s,%s,%s,%s)",
            (p["id"], email, "", "Coordenador"),
        )
        conn.execute(
            "INSERT INTO bt_historico(projeto_id,acao,autor,depois) VALUES (%s,%s,%s,%s)",
            (p["id"], "Projeto criado", email, Jsonb({"nome": nome})),
        )
        return p["id"]


def papel_no_projeto(url, projeto_id, email, administradores):
    with conectar(url) as conn:
        return _acesso(conn, projeto_id, email, administradores)


def listar_membros(url, projeto_id, email, administradores):
    with conectar(url) as conn:
        _acesso(conn, projeto_id, email, administradores)
        return conn.execute(
            "SELECT email,nome,papel FROM bt_membros WHERE projeto_id=%s ORDER BY papel,nome,email",
            (projeto_id,),
        ).fetchall()


def salvar_membro(url, projeto_id, autor, administradores, email, nome, papel):
    email = email_normalizado(email)
    if "@" not in email or papel not in ("Coordenador", "Analista"):
        raise ValueError("Informe o e-mail corporativo e a função da pessoa.")
    with conectar(url) as conn:
        _coordenador(conn, projeto_id, autor, administradores)
        anterior = conn.execute(
            "SELECT nome,papel FROM bt_membros WHERE projeto_id=%s AND email=%s",
            (projeto_id, email),
        ).fetchone()
        if email == autor and papel != "Coordenador" and autor not in administradores:
            raise ValueError("O coordenador não pode retirar seu próprio acesso de coordenação.")
        conn.execute(
            "INSERT INTO bt_membros(projeto_id,email,nome,papel) VALUES (%s,%s,%s,%s) "
            "ON CONFLICT (projeto_id,email) DO UPDATE SET nome=EXCLUDED.nome,papel=EXCLUDED.papel",
            (projeto_id, email, nome.strip(), papel),
        )
        conn.execute(
            "INSERT INTO bt_historico(projeto_id,acao,autor,antes,depois) "
            "VALUES (%s,%s,%s,%s,%s)",
            (projeto_id, "Membro cadastrado ou atualizado", autor,
             Jsonb(anterior) if anterior else None,
             Jsonb({"email": email, "nome": nome.strip(), "papel": papel})),
        )


def listar_itens(url, projeto_id, email, administradores):
    with conectar(url) as conn:
        _acesso(conn, projeto_id, email, administradores)
        return conn.execute(
            "SELECT * FROM bt_itens WHERE projeto_id=%s "
            "ORDER BY length(no), no, length(cenario), cenario, id_cenario",
            (projeto_id,),
        ).fetchall()


def listar_historico(url, projeto_id, email, administradores):
    with conectar(url) as conn:
        _acesso(conn, projeto_id, email, administradores)
        return conn.execute(
            "SELECT h.registrado_em,h.acao,h.autor,i.id_cenario,h.antes,h.depois "
            "FROM bt_historico h LEFT JOIN bt_itens i ON i.id=h.item_id "
            "WHERE h.projeto_id=%s ORDER BY h.id DESC LIMIT 150",
            (projeto_id,),
        ).fetchall()


def cadastrar_itens(url, projeto_id, autor, administradores, candidatos, fonte):
    """Importa itens sem modificar os já cadastrados; valida e audita em uma transação."""
    if len(candidatos) > 2000:
        raise ValueError("Importe até 2.000 cenários por vez.")
    with conectar(url) as conn:
        _coordenador(conn, projeto_id, autor, administradores)
        adicionados = 0
        ignorados = 0
        for c in candidatos:
            identificador = str(c.get("ID") or "").strip()
            no = str(c.get("Nó") or "").strip()
            cenario = str(c.get("Cenário") or "").strip()
            sistema = " ".join(str(c.get("Sistema") or "").split())
            evento = " ".join(str(c.get("Perigo") or "").split())
            if not all((identificador, no, cenario, sistema, evento)):
                raise ValueError(f"Cenário incompleto: {identificador or '(sem ID)'}.")
            pagina = c.get("Página PDF")
            linha = conn.execute(
                "INSERT INTO bt_itens(projeto_id,id_cenario,no,cenario,sistema,evento_topo,"
                "fonte,pagina_pdf,atualizado_por) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s) "
                "ON CONFLICT (projeto_id,id_cenario) DO NOTHING RETURNING id",
                (projeto_id, identificador, no, cenario, sistema, evento,
                 fonte[:255], int(pagina) if pagina else None, autor),
            ).fetchone()
            if linha:
                adicionados += 1
                conn.execute(
                    "INSERT INTO bt_historico(projeto_id,item_id,acao,autor,depois) "
                    "VALUES (%s,%s,%s,%s,%s)",
                    (projeto_id, linha["id"], "BowTie cadastrado", autor,
                     Jsonb({"id_cenario": identificador, "fonte": fonte})),
                )
            else:
                ignorados += 1
        return adicionados, ignorados


def atualizar_item(url, projeto_id, item_id, autor, administradores, versao, alteracoes):
    if not alteracoes or set(alteracoes) - set(COORD_FIELDS):
        raise ValueError("Campos de atualização inválidos.")
    with conectar(url) as conn:
        papel = _acesso(conn, projeto_id, autor, administradores)
        antes = conn.execute(
            "SELECT * FROM bt_itens WHERE id=%s AND projeto_id=%s FOR UPDATE",
            (item_id, projeto_id),
        ).fetchone()
        if not antes:
            raise ValueError("BowTie não encontrado neste projeto.")
        if antes["versao"] != versao:
            raise EdicaoConcorrente(
                "Alguém alterou este BowTie enquanto você editava. Atualize a página e confira os novos dados."
            )
        if papel == "Analista":
            if antes["responsavel_email"] != autor:
                raise SemPermissao("Somente o responsável pode atualizar este BowTie.")
            if set(alteracoes) - set(ANALYST_FIELDS):
                raise SemPermissao("Distribuição, prioridade e conclusão cabem ao coordenador.")
            if alteracoes.get("situacao") == "Concluído":
                raise SemPermissao("Somente o coordenador valida a conclusão.")
        if alteracoes.get("situacao") is not None and alteracoes["situacao"] not in SITUACOES:
            raise ValueError("Situação inválida.")
        if "progresso" in alteracoes and alteracoes["progresso"] is not None:
            if not 0 <= int(alteracoes["progresso"]) <= 100:
                raise ValueError("O progresso deve estar entre 0 e 100%.")
        if "responsavel_email" in alteracoes and alteracoes["responsavel_email"]:
            responsavel = conn.execute(
                "SELECT 1 FROM bt_membros WHERE projeto_id=%s AND email=%s AND papel='Analista'",
                (projeto_id, alteracoes["responsavel_email"]),
            ).fetchone()
            if not responsavel:
                raise ValueError("Escolha um analista cadastrado neste projeto.")
        if papel == "Analista" and antes["situacao"] == "Concluído":
            raise SemPermissao("A conclusão foi validada; peça uma reabertura ao coordenador.")
        if alteracoes.get("situacao") == "Concluído" and not alteracoes.get(
            "conclusao_validada", antes["conclusao_validada"]
        ):
            raise ValueError("Informe a data de conclusão validada.")
        if alteracoes.get("situacao") not in (None, "Concluído") and antes["situacao"] == "Concluído":
            alteracoes["conclusao_validada"] = None
        mudancas = {k: v for k, v in alteracoes.items() if antes[k] != v}
        if not mudancas:
            return False
        # Nomes das colunas provêm exclusivamente da lista fixa COORD_FIELDS.
        atribuicoes = ", ".join(f"{campo}=%s" for campo in mudancas)
        conn.execute(
            f"UPDATE bt_itens SET {atribuicoes}, atualizado_por=%s, "
            "atualizado_em=now(), versao=versao+1 WHERE id=%s AND projeto_id=%s",
            (*mudancas.values(), autor, item_id, projeto_id),
        )
        conn.execute(
            "INSERT INTO bt_historico(projeto_id,item_id,acao,autor,antes,depois) "
            "VALUES (%s,%s,%s,%s,%s,%s)",
            (projeto_id, item_id, "BowTie atualizado", autor,
             Jsonb({k: _auditavel(antes[k]) for k in mudancas}),
             Jsonb({k: _auditavel(v) for k, v in mudancas.items()})),
        )
        return True
