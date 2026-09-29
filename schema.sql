-- Execute uma vez no PostgreSQL compartilhado antes de abrir o acompanhamento.
CREATE EXTENSION IF NOT EXISTS pgcrypto;

CREATE TABLE IF NOT EXISTS bt_projetos (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    nome text NOT NULL UNIQUE CHECK (length(trim(nome)) BETWEEN 3 AND 160),
    descricao text NOT NULL DEFAULT '',
    criado_por text NOT NULL,
    criado_em timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS bt_membros (
    projeto_id uuid NOT NULL REFERENCES bt_projetos(id),
    email text NOT NULL CHECK (email = lower(trim(email))),
    nome text NOT NULL DEFAULT '',
    papel text NOT NULL CHECK (papel IN ('Coordenador', 'Analista')),
    PRIMARY KEY (projeto_id, email)
);

CREATE TABLE IF NOT EXISTS bt_itens (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    projeto_id uuid NOT NULL REFERENCES bt_projetos(id),
    id_cenario text NOT NULL,
    no text NOT NULL,
    cenario text NOT NULL,
    sistema text NOT NULL,
    evento_topo text NOT NULL,
    fonte text NOT NULL DEFAULT '',
    pagina_pdf integer,
    trabalho text CHECK (trabalho IN ('Novo', 'Correção')),
    responsavel_email text,
    situacao text NOT NULL DEFAULT 'A definir' CHECK (situacao IN
       ('A definir','A fazer','Em desenvolvimento','Em revisão',
        'Correção solicitada','Bloqueado','Concluído')),
    prioridade text CHECK (prioridade IN ('Alta','Média','Baixa')),
    inicio_previsto date,
    prazo date,
    inicio_real date,
    entrega_inicial date,
    conclusao_validada date,
    progresso integer CHECK (progresso BETWEEN 0 AND 100),
    proxima_acao text NOT NULL DEFAULT '',
    atualizado_por text NOT NULL,
    atualizado_em timestamptz NOT NULL DEFAULT now(),
    versao integer NOT NULL DEFAULT 1,
    UNIQUE (projeto_id, id_cenario),
    FOREIGN KEY (projeto_id, responsavel_email)
        REFERENCES bt_membros (projeto_id, email)
);

CREATE INDEX IF NOT EXISTS bt_itens_projeto ON bt_itens (projeto_id, no, cenario);

CREATE TABLE IF NOT EXISTS bt_historico (
    id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    projeto_id uuid NOT NULL REFERENCES bt_projetos(id),
    item_id uuid REFERENCES bt_itens(id),
    acao text NOT NULL,
    autor text NOT NULL,
    antes jsonb,
    depois jsonb,
    registrado_em timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS bt_historico_projeto ON bt_historico (projeto_id, registrado_em DESC);
