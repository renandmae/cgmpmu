# app.py - Sistema completo de registro de horas (Flask + SQLite)
# Salve este arquivo como app.py e rode:
# python -m venv venv
# source venv/bin/activate  (ou venv\Scripts\activate no Windows)
# pip install flask
# python app.py

from flask import Flask, render_template_string, request, redirect, session, send_file, Response, jsonify
import io
import csv
import json
from datetime import date
import os
import psycopg2
import psycopg2.extras
from psycopg2 import IntegrityError

def data_padrao_2026():
    hoje = date.today()
    if hoje.year == 2026:
        return hoje.isoformat()
    return "2026-01-01"

from datetime import datetime, date

def fmt(d):
    if not d:
        return ""

    # Se vier como date ou datetime (PostgreSQL / Supabase)
    if isinstance(d, (date, datetime)):
        return d.strftime("%d/%m/%Y")

    # Se vier como string yyyy-mm-dd
    if isinstance(d, str):
        try:
            return datetime.strptime(d[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
        except:
            return d

    return str(d)


app = Flask(__name__)
app.secret_key = 'troque_esta_chave'


# -------------------------
# Helpers
# -------------------------
DATABASE_URL = os.environ.get(
    "DATABASE_URL",
    "postgresql://postgres.rjdfnufoeulbmotxjqwd:cgm2025@aws-1-sa-east-1.pooler.supabase.com:5432/postgres"
)

def get_db():
    return psycopg2.connect(
        DATABASE_URL,
        cursor_factory=psycopg2.extras.RealDictCursor
    )

def init_db():
    con = get_db()
    cur = con.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS colaboradores (
        id SERIAL PRIMARY KEY,
        nome TEXT,
        login TEXT UNIQUE,
        senha TEXT,
        perfil TEXT
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS projeto_paint (
        id SERIAL PRIMARY KEY,
        classificacao TEXT,
        item_paint TEXT UNIQUE,
        tipo_atividade TEXT,
        objeto TEXT,
        objetivo_geral TEXT,
        dt_ini DATE,
        dt_fim DATE,
        hh_atual INTEGER DEFAULT 0
    );
    """)

        # OS
    cur.execute('''
CREATE TABLE IF NOT EXISTS os (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo TEXT UNIQUE,
    item_paint TEXT,
    resumo TEXT,
    unidade TEXT,
    supervisao TEXT,
    coordenacao TEXT,
    equipe TEXT,
    observacao TEXT,
    status TEXT,
    plan INTEGER DEFAULT 0,
    exec INTEGER DEFAULT 0,
    rp INTEGER DEFAULT 0,
    rf INTEGER DEFAULT 0,
    dt_conclusao TEXT
)
''')

    cur.execute("""
    CREATE TABLE IF NOT EXISTS delegacoes (
        id SERIAL PRIMARY KEY,
        requisicoes TEXT,
        os_codigo TEXT,
        colaborador_id INTEGER REFERENCES colaboradores(id),
        data_inicio DATE,
        status TEXT DEFAULT 'Em Andamento',
        grau TEXT,
        data_fim DATE,
        criterio TEXT
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS horas (
        id SERIAL PRIMARY KEY,
        colaborador_id INTEGER REFERENCES colaboradores(id),
        data DATE,
        item_paint TEXT,
        os_codigo TEXT,
        atividade TEXT,
        duracao TEXT,
        hora_inicio TIME,
        hora_fim TIME,
        duracao_minutos INTEGER,
        delegacao_id INTEGER REFERENCES delegacoes(id),
        observacoes TEXT
    );
    """)
    
    cur.execute("""
    CREATE TABLE IF NOT EXISTS atendimentos (
        id SERIAL PRIMARY KEY,
    
        -- vínculo com o lançamento de horas
        hora_id INTEGER REFERENCES horas(id) ON DELETE CASCADE,
        colaborador_id INTEGER REFERENCES colaboradores(id),
    
        -- identificação da OS
        os_codigo TEXT,
        os_resumo TEXT,
    
        -- dados do atendimento
        responsaveis_consultoria TEXT,   -- "Nome A, Nome B"
        macro TEXT,
        diretoria TEXT,
        atividade TEXT,
        data_consultoria DATE,
        assunto TEXT,
        participantes_externos TEXT,
        entidades TEXT,                  -- "CM, SEGOV, SMGAS"
        meio_contato TEXT,               -- Presencial | Email | Telefone
        observacao TEXT,
    
        -- dados derivados da hora
        duracao_minutos INTEGER,
        data_lancamento DATE,
    
        created_at TIMESTAMP DEFAULT NOW()
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS consultorias (
        id SERIAL PRIMARY KEY,
    
        -- vínculo com o lançamento de horas
        hora_id INTEGER REFERENCES horas(id) ON DELETE CASCADE,
        colaborador_id INTEGER REFERENCES colaboradores(id),
    
        -- identificação da OS
        os_codigo TEXT,
        os_resumo TEXT,
    
        -- dados da consultoria
        responsaveis TEXT,   -- "Nome A, Nome B"
        data_consul DATE,
        assunto TEXT,
        secretarias TEXT,                  -- "CM, SEGOV, SMGAS"
        meio TEXT,               -- Presencial | Email | Telefone
        tipo TEXT,
        palavras_dhave TEXT,
        num_oficio TEXT,
        observacao TEXT,
    
        -- dados derivados da hora
        duracao_minutos INTEGER,
        data_lancamento DATE,
    
        created_at TIMESTAMP DEFAULT NOW()
    );
    """)

    con.commit()
    con.close()

# -------------------------
# Seed inicial
# -------------------------
COLABS = [
    "Alexandra Kátia", "Ana Paula Vilela", "Anatole Reis", "Aurélio Feitosa", "Caprice Cardoso", "Carlos Augusto",
    "Fernanda Lima", "Grazielle Carrijo", "Laianne Fogaça", "Marcelo Marques", "Maria Cristina", "Mariana Mota",
    "Mariana Cavanha", "Michelle Terêncio", "Paula Renata", "Paulo Sérgio", "Priscilla da Silva", "Syria Galvão",
    "Thamy Ponciano", "Tiago Pinheiro"
]

ITEMS_PAINT = [
    'O-1', 'O-2', 'O-3', 'O-4', 'O-5', 'O-6', 'O-7',
    'P-1', 'P-2', 'P-3', 'P-4', 'P-5', 'P-6', 'P-7', 'P-8', 'P-10', 'P-11', 'P-12', 'P-13', 'P-14', 'P-15', 'P-16',
    'P-17', 'P-18', 'P-19', 'P-20', 'P-21', 'P-22', 'P-23', 'P-24', 'P-25', 'P-26', 'P-27', 'P-28', 'P-29', 'P-30',
    'P-31', 'P-32', 'P-33', 'P-34', 'P-35', 'P-36', 'P-37', 'P-38', 'P-39', 'P-40', 'P-41', 'P-42', 'P-43', 'P-44',
    'P-45', 'P-46', 'P-47', 'P-48'
]


# --------- Seed compatível com Flask 3 ---------

def executar_seed():
    init_db()
    con = get_db()
    cur = con.cursor()

    # seed colaboradores
    for c in COLABS:
        login = c.lower().replace(' ', '.')
        try:
            cur.execute('INSERT INTO colaboradores (nome,login,senha,perfil) VALUES (%s,%s,%s,%s)',
                        (c, login, '123', 'comum'))
        except IntegrityError:
            con.rollback()

    # admin
    try:
        cur.execute('INSERT INTO colaboradores (nome,login,senha,perfil) VALUES (%s,%s,%s,%s)',
                    ('Renan Justino', 'renan.justino', '123', 'admin'))
    except IntegrityError:
        con.rollback()

    # projeto_paint base
    for it in ITEMS_PAINT:
        try:
            cur.execute(
                'INSERT INTO projeto_paint (classificacao, item_paint, tipo_atividade, objeto, objetivo_geral, dt_ini, dt_fim, hh_atual) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)',
                ('Obrigatório', it, '', '', '', None, None, 0))
        except IntegrityError:
            con.rollback()

    con.commit()
    con.close()


# -------------------------
# Templates base (usamos render_template_string para facilitar entrega única)
# -------------------------
BASE = """
<!doctype html>
<html>
<head>
<style>

    /* ---------------------- TEMA GLOBAL ---------------------- */
    body {
        background: #e9f1fb;                 /* azul clarinho elegante */
        color: #1d2a3a;                      /* cinza-azulado escuro */
        font-family: "Segoe UI", sans-serif;
        margin: 0;
        padding: 0;
    }

    h1, h2, h3, h4 {
        color: #1e4f9c;
        margin-bottom: 10px;
    }

    a { 
        color: #1e4f9c; 
        text-decoration: none; 
    }
    a:hover { 
        color: #3b7ae0; 
    }

    /* ---------------------- CONTAINER ---------------------- */
    .container {
        max-width: 1100px;
        margin: auto;
        padding: 25px;
    }

    /* ---------------------- BOTÕES ---------------------- */
    .btn {
        background: #1e74d9;
        color: white !important;
        padding: 8px 16px;
        border-radius: 6px;
        border: none;
        font-size: 14px;
        cursor: pointer;
        display: inline-block;
        text-align: center;
        transition: 0.2s;
    }

    .btn:hover {
        background: #3a8af0;
    }

    .btn-danger {
        background: #d9534f !important;
    }

    .btn-danger:hover {
        background: #e6736f !important;
    }

    /* ---------------------- INPUTS / SELECTS / TEXTAREA ---------------------- */
    input, select, textarea {
        background: #ffffff;
        border: 1px solid #c3d4ea;
        color: #1d2a3a;
        padding: 8px;
        border-radius: 6px;
        width: 100%;
        margin-top: 4px;
        margin-bottom: 12px;
        font-size: 14px;
    }

    input:focus, select:focus, textarea:focus {
        border-color: #1e74d9;
        outline: none;
        box-shadow: 0 0 0 3px rgba(30, 116, 217, 0.28);
    }

    textarea {
        min-height: 100px;
        resize: vertical;
    }

    /* ---------------------- CARD ---------------------- */
    .card {
        background: #ffffff;
        border: 1px solid #d3e0f0;
        padding: 20px;
        border-radius: 12px;
        margin-bottom: 25px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);

        /* texto corrigido */
        color: #1d2a3a !important;
    }

    /* ---------------------- TABELAS ---------------------- */
    table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 15px;
        background: #ffffff;
        border-radius: 12px;
        overflow: hidden;
        font-size: 14px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
    }

    th {
        background: #dce8fb;
        padding: 10px;
        text-align: left;
        color: #1e4f9c;
        font-weight: 600;
        border-bottom: 1px solid #c7d5ee;
    }

    td {
        padding: 9px 12px;
        border-top: 1px solid #e1e8f2;
        color: #1d2a3a;
    }

    tr:nth-child(even) {
        background: #f4f8ff;
    }

    tr:hover {
        background: #e9f2ff;
    }

    /* ---------------------- TOPBAR ---------------------- */
    .topbar {
        background: #1e74d9;
        padding: 15px 25px;
        color: white !important;
        font-size: 18px;
        font-weight: bold;
        display: flex;
        align-items: center;
        column-gap: 20px;
        flex-wrap: wrap;
    }

    .topbar a {
        color: white !important;
        font-weight: 500;
        opacity: 0.95;
    }

    .topbar a:hover {
        opacity: 1;
    }

    /* ---------------------- MENU LINKS (corrigido) ---------------------- */
.menu-links a {
    margin-right: 16px;
    padding: 7px 12px;
    border-radius: 6px;

    /* Fundo mais sólido e claro */
    background: #ffffff;

    /* Texto azul escuro para boa leitura */
    color: #1e4f9c !important;

    font-size: 14px;
    font-weight: 600;
    border: 1px solid #c5d7f2;

    transition: 0.2s;
}

.menu-links a:hover {
    background: #d9e8ff;        /* azul claro no hover */
    color: #124a99 !important;  /* azul mais forte */
    border-color: #a8c6f2;
}
/* ---------------------- coluna requisicoes largura ---------------------- */
    .col-requisicoes {
        max-width: 320px;
        white-space: normal;
        word-break: break-word;
    }
    
</style>


  <meta charset='utf-8'>
  <title>Sistema de Horas</title>
  <style>
    body{font-family: Arial, Helvetica, sans-serif;max-width:900px;margin:20px auto;color:#222}
    header{display:flex;justify-content:space-between;align-items:center}
    nav a {
    margin-right: 18px;
    font-weight: 500;
    display: inline-block;
}
    table{width:100%;border-collapse:collapse;margin-top:10px}
    th,td{border:1px solid #ddd;padding:6px;text-align:left}
    .small{font-size:0.9em;color:#555}
    .btn{display:inline-block;padding:6px 10px;border-radius:6px;background:#1976d2;color:white;text-decoration:none}
    form div{margin:6px 0}
    input[type=date], input[type=text], select, input[type=password], input[type=time]{padding:6px}
  </style>
</head>
<body>
<header>
  <div>
    <h2>Sistema de Registro de Horas</h2>
    {% if user %}
      <div class='small'>Logado como: <strong>{{user}}</strong> ({{perfil}})</div>
    {% endif %}
  </div>
  <div>
    {% if user %}
      <nav class="menu-links">
        <a href='/menu'>Menu</a>
        <a href='/lancar'>Lançar Horas</a>
        <a href='/relatorios'>Relatórios</a>
        {% if perfil != 'admin' %}
            <a href="/minhas_delegacoes">Requisições Delegadas</a>
        {% endif %}
        {% if perfil == 'admin' %}
          <a href='/atendimentos'>Todos Atendimentos</a>
        {% else %}
          <a href='/atendimentos'>Meus Atendimentos</a>
        {% endif %}
        {% if perfil == 'admin' %}
          <a href='/consultorias'>Todas Consultorias</a>
        {% else %}
          <a href='/consultorias'>Minhas Consultorias</a>
        {% endif %}
        {% if perfil=='admin' %}
            <a href='/colaboradores'>Colaboradores</a>
            <a href='/paint'>Projetos PAINT</a>
            <a href='/os'>O.S</a>
            <a href="/delegar">Delegar Requisições</a>
            <a href="/delegacoes">Delegações Cadastradas</a>
            <a href='/admin_projetos'>Gerenciar Projetos</a>
            <a href='/visao'>Visão Consolidada</a>
        {% endif %}

        <a href='/logout'>Sair</a>
      </nav>
    {% endif %}
  </div>
</header>
<hr>
<div>
  {% block content %}{% endblock %}
</div>
</body>
</html>
"""

@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        login = request.form.get('login')
        senha = request.form.get('senha')

        con = get_db()
        cur = con.cursor()
        cur.execute(
            'SELECT id,nome,perfil FROM colaboradores WHERE login=%s AND senha=%s',
            (login, senha)
        )
        row = cur.fetchone()
        con.close()

        if row:
            session['user'] = row['nome']
            session['perfil'] = row['perfil']
            session['user_id'] = row['id']
            return redirect('/menu')
        else:
            conteudo = f"""
                {HEADER_LOGIN}
                <h3 style="text-align:center;">Login</h3>
                <p style="color:red;text-align:center;">Login inválido</p>
                {LOGIN_FORM}
            """

            return render_template_string(
                BASE.replace('{% block content %}{% endblock %}', conteudo),
                user=None,
                perfil=None
            )

    conteudo = f"""
        {HEADER_LOGIN}
        <h3 style="text-align:center;">Login</h3>
        {LOGIN_FORM}
    """

    return render_template_string(
        BASE.replace('{% block content %}{% endblock %}', conteudo),
        user=None,
        perfil=None
    )

LOGIN_FORM = """
<form method="post" style="max-width:380px;margin:0 auto;">
    <div>
        <label>Login</label>
        <input name="login" required>
    </div>

    <div>
        <label>Senha</label>
        <input type="password" name="senha" required>
    </div>

    <div style="margin-top:15px;">
        <button class="btn" style="width:100%;">Entrar</button>
    </div>
</form>
"""

HEADER_LOGIN = """
<div style="width:100%; text-align:center; margin-bottom:30px;">
    <img
        src="https://i.ibb.co/gFv5XWJp/8-Controladoria-geral.png"
        alt="Controladoria-Geral do Município"
        style="max-height:90px; margin-bottom:15px;"
    >

    <h1 style="
        font-size:28px;
        font-weight:800;
        color:#000;
        margin:0;
        letter-spacing:1px;
    ">
        Controladoria-Geral<br>
    </h1>

    <div style="
        width:100%;
        height:6px;
        background:#000;
        margin-top:15px;
    "></div>
</div>
"""



@app.route('/logout')
def logout():
    session.clear()
    return redirect('/')


@app.route('/menu')
def menu():
    if 'user' not in session:
        return redirect('/')
    content = """
    <h3>Menu</h3>
    <ul>
      <li><a href='/lancar'>Lançar horas</a></li>
      <li><a href='/relatorios'>Relatórios</a></li>
    </ul>
    """
    return render_template_string(BASE.replace('{% block content %}{% endblock %}', content), user=session['user'],
                                  perfil=session['perfil'])

# -------------------------
# Colaboradores (admin)
# -------------------------
@app.route('/colaboradores', methods=['GET', 'POST'])
def colaboradores():
    if 'user' not in session:
        return redirect('/')
    if session['perfil'] != 'admin':
        return 'Acesso negado'

    ADMIN_MASTER_ID = 21
    admin_master = session.get("user_id") == ADMIN_MASTER_ID

    con = get_db()
    cur = con.cursor()
    acao = request.form.get('acao')

    # -------------------------
    # BLOQUEIO DE AÇÕES
    # -------------------------
    if acao and not admin_master:
        con.close()
        return "Ação não permitida para este usuário"

    # -------------------------
    # NOVO COLABORADOR
    # -------------------------
    if acao == 'novo':
        cur.execute("""
            INSERT INTO colaboradores (nome, login, senha, perfil)
            VALUES (%s,%s,%s,%s)
        """, (
            request.form['nome'],
            request.form['login'],
            request.form['senha'],
            request.form['perfil']
        ))
        con.commit()

    # -------------------------
    # EDITAR COLABORADOR
    # -------------------------
    elif acao == 'editar':
        cid = request.form['id']

        cur.execute("SELECT nome FROM colaboradores WHERE id=%s", (cid,))
        antigo = cur.fetchone()
        if not antigo:
            con.close()
            return "Colaborador não encontrado"

        cur.execute("""
            UPDATE colaboradores
            SET nome=%s, login=%s, perfil=%s
            WHERE id=%s
        """, (
            request.form['nome'],
            request.form['login'],
            request.form['perfil'],
            cid
        ))

        if request.form.get('senha'):
            cur.execute(
                "UPDATE colaboradores SET senha=%s WHERE id=%s",
                (request.form['senha'], cid)
            )

        if antigo['nome'] != request.form['nome']:
            cur.execute("""
                UPDATE atendimentos
                SET responsaveis_consultoria =
                    REPLACE(responsaveis_consultoria, %s, %s)
                WHERE responsaveis_consultoria ILIKE %s
            """, (
                antigo['nome'],
                request.form['nome'],
                f"%{antigo['nome']}%"
            ))

        con.commit()

    # -------------------------
    # EXCLUIR COLABORADOR
    # -------------------------
    elif acao == 'excluir':
        cid = request.form['id']

        cur.execute(
            "SELECT COUNT(*) AS qtd FROM horas WHERE colaborador_id=%s",
            (cid,)
        )
        if cur.fetchone()['qtd'] > 0:
            con.close()
            return "Não é possível excluir colaborador com horas lançadas"

        cur.execute("DELETE FROM delegacoes WHERE colaborador_id=%s", (cid,))
        cur.execute("DELETE FROM atendimentos WHERE colaborador_id=%s", (cid,))
        cur.execute("DELETE FROM colaboradores WHERE id=%s", (cid,))
        con.commit()

    # -------------------------
    # LISTAGEM
    # -------------------------
    cur.execute("""
        SELECT
            c.id,
            c.nome,
            c.login,
            c.perfil,
            COALESCE(
                ARRAY_AGG(DISTINCT h.item_paint)
                FILTER (WHERE h.item_paint IS NOT NULL),
                '{}'
            ) AS projetos,
            COALESCE(SUM(h.duracao_minutos), 0) AS total_minutos
        FROM colaboradores c
        LEFT JOIN horas h ON h.colaborador_id = c.id
        GROUP BY c.id
        ORDER BY c.nome
    """)
    colaboradores = cur.fetchall()
    con.close()

    # -------------------------
    # HTML
    # -------------------------
    html = "<h3>Colaboradores</h3>"

    # ➕ NOVO (somente admin 21)
    if admin_master:
        html += """
        <details style="margin-bottom:15px;">
            <summary style="cursor:pointer;">➕ Novo Colaborador</summary>
            <form method="post">
                <input type="hidden" name="acao" value="novo">
                <input name="nome" placeholder="Nome" required>
                <input name="login" placeholder="Login" required>
                <input name="senha" type="password" placeholder="Senha" required>
                <select name="perfil">
                    <option value="comum">comum</option>
                    <option value="admin">admin</option>
                </select>
                <button class="btn">Cadastrar</button>
            </form>
        </details>
        """

    html += """
    <table>
        <tr>
            <th>Nome</th>
            <th>Login</th>
            <th>Perfil</th>
            <th>Projetos PAINT</th>
            <th>Total Horas</th>
            <th>Ações</th>
        </tr>
    """

    for c in colaboradores:
        hh = c['total_minutos'] // 60
        mm = c['total_minutos'] % 60
        total = f"{hh:02d}:{mm:02d}"
        projetos = ", ".join(c['projetos']) if c['projetos'] else "-"

        html += f"""
        <tr>
            <td><a href="/colaborador/{c['id']}">{c['nome']}</a></td>
            <td>{c['login']}</td>
            <td>{c['perfil']}</td>
            <td>{projetos}</td>
            <td>{total}</td>
            <td>
        """

        # ✏️ / 🗑 somente admin 21
        if admin_master:
            html += f"""
                <details style="display:inline-block;">
                    <summary style="cursor:pointer;">✏️</summary>
                    <form method="post">
                        <input type="hidden" name="acao" value="editar">
                        <input type="hidden" name="id" value="{c['id']}">
                        <input name="nome" value="{c['nome']}" required>
                        <input name="login" value="{c['login']}" required>
                        <input name="senha" type="password" placeholder="Nova senha">
                        <select name="perfil">
                            <option value="comum" {'selected' if c['perfil']=='comum' else ''}>comum</option>
                            <option value="admin" {'selected' if c['perfil']=='admin' else ''}>admin</option>
                        </select>
                        <button class="btn">Salvar</button>
                    </form>
                </details>

                <form method="post" style="display:inline;"
                      onsubmit="return confirm('Excluir colaborador?');">
                    <input type="hidden" name="acao" value="excluir">
                    <input type="hidden" name="id" value="{c['id']}">
                    <button class="btn" style="background:#c00;">🗑</button>
                </form>
            """
        else:
            html += "—"

        html += "</td></tr>"

    html += "</table>"

    return render_template_string(
        BASE.replace('{% block content %}{% endblock %}', html),
        user=session['user'],
        perfil=session['perfil']
    )

# -------------------------
# Detalhes do Colaborador (admin)
# -------------------------
@app.route('/colaborador/<int:cid>')
def colaborador_detalhes(cid):
    if 'user' not in session:
        return redirect('/')
    if session['perfil'] != 'admin':
        return 'Acesso negado'

    con = get_db()
    cur = con.cursor()

    # -------------------------
    # Nome do colaborador
    # -------------------------
    cur.execute("SELECT nome FROM colaboradores WHERE id=%s", (cid,))
    col = cur.fetchone()
    if not col:
        con.close()
        return "Colaborador não encontrado"

    nome = col['nome']

    # -------------------------
    # Total por OS (agrupado)
    # -------------------------
    cur.execute("""
        SELECT
            os_codigo,
            item_paint,
            SUM(duracao_minutos) AS minutos
        FROM horas
        WHERE colaborador_id = %s
        GROUP BY os_codigo, item_paint
        ORDER BY item_paint, os_codigo
    """, (cid,))
    por_os = cur.fetchall()

    # -------------------------
    # Total por Item PAINT
    # -------------------------
    cur.execute("""
        SELECT
            item_paint,
            SUM(duracao_minutos) AS minutos
        FROM horas
        WHERE colaborador_id = %s
        GROUP BY item_paint
        ORDER BY item_paint
    """, (cid,))
    por_paint = cur.fetchall()

    con.close()

    # -------------------------
    # Helper
    # -------------------------
    def minutos_para_hhmm(minutos):
        minutos = int(minutos or 0)
        return f"{minutos//60:02d}:{minutos%60:02d}"

    # -------------------------
    # HTML
    # -------------------------
    html = f"<h3>Horas de {nome}</h3>"
    html += "<div style='display:flex; gap:40px; align-items:flex-start;'>"

    # ---- Tabela PAINT
    html += """
    <div style='width:45%;'>
        <h4>Total por Item PAINT</h4>
        <table>
            <tr><th>Item</th><th>Total (HH:MM)</th></tr>
    """
    for r in por_paint:
        html += f"""
        <tr>
            <td>{r['item_paint'] or '-'}</td>
            <td>{minutos_para_hhmm(r['minutos'])}</td>
        </tr>
        """
    html += "</table></div>"

    # ---- Tabela OS
    html += """
    <div style='width:50%;'>
        <h4>Total por O.S. (agrupado por Item PAINT)</h4>
        <table>
            <tr><th>OS</th><th>Item PAINT</th><th>Total (HH:MM)</th></tr>
    """
    for r in por_os:
        html += f"""
        <tr>
            <td>{r['os_codigo'] or '-'}</td>
            <td>{r['item_paint'] or '-'}</td>
            <td>{minutos_para_hhmm(r['minutos'])}</td>
        </tr>
        """
    html += "</table></div>"

    html += "</div>"
    html += "<br><a class='btn' href='/colaboradores'>Voltar</a>"

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        user=session['user'],
        perfil=session['perfil']
    )

from datetime import datetime

# -------------------------
# Editar Projeto PAINT
# -------------------------
@app.route('/projeto/edit/<int:id>', methods=['GET', 'POST'])
def editar_projeto(id):
    if 'user' not in session:
        return redirect('/')
    if session['perfil'] != 'admin':
        return 'Acesso negado'

    con = get_db()
    cur = con.cursor()

    cur.execute("SELECT * FROM projeto_paint WHERE id = %s", (id,))
    projeto = cur.fetchone()

    if not projeto:
        con.close()
        return "Projeto não encontrado", 404

    item_antigo = projeto["item_paint"]

    if request.method == 'POST':
        classificacao = request.form.get('classificacao')
        item_novo = request.form.get('item_paint')
        tipo = request.form.get('tipo_atividade')
        objeto = request.form.get('objeto')
        objetivo = request.form.get('objetivo')
        dt_ini = request.form.get('dt_ini') or None
        dt_fim = request.form.get('dt_fim') or None
        hh_atual = request.form.get('hh_atual') or 0

        try:
            # ---- atualiza projeto ----
            cur.execute("""
                UPDATE projeto_paint
                SET classificacao=%s, item_paint=%s, tipo_atividade=%s, objeto=%s,
                    objetivo_geral=%s, dt_ini=%s, dt_fim=%s, hh_atual=%s
                WHERE id=%s
            """, (classificacao, item_novo, tipo, objeto,
                  objetivo, dt_ini, dt_fim, hh_atual, id))

            # ---- CASCADE MANUAL ----
            if item_antigo != item_novo:
                cur.execute(
                    "UPDATE os SET item_paint=%s WHERE item_paint=%s",
                    (item_novo, item_antigo)
                )
                cur.execute(
                    "UPDATE horas SET item_paint=%s WHERE item_paint=%s",
                    (item_novo, item_antigo)
                )

            con.commit()

        except Exception as e:
            con.rollback()
            con.close()
            return f"Erro ao atualizar projeto: {e}"

        con.close()
        return redirect('/paint')

    con.close()

    # ----- FORMULÁRIO HTML PRÉ-PREENCHIDO -----

    html = f"""
    <h3>Editar Projeto PAINT – {projeto['item_paint']}</h3>

    <form method='post'>
      <div>Classificação:
        <select name='classificacao'>
            <option value='Prioritário'  {"selected" if projeto["classificacao"] == "Prioritário" else ""}>Prioritário</option>
            <option value='Obrigatório'  {"selected" if projeto["classificacao"] == "Obrigatório" else ""}>Obrigatório</option>
            <option value='Complementar' {"selected" if projeto["classificacao"] == "Complementar" else ""}>Complementar</option>
            <option value='Novo'         {"selected" if projeto["classificacao"] == "Novo" else ""}>Novo</option>
        </select>
      </div>

      <div>Item PAINT:
        <input name='item_paint' value='{projeto["item_paint"]}' required>
      </div>

      <div>Tipo de Atividade:
        <input name='tipo_atividade' value='{projeto["tipo_atividade"] or ""}'>
      </div>

      <div>Objeto:
        <textarea name='objeto' rows='3'>{projeto["objeto"] or ""}</textarea>
      </div>

      <div>Objetivo Geral:
        <textarea name='objetivo' rows='3'>{projeto["objetivo_geral"] or ""}</textarea>
      </div>

      <div>Data Inicial:
        <input type='date' name='dt_ini' value='{projeto["dt_ini"] or ""}'>
      </div>

      <div>Data Final:
        <input type='date' name='dt_fim' value='{projeto["dt_fim"] or ""}'>
      </div>

      <div>HH Atual:
        <input type='number' name='hh_atual' value='{projeto["hh_atual"] or 0}'>
      </div>

      <button class='btn'>Salvar alterações</button>
      <a class='btn' href='/paint'>Voltar</a>
    </form>
    """

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        user=session['user'], perfil=session['perfil']
    )

# -------------------------
# Projetos PAINT - list / add
# -------------------------
@app.route('/paint', methods=['GET', 'POST'])
def paint():
    if 'user' not in session:
        return redirect('/')
    if session['perfil'] != 'admin':
        return 'Acesso negado'

    con = get_db()
    cur = con.cursor()

    # ---------------- SALVAR NOVO PROJETO ----------------
    if request.method == 'POST':
        classificacao = request.form.get('classificacao')
        item = request.form.get('item_paint')
        tipo = request.form.get('tipo_atividade')
        objeto = request.form.get('objeto')
        objetivo = request.form.get('objetivo')
        dt_ini = request.form.get('dt_ini') or None
        dt_fim = request.form.get('dt_fim') or None

        try:
            cur.execute("""
                INSERT INTO projeto_paint
                (classificacao, item_paint, tipo_atividade, objeto, objetivo_geral, dt_ini, dt_fim)
                VALUES (%s,%s,%s,%s,%s,%s,%s)
            """, (classificacao, item, tipo, objeto, objetivo, dt_ini, dt_fim))
            con.commit()
        except Exception:
            con.rollback()

    # ---------------- BUSCAR PROJETOS ----------------
    cur.execute("""
        SELECT *
        FROM projeto_paint
        ORDER BY item_paint
    """)
    projetos = cur.fetchall()

    # ---------------- SOMATÓRIO DE HORAS (UMA QUERY) ----------------
    cur.execute("""
        SELECT
            item_paint,
            SUM(
                (CAST(SUBSTRING(duracao FROM 1 FOR 2) AS INTEGER) * 60) +
                 CAST(SUBSTRING(duracao FROM 4 FOR 2) AS INTEGER)
            ) AS minutos
        FROM horas
        GROUP BY item_paint
    """)
    horas_rows = cur.fetchall()

    # mapa item_paint -> minutos
    mapa_horas = {
        r["item_paint"]: r["minutos"] or 0
        for r in horas_rows
    }

    con.close()

    # ---------------- HTML ----------------
    html = "<h3>Cadastrar Projeto PAINT</h3>"
    html += """
    <form method='post'>
      <div>Classificação:
        <select name='classificacao'>
          <option value="Obrigatório">Obrigatório</option>
          <option value="Prioritário">Prioritário</option>
          <option value="Complementar">Complementar</option>
          <option value="Novo">Novo</option>
        </select>
      </div>
      <div>Item PAINT (ex: O-1): <input name='item_paint' required></div>
      <div>Tipo de Atividade: <input name='tipo_atividade'></div>
      <div>Objeto: <input name='objeto'></div>
      <div>Objetivo Geral: <input name='objetivo'></div>
      <div>Data Inicial: <input type='date' name='dt_ini'></div>
      <div>Data Final: <input type='date' name='dt_fim'></div>
      <div><button class='btn'>Adicionar</button></div>
    </form>
    """

    # ---------------- IMPORTAÇÃO EM LOTE ----------------
    html += """
    <h3>Importar múltiplos Projetos PAINT</h3>
    <form method='post' action='/paint/import'>
        <p>Cole os dados abaixo (uma linha por projeto, separando colunas por TAB ou ;):</p>
        <textarea name='bulk_data' rows='15' style='width:100%'></textarea>
        <div><button class='btn'>Importar Projetos</button></div>
    </form>
    """

    # ---------------- LISTAGEM ----------------
    html += "<h4>Projetos cadastrados</h4>"
    html += """
    <input type="text" id="searchPaint" placeholder="Pesquisar projetos..."
           style="width:100%; padding:6px; margin:8px 0;">

    <script>
    document.getElementById("searchPaint").addEventListener("keyup", function() {
        let filter = this.value.toLowerCase();
        let rows = document.querySelectorAll("#tabelaPaint tbody tr");
        rows.forEach(row => {
            row.style.display = row.innerText.toLowerCase().includes(filter) ? "" : "none";
        });
    });
    </script>
    """

    html += """
    <div style="margin-bottom:10px;">
        <a class='btn btn-danger' href='/projeto/delete_all'
           onclick="return confirm('Deseja realmente excluir TODOS os Projetos PAINT?');">
           Excluir todos
        </a>
    </div>
    """

    html += """
    <table id="tabelaPaint">
      <tr>
        <th>Item</th>
        <th>Classif.</th>
        <th>Tipo</th>
        <th>Dt. Início</th>
        <th>Dt. Fim</th>
        <th>HH Atual</th>
        <th>HH Executada</th>
        <th>% Executado</th>
        <th>Ações</th>
      </tr>
    """

    for p in projetos:
        minutos = mapa_horas.get(p["item_paint"], 0)

        hh = minutos // 60
        mm = minutos % 60
        soma = f"{hh:02d}:{mm:02d}"

        if p["hh_atual"] and p["hh_atual"] > 0:
            total_prev = p["hh_atual"] * 60
            percentual = (minutos / total_prev) * 100
            percentual_fmt = f"{percentual:.2f}%"
        else:
            percentual_fmt = "0%"

        html += f"""
        <tr>
          <td>{p['item_paint']}</td>
          <td>{p['classificacao']}</td>
          <td>{p['tipo_atividade'] or ''}</td>
          <td>{fmt(p['dt_ini'])}</td>
          <td>{fmt(p['dt_fim'])}</td>
          <td>{p['hh_atual']}</td>
          <td>{soma}</td>
          <td>{percentual_fmt}</td>
          <td>
            <a class='btn' href='/projeto/edit/{p["id"]}'>Editar</a>
            <a class='btn btn-danger'
               href='/projeto/delete/{p["id"]}'
               onclick="return confirm('Excluir este projeto?');">
               Excluir
            </a>
          </td>
        </tr>
        """

    html += "</table>"

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        user=session['user'],
        perfil=session['perfil']
    )

from datetime import datetime

from datetime import datetime
from psycopg2 import IntegrityError

@app.route('/paint/import', methods=['POST'])
def paint_import():
    if 'user' not in session:
        return redirect('/')
    if session['perfil'] != 'admin':
        return 'Acesso negado'

    data = request.form.get('bulk_data')
    if not data:
        return "Nenhum dado fornecido"

    con = get_db()
    cur = con.cursor()

    linhas = data.strip().splitlines()
    inseridos = 0
    ignorados = 0

    def conv_data(d):
        d = d.strip()
        if not d or d == "***":
            return None
        try:
            return datetime.strptime(d, "%d/%m/%Y").date()
        except:
            return None

    for i, linha in enumerate(linhas):
        linha = linha.strip()
        if not linha:
            continue

        # ignora cabeçalho
        if i == 0 and linha.lower().startswith("classificação"):
            continue

        # Google Sheets → TAB
        cols = linha.split('\t')

        if len(cols) < 6:
            ignorados += 1
            continue

        classificacao = cols[0].strip()
        item_paint   = cols[1].strip()
        tipo         = cols[2].strip()
        objeto       = cols[3].strip()
        objetivo     = cols[4].strip()
        dt_ini       = conv_data(cols[5])

        dt_fim = conv_data(cols[6]) if len(cols) > 6 else None

        # HH atual
        hh_atual = 0
        if len(cols) > 7 and cols[7].strip():
            try:
                hh_atual = int(float(cols[7].replace(",", ".")))
            except:
                hh_atual = 0

        if not item_paint:
            ignorados += 1
            continue

        try:
            cur.execute("""
                INSERT INTO projeto_paint
                (classificacao, item_paint, tipo_atividade, objeto,
                 objetivo_geral, dt_ini, dt_fim, hh_atual)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                classificacao,
                item_paint,
                tipo,
                objeto,
                objetivo,
                dt_ini,
                dt_fim,
                hh_atual
            ))
            inseridos += 1

        except IntegrityError:
            con.rollback()
            ignorados += 1
            continue

    con.commit()
    con.close()

    return f"""
    <h3>Importação concluída</h3>
    <p>✅ Inseridos: <b>{inseridos}</b></p>
    <p>⚠️ Ignorados (duplicados ou inválidos): <b>{ignorados}</b></p>
    <a href="/paint">Voltar</a>
    """

@app.route('/projeto/delete/<int:id>')
def projeto_delete(id):
    if 'user' not in session:
        return redirect('/')
    if session['perfil'] != 'admin':
        return 'Acesso negado'

    con = get_db()
    cur = con.cursor()
    cur.execute("DELETE FROM projeto_paint WHERE id=%s", (id,))
    con.commit()
    con.close()
    return redirect('/paint')


@app.route('/projeto/delete_all')
def delete_all_projetos():
    if 'user' not in session:
        return redirect('/')
    if session['perfil'] != 'admin':
        return 'Acesso negado'

    con = get_db()
    cur = con.cursor()
    cur.execute("DELETE FROM projeto_paint")
    con.commit()
    con.close()
    return redirect('/paint')


# -------------------------
# OS - list / add
# -------------------------
@app.route('/os', methods=['GET', 'POST'])
def os_page():
    if 'user' not in session:
        return redirect('/')
    if session['perfil'] != 'admin':
        return 'Acesso negado'
    con = get_db()
    cur = con.cursor()
    # Carrega colaboradores para uso nos selects
    cur.execute('SELECT nome FROM colaboradores ORDER BY nome')
    colabs = [r['nome'] for r in cur.fetchall()]

    if request.method == 'POST':
        codigo = request.form.get('codigo')
        item = request.form.get('item_paint')
        resumo = request.form.get('resumo')
        unidade = ", ".join(request.form.getlist("unidade"))
        supervisao = ", ".join(request.form.getlist('supervisao'))
        coordenacao = ", ".join(request.form.getlist('coordenacao'))
        equipe = ", ".join(request.form.getlist('equipe'))
        observacao = request.form.get('observacao')
        status = request.form.get('status')
        plan = 1 if request.form.get('plan') == 'on' else 0
        exec_ = 1 if request.form.get('exec') == 'on' else 0
        rp = 1 if request.form.get('rp') == 'on' else 0
        rf = 1 if request.form.get('rf') == 'on' else 0
        dt_conc = request.form.get('dt_conclusao') or None
        try:
            cur.execute(
                'INSERT INTO os (codigo,item_paint,resumo, unidade,supervisao,coordenacao,equipe,observacao,status,plan,exec,rp,rf,dt_conclusao) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)',
                (codigo, item, resumo, unidade, supervisao, coordenacao, equipe, observacao, status, plan, exec_, rp, rf,
                 dt_conc))
            con.commit()
        except IntegrityError:
            con.rollback()

    cur.execute('SELECT * FROM os ORDER BY codigo')
    rows = cur.fetchall()

    cur.execute('SELECT item_paint FROM projeto_paint ORDER BY item_paint')
    items = [r['item_paint'] for r in cur.fetchall()]

    con.close()

    html = '<h3>Cadastrar O.S</h3>'
    html += "<form method='post'>"
    html += "<div>Código (ex: OS-001): <input name='codigo' required></div>"
    html += "<div>Item PAINT: <select name='item_paint'>"
    for it in items:
        html += f"<option value='{it}'>{it}</option>"
    html += "</select></div>"
    # >>> CAMPO NOVO (antes de Unidade)
    html += "<div>Resumo: <input name='resumo' style='width:300px'></div>"
    html += """
    <div>
        Unidade:
        <select name='unidade' multiple size='5'>
            <option value='DAC'>DAC</option>
            <option value='DIRAE'>DIRAE</option>
            <option value='DMAD'>DMAD</option>
            <option value='DOSE'>DOSE</option>
            <option value='DACGR'>DACGR</option>
        </select>
    </div>
    """
    # Supervisão
    html += "<div>Supervisão:<br><select name='supervisao' multiple size='6' style='width:260px'>"
    for c in colabs:
        html += f"<option value='{c}'>{c}</option>"
    html += "</select></div>"

    # Coordenação
    html += "<div>Coordenação:<br><select name='coordenacao' multiple size='6' style='width:260px'>"
    for c in colabs:
        html += f"<option value='{c}'>{c}</option>"
    html += "</select></div>"

    # Equipe
    html += "<div>Equipe:<br><select name='equipe' multiple size='7' style='width:260px'>"
    for c in colabs:
        html += f"<option value='{c}'>{c}</option>"
    html += "</select></div>"

    html += "<div>Observação: <input name='observacao'></div>"
    html += "<div>Status: <select name='status'><option>Andamento</option><option>Concluido</option></select></div>"
    html += "<div>Flags: <label><input type='checkbox' name='plan'> PLAN</label> <label><input type='checkbox' name='exec'> EXEC</label> <label><input type='checkbox' name='rp'> RP</label> <label><input type='checkbox' name='rf'> RF</label></div>"
    html += "<div>Data conclusão: <input type='date' name='dt_conclusao'></div>"
    html += "<div><button class='btn'>Adicionar OS</button></div>"
    html += "</form>"
    # Barra de pesquisa
    html += """
    <div style="margin: 15px 0;">
        <input type="text" id="searchInput" class="form-control" placeholder="Pesquisar..." 
               style="padding: 8px; width: 100%; font-size: 16px;">
    </div>

    <script>
    document.addEventListener("DOMContentLoaded", function() {
        const input = document.getElementById("searchInput");

        input.addEventListener("keyup", function() {
            let filter = input.value.toLowerCase();
            let rows = document.querySelectorAll("table tbody tr");

            rows.forEach(row => {
                let text = row.innerText.toLowerCase();
                row.style.display = text.includes(filter) %s "" : "none";
            });
        });
    });
    </script>
    """
    html += '<h4>O.S cadastradas</h4>'
    html += """
    <div style="margin-bottom:10px;">
        <a class='btn btn-primary' href='/os/import'>Importar por texto</a>
        <a class='btn btn-danger' href='/os/delete_all'
           onclick="return confirm('Deseja realmente excluir TODAS as OS cadastradas?');">
           Excluir todas
        </a>
    </div>
    """

    html += '<table><tr><th>Código</th><th>Item PAINT</th><th>Resumo</th><th>Status</th><th>Ações</th></tr>'
    for r in rows:
        html += f"""
        <tr>
            <td>{r['codigo']}</td>
            <td>{r['item_paint']}</td>
            <td>{r['resumo']}</td>
            <td>{r['status']}</td>
            <td>
                <a class='btn' href='/os/view/{r["id"]}'>Ver</a>
                <a class='btn' href='/os/edit/{r["id"]}'>Editar</a>
                <a class='btn btn-danger' href='/os/delete/{r["id"]}' onclick="return confirm('Deseja realmente excluir esta O.S?');">Excluir</a>

            </td>
        </tr>
        """
    html += '</table>'

    return render_template_string(BASE.replace('{% block content %}{% endblock %}', html), user=session['user'],
                                  perfil=session['perfil'])


@app.route('/os/delete/<int:id>')
def os_delete(id):
    if 'user' not in session:
        return redirect('/')
    if session['perfil'] != 'admin':
        return 'Acesso negado'

    con = get_db()
    cur = con.cursor()
    cur.execute("DELETE FROM os WHERE id=%s", (id,))
    con.commit()
    con.close()
    return redirect('/os')


@app.route('/os/delete_all')
def os_delete_all():
    if 'user' not in session:
        return redirect('/')
    if session['perfil'] != 'admin':
        return 'Acesso negado'

    con = get_db()
    cur = con.cursor()
    cur.execute("DELETE FROM os")  # apaga todos os registros
    con.commit()
    con.close()
    return redirect('/os')


@app.route('/os/view/<int:id>')
def os_view(id):
    if 'user' not in session:
        return redirect('/')
    if session['perfil'] != 'admin':
        return 'Acesso negado'

    con = get_db()
    cur = con.cursor()
    cur.execute("SELECT * FROM os WHERE id=%s", (id,))
    r = cur.fetchone()
    con.close()

    if not r:
        return "O.S não encontrada"

    def icon(v):
        if v == 1:
            return "<span style='color:green; font-weight:bold;'>✔</span>"
        else:
            return "<span style='color:red; font-weight:bold;'>✖</span>"

    html = f"""
    <h3>Visualizar O.S {r['codigo']}</h3>
    <p><strong>Código:</strong> {r['codigo']}</p>
    <p><strong>Item PAINT:</strong> {r['item_paint']}</p>
    <p><strong>Resumo:</strong> {r['resumo']}</p>
    <p><strong>Unidade:</strong> {r['unidade'] if r['unidade'] else ''}</p>
    <p><strong>Supervisão:</strong> {r['supervisao']}</p>
    <p><strong>Coordenação:</strong> {r['coordenacao']}</p>
    <p><strong>Equipe:</strong> {r['equipe']}</p>
    <p><strong>Observação:</strong> {r['observacao']}</p>
    <p><strong>Status:</strong> {r['status']}</p>
    <p><strong>Flags:</strong>
    PLAN={icon(r['plan'])} |
    EXEC={icon(r['exec'])} |
    RP={icon(r['rp'])} |
    RF={icon(r['rf'])}
</p>
    <p><strong>Data Conclusão:</strong> {r['dt_conclusao'] or ''}</p>

    <a class='btn' href='/os/edit/{r["id"]}'>Editar</a>
    <a class='btn' href='/os'>Voltar</a>
    """

    return render_template_string(BASE.replace('{% block content %}{% endblock %}', html),
                                  user=session['user'], perfil=session['perfil'])


@app.route('/os/edit/<int:id>', methods=['GET', 'POST'])
def os_edit(id):
    if 'user' not in session:
        return redirect('/')
    if session['perfil'] != 'admin':
        return 'Acesso negado'

    con = get_db()
    cur = con.cursor()

    cur.execute("SELECT * FROM os WHERE id=%s", (id,))
    os = cur.fetchone()

    if not os:
        con.close()
        return "O.S não encontrada"

    codigo_antigo = os["codigo"]

    # carregar dados auxiliares (igual ao seu código)
    cur.execute('SELECT nome FROM colaboradores ORDER BY nome')
    colabs = [r['nome'] for r in cur.fetchall()]

    cur.execute('SELECT item_paint FROM projeto_paint ORDER BY item_paint')
    items = [r['item_paint'] for r in cur.fetchall()]

    if request.method == 'POST':
        codigo_novo = request.form.get('codigo')
        item = request.form.get('item_paint')
        resumo = request.form.get('resumo') or ""
        unidade = ", ".join(request.form.getlist("unidade"))
        supervisao = ", ".join(request.form.getlist('supervisao'))
        coordenacao = ", ".join(request.form.getlist('coordenacao'))
        equipe = ", ".join(request.form.getlist('equipe'))
        observacao = request.form.get('observacao')
        status = request.form.get('status')
        plan = 1 if request.form.get('plan') else 0
        exec_ = 1 if request.form.get('exec') else 0
        rp = 1 if request.form.get('rp') else 0
        rf = 1 if request.form.get('rf') else 0
        dt_conc = request.form.get('dt_conclusao') or None

        try:
            # ---- atualiza OS ----
            cur.execute("""
                UPDATE os SET 
                    codigo=%s, item_paint=%s, resumo=%s, unidade=%s, supervisao=%s, 
                    coordenacao=%s, equipe=%s, observacao=%s, status=%s, 
                    plan=%s, exec=%s, rp=%s, rf=%s, dt_conclusao=%s
                WHERE id=%s
            """, (
                codigo_novo, item, resumo, unidade, supervisao,
                coordenacao, equipe, observacao, status,
                plan, exec_, rp, rf, dt_conc, id
            ))

            # ---- CASCADE MANUAL ----
            if codigo_antigo != codigo_novo:
                cur.execute(
                    "UPDATE horas SET os_codigo=%s WHERE os_codigo=%s",
                    (codigo_novo, codigo_antigo)
                )
                cur.execute(
                    "UPDATE delegacoes SET os_codigo=%s WHERE os_codigo=%s",
                    (codigo_novo, codigo_antigo)
                )

            con.commit()

        except Exception as e:
            con.rollback()
            con.close()
            return f"Erro ao atualizar O.S: {e}"

        con.close()
        return redirect('/os')

    con.close()

    # ---------------- CARREGAR DADOS ATUAIS ----------------
    unidade_atual = (os['unidade'] or "").split(", ")
    supervisao_atual = (os['supervisao'] or "").split(", ")
    coordenacao_atual = (os['coordenacao'] or "").split(", ")
    equipe_atual = (os['equipe'] or "").split(", ")

    resumo_atual = os['resumo'] if 'resumo' in os.keys() else ""

    # ---------------- FORM HTML ----------------
    html = f"""
    <h3>Editar O.S {os['codigo']}</h3>
    <form method='post'>
      <div>Código: <input name='codigo' value='{os['codigo']}'></div>

      <div>Item PAINT:
        <select name='item_paint'>
    """

    for it in items:
        sel = "selected" if it == os['item_paint'] else ""
        html += f"<option value='{it}' {sel}>{it}</option>"

    html += "</select></div>"

    # >>> CAMPO RESUMO <<<
    html += f"""
    <div>Resumo: 
        <input name='resumo' value="{resumo_atual}" style='width:300px'>
    </div>
    """

    # unidades
    unidades_opcoes = ["DAC", "DIRAE", "DMAD", "DOSE", "DACGR"]
    html += "<div>Unidade:<br><select name='unidade' multiple size='5' style='width:260px'>"
    for u in unidades_opcoes:
        sel = "selected" if u in unidade_atual else ""
        html += f"<option value='{u}' {sel}>{u}</option>"
    html += "</select></div>"

    # supervisão
    html += "<div>Supervisão:<br><select name='supervisao' multiple size='6' style='width:260px'>"
    for c in colabs:
        sel = "selected" if c in supervisao_atual else ""
        html += f"<option value='{c}' {sel}>{c}</option>"
    html += "</select></div>"

    # coordenação
    html += "<div>Coordenação:<br><select name='coordenacao' multiple size='6' style='width:260px'>"
    for c in colabs:
        sel = "selected" if c in coordenacao_atual else ""
        html += f"<option value='{c}' {sel}>{c}</option>"
    html += "</select></div>"

    # equipe
    html += "<div>Equipe:<br><select name='equipe' multiple size='7' style='width:260px'>"
    for c in colabs:
        sel = "selected" if c in equipe_atual else ""
        html += f"<option value='{c}' {sel}>{c}</option>"
    html += "</select></div>"

    # observação
    html += f"<div>Observação: <input name='observacao' value='{os['observacao'] or ''}'></div>"

    # status
    html += "<div>Status: <select name='status'>"
    html += f"<option value='Andamento' {'selected' if os['status']=='Andamento' else ''}>Andamento</option>"
    html += f"<option value='Concluido' {'selected' if os['status']=='Concluido' else ''}>Concluído</option>"
    html += "</select></div>"

    # flags
    html += f"""
    <div>Flags:
      <label><input type='checkbox' name='plan' {'checked' if os['plan'] else ''}> PLAN</label>
      <label><input type='checkbox' name='exec' {'checked' if os['exec'] else ''}> EXEC</label>
      <label><input type='checkbox' name='rp' {'checked' if os['rp'] else ''}> RP</label>
      <label><input type='checkbox' name='rf' {'checked' if os['rf'] else ''}> RF</label>
    </div>
    """

    # data conclusão
    html += f"""
    <div>Data conclusão: 
        <input type='date' name='dt_conclusao' value='{os['dt_conclusao'] or ''}'>
    </div>
    """

    html += "<div><button class='btn btn-primary'>Salvar</button></div>"
    html += "</form>"

    html += "<a class='btn' href='/os'>Voltar</a>"

    return render_template_string(BASE.replace('{% block content %}{% endblock %}', html),
                                  user=session['user'], perfil=session['perfil'])


# -------------------------
# Importar OS por colar texto
# -------------------------
@app.route('/os/import', methods=['GET', 'POST'])
def os_import():
    if 'user' not in session:
        return redirect('/')
    if session['perfil'] != 'admin':
        return "Acesso negado"

    msg = ""

    def conv_bool(v):
        return 1 if v.strip().upper() == "TRUE" else 0

    def conv_data(v):
        v = v.strip()
        if not v:
            return None
        try:
            return datetime.strptime(v, "%d/%m/%Y").date()
        except:
            return None

    if request.method == 'POST':
        texto = request.form.get("texto", "").strip()

        if texto:
            linhas = texto.splitlines()
            con = get_db()
            cur = con.cursor()

            inseridos = 0
            ignorados = 0

            for i, linha in enumerate(linhas):
                linha = linha.strip()
                if not linha:
                    continue

                # ignora cabeçalho
                if i == 0 and linha.lower().startswith("os"):
                    continue

                partes = linha.split("\t")

                if len(partes) < 7:
                    ignorados += 1
                    continue

                codigo      = partes[0].strip()
                item_paint  = partes[1].strip()
                resumo      = partes[2].strip()
                unidade     = partes[3].strip()
                coordenacao = partes[4].strip()
                equipe      = partes[5].strip()
                observacao  = partes[6].strip() if len(partes) > 6 else None

                plan  = conv_bool(partes[7]) if len(partes) > 7 else 0
                exec_ = conv_bool(partes[8]) if len(partes) > 8 else 0
                rp    = conv_bool(partes[9]) if len(partes) > 9 else 0
                rf    = conv_bool(partes[10]) if len(partes) > 10 else 0

                status = partes[11].strip() if len(partes) > 11 else None
                dt_conclusao = conv_data(partes[12]) if len(partes) > 12 else None

                if not codigo or not item_paint:
                    ignorados += 1
                    continue

                try:
                    cur.execute("""
                        INSERT INTO os (
                            codigo,
                            item_paint,
                            resumo,
                            unidade,
                            supervisao,
                            coordenacao,
                            equipe,
                            observacao,
                            plan,
                            exec,
                            rp,
                            rf,
                            status,
                            dt_conclusao
                        )
                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    """, (
                        codigo,
                        item_paint,
                        resumo,        # ID da planilha
                        unidade,
                        None,          # supervisão não existe → NULL
                        coordenacao,
                        equipe,
                        observacao,
                        plan,
                        exec_,
                        rp,
                        rf,
                        status,
                        dt_conclusao
                    ))
                    inseridos += 1

                except IntegrityError:
                    con.rollback()
                    ignorados += 1

            con.commit()
            con.close()

            msg = f"✅ {inseridos} O.S inseridas | ⚠️ {ignorados} ignoradas"

    return render_template_string(
        BASE.replace(
            "{% block content %}{% endblock %}",
            f"""
            <h3>Importar O.S (Copiar e colar do Google Sheets)</h3>

            <form method="post">
                <textarea name="texto" rows="18" style="width:100%;"></textarea><br><br>
                <button class="btn btn-primary">Importar</button>
            </form>

            <p><b>{msg}</b></p>

            <p>Formato esperado (TAB):</p>
            <pre style="background:#eef; padding:10px;">
OS | ITEM PAINT | ID | UNIDADE | COORDENAÇÃO | EQUIPE | OBS | PLAN | EXEC | RP | RF | STATUS | DT_CONCLUSÃO
            </pre>

            <a class="btn" href="/os">Voltar</a>
            """
        ),
        user=session['user'],
        perfil=session['perfil']
    )

# -------------------------
# Lançar horas (colaborador)
# -------------------------
@app.route('/lancar', methods=['GET', 'POST'])
def lancar():
    if 'user' not in session:
        return redirect('/')

    from datetime import datetime, date

    con = get_db()
    cur = con.cursor()

    # -------------------------
    # CARREGAR OS
    # -------------------------
    cur.execute("SELECT codigo, item_paint, resumo FROM os ORDER BY codigo")
    oss = cur.fetchall()

    # -------------------------
    # CARREGAR DELEGAÇÕES
    # -------------------------
    cur.execute("""
        SELECT id, requisicoes, os_codigo, grau
        FROM delegacoes
        WHERE colaborador_id = %s
          AND status = 'Em Andamento'
    """, (session["user_id"],))
    delegacoes = [dict(r) for r in cur.fetchall()]

    # -------------------------
    # PROCESSAR POST
    # -------------------------
    
    cur.execute("SELECT id, nome FROM colaboradores ORDER BY nome")
    colaboradores = cur.fetchall()
    
    if request.method == 'POST':

        item = request.form.get('item')
        os_codigo = request.form.get('os')
        atividade = request.form.get('atividade')
        delegacao_id = request.form.get('delegacao_id') or None
        observacoes = request.form.get('observacoes')

        datas = request.form.getlist("data[]")
        horas_ini = request.form.getlist("hora_ini[]")
        horas_fim = request.form.getlist("hora_fim[]")

        if not datas:
            con.close()
            return "Nenhum lançamento informado"

        total_minutos_paint = 0

        for data, hora_ini, hora_fim in zip(datas, horas_ini, horas_fim):

            # ---- validar data (somente 2026)
            try:
                dt = datetime.strptime(data, "%Y-%m-%d")
                if dt.year != 2026:
                    con.close()
                    return "Só é permitido lançar horas em 2026"
            except:
                con.close()
                return "Data inválida"

            # ---- calcular duração
            try:
                ini = datetime.strptime(hora_ini, "%H:%M")
                fim = datetime.strptime(hora_fim, "%H:%M")
                minutos = (fim - ini).seconds // 60
                duracao = f"{minutos//60:02d}:{minutos%60:02d}"
            except:
                con.close()
                return "Hora inválida"

            # ---- inserir
            cur.execute("""
                INSERT INTO horas
                (colaborador_id, data, item_paint, os_codigo, atividade,
                 delegacao_id, hora_inicio, hora_fim, duracao, duracao_minutos, observacoes)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                RETURNING id
            """, (
                session['user_id'], data, item, os_codigo, atividade,
                delegacao_id, hora_ini, hora_fim, duracao, minutos, observacoes
            ))
            hora_id = cur.fetchone()["id"]

            # -------------------------
            # OS 1.2 – Atendimento
            # -------------------------
            if os_codigo == "1.15/2026":

                responsaveis_ids = request.form.getlist("responsaveis[]")
            
                # lançar horas para participantes (continua igual)
                for resp_id in responsaveis_ids:
                    if not resp_id or not resp_id.isdigit():
                        continue
                    resp_id = int(resp_id)
                    if resp_id == session["user_id"]:
                        continue
            
                    cur.execute("""
                        INSERT INTO horas
                        (colaborador_id, data, item_paint, os_codigo, atividade,
                         delegacao_id, hora_inicio, hora_fim, duracao, duracao_minutos, observacoes)
                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    """, (
                        resp_id, data, item, os_codigo, atividade,
                        None, hora_ini, hora_fim, duracao, minutos,
                        f"Lançamento automático – Atendimento OS {os_codigo}"
                    ))
            
                # 🔐 só insere atendimento se tiver dados
                tem_atendimento = any([
                    request.form.get("data_consultoria"),
                    request.form.get("assunto"),
                    request.form.get("macro"),
                    request.form.get("diretoria"),
                    request.form.get("atividade_atendimento"),
                    request.form.get("meio_contato"),
                    request.form.getlist("entidades[]"),
                    request.form.get("observacao_atendimento")
                ])
            
                if tem_atendimento:
            
                    nomes = []
                    for rid in responsaveis_ids:
                        if rid and rid.isdigit():
                            cur.execute("SELECT nome FROM colaboradores WHERE id=%s", (rid,))
                            r = cur.fetchone()
                            if r:
                                nomes.append(r["nome"])
            
                    os_resumo = next(
                        (o["resumo"] for o in oss if o["codigo"] == os_codigo),
                        None
                    )
            
                    cur.execute("""
                        INSERT INTO atendimentos (
                            hora_id, colaborador_id, os_codigo, os_resumo,
                            responsaveis_consultoria, macro, diretoria, atividade,
                            data_consultoria, assunto, participantes_externos,
                            entidades, meio_contato, observacao,
                            duracao_minutos, data_lancamento
                        ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    """, (
                        hora_id, session["user_id"], os_codigo,
                        os_resumo,
                        ", ".join(nomes),
                        request.form.get("macro"),
                        request.form.get("diretoria"),
                        request.form.get("atividade_atendimento"),
                        request.form.get("data_consultoria"),
                        request.form.get("assunto"),
                        request.form.get("participantes_externos"),
                        ", ".join(request.form.getlist("entidades[]")),
                        request.form.get("meio_contato"),
                        request.form.get("observacao_atendimento"),
                        minutos, data
                    ))

            # -------------------------
            # OS 1.20 – Consultoria
            # -------------------------
            elif os_codigo in ("1.14/2026", "1.16/2026"):
                tipo_consultoria = ("consultoria" if os_codigo == "1.14/2026" else "treinamento")
                responsaveis = request.form.getlist("responsaveis2[]")
            
                for resp_id in responsaveis:
                    if not resp_id or not resp_id.isdigit():
                        continue
                    resp_id = int(resp_id)
                    if resp_id == session["user_id"]:
                        continue
            
                    cur.execute("""
                        INSERT INTO horas
                        (colaborador_id, data, item_paint, os_codigo, atividade,
                         delegacao_id, hora_inicio, hora_fim, duracao, duracao_minutos, observacoes)
                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    """, (
                        resp_id, data, item, os_codigo, atividade,
                        None, hora_ini, hora_fim, duracao, minutos,
                        f"Lançamento automático – Consultoria OS {os_codigo}"
                    ))
            
                tem_consultoria = any([
                    request.form.get("data_consul"),
                    request.form.get("assunto_consultoria"),
                    request.form.get("meio"),
                    request.form.get("num_oficio"),
                    request.form.get("palavras_chave"),
                    request.form.getlist("secretarias[]"),
                    request.form.get("observacao")
                ])
            
                if tem_consultoria:
            
                    nomes = []
                    for rid in responsaveis:
                        if rid and rid.isdigit():
                            cur.execute("SELECT nome FROM colaboradores WHERE id=%s", (rid,))
                            r = cur.fetchone()
                            if r:
                                nomes.append(r["nome"])
            
                    os_resumo = next(
                        (o["resumo"] for o in oss if o["codigo"] == os_codigo),
                        None
                    )
            
                    cur.execute("""
                        INSERT INTO consultorias (
                            hora_id, colaborador_id, os_codigo, os_resumo,
                            responsaveis, tipo, data_consul, assunto,
                            secretarias, meio, palavras_chave,
                            num_oficio, observacao,
                            duracao_minutos, data_lancamento
                        ) VALUES (%s,%s,%s, %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    """, (
                        hora_id, session["user_id"], os_codigo,
                        os_resumo,
                        ", ".join(nomes),
                        tipo_consultoria,
                        request.form.get("data_consul"),
                        request.form.get("assunto_consultoria"),
                        ", ".join(request.form.getlist("secretarias[]")),
                        request.form.get("meio"),
                        request.form.get("palavras_chave"),
                        request.form.get("num_oficio"),
                        request.form.get("observacao"),
                        minutos, data
                    ))

            total_minutos_paint += minutos
            
        con.commit()
        con.close()
        return redirect('/menu')

    # -------------------------
    # HTML
    # -------------------------
    data_padrao = date.today().isoformat()
    con.close()

    form_html = """
<h3>Lançar Horas</h3>

<form method="post">

    <div>O.S:
        <select name="os" id="os_select" required>
            <option value=""></option>
            {% for o in oss %}
                <option value="{{ o.codigo }}" data-item="{{ o.item_paint }}">
                    {{ o.codigo }}{% if o.resumo %} - {{ o.resumo }}{% endif %}
                </option>
            {% endfor %}
        </select>
    </div>

    <div>Item PAINT:
        <input type="text" id="item_paint" name="item" readonly>
    </div>

    <!-- DELEGAÇÃO -->
    <div id="box_delegacao" style="display:none;">
        <div>Delegação:
            <select name="delegacao_id" id="delegacao_select">
                <option value=""></option>
            </select>
        </div>
    </div>

    <div>Atividade:
        <select name="atividade" required>
            <option>1. Planejamento</option>
            <option>2. Execução</option>
            <option>3. Relatório</option>
        </select>
    </div>

    <h4>Registros de Horas</h4>

    <div id="registros">

        <div class="registro">
            <input type="date" name="data[]" value="{{ data_padrao }}"
                   min="2026-01-01" max="2026-12-31" required>

            <input type="time" name="hora_ini[]" required>
            <input type="time" name="hora_fim[]" required>

            <button type="button" onclick="remover(this)">❌</button>
        </div>

    </div>

    <!-- ATENDIMENTO OS 1.15 -->
<div id="box_atendimento" style="display:none; border:1px solid #ccc; padding:10px; margin-top:10px;">

    <h4>Dados do Atendimento (O.S 1.15)</h4>

    <div>
        <label>Responsáveis pelo Atendimento</label><br>
        <select name="responsaveis[]" multiple size="5">
            {% for c in colaboradores %}
                <option value="{{ c.id }}">{{ c.nome }}</option>
            {% endfor %}
        </select>
        
    </div>

    <div>Macro: <input type="text" name="macro"></div>
    <div>Diretoria: <input type="text" name="diretoria"></div>

    <div>Atividade:
        <select name="atividade_atendimento">
            <option value=""></option>
            <option>Consulta</option>
            <option>Esclarecimento</option>
            <option>Orientação</option>
            <option>Preventiva</option>
        </select>
    </div>

    <div>Data do Atendimento:
        <input type="date" name="data_consultoria">
    </div>

    <div>Assunto:
        <input type="text" name="assunto">
    </div>

    <div>Participantes Externos:
        <textarea name="participantes_externos"></textarea>
    </div>

    <div>Entidades:
        <select name="entidades[]" multiple size="6">
            <option>CM</option><option>SEGOV</option><option>SMGAS</option>
            <option>PGM</option><option>SMA</option><option>SMF</option>
            <option>SME</option><option>SMCT</option><option>SMS</option>
            <option>SEDES</option><option>SMAGRO</option><option>SEINFRA</option>
            <option>SETTRAN</option><option>DMAE</option><option>FUTEL</option>
            <option>EMAM</option><option>FERUB</option><option>IPREMU</option>
            <option>SESURB</option><option>SMH</option><option>SEJUV</option>
            <option>SECOM</option><option>SEDEI</option><option>SMGE</option>
            <option>SSEG</option><option>ARESAN</option>
        </select>
    </div>

    <div>Meio de Contato:
        <select name="meio_contato">
            <option value=""></option>
            <option>Presencial</option>
            <option>Email</option>
            <option>Telefone</option>
        </select>
    </div>

    <div>Observação do Atendimento:
        <textarea name="observacao_atendimento"></textarea>
    </div>

</div>

    <!-- CONSULTORIA OS 1.14/1.16 -->
<div id="box_consultoria" style="display:none; border:1px solid #ccc; padding:10px; margin-top:10px;">

    <h4>Dados da Consultoria/Treinamento (O.S 1.16 / O.S 1.14)</h4>

    <div>
        <label>Responsáveis pela Consultoria</label><br>
        <select name="responsaveis2[]" multiple size="5">
            {% for c in colaboradores %}
                <option value="{{ c.id }}">{{ c.nome }}</option>
            {% endfor %}
        </select>
        
    </div>

    <div>Assunto:
        <textarea name="assunto_consultoria"></textarea>
    </div>
    
    <div>Meio: <input type="text" name="meio"></div>

    <div>Oficio Resposta:
        <input type="text" name="num_oficio">
    </div>

    <div>Data da Consultoria:
        <input type="date" name="data_consul">
    </div>

    <div>Palavras Chave:
        <textarea name="palavras_chave"></textarea>
    </div>

    <div>Secretarias:
        <select name="secretarias[]" multiple size="6">
            <option>CM</option><option>SEGOV</option><option>SMGAS</option>
            <option>PGM</option><option>SMA</option><option>SMF</option>
            <option>SME</option><option>SMCT</option><option>SMS</option>
            <option>SEDES</option><option>SMAGRO</option><option>SEINFRA</option>
            <option>SETTRAN</option><option>DMAE</option><option>FUTEL</option>
            <option>EMAM</option><option>FERUB</option><option>IPREMU</option>
            <option>SESURB</option><option>SMH</option><option>SEJUV</option>
            <option>SECOM</option><option>SEDEI</option><option>SMGE</option>
            <option>SSEG</option><option>ARESAN</option>
        </select>
    </div>

    <div>Observação da Consultoria/Treinamento:
        <textarea name="observacao"></textarea>
    </div>

</div>

    <button type="button" onclick="adicionar()">➕ Adicionar registro</button>
    
    <div style="margin-top:10px;">
        <label>Observação Lançamento da Hora (*Campo Não Obrigatório*) Se o lançamento de horas for múltiplo, é enviado para todos os registros, se necessário, alterar individualmente no editor de horas):</label><br>
        <textarea name="observacoes" rows="4" style="width:100%;"></textarea>
    </div>

    <button class="btn" style="margin-top:15px;">
        Registrar Lançamento(s)
    </button>
</form>

<script>
document.addEventListener("DOMContentLoaded", function () {

    const osSelect = document.getElementById("os_select");
    const itemInput = document.getElementById("item_paint");

    const delSelect = document.getElementById("delegacao_select");
    const boxDelegacao = document.getElementById("box_delegacao");

    const boxAtendimento = document.getElementById("box_atendimento");
    const boxConsultoria = document.getElementById("box_consultoria");

    const delegacoes = {{ delegacoes | tojson }};

    osSelect.addEventListener("change", function () {

        const selected = this.selectedOptions[0];
        const codigoOS = this.value;

        // item paint
        itemInput.value = selected ? selected.dataset.item : "";

        // -------- delegações
        delSelect.innerHTML = "<option value=''></option>";
        boxDelegacao.style.display = "none";
        delSelect.required = false;

        let encontrou = false;

        delegacoes.forEach(d => {
            if (d.os_codigo === codigoOS) {
                encontrou = true;
                const opt = document.createElement("option");
                opt.value = d.id;
                opt.textContent = "Reqs: " + d.requisicoes + " | Grau: " + d.grau;
                delSelect.appendChild(opt);
            }
        });

        if (encontrou) {
            boxDelegacao.style.display = "block";
            delSelect.required = true;
        }

        // -------- atendimento OS 1.15
        if (codigoOS === "1.15/2026") {
            boxAtendimento.style.display = "block";
        } else {
            boxAtendimento.style.display = "none";
        }
                // -------- atendimento OS 1.14 e 1.16
        if (codigoOS === "1.14/2026" || codigoOS === "1.16/2026") {
            boxConsultoria.style.display = "block";
        } else {
            boxConsultoria.style.display = "none";
        }
    });

});

// múltiplos registros
function adicionar() {
    const base = document.querySelector(".registro");
    const clone = base.cloneNode(true);
    clone.querySelectorAll("input").forEach(i => i.value = "");
    document.getElementById("registros").appendChild(clone);
}

function remover(btn) {
    const registros = document.querySelectorAll(".registro");
    if (registros.length > 1) {
        btn.parentElement.remove();
    }
}
</script>

"""

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", form_html),
        oss=oss,
        delegacoes=delegacoes,
        colaboradores=colaboradores,
        data_padrao=data_padrao,
        user=session['user'],
        perfil=session['perfil']
    )


# -------------------------
# Relatórios
# -------------------------
@app.route('/relatorios')
def relatorios():
    if 'user' not in session:
        return redirect('/')

    con = get_db()
    cur = con.cursor()

    perfil = session['perfil']
    user_id = session['user_id']

    # ------------------------------------------------------------------
    # 1) Total de horas por colaborador → SOMENTE ADMIN VÊ
    # ------------------------------------------------------------------
    por_colab = []
    if perfil == "admin":
        cur.execute("""
            SELECT c.nome,
                   SUM(
                        (CAST(substr(h.duracao,1,2) AS INTEGER) * 60) +
                         CAST(substr(h.duracao,4,2) AS INTEGER)
                        ) AS minutos
            FROM horas h
            JOIN colaboradores c ON h.colaborador_id = c.id
            GROUP BY c.nome
        """)
        por_colab = cur.fetchall()

    # ------------------------------------------------------------------
    # 2) Total por item_paint → TODOS VEEM
    # ------------------------------------------------------------------
    cur.execute("""
        SELECT item_paint,
               SUM(
               (CAST(substr(duracao,1,2) AS INTEGER) * 60) +
                CAST(substr(duracao,4,2) AS INTEGER)
           ) AS minutos
        FROM horas
        GROUP BY item_paint
    """)
    por_paint = cur.fetchall()

    # ------------------------------------------------------------------
    # 3) Minhas marcações com paginação
    # ------------------------------------------------------------------

    # ---- Tratamento de limite ----
    limit_param = request.args.get("limit", "100")

    if limit_param == "all":
        limite = None
    else:
        try:
            limite = int(limit_param)
        except:
            limite = 100

    mes_filtro = request.args.get("mes", "")

    sql = """
        SELECT
            h.*,
            p.item_paint as item,
            o.resumo AS os_resumo
        FROM horas h
        LEFT JOIN projeto_paint p ON p.item_paint = h.item_paint
        LEFT JOIN os o ON o.codigo = h.os_codigo
        WHERE h.colaborador_id = %s
    """

    params = [user_id]

    # filtro por mês escolhido
    if mes_filtro:
        sql += " AND EXTRACT(MONTH FROM h.data) = %s "
        params.append(mes_filtro)

    sql += " ORDER BY h.data DESC "

    # aplica LIMIT só se limite não for None
    if limite:
        sql += " LIMIT %s "
        params.append(limite)

    cur.execute(sql, tuple(params))
    minhas = cur.fetchall()

    con.close()

    # ------------------------- HTML -------------------------
    html = "<h3>Relatórios</h3>"

    # ==================================================================
    # TOTAL POR COLABORADOR
    # ==================================================================
    if perfil == "admin":
        html += "<h4>Total de horas por colaborador</h4>"
        html += "<table><tr><th>Colaborador</th><th>Total (HH:MM)</th></tr>"
        for r in por_colab:
            minutos = r['minutos'] or 0
            hh = minutos // 60
            mm = minutos % 60
            html += f"<tr><td>{r['nome']}</td><td>{hh:02d}:{mm:02d}</td></tr>"
        html += "</table><br>"

    # ==================================================================
    # TOTAL POR ITEM PAINT
    # ==================================================================
    html += "<h4>Total de horas por Item PAINT</h4>"
    html += "<table><tr><th>Item</th><th>Total (HH:MM)</th></tr>"
    for r in por_paint:
        minutos = r['minutos'] or 0
        hh = minutos // 60
        mm = minutos % 60
        html += f"<tr><td>{r['item_paint']}</td><td>{hh:02d}:{mm:02d}</td></tr>"
    html += "</table><br>"

    # ==================================================================
    # MINHAS MARCAÇÕES
    # ==================================================================
    html += "<h4>Minhas marcações</h4>"

    # ---------------- PAGINAÇÃO ----------------
    html += f"""
        <div style='margin:10px 0;'>
            Mostrar:
            <a href='/relatorios?limit=50'>50</a> |
            <a href='/relatorios?limit=100'>100</a> |
            <a href='/relatorios?limit=200'>200</a> |
            <a href='/relatorios?limit=500'>500</a> |
            <a href='/relatorios?limit=1000'>1000</a> |
            <a href='/relatorios?limit=all'>Todos</a>
        </div>
    """

    # ---------------- BOTÕES POR MÊS ----------------
    html += """
        <div style='margin-bottom:10px;'>
            <strong>Filtrar por mês:</strong><br>
    """

    meses = {
        "01": "Janeiro", "02": "Fevereiro", "03": "Março", "04": "Abril",
        "05": "Maio", "06": "Junho", "07": "Julho", "08": "Agosto",
        "09": "Setembro", "10": "Outubro", "11": "Novembro", "12": "Dezembro"
    }

    for num, nome in meses.items():
        html += f"<a class='btn' style='margin:3px;' href='/relatorios?mes={num}&limit={limit_param}'>{nome}</a>"

    html += """
        <a class='btn' style='margin:3px; background:#444;' href='/relatorios'>Limpar</a>
        </div>
    """

    # ---------------- FILTRO GERAL ----------------
    html += """
        <input type='text' id='filtroGeral' placeholder='Pesquisar em qualquer campo...'
               style='width:100%; padding:6px; margin-bottom:10px;'>
    """

    # ---------------- TABELA ----------------
    html += """
    <table id='tabelaMarcacoes'>
        <tr>
            <th>Data</th>
            <th>Item</th>
            <th>OS</th>
            <th>Atividade</th>
            <th>Observação</th>
            <th>Duração</th>
            <th>Ações</th>
        </tr>
    """
    
    for r in minhas:
        os_visual = r['os_codigo'] or ''
        if r['os_resumo']:
            os_visual += f" - {r['os_resumo']}"
        obs = (r["observacoes"] or "").strip()
        if len(obs) > 90:
            obs = obs[:90] + "..."

        html += f"""
            <tr>
                <td>{fmt(r['data'])}</td>
                <td>{r['item_paint']}</td>
                <td>{os_visual}</td>
                <td>{r['atividade']}</td>
                <td title="{r['observacoes'] or ''}">{obs}</td>
                <td>{r['duracao']}</td>
                <td style="white-space: nowrap;">
                <a class='btn' href='/editar/{r["id"]}'>Editar</a>
                <a class='btn' style='background:#c0392b; margin-left:5px;'
                   href='/excluir_hora/{r["id"]}'
                   onclick="return confirm('Confirma a exclusão deste lançamento?')">
                   Excluir
                </a>
                </td>

            </tr>
        """

    html += "</table>"

    if perfil == "admin":
        html += """
            <div style='margin-top:10px'>
                <a class='btn' href='/export'>Exportar todas as horas (CSV)</a>
                <button class='btn' style='margin-left:10px' onclick='exportarFiltrado()'>
                    Exportar filtrado
                </button>
                    <a class="btn" style="margin-left:10px" href="/export_preventivas">
        Exportar Preventivas
    </a>
            </div>

            <form id='formExportFiltrad
            o' method='POST' action='/export_filtrado'>
                <input type='hidden' name='ids' id='ids_filtrados'>
            </form>
        """

    # ---------------- SCRIPTS ----------------
    html += """
    <script>
    // FILTRO GERAL
    document.getElementById("filtroGeral").addEventListener("keyup", function() {
        let filtro = this.value.toLowerCase();
        let linhas = document.querySelectorAll("#tabelaMarcacoes tr");

        linhas.forEach((tr, i) => {
            if (i === 0) return; // pula cabeçalho
            tr.style.display = tr.innerText.toLowerCase().includes(filtro) ? "" : "none";
        });
    });


function exportarFiltrado() {
    let ids = [];
    document.querySelectorAll("#tabelaMarcacoes tr").forEach((tr, i) => {
        if (i === 0) return; // pula cabeçalho
        if (tr.style.display === "none") return; // pula linhas ocultas

        let id = tr.querySelector("a").href.split("/editar/")[1];
        if (id) ids.push(id);
    });

    if (ids.length === 0) {
        alert("Nenhum registro filtrado para exportar.");
        return;
    }

    document.getElementById("ids_filtrados").value = ids.join(",");
    document.getElementById("formExportFiltrado").submit();
}
    </script>
    """

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        user=session['user'],
        perfil=session['perfil']
    )

# -------------------------
# Editar Registro de Hora
# -------------------------

# -------------------------
# Editar Registro de Hora
# -------------------------
@app.route("/editar/<int:hid>", methods=["GET", "POST"])
def editar(hid):
    if "user" not in session:
        return redirect("/")

    from datetime import datetime, date

    con = get_db()
    cur = con.cursor()

    # -------------------------
    # Registro base (o que veio do relatório)
    # -------------------------
    cur.execute("SELECT * FROM horas WHERE id=%s", (hid,))
    base = cur.fetchone()
    if not base:
        con.close()
        return "Registro não encontrado."

    # -------------------------
    # Segurança
    # -------------------------
    if session["perfil"] != "admin" and base["colaborador_id"] != session["user_id"]:
        con.close()
        return "Acesso negado."

    # -------------------------
    # Registros a editar
    # 👉 SOMENTE o que aparece como agrupado no relatório
    # -------------------------
    cur.execute("""
        SELECT *
        FROM horas
        WHERE colaborador_id = %s
          AND data = %s
          AND os_codigo = %s
          AND atividade = %s
          AND COALESCE(observacoes,'') = COALESCE(%s,'')
        ORDER BY hora_inicio
    """, (
        base["colaborador_id"],
        base["data"],
        base["os_codigo"],
        base["atividade"],
        base["observacoes"]
    ))
    registros = cur.fetchall()

    if not registros:
        con.close()
        return "Registro não encontrado."

    primeiro = registros[0]

    # -------------------------
    # OS
    # -------------------------
    cur.execute("SELECT codigo, item_paint, resumo FROM os ORDER BY codigo")
    oss = cur.fetchall()

    # -------------------------
    # Delegações
    # -------------------------
    cur.execute("""
        SELECT id, requisicoes, os_codigo, grau, data_inicio
        FROM delegacoes
        WHERE colaborador_id = %s
          AND status = 'Em Andamento'
    """, (base["colaborador_id"],))

    delegacoes = []
    ids = set()
    delegacao_atual_id = primeiro.get("delegacao_id")

    for d in cur.fetchall():
        row = dict(d)
        if isinstance(row.get("data_inicio"), (datetime, date)):
            row["data_inicio"] = row["data_inicio"].isoformat()
        delegacoes.append(row)
        ids.add(row["id"])

    # Caso a delegação usada esteja fora do status
    if delegacao_atual_id and delegacao_atual_id not in ids:
        cur.execute("""
            SELECT id, requisicoes, os_codigo, grau, data_inicio
            FROM delegacoes
            WHERE id = %s
        """, (delegacao_atual_id,))
        d = cur.fetchone()
        if d:
            row = dict(d)
            if isinstance(row.get("data_inicio"), (datetime, date)):
                row["data_inicio"] = row["data_inicio"].isoformat()
            delegacoes.append(row)

    # -------------------------
    # POST
    # -------------------------
    if request.method == "POST":

        os_codigo = request.form.get("os")
        item = request.form.get("item")
        atividade = request.form.get("atividade")
        delegacao_id = request.form.get("delegacao_id") or None
        observacoes = request.form.get("observacoes")

        ids_form = request.form.getlist("hora_id[]")
        datas = request.form.getlist("data[]")
        horas_ini = request.form.getlist("hora_ini[]")
        horas_fim = request.form.getlist("hora_fim[]")

        if not datas:
            con.close()
            return "Nenhum registro enviado."

        # 🔐 trava ano
        for d in datas:
            if datetime.strptime(d, "%Y-%m-%d").year != 2026:
                con.close()
                return "Só é permitido editar registros de 2026."

        ids_existentes = {r["id"] for r in registros}
        ids_enviados = set()

        for i in range(len(datas)):
            hid_atual = ids_form[i] or None

            ini = datetime.strptime(horas_ini[i], "%H:%M")
            fim = datetime.strptime(horas_fim[i], "%H:%M")
            minutos = (fim - ini).seconds // 60
            duracao = f"{minutos//60:02d}:{minutos%60:02d}"

            if hid_atual:
                hid_atual = int(hid_atual)
                ids_enviados.add(hid_atual)

                cur.execute("""
                    UPDATE horas SET
                        data=%s,
                        hora_inicio=%s,
                        hora_fim=%s,
                        duracao=%s,
                        duracao_minutos=%s,
                        os_codigo=%s,
                        item_paint=%s,
                        atividade=%s,
                        delegacao_id=%s,
                        observacoes=%s
                    WHERE id=%s
                """, (
                    datas[i],
                    horas_ini[i],
                    horas_fim[i],
                    duracao,
                    minutos,
                    os_codigo,
                    item,
                    atividade,
                    delegacao_id,
                    observacoes,
                    hid_atual
                ))
            else:
                cur.execute("""
                    INSERT INTO horas
                    (colaborador_id, data, item_paint, os_codigo, atividade,
                     delegacao_id, hora_inicio, hora_fim,
                     duracao, duracao_minutos, observacoes)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """, (
                    base["colaborador_id"],
                    datas[i],
                    item,
                    os_codigo,
                    atividade,
                    delegacao_id,
                    horas_ini[i],
                    horas_fim[i],
                    duracao,
                    minutos,
                    observacoes
                ))

        # 🗑️ remove os apagados pelo usuário
        for hid_del in ids_existentes - ids_enviados:
            cur.execute("DELETE FROM horas WHERE id=%s", (hid_del,))

        con.commit()
        con.close()
        return redirect("/relatorios")

    con.close()

    # -------------------------
    # HTML
    # -------------------------
    html = """
<h3>Editar Registro #{{ hid }}</h3>

<form method="post" style="max-width:650px">

<div id="registros">
{% for r in registros %}
<div class="registro">
    <input type="hidden" name="hora_id[]" value="{{ r.id }}">
    <input type="date" name="data[]" value="{{ r.data }}">
    <input type="time" name="hora_ini[]" value="{{ r.hora_inicio }}">
    <input type="time" name="hora_fim[]" value="{{ r.hora_fim }}">
    <button type="button" onclick="remover(this)">🗑</button>
</div>
{% endfor %}
</div>

<button type="button" onclick="adicionar()">➕ Adicionar registro</button>

<br><br>

<label>OS:</label>
<select name="os" id="os_select">
<option></option>
{% for o in oss %}
<option value="{{ o.codigo }}" data-item="{{ o.item_paint }}"
{% if o.codigo == primeiro.os_codigo %}selected{% endif %}>
{{ o.codigo }}{% if o.resumo %} - {{ o.resumo }}{% endif %}
</option>
{% endfor %}
</select>

<br>
Item:
<input name="item" id="item_paint" value="{{ primeiro.item_paint }}" readonly>

<br>
<div id="box_delegacao" style="display:none;">
<label>Delegação:</label>
<select name="delegacao_id" id="delegacao_select">
<option value=""></option>
</select>
</div>

<br>
Atividade:
<select name="atividade">
<option {% if primeiro.atividade.startswith("1") %}selected{% endif %}>1. Planejamento</option>
<option {% if primeiro.atividade.startswith("2") %}selected{% endif %}>2. Execução</option>
<option {% if primeiro.atividade.startswith("3") %}selected{% endif %}>3. Relatório</option>
</select>

<br>
Observações:
<textarea name="observacoes">{{ primeiro.observacoes or '' }}</textarea>

<br><br>
<button class="btn">Salvar Alterações</button>
<a class="btn" href="/relatorios">Cancelar</a>

<script>
const osSelect = document.getElementById("os_select");
const itemInput = document.getElementById("item_paint");
const delSelect = document.getElementById("delegacao_select");
const box = document.getElementById("box_delegacao");

const delegacoes = {{ delegacoes | tojson }};
const delegacaoAtual = {{ primeiro.delegacao_id if primeiro.delegacao_id else "null" }};

function atualizarDelegacoes() {
    delSelect.innerHTML = "<option value=''></option>";
    box.style.display = "none";
    let achou = false;

    delegacoes.forEach(d => {
        if (d.os_codigo === osSelect.value) {
            achou = true;
            let opt = document.createElement("option");
            opt.value = d.id;
            opt.textContent = "Reqs: " + d.requisicoes + " | Grau: " + d.grau;
            if (d.id == delegacaoAtual) opt.selected = true;
            delSelect.appendChild(opt);
        }
    });

    if (achou) box.style.display = "block";
}

osSelect.addEventListener("change", function () {
    itemInput.value = this.selectedOptions[0]?.dataset.item || "";
    atualizarDelegacoes();
});

atualizarDelegacoes();

function adicionar() {
    const base = document.querySelector(".registro");
    const clone = base.cloneNode(true);

    clone.querySelector("input[name='hora_id[]']").value = "";
    clone.querySelectorAll("input[type='date'], input[type='time']").forEach(i => i.value = "");

    document.getElementById("registros").appendChild(clone);
}

function remover(btn) {
    const registros = document.querySelectorAll(".registro");
    if (registros.length > 1) {
        btn.parentElement.remove();
    } else {
        alert("É necessário manter pelo menos um registro.");
    }
}

</script>
"""

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        registros=registros,
        delegacoes=delegacoes,
        primeiro=primeiro,
        hid=hid,
        oss=oss,
        user=session["user"],
        perfil=session["perfil"]
    )

# -------------------------
# Excluir lançamento de horas
# -------------------------
@app.route('/excluir_hora/<int:id>')
def excluir_hora(id):
    if 'user' not in session:
        return redirect('/')

    con = get_db()
    cur = con.cursor()

    # -------------------------------------------------
    # Buscar registro (para validar dono e ajustar HH)
    # -------------------------------------------------
    cur.execute("""
        SELECT colaborador_id, item_paint, duracao_minutos
        FROM horas
        WHERE id = %s
    """, (id,))
    h = cur.fetchone()

    if not h:
        con.close()
        return "Registro não encontrado"

    # -------------------------------------------------
    # Segurança: colaborador só exclui o próprio registro
    # admin pode excluir qualquer um
    # -------------------------------------------------
    if session['perfil'] != 'admin' and h['colaborador_id'] != session['user_id']:
        con.close()
        return "Acesso negado"

    # -------------------------------------------------
    # Excluir registro
    # -------------------------------------------------
    cur.execute("DELETE FROM horas WHERE id = %s", (id,))
    con.commit()
    con.close()

    return redirect('/relatorios')

# -------- ADMIN - GERENCIAR PROJETOS --------
@app.route("/admin_projetos")
def admin_projetos():
    if session.get("perfil") != "admin":
        return redirect("/")

    def icon(v):
        if v == 1:
            return "<span style='color:green; font-weight:bold;'>✔</span>"
        else:
            return "<span style='color:red; font-weight:bold;'>✖</span>"

    conn = get_db()
    cur = conn.cursor()

    # =============================
    # 1) TOTAL DE REGISTROS
    # =============================
    cur.execute("SELECT COUNT(*) AS total FROM projeto_paint")
    total_paint = cur.fetchone()["total"]

    cur.execute("SELECT COUNT(*) AS total FROM os")
    total_os = cur.fetchone()["total"]

    # =============================
    # 2) TOTAL HH (soma hh_atual dos projetos)
    # =============================
    cur.execute("SELECT SUM(COALESCE(hh_atual,0)) AS total_hh FROM projeto_paint")
    total_hh = cur.fetchone()["total_hh"] or 0

    # =============================
    # 3) HH executadas
    # =============================
    cur.execute("""
        SELECT SUM(
            (CAST(SUBSTR(duracao, 1, 2) AS INTEGER) * 60) +
            CAST(SUBSTR(duracao, 4, 2) AS INTEGER)
        ) AS minutos
        FROM horas
    """)
    total_exec_min = cur.fetchone()["minutos"] or 0

    exec_hh = total_exec_min // 60
    exec_mm = total_exec_min % 60
    total_exec_hhmm = f"{int(exec_hh):02d}:{int(exec_mm):02d}"

    total_hh_min = int(total_hh) * 60
    percentual_global = (total_exec_min / total_hh_min * 100.0) if total_hh_min > 0 else 0
    percentual_global_fmt = f"{percentual_global:.2f}%"

    # =============================
    # 4) LISTAR PROJETOS PAINT
    # =============================
    cur.execute("SELECT * FROM projeto_paint ORDER BY item_paint")
    paint_rows = cur.fetchall()

    # ---- horas executadas por item_paint (1 query só) ----
    cur.execute("""
        SELECT 
            item_paint,
            SUM(
                (CAST(SUBSTR(duracao, 1, 2) AS INTEGER) * 60) +
                CAST(SUBSTR(duracao, 4, 2) AS INTEGER)
            ) AS minutos
        FROM horas
        GROUP BY item_paint
    """)

    horas_por_paint = {
        row["item_paint"]: row["minutos"] or 0
        for row in cur.fetchall()
    }

    paint_data = []

    for r in paint_rows:
        minutos = horas_por_paint.get(r["item_paint"], 0)

        hh = minutos // 60
        mm = minutos % 60
        soma = f"{hh:02d}:{mm:02d}"

        percentual_fmt = (
        f"{((minutos / 60) / r['hh_atual'] * 100):.2f}%"
        if r["hh_atual"] else "0%"
        )

        paint_data.append({
        "classificacao": r["classificacao"],
        "item_paint": r["item_paint"],
        "tipo_atividade": r["tipo_atividade"],
        "objeto": r["objeto"],
        "objetivo_geral": r["objetivo_geral"],
        "dt_ini": fmt(r["dt_ini"]),
        "dt_fim": fmt(r["dt_fim"]),
        "hh_atual": r["hh_atual"],
        "hh_exec": soma,
        "percentual": percentual_fmt
        })


    # =============================
    # 5) LISTAR OS
    # =============================
    cur.execute("SELECT * FROM os ORDER BY codigo")
    os_rows = cur.fetchall()
    
    os_data = []
    for r in os_rows:
        os_data.append({
            "codigo": r["codigo"],
            "item_paint": r["item_paint"],
            "resumo": r["resumo"],
            "unidade": r["unidade"],
            "coordenacao": r["coordenacao"],
            "equipe": r["equipe"],
            "observacao": r["observacao"],
            "status": r["status"],
            "plan": r["plan"],
            "exec": r["exec"],
            "rp": r["rp"],
            "rf": r["rf"],
            "dt_conclusao": fmt(r["dt_conclusao"])
        })

    # =============================
    # 6) HTML
    # =============================
    def hh_from_hours_integer(h):
        try:
            h = int(h or 0)
        except:
            h = 0
        return f"{h:02d}:00"

    total_hh_display = hh_from_hours_integer(total_hh)

    html = """
    <style>
        body, html { margin:0; padding:0; width:100vw; }
        .container, #content { width:100% !important; max-width:100% !important; margin:0 !important; padding:10px 20px !important; }
        table { width:100% !important; border-collapse: collapse; margin-bottom: 20px; }
        table th, table td { padding:10px 15px; border:1px solid #ccc; text-align:left; vertical-align:top; }
        table th { background:#f0f0f0; }
        input[type="text"] { width:400px; padding:6px 10px; margin-bottom:10px; }
    </style>
    
<style>
    .navbar {
        background: #f4f4f4;
        padding: 10px 15px;
        border-radius: 6px;
        margin-bottom: 20px;
        border: 1px solid #ddd;
        display: flex;
        gap: 15px;
        flex-wrap: wrap;
    }
    .navbar a {
        text-decoration: none;
        padding: 6px 12px;
        background: #1976d2;
        color: white;
        border-radius: 6px;
        font-weight: bold;
        transition: 0.2s;
    }
    .navbar a:hover {
        background: #0f4fa8;
    }
</style>

<div class="navbar">
    <a href='/menu'>Menu</a>
    <a href='/lancar'>Lançar Horas</a>
    <a href='/relatorios'>Relatórios</a>

    <a href='/colaboradores'>Colaboradores</a>
    <a href='/paint'>Projetos PAINT</a>
    <a href='/os'>O.S</a>
    <a href='/admin_projetos'>Gerenciar Projetos</a>
    <a href='/visao'>Visão Consolidada</a>

    <a href='/logout'>Sair</a>
</div>
    <h2>Gerenciar Projetos</h2>
    <div style='display:flex; gap:20px; margin-bottom:18px; flex-wrap:wrap;'>
    """

    # === cards + gauge ===
    # garantir percent 0..100
    percent = max(0.0, min(100.0, percentual_global))
    # parâmetros do SVG
    import math
    circumference = 2 * math.pi * 30
    dash = (percent / 100.0) * circumference

    html += f"""
    <div style="display:flex; gap:20px; margin-bottom:18px; align-items:stretch; flex-wrap:wrap;">

        <div style='padding:18px; background:#d6e4ff; border-radius:12px;
                    text-align:center; flex:1; min-width:160px;
                    border:1px solid #9bbcff; box-shadow:0 2px 6px rgba(0,0,0,0.08);'>
            <h4 style='margin:6px 0; color:#1e3a8a;'>Total PAINT</h4>
            <p style='font-size:28px; font-weight:700; margin:6px 0; color:#1e40af;'>{total_paint}</p>
        </div>

        <div style='padding:18px; background:#d1fae5; border-radius:12px;
                    text-align:center; flex:1; min-width:160px;
                    border:1px solid #6ee7b7; box-shadow:0 2px 6px rgba(0,0,0,0.08);'>
            <h4 style='margin:6px 0; color:#065f46;'>Total OS</h4>
            <p style='font-size:28px; font-weight:700; margin:6px 0; color:#047857;'>{total_os}</p>
        </div>

        <div style='padding:18px; background:#e0f2fe; border-radius:12px;
                    text-align:center; flex:1; min-width:200px;
                    border:1px solid #7dd3fc; box-shadow:0 2px 6px rgba(0,0,0,0.08);'>
            <h4 style='margin:6px 0; color:#075985;'>Total HH (planejado)</h4>
            <p style='font-size:22px; font-weight:700; margin:6px 0; color:#0369a1;'>{total_hh_display}</p>
            <div class='small' style='color:#334155'>soma de hh_atual dos projetos</div>
        </div>

        <div style='padding:18px; background:#ffedd5; border-radius:12px;
                    text-align:center; flex:1; min-width:200px;
                    border:1px solid #fdba74; box-shadow:0 2px 6px rgba(0,0,0,0.08);'>
            <h4 style='margin:6px 0; color:#9a3412;'>HH Executadas</h4>
            <p style='font-size:22px; font-weight:700; margin:6px 0; color:#c2410c;'>{total_exec_hhmm}</p>
            <div class='small' style='color:#4b5563'>total de horas registradas</div>
        </div>

        <div style='padding:12px; background:#e8f1ff; border-radius:12px;
                    width:220px; text-align:center; min-width:220px;
                    border:1px solid #93c5fd; box-shadow:0 2px 6px rgba(0,0,0,0.08);'>
            <h4 style='margin:6px 0; color:#1e3a8a;'>% Executado</h4>

            <svg width='120' height='120' viewBox='0 0 120 120' style='display:block;margin:auto'>
              <defs>
                <linearGradient id='gaugeGrad' x1='0%' y1='0%' x2='100%' y2='0%'>
                  <stop offset='0%' stop-color='#2563eb'/>
                  <stop offset='100%' stop-color='#06b6d4'/>
                </linearGradient>
              </defs>
              <g transform='translate(60,60)'>
                <circle r='44' fill='transparent' stroke='#c7d2fe' stroke-width='16'/>
                <circle r='38' fill='transparent' stroke='url(#gaugeGrad)' stroke-width='12'
                        stroke-dasharray='{dash:.2f} {circumference - dash:.2f}'
                        stroke-linecap='round' transform='rotate(-90)' />
                <text x='0' y='6' text-anchor='middle' font-size='18' font-weight='700' fill='#1e40af'>
                    {percent:.2f}%
                </text>
              </g>
            </svg>

            <div class='small' style='color:#334155; margin-top:6px;'>
                {percentual_global_fmt} do total planejado
            </div>
        </div>

    </div>
    """

    # PAINT Table
    html += """
    <h3>Projetos PAINT</h3>
    <input type='text' id='searchPaint' onkeyup="filterTable('searchPaint','paintTable')" placeholder='Pesquisar...'>
    <table id='paintTable'>
        <tr>
            <th>Classificação</th><th>Item</th><th>Tipo</th><th>Objeto</th><th>Objetivo Geral</th>
            <th>Início</th><th>Fim</th><th>HH Atual</th><th>HH Executada</th><th>% Executado</th>
        </tr>
    """
    for r in paint_data:
        html += f"""
        <tr>
            <td>{r['classificacao']}</td><td>{r['item_paint']}</td><td>{r['tipo_atividade']}</td>
            <td>{r['objeto']}</td><td>{r['objetivo_geral']}</td><td>{r['dt_ini']}</td><td>{r['dt_fim']}</td>
            <td>{r['hh_atual']}</td><td>{r['hh_exec']}</td><td>{r['percentual']}</td>
        </tr>
        """
    html += "</table>"

    # OS Table
    html += """
    <h3>Ordens de Serviço (OS)</h3>
    <input type='text' id='searchOS' onkeyup="filterTable('searchOS','osTable')" placeholder='Pesquisar...'>
    <table id='osTable'>
        <tr>
            <th>Código</th><th>Item PAINT</th><th>Resumo</th><th>Unidade</th><th>Coordenação</th>
            <th>Equipe</th><th>Observação</th><th>Status</th><th>PLAN</th><th>EXEC</th>
            <th>RP</th><th>RF</th><th>Conclusão</th>
        </tr>
    """
    for r in os_data:
        html += f"""
        <tr>
            <td>{r['codigo']}</td><td>{r['item_paint']}</td><td>{r['resumo']}</td><td>{r['unidade']}</td>
            <td>{r['coordenacao']}</td><td>{r['equipe']}</td><td>{r['observacao']}</td><td>{r['status']}</td>
            <td>{icon(r['plan'])}</td><td>{icon(r['exec'])}</td><td>{icon(r['rp'])}</td><td>{icon(r['rf'])}</td>
            <td>{r['dt_conclusao']}</td>
        </tr>
        """
    html += "</table>"

    # JS filter
    html += """
    <script>
    function filterTable(inputId, tableId) {
        var input = document.getElementById(inputId);
        var filter = input.value.toLowerCase();
        var table = document.getElementById(tableId);
        var tr = table.getElementsByTagName("tr");
        for (var i = 1; i < tr.length; i++) {
            var tds = tr[i].getElementsByTagName("td");
            var found = false;
            for (var j = 0; j < tds.length; j++) {
                if (tds[j] && tds[j].innerText.toLowerCase().indexOf(filter) > -1) {
                    found = true; break;
                }
            }
            tr[i].style.display = found ? "" : "none";
        }
    }
    </script>
    """

    return html

@app.route('/visao')
def visao_consolidada():
    if 'user' not in session:
        return redirect('/')

    import json
    con = get_db()
    cur = con.cursor()

    MESES = [
        ("01", "janeiro"), ("02", "fevereiro"), ("03", "marco"),
        ("04", "abril"), ("05", "maio"), ("06", "junho"),
        ("07", "julho"), ("08", "agosto"), ("09", "setembro"),
        ("10", "outubro"), ("11", "novembro"), ("12", "dezembro"),
    ]

    # ============================================================
    # TABELA 1 – HORAS POR COLABORADOR
    # ============================================================
    selects_mes = []
    for num, nome in MESES:
        selects_mes.append(
            f"""
            SUM(CASE WHEN EXTRACT(MONTH FROM h.data) = {int(num)}
                THEN h.duracao_minutos ELSE 0 END) AS {nome}
            """
        )

    sql_colab = f"""
        SELECT c.nome,
               {",".join(selects_mes)},
               SUM(h.duracao_minutos) AS total
        FROM horas h
        JOIN colaboradores c ON c.id = h.colaborador_id
        GROUP BY c.nome
        ORDER BY c.nome
    """

    cur.execute(sql_colab)
    tabela_colab = cur.fetchall()

    total_colab = {nome: 0 for _, nome in MESES}
    total_colab["total"] = 0

    for row in tabela_colab:
        for _, nome in MESES:
            total_colab[nome] += row[nome] or 0
        total_colab["total"] += row["total"] or 0

    # ============================================================
    # TABELA 2 – HORAS POR OS
    # ============================================================
    selects_mes_os = []
    for num, nome in MESES:
        selects_mes_os.append(
            f"""
            SUM(CASE WHEN EXTRACT(MONTH FROM h.data) = {int(num)}
                THEN h.duracao_minutos ELSE 0 END) AS {nome}
            """
        )

    sql_os = f"""
        SELECT o.codigo, o.item_paint, o.resumo,
               p.tipo_atividade, p.objeto, p.objetivo_geral,
               {",".join(selects_mes_os)},
               SUM(h.duracao_minutos) AS total
        FROM horas h
        JOIN os o ON o.codigo = h.os_codigo
        LEFT JOIN projeto_paint p ON p.item_paint = o.item_paint
        GROUP BY o.codigo, o.item_paint, o.resumo,
                 p.tipo_atividade, p.objeto, p.objetivo_geral
        ORDER BY o.item_paint, o.codigo
    """

    cur.execute(sql_os)
    tabela_os = cur.fetchall()

    total_os = {nome: 0 for _, nome in MESES}
    total_os["total"] = 0

    for row in tabela_os:
        for _, nome in MESES:
            total_os[nome] += row[nome] or 0
        total_os["total"] += row["total"] or 0

    # ============================================================
    # GRÁFICO – UMA QUERY SÓ (🔥 ganho grande)
    # ============================================================
    cur.execute("""
        SELECT EXTRACT(MONTH FROM data) AS mes,
               SUM(duracao_minutos) AS minutos
        FROM horas
        GROUP BY mes
        ORDER BY mes
    """)

    mapa = {int(r["mes"]): round((r["minutos"] or 0) / 60, 2) for r in cur.fetchall()}
    totais_mensais = [mapa.get(i, 0) for i in range(1, 13)]
    labels_meses = [nome.capitalize() for _, nome in MESES]

    con.close()

    # ---------------------------------------------------------------------- #
    #                    HTML – RENDERIZAÇÃO                                #
    # ---------------------------------------------------------------------- #

    html = """
    <h2>Visão Consolidada</h2>

    <div class='card'>
        <h3>Horas por Colaborador</h3>
        <input id='filtroColab' onkeyup='filtrar("tabelaColab","filtroColab")' placeholder='Pesquisar colaborador...'>
        <table id='tabelaColab'>
            <thead>
                <tr>
                    <th>Nome</th>
    """

    for _, nome in MESES:
        html += f"<th>{nome.capitalize()}</th>"

    html += "<th>Total Geral</th></tr></thead><tbody>"

    # ---- Corpo ----
    for row in tabela_colab:
        html += "<tr>"
        html += f"<td>{row['nome']}</td>"
        for _, nome in MESES:
            minutos = row[nome] or 0
            html += f"<td>{minutos//60:02d}:{minutos%60:02d}</td>"
        tg = row["total"] or 0
        html += f"<td><b>{tg//60:02d}:{tg%60:02d}</b></td>"
        html += "</tr>"

    # ---- Linha TOTALIZADORA ----
    html += "<tr style='font-weight:bold;background:#eef;'>"
    html += "<td>TOTAL</td>"
    for _, nome in MESES:
        m = total_colab[nome]
        html += f"<td>{m//60:02d}:{m%60:02d}</td>"
    t = total_colab["total"]
    html += f"<td>{t//60:02d}:{t%60:02d}</td>"
    html += "</tr>"

    html += "</tbody></table></div>"
    
    # ------------------------- TABELA 2 (OS) -------------------------
    html += """
    <div class='card'>
        <h3>Horas por O.S</h3>
        <input id='filtroOS' onkeyup='filtrar("tabelaOS","filtroOS")' placeholder='Pesquisar OS / PAINT / Atividade...'>
        <table id='tabelaOS'>
            <thead>
                <tr>
                    <th>Código OS</th>
                    <th>Item PAINT</th>
                    <th>Resumo</th>
    """

    for _, nome in MESES:
        html += f"<th>{nome.capitalize()}</th>"

    html += "<th>Total Geral</th></tr></thead><tbody>"

    # ---- Corpo ----
    for row in tabela_os:
        html += "<tr>"
        html += f"<td>{row['codigo']}</td>"
        html += f"<td>{row['item_paint']}</td>"
        html += f"<td>{row['resumo'] or ''}</td>"

        for _, nome in MESES:
            minutos = row[nome] or 0
            html += f"<td>{minutos//60:02d}:{minutos%60:02d}</td>"

        tg = row["total"] or 0
        html += f"<td><b>{tg//60:02d}:{tg%60:02d}</b></td>"
        html += "</tr>"

    # ---- TOTALIZADOR ----
    html += "<tr style='font-weight:bold;background:#eef;'>"
    html += "<td colspan='3'>TOTAL</td>"
    for _, nome in MESES:
        m = total_os[nome]
        html += f"<td>{m//60:02d}:{m%60:02d}</td>"
    t = total_os["total"]
    html += f"<td>{t//60:02d}:{t%60:02d}</td>"
    html += "</tr>"

    html += "</tbody></table></div>"

    # Filtro JS
    html += """
    <script>
    function filtrar(idTabela, idFiltro) {
        let filtro = document.getElementById(idFiltro).value.toLowerCase();
        let linhas = document.getElementById(idTabela).getElementsByTagName("tr");

        for (let i = 1; i < linhas.length; i++) {
            let texto = linhas[i].innerText.toLowerCase();
            linhas[i].style.display = texto.includes(filtro) ? "" : "none";
        }
    }
    </script>
    """

    # -------------------- FILTROS (LISTAS) --------------------
    cur = get_db().cursor()

    cur.execute("SELECT DISTINCT nome FROM colaboradores ORDER BY nome")
    lista_colab = [r["nome"] for r in cur.fetchall()]
    
    cur.execute("SELECT DISTINCT item_paint FROM os WHERE item_paint IS NOT NULL ORDER BY item_paint")
    lista_paint = [r["item_paint"] for r in cur.fetchall()]
    
    cur.execute("SELECT DISTINCT codigo FROM os ORDER BY codigo")
    lista_os = [r["codigo"] for r in cur.fetchall()]

    f_colab = request.args.get("colaborador")
    f_paint = request.args.get("item_paint")
    f_os = request.args.get("os")

    filtros = []
    params = {}

    if f_colab:
        filtros.append("c.nome = %(colab)s")
        params["colab"] = f_colab
    
    if f_paint:
        filtros.append("o.item_paint = %(paint)s")
        params["paint"] = f_paint
    
    if f_os:
        filtros.append("o.codigo = %(os)s")
        params["os"] = f_os
    
    where_sql = ""
    if filtros:
        where_sql = "WHERE " + " AND ".join(filtros)


    # ---------------------- GRÁFICO (TOTAL MENSAL) ----------------------
    # Prepara lista de totais mensais (em horas float)
    totais_mensais = []

    for num, _ in MESES:
        where_mes = "AND" if where_sql else "WHERE"

        cur.execute(f"""
            SELECT
                SUM(
                    CAST(SUBSTR(h.duracao,1,2) AS INTEGER) * 60 +
                    CAST(SUBSTR(h.duracao,4,2) AS INTEGER)
                ) AS minutos
            FROM horas h
            LEFT JOIN colaboradores c ON c.id = h.colaborador_id
            LEFT JOIN os o ON o.codigo = h.os_codigo
            {where_sql}
            {where_mes} EXTRACT(MONTH FROM h.data) = {int(num)}
        """, params)

        minutos = cur.fetchone()["minutos"] or 0
        totais_mensais.append(round(minutos / 60, 2))

        labels_meses = [nome.capitalize() for _, nome in MESES]
    html += f"""
<div class='card'>
    <h3>Total de Horas por Mês</h3>

    <form method="get" style="margin-bottom:15px;">
        <label>Colaborador:</label>
        <select name="colaborador">
            <option value="">Todos</option>
            {''.join(f"<option {'selected' if f_colab==c else ''}>{c}</option>" for c in lista_colab)}
        </select>

        <label>Item PAINT:</label>
        <select name="item_paint">
            <option value="">Todos</option>
            {''.join(f"<option {'selected' if f_paint==p else ''}>{p}</option>" for p in lista_paint)}
        </select>

        <label>OS:</label>
        <select name="os">
            <option value="">Todas</option>
            {''.join(f"<option {'selected' if f_os==o else ''}>{o}</option>" for o in lista_os)}
        </select>

        <button class="btn">Filtrar</button>
    </form>

    <canvas id="graficoHoras"></canvas>
</div>

<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels"></script>

<script>
new Chart(document.getElementById('graficoHoras'), {{
    type: 'bar',
    data: {{
        labels: {json.dumps(labels_meses)},
        datasets: [{{
            label: 'Total de horas',
            data: {json.dumps(totais_mensais)},
            borderWidth: 2
        }}]
    }},
    plugins: [ChartDataLabels],
    options: {{
        responsive: true,
        plugins: {{
            datalabels: {{
                anchor: 'end',
                align: 'top'
            }},
            tooltip: {{
                enabled: false
            }}
        }},
        scales: {{
            y: {{
                beginAtZero: true,
                title: {{
                    display: true,
                    text: 'Horas'
                }}
            }}
        }}
    }}
}});
</script>
"""

    return render_template_string(BASE.replace("{% block content %}{% endblock %}", html),
                                  user=session["user"], perfil=session["perfil"])

@app.route('/export')
def export_csv():
    if 'user' not in session:
        return redirect('/')

    import csv, io
    from datetime import datetime

    con = get_db()
    cur = con.cursor()

    cur.execute("""
        SELECT
            h.id,
            c.nome AS colaborador,
            h.data,
            h.hora_inicio,
            h.hora_fim,
            h.item_paint,
            h.os_codigo,
            h.atividade,
            h.duracao,
            h.observacoes,
            d.data_inicio AS data_inicio_delegacao,
            d.data_fim AS data_fim_delegacao,
            d.requisicoes,
            d.status AS status_delegacao,
            d.grau AS grau_delegacao,
            d.id as delegacao_id
        FROM horas h
        LEFT JOIN colaboradores c ON c.id = h.colaborador_id
        LEFT JOIN delegacoes d ON d.id = h.delegacao_id
        ORDER BY h.data
    """)
    rows = cur.fetchall()
    con.close()

    si = io.StringIO()
    cw = csv.writer(si, delimiter=";")

    # Cabeçalho
    cw.writerow([
        "ID", "Colaborador", "Data", "Hora Início", "Hora Fim",
        "Item PAINT", "OS", "Atividade", "Duração", "Obs",
        "Data Início Delegação", "Data Fim Delegação",
        "Requisições", "Qtd Requisições",
        "Status Delegação", "Grau", "delegacao_id"
    ])

    for r in rows:
        requisicoes = r["requisicoes"] or ""
        qtd_req = len([x for x in requisicoes.split(",") if x.strip()]) if requisicoes else 0

        # formatar data
        try:
            data_fmt = datetime.strptime(r["data"], "%Y-%m-%d").strftime("%d/%m/%Y")
        except:
            data_fmt = r["data"]

        cw.writerow([
            r["id"],
            r["colaborador"],
            data_fmt,
            r["hora_inicio"],
            r["hora_fim"],
            r["item_paint"],
            r["os_codigo"],
            r["atividade"],
            r["duracao"],
            r["observacoes"],
            r["data_inicio_delegacao"],
            r["data_fim_delegacao"],
            requisicoes,
            qtd_req,
            r["status_delegacao"],
            r["grau_delegacao"],
            r["delegacao_id"]
        ])

    output = io.BytesIO()
    output.write("\ufeff".encode("utf-8"))
    output.write(si.getvalue().encode("utf-8"))
    output.seek(0)

    return send_file(
        output,
        mimetype="text/csv",
        as_attachment=True,
        download_name="horas_completo.csv"
    )

@app.route('/export_filtrado', methods=['POST'])
def export_filtrado():
    if 'user' not in session or session['perfil'] != 'admin':
        return redirect('/')

    ids_raw = request.form.get("ids", "")
    if not ids_raw:
        return "Nenhum ID recebido.", 400

    ids = ids_raw.split(",")

    import csv
    from io import StringIO
    from datetime import datetime

    con = get_db()
    cur = con.cursor()

    sql = f"""
        SELECT
            h.id,
            h.data,
            h.hora_inicio,
            h.hora_fim,
            h.item_paint,
            h.os_codigo,
            h.atividade,
            h.duracao,
            h.observacoes,
            d.data_inicio AS data_inicio_delegacao,
            d.data_fim AS data_fim_delegacao,
            d.requisicoes,
            d.status AS status_delegacao,
            d.grau AS grau_delegacao
        FROM horas h
        LEFT JOIN delegacoes d ON d.id = h.delegacao_id
        WHERE h.id IN ({",".join(["%s"] * len(ids))})
        ORDER BY h.id
    """

    cur.execute(sql, ids)
    rows = cur.fetchall()
    con.close()

    output = StringIO()
    output.write("\ufeff")
    writer = csv.writer(output, delimiter=";")

    writer.writerow([
        "ID", "Data", "Hora Início", "Hora Fim",
        "Item PAINT", "OS", "Atividade", "Duração", "Obs",
        "Data Início Delegação", "Data Fim Delegação",
        "Requisições", "Qtd Requisições",
        "Status Delegação", "Grau"
    ])

    for r in rows:
        requisicoes = r["requisicoes"] or ""
        qtd_req = len([x for x in requisicoes.split(",") if x.strip()]) if requisicoes else 0

        try:
            data_fmt = datetime.strptime(r["data"], "%Y-%m-%d").strftime("%d/%m/%Y")
        except:
            data_fmt = r["data"]

        writer.writerow([
            r["id"],
            data_fmt,
            r["hora_inicio"],
            r["hora_fim"],
            r["item_paint"],
            r["os_codigo"],
            r["atividade"],
            r["duracao"],
            r["observacoes"],
            r["data_inicio_delegacao"],
            r["data_fim_delegacao"],
            requisicoes,
            qtd_req,
            r["status_delegacao"],
            r["grau_delegacao"]
        ])

    return Response(
        output.getvalue(),
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": "attachment; filename=horas_filtradas_completo.csv"}
    )


def minutos_para_hhmm(minutos):
    horas = minutos // 60
    mins = minutos % 60
    return f"{int(horas):02d}:{int(mins):02d}"

@app.route('/export_preventivas')
def export_preventivas():
    if 'user' not in session or session['perfil'] != 'admin':
        return redirect('/')

    import csv, io
    from datetime import datetime
    from collections import defaultdict
    from flask import Response

    con = get_db()
    cur = con.cursor()

    cur.execute("""
        SELECT
            c.nome AS colaborador,
            -- horas (podem ser NULL)
            h.data,
            h.hora_inicio,
            h.hora_fim,
            h.duracao_minutos,
            -- delegação
            d.id AS delegacao_id,
            d.data_inicio,
            d.data_fim,
            d.requisicoes,
            d.status,
            d.grau,
            d.criterio,
            d.os_codigo,
            -- OS
            o.item_paint
    
        FROM delegacoes d
        
        JOIN colaboradores c ON c.id = d.colaborador_id
        -- OS é obrigatória para a delegação
        LEFT JOIN os o ON o.codigo = d.os_codigo
        -- horas são opcionais
        LEFT JOIN horas h ON h.delegacao_id = d.id
    
        ORDER BY d.id, h.data, h.hora_inicio
    """)

    rows = cur.fetchall()
    con.close()

    # ---------------- AGRUPAR POR DELEGAÇÃO ----------------
    grupos = defaultdict(lambda: {
        "colaborador": "",
        "data_inicio": "",
        "data_fim": "",
        "datas": [],
        "hora_ini": [],
        "hora_fim": [],
        "duracao_total_min": 0,
        "os": "",
        "item": "",
        "requisicoes": "",
        "qtd_req": 0,
        "grau": "",
        "criterio": "",
        "status": ""
    })

    for r in rows:
        g = grupos[r["delegacao_id"]]

        g["colaborador"] = r["colaborador"]
        g["data_inicio"] = r["data_inicio"]
        g["data_fim"] = r["data_fim"]
        g["os"] = r["os_codigo"]
        g["item"] = r["item_paint"]
        g["requisicoes"] = r["requisicoes"] or ""
        g["grau"] = r["grau"]
        g["criterio"] = r["criterio"]
        g["status"] = r["status"]

        # 👉 só adiciona se existir hora lançada
        if r["data"]:
            if isinstance(r["data"], (datetime, date)):
                data_fmt = r["data"].strftime("%d/%m/%Y")
            else:
                data_fmt = str(r["data"])
        
            g["datas"].append(data_fmt)
            g["hora_ini"].append(str(r["hora_inicio"]))
            g["hora_fim"].append(str(r["hora_fim"]))
        
            g["duracao_total_min"] += r["duracao_minutos"] or 0


        if g["requisicoes"]:
            g["qtd_req"] = len([x for x in g["requisicoes"].split(",") if x.strip()])

    # ---------------- GERAR CSV ----------------
    output = io.StringIO()
    output.write("\ufeff")  # BOM Excel
    writer = csv.writer(output, delimiter=";")

    writer.writerow([
        "Colaborador",
        "Data Início Delegação",
        "Data Fim Delegação",
        "Datas Trabalhadas",
        "Horas Início",
        "Horas Fim",
        "Horas Totais",
        "Qtd Requisições",
        "Requisições",
        "Grau",
        "Criterio",
        "Status",
        "OS",
        "Item PAINT"
    ])

    for g in grupos.values():
        duracao_total = minutos_para_hhmm(g["duracao_total_min"])

        writer.writerow([
            g["colaborador"],
            g["data_inicio"],
            g["data_fim"],
            "\n".join(g["datas"]),        # ALT + ENTER
            "\n".join(g["hora_ini"]),     # ALT + ENTER
            "\n".join(g["hora_fim"]),     # ALT + ENTER
            duracao_total,
            g["qtd_req"],
            g["requisicoes"].replace(",", "\n"),  # opcional: reqs em linhas
            g["grau"],
            g["criterio"],
            g["status"],
            g["os"],
            g["item"]
        ])

    return Response(
        output.getvalue(),
        mimetype="text/csv; charset=utf-8",
        headers={
            "Content-Disposition": "attachment; filename=preventivas.csv"
        }
    )

# -------------------------
# Delegar Requisições (ADMIN)
# -------------------------
@app.route("/delegar", methods=["GET", "POST"])
def delegar():
    if "user" not in session:
        return redirect("/")
    if session["perfil"] != "admin":
        return "Acesso negado"

    con = get_db()
    cur = con.cursor()

    # Carregar colaboradores
    cur.execute("SELECT id, nome FROM colaboradores ORDER BY nome")
    colaboradores = cur.fetchall()

    # Carregar OS permitidas
    cur.execute("""
        SELECT codigo, resumo FROM os 
        WHERE codigo IN ('1.4/2026','1.5/2026','1.6/2026')
    """)
    oss = cur.fetchall()

    if request.method == "POST":
        qtd = int(request.form.get("qtd"))
        requisicoes = []

        for i in range(qtd):
            req = request.form.get(f"req{i+1}")
            if req:
                requisicoes.append(req)

        req_str = ",".join(requisicoes)

        os_codigo = request.form.get("os_codigo")
        colaborador = request.form.get("colaborador_id")
        data_inicio = request.form.get("data_inicio")
        grau = request.form.get("grau")
        criterio = request.form.get("criterio")

        # Validação do ano
        if not data_inicio.startswith("2026-"):
            return "A data deve ser do ano de 2026"

        # Status inicial sempre "Em Andamento"
        status = "Em Andamento"

        cur.execute("""
            INSERT INTO delegacoes 
                (requisicoes, os_codigo, colaborador_id, data_inicio, status, grau, criterio)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """, (req_str, os_codigo, colaborador, data_inicio, status, grau, criterio))

        con.commit()
        con.close()
        return redirect("/menu")

    con.close()

    # ---- HTML ----
    html = """
    <h2>Delegar Requisições</h2>

    <form method="post">
        <div>Quantidade de requisições:
            <input type="number" name="qtd" id="qtd" min="1" max="50" required>
            <button type="button" onclick="gerar()">Gerar Campos</button>
        </div>

        <div id="campos"></div>

        <div>O.S:
            <select name="os_codigo" required>
                <option></option>
    """
    for os in oss:
        texto = os["codigo"]
        if os["resumo"]:
            texto += " - " + os["resumo"]

        html += f"""
            <option value="{os['codigo']}">
                {texto}
            </option>
    """
    
    html += """
            </select>
        </div>

        <div>Colaborador:
            <select name="colaborador_id" required>
                <option></option>
    """
    for c in colaboradores:
        html += f"<option value='{c['id']}'>{c['nome']}</option>"

    html += """
            </select>
        </div>

        <div>Tipo da Requisição:
            <select name="grau" required>
                <option></option>
                <option value="Contratação">Contratação</option>
                <option value="Liquidação">Liquidação</option>
                <option value="Aditamento">Aditamento</option>
            </select>
        </div>

        <div>Critério:
            <select name="criterio" required>
                <option></option>
                <option value="Materialidade">Materialidade</option>
                <option value="Relevância">Relevância</option>
                <option value="Risco">Risco</option>
                <option value="Engenharia">Engenharia</option>
            </select>
        </div>

        <div>Data de início:
            <input type="date" name="data_inicio" min="2026-01-01" max="2026-12-31" required>
        </div>

        <button class='btn'>Salvar</button>
    </form>

    <script>
    function gerar() {
        let qtd = document.getElementById("qtd").value;
        let div = document.getElementById("campos");
        div.innerHTML = "";
        for (let i = 1; i <= qtd; i++) {
            div.innerHTML += "<div>Requisição " + i + 
                             ": <input name='req" + i + "' required></div>";
        }
    }
    </script>
    """

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        user=session["user"],
        perfil=session["perfil"]
    )

# -------------------------
# Listar Delegações (ADMIN)
# -------------------------
@app.route("/delegacoes")
def listar_delegacoes():
    if "user" not in session:
        return redirect("/")
    if session["perfil"] != "admin":
        return "Acesso negado"

    con = get_db()
    cur = con.cursor()

   # ---------------- PAGINAÇÃO ----------------
    limit_param = request.args.get("limit", "50")
    
    if limit_param == "all":
        limite = None
    else:
        try:
            limite = int(limit_param)
        except:
            limite = 50

    # ---------------- FILTRO STATUS ----------------
    status_param = request.args.get("status", "Em Andamento")

    sql = """
        SELECT d.*, c.nome AS colaborador
        FROM delegacoes d
        LEFT JOIN colaboradores c ON c.id = d.colaborador_id
    """
    params = []

    # filtro por status
    if status_param != "all":
        sql += " WHERE d.status = %s "
        params.append(status_param)
    
    sql += " ORDER BY d.id DESC "
    
    # limite
    if limite is not None:
        sql += " LIMIT %s "
        params.append(limite)
    
    cur.execute(sql, tuple(params))
    delegacoes = cur.fetchall()
    con.close()

    html = """
    <h2>Delegações Cadastradas</h2>

<!-- PAGINAÇÃO -->
<div style="margin-bottom:6px">
    Mostrar:
    <a class="btn" href="/delegacoes?limit=20&status={{status}}">20</a>
    <a class="btn" href="/delegacoes?limit=50&status={{status}}">50</a>
    <a class="btn" href="/delegacoes?limit=100&status={{status}}">100</a>
    <a class="btn" href="/delegacoes?limit=200&status={{status}}">200</a>
    <a class="btn" href="/delegacoes?limit=all&status={{status}}">Todos</a>
</div>

<!-- FILTRO STATUS -->
<div style="margin-bottom:12px">
    <strong>Status:</strong><br>
    <a class="btn"
       href="/delegacoes?status=Em Andamento&limit={{limit_param}}">
       Em Andamento
    </a>

    <a class="btn"
       href="/delegacoes?status=Concluída&limit={{limit_param}}">
       Concluídas
    </a>

    <a class="btn"
       href="/delegacoes?status=all&limit={{limit_param}}">
       Todas
    </a>
</div>

<!-- FILTRO GERAL -->
<input type="text" id="filtroGeral"
       placeholder="Pesquisar em qualquer campo..."
       style="width:100%; padding:8px; margin-bottom:12px;">

<table id="tabelaDelegacoes" border=1 cellpadding=5>
    <tr>
        <th>ID</th>
        <th>Requisições</th>
        <th>O.S</th>
        <th>Colaborador</th>
        <th>Data Início</th>
        <th>Status</th>
        <th>Tipo</th>
        <th>Critério</th>
        <th>Ações</th>
    </tr>
    """

    for d in delegacoes:
        html += f"""
            <tr>
                <td>{d['id']}</td>
                <td class="col-requisicoes">{d['requisicoes']}</td>
                <td>{d['os_codigo']}</td>
                <td>{d['colaborador']}</td>
                <td>{fmt(d['data_inicio'])}</td>
                <td style="min-width:140px">
                <select style="width:100%" onchange="alterarStatus({d['id']}, this.value)">
                    <option {"selected" if d['status']=="Em Andamento" else ""}>Em Andamento</option>
                    <option {"selected" if d['status']=="Concluída" else ""}>Concluída</option>
                    <option {"selected" if d['status']=="Cancelada" else ""}>Cancelada</option>
                </select>
                </td>
                <td>{d['grau']}</td>
                <td>{d['criterio']}</td>
                <td>
                    <a href='/delegacao/{d['id']}'>Ver</a> |
                    <a href='/editar_delegacao/{d['id']}'>Editar</a> |
                    <a href='/excluir_delegacao/{d['id']}'
                       onclick="return confirm('Excluir delegação?')">Excluir</a>
                </td>
            </tr>
        """

    html += "</table>"

    html += """
    <script>
    function alterarStatus(id, status) {
        fetch('/alterar_status_delegacao', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                id: id,
                status: status
            })
        })
        .then(r => r.json())
        .then(resp => {
            if (!resp.ok) {
                alert(resp.msg || "Erro ao atualizar status");
            }
        })
        .catch(() => alert("Erro de comunicação com o servidor"));
    }
    </script>
    """

    html += """
    <script>
    // FILTRO GERAL
    document.getElementById("filtroGeral").addEventListener("keyup", function () {
        let filtro = this.value.toLowerCase();
        let linhas = document.querySelectorAll("#tabelaDelegacoes tr");

        linhas.forEach((tr, i) => {
            if (i === 0) return; // cabeçalho
            tr.style.display = tr.innerText.toLowerCase().includes(filtro)
                ? ""
                : "none";
        });
    });

    function alterarStatus(id, status) {
        fetch('/alterar_status_delegacao', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: id, status: status })
        })
        .then(r => r.json())
        .then(resp => {
            if (!resp.ok) alert(resp.msg || "Erro ao atualizar status");
        })
        .catch(() => alert("Erro de comunicação"));
    }
    </script>
    """
    
    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        user=session["user"], perfil=session["perfil"],
        status=status_param,
        limit_param=limit_param
    )

# -------------------------------------------------
# Atualizar status da delegação (inline)
# -------------------------------------------------
@app.route("/alterar_status_delegacao", methods=["POST"])
def alterar_status_delegacao():
    if "user" not in session:
        return {"ok": False, "msg": "Não autenticado"}, 403

    data = request.get_json()
    delegacao_id = data.get("id")
    novo_status = data.get("status")

    if novo_status not in ["Em Andamento", "Concluída", "Cancelada"]:
        return {"ok": False, "msg": "Status inválido"}, 400

    con = get_db()
    cur = con.cursor()

    # 🔎 Busca delegação
    cur.execute("""
        SELECT id, colaborador_id
        FROM delegacoes
        WHERE id = %s
    """, (delegacao_id,))
    delegacao = cur.fetchone()

    if not delegacao:
        con.close()
        return {"ok": False, "msg": "Delegação não encontrada"}, 404

    # 🔐 Permissão
    if session["perfil"] != "admin":
        if delegacao["colaborador_id"] != session["user_id"]:
            con.close()
            return {"ok": False, "msg": "Sem permissão"}, 403

        if novo_status == "Cancelada":
            con.close()
            return {"ok": False, "msg": "Ação não permitida"}, 403

    # 🧠 Data fim automática ao concluir
    data_fim = None
    if novo_status == "Concluída":
        cur.execute("""
            SELECT MAX(data) AS ultima_data
            FROM horas
            WHERE delegacao_id = %s
        """, (delegacao_id,))
        r = cur.fetchone()
        data_fim = r["ultima_data"] if r else None

    cur.execute("""
        UPDATE delegacoes
        SET status = %s,
            data_fim = %s
        WHERE id = %s
    """, (novo_status, data_fim, delegacao_id))

    con.commit()
    con.close()

    return {"ok": True}

@app.route("/minhas_delegacoes")
def minhas_delegacoes():
    if "user" not in session:
        return redirect("/")

    if session["perfil"] == "admin":
        return redirect("/delegacoes")

    con = get_db()
    cur = con.cursor()

    # ---------------- PAGINAÇÃO ----------------
    limit_param = request.args.get("limit", "50")
    status_param = request.args.get("status", "Em Andamento")

    if limit_param == "all":
        limite = None
    else:
        try:
            limite = int(limit_param)
        except:
            limite = 50

    sql = """
    SELECT d.*, c.nome AS colaborador
    FROM delegacoes d
    LEFT JOIN colaboradores c ON c.id = d.colaborador_id
    WHERE d.colaborador_id = %s
    """
    params = [session["user_id"]]
    
    # filtro por status (default = Em Andamento)
    if status_param:
        sql += " AND d.status = %s "
        params.append(status_param)
    
    sql += " ORDER BY d.id DESC "
    
    if limite:
        sql += " LIMIT %s "
        params.append(limite)

    cur.execute(sql, tuple(params))
    delegacoes = cur.fetchall()
    con.close()

    html = """
    <div class="card">
    <h2>Requisições Delegadas</h2>
    
    <!-- PAGINAÇÃO -->
   <div style="margin-bottom:10px">
    <strong>Mostrar:</strong><br>

    <a class="btn" href="/minhas_delegacoes?limit=20&status={{ status_param }}">20</a>
    <a class="btn" href="/minhas_delegacoes?limit=50&status={{ status_param }}">50</a>
    <a class="btn" href="/minhas_delegacoes?limit=100&status={{ status_param }}">100</a>
    <a class="btn" href="/minhas_delegacoes?limit=200&status={{ status_param }}">200</a>
    <a class="btn" href="/minhas_delegacoes?limit=all&status={{ status_param }}">Todos</a>

    <br><br>

    <strong>Filtrar por status:</strong><br>

    <a class="btn"
       style="background:#ffc107;color:#000;"
       href="/minhas_delegacoes?status=Em Andamento&limit={{ limit_param }}">
        Em Andamento
    </a>

    <a class="btn"
       style="background:#28a745;"
       href="/minhas_delegacoes?status=Concluída&limit={{ limit_param }}">
        Concluídas
    </a>
    </div>

    
    <!-- FILTRO GERAL -->
    <input type="text" id="filtroGeral"
           placeholder="Pesquisar em qualquer campo..."
           style="width:100%; padding:8px; margin-bottom:12px;">
    
    {% if delegacoes %}
    <table id="tabelaDelegacoes">
        <tr>
            <th>ID</th>
            <th>Requisições</th>
            <th>O.S</th>
            <th>Data Início</th>
            <th>Status</th>
            <th>Tipo</th>
            <th>Critério</th>
            <th>Ação</th>
        </tr>
    
        {% for d in delegacoes %}
        <tr>
            <td>{{ d.id }}</td>
            <td class="col-requisicoes">{{ d.requisicoes }}</td>
            <td>{{ d.os_codigo }}</td>
            <td>{{fmt(d.data_inicio)}}</td>
    
            <td>
                <select onchange="alterarStatus({{ d.id }}, this.value)">
                    <option value="Em Andamento"
                        {% if d.status == "Em Andamento" %}selected{% endif %}>
                        Em Andamento
                    </option>
            
                    <option value="Concluída"
                        {% if d.status == "Concluída" %}selected{% endif %}>
                        Concluída
                    </option>
                </select>
            </td>
            
            <td>{{ d.grau }}</td>
            <td>{{ d.criterio }}</td>
            <td>
                <a class="btn" href="/delegacao/{{ d.id }}">Ver</a>
            </td>
        </tr>
        {% endfor %}
    </table>
    {% else %}
        <p class="small">Nenhuma requisição delegada para você.</p>
    {% endif %}
</div>
    <script>
    document.getElementById("filtroGeral").addEventListener("keyup", function () {
        let filtro = this.value.toLowerCase();
        let linhas = document.querySelectorAll("#tabelaDelegacoes tr");

        linhas.forEach((tr, i) => {
            if (i === 0) return; // cabeçalho
            tr.style.display = tr.innerText.toLowerCase().includes(filtro)
                ? ""
                : "none";
        });
    });
    
    function alterarStatus(id, status) {
        fetch('/alterar_status_delegacao', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                id: id,
                status: status
            })
        })
        .then(r => r.json())
        .then(resp => {
            if (!resp.ok) {
                alert(resp.msg || "Erro ao atualizar status");
            }
        })
        .catch(() => alert("Erro de comunicação com o servidor"));
    }
    </script>
    """

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        delegacoes=delegacoes,
        fmt=fmt,
        status_param=status_param,
        limit_param=limit_param,
        user=session["user"],
        perfil=session["perfil"]
    )

@app.route("/excluir_delegacao/<int:id>")
def excluir_delegacao(id):
    if "user" not in session:
        return redirect("/")
    if session["perfil"] != "admin":
        return "Acesso negado"

    con = get_db()
    cur = con.cursor()
    cur.execute("DELETE FROM delegacoes WHERE id = %s", (id,))
    con.commit()
    con.close()

    return redirect("/delegacoes")

@app.route("/editar_delegacao/<int:id>", methods=["GET", "POST"])
def editar_delegacao(id):
    if "user" not in session:
        return redirect("/")
    if session["perfil"] != "admin":
        return "Acesso negado"

    con = get_db()
    cur = con.cursor()

    # ================= POST =================
    if request.method == "POST":
        requisicoes = request.form.get("requisicoes")
        os_codigo = request.form.get("os_codigo")
        colaborador_id = request.form.get("colaborador_id")
        data_inicio = request.form.get("data_inicio")
        status = request.form.get("status")
        grau = request.form.get("grau")
        criterio = request.form.get("criterio")
        data_fim = request.form.get("data_fim") or None
        cur.execute("SELECT os_codigo FROM delegacoes WHERE id=%s", (id,))
        os_antiga = cur.fetchone()["os_codigo"]
        
        if data_fim and not data_fim.startswith("2026-"):
            return "A data de conclusão deve ser dentro do ano de 2026."

        # -------------------------------
        # REGRA DE STATUS / DATA_FIM
        # -------------------------------
        if status == "Concluída":
            # se admin NÃO informou data_fim → pega última data de horas
            if not data_fim:
                cur.execute("""
                    SELECT MAX(data) AS ultima_data
                    FROM horas
                    WHERE delegacao_id = %s
                """, (id,))
                r = cur.fetchone()
                data_fim = r["ultima_data"] if r and r["ultima_data"] else None
        else:
            # se não for concluída → limpa data_fim
            data_fim = None
        
        cur.execute("SELECT os_codigo FROM delegacoes WHERE id=%s", (id,))
        os_antiga = cur.fetchone()["os_codigo"]
        
        cur.execute("""
            UPDATE delegacoes
            SET requisicoes=%s,
                os_codigo=%s,
                colaborador_id=%s,
                data_inicio=%s,
                status=%s,
                grau=%s,
                criterio=%s,
                data_fim=%s
            WHERE id=%s
        """, (
            requisicoes,
            os_codigo,
            colaborador_id,
            data_inicio,
            status,
            grau,
            criterio,
            data_fim,
            id
        ))
        
        # Atualiza horas vinculadas à delegação
        if os_antiga != os_codigo:
            cur.execute("""
                UPDATE horas
                SET os_codigo = %s
                WHERE delegacao_id = %s
            """, (os_codigo, id))
        
        con.commit()
        con.close()
        return redirect("/delegacoes")

    # ================= GET =================
    cur.execute("SELECT * FROM delegacoes WHERE id=%s", (id,))
    d = cur.fetchone()

    cur.execute("SELECT id, nome FROM colaboradores ORDER BY nome")
    colaboradores = cur.fetchall()

    cur.execute("""
        SELECT codigo, resumo
        FROM os
        WHERE codigo IN ('1.4/2026','1.5/2026','1.6/2026')
    """)
    oss = cur.fetchall()

    con.close()

    # ================= HTML =================
    html = f"""
    <h2>Editar Delegação #{id}</h2>

    <form method="post">
        Requisições:<br>
        <input name="requisicoes" value="{d['requisicoes']}" required><br><br>

        O.S:<br>
        <select name="os_codigo" required>
    """

    for os in oss:
        texto = os["codigo"]
        if os["resumo"]:
            texto += " - " + os["resumo"]

        sel = "selected" if os["codigo"] == d["os_codigo"] else ""
        html += f"""
            <option value="{os['codigo']}" {sel}>
                {texto}
            </option>
        """

    html += """
        </select><br><br>

        Colaborador:<br>
        <select name="colaborador_id" required>
    """

    for c in colaboradores:
        sel = "selected" if c["id"] == d["colaborador_id"] else ""
        html += f"<option value='{c['id']}' {sel}>{c['nome']}</option>"

    html += "</select><br><br>"

    html += f"""
        Data início:<br>
        <input type="date" name="data_inicio"
               value="{d['data_inicio']}"
               min="2026-01-01" max="2026-12-31"
               required><br><br>

        Status:<br>
        <select name="status">
            <option {"selected" if d['status']=="Em Andamento" else ""}>Em Andamento</option>
            <option {"selected" if d['status']=="Concluída" else ""}>Concluída</option>
            <option {"selected" if d['status']=="Cancelada" else ""}>Cancelada</option>
        </select><br><br>

        Grau:<br>
        <select name="grau">
            <option {"selected" if d['grau']=="Contratação" else ""}>Contratação</option>
            <option {"selected" if d['grau']=="Liquidação" else ""}>Liquidação</option>
            <option {"selected" if d['grau']=="Aditamento" else ""}>Aditamento</option>
        </select><br><br>

        Critério:<br>
        <select name="criterio">
            <option {"selected" if d['criterio']=="Materialidade" else ""}>Materialidade</option>
            <option {"selected" if d['criterio']=="Relevância" else ""}>Relevância</option>
            <option {"selected" if d['criterio']=="Risco" else ""}>Risco</option>
            <option {"selected" if d['criterio']=="Engenharia" else ""}>Engenharia</option>
        </select><br><br>

        Data de conclusão (opcional):<br>
        <input type="date" name="data_fim"
               value="{d['data_fim'] or ''}"
               min="2026-01-01" max="2026-12-31"><br><br>

        <button class="btn">Salvar</button>
    </form>
    """

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        user=session["user"],
        perfil=session["perfil"]
    )

# ---------------------------------------------------------
# Visualizar uma delegação + histórico de horas associadas
# ---------------------------------------------------------
@app.route("/delegacao/<int:id>")
def ver_delegacao(id):
    if "user" not in session:
        return redirect("/")

    con = get_db()
    cur = con.cursor()

    # Buscar dados da delegação
    cur.execute("""
        SELECT d.*, c.nome AS colaborador
        FROM delegacoes d
        LEFT JOIN colaboradores c ON c.id = d.colaborador_id
        WHERE d.id = %s
    """, (id,))
    deleg = cur.fetchone()

    if not deleg:
        con.close()
        return "Delegação não encontrada"

    # -------------------------------------------------
    # REGRA DE ACESSO
    # -------------------------------------------------
    if session["perfil"] != "admin":
        if deleg["colaborador_id"] != session["user_id"]:
            con.close()
            return "Acesso negado"

    # Lista de requisições
    cur.execute("""
    SELECT h.*, col.nome AS colaborador_nome
    FROM horas h
    LEFT JOIN colaboradores col ON col.id = h.colaborador_id
    WHERE h.delegacao_id = %s
    ORDER BY h.data, h.hora_inicio
    """, (id,))

    horas = cur.fetchall()

    con.close()

    # -------------------------------
    # Construção do HTML
    # -------------------------------
    html = f"""
    <h2>Delegação #{id}</h2>

    <b>Requisições:</b> {deleg['requisicoes']}<br>
    <b>O.S:</b> {deleg['os_codigo']}<br>
    <b>Colaborador:</b> {deleg['colaborador']}<br>
    <b>Data Início:</b> {fmt(deleg['data_inicio'])}<br>
    <b>Status:</b> {deleg['status']}<br>
    <b>Grau:</b> {deleg['grau']}<br>
    <b>Critério:</b> {deleg['criterio']}<br><br>

    <h3>Horas Lançadas nesta Delegação</h3>
    """

    if not horas:
        html += "<p>Nenhum lançamento encontrado.</p>"
    else:
        html += """
        <table border=1 cellpadding=5>
            <tr>
                <th>Data</th>
                <th>Início</th>
                <th>Fim</th>
                <th>Duração</th>
                <th>Atividade</th>
                <th>Requisições</th>
                <th>Colaborador</th>
            </tr>
        """
        for h in horas:
            html += f"""
            <tr>
                <td>{fmt(h['data'])}</td>
                <td>{h['hora_inicio']}</td>
                <td>{h['hora_fim']}</td>
                <td>{h['duracao']}</td>
                <td>{h['atividade']}</td>
                <td>{deleg['requisicoes']}</td>
                <td>{h['colaborador_nome']}</td>
            </tr>
            """
        html += "</table>"

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        user=session["user"], perfil=session["perfil"]
    )

def minutos_para_hhmm(minutos):
    if minutos is None:
        return "00:00"
    horas = minutos // 60
    mins = minutos % 60
    return f"{horas:02d}:{mins:02d}"

# -------------------------
# Atendimentos (tela)
# -------------------------
@app.route('/atendimentos')
def atendimentos():
    if 'user' not in session:
        return redirect('/')

    mes = request.args.get('mes')  # YYYY-MM
    perfil = session.get("perfil")
    colaborador_id = session.get("user_id")
    nome_user = session.get("user")

    con = get_db()
    cur = con.cursor()

    where = []
    params = []

    # -------------------------
    # REGRA DE VISUALIZAÇÃO
    # -------------------------
    if perfil != "admin":
        where.append("""
            (
                h.colaborador_id = %s
                OR a.responsaveis_consultoria ILIKE %s
            )
        """)
        params.extend([colaborador_id, f"%{nome_user}%"])

    if mes:
        where.append("to_char(a.data_consultoria, 'YYYY-MM') = %s")
        params.append(mes)

    where_sql = "WHERE " + " AND ".join(where) if where else ""

    # -------------------------
    # CONSULTA
    # -------------------------
    sql = f"""
        SELECT
            a.id,
            a.data_consultoria,
            a.assunto,
            a.macro,
            a.diretoria,
            a.atividade,
            a.meio_contato,
            a.duracao_minutos,
            c.nome AS colaborador,
            h.os_codigo,
            h.item_paint
        FROM atendimentos a
        JOIN horas h ON h.id = a.hora_id
        JOIN colaboradores c ON c.id = h.colaborador_id
        {where_sql}
        ORDER BY a.data_consultoria DESC
        {"LIMIT 100" if not mes else ""}
    """

    cur.execute(sql, tuple(params))
    rows = cur.fetchall()
    con.close()

    # -------------------------
    # CONVERTER MINUTOS → HH:MM
    # -------------------------
    atendimentos = []
    for a in rows:
        a = dict(a)
        a["duracao_hhmm"] = minutos_para_hhmm(a["duracao_minutos"])
        a["data_consultoria"] = fmt(a["data_consultoria"])
        atendimentos.append(a)

    # -------------------------
    # HTML
    # -------------------------
    html = """
<h3>Atendimentos</h3>

<form method="get" style="margin-bottom:15px;">
    <label>Mês:</label>
    <input type="month" name="mes" value="{{ mes or '' }}">
    <button class="btn">Filtrar</button>
    <a href="/atendimentos">Limpar</a>
</form>

<a href="/atendimentos/exportar" class="btn" style="margin-bottom:10px; display:inline-block;">
    Exportar atendimentos
</a>

<table border="1" cellpadding="6" cellspacing="0" width="100%">
<tr>
    <th>Data</th>
    {% if perfil == 'admin' %}<th>Colaborador</th>{% endif %}
    <th>OS</th>
    <th>Item</th>
    <th>Assunto</th>
    <th>Macro</th>
    <th>Diretoria</th>
    <th>Meio</th>
    <th>Duração</th>
    <th>Ações</th>
</tr>

{% for a in atendimentos %}
<tr>
    <td>{{ a.data_consultoria }}</td>

    {% if perfil == 'admin' %}
        <td>{{ a.colaborador }}</td>
    {% endif %}

    <td>{{ a.os_codigo }}</td>
    <td>{{ a.item_paint }}</td>
    <td>{{ a.assunto }}</td>
    <td>{{ a.macro }}</td>
    <td>{{ a.diretoria }}</td>
    <td>{{ a.meio_contato }}</td>
    <td style="text-align:right;">{{ a.duracao_hhmm }}</td>
    <td>
        <a href="/atendimentos/ver/{{ a.id }}">Ver</a> |
        <a href="/atendimentos/editar/{{ a.id }}">Editar</a>
    </td>
</tr>
{% endfor %}
</table>
"""

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        atendimentos=atendimentos,
        mes=mes,
        perfil=perfil,
        user=session['user']
    )

@app.route('/atendimentos/ver/<int:id>')
def ver_atendimento(id):
    if 'user' not in session:
        return redirect('/')

    con = get_db()
    cur = con.cursor()

    cur.execute("""
        SELECT *
        FROM atendimentos
        WHERE id = %s
    """, (id,))

    a = cur.fetchone()
    a = dict(a)
    a["duracao_hhmm"] = minutos_para_hhmm(a["duracao_minutos"])
    a["data_consultoria"] = fmt(a["data_consultoria"])
    con.close()

    html = """
<h3>Detalhes do Atendimento</h3>

<p><b>Data:</b> {{ a.data_consultoria }}</p>
<p><b>Assunto:</b> {{ a.assunto }}</p>
<p><b>Macro:</b> {{ a.macro }}</p>
<p><b>Diretoria:</b> {{ a.diretoria }}</p>
<p><b>Atividade:</b> {{ a.atividade }}</p>
<p><b>Meio:</b> {{ a.meio_contato }}</p>
<p><b>Responsáveis:</b> {{ a.responsaveis_consultoria }}</p>
<p><b>Duração:</b> {{ a.duracao_hhmm }}</p>
<p><b>Observação:</b> {{ a.observacao }}</p>

<a href="/atendimentos">Voltar</a>
"""

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        a=a,
        user=session['user'],
        perfil=session['perfil']
    )

@app.route('/atendimentos/editar/<int:id>', methods=['GET', 'POST'])
def editar_atendimento(id):
    if 'user' not in session:
        return redirect('/')

    con = get_db()
    cur = con.cursor()

    if request.method == 'POST':
        cur.execute("""
            UPDATE atendimentos SET
                assunto = %s,
                macro = %s,
                diretoria = %s,
                atividade = %s,
                meio_contato = %s,
                observacao = %s,
                responsaveis_consultoria = %s
            WHERE id = %s
        """, (
            request.form['assunto'],
            request.form['macro'],
            request.form['diretoria'],
            request.form['atividade_atendimento'],
            request.form['meio_contato'],
            request.form['observacao'],
            request.form['responsaveis'],
            id
        ))
        con.commit()
        con.close()
        return redirect('/atendimentos')

    cur.execute("SELECT * FROM atendimentos WHERE id = %s", (id,))
    a = cur.fetchone()
    con.close()

    html = """
<h3>Editar Atendimento</h3>

<form method="post">
    <label>Assunto</label><br>
    <input name="assunto" value="{{ a.assunto }}"><br><br>

    <label>Macro</label><br>
    <input name="macro" value="{{ a.macro }}"><br><br>

    <label>Diretoria</label><br>
    <input name="diretoria" value="{{ a.diretoria }}"><br><br>

    <label>Atividade</label><br>
    <input name="atividade_atendimento" value="{{ a.atividade }}"><br><br>

    <label>Meio</label><br>
    <input name="meio_contato" value="{{ a.meio_contato }}"><br><br>

    <label>Responsáveis</label><br>
    <input name="responsaveis" value="{{ a.responsaveis_consultoria }}"><br><br>

    <label>Observação</label><br>
    <textarea name="observacao">{{ a.observacao }}</textarea><br><br>

    <button class="btn">Salvar</button>
    <a href="/atendimentos">Cancelar</a>
</form>
"""

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        a=a,
        user=session['user'],
        perfil=session["perfil"]
    )

# -------------------------
# Exportar Atendimentos (CSV)
# -------------------------
@app.route('/atendimentos/exportar')
def exportar_atendimentos():
    if 'user' not in session:
        return redirect('/')

    perfil = session.get("perfil")
    colaborador_id = session.get("user_id")

    con = get_db()
    cur = con.cursor()

    where = []
    params = []

    if perfil != "admin":
        where.append("h.colaborador_id = %s")
        params.append(colaborador_id)

    where_sql = "WHERE " + " AND ".join(where) if where else ""

    cur.execute(f"""
        SELECT
            a.data_consultoria,
            c.nome AS colaborador,
            h.os_codigo,
            h.item_paint,
            a.macro,
            a.diretoria,
            a.responsaveis_consultoria,
            a.atividade,
            a.data_consultoria,
            a.assunto,
            a.participantes_externos,
            a.entidades,
            a.meio_contato,
            a.observacao,
            a.duracao_minutos
        FROM atendimentos a
        JOIN horas h ON h.id = a.hora_id
        JOIN colaboradores c ON c.id = h.colaborador_id
        {where_sql}
        ORDER BY a.data_consultoria
    """, tuple(params))

    rows = cur.fetchall()
    con.close()

    import csv, io
    from flask import send_file

    output = io.StringIO()
    writer = csv.writer(output, delimiter=';')

    # ---- cabeçalho amigável
    writer.writerow([
        "Data",
        "Colaborador",
        "OS",
        "Item Paint",
        "Macro",
        "Diretoria",
        "Responsaveis",
        "Atividade",
        "Data",
        "Assunto",
        "Participantes",
        "Entidades",
        "Meio Contato",
        "Obs",
        "Duração (min)"
    ])

    # ---- dados CORRETOS (valores, não as chaves)
    for r in rows:
        writer.writerow([
            r["data_consultoria"],
            r["colaborador"],
            r["os_codigo"],
            r["item_paint"],
            r["macro"],
            r["diretoria"],
            r["responsaveis_consultoria"],
            r["atividade"],
            r["data_consultoria"],
            r["assunto"],
            r["participantes_externos"],
            r["entidades"],
            r["meio_contato"],
            r["observacao"],
            r["duracao_minutos"]
        ])

    output.seek(0)

    return send_file(
        io.BytesIO(output.getvalue().encode("utf-8-sig")),
        mimetype="text/csv",
        as_attachment=True,
        download_name="atendimentos.csv"
    )

@app.route('/consultorias')
def consultorias():
    if 'user' not in session:
        return redirect('/')

    mes = request.args.get('mes')  # YYYY-MM
    perfil = session.get("perfil")
    colaborador_id = session.get("user_id")
    nome_user = session.get("user")

    con = get_db()
    cur = con.cursor()

    where = []
    params = []

    # -------------------------
    # REGRA DE VISUALIZAÇÃO
    # -------------------------
    if perfil != "admin":
        where.append("""
            (
                h.colaborador_id = %s
                OR c.responsaveis ILIKE %s
            )
        """)
        params.extend([colaborador_id, f"%{nome_user}%"])

    if mes:
        where.append("to_char(c.data_consul, 'YYYY-MM') = %s")
        params.append(mes)

    where_sql = "WHERE " + " AND ".join(where) if where else ""

    sql = f"""
        SELECT
            c.id,
            c.data_consul,
            c.assunto,
            c.secretarias,
            c.meio,
            c.tipo,
            c.duracao_minutos,
            col.nome AS colaborador,
            h.os_codigo,
            h.item_paint
        FROM consultorias c
        JOIN horas h ON h.id = c.hora_id
        JOIN colaboradores col ON col.id = h.colaborador_id
        {where_sql}
        ORDER BY c.data_consul DESC
        {"LIMIT 100" if not mes else ""}
    """

    cur.execute(sql, tuple(params))
    rows = cur.fetchall()
    con.close()

    consultorias = []
    for c in rows:
        c = dict(c)
        c["duracao_hhmm"] = minutos_para_hhmm(c["duracao_minutos"])
        c["data_consul"] = fmt(c["data_consul"])
        consultorias.append(c)

    html = """
<h3>Consultorias</h3>

<form method="get" style="margin-bottom:15px;">
    <label>Mês:</label>
    <input type="month" name="mes" value="{{ mes or '' }}">
    <button class="btn">Filtrar</button>
    <a href="/consultorias">Limpar</a>
</form>

<a href="/consultorias/exportar" class="btn" style="margin-bottom:10px; display:inline-block;">
    Exportar consultorias
</a>

<table border="1" cellpadding="6" cellspacing="0" width="100%">
<tr>
    <th>Data</th>
    {% if perfil == 'admin' %}<th>Colaborador</th>{% endif %}
    <th>OS</th>
    <th>Item</th>
    <th>Assunto</th>
    <th>Secretarias</th>
    <th>Meio</th>
    <th>Tipo</th>
    <th>Duração</th>
    <th>Ações</th>
</tr>

{% for c in consultorias %}
<tr>
    <td>{{ c.data_consul }}</td>

    {% if perfil == 'admin' %}
        <td>{{ c.colaborador }}</td>
    {% endif %}

    <td>{{ c.os_codigo }}</td>
    <td>{{ c.item_paint }}</td>
    <td>{{ c.assunto }}</td>
    <td>{{ c.secretarias }}</td>
    <td>{{ c.meio }}</td>
    <td>{{ c.tipo }}</td>
    <td style="text-align:right;">{{ c.duracao_hhmm }}</td>
    <td>
        <a href="/consultorias/ver/{{ c.id }}">Ver</a> |
        <a href="/consultorias/editar/{{ c.id }}">Editar</a>
    </td>
</tr>
{% endfor %}
</table>
"""

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        consultorias=consultorias,
        mes=mes,
        perfil=perfil,
        user=session['user']
    )

@app.route('/consultorias/ver/<int:id>')
def ver_consultoria(id):
    if 'user' not in session:
        return redirect('/')

    con = get_db()
    cur = con.cursor()

    cur.execute("SELECT * FROM consultorias WHERE id=%s", (id,))
    c = dict(cur.fetchone())
    con.close()

    c["duracao_hhmm"] = minutos_para_hhmm(c["duracao_minutos"])
    c["data_consul"] = fmt(c["data_consul"])

    html = """
<h3>Detalhes da Consultoria / Treinamento</h3>
<p><b>Tipo:</b> {{ c.tipo }}</p>
<p><b>Data:</b> {{ c.data_consul }}</p>
<p><b>Assunto:</b> {{ c.assunto }}</p>
<p><b>Secretarias:</b> {{ c.secretarias }}</p>
<p><b>Meio:</b> {{ c.meio }}</p>
<p><b>Responsáveis:</b> {{ c.responsaveis }}</p>
<p><b>Palavras-chave:</b> {{ c.palavras_chave }}</p>
<p><b>Nº Ofício:</b> {{ c.num_oficio }}</p>
<p><b>Duração:</b> {{ c.duracao_hhmm }}</p>
<p><b>Observação:</b> {{ c.observacao }}</p>

<a href="/consultorias">Voltar</a>
"""

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        c=c,
        user=session['user'],
        perfil=session['perfil']
    )

@app.route('/consultorias/editar/<int:id>', methods=['GET', 'POST'])
def editar_consultoria(id):
    if 'user' not in session:
        return redirect('/')

    con = get_db()
    cur = con.cursor()

    if request.method == 'POST':
        cur.execute("""
            UPDATE consultorias SET
                assunto = %s,
                secretarias = %s,
                meio = %s,
                tipo = %s,
                responsaveis = %s,
                palavras_chave = %s,
                num_oficio = %s,
                observacao = %s
            WHERE id = %s
        """, (
            request.form['assunto'],
            request.form['secretarias'],
            request.form['meio'],
            request.form['tipo'],
            request.form['responsaveis'],
            request.form['palavras_chave'],
            request.form['num_oficio'],
            request.form['observacao'],
            id
        ))
        con.commit()
        con.close()
        return redirect('/consultorias')

    cur.execute("SELECT * FROM consultorias WHERE id=%s", (id,))
    c = cur.fetchone()
    con.close()

    html = """
<h3>Editar Consultoria</h3>

<form method="post">
    <label>Assunto</label><br>
    <input name="assunto" value="{{ c.assunto }}"><br><br>

    <label>Secretarias</label><br>
    <input name="secretarias" value="{{ c.secretarias }}"><br><br>

    <label>Meio</label><br>
    <input name="meio" value="{{ c.meio }}"><br><br>

    <label>Tipo</label><br>
    <input name="tipo" value="{{ c.tipo }}"><br><br>

    <label>Responsáveis</label><br>
    <input name="responsaveis" value="{{ c.responsaveis }}"><br><br>

    <label>Palavras-chave</label><br>
    <input name="palavras_chave" value="{{ c.palavras_chave }}"><br><br>

    <label>Nº Ofício</label><br>
    <input name="num_oficio" value="{{ c.num_oficio }}"><br><br>

    <label>Observação</label><br>
    <textarea name="observacao">{{ c.observacao }}</textarea><br><br>

    <button class="btn">Salvar</button>
    <a href="/consultorias">Cancelar</a>
</form>
"""

    return render_template_string(
        BASE.replace("{% block content %}{% endblock %}", html),
        c=c,
        user=session['user'],
        perfil=session["perfil"]
    )

@app.route('/consultorias/exportar')
def exportar_consultorias():
    if 'user' not in session:
        return redirect('/')

    perfil = session.get("perfil")
    colaborador_id = session.get("user_id")

    con = get_db()
    cur = con.cursor()

    where = []
    params = []

    if perfil != "admin":
        where.append("h.colaborador_id = %s")
        params.append(colaborador_id)

    where_sql = "WHERE " + " AND ".join(where) if where else ""

    cur.execute(f"""
        SELECT
            c.data_consul,
            col.nome AS colaborador,
            h.os_codigo,
            h.item_paint,
            c.secretarias,
            c.responsaveis,
            c.assunto,
            c.meio,
            c.tipo,
            c.palavras_chave,
            c.num_oficio,
            c.observacao,
            c.duracao_minutos
        FROM consultorias c
        JOIN horas h ON h.id = c.hora_id
        JOIN colaboradores col ON col.id = h.colaborador_id
        {where_sql}
        ORDER BY c.data_consul
    """, tuple(params))

    rows = cur.fetchall()
    con.close()

    import csv, io
    from flask import send_file

    output = io.StringIO()
    writer = csv.writer(output, delimiter=';')

    writer.writerow([
        "Data",
        "Colaborador",
        "OS",
        "Item Paint",
        "Secretarias",
        "Responsáveis",
        "Assunto",
        "Meio",
        "Tipo",
        "Palavras-chave",
        "Nº Ofício",
        "Observação",
        "Duração (min)"
    ])

    for r in rows:
        writer.writerow([
            r["data_consul"],
            r["colaborador"],
            r["os_codigo"],
            r["item_paint"],
            r["secretarias"],
            r["responsaveis"],
            r["assunto"],
            r["meio"],
            r["tipo"],
            r["palavras_chave"],
            r["num_oficio"],
            r["observacao"],
            r["duracao_minutos"]
        ])

    output.seek(0)

    return send_file(
        io.BytesIO(output.getvalue().encode("utf-8-sig")),
        mimetype="text/csv",
        as_attachment=True,
        download_name="consultorias.csv"
    )

@app.route("/seed")
def seed():
    executar_seed()
    return "Seed executado com sucesso!"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
    
