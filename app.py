# -*- coding: utf-8 -*-
import os
import io
import time
import json
import sqlite3
from datetime import datetime
from flask import (
    Flask, request, jsonify, send_file, render_template_string
)
import pypdf
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, HRFlowable
)
from PIL import Image

from seed_data import (
    DEMANDAS_FIXAS_PADRAO, ITENS_PRASS_INICIAIS, CHECKLISTS_INICIAIS_COMPLEMENTO
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # Limite de 64MB para uploads


# ==============================================================================
# CAMADA DE BANCO DE DADOS (AUTO-DETECÇÃO POSTGRESQL RAILWAY + FALLBACK SQLITE)
# ==============================================================================
def obter_url_postgres():
    """Localiza automaticamente a URL do PostgreSQL no Railway em qualquer variável disponível."""
    # 1. Tenta encontrar uma URL completa já definida
    for nome_var in ("DATABASE_PUBLIC_URL", "DATABASE_URL", "POSTGRES_URL", "POSTGRESQL_URL"):
        val = os.environ.get(nome_var, "").strip()
        # Ignora se o Railway não resolveu a variável (ainda contém '${{')
        if val and not val.startswith("${{"):
            if val.startswith("postgres://"):
                val = val.replace("postgres://", "postgresql://", 1)
            return val

    # 2. Constrói a URL a partir das variáveis individuais (método mais seguro no Railway)
    pghost = os.environ.get("PGHOST", "").strip()
    pguser = os.environ.get("PGUSER", "").strip()
    pgpass = os.environ.get("PGPASSWORD", "").strip()
    pgdb = os.environ.get("PGDATABASE", "").strip()
    pgport = os.environ.get("PGPORT", "5432").strip()
    
    if pghost and pguser and pgdb:
        return f"postgresql://{pguser}:{pgpass}@{pghost}:{pgport}/{pgdb}"

    return ""


DATABASE_URL = obter_url_postgres()
USE_POSTGRES = bool(DATABASE_URL)
DB_INITIALIZED = False

if USE_POSTGRES:
    import psycopg2
    from psycopg2.extras import RealDictCursor


def get_db_conn():
    if USE_POSTGRES:
        ultimo_erro = None
        for _ in range(5):
            try:
                return psycopg2.connect(DATABASE_URL, connect_timeout=10)
            except Exception as e:
                ultimo_erro = e
                time.sleep(1.2)
        raise ultimo_erro
    conn = sqlite3.connect("prass_local.db")
    conn.row_factory = sqlite3.Row
    return conn


def db_query(sql, params=(), fetch=True, commit=False):
    """Executa queries no banco de forma agnóstica (PostgreSQL ou SQLite)"""
    conn = get_db_conn()
    try:
        if USE_POSTGRES:
            cur = conn.cursor(cursor_factory=RealDictCursor)
            cur.execute(sql, params)
        else:
            sql_sqlite = sql.replace("%s", "?")
            cur = conn.cursor()
            cur.execute(sql_sqlite, params)

        res = None
        if fetch:
            rows = cur.fetchall()
            res = [dict(r) for r in rows]
        if commit:
            conn.commit()
        return res
    finally:
        conn.close()


def init_db():
    conn = get_db_conn()
    try:
        cur = conn.cursor()
        if USE_POSTGRES:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS prass_items (
                    codigo VARCHAR(20) PRIMARY KEY,
                    titulo TEXT NOT NULL,
                    apontamento_ans TEXT NOT NULL,
                    medidas_implementadas TEXT NOT NULL,
                    observacoes TEXT NOT NULL,
                    historico_json TEXT NOT NULL,
                    areas_atribuidas TEXT NOT NULL,
                    oficio_rascunho TEXT DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS demandas (
                    id SERIAL PRIMARY KEY,
                    item_codigo VARCHAR(20) NOT NULL,
                    nome VARCHAR(180) NOT NULL,
                    tipo VARCHAR(20) NOT NULL,
                    ordem INT DEFAULT 100,
                    oficio_texto TEXT DEFAULT '',
                    UNIQUE(item_codigo, nome)
                );
                CREATE TABLE IF NOT EXISTS checklist_evidencias (
                    id SERIAL PRIMARY KEY,
                    demanda_id INT NOT NULL REFERENCES demandas(id) ON DELETE CASCADE,
                    setor VARCHAR(20) NOT NULL,
                    titulo TEXT NOT NULL,
                    analise_evidencia TEXT DEFAULT '',
                    concluido BOOLEAN DEFAULT FALSE,
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS anexos (
                    id SERIAL PRIMARY KEY,
                    checklist_id INT NOT NULL REFERENCES checklist_evidencias(id) ON DELETE CASCADE,
                    nome_arquivo VARCHAR(255) NOT NULL,
                    mime_type VARCHAR(120) NOT NULL,
                    tamanho INT NOT NULL,
                    resumo_extraido TEXT DEFAULT '',
                    conteudo BYTEA NOT NULL,
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
            """)
        else:
            cur.executescript("""
                CREATE TABLE IF NOT EXISTS prass_items (
                    codigo TEXT PRIMARY KEY,
                    titulo TEXT NOT NULL,
                    apontamento_ans TEXT NOT NULL,
                    medidas_implementadas TEXT NOT NULL,
                    observacoes TEXT NOT NULL,
                    historico_json TEXT NOT NULL,
                    areas_atribuidas TEXT NOT NULL,
                    oficio_rascunho TEXT DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS demandas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    item_codigo TEXT NOT NULL,
                    nome TEXT NOT NULL,
                    tipo TEXT NOT NULL,
                    ordem INTEGER DEFAULT 100,
                    oficio_texto TEXT DEFAULT '',
                    UNIQUE(item_codigo, nome)
                );
                CREATE TABLE IF NOT EXISTS checklist_evidencias (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    demanda_id INTEGER NOT NULL,
                    setor TEXT NOT NULL,
                    titulo TEXT NOT NULL,
                    analise_evidencia TEXT DEFAULT '',
                    concluido INTEGER DEFAULT 0,
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS anexos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    checklist_id INTEGER NOT NULL,
                    nome_arquivo TEXT NOT NULL,
                    mime_type TEXT NOT NULL,
                    tamanho INTEGER NOT NULL,
                    resumo_extraido TEXT DEFAULT '',
                    conteudo BLOB NOT NULL,
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
            """)
        conn.commit()
    finally:
        conn.close()

    # Popula os itens do PRASS e as demandas fixas (08/2026 a 08/2027) se ainda não existirem
    existentes = db_query("SELECT codigo FROM prass_items")
    codigos_existentes = {r["codigo"] for r in existentes}

    for item in ITENS_PRASS_INICIAIS:
        cod = item["codigo"]
        if cod not in codigos_existentes:
            db_query(
                """INSERT INTO prass_items
                   (codigo, titulo, apontamento_ans, medidas_implementadas, observacoes, historico_json, areas_atribuidas)
                   VALUES (%s, %s, %s, %s, %s, %s, %s)""",
                (
                    cod,
                    item["titulo"],
                    item["apontamento_ans"],
                    item["medidas_implementadas"],
                    item["observacoes"],
                    json.dumps(item.get("historico_planilha", {}), ensure_ascii=False),
                    json.dumps(item.get("areas_atribuidas", ["CODER", "COCAP", "CONEC"]), ensure_ascii=False),
                ),
                fetch=False,
                commit=True,
            )

        for idx_d, nome_dem in enumerate(DEMANDAS_FIXAS_PADRAO, start=1):
            ja_tem = db_query(
                "SELECT id FROM demandas WHERE item_codigo = %s AND nome = %s",
                (cod, nome_dem)
            )
            if not ja_tem:
                db_query(
                    "INSERT INTO demandas (item_codigo, nome, tipo, ordem) VALUES (%s, %s, %s, %s)",
                    (cod, nome_dem, "FIXA", idx_d),
                    fetch=False,
                    commit=True
                )

        if cod in CHECKLISTS_INICIAIS_COMPLEMENTO:
            dem_row = db_query(
                "SELECT id FROM demandas WHERE item_codigo = %s AND nome = %s",
                (cod, "Complemento ANS 26/10")
            )
            if dem_row:
                dem_id = dem_row[0]["id"]
                exist_chk = db_query(
                    "SELECT COUNT(*) AS qtd FROM checklist_evidencias WHERE demanda_id = %s",
                    (dem_id,)
                )
                if exist_chk and int(exist_chk[0]["qtd"]) == 0:
                    for setor, lista_titulos in CHECKLISTS_INICIAIS_COMPLEMENTO[cod].items():
                        for tit in lista_titulos:
                            db_query(
                                """INSERT INTO checklist_evidencias (demanda_id, setor, titulo, concluido)
                                   VALUES (%s, %s, %s, %s)""",
                                (dem_id, setor, tit, False),
                                fetch=False,
                                commit=True
                            )


@app.before_request
def garantir_banco_inicializado():
    global DB_INITIALIZED
    if not DB_INITIALIZED:
        try:
            init_db()
            DB_INITIALIZED = True
        except Exception as e:
            print(f"Aviso ao inicializar banco: {e}")


# ==============================================================================
# FUNÇÕES AUXILIARES DE ANÁLISE DE ANEXOS E GERAÇÃO DE PDF UNIFICADO
# ==============================================================================
def extrair_resumo_arquivo(nome_arquivo, raw_bytes):
    ext = os.path.splitext(nome_arquivo.lower())[1]
    try:
        if ext == ".pdf":
            reader = pypdf.PdfReader(io.BytesIO(raw_bytes))
            num_pags = len(reader.pages)
            trechos = []
            for p in reader.pages[:3]:
                txt = (p.extract_text() or "").strip()
                if txt:
                    trechos.append(" ".join(txt.split())[:350])
            resumo_txt = " | ".join(trechos)[:600] if trechos else "Documento PDF composto por imagens/capturas de tela."
            return f"Documento PDF com {num_pags} página(s). Síntese extraída: {resumo_txt}"
        elif ext in (".xlsx", ".xls", ".csv"):
            if ext == ".csv":
                linhas = raw_bytes.decode("latin-1", errors="ignore").splitlines()
                cabecalho = linhas[0][:250] if linhas else ""
                return (
                    f"Planilha/Base de Dados em formato {ext.upper()} (referenciada no Ofício pelo nome '{nome_arquivo}'), "
                    f"contendo {max(0, len(linhas)-1)} registros. Colunas identificadas: {cabecalho}"
                )
            return (
                f"Planilha eletrônica Excel ({ext.upper()}) mantida como arquivo externo referenciado sob o nome "
                f"'{nome_arquivo}' ({len(raw_bytes)/1024:.1f} KB)."
            )
        elif ext in (".png", ".jpg", ".jpeg"):
            im = Image.open(io.BytesIO(raw_bytes))
            return f"Captura de tela / Evidência visual ({ext.upper()}, resolução {im.size[0]}x{im.size[1]}px)."
        else:
            return f"Documento comprobatório ({ext.upper()}, {len(raw_bytes)/1024:.1f} KB)."
    except Exception:
        return f"Arquivo comprobatório anexado ({nome_arquivo})."


def calcular_semaforo_setores(demanda_id, areas_atribuidas):
    rows = db_query(
        """SELECT c.id, c.setor, c.concluido, COUNT(a.id) AS qtd_anexos
           FROM checklist_evidencias c
           LEFT JOIN anexos a ON a.checklist_id = c.id
           WHERE c.demanda_id = %s
           GROUP BY c.id, c.setor, c.concluido""",
        (demanda_id,)
    )
    por_setor = {"CODER": [], "COCAP": [], "CONEC": []}
    for r in rows:
        s = r["setor"]
        if s in por_setor:
            por_setor[s].append(r)

    resultado = {}
    for setor in ["CODER", "COCAP", "CONEC"]:
        itens = por_setor[setor]
        atribuido = setor in areas_atribuidas
        total = len(itens)
        concluidos = sum(1 for x in itens if bool(x["concluido"]))
        com_anexo = sum(1 for x in itens if int(x["qtd_anexos"]) > 0)

        if total == 0:
            cor = "VERMELHO"
            label = "Não Respondido"
        elif concluidos == total and total > 0:
            cor = "VERDE"
            label = f"Respondido ({concluidos}/{total})"
        elif concluidos > 0 or com_anexo > 0:
            cor = "AMARELO"
            label = f"Respondendo ({concluidos}/{total})"
        else:
            cor = "VERMELHO"
            label = f"Não Respondido (0/{total})"

        resultado[setor] = {
            "atribuido": atribuido,
            "cor": cor,
            "label": label,
            "total": total,
            "concluidos": concluidos,
            "percentual": round((concluidos / total * 100), 1) if total > 0 else 0.0
        }
    return resultado


def construir_texto_oficio_automatico(item_row, demanda_row, checklists_com_anexos):
    data_str = datetime.now().strftime("%d/%m/%Y")
    linhas = [
        f"Brasília, {data_str}",
        "",
        "RESPOSTA AO OFÍCIO DA ANS — MONITORAMENTO PRASS POSTAL SAÚDE",
        f"ITEM DO PRASS: {item_row['titulo']}",
        f"REFERÊNCIA / DEMANDA: {demanda_row['nome']} ({demanda_row['tipo']})",
        f"COORDENAÇÕES ATRIBUÍDAS (GERED): {', '.join(json.loads(item_row['areas_atribuidas']))}",
        "",
        "1. CONTEXTUALIZAÇÃO E DIAGNÓSTICO INSTITUCIONAL",
        "Em atendimento às determinações da Agência Nacional de Saúde Suplementar (ANS) no âmbito do Plano de "
        "Recuperação Assistencial (PRASS - Processo ANS nº 33910.030426/2026-21), a Gerência Executiva de Rede (GERED) "
        "apresenta os esclarecimentos técnicos, a origem das extrações sistêmicas e a documentação comprobatória "
        f"referente ao item {item_row['codigo']}, consolidada pelas coordenações técnicas vinculadas.",
        ""
    ]

    hist = json.loads(item_row.get("historico_json") or "{}")
    if demanda_row["nome"] in hist and hist[demanda_row["nome"]].strip() and hist[demanda_row["nome"]].strip() != "N/A":
        linhas.append("2. CONSIDERAÇÕES EXECUTIVAS DO PERÍODO")
        linhas.append(hist[demanda_row["nome"]].strip())
        linhas.append("")

    linhas.append("3. ANÁLISE TÉCNICA DAS EVIDÊNCIAS E DOCUMENTOS COMPROBATÓRIOS")
    linhas.append("")

    if not checklists_com_anexos:
        linhas.append("Nenhum item de checklist cadastrado para esta demanda até o momento.")
    else:
        for idx, chk in enumerate(checklists_com_anexos, start=1):
            status_txt = "CONCLUÍDO" if chk["concluido"] else "EM ANDAMENTO / PENDENTE DE COMPLEMENTAÇÃO"
            linhas.append(f"3.{idx}. [{chk['setor']}] {chk['titulo']}")

            anexos = chk.get("anexos", [])
            if anexos:
                nomes_arqs = []
                notas_excel = []
                resumos_arqs = []
                for a in anexos:
                    n_arq = a["nome_arquivo"]
                    nomes_arqs.append(n_arq)
                    ext = os.path.splitext(n_arq.lower())[1]
                    if ext in (".xlsx", ".xls", ".csv"):
                        notas_excel.append(
                            f"• Planilha/Base Externa Referenciada: '{n_arq}' (mantida em formato original de planilha "
                            f"eletrônica e não mesclada ao corpo do PDF para preservação das linhas e colunas auditáveis)."
                        )
                    else:
                        resumos_arqs.append(f"• {n_arq}: {a.get('resumo_extraido', 'Anexado integralmente ao dossiê PDF.')}")

                linhas.append(f"Evidência(s) apresentada(s): {', '.join(nomes_arqs)}")
                linhas.append("")
                if chk.get("analise_evidencia") and chk["analise_evidencia"].strip():
                    linhas.append(chk["analise_evidencia"].strip())
                else:
                    linhas.append(
                        f"A coordenação {chk['setor']}/GERED disponibiliza a documentação comprobatória referente ao item "
                        f"'{chk['titulo']}', extraída dos sistemas corporativos da Postal Saúde (BENNER / GLPI / Portal Guia de Saúde) "
                        f"e validada conforme os procedimentos operacionais vigentes."
                    )
                if resumos_arqs:
                    linhas.append("Detalhamento dos documentos incorporados ao PDF consolidado:")
                    linhas.extend(resumos_arqs)
                if notas_excel:
                    linhas.append("Referência de Bases em Excel / CSV:")
                    linhas.extend(notas_excel)
            else:
                linhas.append("Evidência(s) apresentada(s): (Sem anexo vinculado até o momento)")
                if chk.get("analise_evidencia") and chk["analise_evidencia"].strip():
                    linhas.append(chk["analise_evidencia"].strip())
                else:
                    linhas.append(
                        f"Item registrado sob responsabilidade da coordenação {chk['setor']}, aguardando juntada do "
                        f"documento primário comprobatório."
                    )

            linhas.append(f"Situação: {status_txt}.")
            linhas.append("-" * 78)
            linhas.append("")

    linhas.extend([
        "",
        "Atenciosamente,",
        "",
        "João Victor Peres de Moura",
        "Gerente Executivo de Rede",
        "CODER / COCAP / CONEC — GERED / POSTAL SAÚDE"
    ])
    return "\n".join(linhas)


def gerar_pdf_capa_oficio(titulo_item, nome_demanda, texto_oficio, lista_arquivos_excel, lista_pdfs_anexados):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=2.2 * cm,
        rightMargin=2.2 * cm,
        topMargin=2.0 * cm,
        bottomMargin=2.0 * cm
    )

    styles = getSampleStyleSheet()
    st_header = ParagraphStyle(
        "HeaderPostal",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#003366")
    )
    st_sub = ParagraphStyle(
        "SubPostal",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#475569")
    )
    st_secao = ParagraphStyle(
        "SecaoOficio",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10.5,
        leading=14,
        spaceBefore=8,
        spaceAfter=4,
        textColor=colors.HexColor("#003366")
    )
    st_corpo = ParagraphStyle(
        "CorpoOficio",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9.8,
        leading=13.5,
        spaceBefore=2,
        spaceAfter=5,
        textColor=colors.HexColor("#1E293B")
    )

    story = []
    story.append(Paragraph("POSTAL SAÚDE — CAIXA DE ASSISTÊNCIA E SAÚDE DOS EMPREGADOS DOS CORREIOS", st_header))
    story.append(Paragraph("PRASS — Plano de Recuperação Assistencial | Processo ANS nº 33910.030426/2026-21 | GERED", st_sub))
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#003366")))
    story.append(Spacer(1, 10))

    for linha in texto_oficio.splitlines():
        l_strip = linha.strip()
        if not l_strip:
            story.append(Spacer(1, 6))
            continue
        if l_strip.startswith("---"):
            story.append(Spacer(1, 4))
            story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#CBD5E1")))
            story.append(Spacer(1, 4))
            continue

        l_safe = (
            l_strip.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

        if (
            l_strip.startswith("RESPOSTA AO OFÍCIO")
            or l_strip.startswith("ITEM DO PRASS:")
            or l_strip.startswith("1.")
            or l_strip.startswith("2.")
            or l_strip.startswith("3.")
            or l_strip.startswith("Evidência(s)")
            or l_strip.startswith("Situação:")
        ):
            story.append(Paragraph(l_safe, st_secao))
        else:
            story.append(Paragraph(l_safe, st_corpo))

    story.append(Spacer(1, 14))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#003366")))
    story.append(Spacer(1, 6))
    story.append(Paragraph("ÍNDICE DE ANEXOS CONSOLIDADOS NESTE DOSSIÊ", st_secao))

    if lista_pdfs_anexados:
        story.append(Paragraph("<b>A) Documentos PDF / Imagens unificados a seguir neste próprio arquivo PDF:</b>", st_corpo))
        for idx_p, info_p in enumerate(lista_pdfs_anexados, start=1):
            story.append(Paragraph(f"{idx_p}. [{info_p['setor']}] {info_p['nome_arquivo']} — ({info_p['titulo_chk']})", st_corpo))

    if lista_arquivos_excel:
        story.append(Spacer(1, 6))
        story.append(Paragraph("<b>B) Arquivos em formato Planilha Eletrônica (Excel / CSV) referenciados pelo nome original:</b>", st_corpo))
        for idx_x, info_x in enumerate(lista_arquivos_excel, start=1):
            story.append(Paragraph(
                f"{idx_x}. <b>{info_x['nome_arquivo']}</b> — Setor: {info_x['setor']} | Item: {info_x['titulo_chk']}",
                st_corpo
            ))

    doc.build(story)
    buffer.seek(0)
    return buffer


def converter_imagem_para_pdf_bytes(img_bytes):
    im = Image.open(io.BytesIO(img_bytes))
    if im.mode in ("RGBA", "P"):
        im = im.convert("RGB")
    out = io.BytesIO()
    im.save(out, format="PDF", resolution=100.0)
    out.seek(0)
    return out


# ==============================================================================
# ROTAS DE API (BACKEND FULL PYTHON)
# ==============================================================================
@app.route("/api/bootstrap", methods=["GET"])
def api_bootstrap():
    itens = db_query("SELECT * FROM prass_items ORDER BY codigo")
    demanda_selecionada_nome = request.args.get("demanda_nome", "Complemento ANS 26/10")

    resultado_itens = []
    for it in itens:
        cod = it["codigo"]
        demandas = db_query(
            "SELECT * FROM demandas WHERE item_codigo = %s ORDER BY ordem ASC, id ASC",
            (cod,)
        )
        dem_ativa = next((d for d in demandas if d["nome"] == demanda_selecionada_nome), None)
        if not dem_ativa and demandas:
            dem_ativa = demandas[0]

        areas = json.loads(it["areas_atribuidas"] or '["CODER","COCAP","CONEC"]')
        semaforo = calcular_semaforo_setores(dem_ativa["id"], areas) if dem_ativa else {}

        resultado_itens.append({
            "codigo": cod,
            "titulo": it["titulo"],
            "apontamento_ans": it["apontamento_ans"],
            "medidas_implementadas": it["medidas_implementadas"],
            "observacoes": it["observacoes"],
            "historico_planilha": json.loads(it["historico_json"] or "{}"),
            "areas_atribuidas": areas,
            "demandas": demandas,
            "semaforo": semaforo
        })

    return jsonify({"itens": resultado_itens, "postgres_ativo": USE_POSTGRES})


@app.route("/api/item/<path:codigo>/areas", methods=["POST"])
def api_atualizar_areas(codigo):
    body = request.get_json(force=True)
    areas = body.get("areas_atribuidas", ["CODER", "COCAP", "CONEC"])
    db_query(
        "UPDATE prass_items SET areas_atribuidas = %s WHERE codigo = %s",
        (json.dumps(areas, ensure_ascii=False), codigo),
        fetch=False,
        commit=True
    )
    return jsonify({"ok": True, "areas_atribuidas": areas})


@app.route("/api/demandas", methods=["POST"])
def api_criar_demanda():
    body = request.get_json(force=True)
    item_codigo = body.get("item_codigo", "").strip()
    nome = body.get("nome", "").strip()
    replicar_todos = bool(body.get("replicar_todos", True))

    if not item_codigo or not nome:
        return jsonify({"erro": "Informe o código do item e o nome da demanda."}), 400

    codigos_alvo = [r["codigo"] for r in db_query("SELECT codigo FROM prass_items")] if replicar_todos else [item_codigo]
    for c in codigos_alvo:
        existe = db_query("SELECT id FROM demandas WHERE item_codigo = %s AND nome = %s", (c, nome))
        if not existe:
            db_query(
                "INSERT INTO demandas (item_codigo, nome, tipo, ordem) VALUES (%s, %s, %s, %s)",
                (c, nome, "AVULSA", 500),
                fetch=False,
                commit=True
            )
    return jsonify({"ok": True, "nome": nome})


@app.route("/api/checklist/<int:demanda_id>", methods=["GET"])
def api_listar_checklist(demanda_id):
    dem = db_query("SELECT * FROM demandas WHERE id = %s", (demanda_id,))
    if not dem:
        return jsonify({"erro": "Demanda não encontrada."}), 404
    dem_row = dem[0]
    it_row = db_query("SELECT * FROM prass_items WHERE codigo = %s", (dem_row["item_codigo"],))[0]

    chks = db_query(
        "SELECT * FROM checklist_evidencias WHERE demanda_id = %s ORDER BY id ASC",
        (demanda_id,)
    )
    for c in chks:
        c["concluido"] = bool(c["concluido"])
        anexos = db_query(
            """SELECT id, checklist_id, nome_arquivo, mime_type, tamanho, resumo_extraido, criado_em
               FROM anexos WHERE checklist_id = %s ORDER BY id ASC""",
            (c["id"],)
        )
        for a in anexos:
            a["criado_em"] = str(a["criado_em"])
        c["anexos"] = anexos
        c["criado_em"] = str(c["criado_em"])

    areas = json.loads(it_row["areas_atribuidas"] or '["CODER","COCAP","CONEC"]')
    semaforo = calcular_semaforo_setores(demanda_id, areas)

    return jsonify({
        "demanda": dem_row,
        "checklist": chks,
        "semaforo": semaforo
    })


@app.route("/api/checklist", methods=["POST"])
def api_adicionar_checklist():
    body = request.get_json(force=True)
    demanda_id = int(body["demanda_id"])
    setor = body["setor"].strip().upper()
    titulo = body["titulo"].strip()
    analise = body.get("analise_evidencia", "").strip()

    if not titulo:
        return jsonify({"erro": "Digite a descrição da evidência para o checklist."}), 400

    db_query(
        """INSERT INTO checklist_evidencias (demanda_id, setor, titulo, analise_evidencia, concluido)
           VALUES (%s, %s, %s, %s, %s)""",
        (demanda_id, setor, titulo, analise, False),
        fetch=False,
        commit=True
    )
    return jsonify({"ok": True})


@app.route("/api/checklist/<int:chk_id>", methods=["PUT"])
def api_editar_checklist(chk_id):
    body = request.get_json(force=True)
    if "titulo" in body:
        db_query(
            "UPDATE checklist_evidencias SET titulo = %s, analise_evidencia = %s WHERE id = %s",
            (body["titulo"].strip(), body.get("analise_evidencia", "").strip(), chk_id),
            fetch=False,
            commit=True
        )
    if "concluido" in body:
        db_query(
            "UPDATE checklist_evidencias SET concluido = %s WHERE id = %s",
            (bool(body["concluido"]), chk_id),
            fetch=False,
            commit=True
        )
    return jsonify({"ok": True})


@app.route("/api/checklist/<int:chk_id>", methods=["DELETE"])
def api_excluir_checklist(chk_id):
    db_query("DELETE FROM anexos WHERE checklist_id = %s", (chk_id,), fetch=False, commit=True)
    db_query("DELETE FROM checklist_evidencias WHERE id = %s", (chk_id,), fetch=False, commit=True)
    return jsonify({"ok": True})


@app.route("/api/checklist/<int:chk_id>/anexos", methods=["POST"])
def api_upload_anexos(chk_id):
    arquivos = request.files.getlist("arquivos")
    if not arquivos:
        return jsonify({"erro": "Nenhum arquivo enviado."}), 400

    conn = get_db_conn()
    try:
        cur = conn.cursor()
        for f in arquivos:
            if not f.filename:
                continue
            raw_bytes = f.read()
            nome_arq = f.filename
            mime = f.mimetype or "application/octet-stream"
            tamanho = len(raw_bytes)
            resumo = extrair_resumo_arquivo(nome_arq, raw_bytes)

            if USE_POSTGRES:
                cur.execute(
                    """INSERT INTO anexos (checklist_id, nome_arquivo, mime_type, tamanho, resumo_extraido, conteudo)
                       VALUES (%s, %s, %s, %s, %s, %s)""",
                    (chk_id, nome_arq, mime, tamanho, resumo, psycopg2.Binary(raw_bytes))
                )
            else:
                cur.execute(
                    """INSERT INTO anexos (checklist_id, nome_arquivo, mime_type, tamanho, resumo_extraido, conteudo)
                       VALUES (?, ?, ?, ?, ?, ?)""",
                    (chk_id, nome_arq, mime, tamanho, resumo, sqlite3.Binary(raw_bytes))
                )
        conn.commit()
    finally:
        conn.close()

    return jsonify({"ok": True})


@app.route("/api/anexos/<int:anexo_id>", methods=["DELETE"])
def api_remover_anexo(anexo_id):
    row = db_query("SELECT checklist_id FROM anexos WHERE id = %s", (anexo_id,))
    if row:
        chk_id = row[0]["checklist_id"]
        db_query("DELETE FROM anexos WHERE id = %s", (anexo_id,), fetch=False, commit=True)
        db_query(
            "UPDATE checklist_evidencias SET concluido = %s WHERE id = %s",
            (False, chk_id),
            fetch=False,
            commit=True
        )
    return jsonify({"ok": True})


@app.route("/api/anexos/<int:anexo_id>/view", methods=["GET"])
def api_visualizar_anexo(anexo_id):
    row = db_query("SELECT nome_arquivo, mime_type, conteudo FROM anexos WHERE id = %s", (anexo_id,))
    if not row:
        return "Arquivo não encontrado", 404
    r = row[0]
    dados = bytes(r["conteudo"])
    return send_file(
        io.BytesIO(dados),
        mimetype=r["mime_type"],
        as_attachment=False,
        download_name=r["nome_arquivo"]
    )


@app.route("/api/anexos/<int:anexo_id>/download", methods=["GET"])
def api_baixar_anexo(anexo_id):
    row = db_query("SELECT nome_arquivo, mime_type, conteudo FROM anexos WHERE id = %s", (anexo_id,))
    if not row:
        return "Arquivo não encontrado", 404
    r = row[0]
    dados = bytes(r["conteudo"])
    return send_file(
        io.BytesIO(dados),
        mimetype=r["mime_type"],
        as_attachment=True,
        download_name=r["nome_arquivo"]
    )


@app.route("/api/demanda/<int:demanda_id>/gerar-minuta-oficio", methods=["POST"])
def api_gerar_minuta_oficio(demanda_id):
    dem = db_query("SELECT * FROM demandas WHERE id = %s", (demanda_id,))
    if not dem:
        return jsonify({"erro": "Demanda não encontrada."}), 404
    dem_row = dem[0]
    it_row = db_query("SELECT * FROM prass_items WHERE codigo = %s", (dem_row["item_codigo"],))[0]

    chks = db_query(
        "SELECT * FROM checklist_evidencias WHERE demanda_id = %s ORDER BY setor ASC, id ASC",
        (demanda_id,)
    )
    arquivos_pdf = []
    arquivos_excel = []
    for c in chks:
        c["concluido"] = bool(c["concluido"])
        anexos = db_query(
            "SELECT id, nome_arquivo, mime_type, tamanho, resumo_extraido FROM anexos WHERE checklist_id = %s ORDER BY id ASC",
            (c["id"],)
        )
        c["anexos"] = anexos
        for a in anexos:
            ext = os.path.splitext(a["nome_arquivo"].lower())[1]
            if ext in (".xlsx", ".xls", ".csv"):
                arquivos_excel.append({"setor": c["setor"], "nome_arquivo": a["nome_arquivo"], "titulo_chk": c["titulo"]})
            else:
                arquivos_pdf.append({"setor": c["setor"], "nome_arquivo": a["nome_arquivo"], "titulo_chk": c["titulo"]})

    body = request.get_json(silent=True) or {}
    forcar_regeneracao = bool(body.get("regenerar", False))

    if dem_row.get("oficio_texto") and dem_row["oficio_texto"].strip() and not forcar_regeneracao:
        texto_minuta = dem_row["oficio_texto"]
    else:
        texto_minuta = construir_texto_oficio_automatico(it_row, dem_row, chks)
        db_query(
            "UPDATE demandas SET oficio_texto = %s WHERE id = %s",
            (texto_minuta, demanda_id),
            fetch=False,
            commit=True
        )

    return jsonify({
        "texto_oficio": texto_minuta,
        "arquivos_pdf": arquivos_pdf,
        "arquivos_excel": arquivos_excel,
        "item_codigo": it_row["codigo"],
        "demanda_nome": dem_row["nome"]
    })


@app.route("/api/demanda/<int:demanda_id>/salvar-oficio", methods=["POST"])
def api_salvar_oficio(demanda_id):
    body = request.get_json(force=True)
    texto = body.get("texto_oficio", "")
    db_query(
        "UPDATE demandas SET oficio_texto = %s WHERE id = %s",
        (texto, demanda_id),
        fetch=False,
        commit=True
    )
    return jsonify({"ok": True})


@app.route("/api/demanda/<int:demanda_id>/exportar-pdf-unificado", methods=["POST"])
def api_exportar_pdf_unificado(demanda_id):
    body = request.get_json(force=True)
    texto_oficio = body.get("texto_oficio", "").strip()

    dem = db_query("SELECT * FROM demandas WHERE id = %s", (demanda_id,))
    if not dem:
        return jsonify({"erro": "Demanda não encontrada."}), 404
    dem_row = dem[0]
    it_row = db_query("SELECT * FROM prass_items WHERE codigo = %s", (dem_row["item_codigo"],))[0]

    db_query(
        "UPDATE demandas SET oficio_texto = %s WHERE id = %s",
        (texto_oficio, demanda_id),
        fetch=False,
        commit=True
    )

    chks = db_query(
        "SELECT * FROM checklist_evidencias WHERE demanda_id = %s ORDER BY setor ASC, id ASC",
        (demanda_id,)
    )

    lista_pdfs_info = []
    lista_excel_info = []
    anexos_para_merge = []

    for c in chks:
        anexos = db_query(
            "SELECT id, nome_arquivo, mime_type, conteudo FROM anexos WHERE checklist_id = %s ORDER BY id ASC",
            (c["id"],)
        )
        for a in anexos:
            ext = os.path.splitext(a["nome_arquivo"].lower())[1]
            if ext in (".xlsx", ".xls", ".csv"):
                lista_excel_info.append({
                    "setor": c["setor"],
                    "nome_arquivo": a["nome_arquivo"],
                    "titulo_chk": c["titulo"]
                })
            elif ext in (".pdf", ".png", ".jpg", ".jpeg"):
                lista_pdfs_info.append({
                    "setor": c["setor"],
                    "nome_arquivo": a["nome_arquivo"],
                    "titulo_chk": c["titulo"]
                })
                anexos_para_merge.append((ext, a["nome_arquivo"], bytes(a["conteudo"])))

    capa_buffer = gerar_pdf_capa_oficio(
        it_row["titulo"],
        dem_row["nome"],
        texto_oficio,
        lista_excel_info,
        lista_pdfs_info
    )

    writer = pypdf.PdfWriter()
    capa_reader = pypdf.PdfReader(capa_buffer)
    for page in capa_reader.pages:
        writer.add_page(page)

    for ext, nome_arq, raw_bytes in anexos_para_merge:
        try:
            if ext == ".pdf":
                r_anexo = pypdf.PdfReader(io.BytesIO(raw_bytes))
                for p in r_anexo.pages:
                    writer.add_page(p)
            elif ext in (".png", ".jpg", ".jpeg"):
                pdf_img_buf = converter_imagem_para_pdf_bytes(raw_bytes)
                r_img = pypdf.PdfReader(pdf_img_buf)
                for p in r_img.pages:
                    writer.add_page(p)
        except Exception as err:
            print(f"Aviso ao mesclar {nome_arq}: {err}")

    pdf_final = io.BytesIO()
    writer.write(pdf_final)
    pdf_final.seek(0)

    nome_download = f"Dossie_Oficio_PRASS_{it_row['codigo'].replace('.', '_')}_{dem_row['nome'].replace('/', '-').replace(' ', '_')}.pdf"
    return send_file(
        pdf_final,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=nome_download
    )


# ==============================================================================
# FRONTEND WEB COMPLETO EMBUTIDO (HTML5 + CSS3 + JAVASCRIPT VANILLA)
# ==============================================================================
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Postal Saúde | Sistema de Automação PRASS — GERED (CODER • COCAP • CONEC)</title>
<style>
    :root {
        --azul-postal: #003366;
        --azul-escuro: #001F3F;
        --amarelo-postal: #F2A900;
        --amarelo-claro: #FFF8E1;
        --bg-geral: #F0F4F8;
        --card-bg: #FFFFFF;
        --texto: #0F172A;
        --cinza-borda: #CBD5E1;
        --verde: #198754;
        --verde-bg: #D1E7DD;
        --amarelo-status: #B45309;
        --amarelo-status-bg: #FEF3C7;
        --vermelho: #DC3545;
        --vermelho-bg: #F8D7DA;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
        font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
        background-color: var(--bg-geral);
        color: var(--texto);
        display: flex;
        flex-direction: column;
        height: 100vh;
        overflow: hidden;
    }
    header {
        background: linear-gradient(90deg, var(--azul-escuro) 0%, var(--azul-postal) 100%);
        color: #FFF;
        padding: 12px 24px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        border-bottom: 4px solid var(--amarelo-postal);
        flex-shrink: 0;
    }
    .brand-title { font-size: 18px; font-weight: 800; letter-spacing: 0.3px; }
    .brand-sub { font-size: 12px; color: #FFD100; margin-top: 2px; }
    .db-badge {
        background: rgba(255,255,255,0.12);
        border: 1px solid rgba(255,255,255,0.25);
        padding: 6px 12px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 600;
    }
    .main-container {
        display: flex;
        flex: 1;
        overflow: hidden;
    }
    .sidebar {
        width: 390px;
        background: #FFFFFF;
        border-right: 1px solid var(--cinza-borda);
        display: flex;
        flex-direction: column;
        flex-shrink: 0;
    }
    .sidebar-header {
        padding: 14px 16px;
        background: #F8FAFC;
        border-bottom: 1px solid var(--cinza-borda);
        font-weight: 700;
        color: var(--azul-postal);
        font-size: 13px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .items-list {
        flex: 1;
        overflow-y: auto;
        padding: 10px;
    }
    .prass-item-card {
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 12px;
        margin-bottom: 10px;
        cursor: pointer;
        transition: all 0.15s ease;
        background: #FFFFFF;
        position: relative;
    }
    .prass-item-card:hover {
        border-color: #94A3B8;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
    }
    .prass-item-card.active {
        border: 2px solid var(--azul-postal);
        background: #F0F7FF;
    }
    .item-code-badge {
        display: inline-block;
        background: var(--azul-postal);
        color: #FFF;
        font-weight: 800;
        font-size: 11px;
        padding: 2px 8px;
        border-radius: 4px;
        margin-bottom: 6px;
    }
    .prass-item-card.active .item-code-badge {
        background: var(--amarelo-postal);
        color: var(--azul-escuro);
    }
    .item-title-short {
        font-size: 12.5px;
        font-weight: 700;
        color: #1E293B;
        line-height: 1.35;
        margin-bottom: 8px;
    }
    .semaforo-row {
        display: flex;
        align-items: center;
        gap: 6px;
        flex-wrap: wrap;
        margin-top: 6px;
    }
    .area-pill {
        font-size: 10.5px;
        font-weight: 700;
        padding: 2px 8px;
        border-radius: 12px;
        display: inline-flex;
        align-items: center;
        gap: 4px;
    }
    .area-pill.VERMELHO { background: var(--vermelho-bg); color: #842029; border: 1px solid #F5C2C7; }
    .area-pill.AMARELO { background: var(--amarelo-status-bg); color: #92400E; border: 1px solid #FDE68A; }
    .area-pill.VERDE { background: var(--verde-bg); color: #0F5132; border: 1px solid #BADBCC; }
    .area-pill.INATIVO { background: #F1F5F9; color: #94A3B8; border: 1px dashed #CBD5E1; text-decoration: line-through; }
    .workspace {
        flex: 1;
        overflow-y: auto;
        padding: 18px 24px;
        display: flex;
        flex-direction: column;
        gap: 16px;
    }
    .top-item-banner {
        background: #FFFFFF;
        border: 1px solid var(--cinza-borda);
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .banner-header-row {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 16px;
    }
    .banner-title {
        font-size: 16px;
        font-weight: 800;
        color: var(--azul-postal);
        line-height: 1.35;
    }
    .banner-actions {
        display: flex;
        gap: 8px;
        flex-shrink: 0;
        align-items: center;
    }
    .dropdown-areas {
        position: relative;
        display: inline-block;
    }
    .dropdown-btn {
        background: #F8FAFC;
        border: 1px solid var(--cinza-borda);
        padding: 7px 12px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 700;
        color: var(--azul-postal);
        cursor: pointer;
    }
    .dropdown-menu {
        display: none;
        position: absolute;
        right: 0;
        top: 110%;
        background: #FFF;
        border: 1px solid var(--cinza-borda);
        border-radius: 8px;
        box-shadow: 0 8px 20px rgba(0,0,0,0.12);
        padding: 10px 14px;
        z-index: 100;
        min-width: 210px;
    }
    .dropdown-menu.show { display: block; }
    .dropdown-menu label {
        display: flex;
        align-items: center;
        gap: 8px;
        font-size: 12.5px;
        font-weight: 600;
        padding: 5px 0;
        cursor: pointer;
    }
    .planilha-buttons-row {
        display: flex;
        gap: 10px;
        margin-top: 14px;
        flex-wrap: wrap;
    }
    .btn-info-planilha {
        flex: 1;
        min-width: 220px;
        padding: 9px 14px;
        border-radius: 7px;
        border: 1px solid #CBD5E1;
        background: #F8FAFC;
        color: #1E293B;
        font-weight: 700;
        font-size: 12.5px;
        cursor: pointer;
        text-align: left;
        display: flex;
        align-items: center;
        justify-content: space-between;
        transition: all 0.15s;
    }
    .btn-info-planilha:hover { background: #E2E8F0; }
    .btn-info-planilha.active {
        background: var(--azul-postal);
        color: #FFFFFF;
        border-color: var(--azul-postal);
    }
    .info-drawer {
        display: none;
        margin-top: 12px;
        background: #F8FAFC;
        border-left: 4px solid var(--azul-postal);
        border-radius: 6px;
        padding: 14px 18px;
        font-size: 13px;
        line-height: 1.55;
        white-space: pre-wrap;
        color: #1E293B;
        max-height: 280px;
        overflow-y: auto;
    }
    .info-drawer.open { display: block; }
    .controls-bar {
        background: #FFFFFF;
        border: 1px solid var(--cinza-borda);
        border-radius: 10px;
        padding: 14px 20px;
        display: flex;
        flex-wrap: wrap;
        gap: 18px;
        align-items: center;
        justify-content: space-between;
    }
    .control-group {
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
    }
    .control-label {
        font-size: 12px;
        font-weight: 800;
        text-transform: uppercase;
        color: #475569;
    }
    .seg-btn {
        padding: 7px 14px;
        border-radius: 6px;
        border: 1px solid var(--cinza-borda);
        background: #F8FAFC;
        font-size: 12.5px;
        font-weight: 700;
        cursor: pointer;
        color: #334155;
    }
    .seg-btn.active {
        background: var(--azul-postal);
        color: #FFFFFF;
        border-color: var(--azul-postal);
    }
    select.select-demanda {
        padding: 7px 12px;
        border-radius: 6px;
        border: 1px solid var(--azul-postal);
        font-size: 13px;
        font-weight: 700;
        color: var(--azul-postal);
        background: #FFF;
        min-width: 220px;
    }
    .setor-btn {
        padding: 8px 18px;
        border-radius: 7px;
        border: 2px solid transparent;
        font-size: 13px;
        font-weight: 800;
        cursor: pointer;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #E2E8F0;
        color: #1E293B;
    }
    .setor-btn.active {
        background: var(--azul-postal);
        color: #FFFFFF;
        border-color: var(--amarelo-postal);
    }
    .checklist-panel {
        background: #FFFFFF;
        border: 1px solid var(--cinza-borda);
        border-radius: 10px;
        padding: 18px 20px;
        display: flex;
        flex-direction: column;
        gap: 14px;
    }
    .chk-header-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 12px;
        flex-wrap: wrap;
    }
    .progress-wrap {
        flex: 1;
        min-width: 260px;
    }
    .progress-bar-outer {
        width: 100%;
        height: 12px;
        background: #E2E8F0;
        border-radius: 6px;
        overflow: hidden;
        margin-top: 5px;
    }
    .progress-bar-inner {
        height: 100%;
        background: var(--verde);
        width: 0%;
        transition: width 0.25s ease;
    }
    .add-chk-row {
        display: flex;
        gap: 8px;
        background: #F8FAFC;
        padding: 10px;
        border-radius: 8px;
        border: 1px dashed #94A3B8;
    }
    .add-chk-row input {
        flex: 1;
        padding: 8px 12px;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        font-size: 13px;
    }
    .btn-primary {
        background: var(--azul-postal);
        color: #FFF;
        border: none;
        padding: 8px 14px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 12.5px;
        cursor: pointer;
    }
    .btn-success {
        background: var(--verde);
        color: #FFF;
        border: none;
        padding: 8px 14px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 12.5px;
        cursor: pointer;
    }
    .btn-warning {
        background: var(--amarelo-postal);
        color: var(--azul-escuro);
        border: none;
        padding: 8px 14px;
        border-radius: 6px;
        font-weight: 800;
        font-size: 12.5px;
        cursor: pointer;
    }
    .chk-item-card {
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 12px 14px;
        background: #FFFFFF;
        display: flex;
        flex-direction: column;
        gap: 8px;
    }
    .chk-item-card.concluido {
        border-left: 5px solid var(--verde);
        background: #F8FFF9;
    }
    .chk-item-card.pendente {
        border-left: 5px solid var(--vermelho);
    }
    .chk-top-line {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 10px;
    }
    .chk-title-group {
        display: flex;
        align-items: center;
        gap: 10px;
        flex: 1;
    }
    .chk-title-text {
        font-size: 13.5px;
        font-weight: 700;
        color: #0F172A;
    }
    .chk-tools {
        display: flex;
        align-items: center;
        gap: 6px;
        flex-wrap: wrap;
    }
    .tool-btn {
        border: 1px solid #CBD5E1;
        background: #F8FAFC;
        color: #1E293B;
        padding: 5px 9px;
        border-radius: 5px;
        font-size: 11.5px;
        font-weight: 700;
        cursor: pointer;
    }
    .tool-btn:hover { background: #E2E8F0; }
    .tool-btn.concluir-on { background: var(--verde); color: #FFF; border-color: var(--verde); }
    .tool-btn.danger { background: #FEE2E2; color: #991B1B; border-color: #FCA5A5; }
    .anexos-box {
        background: #F8FAFC;
        border-radius: 6px;
        padding: 8px 10px;
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        align-items: center;
    }
    .anexo-chip {
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        padding: 4px 8px;
        font-size: 11.5px;
        display: inline-flex;
        align-items: center;
        gap: 6px;
    }
    .anexo-chip a {
        text-decoration: none;
        font-weight: 700;
        color: var(--azul-postal);
    }
    .modal-backdrop {
        display: none;
        position: fixed;
        inset: 0;
        background: rgba(15, 23, 42, 0.65);
        z-index: 1000;
        align-items: center;
        justify-content: center;
        padding: 20px;
    }
    .modal-backdrop.open { display: flex; }
    .modal-window {
        background: #FFFFFF;
        width: 100%;
        max-width: 980px;
        max-height: 90vh;
        border-radius: 12px;
        display: flex;
        flex-direction: column;
        overflow: hidden;
        box-shadow: 0 20px 50px rgba(0,0,0,0.3);
    }
    .modal-header {
        background: var(--azul-postal);
        color: #FFF;
        padding: 14px 20px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .modal-body {
        padding: 18px 20px;
        overflow-y: auto;
        flex: 1;
        display: flex;
        flex-direction: column;
        gap: 12px;
    }
    .oficio-textarea {
        width: 100%;
        min-height: 360px;
        font-family: 'Consolas', 'Courier New', monospace;
        font-size: 12.5px;
        line-height: 1.5;
        padding: 14px;
        border: 1px solid #94A3B8;
        border-radius: 8px;
        resize: vertical;
    }
    .modal-footer {
        background: #F1F5F9;
        padding: 12px 20px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-top: 1px solid #CBD5E1;
    }
</style>
</head>
<body>

<header>
    <div>
        <div class="brand-title">POSTAL SAÚDE | SISTEMA DE AUTOMAÇÃO PRASS — GERED</div>
        <div class="brand-sub">Módulos 8.1.2 a 8.6.1 • Coordenações CODER, COCAP e CONEC • Geração Automática de Ofício + Dossiê PDF Unificado</div>
    </div>
    <div class="db-badge" id="dbBadge">Conectando ao Banco...</div>
</header>

<div class="main-container">
    <aside class="sidebar">
        <div class="sidebar-header">
            <span>MÓDULOS DO PRASS (GERED)</span>
            <span id="lblDemandaSidebar" style="font-size:11px; color:#0D6EFD;"></span>
        </div>
        <div class="items-list" id="sidebarItems"></div>
    </aside>

    <main class="workspace">
        <section class="top-item-banner">
            <div class="banner-header-row">
                <div>
                    <div class="banner-title" id="mainItemTitulo">Carregando...</div>
                    <div id="mainSemaforoBadges" class="semaforo-row" style="margin-top:8px;"></div>
                </div>
                <div class="banner-actions">
                    <div class="dropdown-areas">
                        <button class="dropdown-btn" onclick="toggleAreasDropdown()">⚙️ Áreas Atribuídas ▼</button>
                        <div class="dropdown-menu" id="areasDropdownMenu">
                            <div style="font-size:11px; color:#64748B; margin-bottom:6px; font-weight:700;">Atribuir a este Item do PRASS:</div>
                            <label><input type="checkbox" value="CODER" onchange="salvarAreasAtribuidas()"> CODER</label>
                            <label><input type="checkbox" value="COCAP" onchange="salvarAreasAtribuidas()"> COCAP</label>
                            <label><input type="checkbox" value="CONEC" onchange="salvarAreasAtribuidas()"> CONEC</label>
                        </div>
                    </div>

                    <button class="btn-warning" onclick="abrirModalOficio()">
                        📄 Gerar Ofício + Unificar PDF
                    </button>
                </div>
            </div>

            <div class="planilha-buttons-row">
                <button class="btn-info-planilha" id="btnDrawerApontamento" onclick="toggleDrawer('apontamento')">
                    <span>📌 1. Apontamento da ANS & Diagnóstico (Col. B)</span>
                    <span>▼</span>
                </button>
                <button class="btn-info-planilha" id="btnDrawerMedidas" onclick="toggleDrawer('medidas')">
                    <span>🛠️ 2. Medidas Implementadas e em Implementação (Col. B)</span>
                    <span>▼</span>
                </button>
                <button class="btn-info-planilha" id="btnDrawerObs" onclick="toggleDrawer('obs')">
                    <span>👁️ 3. Observações — Tratamento e Monitoramento (Col. C)</span>
                    <span>▼</span>
                </button>
            </div>

            <div class="info-drawer" id="drawerContent"></div>
        </section>

        <section class="controls-bar">
            <div class="control-group">
                <span class="control-label">Tipo:</span>
                <button class="seg-btn active" id="btnTipoFixa" onclick="selecionarTipoDemanda('FIXA')">Demanda Fixa (08/26 a 08/27)</button>
                <button class="seg-btn" id="btnTipoAvulsa" onclick="selecionarTipoDemanda('AVULSA')">Demanda Avulsa</button>
            </div>

            <div class="control-group">
                <span class="control-label">Competência / Demanda:</span>
                <select class="select-demanda" id="selectDemanda" onchange="aoMudarDemanda()"></select>
                <button class="btn-primary" onclick="criarNovaDemandaAvulsa()">+ Nova Demanda</button>
            </div>

            <div class="control-group">
                <span class="control-label">Setor (GERED):</span>
                <button class="setor-btn" id="btnSetorCODER" onclick="selecionarSetor('CODER')">CODER</button>
                <button class="setor-btn active" id="btnSetorCOCAP" onclick="selecionarSetor('COCAP')">COCAP</button>
                <button class="setor-btn" id="btnSetorCONEC" onclick="selecionarSetor('CONEC')">CONEC</button>
            </div>
        </section>

        <section class="checklist-panel">
            <div class="chk-header-row">
                <div>
                    <h3 id="tituloChecklistSetor" style="color:var(--azul-postal); font-size:15px;">Checklist de Evidências — COCAP</h3>
                    <p id="subtituloHistoricoPlanilha" style="font-size:12px; color:#475569; margin-top:3px;"></p>
                </div>
                <div class="progress-wrap">
                    <div style="display:flex; justify-content:space-between; font-size:12px; font-weight:700;">
                        <span id="lblProgressoSetor">Progresso do Setor: 0/0 (0%)</span>
                        <span id="lblProgressoDemandaGeral" style="color:var(--azul-postal);">Total Geral da Demanda: 0%</span>
                    </div>
                    <div class="progress-bar-outer">
                        <div class="progress-bar-inner" id="barraProgressoSetor"></div>
                    </div>
                </div>
            </div>

            <div class="add-chk-row">
                <input type="text" id="inputNovaEvidencia" placeholder="Digite o título da nova evidência para o checklist deste setor e pressione Enter ou clique em (+)..." onkeydown="if(event.key==='Enter') adicionarEvidencia()">
                <button class="btn-success" onclick="adicionarEvidencia()">+ Adicionar Evidência</button>
            </div>

            <div id="listaChecklistContainer" style="display:flex; flex-direction:column; gap:10px;"></div>
        </section>
    </main>
</div>

<div class="modal-backdrop" id="modalOficio">
    <div class="modal-window">
        <div class="modal-header">
            <div>
                <div style="font-size:15px; font-weight:800;">📄 ANÁLISE AUTOMÁTICA DE EVIDÊNCIAS & MINUTA DO OFÍCIO</div>
                <div style="font-size:11.5px; color:#FFD100;" id="modalSubtitulo">Revise ou edite o texto abaixo antes de aprovar e gerar o PDF único com todas as evidências mescladas.</div>
            </div>
            <button onclick="fecharModalOficio()" style="background:none; border:none; color:#FFF; font-size:20px; cursor:pointer;">✖</button>
        </div>
        <div class="modal-body">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:12.5px; font-weight:700; color:#334155;">Texto do Ofício (Editável pelo Analista):</span>
                <button class="tool-btn" onclick="regenerarTextoOficio()">🔄 Atualizar Análise com Base nos Anexos Atuais</button>
            </div>
            <textarea class="oficio-textarea" id="textareaOficio"></textarea>
            <div id="resumoArquivosModal" style="background:#F8FAFC; border:1px solid #CBD5E1; border-radius:6px; padding:10px 14px; font-size:12px;"></div>
        </div>
        <div class="modal-footer">
            <button class="tool-btn" onclick="salvarRascunhoOficio()">💾 Salvar Edição no Banco</button>
            <div style="display:flex; gap:10px;">
                <button class="tool-btn" onclick="fecharModalOficio()">Cancelar</button>
                <button class="btn-success" onclick="aprovarEBaixarPDFUnificado()">✅ Aprovar Ofício e Baixar PDF Unificado (Ofício + Anexos)</button>
            </div>
        </div>
    </div>
</div>

<input type="file" id="hiddenFileInput" multiple style="display:none" onchange="executarUploadAnexos()">

<script>
let ESTADO_ITENS = [];
let ITEM_ATUAL = null;
let TIPO_DEMANDA_ATUAL = 'FIXA';
let DEMANDA_ATUAL_ID = null;
let DEMANDA_ATUAL_NOME = 'Complemento ANS 26/10';
let SETOR_ATUAL = 'COCAP';
let DRAWER_ABERTO = null;
let CHECKLIST_ATUAL = [];
let CHK_ID_PARA_UPLOAD = null;

async function carregarSistema() {
    try {
        const resp = await fetch(`/api/bootstrap?demanda_nome=${encodeURIComponent(DEMANDA_ATUAL_NOME)}`);
        if (!resp.ok) throw new Error('Falha ao buscar dados no servidor');
        const dados = await resp.json();
        ESTADO_ITENS = dados.itens;

        const badge = document.getElementById('dbBadge');
        badge.textContent = dados.postgres_ativo ? '🟢 PostgreSQL Railway Ativo' : '🟡 Banco Interno SQLite Ativo';

        if (!ITEM_ATUAL && ESTADO_ITENS.length > 0) {
            ITEM_ATUAL = ESTADO_ITENS.find(i => i.codigo === '8.5.1') || ESTADO_ITENS[0];
        } else if (ITEM_ATUAL) {
            ITEM_ATUAL = ESTADO_ITENS.find(i => i.codigo === ITEM_ATUAL.codigo) || ESTADO_ITENS[0];
        }

        renderizarSidebar();
        renderizarItemAtual();
    } catch (e) {
        console.error(e);
        document.getElementById('dbBadge').textContent = '🔴 Erro na API (Verifique Logs)';
    }
}

function renderizarSidebar() {
    const container = document.getElementById('sidebarItems');
    document.getElementById('lblDemandaSidebar').textContent = DEMANDA_ATUAL_NOME;
    container.innerHTML = '';

    ESTADO_ITENS.forEach(item => {
        const card = document.createElement('div');
        card.className = 'prass-item-card' + (ITEM_ATUAL && ITEM_ATUAL.codigo === item.codigo ? ' active' : '');
        card.onclick = () => {
            ITEM_ATUAL = item;
            renderizarSidebar();
            renderizarItemAtual();
        };

        let badgesHtml = '';
        ['CODER', 'COCAP', 'CONEC'].forEach(setor => {
            const st = item.semaforo && item.semaforo[setor];
            if (!item.areas_atribuidas.includes(setor)) {
                badgesHtml += `<span class="area-pill INATIVO">${setor}</span>`;
            } else if (st) {
                const icone = st.cor === 'VERDE' ? '🟢' : (st.cor === 'AMARELO' ? '🟡' : '🔴');
                badgesHtml += `<span class="area-pill ${st.cor}">${icone} ${setor}</span>`;
            }
        });

        card.innerHTML = `
            <div class="item-code-badge">ITEM ${item.codigo}</div>
            <div class="item-title-short">${item.titulo}</div>
            <div class="semaforo-row">${badgesHtml}</div>
        `;
        container.appendChild(card);
    });
}

function renderizarItemAtual() {
    if (!ITEM_ATUAL) return;
    document.getElementById('mainItemTitulo').textContent = ITEM_ATUAL.titulo;

    const checks = document.querySelectorAll('#areasDropdownMenu input[type="checkbox"]');
    checks.forEach(chk => {
        chk.checked = ITEM_ATUAL.areas_atribuidas.includes(chk.value);
    });

    if (DRAWER_ABERTO) {
        atualizarTextoDrawer();
    }

    populaSelectDemandas();
}

function toggleAreasDropdown() {
    document.getElementById('areasDropdownMenu').classList.toggle('show');
}

async function salvarAreasAtribuidas() {
    const checks = document.querySelectorAll('#areasDropdownMenu input[type="checkbox"]');
    const selecionadas = [];
    checks.forEach(c => { if (c.checked) selecionadas.push(c.value); });
    if (selecionadas.length === 0) {
        alert('Pelo menos uma coordenação deve estar atribuída ao item.');
        renderizarItemAtual();
        return;
    }
    await fetch(`/api/item/${encodeURIComponent(ITEM_ATUAL.codigo)}/areas`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({areas_atribuidas: selecionadas})
    });
    await carregarSistema();
}

function toggleDrawer(qual) {
    const drawer = document.getElementById('drawerContent');
    if (DRAWER_ABERTO === qual) {
        DRAWER_ABERTO = null;
        drawer.classList.remove('open');
    } else {
        DRAWER_ABERTO = qual;
        drawer.classList.add('open');
        atualizarTextoDrawer();
    }
    document.getElementById('btnDrawerApontamento').classList.toggle('active', DRAWER_ABERTO === 'apontamento');
    document.getElementById('btnDrawerMedidas').classList.toggle('active', DRAWER_ABERTO === 'medidas');
    document.getElementById('btnDrawerObs').classList.toggle('active', DRAWER_ABERTO === 'obs');
}

function atualizarTextoDrawer() {
    const drawer = document.getElementById('drawerContent');
    if (!ITEM_ATUAL || !DRAWER_ABERTO) return;
    if (DRAWER_ABERTO === 'apontamento') {
        drawer.textContent = ITEM_ATUAL.apontamento_ans;
    } else if (DRAWER_ABERTO === 'medidas') {
        drawer.textContent = ITEM_ATUAL.medidas_implementadas;
    } else if (DRAWER_ABERTO === 'obs') {
        drawer.textContent = ITEM_ATUAL.observacoes;
    }
}

function selecionarTipoDemanda(tipo) {
    TIPO_DEMANDA_ATUAL = tipo;
    document.getElementById('btnTipoFixa').classList.toggle('active', tipo === 'FIXA');
    document.getElementById('btnTipoAvulsa').classList.toggle('active', tipo === 'AVULSA');
    populaSelectDemandas();
}

function populaSelectDemandas() {
    const sel = document.getElementById('selectDemanda');
    sel.innerHTML = '';
    const filtradas = ITEM_ATUAL.demandas.filter(d => d.tipo === TIPO_DEMANDA_ATUAL);

    if (filtradas.length === 0) {
        const opt = document.createElement('option');
        opt.value = '';
        opt.textContent = TIPO_DEMANDA_ATUAL === 'AVULSA' ? '(Nenhuma demanda avulsa — clique em + Nova Demanda)' : '(Sem demandas)';
        sel.appendChild(opt);
        DEMANDA_ATUAL_ID = null;
        document.getElementById('listaChecklistContainer').innerHTML = '';
        return;
    }

    filtradas.forEach(d => {
        const opt = document.createElement('option');
        opt.value = d.id;
        opt.textContent = d.nome;
        if (d.nome === DEMANDA_ATUAL_NOME) opt.selected = true;
        sel.appendChild(opt);
    });

    DEMANDA_ATUAL_ID = parseInt(sel.value, 10);
    const demObj = filtradas.find(d => d.id === DEMANDA_ATUAL_ID);
    if (demObj) DEMANDA_ATUAL_NOME = demObj.nome;

    carregarChecklistDemanda();
}

async function aoMudarDemanda() {
    const sel = document.getElementById('selectDemanda');
    if (!sel.value) return;
    DEMANDA_ATUAL_ID = parseInt(sel.value, 10);
    const demObj = ITEM_ATUAL.demandas.find(d => d.id === DEMANDA_ATUAL_ID);
    if (demObj) DEMANDA_ATUAL_NOME = demObj.nome;
    await carregarSistema();
}

async function criarNovaDemandaAvulsa() {
    const nome = prompt('Digite o nome da nova Demanda Avulsa (ex: Ofício 108/2026 - Revisão Extra):');
    if (!nome || !nome.trim()) return;
    await fetch('/api/demandas', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            item_codigo: ITEM_ATUAL.codigo,
            nome: nome.trim(),
            replicar_todos: true
        })
    });
    TIPO_DEMANDA_ATUAL = 'AVULSA';
    DEMANDA_ATUAL_NOME = nome.trim();
    document.getElementById('btnTipoFixa').classList.remove('active');
    document.getElementById('btnTipoAvulsa').classList.add('active');
    await carregarSistema();
}

function selecionarSetor(setor) {
    SETOR_ATUAL = setor;
    ['CODER', 'COCAP', 'CONEC'].forEach(s => {
        document.getElementById('btnSetor' + s).classList.toggle('active', s === setor);
    });
    renderizarChecklistNaTela();
}

async function carregarChecklistDemanda() {
    if (!DEMANDA_ATUAL_ID) return;
    const resp = await fetch(`/api/checklist/${DEMANDA_ATUAL_ID}`);
    const dados = await resp.json();
    CHECKLIST_ATUAL = dados.checklist || [];
    ITEM_ATUAL.semaforo = dados.semaforo;

    renderizarSemaforoTopo(dados.semaforo);
    renderizarChecklistNaTela();
}

function renderizarSemaforoTopo(semaforo) {
    const box = document.getElementById('mainSemaforoBadges');
    box.innerHTML = '';
    ['CODER', 'COCAP', 'CONEC'].forEach(setor => {
        const st = semaforo[setor];
        const btnSetor = document.getElementById('btnSetor' + setor);
        if (!ITEM_ATUAL.areas_atribuidas.includes(setor)) {
            box.innerHTML += `<span class="area-pill INATIVO">${setor}: Não atribuído</span>`;
            btnSetor.innerHTML = `${setor}`;
        } else if (st) {
            const icone = st.cor === 'VERDE' ? '🟢' : (st.cor === 'AMARELO' ? '🟡' : '🔴');
            box.innerHTML += `<span class="area-pill ${st.cor}">${icone} ${setor}: ${st.label}</span>`;
            btnSetor.innerHTML = `${icone} ${setor} (${st.concluidos}/${st.total})`;
        }
    });
}

function renderizarChecklistNaTela() {
    document.getElementById('tituloChecklistSetor').textContent =
        `Checklist de Evidências — Coordenação ${SETOR_ATUAL} | Demanda: ${DEMANDA_ATUAL_NOME}`;

    const histTxt = (ITEM_ATUAL.historico_planilha && ITEM_ATUAL.historico_planilha[DEMANDA_ATUAL_NOME]) || '';
    document.getElementById('subtituloHistoricoPlanilha').textContent =
        histTxt && histTxt !== 'N/A' ? `Resumo registrado na planilha para ${DEMANDA_ATUAL_NOME}: ${histTxt.substring(0, 220)}...` : '';

    const itensSetor = CHECKLIST_ATUAL.filter(c => c.setor === SETOR_ATUAL);
    const concluidosSetor = itensSetor.filter(c => c.concluido).length;
    const percSetor = itensSetor.length > 0 ? Math.round((concluidosSetor / itensSetor.length) * 100) : 0;

    document.getElementById('lblProgressoSetor').textContent =
        `Progresso ${SETOR_ATUAL}: ${concluidosSetor} de ${itensSetor.length} concluídos (${percSetor}%)`;
    document.getElementById('barraProgressoSetor').style.width = `${percSetor}%`;

    const totalGeral = CHECKLIST_ATUAL.length;
    const concGeral = CHECKLIST_ATUAL.filter(c => c.concluido).length;
    const percGeral = totalGeral > 0 ? Math.round((concGeral / totalGeral) * 100) : 0;
    document.getElementById('lblProgressoDemandaGeral').textContent =
        `Total da Demanda (3 Coordenações): ${concGeral}/${totalGeral} (${percGeral}%)`;

    const container = document.getElementById('listaChecklistContainer');
    container.innerHTML = '';

    if (itensSetor.length === 0) {
        container.innerHTML = `
            <div style="padding:24px; text-align:center; color:#64748B; background:#F8FAFC; border-radius:8px;">
                Nenhuma evidência cadastrada para <b>${SETOR_ATUAL}</b> em <b>${DEMANDA_ATUAL_NOME}</b>.<br>
                Use o campo acima para adicionar itens ao checklist.
            </div>`;
        return;
    }

    itensSetor.forEach((chk, idx) => {
        const card = document.createElement('div');
        card.className = 'chk-item-card ' + (chk.concluido ? 'concluido' : 'pendente');

        let anexosHtml = '';
        if (chk.anexos && chk.anexos.length > 0) {
            chk.anexos.forEach(a => {
                anexosHtml += `
                    <div class="anexo-chip">
                        <span>📄 ${a.nome_arquivo}</span>
                        <a href="/api/anexos/${a.id}/view" target="_blank" title="Visualizar Anexo">👁️ Ver</a>
                        <a href="/api/anexos/${a.id}/download" title="Baixar Anexo">⬇️ Baixar</a>
                        <button onclick="removerAnexo(${a.id})" style="border:none; background:none; color:#DC3545; font-weight:800; cursor:pointer;" title="Remover Anexo (volta para Pendente)">✖</button>
                    </div>
                `;
            });
        } else {
            anexosHtml = `<span style="font-size:11.5px; color:#64748B;">Nenhum anexo incluído ainda.</span>`;
        }

        card.innerHTML = `
            <div class="chk-top-line">
                <div class="chk-title-group">
                    <span style="font-weight:800; font-size:12px; color:#003366;">#${idx + 1}</span>
                    <span class="chk-title-text">${chk.titulo}</span>
                </div>
                <div class="chk-tools">
                    <button class="tool-btn" onclick="acionarUpload(${chk.id})">📎 Incluir Anexo(s)</button>
                    <button class="tool-btn" onclick="editarEvidencia(${chk.id})">✏️ Editar</button>
                    <button class="tool-btn danger" onclick="excluirEvidencia(${chk.id})">🗑️ Excluir</button>
                    <button class="tool-btn ${chk.concluido ? 'concluir-on' : ''}" onclick="alternarConclusao(${chk.id}, ${!chk.concluido}, ${chk.anexos.length})">
                        ${chk.concluido ? '✔ Concluído (Reabrir)' : '✔ Concluir'}
                    </button>
                </div>
            </div>
            <div class="anexos-box">${anexosHtml}</div>
        `;
        container.appendChild(card);
    });
}

async function adicionarEvidencia() {
    if (!DEMANDA_ATUAL_ID) {
        alert('Selecione ou crie uma demanda primeiro.');
        return;
    }
    const input = document.getElementById('inputNovaEvidencia');
    const titulo = input.value.trim();
    if (!titulo) return;

    await fetch('/api/checklist', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            demanda_id: DEMANDA_ATUAL_ID,
            setor: SETOR_ATUAL,
            titulo: titulo
        })
    });
    input.value = '';
    await carregarChecklistDemanda();
    await carregarSistema();
}

async function editarEvidencia(chkId) {
    const chk = CHECKLIST_ATUAL.find(c => c.id === chkId);
    if (!chk) return;
    const novoTitulo = prompt('Editar descrição da evidência:', chk.titulo);
    if (!novoTitulo || !novoTitulo.trim()) return;
    const novaAnalise = prompt('Observação/Explicação de como obteve esta evidência (para sair no Ofício):', chk.analise_evidencia || '');

    await fetch(`/api/checklist/${chkId}`, {
        method: 'PUT',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            titulo: novoTitulo.trim(),
            analise_evidencia: novaAnalise !== null ? novaAnalise : (chk.analise_evidencia || '')
        })
    });
    await carregarChecklistDemanda();
}

async function excluirEvidencia(chkId) {
    if (!confirm('Deseja realmente excluir este item do checklist e seus anexos?')) return;
    await fetch(`/api/checklist/${chkId}`, { method: 'DELETE' });
    await carregarChecklistDemanda();
    await carregarSistema();
}

async function alternarConclusao(chkId, novoStatus, qtdAnexos) {
    if (novoStatus && qtdAnexos === 0) {
        const ok = confirm('Atenção: Este item ainda não possui anexo importado. Deseja marcá-lo como CONCLUÍDO mesmo assim?');
        if (!ok) return;
    }
    await fetch(`/api/checklist/${chkId}`, {
        method: 'PUT',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ concluido: novoStatus })
    });
    await carregarChecklistDemanda();
    await carregarSistema();
}

function acionarUpload(chkId) {
    CHK_ID_PARA_UPLOAD = chkId;
    const inp = document.getElementById('hiddenFileInput');
    inp.value = '';
    inp.click();
}

async function executarUploadAnexos() {
    const inp = document.getElementById('hiddenFileInput');
    if (!inp.files || inp.files.length === 0 || !CHK_ID_PARA_UPLOAD) return;

    const formData = new FormData();
    for (let i = 0; i < inp.files.length; i++) {
        formData.append('arquivos', inp.files[i]);
    }

    await fetch(`/api/checklist/${CHK_ID_PARA_UPLOAD}/anexos`, {
        method: 'POST',
        body: formData
    });
    await carregarChecklistDemanda();
    await carregarSistema();
}

async function removerAnexo(anexoId) {
    if (!confirm('Remover este anexo? O status do item voltará automaticamente para PENDENTE.')) return;
    await fetch(`/api/anexos/${anexoId}`, { method: 'DELETE' });
    await carregarChecklistDemanda();
    await carregarSistema();
}

async function abrirModalOficio() {
    if (!DEMANDA_ATUAL_ID) {
        alert('Selecione uma demanda primeiro.');
        return;
    }
    const resp = await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/gerar-minuta-oficio`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ regenerar: false })
    });
    const dados = await resp.json();
    preencherModalOficio(dados);
    document.getElementById('modalOficio').classList.add('open');
}

async function regenerarTextoOficio() {
    const resp = await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/gerar-minuta-oficio`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ regenerar: true })
    });
    const dados = await resp.json();
    preencherModalOficio(dados);
}

function preencherModalOficio(dados) {
    document.getElementById('textareaOficio').value = dados.texto_oficio || '';
    const box = document.getElementById('resumoArquivosModal');
    const qtdPdf = (dados.arquivos_pdf || []).length;
    const qtdXls = (dados.arquivos_excel || []).length;

    let html = `<b>📦 Composição do Dossiê Unificado que será gerado após sua aprovação:</b><br>`;
    html += `• <b>1º) Ofício de Resposta (acima)</b> convertido automaticamente nas primeiras páginas do PDF;<br>`;
    html += `• <b>2º) ${qtdPdf} anexo(s) PDF/Imagem</b> mesclados integralmente logo abaixo do Ofício em um único PDF;<br>`;
    if (qtdXls > 0) {
        const nomesXls = dados.arquivos_excel.map(x => x.nome_arquivo).join(', ');
        html += `• <b>3º) ${qtdXls} planilha(s) Excel/CSV</b> citadas expressamente pelo nome no corpo e no índice do Ofício: <code>${nomesXls}</code>.`;
    } else {
        html += `• Nenhuma planilha Excel/CSV externa pendente de citação nominal.`;
    }
    box.innerHTML = html;
}

function fecharModalOficio() {
    document.getElementById('modalOficio').classList.remove('open');
}

async function salvarRascunhoOficio() {
    const texto = document.getElementById('textareaOficio').value;
    await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/salvar-oficio`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ texto_oficio: texto })
    });
    alert('Texto do Ofício salvo com sucesso no banco de dados!');
}

async function aprovarEBaixarPDFUnificado() {
    const texto = document.getElementById('textareaOficio').value;
    const resp = await fetch(`/api/demanda/${DEMANDA_ATUAL_ID}/exportar-pdf-unificado`, {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ texto_oficio: texto })
    });
    if (!resp.ok) {
        alert('Erro ao gerar o PDF unificado.');
        return;
    }
    const blob = await resp.blob();
    const nomeSugerido = `Oficio_e_Evidencias_PRASS_${ITEM_ATUAL.codigo.replace(/\./g, '_')}_${DEMANDA_ATUAL_NOME.replace(/[^a-zA-Z0-9]/g, '_')}.pdf`;

    if (window.showSaveFilePicker) {
        try {
            const handle = await window.showSaveFilePicker({
                suggestedName: nomeSugerido,
                types: [{
                    description: 'Documento PDF Unificado',
                    accept: {'application/pdf': ['.pdf']}
                }]
            });
            const writable = await handle.createWritable();
            await writable.write(blob);
            await writable.close();
            fecharModalOficio();
            return;
        } catch (err) {
            if (err.name === 'AbortError') return;
        }
    }

    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = nomeSugerido;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
    fecharModalOficio();
}

window.addEventListener('DOMContentLoaded', carregarSistema);
</script>
</body>
</html>
