"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  SM · Sistema 2 — Maestro de Clientes                                       ║
║  Backend Python: lectura SAP, alertas A1-A5, SQLite, exportación Excel,     ║
║  servidor HTTP local, dashboard HTML autocontenido.                          ║
║  Sofgen Pharma · Cumplimiento Legal Corporativo · SAGRILAFT                  ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import pandas as pd
import os
import json
import base64
import webbrowser
import urllib.parse
import threading
import sqlite3
import re
from http.server import SimpleHTTPRequestHandler, HTTPServer
from pathlib import Path
from datetime import datetime, date
from socketserver import ThreadingMixIn

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN MAESTRA
# ══════════════════════════════════════════════════════════════════════════════
SM_DIR = Path(__file__).resolve().parent
NOMBRE_EXCEL    = "Maestro_Prueba_SAGRILAFT.xlsx"
ARCHIVO_EXCEL   = SM_DIR / NOMBRE_EXCEL
HOJA_PRINCIPAL  = "Maestro_Clientes"
HOJA_CIIU       = "CIIU_Clasificacion"
HOJA_PAISES     = "Paises_Riesgo"
HOJA_SOCIEDADES = "Sociedades_Ref"
DASHBOARD_HTML  = SM_DIR / "dashboard_maestro.html"
LOGO_FILE       = SM_DIR.parent / "Assets" / "logo.png"
SM_LOGO_FILE    = SM_DIR.parent / "Assets" / "logo_sm.png"
DB_FILE         = SM_DIR / "sm_maestro_historico.db"
PUERTO          = 8081

def _cargar_config():
    config_path = SM_DIR.parent / "sm_config.json"
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"   ℹ️  sm_config.json no encontrado. Usando configuración por defecto.")
        return {}
    except Exception as e:
        print(f"   ⚠️  Error leyendo sm_config.json: {e}. Usando configuración por defecto.")
        return {}

_CONFIG = _cargar_config()
_RUTA_DEFAULT = str(SM_DIR.parent / "SAGRILAFT" / "Maestro de clientes")
CARPETA_RAIZ = Path(_CONFIG.get('rutas', {}).get('arc_exportacion', _RUTA_DEFAULT))

MESES_ES     = ['01 - Enero','02 - Febrero','03 - Marzo','04 - Abril','05 - Mayo','06 - Junio','07 - Julio','08 - Agosto','09 - Septiembre','10 - Octubre','11 - Noviembre','12 - Diciembre']
MESES_NOMBRE = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
NIVEL_FOLDER_MAP = {'ALTO': '1. Riesgo Alto', 'MEDIO': '2. Riesgo Medio', 'BAJO': '3. Riesgo Bajo'}

SOCIEDADES_NOMBRES = {
    "CO01":"PROCAPS S.A.","CO02":"PHARMAYECT SA","CO03":"C.I. PROCAPS SA","CO04":"C.I. NATURMEGA SA","CO05":"USGP A DIVISION OF PROCAP",
    "CO06":"DIABETRICS HEALTHCARE SAS","CO07":"COLMED LTDA.","CO08":"FUNDACION PROCAPS","CO09":"IGT","CO10":"INVERSIONES HENIA SAS",
    "CO11":"INVERSIONES CRYNSEEN SAS","CO12":"INVERSIONES JADES SAS","CO13":"INVERSIONES GANEDEN SAS","CO14":"INDUSTRIAS KADIMA SAS",
    "CO19":"CRYNSSEN PHARMA S.A.S.","CO20":"FUNTRITION SAS","CO21":"RYMCO S.A.","CO22":"RYMCO MEDICAL S.A.S",
    "CR01":"PHARMARKETING COSTA RICA","DO01":"PHARMARKETING DOMINICANA","EC01":"RODDOME PHARMACEUTICAL SA","ES02":"ALLOPHAME HOLDING",
    "ES03":"UNIMED FARMACEUTICA HOLDI","GT01":"CDI, S.A. Guatemala","NI01":"CDI, S.A. Nicaragua","PA01":"PHARMARKETING SA"
}

_arc_cfg = _CONFIG.get('arc', {}).get('alertas', {})
UMBRAL_PAGO_ALTO  = _arc_cfg.get('a1_condicion_pago', {}).get('umbral_alto_dias',  90)
UMBRAL_PAGO_MEDIO = _arc_cfg.get('a1_condicion_pago', {}).get('umbral_medio_dias', 60)
UMBRAL_RECIENTE   = _arc_cfg.get('a4_empresa_reciente', {}).get('umbral_meses',    24)
UMBRAL_REP_DUP    = _arc_cfg.get('a3_representante_duplicado', {}).get('umbral_empresas', 2)

# ══════════════════════════════════════════════════════════════════════════════
# PASO 1 — LECTURA Y NORMALIZACIÓN DEL EXCEL
# ══════════════════════════════════════════════════════════════════════════════
def leer_excel():
    print(f"📊 Leyendo archivo Excel: {ARCHIVO_EXCEL.name}")
    if not ARCHIVO_EXCEL.exists(): raise FileNotFoundError(f"\n❌ Archivo no encontrado: {ARCHIVO_EXCEL}")
    xl = pd.ExcelFile(ARCHIVO_EXCEL); hojas = xl.sheet_names
    df_raw = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA_PRINCIPAL, dtype=str, header=None)
    header_row = 0
    for i in range(min(5, len(df_raw))):
        row_vals = [str(v).strip() for v in df_raw.iloc[i] if pd.notna(v)]
        if any(k in row_vals for k in ['Cliente', 'Razón Social', 'Razon Social', 'CLIENTE']):
            header_row = i; break
    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA_PRINCIPAL, dtype=str, header=header_row).fillna('')
    df = df[df.iloc[:, 0].str.strip() != ''].reset_index(drop=True)

    def leer_hoja_ref(nombre_hoja):
        if nombre_hoja not in hojas: return pd.DataFrame()
        df_r = pd.read_excel(ARCHIVO_EXCEL, sheet_name=nombre_hoja, dtype=str, header=None)
        hr = 0
        for i in range(min(5, len(df_r))):
            non_null = [str(v).strip() for v in df_r.iloc[i] if pd.notna(v)]
            if len(non_null) >= 2 and not any('·' in v or '—' in v for v in non_null): hr = i; break
        return pd.read_excel(ARCHIVO_EXCEL, sheet_name=nombre_hoja, dtype=str, header=hr).fillna('')

    ciiu_ref = {}; df_ciiu = leer_hoja_ref(HOJA_CIIU)
    if not df_ciiu.empty:
        for _, row in df_ciiu.iterrows():
            codigo = str(row.get('Código CIIU', row.get('codigo', ''))).strip()
            nivel  = str(row.get('Nivel de Riesgo', row.get('nivel', 'ALTO'))).strip().upper()
            if codigo and codigo != 'Código CIIU': ciiu_ref[codigo] = nivel

    # ── Lectura de países: compatible con la hoja EVALUACION DE RIESGO
    # de la Matriz V6.0 (header en fila 6, col 2 = País, col 11 = Estatus)
    # y también con el formato anterior (columnas 'País' / 'Nivel de Riesgo').
    paises_ref = {}
    if HOJA_PAISES in hojas:
        df_raw_p = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA_PAISES,
                                  header=None, dtype=str)
        # Detectar si es formato Matriz V6.0:
        # la fila 6 (índice 6) tiene 'País ' y 'Estatus' como encabezados
        es_matriz_v6 = False
        if len(df_raw_p) > 7:
            fila6 = [str(v).strip() for v in df_raw_p.iloc[6]]
            if 'País' in fila6 and 'Estatus' in fila6:
                es_matriz_v6 = True

        if es_matriz_v6:
            # Formato Matriz V6.0: datos desde fila 7 (índice 7)
            # col 2 (índice 2) = País, col 11 (índice 11) = Estatus
            ESTATUS_MAP = {'ALTO': 'ALTO', 'MEDIO': 'MEDIO', 'BAJO': 'BAJO'}
            for _, row in df_raw_p.iloc[7:].iterrows():
                pais   = str(row.iloc[2]).strip()  if len(row) > 2  else ''
                estatus = str(row.iloc[11]).strip() if len(row) > 11 else ''
                if not pais or pais in ('', 'nan', 'País', 'País '): continue
                nivel = ESTATUS_MAP.get(estatus.upper(), None)
                if nivel in ('ALTO', 'MEDIO'):
                    paises_ref[pais] = nivel
        else:
            # Formato anterior: columnas con nombre 'País' / 'Nivel de Riesgo'
            df_paises = leer_hoja_ref(HOJA_PAISES)
            for _, row in df_paises.iterrows():
                pais  = str(row.get('País', row.get('Pais', ''))).strip()
                nivel = str(row.get('Nivel de Riesgo', row.get('nivel', ''))).strip().upper()
                if pais and pais != 'País' and nivel in ('ALTO', 'MEDIO'):
                    paises_ref[pais] = nivel

    print(f"   ✅ {len(paises_ref)} países de riesgo cargados "
          f"(ALTO: {sum(1 for v in paises_ref.values() if v=='ALTO')}, "
          f"MEDIO: {sum(1 for v in paises_ref.values() if v=='MEDIO')})")

    return df, ciiu_ref, paises_ref

def normalizar_columnas(df):
    MAPA = {
        'cliente':      ['Cliente','Código','Cod. Cliente','Customer','CLIENTE'],
        'razon_social': ['Razón Social','Razon Social','Nombre','Name','RAZON SOCIAL'],
        'pais':         ['País','Pais','Country','Clave de país','PAIS'],
        'sociedad':     ['Sociedad','Society','Soc.','SOCIEDAD'],
        'nombre_soc':   ['Nombre Sociedad','Nombre Emp.','Empresa'],
        'pais_soc':     ['País Sociedad','Pais Sociedad'],
        'region':       ['Región','Region','Departamento'],
        'cod_pago':     ['Cód. Cond. Pago','Cod. Pago','Condición Pago'],
        'desc_pago':    ['Desc. Cond. Pago','Condición de Pago','Cond. Pago','Payment Terms'],
        'dias_pago':    ['Días Plazo','Dias Plazo','Días','Days','DIAS PLAZO'],
        'ciiu':         ['Ramo CIIU','Ramo','CIIU','Industry','RAMO'],
        'desc_ciiu':    ['Descripción Ramo','Desc. Ramo','Actividad','DESC RAMO'],
        'representante':['Representante Legal','Representante','RL','Legal Rep'],
        'fecha_creacion':['Fecha Creación','Fecha Creacion','Fecha Constitución','Creation Date'],
        'a1_pre':       ['A1: Cond. Pago','A1'], 'a2_pre': ['A2: Coherencia CIIU','A2'],
        'a3_pre':       ['A3: Rep. Duplicado','A3'], 'a4_pre': ['A4: Empresa Reciente (<2 años)','A4: Empresa Reciente','A4'],
        'a5_pre':       ['A5: Riesgo País','A5'], 'n_alertas_pre': ['# Alertas','Num Alertas'], 'riesgo_pre': ['RIESGO GLOBAL','Riesgo Global','Riesgo'],
    }
    cols_actuales = {c.strip(): c for c in df.columns}; rename_map = {}
    for campo, variantes in MAPA.items():
        for v in variantes:
            if v in cols_actuales: rename_map[cols_actuales[v]] = campo; break
    df = df.rename(columns=rename_map)
    for col in ['cliente','razon_social','pais','sociedad','dias_pago','ciiu','representante','fecha_creacion']:
        if col not in df.columns: df[col] = ''
    return df

# ══════════════════════════════════════════════════════════════════════════════
# PASO 2 — MOTOR DE ALERTAS A1-A5
# ══════════════════════════════════════════════════════════════════════════════
def extraer_dias(val):
    if pd.isna(val) or val == '': return 0
    nums = re.findall(r'\d+', str(val))
    return int(nums[0]) if nums else 0

def calcular_alertas(df, ciiu_ref, paises_ref):
    if all(c in df.columns for c in ['a1_pre','a2_pre','a3_pre','a4_pre','a5_pre','riesgo_pre']):
        df['a1'] = df['a1_pre'].str.upper().str.strip().fillna('BAJO')
        df['a2'] = df['a2_pre'].str.upper().str.strip().fillna('BAJO')
        df['a3'] = df['a3_pre'].str.upper().str.strip().fillna('BAJO')
        df['a4'] = df['a4_pre'].str.upper().str.strip().fillna('BAJO')
        df['a5'] = df['a5_pre'].str.upper().str.strip().fillna('BAJO')
        df['riesgo'] = df['riesgo_pre'].str.upper().str.strip().fillna('BAJO')
        df['n_alertas'] = ((df['a1'] != 'BAJO').astype(int) + (df['a2'] != 'BAJO').astype(int) + (df['a3'] != 'BAJO').astype(int) + (df['a4'] != 'BAJO').astype(int) + (df['a5'] != 'BAJO').astype(int))
        return df

    df['dias_num'] = df['dias_pago'].apply(extraer_dias)
    df['a1'] = df['dias_num'].apply(lambda d: 'ALTO' if d > UMBRAL_PAGO_ALTO else ('MEDIO' if d > UMBRAL_PAGO_MEDIO else 'BAJO'))
    df['ciiu_str'] = df['ciiu'].astype(str).str.strip()
    df['a2'] = df['ciiu_str'].apply(lambda c: ciiu_ref.get(c, 'ALTO'))
    rep_counts = df['representante'].str.strip().value_counts()
    df['a3'] = df['representante'].str.strip().apply(lambda r: 'ALTO' if (r and rep_counts.get(r, 0) >= UMBRAL_REP_DUP) else 'BAJO')

    hoy = date.today()
    def calcular_a4(fecha_str):
        if not fecha_str: return 'BAJO'
        try:
            if isinstance(fecha_str, str):
                for fmt in ['%Y-%m-%d','%d/%m/%Y','%Y/%m/%d','%d-%m-%Y']:
                    try:
                        fecha = datetime.strptime(fecha_str[:10], fmt).date(); break
                    except ValueError: continue
                else: return 'BAJO'
            else: fecha = fecha_str
            meses = (hoy.year - fecha.year) * 12 + (hoy.month - fecha.month)
            return 'ALTO' if meses < UMBRAL_RECIENTE else 'BAJO'
        except Exception: return 'BAJO'
    df['a4'] = df['fecha_creacion'].apply(calcular_a4)
    df['a5'] = df['pais'].str.strip().apply(lambda p: paises_ref.get(p, 'BAJO'))

    def riesgo_global(row):
        alertas = [row['a1'], row['a2'], row['a3'], row['a4'], row['a5']]
        n_alto  = alertas.count('ALTO'); n_medio = alertas.count('MEDIO')
        if n_alto >= 1 or n_medio >= 2: return 'ALTO'
        elif n_medio == 1: return 'MEDIO'
        return 'BAJO'

    df['riesgo']    = df.apply(riesgo_global, axis=1)
    df['n_alertas'] = ((df['a1'] != 'BAJO').astype(int) + (df['a2'] != 'BAJO').astype(int) + (df['a3'] != 'BAJO').astype(int) + (df['a4'] != 'BAJO').astype(int) + (df['a5'] != 'BAJO').astype(int))
    return df

# ══════════════════════════════════════════════════════════════════════════════
# PASO 3 — CONSTRUCCIÓN DEL BUNDLE JSON COMPACTO
# ══════════════════════════════════════════════════════════════════════════════
def construir_bundle(df):
    NL = {'ALTO': 1, 'MEDIO': 2, 'BAJO': 3}
    total = len(df); alto = (df['riesgo'] == 'ALTO').sum(); medio = (df['riesgo'] == 'MEDIO').sum(); bajo = (df['riesgo'] == 'BAJO').sum()

    pa_list  = sorted(df['pais'].dropna().unique().tolist())
    so_list  = sorted(df['sociedad'].dropna().unique().tolist())
    rz_list  = sorted(df['razon_social'].dropna().unique().tolist())
    dd_list  = sorted(df.get('desc_ciiu', pd.Series(dtype=str)).dropna().unique().tolist())
    nd_list  = sorted(df.get('nombre_soc', pd.Series(dtype=str)).dropna().unique().tolist())

    pa_map = {v: i for i, v in enumerate(pa_list)}; so_map = {v: i for i, v in enumerate(so_list)}
    rz_map = {v: i for i, v in enumerate(rz_list)}; dd_map = {v: i for i, v in enumerate(dd_list)}
    nd_map = {v: i for i, v in enumerate(nd_list)}

    def alert_counts(df_sub):
        if df_sub.empty: return [0, 0, 0, 0, 0]
        return [int((df_sub['a1'] != 'BAJO').sum()), int((df_sub['a2'] != 'BAJO').sum()), int((df_sub['a3'] != 'BAJO').sum()), int((df_sub['a4'] != 'BAJO').sum()), int((df_sub['a5'] != 'BAJO').sum())]

    ac = [alert_counts(df[df['riesgo'] == 'ALTO']), alert_counts(df[df['riesgo'] == 'MEDIO']), alert_counts(df[df['riesgo'] == 'BAJO'])]

    sa = []
    for so, grp in df.groupby('sociedad'): sa.append([so, SOCIEDADES_NOMBRES.get(so, grp['nombre_soc'].iloc[0] if 'nombre_soc' in grp else so)[:30], int((grp['riesgo'] == 'ALTO').sum()), int((grp['riesgo'] == 'MEDIO').sum()), int((grp['riesgo'] == 'BAJO').sum()), len(grp)])
    sa.sort(key=lambda x: -x[5])

    paa = []
    for pa, grp in df.groupby('pais'): paa.append([pa, int((grp['riesgo'] == 'ALTO').sum()), int((grp['riesgo'] == 'MEDIO').sum()), int((grp['riesgo'] == 'BAJO').sum()), len(grp)])
    paa.sort(key=lambda x: -x[1]); paa = paa[:15]

    rep_counts_map = df['representante'].str.strip().value_counts()
    dup_reps = rep_counts_map[rep_counts_map > 1].index.tolist()
    ra = []
    for rep in dup_reps[:15]:
        grp = df[df['representante'].str.strip() == rep]
        ra.append([rep[:45], len(grp), int((grp['riesgo'] == 'ALTO').sum()), int((grp['riesgo'] == 'MEDIO').sum()), int((grp['riesgo'] == 'BAJO').sum())])
    ra.sort(key=lambda x: -x[1])

    ca = []
    for ciiu, grp in df.groupby('ciiu'): ca.append([ciiu, grp['desc_ciiu'].iloc[0][:55] if 'desc_ciiu' in grp else ciiu, int((grp['riesgo'] == 'ALTO').sum()), int((grp['riesgo'] == 'MEDIO').sum()), int((grp['riesgo'] == 'BAJO').sum())])
    ca.sort(key=lambda x: -x[2]); ca = ca[:10]

    combo_map = {}
    for _, row in df[df['riesgo'] == 'ALTO'].iterrows():
        key = ''.join(['1' if row[f'a{i}'] != 'BAJO' else '0' for i in range(1, 6)])
        combo_map[key] = combo_map.get(key, 0) + 1
    cb = sorted([[k, v] for k, v in combo_map.items()], key=lambda x: -x[1])[:8]

    df_embed = pd.concat([df[df['riesgo'] == 'ALTO'].sort_values('n_alertas', ascending=False).head(300), df[df['riesgo'] == 'MEDIO'], df[df['riesgo'] == 'BAJO'].head(50)])

    rn_used = sorted(set(df_embed['representante'].dropna().str.strip().tolist()) | set(r[0] for r in ra))
    rn_map  = {n: i for i, n in enumerate(rn_used)}

    def encode_row(row):
        return [
            str(row.get('cliente', '')), rz_map.get(str(row.get('razon_social', '')), 0), pa_map.get(str(row.get('pais', '')), 0),
            so_map.get(str(row.get('sociedad', '')), 0), int(row.get('dias_num', 0)) if 'dias_num' in row.index else extraer_dias(row.get('dias_pago', '')),
            str(row.get('ciiu', '')), rn_map.get(str(row.get('representante', '')).strip(), 0), str(row.get('fecha_creacion', ''))[:7],
            NL.get(row.get('a1', 'BAJO'), 3), NL.get(row.get('a2', 'BAJO'), 3), NL.get(row.get('a3', 'BAJO'), 3), NL.get(row.get('a4', 'BAJO'), 3),
            NL.get(row.get('a5', 'BAJO'), 3), int(row.get('n_alertas', 0)), NL.get(row.get('riesgo', 'BAJO'), 3),
            dd_map.get(str(row.get('desc_ciiu', '')), 0), nd_map.get(str(row.get('nombre_soc', '')), 0)
        ]

    tc = [encode_row(row) for _, row in df_embed.iterrows()]
    ra_final = []
    for i, r in enumerate(ra): ra_final.append([rn_map.get(r[0], 0), r[0], r[1], r[2], r[3], r[4]])

    bundle = {
        "meta": {"total": total, "alto": int(alto), "medio": int(medio), "bajo": int(bajo), "embedded": len(tc), "generated": date.today().isoformat()},
        "pa": pa_list, "so": so_list, "rn": rn_used, "rz": rz_list, "dd": dd_list, "nd": nd_list, "rc": {1: int(alto), 2: int(medio), 3: int(bajo)},
        "ac": ac, "sa": sa, "paa": paa, "ra": ra_final, "ca": ca, "cb": cb, "tc": tc
    }
    bundle_json = json.dumps(bundle, separators=(',', ':'), ensure_ascii=False)
    return bundle, bundle_json, df

# ══════════════════════════════════════════════════════════════════════════════
# PASO 4 — BASE DE DATOS SQLITE
# ══════════════════════════════════════════════════════════════════════════════
def inicializar_db():
    DB_FILE.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_FILE, timeout=10) as conn:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS historico_clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT, cliente TEXT NOT NULL, razon_social TEXT, sociedad TEXT,
                riesgo TEXT, n_alertas INTEGER, decision TEXT DEFAULT 'Pendiente', observacion TEXT DEFAULT '', fecha_registro TEXT
            )
        """)
        conn.commit()

def guardar_decision(cliente, razon_social, sociedad, riesgo, n_alertas, decision, observacion):
    if not cliente or not decision or decision.lower() == 'pendiente': return
    with sqlite3.connect(DB_FILE, timeout=10) as conn:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("INSERT INTO historico_clientes (cliente, razon_social, sociedad, riesgo, n_alertas, decision, observacion, fecha_registro) VALUES (?,?,?,?,?,?,?,?)", 
                     (cliente, razon_social, sociedad, riesgo, int(n_alertas or 0), decision.capitalize(), observacion, datetime.now().isoformat()))
        conn.commit()

def obtener_historico_cliente(cliente):
    try:
        with sqlite3.connect(DB_FILE, timeout=10) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            rows = conn.execute("SELECT decision, observacion, fecha_registro, riesgo FROM historico_clientes WHERE cliente=? AND decision IN ('Confirmado','Descartado') ORDER BY fecha_registro DESC LIMIT 1", (cliente,)).fetchall()
        return [{'decision': r[0], 'observacion': r[1], 'fecha': r[2][:16].replace('T', ' ') if r[2] else '', 'riesgo': r[3]} for r in rows]
    except Exception: return []

# ══════════════════════════════════════════════════════════════════════════════
# PASO 5 — EXPORTACIÓN EXCEL (ESTANDARIZADA SMT)
# ══════════════════════════════════════════════════════════════════════════════
def autoajustar_columnas(ws, df_ref, fmt_default=None, max_width=50):
    for idx, col in enumerate(df_ref.columns):
        try:
            max_len = max(df_ref[col].astype(str).map(len).max() if len(df_ref) > 0 else 10, len(str(col).replace('_', ' '))) + 2
            ws.set_column(idx, idx, min(max_len, max_width), fmt_default)
        except Exception: ws.set_column(idx, idx, 18, fmt_default)

def exportar_expediente_individual(cliente, df_full, decision='', observacion=''):
    df_t = df_full[df_full['cliente'] == cliente].copy()
    if df_t.empty: raise ValueError(f"Cliente '{cliente}' no encontrado.")

    row = df_t.iloc[0]
    riesgo = row.get('riesgo', 'BAJO')
    nivel_f = NIVEL_FOLDER_MAP.get(riesgo, '3. Riesgo Bajo')
    mes_act = datetime.now().month
    año_act = datetime.now().year
    carpeta = CARPETA_RAIZ / str(año_act) / MESES_ES[mes_act - 1] / nivel_f
    carpeta.mkdir(parents=True, exist_ok=True)

    nombre_limpio = "".join(c for c in str(row.get('razon_social', cliente)) if c.isalnum() or c in " _-")[:40].strip()
    timestamp = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    ruta = carpeta / f"Expediente_ARC_{nombre_limpio}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    decision_cap = decision.capitalize() if decision else 'Pendiente'

    with pd.ExcelWriter(ruta, engine='xlsxwriter') as writer:
        wb = writer.book
        fmt_title   = wb.add_format({'bold': True, 'font_size': 14, 'font_color': '#1a3a6b'})
        fmt_meta    = wb.add_format({'italic': True, 'font_color': '#7a93ad', 'font_size': 9})
        fmt_sub     = wb.add_format({'bold': True, 'font_color': '#1a3a6b'})
        fmt_section = wb.add_format({'bold': True, 'font_color': '#1a3a6b', 'font_size': 12})
        fmt_hdr     = wb.add_format({'bold': True, 'bg_color': '#1a3a6b', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_data    = wb.add_format({'border': 1})
        fmt_center  = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_num     = wb.add_format({'border': 1, 'num_format': '#,##0', 'align': 'center'})
        fmt_alto    = wb.add_format({'bg_color': '#fdeef2', 'font_color': '#e8365d', 'bold': True, 'border': 1, 'align': 'center'})
        fmt_medio   = wb.add_format({'bg_color': '#fef6e7', 'font_color': '#d4860a', 'bold': True, 'border': 1, 'align': 'center'})
        fmt_bajo    = wb.add_format({'bg_color': '#e8f8f7', 'font_color': '#00a99d', 'bold': True, 'border': 1, 'align': 'center'})
        def fn(n): return {'ALTO': fmt_alto, 'MEDIO': fmt_medio}.get(str(n).upper(), fmt_bajo)

        # ── HOJA 1: Ficha del Cliente ──────────────────────────────────────────
        ws = wb.add_worksheet('Ficha del Cliente')
        ws.freeze_panes(3, 0)
        ws.write('A1', f'Expediente SAGRILAFT (ARC): {row.get("razon_social", cliente)}', fmt_title)
        ws.write('A2', f'Generado: {timestamp}  |  SM  |  Sofgen Pharma', fmt_meta)

        info = [
            ('Código SAP',           row.get('cliente', '')),
            ('Razón Social',         row.get('razon_social', '')),
            ('País',                 row.get('pais', '')),
            ('Sociedad',             row.get('sociedad', '') + ' — ' + SOCIEDADES_NOMBRES.get(str(row.get('sociedad', '')), '')),
            ('CIIU',                 str(row.get('ciiu', '')) + ' — ' + str(row.get('desc_ciiu', ''))),
            ('Representante Legal',  row.get('representante', '')),
            ('Fecha Constitución',   str(row.get('fecha_creacion', ''))[:10]),
            ('Condición de Pago',    f"{row.get('dias_num', 0)}d — {row.get('desc_pago', '')}"),
            ('Región',               row.get('region', '')),
            ('Riesgo Global',        riesgo),
            ('Nº Alertas Activas',   int(row.get('n_alertas', 0))),
            ('Estado de Decisión',   decision_cap),
            ('Observación',          observacion if observacion else '(Sin observación)'),
        ]
        for r, (label, val) in enumerate(info, 4):
            ws.write(r, 0, label, fmt_sub)
            if label == 'Riesgo Global':
                ws.write(r, 1, str(val), fn(str(val)))
            elif label == 'Nº Alertas Activas':
                ws.write_number(r, 1, int(val), fmt_num)
            else:
                ws.write(r, 1, str(val), fmt_center)

        # Resumen de Alertas (tabla en Hoja 1 — igual que SMT)
        fila_sec = len(info) + 6
        ws.write(fila_sec, 0, 'Resumen de Alertas SAGRILAFT', fmt_section)
        for c, h in enumerate(['Alerta', 'Nivel Detectado', 'Detalle Forense', 'Resultado']):
            ws.write(fila_sec + 1, c, h, fmt_hdr)

        alertas_resumen = [
            ('A1: Condición de Pago',      row.get('a1', 'BAJO'), f"Días plazo detectados: {row.get('dias_num', 0)}d",             'ALTO' if row.get('a1') == 'ALTO' else ('MEDIO' if row.get('a1') == 'MEDIO' else 'Sin alerta')),
            ('A2: Coherencia CIIU',        row.get('a2', 'BAJO'), f"CIIU: {row.get('ciiu', '')} — {str(row.get('desc_ciiu', ''))[:55]}", 'ALTO' if row.get('a2') == 'ALTO' else 'Sin alerta'),
            ('A3: Representante Duplicado',row.get('a3', 'BAJO'), f"RL: {row.get('representante', '')}",                           'ALTO' if row.get('a3') == 'ALTO' else 'Sin alerta'),
            ('A4: Empresa Reciente',       row.get('a4', 'BAJO'), f"Constitución: {str(row.get('fecha_creacion', ''))[:10]}",      'ALTO' if row.get('a4') == 'ALTO' else 'Sin alerta'),
            ('A5: Riesgo País',            row.get('a5', 'BAJO'), f"País: {row.get('pais', '')}",                                 'ALTO' if row.get('a5') == 'ALTO' else 'Sin alerta'),
        ]
        for i, (alerta, nivel, detalle, resultado) in enumerate(alertas_resumen):
            ws.write(fila_sec + 2 + i, 0, alerta,    fmt_data)
            ws.write(fila_sec + 2 + i, 1, nivel,     fn(nivel))
            ws.write(fila_sec + 2 + i, 2, detalle,   fmt_data)
            ws.write(fila_sec + 2 + i, 3, resultado,
                     fmt_alto if resultado == 'ALTO' else (fmt_medio if resultado == 'MEDIO' else fmt_bajo))

        ws.set_column('A:A', 30)
        ws.set_column('B:B', 55)
        ws.set_column('C:C', 22)
        ws.set_column('D:D', 15)

        # ── HOJA 2: Desglose de Alertas ───────────────────────────────────────
        ws2 = wb.add_worksheet('Desglose de Alertas')
        ws2.freeze_panes(3, 0)
        ws2.write('A1', 'Detalle Técnico de Alertas A1–A5', fmt_title)
        ws2.write('A2', f'Generado: {timestamp}  |  SM  |  Sofgen Pharma', fmt_meta)
        for c, h in enumerate(['Alerta', 'Nivel', 'Detalle Forense']):
            ws2.write(2, c, h, fmt_hdr)

        alertas_detalle = [
            ('A1: Condición de Pago',       row.get('a1', 'BAJO'), f"Días plazo detectados: {row.get('dias_num', 0)}d — {row.get('desc_pago', '')}"),
            ('A2: Coherencia CIIU',         row.get('a2', 'BAJO'), f"CIIU: {row.get('ciiu', '')} — {str(row.get('desc_ciiu', ''))[:80]}"),
            ('A3: Representante Duplicado', row.get('a3', 'BAJO'), f"Representante Legal: {row.get('representante', '')}"),
            ('A4: Empresa Reciente',        row.get('a4', 'BAJO'), f"Fecha Constitución: {str(row.get('fecha_creacion', ''))[:10]}"),
            ('A5: Riesgo País',             row.get('a5', 'BAJO'), f"País: {row.get('pais', '')}"),
        ]
        for r, (nombre, nivel, detalle) in enumerate(alertas_detalle, 3):
            ws2.write(r, 0, nombre,  fmt_data)
            ws2.write(r, 1, nivel,   fn(nivel))
            ws2.write(r, 2, detalle, fmt_data)

        ws2.set_column('A:A', 30)
        ws2.set_column('B:B', 15)
        ws2.set_column('C:C', 80)

    return str(ruta)

def exportar_reporte_maestro(tipo, df_full, riesgo_filtro=None, decisiones=None):
    if decisiones is None: decisiones = {}
    df_exp = df_full.copy()
    if riesgo_filtro and riesgo_filtro != 'all': df_exp = df_exp[df_exp['riesgo'] == riesgo_filtro.upper()]
    if df_exp.empty: raise ValueError("Sin datos para exportar con los filtros aplicados.")
    mes_act = datetime.now().month; año_act = datetime.now().year; timestamp = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    if tipo == 'mensual':
        nivel_dom = df_exp['riesgo'].mode()[0] if not df_exp.empty else 'BAJO'
        nivel_f   = NIVEL_FOLDER_MAP.get(nivel_dom, '3. Riesgo Bajo')
        carpeta   = CARPETA_RAIZ / str(año_act) / MESES_ES[mes_act - 1] / nivel_f
        nombre_arch = f"Maestro_{MESES_NOMBRE[mes_act-1]}_{año_act}_Nivel_{nivel_dom}.xlsx"
    else:
        carpeta   = CARPETA_RAIZ / str(año_act)
        nombre_arch = f"{año_act}_Consolidado_Maestro_SAGRILAFT.xlsx"
    carpeta.mkdir(parents=True, exist_ok=True); ruta = carpeta / nombre_arch

    columnas_export = ['cliente', 'razon_social', 'pais', 'sociedad', 'nombre_soc', 'region', 'dias_pago', 'ciiu', 'desc_ciiu', 'representante', 'fecha_creacion', 'a1', 'a2', 'a3', 'a4', 'a5', 'n_alertas', 'riesgo']
    columnas_exist  = [c for c in columnas_export if c in df_exp.columns]
    df_out = df_exp[columnas_exist].copy()
    df_out['Decision Analista'] = df_out['cliente'].apply(lambda c: decisiones.get(c, {}).get('decision', 'Pendiente'))
    df_out['Observacion']       = df_out['cliente'].apply(lambda c: decisiones.get(c, {}).get('observacion', ''))

    with pd.ExcelWriter(ruta, engine='xlsxwriter') as writer:
        wb = writer.book
        fmt_title = wb.add_format({'bold': True, 'font_size': 14, 'font_color': '#1a3a6b'})
        fmt_meta  = wb.add_format({'italic': True, 'font_color': '#7a93ad', 'font_size': 9})
        fmt_hdr   = wb.add_format({'bold': True, 'bg_color': '#1a3a6b', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_data  = wb.add_format({'border': 1})
        fmt_num   = wb.add_format({'border': 1, 'num_format': '#,##0', 'align': 'center'})
        fmt_alto  = wb.add_format({'bg_color': '#fdeef2', 'font_color': '#e8365d', 'bold': True, 'border': 1, 'align': 'center'})
        fmt_medio = wb.add_format({'bg_color': '#fef6e7', 'font_color': '#d4860a', 'bold': True, 'border': 1, 'align': 'center'})
        fmt_bajo  = wb.add_format({'bg_color': '#e8f8f7', 'font_color': '#00a99d', 'bold': True, 'border': 1, 'align': 'center'})

        ws = wb.add_worksheet('Maestro Clientes')
        ws.freeze_panes(3, 0)
        ws.write('A1', f'SM · Reporte de Clientes {tipo.capitalize()} · Sofgen Pharma', fmt_title)
        ws.write('A2', f'Generado: {timestamp}  |  SM  |  Sofgen Pharma', fmt_meta)

        headers = [c.replace('_', ' ').title() for c in df_out.columns]
        for c, h in enumerate(headers): ws.write(2, c, h, fmt_hdr)

        for r, (_, row_data) in enumerate(df_out.iterrows(), 3):
            for c, (col, val) in enumerate(zip(df_out.columns, row_data)):
                if col in ('a1', 'a2', 'a3', 'a4', 'a5', 'riesgo'):
                    nivel = str(val).upper()
                    fmt_n = fmt_alto if nivel == 'ALTO' else (fmt_medio if nivel == 'MEDIO' else fmt_bajo)
                    ws.write(r, c, val, fmt_n)
                elif col == 'n_alertas' and isinstance(val, (int, float)) and not pd.isna(val):
                    ws.write_number(r, c, int(val), fmt_num)
                elif isinstance(val, (int, float)) and not pd.isna(val):
                    ws.write_number(r, c, val, fmt_data)
                else:
                    ws.write(r, c, str(val) if val is not None else '', fmt_data)
        autoajustar_columnas(ws, df_out)
    print(f"   ✅ Reporte {tipo} exportado: {ruta}")
    return str(ruta)

# ══════════════════════════════════════════════════════════════════════════════
# PASO 6 — SERVIDOR HTTP (NUEVAS Y VIEJAS RUTAS PARA CERO FALLOS)
# ══════════════════════════════════════════════════════════════════════════════
class ThreadedHTTPServer(ThreadingMixIn, HTTPServer): daemon_threads = True

class MaestroHandler(SimpleHTTPRequestHandler):
    def _send_json(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(code); self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*'); self.send_header('Content-Length', len(body))
        self.end_headers(); self.wfile.write(body)

    def _send_text(self, code, text):
        body = text.encode('utf-8')
        self.send_response(code); self.send_header('Content-Type', 'text/plain; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*'); self.send_header('Content-Length', len(body))
        self.end_headers(); self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(200, "ok"); self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS'); self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_POST(self):
        if self.path == '/export_arc':
            content_length = int(self.headers['Content-Length'])
            try:
                data = json.loads(self.rfile.read(content_length).decode('utf-8'))
                cliente_id = data.get('id'); observacion = data.get('obs', ''); estado = data.get('est', 'Pendiente')
                
                df_t = DF_GLOBAL[DF_GLOBAL['cliente'] == cliente_id].copy()
                if df_t.empty:
                    self._send_json(404, {"status":"error","msg":"Cliente no encontrado"}); return
                
                r = df_t.iloc[0]
                guardar_decision(cliente_id, r.get('razon_social', ''), r.get('sociedad', ''), r.get('riesgo', ''), r.get('n_alertas', 0), estado, observacion)
                exportar_expediente_individual(cliente_id, DF_GLOBAL, estado, observacion)
                
                self._send_json(200, {"status":"ok"})
            except Exception as e:
                self._send_json(500, {"status":"error", "msg": str(e)})

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path); path = parsed.path; query = urllib.parse.parse_qs(parsed.query)

        if path in ('/', '/dashboard_maestro.html', '/dashboard-2.html'):
            if HTML_GLOBAL:
                self.send_response(200); self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.send_header('Access-Control-Allow-Origin', '*'); self.send_header('Content-Length', len(HTML_GLOBAL))
                self.end_headers(); self.wfile.write(HTML_GLOBAL)
            else: self._send_text(503, "Dashboard aún no disponible.")
        elif path == '/bundle': self._send_json(200, BUNDLE_GLOBAL)
        elif path == '/historico': self._send_json(200, obtener_historico_cliente(query.get('cliente', [''])[0]))
        elif path == '/stats':
            try:
                with sqlite3.connect(DB_FILE, timeout=10) as conn:
                    conf = conn.execute("SELECT COUNT(DISTINCT cliente) FROM historico_clientes WHERE decision='Confirmado'").fetchone()[0]
                    desc = conn.execute("SELECT COUNT(DISTINCT cliente) FROM historico_clientes WHERE decision='Descartado'").fetchone()[0]
                self._send_json(200, {'confirmados': conf, 'descartados': desc})
            except Exception as e: self._send_text(500, str(e))
        elif path == '/logo.png':
            try:
                with open(LOGO_FILE, 'rb') as f: img = f.read()
                self.send_response(200); self.send_header('Content-Type', 'image/png'); self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers(); self.wfile.write(img)
            except Exception: self._send_text(404, "logo.png no encontrado")
        elif path == '/logo_sm.png':
            try:
                with open(SM_LOGO_FILE, 'rb') as f: img = f.read()
                self.send_response(200); self.send_header('Content-Type', 'image/png'); self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers(); self.wfile.write(img)
            except Exception: self._send_text(404, "logo_sm.png no encontrado")

        # 👇 RUTA RESTAURADA PARA COMPATIBILIDAD CON CACHÉ Y REPORTES MASIVOS 👇
        elif path == '/exportar':
            tipo      = query.get('tipo', [''])[0]
            cliente   = query.get('cliente', [''])[0]
            obs       = urllib.parse.unquote(query.get('obs', [''])[0])
            decision  = query.get('decision', ['Pendiente'])[0]
            riesgo_f  = query.get('riesgo', ['all'])[0]

            try:
                if cliente and decision and decision.lower() != 'pendiente':
                    row = DF_GLOBAL[DF_GLOBAL['cliente'] == cliente]
                    if not row.empty:
                        r = row.iloc[0]
                        guardar_decision(cliente, r.get('razon_social', ''), r.get('sociedad', ''), r.get('riesgo', ''), r.get('n_alertas', 0), decision, obs)
            except Exception as e_sq: print(f"SQLite: {e_sq}")

            try:
                if tipo == 'individual' and cliente:
                    exportar_expediente_individual(cliente, DF_GLOBAL, decision, obs)
                elif tipo in ('mensual', 'consolidado'):
                    exportar_reporte_maestro(tipo, DF_GLOBAL, riesgo_f)
                self._send_text(200, "OK")
            except Exception as e:
                print(f"Error exportación: {e}"); self._send_text(500, str(e))
        else: self._send_text(404, "Ruta no encontrada")

    def log_message(self, format, *args): pass

# ══════════════════════════════════════════════════════════════════════════════
# PASO 7 — GENERACIÓN DEL DASHBOARD HTML AUTOCONTENIDO
# ══════════════════════════════════════════════════════════════════════════════
def generar_dashboard_autocontenido(bundle_json, ruta_html):
    import re as _re
    print("🌐 Construyendo dashboard con datos reales...")
    nombres_posibles = ["dashboard-2.html", "dashboard_maestro_template.html", "SM_Maestro_Clientes.html", "dashboard_maestro.html"]
    html_fuente = None
    for nombre in nombres_posibles:
        candidato = SM_DIR / nombre
        if candidato.exists() and candidato != ruta_html:
            html_fuente = candidato; print(f"   📄 Fuente encontrada: {nombre}"); break

    if html_fuente is None:
        if DASHBOARD_HTML.exists(): html_fuente = DASHBOARD_HTML
        else: raise FileNotFoundError(f"\n❌ No se encontró el archivo HTML fuente.")

    with open(html_fuente, 'r', encoding='utf-8') as f: html = f.read()
    nuevo_b = f"var B={bundle_json};"
    html_nuevo, n = _re.subn(r'var\s+B\s*=\s*\{[\s\S]*?\};(?=\s*var\s+NL)', nuevo_b, html, count=1)
    if n == 0: html_nuevo, n = _re.subn(r'var\s+B\s*=\s*\{[\s\S]*?\};', nuevo_b, html, count=1)
    
    with open(ruta_html, 'w', encoding='utf-8') as f: f.write(html_nuevo)
    return html_nuevo.encode('utf-8')

# ══════════════════════════════════════════════════════════════════════════════
# PASO 8 — INICIALIZACIÓN DE CARPETAS Y ARRANQUE
# ══════════════════════════════════════════════════════════════════════════════
def crear_estructura_carpetas():
    print("📁 Creando estructura de carpetas SAGRILAFT...")
    año_act = datetime.now().year; mes_act = datetime.now().month
    for mes in range(1, mes_act + 1):
        for nivel_f in NIVEL_FOLDER_MAP.values():
            (CARPETA_RAIZ / str(año_act) / MESES_ES[mes - 1] / nivel_f).mkdir(parents=True, exist_ok=True)
    (CARPETA_RAIZ / str(año_act)).mkdir(parents=True, exist_ok=True)

BUNDLE_GLOBAL = {}; DF_GLOBAL = pd.DataFrame(); HTML_GLOBAL = b''

if __name__ == '__main__':
    print("=" * 60)
    print("  SM v2.0 · Sistema 2 — Maestro de Clientes")
    print("  Sofgen Pharma · Cumplimiento Legal Corporativo")
    print("=" * 60)

    df_raw, ciiu_ref, paises_ref = leer_excel()
    df_norm = normalizar_columnas(df_raw)
    df_final = calcular_alertas(df_norm, ciiu_ref, paises_ref)
    DF_GLOBAL = df_final

    BUNDLE_GLOBAL, bundle_json, _ = construir_bundle(df_final)

    inicializar_db()
    crear_estructura_carpetas()
    HTML_GLOBAL = generar_dashboard_autocontenido(bundle_json, DASHBOARD_HTML)

    def iniciar_servidor():
        server = ThreadedHTTPServer(('0.0.0.0', PUERTO), MaestroHandler)
        server.serve_forever()

    threading.Thread(target=iniciar_servidor, daemon=True).start()

    print("\n" + "=" * 60)
    print("  🚀 SM SISTEMA 2 ACTIVADO")
    print("=" * 60)
    print(f"  Dashboard (PC):     http://localhost:{PUERTO}/\n")
    
    if not os.environ.get('SM_UNIFIED'): webbrowser.open(f'http://localhost:{PUERTO}/')

    try:
        while True: pass
    except KeyboardInterrupt:
        print("\n🛑 Apagando SM Sistema 2 de manera segura...")
