"""Dashboard de avance de campaña en Streamlit (un solo archivo).

Lee tu base YA limpia (Query_Preventas) tal como viene de MySQL: usa sus columnas Coordinador, Supervisor,
ESTADO.1 (o ESTADO_ACTIVACION), Linea Adicional neta/bruta, CARGO_FIJO_ACTIVACION, PLAN_COMERCIAL y DERIVACION_CANAL.

Carpeta del proyecto:  dashboard.py + Query_Preventas.xls|xlsx|csv
Instalar:              pip install streamlit pandas numpy openpyxl xlrd xlsxwriter
Ejecutar:              python -m streamlit run dashboard.py
"""
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# =============================== CONFIGURACIÓN (edita aquí cada mes) ===============================
CFG = {
    "titulo": "Avance de campaña · Portabilidad Out Masivo",
    "servicio_preventa": "PORTABILIDAD",
    "fecha_corte": None,                          # solo para el subtítulo; None = última fecha de la base
    "dias_mes": 23,                               # días laborables del mes
    "conversion_objetivo": 0.70,                  # objetivo bruto = objetivo neto / 0.70
    "plan_rentable_min": 69.90,                   # SUPUESTO: plan rentable = CF de activación >= este valor
    "cf_tope": 159.90,                            # excepciones = activas con CF por encima de este valor
    "coordinadores": {                            # objetivo neto, asesores objetivo y objetivo diario por coordinador
        "Baleryn Camacho": {"obj_neta": 656, "as_obj": 30, "obj_diario": 50},
        "Jerry Vilchez": {"obj_neta": 656, "as_obj": 30, "obj_diario": 50},
        "Juan Arismendiz": {"obj_neta": 2501, "as_obj": 10, "obj_diario": 20},
    },
    "fuera": "Fuera de campaña",
}
ESTADOS = {"VENTA ACTIVA": "Activo", "VENTA BRUTA": "En Proceso", "VENTA PERDIDA": "Caida", "VENTA ENTREGADA": "Entregado"}


# =============================== CÁLCULOS ===============================
def sdiv(a, b):
    return a / b.where(b != 0) if isinstance(b, pd.Series) else (a / b if b else np.nan)


def derivar(g, cfg, dias):
    g["obj_bruta"] = sdiv(g["obj_neta"], pd.Series(cfg["conversion_objetivo"], index=g.index))
    g["pendiente_bruta"] = g["obj_bruta"] - g["brutas"]
    g["avance_brutas"] = sdiv(g["brutas"], g["obj_bruta"])
    g["pro_cf_bruto"] = sdiv(g["cf_sum_b"], g["brutas"])
    g["conversion"] = sdiv(g["activas"], g["brutas"])
    g["pct_perdida"] = sdiv(g["perdida"], g["brutas"])
    g["proyeccion"] = g["activas"] / dias * cfg["dias_mes"]
    g["pct_avance"] = sdiv(g["activas"], g["obj_neta"])
    g["pct_proy"] = sdiv(g["proyeccion"], g["obj_neta"])
    g["pendiente_obj"] = g["obj_neta"] - g["activas"]
    for pref, num in [("b", "brutas"), ("n", "activas")]:
        g[f"spd_{pref}_obj"] = sdiv(g[num], g["as_obj"]) / dias
        g[f"spd_{pref}_con"] = sdiv(g[num], g["q_as"]) / dias
    g["pct_la"] = sdiv(g["lineas_adic"], g["activas"])
    g["pct_pr"] = sdiv(g["plan_rentable"], g["activas"])
    g["pro_cf_neto"] = sdiv(g["cf_sum_n"], g["activas"])
    g.loc[g["obj_neta"] == 0, ["pct_avance", "pct_proy", "pendiente_obj", "avance_brutas", "pendiente_bruta", "obj_bruta"]] = np.nan
    return g

def calcular(f, cfg, dias):
    """Una fila por coordinador + total. Todo sale de las columnas de la base tal como vienen."""
    act = f["estado"] == "Activo"
    cf = f["CARGO_FIJO_ACTIVACION"].fillna(f["CARGO_FIJO"])
    t = pd.DataFrame({
        "grupo": f["grupo"], "brutas": 1, "activas": act.astype(int),
        "entregado": (f["estado"] == "Entregado").astype(int), "perdida": (f["estado"] == "Caida").astype(int),
        "en_proceso": (f["estado"] == "En Proceso").astype(int),
        "cf_sum_b": f["CARGO_FIJO"].fillna(0), "cf_sum_n": np.where(act, cf.fillna(0), 0.0),
        "lineas_adic": ((f["Linea Adicional neta"] == "Linea Adicional") & act).astype(int),
        "lineas_adic_bruta": (f["LINEA ADICIONAL bruta"] == "Linea Adicional").astype(int),
        "plan_rentable": (act & (cf >= cfg["plan_rentable_min"])).astype(int),
        "plan_9590": (act & f["PLAN_COMERCIAL"].astype("string").str.contains("95.90", na=False)).astype(int),
        "cf_exc": (act & (cf > cfg["cf_tope"])).astype(int),
    })
    g = t.groupby("grupo").sum().reindex(list(cfg["coordinadores"]) + [cfg["fuera"]], fill_value=0)
    for k in ["obj_neta", "as_obj", "obj_diario"]:
        g[k] = [cfg["coordinadores"].get(c, {}).get(k, 0) for c in g.index]
    q = f.groupby("grupo")["usuario"].nunique()               # asesores con al menos una venta
    g["q_as"] = [q.get(c, 0) for c in g.index]
    g.loc["Total general"] = g.sum()
    return derivar(g, cfg, dias)


# =============================== DASHBOARD HTML (mismo diseño del dashboard de avance) ===============================
# =============================== DASHBOARD HTML ===============================
def n(x, d=0):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "–"
    return f"{x:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pc(x, d=1):
    return "–" if x is None or pd.isna(x) else n(x * 100, d) + "%"


def tabla(titulo, cols, g, extra=""):
    cab = "".join(f"<th>{c}</th>" for c, _, _ in cols)
    filas = ""
    for i, (nombre, r) in enumerate(g.iterrows()):
        cls = ' class="tot"' if nombre == "Total general" else ""
        filas += f"<tr{cls}><td>{nombre}</td>" + "".join(f"<td>{fn(r[k])}</td>" for _, k, fn in cols) + "</tr>"
    return (f'<div class="card" style="overflow-x:auto;{extra}"><p class="lbl">{titulo}</p>'
            f'<table class="t"><tr><th>Coordinador</th>{cab}</tr>{filas}</table></div>')


def barras(items, total):
    h = ""
    for etq, v, col in items:
        w = 100 * v / total if total else 0
        h += (f'<div>{etq} · {n(v)}<div style="height:8px;margin-top:6px;border-radius:4px;background:{col};'
              f'width:{w:.1f}%"></div></div>')
    return h


def grafico_diario(f):
    d = f.groupby(f["FECHA_EVALUACION"].dt.date).agg(b=("_b", "sum"), a=("_n", "sum"))
    if d.empty:
        return ""
    mx, k = d["b"].max(), len(d)
    paso = 600 / k
    s = ""
    for i, (fecha, r) in enumerate(d.iterrows()):
        x = i * paso + 2
        w = max(paso / 2 - 3, 2)
        hb, ha = 170 * r["b"] / mx, 170 * r["a"] / mx
        s += (f'<rect x="{x:.1f}" y="{180 - hb:.1f}" width="{w:.1f}" height="{hb:.1f}" fill="#B8C4CC"/>'
              f'<rect x="{x + w + 1:.1f}" y="{180 - ha:.1f}" width="{w:.1f}" height="{ha:.1f}" fill="#0F766E"/>')
        if k <= 16 or i % 2 == 0:
            s += f'<text x="{x + w:.1f}" y="198" font-size="10" text-anchor="middle" fill="#5B6B76">{fecha.strftime("%d/%m")}</text>'
    return f'<svg viewBox="0 0 600 205" style="width:100%;height:auto" role="img" aria-label="Ventas diarias brutas y activas">{s}</svg>'


CSS = """body{margin:0}.card{background:#fff;border:1px solid #E1E6EA;border-radius:12px;padding:20px;min-width:0}
.lbl{font-size:13px;color:#5B6B76;margin:0 0 12px}.big{font-size:30px;font-weight:700;margin:6px 0 4px}.sub{font-size:13px;color:#5B6B76}
.t{width:100%;border-collapse:collapse;font-size:13px}.t th,.t td{padding:10px 8px;text-align:right;white-space:nowrap}
.t th:first-child,.t td:first-child{text-align:left}.t th{color:#5B6B76;font-weight:500}.t tr+tr td{border-top:1px solid #E1E6EA}
.t .tot td{border-top:2px solid #12202B;font-weight:700}
.grid{display:grid;gap:20px;grid-template-columns:repeat(auto-fit,minmax(320px,1fr))}.k{grid-template-columns:repeat(auto-fit,minmax(240px,1fr))}
.wrap{max-width:1280px;margin:0 auto;padding:32px;display:flex;flex-direction:column;gap:20px}"""


def kpi(l, big, sub, w=None):
    barra = (f'<div style="height:8px;background:#E1E6EA;border-radius:4px"><div style="width:{min(w, 100):.1f}%;height:8px;'
             f'background:#0F766E;border-radius:4px"></div></div>') if w is not None else ""
    return f'<div class="card"><p class="lbl" style="margin:0">{l}</p><p class="big">{big}</p>{barra}<span class="sub">{sub}</span></div>'


def construir_html(f, g, cfg, dias):
    t = g.loc["Total general"]
    kp = "".join([
        kpi("Ventas netas (activas)", n(t.activas), f"{pc(t.pct_avance)} de {n(t.obj_neta)} · pendiente {n(t.pendiente_obj)}", 100 * (t.pct_avance or 0)),
        kpi("Ventas brutas", n(t.brutas), f"Avance brutas {pc(t.avance_brutas)} · obj. {n(t.obj_bruta)}"),
        kpi("Proyección lineal", n(t.proyeccion), f"{pc(t.pct_proy)} del objetivo"),
        kpi("Conversión", pc(t.conversion), f"Venta perdida: {n(t.perdida)} ({pc(t.pct_perdida)})"),
        kpi("Plan rentable", pc(t.pct_pr), f"{n(t.plan_rentable)} ventas · {n(t.plan_9590)} con plan 95,90"),
        kpi("Líneas adicionales", n(t.lineas_adic), f"{pc(t.pct_la)} · CF > {n(cfg['cf_tope'], 2)}: {n(t.cf_exc)}"),
        kpi("Asesores activos", n(t.q_as), f"Objetivo {n(t.as_obj)} · obj. diario {n(t.obj_diario)}"),
        kpi("Días trabajados", n(dias, 1), f"de {cfg['dias_mes']} días del mes", 100 * dias / cfg["dias_mes"]),
    ])
    estados = barras([("Activas", t.activas, "#0F766E"), ("En proceso", t.en_proceso, "#D9480F"),
                      ("Caídas", t.perdida, "#B42318"), ("Entregadas", t.entregado, "#5B6B76")], t.brutas)
    canal = f[f["_n"] == 1].groupby(f["DERIVACION_CANAL"].fillna("SIN CANAL")).size().sort_values(ascending=False)
    canales = barras([(k, v, "#0F766E") for k, v in canal.items()], canal.sum())
    I, P = lambda x: n(x), lambda x: pc(x)
    D1, D2 = lambda x: n(x, 1), lambda x: n(x, 2)
    t_brutas = tabla("Ventas brutas por coordinador", [
        ("Q ventas", "brutas", I), ("AS obj.", "as_obj", I), ("Q AS", "q_as", I), ("Obj. diario", "obj_diario", I),
        ("Obj. mes venta bruta", "obj_bruta", D1), ("Pendiente", "pendiente_bruta", D1),
        ("Pro. CF bruto", "pro_cf_bruto", D1), ("Avance brutas", "avance_brutas", P)], g)
    t_netas = tabla("Ventas netas y cumplimiento", [
        ("Obj. mes venta neta", "obj_neta", I), ("Activas", "activas", I), ("Entregado", "entregado", I),
        ("Proyección", "proyeccion", D1), ("% avance cumpl.", "pct_avance", P), ("% cumpl. proyectado", "pct_proy", P),
        ("Pendiente obj.", "pendiente_obj", I), ("Venta perdida", "perdida", I), ("% venta perdida", "pct_perdida", P),
        ("Conversión", "conversion", P)], g)
    t_prod = tabla("Productividad y calidad de venta", [
        ("SPD B obj.", "spd_b_obj", D2), ("SPD N obj.", "spd_n_obj", D2), ("SPD B conect.", "spd_b_con", D2),
        ("SPD N conect.", "spd_n_con", D2), ("Líneas adic. neta", "lineas_adic", I), ("% líneas adic.", "pct_la", P), ("Líneas adic. bruta", "lineas_adic_bruta", I),
        ("Plan rentable", "plan_rentable", I), ("% plan rentable", "pct_pr", P), ("Pro. CF neto", "pro_cf_neto", D2),
        ("Plan 95,90", "plan_9590", I), (f"CF > {n(cfg['cf_tope'], 2)}", "cf_exc", I)], g)
    t_kon = tabla("Proyección lineal · Konecta", [
        ("Obj. Konecta", "obj_neta", I), ("Proyección lineal", "proyeccion", D1),
        ("% avance cumpl.", "pct_avance", P), ("% cumpl. proyectado", "pct_proy", P)], g, "grid-column:span 2")
    corte = cfg["fecha_corte"] or f["FECHA_EVALUACION"].max().strftime("%Y-%m-%d")
    return f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{cfg['titulo']}</title><style>{CSS}body{{background:#F4F6F8;color:#12202B;font-family:'DM Sans',Arial,sans-serif}}</style></head>
<body><div class="wrap">
<div><p class="lbl" style="margin:0">{cfg['servicio_preventa'].title()} · corte {corte} · {n(dias, 1)} de {cfg['dias_mes']} días trabajados</p>
<h1 style="margin:2px 0 0;font-size:28px">{cfg['titulo']}</h1></div>
<div class="grid k">{kp}</div>
<div class="grid"><div class="card" style="grid-column:span 2"><p class="lbl">Ventas diarias: brutas (gris) y activas (verde)</p>{grafico_diario(f)}</div>
<div class="card"><p class="lbl">Estado de ventas brutas</p><div style="display:flex;flex-direction:column;gap:14px;font-size:14px">{estados}</div></div></div>
{t_brutas}{t_netas}{t_prod}
<div class="grid">{t_kon}<div class="card"><p class="lbl">Ventas netas por canal</p><div style="display:flex;flex-direction:column;gap:14px;font-size:14px">{canales}</div></div></div>
</div></body></html>"""


# =============================== STREAMLIT ===============================
st.set_page_config(page_title="Avance de campaña", page_icon="📊", layout="wide")


@st.cache_data(show_spinner="Leyendo la base...")
def leer(nombre: str, contenido: bytes | None):
    if contenido is not None:
        import io
        fuente = io.BytesIO(contenido)
        df = pd.read_csv(fuente) if nombre.lower().endswith(".csv") else pd.read_excel(fuente)
    else:
        ruta = next((r for ext in ("xlsx", "xls", "csv") if (r := Path(__file__).parent / f"Query_Preventas.{ext}").exists()), None)
        if ruta is None:
            return None
        df = pd.read_csv(ruta) if ruta.suffix == ".csv" else pd.read_excel(ruta)
    df.columns = df.columns.str.strip()
    df["FECHA_EVALUACION"] = pd.to_datetime(df["FECHA_EVALUACION"])
    for c in ["CARGO_FIJO", "CARGO_FIJO_ACTIVACION"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["estado"] = df["ESTADO.1"] if "ESTADO.1" in df else df["ESTADO_ACTIVACION"].map(ESTADOS)
    coord = df["Coordinador"].astype("string").str.strip()
    df["coordinador"] = coord.where(coord.str.upper() != "FUERA DE CAMPAÑA", CFG["fuera"])
    df["grupo"] = df["coordinador"].where(df["coordinador"].isin(CFG["coordinadores"]), CFG["fuera"])
    df["supervisor"] = df["Supervisor"].astype("string").str.strip().fillna("Sin supervisor")
    df["usuario"] = df["E_REAL_FINAL"] if "E_REAL_FINAL" in df else df["VENTA_USUARIO"]
    return df


subido = st.sidebar.file_uploader("Carga Query_Preventas (opcional)", type=["xls", "xlsx", "csv"])
df = leer(subido.name, subido.getvalue()) if subido else leer("", None)
if df is None:
    st.error("No encuentro Query_Preventas.xls/.xlsx/.csv junto a dashboard.py. Cárgalo desde la barra lateral.")
    st.stop()

st.sidebar.header("Filtros")
quitar_dup = st.sidebar.checkbox("Quitar filas duplicadas", value=True)
if quitar_dup:
    cols = [c for c in df.columns if c not in ("ESTADO.1", "Linea Adicional neta", "LINEA ADICIONAL bruta")]
    df = df.drop_duplicates(subset=cols)

periodos = sorted(df["TIPO_PERIODO"].dropna().unique())
servicios = sorted(df["SERVICIO"].dropna().unique())
grupos = [x for x in list(CFG["coordinadores"]) + [CFG["fuera"]] if x in set(df["grupo"])]
sups = sorted(df["supervisor"].unique())
estados = ["Activo", "En Proceso", "Entregado", "Caida"]
sel_p = st.sidebar.multiselect("Periodo:", periodos, default=[p for p in periodos if p == "MES ACTUAL"] or periodos)
sel_srv = st.sidebar.multiselect("Servicio:", servicios, default=[s for s in servicios if s == CFG["servicio_preventa"]] or servicios)
sel_g = st.sidebar.multiselect("Coordinador:", grupos, default=grupos)
sel_s = st.sidebar.multiselect("Supervisor:", sups, default=sups)
sel_e = st.sidebar.multiselect("Estado:", estados, default=estados)

f = df[df["TIPO_PERIODO"].isin(sel_p) & df["SERVICIO"].isin(sel_srv) & df["grupo"].isin(sel_g)
       & df["supervisor"].isin(sel_s) & df["estado"].isin(sel_e)]
if f.empty:
    st.warning("Los filtros no dejan ninguna venta.")
    st.stop()
f = f.assign(_b=1, _n=(f["estado"] == "Activo").astype(int))

st.sidebar.header("Parámetros")
dias_mes = st.sidebar.number_input("Días del mes", 1, 31, int(CFG["dias_mes"]))
dias = st.sidebar.number_input("Días trabajados", 0.5, 31.0, float(f["FECHA_EVALUACION"].dt.date.nunique()), step=0.5)
cfg = {**CFG, "dias_mes": dias_mes}

g = calcular(f, cfg, dias)
sin_doc = df["NRO DOCUMENTO CLIENTE"].isna().mean() if "NRO DOCUMENTO CLIENTE" in df else 0
if sin_doc > 0.10:
    st.warning(f"{sin_doc:.0%} de las ventas no tiene documento: las líneas adicionales (vienen de tu SQL) pueden estar subestimadas.")

html = construir_html(f, g, cfg, dias)
components.html(html, height=2300, scrolling=True)

with st.expander("Ver base filtrada"):
    st.dataframe(f.drop(columns=["_b", "_n"]))

out = BytesIO()
with pd.ExcelWriter(out, engine="xlsxwriter") as w:
    f.drop(columns=["_b", "_n"]).to_excel(w, index=False, sheet_name="Data_Filtrada")
    g.drop(columns=["cf_sum_b", "cf_sum_n"]).reset_index(names="coordinador").to_excel(w, index=False, sheet_name="Resumen")
st.sidebar.download_button("📥 Descargar Excel filtrado", out.getvalue(), "Reporte_Preventas_Filtrado.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.sidebar.download_button("📥 Descargar dashboard (HTML)", html.encode("utf-8"), "dashboard_avance.html", mime="text/html")