import base64
import hmac
import html
import json
from datetime import datetime
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill

CARPETA = Path(__file__).parent
ICONO = CARPETA / "icono.png"

st.set_page_config(
    page_title="Evaluación de la Oportunidad de Formación",
    page_icon=str(ICONO) if ICONO.exists() else "📝",
    layout="centered",
)

ZONA = ZoneInfo("America/Caracas")
TITULO = "Evaluación de la Oportunidad de Formación"
ESCALA = ["5 Excelente", "4 Bueno", "3 Satisfactorio", "2 Regular", "1 Deficiente"]
COLUMNAS = {
    "encuestas": ["id", "tipo", "evento", "director_genero", "director",
                  "fechas", "facilitadores", "activa", "creada"],
    "respuestas": ["id_encuesta", "fecha_hora", "nombre", "impacto", "director",
                   "facilitadores", "administrador", "ambiente", "comentario"],
}

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@500;700;800&display=swap');
#MainMenu, footer {visibility: hidden;}
.block-container {max-width: 760px; padding-top: 2.5rem;}
h1, h2, h3 {font-family: 'Montserrat', sans-serif; color: #622599;}
.cabecera {background: #fff; border-radius: 18px; overflow: hidden; border: 1px solid #E6DCF0;
  box-shadow: 0 10px 30px rgba(98, 37, 153, .10); margin-bottom: 1.5rem;}
.cabecera-logo {padding: 1.3rem 1rem .9rem; text-align: center;}
.cabecera-logo img {max-height: 96px; max-width: 70%;}
.cabecera-banda {background: linear-gradient(135deg, #622599 0%, #4A1A78 100%); color: #fff;
  text-align: center; padding: 1.6rem 1.5rem 1.7rem;}
.cabecera-etiqueta {font-family: 'Montserrat', sans-serif; font-size: .72rem; letter-spacing: .22em;
  text-transform: uppercase; opacity: .85;}
.cabecera-titulo {font-family: 'Montserrat', sans-serif; font-weight: 800; font-size: 1.6rem;
  line-height: 1.2; margin-top: .55rem;}
.cabecera-linea {width: 56px; height: 3px; background: #fff; opacity: .6; margin: .85rem auto;
  border-radius: 2px;}
.cabecera-detalle {font-size: .98rem; opacity: .95; line-height: 1.5;}
[data-testid="stForm"] {background: #fff; border: 1px solid #E6DCF0; border-radius: 18px;
  padding: 1.8rem 1.8rem 1.4rem; box-shadow: 0 10px 30px rgba(98, 37, 153, .06);}
[data-testid="stWidgetLabel"] p {font-weight: 600; color: #2B1B3D;}
.stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] > button {
  border-radius: 999px; font-weight: 700; letter-spacing: .04em; padding: .55rem 1.6rem;}
.escala {background: #F3ECF9; border-left: 4px solid #622599; border-radius: 8px;
  padding: .8rem 1rem; margin: .4rem 0 1rem; font-size: .93rem; color: #2B1B3D;}
.aviso {background: #fff; border: 1px solid #E6DCF0; border-radius: 18px; text-align: center;
  padding: 2.2rem 1.5rem; box-shadow: 0 10px 30px rgba(98, 37, 153, .06);}
.aviso-titulo {font-family: 'Montserrat', sans-serif; font-weight: 800; font-size: 1.4rem;
  color: #622599; margin-bottom: .4rem;}
</style>
"""


# ---------------------------------------------------------------- Diseño
@st.cache_data
def imagen_base64(nombre):
    ruta = CARPETA / nombre
    if not ruta.exists():
        return ""
    return base64.b64encode(ruta.read_bytes()).decode()


def encabezado(etiqueta, titulo, detalle=""):
    logo = imagen_base64("logo.png")
    partes = ['<div class="cabecera">']
    if logo:
        partes.append(
            f'<div class="cabecera-logo"><img src="data:image/png;base64,{logo}" alt="Logo"></div>'
        )
    partes.append('<div class="cabecera-banda">')
    partes.append(f'<div class="cabecera-etiqueta">{html.escape(etiqueta)}</div>')
    partes.append(f'<div class="cabecera-titulo">{html.escape(titulo)}</div>')
    if detalle:
        partes.append('<div class="cabecera-linea"></div>')
        partes.append(f'<div class="cabecera-detalle">{detalle}</div>')
    partes.append("</div></div>")
    st.markdown("".join(partes), unsafe_allow_html=True)


def aviso(titulo, texto):
    st.markdown(
        f'<div class="aviso"><div class="aviso-titulo">{html.escape(titulo)}</div>'
        f"<div>{html.escape(texto)}</div></div>",
        unsafe_allow_html=True,
    )


def detalle_evento(e):
    return (
        f"{html.escape(str(e['director_genero']))}: <b>{html.escape(str(e['director']))}</b>"
        f"<br>{html.escape(str(e['fechas']))}"
    )


# ---------------------------------------------------------------- Google Sheets
@st.cache_resource
def abrir_libro():
    cliente = gspread.service_account_from_dict(json.loads(st.secrets["gcp_json"]))
    return cliente.open_by_key(st.secrets["sheet_id"])


def hoja(nombre):
    libro = abrir_libro()
    try:
        return libro.worksheet(nombre)
    except gspread.exceptions.WorksheetNotFound:
        try:
            ws = libro.add_worksheet(title=nombre, rows=1000, cols=len(COLUMNAS[nombre]))
            ws.append_row(COLUMNAS[nombre])
            return ws
        except gspread.exceptions.APIError:
            return libro.worksheet(nombre)


@st.cache_data(ttl=60)
def leer(nombre):
    registros = hoja(nombre).get_all_records()
    return pd.DataFrame(registros, columns=COLUMNAS[nombre])


def ahora():
    return datetime.now(ZONA).strftime("%d/%m/%Y %H:%M")


def buscar_encuesta(id_txt):
    encuestas = leer("encuestas")
    fila = encuestas[encuestas["id"].astype(str) == str(id_txt)]
    return None if fila.empty else fila.iloc[0]


def respuestas_de(e):
    respuestas = leer("respuestas")
    return respuestas[respuestas["id_encuesta"].astype(str) == str(e["id"])]


# ---------------------------------------------------------------- Textos y enlaces
def texto_director(e):
    return "de la Directora" if e["director_genero"] == "Directora" else "del Director"


def texto_facilitador(f):
    articulo = "de la facilitadora" if f["genero"] == "facilitadora" else "del facilitador"
    return f"{articulo} {f['nombre']}"


def direccion_app():
    return st.secrets.get("app_url", "").rstrip("/")


def enlace(id_encuesta):
    return f"{direccion_app()}/?encuesta={id_encuesta}"


def clave_reporte(id_encuesta):
    secreto = st.secrets["admin_password"].encode()
    return hmac.new(secreto, f"reporte-{id_encuesta}".encode(), "sha256").hexdigest()[:20]


def enlace_reporte(id_encuesta):
    return f"{direccion_app()}/?reporte={id_encuesta}&clave={clave_reporte(id_encuesta)}"


# ---------------------------------------------------------------- Encuesta pública
def pagina_encuesta(id_txt):
    e = buscar_encuesta(id_txt)
    if e is None:
        encabezado("Oportunidad de formación", TITULO)
        st.error("Esta encuesta no existe. Verifica el enlace.")
        return
    tipo = e["tipo"]
    facilitadores = json.loads(e["facilitadores"])
    encabezado(TITULO, str(e["evento"]).upper(), detalle_evento(e))

    if str(e["activa"]).upper() != "SI":
        aviso("Encuesta cerrada", "Esta evaluación ya no recibe respuestas. ¡Gracias por tu interés!")
        return
    if st.session_state.get(f"enviada_{id_txt}"):
        aviso("¡Gracias!", "Tu evaluación fue registrada. Siempre listos para servir.")
        return

    with st.form("encuesta"):
        nombre = st.text_input("Nombre y Apellido *")
        impacto = st.text_area(f"¿Qué fue lo que más te impactó del contenido del {tipo}? *")
        st.markdown(
            '<div class="escala"><b>Haga uso de la escala de estimación abajo indicada</b><br>'
            "Excelente = 5 · Bueno = 4 · Satisfactorio = 3 · Regular = 2 · Deficiente = 1</div>",
            unsafe_allow_html=True,
        )
        p_director = st.radio(
            f"¿Cómo fue el desempeño {texto_director(e)}? *", ESCALA, index=None
        )
        p_facilitadores = {}
        for i, f in enumerate(facilitadores):
            p_facilitadores[f["nombre"]] = st.radio(
                f"¿Cómo fue el desempeño {texto_facilitador(f)}? *",
                ESCALA, index=None, key=f"fac_{i}",
            )
        p_admin = st.radio("¿Cómo fue el desempeño del Administrador? *", ESCALA, index=None)
        p_ambiente = st.radio(
            f"¿Cómo le pareció el ambiente que predominó en el {tipo}? *", ESCALA, index=None
        )
        comentario = st.text_area("Comentario u Observaciones finales")
        enviar = st.form_submit_button("Enviar evaluación", type="primary")

    if enviar:
        escalas = [p_director, p_admin, p_ambiente, *p_facilitadores.values()]
        if not nombre.strip() or not impacto.strip() or any(v is None for v in escalas):
            st.error("Por favor responde todas las preguntas marcadas con *.")
            return
        notas = {n: int(v[0]) for n, v in p_facilitadores.items()}
        hoja("respuestas").append_row([
            int(e["id"]), ahora(), nombre.strip(), impacto.strip(),
            int(p_director[0]), json.dumps(notas, ensure_ascii=False),
            int(p_admin[0]), int(p_ambiente[0]), comentario.strip(),
        ])
        leer.clear()
        st.session_state[f"enviada_{id_txt}"] = True
        st.rerun()


# ---------------------------------------------------------------- Reporte Excel
def generar_excel(e, respuestas):
    tipo = e["tipo"]
    facilitadores = json.loads(e["facilitadores"])
    col_impacto = f"Lo que más impactó del {tipo}"
    col_director = f"{e['director_genero']}: {e['director']}"
    cols_fac = {f["nombre"]: f"{f['genero'].capitalize()}: {f['nombre']}" for f in facilitadores}
    col_admin = "Administrador"
    col_ambiente = f"Ambiente del {tipo}"
    cols_escala = [col_director, *cols_fac.values(), col_admin, col_ambiente]
    columnas = ["Fecha y hora", "Nombre y Apellido", col_impacto, *cols_escala, "Comentario"]

    filas = []
    for _, r in respuestas.iterrows():
        notas = json.loads(r["facilitadores"]) if r["facilitadores"] else {}
        fila = {
            "Fecha y hora": r["fecha_hora"],
            "Nombre y Apellido": r["nombre"],
            col_impacto: r["impacto"],
            col_director: r["director"],
            col_admin: r["administrador"],
            col_ambiente: r["ambiente"],
            "Comentario": r["comentario"],
        }
        for nombre, columna in cols_fac.items():
            fila[columna] = notas.get(nombre)
        filas.append(fila)
    detalle = pd.DataFrame(filas, columns=columnas)

    resumen = []
    for c in cols_escala:
        s = pd.to_numeric(detalle[c], errors="coerce")
        resumen.append({
            "Aspecto evaluado": c,
            "Promedio": round(s.mean(), 2) if s.notna().any() else None,
            "5 Excelente": int((s == 5).sum()),
            "4 Bueno": int((s == 4).sum()),
            "3 Satisfactorio": int((s == 3).sum()),
            "2 Regular": int((s == 2).sum()),
            "1 Deficiente": int((s == 1).sum()),
        })
    resumen = pd.DataFrame(resumen)

    morado = PatternFill("solid", fgColor="622599")
    blanco = Font(bold=True, color="FFFFFF")
    archivo = BytesIO()
    with pd.ExcelWriter(archivo, engine="openpyxl") as w:
        resumen.to_excel(w, sheet_name="Resumen", index=False, startrow=5)
        detalle.to_excel(w, sheet_name="Respuestas", index=False)
        h = w.sheets["Resumen"]
        h["A1"] = TITULO
        h["A2"] = str(e["evento"]).upper()
        h["A3"] = f"{e['director_genero']}: {e['director']}"
        h["A4"] = str(e["fechas"])
        h["A5"] = f"Total de respuestas: {len(detalle)}"
        h["A1"].font = Font(bold=True, size=14, color="622599")
        h["A2"].font = Font(bold=True, size=12)
        for celda in list(h[6]) + list(w.sheets["Respuestas"][1]):
            celda.fill = morado
            celda.font = blanco
        for ws in w.sheets.values():
            for columna in ws.columns:
                ws.column_dimensions[columna[0].column_letter].width = 30
    return archivo.getvalue()


def boton_descarga(e, respuestas):
    st.download_button(
        "Descargar reporte en Excel",
        data=generar_excel(e, respuestas),
        file_name=f"Evaluacion_{e['id']}_{str(e['evento']).replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


# ---------------------------------------------------------------- Reporte para el Director
def pagina_reporte(id_txt, clave):
    e = buscar_encuesta(id_txt)
    valida = hmac.compare_digest(str(clave or "").encode(), clave_reporte(id_txt).encode())
    if e is None or not valida:
        encabezado("Oportunidad de formación", "Reporte de evaluación")
        st.error("Este enlace no es válido. Solicita uno nuevo al administrador.")
        return
    encabezado("Reporte de evaluación", str(e["evento"]).upper(), detalle_evento(e))

    if st.button("Actualizar datos"):
        leer.clear()
    respuestas = respuestas_de(e)
    activa = str(e["activa"]).upper() == "SI"

    c1, c2 = st.columns(2)
    c1.metric("Respuestas recibidas", len(respuestas))
    c2.metric("Estado de la encuesta", "Abierta" if activa else "Cerrada")
    boton_descarga(e, respuestas)
    st.caption("El reporte se genera con las respuestas recibidas hasta este momento.")


# ---------------------------------------------------------------- Administración
def acceso_admin():
    if st.session_state.get("admin"):
        return True
    clave = st.text_input("Clave de administrador", type="password")
    if st.button("Entrar", type="primary"):
        if hmac.compare_digest(clave.encode(), st.secrets["admin_password"].encode()):
            st.session_state["admin"] = True
            st.rerun()
        else:
            st.error("Clave incorrecta.")
    return False


def nueva_encuesta():
    cantidad = st.number_input("Cantidad de facilitadores", 0, 15, 2)
    with st.form("nueva"):
        tipo = st.selectbox("Tipo de evento", ["Taller", "Curso"])
        evento = st.text_input("Nombre del evento", placeholder="TALLER DE CONDUCCIÓN DE TROPA")
        c1, c2 = st.columns([1, 2])
        director_genero = c1.selectbox("Dirección", ["Director", "Directora"])
        director = c2.text_input("Nombre del Director o Directora")
        fechas = st.text_input("Fechas", placeholder="27 al 29 de noviembre de 2026")
        facilitadores = []
        for i in range(int(cantidad)):
            c1, c2 = st.columns([1, 2])
            genero = c1.selectbox(f"Facilitador(a) {i + 1}", ["facilitador", "facilitadora"], key=f"g{i}")
            nombre = c2.text_input("Nombre", key=f"n{i}")
            facilitadores.append({"genero": genero, "nombre": nombre.strip()})
        crear = st.form_submit_button("Crear encuesta", type="primary")

    if crear:
        nombres = [f["nombre"] for f in facilitadores]
        if not evento.strip() or not director.strip() or not fechas.strip() or "" in nombres:
            st.error("Completa todos los campos.")
            return
        if len(set(nombres)) != len(nombres):
            st.error("Hay facilitadores con el mismo nombre.")
            return
        encuestas = leer("encuestas")
        nuevo_id = 1 if encuestas.empty else int(pd.to_numeric(encuestas["id"]).max()) + 1
        hoja("encuestas").append_row([
            nuevo_id, tipo, evento.strip(), director_genero, director.strip(),
            fechas.strip(), json.dumps(facilitadores, ensure_ascii=False), "SI", ahora(),
        ])
        leer.clear()
        st.success("Encuesta creada.")
        st.caption("Enlace para los participantes")
        st.code(enlace(nuevo_id), language=None)
        st.caption("Enlace del reporte para el Director o Directora")
        st.code(enlace_reporte(nuevo_id), language=None)


def lista_encuestas():
    if st.button("Actualizar datos"):
        leer.clear()
    encuestas = leer("encuestas")
    if encuestas.empty:
        st.info("Todavía no hay encuestas. Crea la primera en la otra pestaña.")
        return
    encuestas = encuestas.iloc[::-1]
    opciones = {f"{r['id']} · {r['evento']} ({r['fechas']})": i for i, r in encuestas.iterrows()}
    elegida = st.selectbox("Encuesta", list(opciones))
    indice = opciones[elegida]
    e = encuestas.loc[indice]

    respuestas = respuestas_de(e)
    activa = str(e["activa"]).upper() == "SI"

    st.caption("Enlace para los participantes")
    st.code(enlace(e["id"]), language=None)
    st.caption("Enlace del reporte para el Director o Directora")
    st.code(enlace_reporte(e["id"]), language=None)

    c1, c2 = st.columns(2)
    c1.metric("Respuestas", len(respuestas))
    c2.metric("Estado", "Abierta" if activa else "Cerrada")

    c1, c2 = st.columns(2)
    with c1:
        boton_descarga(e, respuestas)
    if c2.button("Cerrar encuesta" if activa else "Reabrir encuesta"):
        columna = COLUMNAS["encuestas"].index("activa") + 1
        hoja("encuestas").update_cell(int(indice) + 2, columna, "NO" if activa else "SI")
        leer.clear()
        st.rerun()


def pagina_admin():
    encabezado("Oportunidades de formación", "Administración de encuestas")
    if not acceso_admin():
        return
    t1, t2 = st.tabs(["Encuestas y reportes", "Nueva encuesta"])
    with t1:
        lista_encuestas()
    with t2:
        nueva_encuesta()


# ---------------------------------------------------------------- Inicio
st.markdown(CSS, unsafe_allow_html=True)
id_encuesta = st.query_params.get("encuesta")
id_reporte = st.query_params.get("reporte")
if id_encuesta:
    pagina_encuesta(id_encuesta)
elif id_reporte:
    pagina_reporte(id_reporte, st.query_params.get("clave"))
else:
    pagina_admin()
