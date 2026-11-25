import hmac
import json
from datetime import datetime
from io import BytesIO
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Evaluación de la Oportunidad de Formación",
    page_icon="📝",
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


# ---------------------------------------------------------------- Google Sheets
@st.cache_resource
def abrir_libro():
    cliente = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
    return cliente.open_by_key(st.secrets["sheet_id"])


def hoja(nombre):
    libro = abrir_libro()
    try:
        return libro.worksheet(nombre)
    except gspread.exceptions.WorksheetNotFound:
        ws = libro.add_worksheet(title=nombre, rows=1000, cols=len(COLUMNAS[nombre]))
        ws.append_row(COLUMNAS[nombre])
        return ws


@st.cache_data(ttl=60)
def leer(nombre):
    registros = hoja(nombre).get_all_records()
    return pd.DataFrame(registros, columns=COLUMNAS[nombre])


def ahora():
    return datetime.now(ZONA).strftime("%d/%m/%Y %H:%M")


# ---------------------------------------------------------------- Textos
def texto_director(e):
    return "de la Directora" if e["director_genero"] == "Directora" else "del Director"


def texto_facilitador(f):
    articulo = "de la facilitadora" if f["genero"] == "facilitadora" else "del facilitador"
    return f"{articulo} {f['nombre']}"


# ---------------------------------------------------------------- Encuesta pública
def pagina_encuesta(id_txt):
    encuestas = leer("encuestas")
    fila = encuestas[encuestas["id"].astype(str) == str(id_txt)]
    if fila.empty:
        st.error("Esta encuesta no existe. Verifica el enlace.")
        return
    e = fila.iloc[0]
    tipo = e["tipo"]
    facilitadores = json.loads(e["facilitadores"])

    st.title(TITULO)
    st.subheader(str(e["evento"]).upper())
    st.write(f"**{e['director_genero']}:** {e['director']}  \n{e['fechas']}")
    st.divider()

    if str(e["activa"]).upper() != "SI":
        st.info("Esta encuesta ya está cerrada. ¡Gracias por tu interés!")
        return
    if st.session_state.get(f"enviada_{id_txt}"):
        st.success("¡Gracias! Tu evaluación fue registrada.")
        return

    with st.form("encuesta"):
        nombre = st.text_input("Nombre y Apellido *")
        impacto = st.text_area(f"¿Qué fue lo que más te impactó del contenido del {tipo}? *")
        st.markdown(
            "***Haga uso de la escala de estimación abajo indicada***  \n"
            "**(Excelente = 5 / Bueno = 4 / Satisfactorio = 3 / Regular = 2 / Deficiente = 1)**"
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
        enviar = st.form_submit_button("Enviar", type="primary")

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
        for ws in w.sheets.values():
            for columna in ws.columns:
                ws.column_dimensions[columna[0].column_letter].width = 30
    return archivo.getvalue()


# ---------------------------------------------------------------- Administración
def acceso_admin():
    if st.session_state.get("admin"):
        return True
    st.title("Administración de encuestas")
    clave = st.text_input("Clave de administrador", type="password")
    if st.button("Entrar", type="primary"):
        if hmac.compare_digest(clave, st.secrets["admin_password"]):
            st.session_state["admin"] = True
            st.rerun()
        else:
            st.error("Clave incorrecta.")
    return False


def enlace(id_encuesta):
    base = st.secrets.get("app_url", "").rstrip("/")
    return f"{base}/?encuesta={id_encuesta}"


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
        st.success(f"Encuesta creada. Enlace para los participantes: {enlace(nuevo_id)}")


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

    respuestas = leer("respuestas")
    respuestas = respuestas[respuestas["id_encuesta"].astype(str) == str(e["id"])]
    activa = str(e["activa"]).upper() == "SI"

    st.write(f"**Enlace:** {enlace(e['id'])}")
    c1, c2 = st.columns(2)
    c1.metric("Respuestas", len(respuestas))
    c2.metric("Estado", "Abierta" if activa else "Cerrada")

    c1, c2 = st.columns(2)
    c1.download_button(
        "Descargar reporte en Excel",
        data=generar_excel(e, respuestas),
        file_name=f"Evaluacion_{e['id']}_{str(e['evento']).replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    if c2.button("Cerrar encuesta" if activa else "Reabrir encuesta"):
        columna = COLUMNAS["encuestas"].index("activa") + 1
        hoja("encuestas").update_cell(int(indice) + 2, columna, "NO" if activa else "SI")
        leer.clear()
        st.rerun()


def pagina_admin():
    if not acceso_admin():
        return
    st.title("Administración de encuestas")
    t1, t2 = st.tabs(["Encuestas y reportes", "Nueva encuesta"])
    with t1:
        lista_encuestas()
    with t2:
        nueva_encuesta()


# ---------------------------------------------------------------- Inicio
id_encuesta = st.query_params.get("encuesta")
if id_encuesta:
    pagina_encuesta(id_encuesta)
else:
    pagina_admin()
