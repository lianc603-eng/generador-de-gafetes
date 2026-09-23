import copy
import io
import os
import xml.etree.ElementTree as ET
import zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Gafetes", page_icon="🪪", layout="wide")

st.title("🪪 Generador de Gafetes Oficiales")
st.write(
    "Selecciona a los empleados requeridos para generar sus gafetes en formato **ODP** conservando exactamente el diseño original."
)

# --- 1. DETECCIÓN AUTOMÁTICA DE ARCHIVOS ---
def buscar_archivo_plantilla():
    # Busca variantes de nombre comunes para evitar errores de mayúsculas/minúsculas
    posibles_nombres = [
        "PLANTILLA MAESTRA.odp",
        "plantilla_maestra.odp",
        "Plantilla_Maestra.odp",
        "Plantilla Maestra.odp",
        "PLANTILLA_MAESTRA.odp",
    ]
    for nombre in posibles_nombres:
        if os.path.exists(nombre):
            return nombre
    # Si hay algún archivo con extensión .odp en la carpeta, toma ese
    for f in os.listdir("."):
        if f.lower().endswith(".odp"):
            return f
    return None

archivo_plantilla = buscar_archivo_plantilla()

# Cargar base de datos
@st.cache_data
def cargar_base():
    posibles_excel = ["base.xlsx", "BASE.xlsx", "Base.xlsx"]
    for f in posibles_excel:
        if os.path.exists(f):
            df = pd.read_excel(f)
            df.columns = [str(c).strip() for c in df.columns]
            return df
    # Si hay algún .xlsx
    for f in os.listdir("."):
        if f.lower().endswith(".xlsx") and not f.startswith("~$"):
            df = pd.read_excel(f)
            df.columns = [str(c).strip() for c in df.columns]
            return df
    return None

df = cargar_base()

# Validaciones iniciales
if df is None:
    st.error("❌ No se encontró el archivo 'base.xlsx' en el repositorio.")
    st.stop()

if archivo_plantilla is None:
    st.error("❌ No se encontró el archivo de la plantilla ODP en el repositorio.")
    st.stop()

# --- 2. DETECCIÓN DE COLUMNAS ---
col_num = next((c for c in df.columns if "NUM" in c.upper() or "EMPLEADO" in c.upper()), "NUMERO DE EMPLEADO")
col_nom = next((c for c in df.columns if "NOMBRE" in c.upper()), "Nombre completo")
col_pto = next((c for c in df.columns if "PUESTO" in c.upper()), "Puesto")

# --- 3. INTERFAZ DE USUARIO ---
col_chk, _ = st.columns([2, 2])
with col_chk:
    seleccionar_todos = st.checkbox("Seleccionar todos los empleados de la lista")

opciones = df[col_nom].dropna().tolist()

if seleccionar_todos:
    seleccionados = st.multiselect("Empleados a generar:", opciones, default=opciones)
else:
    seleccionados = st.multiselect("Empleados a generar:", opciones)

# --- 4. MOTOR DE GENERACIÓN ODP ---
def generar_odp(df_seleccionados, ruta_plantilla):
    with open(ruta_plantilla, "rb") as f:
        template_bytes = f.read()

    in_zip = zipfile.ZipFile(io.BytesIO(template_bytes), "r")
    out_buffer = io.BytesIO()
    out_zip = zipfile.ZipFile(out_buffer, "w", zipfile.ZIP_DEFLATED)

    # Copiar todos los archivos internos excepto content.xml
    for item in in_zip.infolist():
        if item.filename != "content.xml":
            out_zip.writestr(item, in_zip.read(item.filename))

    xml_text = in_zip.read("content.xml").decode("utf-8")

    # Mantener los namespaces originales
    namespaces = dict([node for _, node in ET.iterparse(io.StringIO(xml_text), events=["start-ns"])])
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)

    root = ET.fromstring(xml_text)
    draw_ns = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"

    body = root.find(".//{urn:oasis:names:tc:opendocument:xmlns:office:1.0}presentation")
    if body is None:
        body = root.find(".//{urn:oasis:names:tc:opendocument:xmlns:office:1.0}body/*")

    paginas = body.findall(f"{{{draw_ns}}}page")
    if not paginas:
        raise ValueError("No se encontraron páginas en la plantilla.")

    pagina_maestra = paginas[0]

    # Limpiar las páginas base para colocar las generadas
    for p in paginas:
        body.remove(p)

    def reemplazar_en_nodo(elem, mapa):
        if elem.text:
            for k, v in mapa.items():
                if k in elem.text:
                    elem.text = elem.text.replace(k, str(v))
        if elem.tail:
            for k, v in mapa.items():
                if k in elem.tail:
                    elem.tail = elem.tail.replace(k, str(v))
        for hijo in elem:
            reemplazar_en_nodo(hijo, mapa)

    filas = list(df_seleccionados.iterrows())
    pagina_idx = 1

    # Procesar de 2 en 2 por hoja Carta (gafete arriba y gafete abajo)
    for i in range(0, len(filas), 2):
        emp1 = filas[i][1]
        emp2 = filas[i + 1][1] if i + 1 < len(filas) else None

        nueva_pagina = copy.deepcopy(pagina_maestra)
        nueva_pagina.set(f"{{{draw_ns}}}name", f"Hoja_{pagina_idx}")
        pagina_idx += 1

        elementos = [e for e in nueva_pagina if not e.tag.endswith("notes")]
        mitad = len(elementos) // 2
        gafete_arriba = elementos[:mitad]
        gafete_abajo = elementos[mitad:]

        # Mapeo primer empleado (arriba)
        nom1 = str(emp1.get(col_nom, "")).strip()
        num1 = str(emp1.get(col_num, "")).strip()
        pto1 = str(emp1.get(col_pto, "")).strip() if pd.notna(emp1.get(col_pto)) else ""
        folio1 = f"{i + 1:03d}"

        mapa_g1 = {
            "*Citlalli Estefanía Brown Ocaña*": nom1,
            "*Citlalli": nom1,
            "Ocaña*": "",
            "*ANALISTA*": pto1,
            "*9820*": num1,
            "*041*": folio1,
        }
        for el in gafete_arriba:
            reemplazar_en_nodo(el, mapa_g1)

        # Mapeo segundo empleado (abajo)
        if emp2 is not None:
            nom2 = str(emp2.get(col_nom, "")).strip()
            num2 = str(emp2.get(col_num, "")).strip()
            pto2 = str(emp2.get(col_pto, "")).strip() if pd.notna(emp2.get(col_pto)) else ""
            folio2 = f"{i + 2:03d}"

            mapa_g2 = {
                "*Citlalli Estefanía Brown Ocaña*": nom2,
                "*Citlalli": nom2,
                "Ocaña*": "",
                "*ANALISTA*": pto2,
                "*9820*": num2,
                "*041*": folio2,
            }
            for el in gafete_abajo:
                reemplazar_en_nodo(el, mapa_g2)
        else:
            # Si el total seleccionado es impar, se elimina el segundo gafete en blanco
            for el in gafete_abajo:
                nueva_pagina.remove(el)

        body.append(nueva_pagina)

    out_zip.writestr("content.xml", ET.tostring(root, encoding="utf-8", xml_declaration=True))
    in_zip.close()
    out_zip.close()
    out_buffer.seek(0)
    return out_buffer

# --- 5. BOTÓN DE EXPORTACIÓN ---
if st.button("Generar y Exportar ODP", type="primary"):
    if not seleccionados:
        st.warning("⚠️ Debes seleccionar al menos a un empleado.")
    else:
        df_sel = df[df[col_nom].isin(seleccionados)]
        with st.spinner("Procesando y generando archivo ODP..."):
            try:
                archivo_odp_listo = generar_odp(df_sel, archivo_plantilla)
                st.success(f"✅ ¡Gafetes generados exitosamente para {len(df_sel)} empleado(s)!")
                st.download_button(
                    label="📥 Descargar Gafetes (.odp)",
                    data=archivo_odp_listo,
                    file_name="Gafetes_Generados.odp",
                    mime="application/vnd.oasis.opendocument.presentation",
                )
            except Exception as e:
                st.error(f"Error durante el procesamiento: {e}")
