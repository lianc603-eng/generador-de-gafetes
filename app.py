import copy
import io
import os
import re
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Gafetes", page_icon="🪪", layout="wide")

st.title("🪪 Generador de Gafetes Oficiales")
st.write(
    "Genera tus gafetes en formato **ODP** o **PowerPoint (PPTX)** con 2 trabajadores distintos por hoja tamaño Carta."
)

# --- 1. DETECCIÓN DE ARCHIVOS ---
def buscar_plantilla():
    for f in os.listdir("."):
        if f.lower().endswith(".odp"):
            return f
    return None

archivo_plantilla = buscar_plantilla()

@st.cache_data
def cargar_base():
    for f in os.listdir("."):
        if f.lower().endswith(".xlsx") and not f.startswith("~$"):
            df = pd.read_excel(f)
            df.columns = [str(c).strip() for c in df.columns]
            return df
    return None

df = cargar_base()

if df is None:
    st.error("❌ No se encontró el archivo Excel de empleados en el repositorio.")
    st.stop()

if archivo_plantilla is None:
    st.error("❌ No se encontró la plantilla ODP en el repositorio.")
    st.stop()

col_num = next((c for c in df.columns if "NUM" in c.upper() or "EMPLEADO" in c.upper()), "NUMERO DE EMPLEADO")
col_nom = next((c for c in df.columns if "NOMBRE" in c.upper()), "Nombre completo")
col_pto = next((c for c in df.columns if "PUESTO" in c.upper()), "Puesto")

# --- 2. SELECTOR Y FORMATO ---
col1, col2 = st.columns(2)
with col1:
    seleccionar_todos = st.checkbox("Seleccionar todos los empleados de la lista")
with col2:
    formato_salida = st.radio(
        "Formato de descarga:",
        options=["ODP (LibreOffice)", "PPTX (PowerPoint)"],
        horizontal=True,
    )

opciones = df[col_nom].dropna().tolist()

if seleccionar_todos:
    seleccionados = st.multiselect("Empleados a generar:", opciones, default=opciones)
else:
    seleccionados = st.multiselect("Empleados a generar:", opciones)

# --- 3. FUNCIONES DE PROCESAMIENTO ---
def limpiar_id(val):
    if pd.isna(val):
        return ""
    try:
        return str(int(float(val))).strip()
    except Exception:
        return str(val).strip()

def parse_y_cm(elem):
    """Calcula la posición vertical del elemento en cm."""
    y_str = elem.attrib.get("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}y")
    if not y_str:
        for ch in elem.iter():
            y_str = ch.attrib.get("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}y")
            if y_str:
                break
    if not y_str:
        return None
    y_clean = str(y_str).strip().lower()
    try:
        if y_clean.endswith("cm"):
            return float(y_clean[:-2])
        elif y_clean.endswith("mm"):
            return float(y_clean[:-2]) / 10.0
        elif y_clean.endswith("in"):
            return float(y_clean[:-2]) * 2.54
        elif y_clean.endswith("pt"):
            return float(y_clean[:-2]) * (2.54 / 72.0)
        return float(y_clean)
    except Exception:
        return None

def asignar_texto_limpio(parrafo, texto_nuevo):
    """Asigna el texto nuevo de forma limpia sin asteriscos ni duplicaciones."""
    text_ns = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    spans = parrafo.findall(f"{{{text_ns}}}span")
    if spans:
        spans[0].text = texto_nuevo
        for s in spans[1:]:
            s.text = ""
        parrafo.text = None
    else:
        parrafo.text = texto_nuevo

def procesar_elemento(elem, nombre, puesto, num_emp, folio):
    text_ns = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"

    for p in elem.iter(f"{{{text_ns}}}p"):
        p_text = "".join(p.itertext()).strip()
        if not p_text:
            continue

        # Evitar sobreescribir las etiquetas fijas
        if p_text in ["Se autoriza al", "Como:", "Firma del Trabajador"]:
            continue

        # 1. Caja de ID / Número de empleado (detecta 'ID:', 'ID :', 'NO. DE EMPLEADO', etc.)
        if p_text.upper().startswith("ID:") or p_text.upper().startswith("ID :") or "ID:" in p_text.upper() or "ID :" in p_text.upper():
            asignar_texto_limpio(p, f"ID: {num_emp}")

        elif "NO. DE EMPLEADO" in p_text.upper() or "EMPLEADO:" in p_text.upper():
            asignar_texto_limpio(p, f"NO. DE EMPLEADO:{num_emp}")

        # 2. Código institucional DDUMA-EMP-
        elif "DDUMA-EMP" in p_text:
            asignar_texto_limpio(p, f"DDUMA-EMP-{num_emp}")

        # 3. Nombre del trabajador (Frente y Firma)
        elif "Citlalli" in p_text or "Brown" in p_text or p_text.startswith("C. *") or p_text.startswith("C.*") or p_text.startswith("C. "):
            asignar_texto_limpio(p, f"C. {nombre}")

        # 4. Puesto del trabajador
        elif "ANALISTA" in p_text or ("*" in p_text and any(c.isalpha() for c in p_text) and not "DDUMA" in p_text):
            asignar_texto_limpio(p, puesto)

        # 5. Folio Consecutivo
        elif "/DDUMA/" in p_text:
            asignar_texto_limpio(p, f"{folio:03d}/DDUMA/2026")

        # 6. Reemplazo de respaldo si hay un número suelto entre asteriscos
        elif re.search(r'\*\d+\*', p_text):
            nuevo = re.sub(r'\*\d+\*', num_emp, p_text)
            asignar_texto_limpio(p, nuevo)

        # Limpieza residual de cualquier asterisco que hubiera quedado
        if p.text and "*" in p.text:
            p.text = p.text.replace("*", "")
        for s in p.findall(f"{{{text_ns}}}span"):
            if s.text and "*" in s.text:
                s.text = s.text.replace("*", "")

# --- 4. MOTOR DE GENERACIÓN ODP ---
def generar_odp(df_sel, ruta_plantilla):
    with open(ruta_plantilla, "rb") as f:
        template_bytes = f.read()

    in_zip = zipfile.ZipFile(io.BytesIO(template_bytes), "r")
    out_buffer = io.BytesIO()
    out_zip = zipfile.ZipFile(out_buffer, "w", zipfile.ZIP_DEFLATED)

    for item in in_zip.infolist():
        if item.filename != "content.xml":
            out_zip.writestr(item, in_zip.read(item.filename))

    xml_text = in_zip.read("content.xml").decode("utf-8")

    namespaces = dict([node for _, node in ET.iterparse(io.StringIO(xml_text), events=["start-ns"])])
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)

    root = ET.fromstring(xml_text)
    draw_ns = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"

    body = root.find(".//{urn:oasis:names:tc:opendocument:xmlns:office:1.0}presentation")
    if body is None:
        body = root.find(".//{urn:oasis:names:tc:opendocument:xmlns:office:1.0}body/*")

    paginas = body.findall(f"{{{draw_ns}}}page")
    pagina_maestra = paginas[0]

    for p in paginas:
        body.remove(p)

    filas = list(df_sel.iterrows())
    num_hoja = 1

    # Agrupar de 2 en 2 trabajadores por página Carta
    for i in range(0, len(filas), 2):
        emp1 = filas[i][1]
        emp2 = filas[i + 1][1] if (i + 1 < len(filas)) else None

        nueva_pagina = copy.deepcopy(pagina_maestra)
        nueva_pagina.set(f"{{{draw_ns}}}name", f"Hoja_{num_hoja}")
        num_hoja += 1

        elementos = [e for e in nueva_pagina if not e.tag.endswith("notes")]

        # Separar por coordenada vertical: arriba (< 14 cm) y abajo (>= 14 cm)
        elementos_arriba = []
        elementos_abajo = []

        for el in elementos:
            y = parse_y_cm(el)
            if y is not None:
                if y < 14.0:
                    elementos_arriba.append(el)
                else:
                    elementos_abajo.append(el)
            else:
                elementos_arriba.append(el)

        # 1. Aplicar datos al Trabajador 1 (Arriba)
        nom1 = str(emp1.get(col_nom, "")).strip().upper()
        num1 = limpiar_id(emp1.get(col_num, ""))
        pto1 = str(emp1.get(col_pto, "")).strip().upper() if pd.notna(emp1.get(col_pto)) else ""

        for el in elementos_arriba:
            procesar_elemento(el, nom1, pto1, num1, i + 1)

        # 2. Aplicar datos al Trabajador 2 (Abajo)
        if emp2 is not None:
            nom2 = str(emp2.get(col_nom, "")).strip().upper()
            num2 = limpiar_id(emp2.get(col_num, ""))
            pto2 = str(emp2.get(col_pto, "")).strip().upper() if pd.notna(emp2.get(col_pto)) else ""

            for el in elementos_abajo:
                procesar_elemento(el, nom2, pto2, num2, i + 2)
        else:
            # Si el total seleccionado es impar, se eliminan los elementos del gafete de abajo
            for el in elementos_abajo:
                nueva_pagina.remove(el)

        body.append(nueva_pagina)

    out_zip.writestr("content.xml", ET.tostring(root, encoding="utf-8", xml_declaration=True))
    in_zip.close()
    out_zip.close()
    out_buffer.seek(0)
    return out_buffer

# --- 5. EXPORTACIÓN ---
if st.button("Generar Gafetes", type="primary"):
    if not seleccionados:
        st.warning("⚠️ Debes seleccionar al menos a un empleado.")
    else:
        df_filtrado = df[df[col_nom].isin(seleccionados)].copy()
        df_filtrado["_orden"] = df_filtrado[col_nom].map({nombre: idx for idx, nombre in enumerate(seleccionados)})
        df_final = df_filtrado.sort_values("_orden").drop(columns=["_orden"])

        with st.spinner("Creando archivo con los gafetes..."):
            try:
                odp_bytes = generar_odp(df_final, archivo_plantilla)
                st.success(f"✅ ¡Gafetes listos para {len(df_final)} persona(s)!")

                if "ODP" in formato_salida:
                    st.download_button(
                        label="📥 Descargar en ODP (LibreOffice)",
                        data=odp_bytes.getvalue(),
                        file_name="Gafetes_Generados.odp",
                        mime="application/vnd.oasis.opendocument.presentation",
                    )
                else:
                    with tempfile.TemporaryDirectory() as tmpdir:
                        odp_temp = os.path.join(tmpdir, "temp.odp")
                        with open(odp_temp, "wb") as f:
                            f.write(odp_bytes.getvalue())
                        subprocess.run(
                            ["libreoffice", "--headless", "--convert-to", "pptx", odp_temp, "--outdir", tmpdir],
                            check=True,
                            timeout=60,
                        )
                        with open(os.path.join(tmpdir, "temp.pptx"), "rb") as f:
                            st.download_button(
                                label="📥 Descargar en PowerPoint (.pptx)",
                                data=f.read(),
                                file_name="Gafetes_Generados.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            )
            except Exception as e:
                st.error(f"Error: {e}")
