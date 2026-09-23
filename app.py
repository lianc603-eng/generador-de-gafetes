import copy
import io
import os
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
    "Selecciona a los empleados y personaliza el formato de visualización y exportación."
)

# --- 1. DETECCIÓN Y CARGA DE ARCHIVOS ---
st.sidebar.header("📁 Configuración de Plantilla")

def buscar_plantilla_local():
    posibles = [
        "PLANTILLA MAESTRA.odp",
        "plantilla_maestra.odp",
        "Plantilla_Maestra.odp",
        "Plantilla Maestra.odp",
        "PLANTILLA_MAESTRA.odp",
    ]
    for p in posibles:
        if os.path.exists(p):
            return p
    for f in os.listdir("."):
        if f.lower().endswith(".odp"):
            return f
    return None

plantilla_subida = st.sidebar.file_uploader("Subir nueva plantilla (.odp):", type=["odp"])
archivo_plantilla = plantilla_subida if plantilla_subida is not None else buscar_plantilla_local()

@st.cache_data
def cargar_base():
    posibles_excel = ["base.xlsx", "BASE.xlsx", "Base.xlsx"]
    for f in posibles_excel:
        if os.path.exists(f):
            df = pd.read_excel(f)
            df.columns = [str(c).strip() for c in df.columns]
            return df
    for f in os.listdir("."):
        if f.lower().endswith(".xlsx") and not f.startswith("~$"):
            df = pd.read_excel(f)
            df.columns = [str(c).strip() for c in df.columns]
            return df
    return None

df = cargar_base()

if df is None:
    st.error("❌ No se encontró el archivo 'base.xlsx' en el repositorio.")
    st.stop()

if archivo_plantilla is None:
    st.error("❌ No se encontró la plantilla ODP. Súbela en la barra lateral o al repositorio de GitHub.")
    st.stop()

# Identificar columnas
col_num = next((c for c in df.columns if "NUM" in c.upper() or "EMPLEADO" in c.upper()), "NUMERO DE EMPLEADO")
col_nom = next((c for c in df.columns if "NOMBRE" in c.upper()), "Nombre completo")
col_pto = next((c for c in df.columns if "PUESTO" in c.upper()), "Puesto")

# --- 2. OPCIONES DE FORMATO ---
col_opt1, col_opt2, col_opt3 = st.columns(3)

with col_opt1:
    orden_nombre = st.radio(
        "Orden del Nombre:",
        options=["Apellidos primero (como en Excel)", "Nombre(s) primero"],
        index=0,
        help="El archivo Excel viene por defecto con los apellidos primero."
    )

with col_opt2:
    estilo_capitalizacion = st.radio(
        "Formato de texto:",
        options=["MAYÚSCULAS COMPLETAS", "Tipo Nombre Propio (Título)"],
        index=0
    )

with col_opt3:
    formato_salida = st.radio(
        "Formato de descarga:",
        options=["ODP (LibreOffice)", "PPTX (PowerPoint)", "Ambos formatos (ZIP)"],
        index=0
    )

# --- 3. SELECTOR DE EMPLEADOS ---
col_chk, _ = st.columns([2, 2])
with col_chk:
    seleccionar_todos = st.checkbox("Seleccionar todos los empleados de la lista")

opciones = df[col_nom].dropna().tolist()

if seleccionar_todos:
    seleccionados = st.multiselect("Empleados a generar:", opciones, default=opciones)
else:
    seleccionados = st.multiselect("Empleados a generar:", opciones)

# --- 4. FUNCIONES DE FORMATEO Y XML ---
def transformar_nombre(texto_nombre, orden, estilo):
    if not texto_nombre or pd.isna(texto_nombre):
        return ""
    texto = str(texto_nombre).strip()
    partes = texto.split()
    
    if orden == "Nombre(s) primero" and len(partes) >= 3:
        # En la base: [Apellido Paterno] [Apellido Materno] [Nombres...]
        apellidos = " ".join(partes[:2])
        nombres = " ".join(partes[2:])
        resultado = f"{nombres} {apellidos}"
    else:
        # Apellidos primero (tal como viene en la base de datos)
        resultado = texto

    if estilo == "Tipo Nombre Propio (Título)":
        return resultado.title()
    return resultado.upper()

def limpiar_id_empleado(val):
    if pd.isna(val):
        return ""
    try:
        # Evitar que aparezca con punto decimal (ej. 7394.0 -> 7394)
        return str(int(float(val))).strip()
    except Exception:
        return str(val).strip()

def parse_coord_cm(val_str):
    if not val_str:
        return None
    s = str(val_str).strip().lower()
    try:
        if s.endswith("cm"):
            return float(s[:-2])
        elif s.endswith("mm"):
            return float(s[:-2]) / 10.0
        elif s.endswith("in"):
            return float(s[:-2]) * 2.54
        elif s.endswith("pt"):
            return float(s[:-2]) * (2.54 / 72.0)
        else:
            return float(s)
    except Exception:
        return None

def obtener_y_cm(elem):
    y_attr = elem.attrib.get("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}y")
    if y_attr:
        v = parse_coord_cm(y_attr)
        if v is not None:
            return v
    for child in elem.iter():
        y_c = child.attrib.get("{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}y")
        if y_c:
            v = parse_coord_cm(y_c)
            if v is not None:
                return v
    return None

def asignar_texto_nodo(parrafo, nuevo_texto):
    text_ns = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    spans = parrafo.findall(f"{{{text_ns}}}span")
    if spans:
        spans[0].text = nuevo_texto
        for s in spans[1:]:
            s.text = ""
        parrafo.text = None
    else:
        parrafo.text = nuevo_texto

def actualizar_textos_gafete(elemento, nombre, puesto, num_emp, folio_num):
    text_ns = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    
    for p in elemento.iter(f"{{{text_ns}}}p"):
        p_text = "".join(p.itertext()).strip()
        if not p_text:
            continue

        # 1. Nombre completo (Frente y Firma)
        if "Citlalli" in p_text or "Brown" in p_text:
            asignar_texto_nodo(p, f"C. {nombre}")

        # 2. Puesto
        elif "ANALISTA" in p_text:
            puesto_val = puesto if (pd.notna(puesto) and str(puesto).strip()) else ""
            asignar_texto_nodo(p, puesto_val)

        # 3. Código institucional DDUMA-EMP-[ID]
        elif "DDUMA-EMP" in p_text:
            asignar_texto_nodo(p, f"DDUMA-EMP-{num_emp}")

        # 4. Número de empleado / ID
        elif "NO. DE EMPLEADO" in p_text or "EMPLEADO:" in p_text:
            asignar_texto_nodo(p, f"NO. DE EMPLEADO:{num_emp}")

        # 5. Folio consecutivo
        elif "/DDUMA/" in p_text:
            asignar_texto_nodo(p, f"{folio_num:03d}/DDUMA/2026")

        # 6. Reemplazo de respaldo por si el número de empleado '9820' aparece suelto
        elif "9820" in p_text:
            texto_actualizado = p_text.replace("*9820*", num_emp).replace("9820", num_emp)
            asignar_texto_nodo(p, texto_actualizado)

        # Limpieza residual de cualquier asterisco que hubiera quedado
        if p.text and "*" in p.text:
            p.text = p.text.replace("*", "")
        for s in p.findall(f"{{{text_ns}}}span"):
            if s.text and "*" in s.text:
                s.text = s.text.replace("*", "")

# --- 5. MOTOR DE GENERACIÓN ODP ---
def generar_odp(df_seleccionados, fuente_plantilla, orden_nom, estilo_nom):
    if isinstance(fuente_plantilla, str):
        with open(fuente_plantilla, "rb") as f:
            template_bytes = f.read()
    else:
        fuente_plantilla.seek(0)
        template_bytes = fuente_plantilla.read()

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
    if not paginas:
        raise ValueError("No se encontraron páginas en la plantilla.")

    pagina_maestra = paginas[0]

    for p in paginas:
        body.remove(p)

    filas = list(df_seleccionados.iterrows())
    pagina_idx = 1

    for i in range(0, len(filas), 2):
        emp1 = filas[i][1]
        emp2 = filas[i + 1][1] if (i + 1 < len(filas)) else None

        nueva_pagina = copy.deepcopy(pagina_maestra)
        nueva_pagina.set(f"{{{draw_ns}}}name", f"Hoja_{pagina_idx}")
        pagina_idx += 1

        elementos = [e for e in nueva_pagina if not e.tag.endswith("notes")]

        elementos_arriba = []
        elementos_abajo = []

        for el in elementos:
            y_pos = obtener_y_cm(el)
            if y_pos is None:
                continue
            if y_pos < 14.0:
                elementos_arriba.append(el)
            else:
                elementos_abajo.append(el)

        # 1. Asignar datos al gafete de arriba (Empleado 1)
        nom1 = transformar_nombre(emp1.get(col_nom, ""), orden_nom, estilo_nom)
        num1 = limpiar_id_empleado(emp1.get(col_num, ""))
        pto1 = str(emp1.get(col_pto, "")).strip() if pd.notna(emp1.get(col_pto)) else ""
        if estilo_nom == "Tipo Nombre Propio (Título)":
            pto1 = pto1.title()
        else:
            pto1 = pto1.upper()

        for el in elementos_arriba:
            actualizar_textos_gafete(el, nom1, pto1, num1, i + 1)

        # 2. Asignar datos al gafete de abajo (Empleado 2)
        if emp2 is not None:
            nom2 = transformar_nombre(emp2.get(col_nom, ""), orden_nom, estilo_nom)
            num2 = limpiar_id_empleado(emp2.get(col_num, ""))
            pto2 = str(emp2.get(col_pto, "")).strip() if pd.notna(emp2.get(col_pto)) else ""
            if estilo_nom == "Tipo Nombre Propio (Título)":
                pto2 = pto2.title()
            else:
                pto2 = pto2.upper()

            for el in elementos_abajo:
                actualizar_textos_gafete(el, nom2, pto2, num2, i + 2)
        else:
            for el in elementos_abajo:
                nueva_pagina.remove(el)

        body.append(nueva_pagina)

    out_zip.writestr("content.xml", ET.tostring(root, encoding="utf-8", xml_declaration=True))
    in_zip.close()
    out_zip.close()
    out_buffer.seek(0)
    return out_buffer

# --- 6. CONVERSIÓN A POWERPOINT (PPTX) ---
def convertir_odp_a_pptx(odp_bytes):
    with tempfile.TemporaryDirectory() as tmpdir:
        input_odp = os.path.join(tmpdir, "gafetes.odp")
        with open(input_odp, "wb") as f:
            f.write(odp_bytes.getvalue())

        cmd = None
        for executable in ["libreoffice", "soffice"]:
            if shutil.which(executable):
                cmd = executable
                break

        if not cmd:
            raise RuntimeError(
                "LibreOffice no está instalado en el servidor. Asegúrate de tener 'packages.txt' con 'libreoffice' en tu repositorio de GitHub."
            )

        subprocess.run(
            [cmd, "--headless", "--convert-to", "pptx", input_odp, "--outdir", tmpdir],
            capture_output=True,
            text=True,
            timeout=120,
        )

        expected_pptx = os.path.join(tmpdir, "gafetes.pptx")
        if os.path.exists(expected_pptx):
            with open(expected_pptx, "rb") as f:
                return io.BytesIO(f.read())
        else:
            raise RuntimeError("No se pudo completar la conversión a PPTX.")

# --- 7. BOTÓN DE EXPORTACIÓN ---
if st.button("Generar Gafetes", type="primary"):
    if not seleccionados:
        st.warning("⚠️ Debes seleccionar al menos a un empleado.")
    else:
        df_filtrado = df[df[col_nom].isin(seleccionados)].copy()
        df_filtrado["_orden_sel"] = df_filtrado[col_nom].map({nombre: idx for idx, nombre in enumerate(seleccionados)})
        df_sel = df_filtrado.sort_values("_orden_sel").drop(columns=["_orden_sel"])

        with st.spinner("Generando gafetes con el formato seleccionado..."):
            try:
                odp_buffer = generar_odp(df_sel, archivo_plantilla, orden_nombre, estilo_capitalizacion)
                st.success(f"✅ ¡Gafetes listos para {len(df_sel)} empleado(s)!")

                if "ODP" in formato_salida:
                    st.download_button(
                        label="📥 Descargar en ODP (LibreOffice)",
                        data=odp_buffer.getvalue(),
                        file_name="Gafetes_Generados.odp",
                        mime="application/vnd.oasis.opendocument.presentation",
                    )

                if "PPTX" in formato_salida:
                    pptx_buffer = convertir_odp_a_pptx(odp_buffer)
                    st.download_button(
                        label="📥 Descargar en PowerPoint (.pptx)",
                        data=pptx_buffer.getvalue(),
                        file_name="Gafetes_Generados.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )

                if "Ambos" in formato_salida:
                    pptx_buffer = convertir_odp_a_pptx(odp_buffer)
                    zip_salida = io.BytesIO()
                    with zipfile.ZipFile(zip_salida, "w", zipfile.ZIP_DEFLATED) as z:
                        z.writestr("Gafetes_Generados.odp", odp_buffer.getvalue())
                        z.writestr("Gafetes_Generados.pptx", pptx_buffer.getvalue())
                    zip_salida.seek(0)

                    st.download_button(
                        label="📥 Descargar Ambos Formatos (.zip)",
                        data=zip_salida.getvalue(),
                        file_name="Gafetes_ODP_y_PPTX.zip",
                        mime="application/zip",
                    )

            except Exception as e:
                st.error(f"Error durante el procesamiento: {e}")
