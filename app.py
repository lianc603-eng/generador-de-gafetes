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
    "Selecciona a los empleados y elige el formato deseado (**ODP** de LibreOffice o **PPTX** de PowerPoint)."
)

# --- 1. DETECCIÓN / CARGA DE ARCHIVOS ---
st.sidebar.header("📁 Configuración de Plantilla")

def buscar_plantilla_local():
    posibles = [
        "PLANTILLA MAESTRA.odp",
        "plantilla_maestra.odp",
        "Plantilla_Maestra.odp",
        "Plantilla Maestra.odp",
    ]
    for p in posibles:
        if os.path.exists(p):
            return p
    for f in os.listdir("."):
        if f.lower().endswith(".odp"):
            return f
    return None

plantilla_subida = st.sidebar.file_uploader("Subir plantilla diferente (.odp):", type=["odp"])
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
    st.error("❌ No se encontró la plantilla ODP. Súbela en la barra lateral o a GitHub.")
    st.stop()

# --- 2. COLUMNAS ---
col_num = next((c for c in df.columns if "NUM" in c.upper() or "EMPLEADO" in c.upper()), "NUMERO DE EMPLEADO")
col_nom = next((c for c in df.columns if "NOMBRE" in c.upper()), "Nombre completo")
col_pto = next((c for c in df.columns if "PUESTO" in c.upper()), "Puesto")

# --- 3. SELECTOR DE EMPLEADOS Y FORMATO ---
col_chk, col_fmt = st.columns([2, 2])
with col_chk:
    seleccionar_todos = st.checkbox("Seleccionar todos los empleados de la lista")

with col_fmt:
    formato_salida = st.radio(
        "Formato de descarga:",
        options=["ODP (LibreOffice / OpenDocument)", "PPTX (Microsoft PowerPoint)", "Ambos formatos (ZIP)"],
        horizontal=True,
    )

opciones = df[col_nom].dropna().tolist()

if seleccionar_todos:
    seleccionados = st.multiselect("Empleados a generar:", opciones, default=opciones)
else:
    seleccionados = st.multiselect("Empleados a generar:", opciones)

# --- 4. FUNCIONES DE APOYO XML / COORDENADAS ---
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

def actualizar_textos_gafete(elemento, nombre, puesto, num_emp, folio_num):
    text_ns = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    
    for p in elemento.iter(f"{{{text_ns}}}p"):
        p_text = "".join(p.itertext())
        if not p_text.strip():
            continue
            
        # 1. Nombre del empleado (Frente y Firma) - Limpio sin asteriscos
        if "Citlalli" in p_text or "*Citlalli" in p_text:
            spans = p.findall(f"{{{text_ns}}}span")
            if spans:
                spans[0].text = f"C. {nombre}"
                for s in spans[1:]:
                    s.text = ""
            else:
                p.text = f"C. {nombre}"

        # 2. Puesto - Limpio sin asteriscos
        elif "ANALISTA" in p_text or "*ANALISTA*" in p_text:
            puesto_val = puesto if (pd.notna(puesto) and str(puesto).strip()) else ""
            spans = p.findall(f"{{{text_ns}}}span")
            if spans:
                reemplazado = False
                for s in spans:
                    if s.text and "ANALISTA" in s.text:
                        s.text = puesto_val
                        reemplazado = True
                    elif s.text:
                        s.text = s.text.replace("*", "")
                if not reemplazado:
                    spans[0].text = puesto_val
            else:
                p.text = puesto_val

        # 3. Código DDUMA-EMP- - Limpio sin asteriscos
        elif "DDUMA-EMP-" in p_text:
            spans = p.findall(f"{{{text_ns}}}span")
            if spans:
                spans[0].text = f"DDUMA-EMP-{num_emp}"
                for s in spans[1:]:
                    s.text = ""
            else:
                p.text = f"DDUMA-EMP-{num_emp}"

        # 4. Número de empleado - Limpio sin asteriscos
        elif "NO. DE EMPLEADO" in p_text:
            spans = p.findall(f"{{{text_ns}}}span")
            if spans:
                spans[0].text = f"NO. DE EMPLEADO:{num_emp}"
                for s in spans[1:]:
                    s.text = ""
            else:
                p.text = f"NO. DE EMPLEADO:{num_emp}"

        # 5. Folio consecutivo - Limpio sin asteriscos
        elif "/DDUMA/" in p_text:
            nuevo_folio = f"{folio_num:03d}/DDUMA/2026 "
            spans = p.findall(f"{{{text_ns}}}span")
            if spans:
                spans[0].text = nuevo_folio
                for s in spans[1:]:
                    s.text = ""
            else:
                p.text = nuevo_folio

        # Limpieza residual
        if p.text and "*" in p.text:
            p.text = p.text.replace("*", "")
        for s in p.findall(f"{{{text_ns}}}span"):
            if s.text and "*" in s.text:
                s.text = s.text.replace("*", "")

# --- 5. GENERACIÓN ODP ---
def generar_odp(df_seleccionados, fuente_plantilla):
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

        nom1 = str(emp1.get(col_nom, "")).strip()
        num1 = str(emp1.get(col_num, "")).strip()
        pto1 = str(emp1.get(col_pto, "")).strip() if pd.notna(emp1.get(col_pto)) else ""
        for el in elementos_arriba:
            actualizar_textos_gafete(el, nom1, pto1, num1, i + 1)

        if emp2 is not None:
            nom2 = str(emp2.get(col_nom, "")).strip()
            num2 = str(emp2.get(col_num, "")).strip()
            pto2 = str(emp2.get(col_pto, "")).strip() if pd.notna(emp2.get(col_pto)) else ""
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
                "El comando LibreOffice no está disponible en el servidor para convertir a PowerPoint. "
                "Asegúrate de agregar 'packages.txt' con la palabra 'libreoffice' en tu repositorio de GitHub."
            )

        proceso = subprocess.run(
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
            raise RuntimeError(f"Error al convertir a PPTX: {proceso.stderr}")

# --- 7. BOTÓN DE EXPORTACIÓN ---
if st.button("Generar Gafetes", type="primary"):
    if not seleccionados:
        st.warning("⚠️ Selecciona al menos a un empleado.")
    else:
        df_filtrado = df[df[col_nom].isin(seleccionados)].copy()
        df_filtrado["_orden_sel"] = df_filtrado[col_nom].map({nombre: idx for idx, nombre in enumerate(seleccionados)})
        df_sel = df_filtrado.sort_values("_orden_sel").drop(columns=["_orden_sel"])

        with st.spinner("Generando archivos..."):
            try:
                odp_buffer = generar_odp(df_sel, archivo_plantilla)
                st.success(f"✅ ¡Gafetes procesados para {len(df_sel)} empleado(s)!")

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
