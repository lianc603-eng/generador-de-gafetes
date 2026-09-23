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

# Configuración de la interfaz
st.set_page_config(page_title="Generador de Gafetes", page_icon="🪪", layout="wide")

st.title("🪪 Generador de Gafetes Oficiales")
st.write(
    "Selecciona a los empleados requeridos para generar sus gafetes en formato **ODP** o **PowerPoint (PPTX)** "
    "respetando la plantilla institucional (2 trabajadores por hoja)."
)

# --- 1. DETECCIÓN AUTOMÁTICA DE ARCHIVOS ---
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

archivo_plantilla = buscar_plantilla_local()

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
    st.error("❌ No se encontró la plantilla ODP en el repositorio.")
    st.stop()

# Detección de columnas de la base de datos
col_num = next((c for c in df.columns if "NUM" in c.upper() or "EMPLEADO" in c.upper()), "NUMERO DE EMPLEADO")
col_nom = next((c for c in df.columns if "NOMBRE" in c.upper()), "Nombre completo")
col_pto = next((c for c in df.columns if "PUESTO" in c.upper()), "Puesto")

# --- 2. SELECTOR Y OPCIONES ---
col1, col2 = st.columns(2)
with col1:
    seleccionar_todos = st.checkbox("Seleccionar todos los empleados de la lista")

with col2:
    formato_salida = st.radio(
        "Formato de descarga:",
        options=["ODP (LibreOffice)", "PPTX (PowerPoint)", "Ambos formatos (ZIP)"],
        horizontal=True,
    )

opciones = df[col_nom].dropna().tolist()

if seleccionar_todos:
    seleccionados = st.multiselect("Empleados a generar:", opciones, default=opciones)
else:
    seleccionados = st.multiselect("Empleados a generar:", opciones)

# --- 3. REEMPLAZO DINÁMICO (EXCLUSIVAMENTE LO MARCADO CON *) ---
def limpiar_id(val):
    if pd.isna(val):
        return ""
    try:
        return str(int(float(val))).strip()
    except Exception:
        return str(val).strip()

def aplicar_datos_a_elemento(elem, nombre, puesto, num_emp, folio):
    text_ns = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"

    for p in elem.iter(f"{{{text_ns}}}p"):
        p_text = "".join(p.itertext()).strip()
        if not p_text:
            continue

        spans = p.findall(f"{{{text_ns}}}span")

        # 1. Nombre completo (delimitado entre *...*)
        if "C. *" in p_text or "autoriza al" in p_text or "Firma del" in p_text:
            nuevo_texto = f"C. {nombre}"
            if spans:
                spans[0].text = nuevo_texto
                for s in spans[1:]:
                    s.text = ""
                p.text = None
            else:
                p.text = nuevo_texto

        # 2. Puesto / Cargo (marcado con *...*)
        elif "Como:" in p_text or ("*" in p_text and any(c.isalpha() for c in p_text) and not "/" in p_text and not "EMP" in p_text and not "EMPLEADO" in p_text):
            if spans:
                spans[0].text = puesto
                for s in spans[1:]:
                    s.text = ""
                p.text = None
            else:
                p.text = puesto

        # 3. Código institucional DDUMA-EMP-*[ID]*
        elif "DDUMA-EMP" in p_text:
            nuevo_texto = f"DDUMA-EMP-{num_emp}"
            if spans:
                spans[0].text = nuevo_texto
                for s in spans[1:]:
                    s.text = ""
                p.text = None
            else:
                p.text = nuevo_texto

        # 4. Etiqueta NO. DE EMPLEADO:*[ID]*
        elif "NO. DE EMPLEADO" in p_text or "EMPLEADO:" in p_text:
            nuevo_texto = f"NO. DE EMPLEADO:{num_emp}"
            if spans:
                spans[0].text = nuevo_texto
                for s in spans[1:]:
                    s.text = ""
                p.text = None
            else:
                p.text = nuevo_texto

        # 5. Folio Consecutivo (*XXX*/DDUMA/2026)
        elif "/DDUMA/" in p_text:
            nuevo_texto = f"{folio:03d}/DDUMA/2026"
            if spans:
                spans[0].text = nuevo_texto
                for s in spans[1:]:
                    s.text = ""
                p.text = None
            else:
                p.text = nuevo_texto

        # 6. ID individual numérico delimitado con asteriscos (*...*)
        elif re.search(r'\*\d+\*', p_text):
            nuevo_texto = re.sub(r'\*\d+\*', num_emp, p_text)
            if spans:
                spans[0].text = nuevo_texto
                for s in spans[1:]:
                    s.text = ""
                p.text = None
            else:
                p.text = nuevo_texto

        # 7. Limpieza final: eliminar cualquier asterisco residual
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

    # Agrupar de 2 en 2 empleados por cada hoja Carta
    for i in range(0, len(filas), 2):
        emp1 = filas[i][1]
        emp2 = filas[i + 1][1] if (i + 1 < len(filas)) else None

        nueva_pagina = copy.deepcopy(pagina_maestra)
        nueva_pagina.set(f"{{{draw_ns}}}name", f"Hoja_{num_hoja}")
        num_hoja += 1

        elementos = [e for e in nueva_pagina if not e.tag.endswith("notes")]
        mitad = len(elementos) // 2

        elementos_arriba = elementos[:mitad]
        elementos_abajo = elementos[mitad:]

        # 1. Asignar datos al Trabajador 1 (Arriba)
        nom1 = str(emp1.get(col_nom, "")).strip().upper()
        num1 = limpiar_id(emp1.get(col_num, ""))
        pto1 = str(emp1.get(col_pto, "")).strip().upper() if pd.notna(emp1.get(col_pto)) else ""

        for el in elementos_arriba:
            aplicar_datos_a_elemento(el, nom1, pto1, num1, i + 1)

        # 2. Asignar datos al Trabajador 2 (Abajo)
        if emp2 is not None:
            nom2 = str(emp2.get(col_nom, "")).strip().upper()
            num2 = limpiar_id(emp2.get(col_num, ""))
            pto2 = str(emp2.get(col_pto, "")).strip().upper() if pd.notna(emp2.get(col_pto)) else ""

            for el in elementos_abajo:
                aplicar_datos_a_elemento(el, nom2, pto2, num2, i + 2)
        else:
            # Si se seleccionó un número impar de empleados, se retira el gafete sobrante
            for el in elementos_abajo:
                nueva_pagina.remove(el)

        body.append(nueva_pagina)

    out_zip.writestr("content.xml", ET.tostring(root, encoding="utf-8", xml_declaration=True))
    in_zip.close()
    out_zip.close()
    out_buffer.seek(0)
    return out_buffer

# --- 5. CONVERSIÓN A POWERPOINT (PPTX) ---
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

# --- 6. BOTÓN DE EXPORTACIÓN ---
if st.button("Generar Gafetes", type="primary"):
    if not seleccionados:
        st.warning("⚠️ Debes seleccionar al menos a un empleado.")
    else:
        df_filtrado = df[df[col_nom].isin(seleccionados)].copy()
        df_filtrado["_orden"] = df_filtrado[col_nom].map({nombre: idx for idx, nombre in enumerate(seleccionados)})
        df_final = df_filtrado.sort_values("_orden").drop(columns=["_orden"])

        with st.spinner("Generando gafetes sin asteriscos..."):
            try:
                odp_buffer = generar_odp(df_final, archivo_plantilla)
                st.success(f"✅ ¡Gafetes listos para {len(df_final)} persona(s)!")

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
