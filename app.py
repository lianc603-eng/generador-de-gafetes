import copy
import io
import re
import xml.etree.ElementTree as ET
import zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Gafetes", page_icon="🪪", layout="wide")

st.title("🪪 Generador de Gafetes Oficiales")
st.write(
    "Selecciona a los empleados requeridos para generar sus gafetes en formato **ODP** respetando la plantilla institucional."
)

# --- Cargar datos ---
@st.cache_data
def cargar_datos():
    try:
        df = pd.read_excel("base.xlsx")
        # Limpieza básica de espacios en nombres de columnas
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"Error al leer base.xlsx: {e}")
        return None

df = cargar_datos()

if df is not None:
    # Identificar columna para el selector
    col_nombre = next((c for c in df.columns if "NOMBRE" in c.upper()), df.columns[0])

    col1, col2 = st.columns([3, 1])
    with col1:
        todos = st.checkbox("Seleccionar todos los empleados")
    
    opciones = df[col_nombre].dropna().unique().tolist()
    
    if todos:
        seleccionados = st.multiselect("Empleados seleccionados:", opciones, default=opciones)
    else:
        seleccionados = st.multiselect("Empleados seleccionados:", opciones)

    # --- Función para procesar ODP ---
    def generar_odp(df_filtrado, template_path="plantilla_maestra.odp"):
        with open(template_path, "rb") as f:
            template_bytes = f.read()

        in_zip = zipfile.ZipFile(io.BytesIO(template_bytes), "r")
        out_buffer = io.BytesIO()
        out_zip = zipfile.ZipFile(out_buffer, "w", zipfile.ZIP_DEFLATED)

        # Copiar todos los archivos excepto content.xml
        for item in in_zip.infolist():
            if item.filename != "content.xml":
                out_zip.writestr(item, in_zip.read(item.filename))

        # Procesar content.xml
        xml_content = in_zip.read("content.xml").decode("utf-8")
        
        # Registrar namespaces para evitar prefijos tipo ns0:
        namespaces = dict([node for _, node in ET.iterparse(io.StringIO(xml_content), events=["start-ns"])])
        for prefix, uri in namespaces.items():
            ET.register_namespace(prefix, uri)

        root = ET.fromstring(xml_content)

        # Namespaces XML estándar de ODF
        draw_ns = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
        text_ns = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"

        # Localizar el contenedor de diapositivas y la primera página
        body = root.find(".//{urn:oasis:names:tc:opendocument:xmlns:office:1.0}presentation")
        if body is None:
            body = root.find(".//{urn:oasis:names:tc:opendocument:xmlns:office:1.0}body/*")

        paginas = body.findall(f"{{{draw_ns}}}page")
        if not paginas:
            raise ValueError("No se encontraron páginas en la plantilla maestra.")

        pagina_maestra = paginas[0]

        # Eliminar las páginas actuales del body para reinsertar las generadas
        for p in paginas:
            body.remove(p)

        # Función auxiliar para reemplazar placeholders en un elemento XML
        def reemplazar_texto(elem, mapa_reemplazos):
            if elem.text:
                for k, v in mapa_reemplazos.items():
                    if k in elem.text:
                        elem.text = elem.text.replace(k, str(v))
            if elem.tail:
                for k, v in mapa_reemplazos.items():
                    if k in elem.tail:
                        elem.tail = elem.tail.replace(k, str(v))
            for child in elem:
                reemplazar_texto(child, mapa_reemplazos)

        # Duplicar y personalizar por cada empleado
        for idx, (_, row) in enumerate(df_filtrado.iterrows(), start=1):
            nueva_pag = copy.deepcopy(pagina_maestra)
            nueva_pag.set(f"{{{draw_ns}}}name", f"Gafete_{idx}")

            # Mapeo flexible de columnas a los textos con *
            # Por ejemplo, si en la plantilla está "*NOMBRE*" o "*CARGO*"
            mapa = {}
            for col in df_filtrado.columns:
                val = "" if pd.isna(row[col]) else str(row[col]).strip()
                # Reemplaza variantes: *COLUMNA*, *columna*, etc.
                mapa[f"*{col}*"] = val
                mapa[f"*{col.upper()}*"] = val
                mapa[f"*{col.lower()}*"] = val

            reemplazar_texto(nueva_pag, mapa)
            body.append(nueva_pag)

        # Guardar content.xml modificado
        out_zip.writestr("content.xml", ET.tostring(root, encoding="utf-8", xml_declaration=True))
        in_zip.close()
        out_zip.close()
        out_buffer.seek(0)
        return out_buffer

    # --- Botón de Generación ---
    if st.button("Generar y Exportar ODP", type="primary"):
        if not seleccionados:
            st.warning("Por favor selecciona al menos a una persona.")
        else:
            df_sel = df[df[col_nombre].isin(seleccionados)]
            with st.spinner("Generando archivo ODP..."):
                try:
                    odp_resultado = generar_odp(df_sel)
                    st.success(f"¡Listo! Se han generado {len(df_sel)} gafete(s).")
                    st.download_button(
                        label="📥 Descargar Gafetes (.odp)",
                        data=odp_resultado,
                        file_name="Gafetes_Generados.odp",
                        mime="application/vnd.oasis.opendocument.presentation",
                    )
                except Exception as err:
                    st.error(f"Error durante el procesamiento: {err}")
