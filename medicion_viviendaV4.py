from datetime import date
from io import BytesIO
from fpdf import FPDF
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Medición de Vivienda", page_icon="📐", layout="centered"
)


def texto_pdf_seguro(texto: str) -> str:
    """Evita errores de codificación en FPDF con caracteres fuera de latin-1."""
    return str(texto).encode("latin-1", "replace").decode("latin-1")


class PDFInforme(FPDF):
    """Clase personalizada para gestionar cabeceras y saltos de página limpios."""

    def header(self):
        if self.page_no() > 1:
            self.set_font("Helvetica", "I", 8)
            self.cell(
                0,
                6,
                texto_pdf_seguro("Informe de Mediciones (Continuación)"),
                ln=True,
                align="R",
            )
            self.ln(2)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.cell(0, 10, f"Página {self.page_no()}", align="C")


def generar_pdf_bytes(
    estancias: list,
    tot_paredes: float,
    tot_suelos: float,
    tot_techos: float,
    cliente: str,
    direccion: str,
    fecha_str: str,
) -> bytes:
    pdf = PDFInforme(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()

    # --- Cabecera personalizada ---
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(
        0,
        10,
        texto_pdf_seguro("Informe de Mediciones y Superficies de Vivienda"),
        ln=True,
    )

    pdf.set_font("Helvetica", "", 10)
    if cliente.strip():
        pdf.cell(0, 6, texto_pdf_seguro(f"Cliente: {cliente.strip()}"), ln=True)
    if direccion.strip():
        pdf.cell(
            0, 6, texto_pdf_seguro(f"Obra / Dirección: {direccion.strip()}"), ln=True
        )
    pdf.cell(0, 6, texto_pdf_seguro(f"Fecha: {fecha_str}"), ln=True)
    pdf.ln(5)

    # --- Función interna para dibujar cabecera de tabla ---
    anchos = (42, 22, 22, 22, 28, 27, 27)
    headers = ("Estancia", "Largo", "Ancho", "Altura", "Paredes", "Suelo", "Techo")

    def dibujar_cabecera_tabla():
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_fill_color(217, 225, 242)
        for h, w in zip(headers, anchos):
            pdf.cell(w, 8, texto_pdf_seguro(h), border=1, align="C", fill=True)
        pdf.ln()

    dibujar_cabecera_tabla()

    pdf.set_font("Helvetica", "", 9)
    for e in estancias:
        # Si queda poco espacio en la página, salta y repite cabecera de tabla
        if pdf.get_y() > 265:
            pdf.add_page()
            dibujar_cabecera_tabla()
            pdf.set_font("Helvetica", "", 9)

        fila = (
            texto_pdf_seguro(e["Estancia"]),
            f"{e['Largo (m)']:.2f} m".replace(".", ","),
            f"{e['Ancho (m)']:.2f} m".replace(".", ","),
            f"{e['Altura (m)']:.2f} m".replace(".", ","),
            f"{e['Paredes (m²)']:.2f} m2".replace(".", ","),
            f"{e['Suelo (m²)']:.2f} m2".replace(".", ","),
            f"{e['Techo (m²)']:.2f} m2".replace(".", ","),
        )
        for idx, (val, w) in enumerate(zip(fila, anchos)):
            pdf.cell(w, 7, val, border=1, align="L" if idx == 0 else "R")
        pdf.ln()

    # Fila de totales
    if pdf.get_y() > 265:
        pdf.add_page()

    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(sum(anchos[:4]), 8, "TOTALES GENERALES", border=1, align="R")
    pdf.cell(
        anchos, 8, f"{tot_paredes:.2f} m2".replace(".", ","), border=1, align="R"
    )
    pdf.cell(
        anchos, 8, f"{tot_suelos:.2f} m2".replace(".", ","), border=1, align="R"
    )
    pdf.cell(
        anchos, 8, f"{tot_techos:.2f} m2".replace(".", ","), border=1, align="R"
    )

    return bytes(pdf.output())


def generar_excel_bytes(
    estancias: list, cliente: str, direccion: str, fecha_str: str
) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mediciones"

    # --- Cabecera personalizada en Excel ---
    ws["A1"] = "INFORME DE MEDICIONES Y SUPERFICIES"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Cliente: {cliente.strip() or 'No especificado'}"
    ws["A3"] = f"Dirección: {direccion.strip() or 'No especificada'}"
    ws["A4"] = f"Fecha: {fecha_str}"

    fila_inicio_tabla = 6
    headers = [
        "Estancia",
        "Largo (m)",
        "Ancho (m)",
        "Altura (m)",
        "Paredes (m²)",
        "Suelo (m²)",
        "Techo (m²)",
    ]

    fill_header = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )
    for col_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=fila_inicio_tabla, column=col_idx, value=text)
        cell.font = Font(bold=True)
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center")

    # --- Filas de datos con fórmulas dinámicas ---
    for i, e in enumerate(estancias, start=fila_inicio_tabla + 1):
        ws.cell(row=i, column=1, value=e["Estancia"])
        ws.cell(row=i, column=2, value=float(e["Largo (m)"])).number_format = "0.00"
        ws.cell(row=i, column=3, value=float(e["Ancho (m)"])).number_format = "0.00"
        ws.cell(row=i, column=4, value=float(e["Altura (m)"])).number_format = "0.00"

        c_paredes = ws.cell(row=i, column=5, value=f"=(B{i}+C{i})*2*D{i}")
        c_suelo = ws.cell(row=i, column=6, value=f"=B{i}*C{i}")
        c_techo = ws.cell(row=i, column=7, value=f"=B{i}*C{i}")

        for c in (c_paredes, c_suelo, c_techo):
            c.number_format = "0.00"

    # --- Fila de Totales con =SUM(...) ---
    fila_primera_dato = fila_inicio_tabla + 1
    fila_ultima_dato = fila_inicio_tabla + len(estancias)
    fila_total = fila_ultima_dato + 1

    ws.cell(row=fila_total, column=1, value="TOTALES GENERALES").font = Font(
        bold=True
    )
    for col_idx, col_letra in [(5, "E"), (6, "F"), (7, "G")]:
        cell = ws.cell(
            row=fila_total,
            column=col_idx,
            value=f"=SUM({col_letra}{fila_primera_dato}:{col_letra}{fila_ultima_dato})",
        )
        cell.font = Font(bold=True)
        cell.number_format = "0.00"

    # Ajuste automático de ancho de columnas (ignorando filas de título A1:A4)
    for col_idx in range(1, len(headers) + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(fila_inicio_tabla, fila_total + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is not None and not str(val).startswith("="):
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# --- INTERFAZ ---
st.title("📐 Medición de Vivienda")

if "estancias" not in st.session_state:
    st.session_state.estancias = []

# Bloque de datos de cabecera
with st.expander("📋 Datos del proyecto / cabecera (Opcional)", expanded=False):
    c_cli, c_fec = st.columns(2)
    with c_cli:
        cliente = st.text_input("Nombre del cliente", placeholder="Ej: Juan Pérez")
    with c_fec:
        fecha_proyecto = st.date_input("Fecha", value=date.today())
    direccion = st.text_input(
        "Dirección de la vivienda", placeholder="Ej: Calle Mayor 12, 3ºB"
    )

fecha_formateada = fecha_proyecto.strftime("%d/%m/%Y")

with st.form("form_estancia", clear_on_submit=True):
    st.subheader("Añadir nueva estancia")

    opciones = (
        "Hall",
        "Pasillo 1",
        "Pasillo 2",
        "Cocina",
        "Baño",
        "Aseo",
        "Despensa",
        "Salón",
        "Dormitorio 1",
        "Dormitorio 2",
        "Dormitorio 3",
        "Terraza",
        "Galería",
        "Otra (escribir abajo)",
    )
    seleccion = st.selectbox("Estancia", opciones)
    nombre_custom = st.text_input("Nombre personalizado (si elegiste 'Otra')")

    # st.number_input despliega el teclado numérico en móviles y acepta coma o punto
    col1, col2, col3 = st.columns(3)
    with col1:
        largo = st.number_input(
            "Largo (m)", min_value=0.0, value=None, step=0.05, format="%.2f", placeholder="0,00"
        )
    with col2:
        ancho = st.number_input(
            "Ancho (m)", min_value=0.0, value=None, step=0.05, format="%.2f", placeholder="0,00"
        )
    with col3:
        altura = st.number_input(
            "Altura (m)", min_value=0.0, value=2.60, step=0.05, format="%.2f", placeholder="2,60"
        )

    enviado = st.form_submit_button("➕ Añadir estancia", use_container_width=True)

    if enviado:
        if seleccion == "Otra (escribir abajo)" and not nombre_custom.strip():
            st.error("Escribe un nombre personalizado para la estancia.")
        elif not largo or not ancho or not altura or largo <= 0 or ancho <= 0 or altura <= 0:
            st.error("Introduce medidas mayores que 0 en largo, ancho y altura.")
        else:
            nombre_final = (
                nombre_custom.strip()
                if seleccion == "Otra (escribir abajo)"
                else seleccion
            )
            paredes = (largo + ancho) * 2 * altura
            area_planta = largo * ancho

            st.session_state.estancias.append({
                "Estancia": nombre_final,
                "Largo (m)": round(float(largo), 2),
                "Ancho (m)": round(float(ancho), 2),
                "Altura (m)": round(float(altura), 2),
                "Paredes (m²)": round(paredes, 2),
                "Suelo (m²)": round(area_planta, 2),
                "Techo (m²)": round(area_planta, 2),
            })
            st.success(f"Añadida: {nombre_final}")

if st.session_state.estancias:
    st.subheader("Resumen de estancias medidas")
    st.caption("💡 Puedes modificar valores o eliminar filas directamente en la tabla.")

    df_estancias = pd.DataFrame(st.session_state.estancias)
    df_editado = st.data_editor(
        df_estancias,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Estancia": st.column_config.TextColumn("Estancia", required=True),
            "Largo (m)": st.column_config.NumberColumn("Largo (m)", min_value=0.01, format="%.2f"),
            "Ancho (m)": st.column_config.NumberColumn("Ancho (m)", min_value=0.01, format="%.2f"),
            "Altura (m)": st.column_config.NumberColumn("Altura (m)", min_value=0.01, format="%.2f"),
            "Paredes (m²)": st.column_config.NumberColumn("Paredes (m²)", disabled=True, format="%.2f"),
            "Suelo (m²)": st.column_config.NumberColumn("Suelo (m²)", disabled=True, format="%.2f"),
            "Techo (m²)": st.column_config.NumberColumn("Techo (m²)", disabled=True, format="%.2f"),
        },
        key="editor_estancias",
    )

    # Recalcular superficies si el usuario editó largo, ancho o alto en la tabla
    nuevas_estancias = []
    for _, row in df_editado.iterrows():
        if pd.notna(row["Largo (m)"]) and pd.notna(row["Ancho (m)"]) and pd.notna(row["Altura (m)"]):
            l_val = float(row["Largo (m)"])
            a_val = float(row["Ancho (m)"])
            h_val = float(row["Altura (m)"])
            nuevas_estancias.append({
                "Estancia": str(row["Estancia"] or "Estancia"),
                "Largo (m)": round(l_val, 2),
                "Ancho (m)": round(a_val, 2),
                "Altura (m)": round(h_val, 2),
                "Paredes (m²)": round((l_val + a_val) * 2 * h_val, 2),
                "Suelo (m²)": round(l_val * a_val, 2),
                "Techo (m²)": round(l_val * a_val, 2),
            })

    if nuevas_estancias != st.session_state.estancias:
        st.session_state.estancias = nuevas_estancias
        st.rerun()

if st.session_state.estancias:
    tot_paredes = sum(e["Paredes (m²)"] for e in st.session_state.estancias)
    tot_suelos = sum(e["Suelo (m²)"] for e in st.session_state.estancias)
    tot_techos = sum(e["Techo (m²)"] for e in st.session_state.estancias)

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Paredes", f"{tot_paredes:.2f} m²".replace(".", ","))
    c2.metric("Total Suelos", f"{tot_suelos:.2f} m²".replace(".", ","))
    c3.metric("Total Techos", f"{tot_techos:.2f} m²".replace(".", ","))

    st.divider()

    col_pdf, col_xlsx = st.columns(2)
    with col_pdf:
        st.download_button(
            label="📄 Descargar PDF (DIN A4)",
            data=generar_pdf_bytes(
                st.session_state.estancias,
                tot_paredes,
                tot_suelos,
                tot_techos,
                cliente,
                direccion,
                fecha_formateada,
            ),
            file_name="informe_calculo_vivienda.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    with col_xlsx:
        st.download_button(
            label="📊 Descargar Excel (.xlsx)",
            data=generar_excel_bytes(
                st.session_state.estancias, cliente, direccion, fecha_formateada
            ),
            file_name="informe_calculo_vivienda.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    col_del1, col_del2 = st.columns(2)
    with col_del1:
        if st.button("↩️ Borrar última estancia", use_container_width=True):
            st.session_state.estancias.pop()
            st.rerun()
    with col_del2:
        if st.button("🗑️ Vaciar todas las estancias", use_container_width=True):
            st.session_state.estancias.clear()
            st.rerun()
