from datetime import datetime, date
import calendar
import io
import json
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import streamlit as st
from supabase import create_client

# Configuración de la página en español
st.set_page_config(
    page_title="Gestión de Traslados - Diálisis", page_icon="🚗", layout="centered"
)

# --- CONEXIÓN A SUPABASE ---
try:
  supabase_url = st.secrets["SUPABASE_URL"].strip().rstrip("/")
  supabase_key = st.secrets["SUPABASE_KEY"].strip()
  supabase = create_client(supabase_url, supabase_key)
except Exception as e:
  st.error(
      "⚠️ Error de configuración: Faltan las credenciales de Supabase en los"
      f" Secrets de Streamlit. Detalle: {e}"
  )
  st.stop()

# --- ESTILOS VISUALES MINIMALISTAS Y TIPOGRAFÍA DE ALTA CALIDAD ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    html, body, [class*="css"], .stApp {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        background-color: #0D1117;
        color: #E6EDF3;
    }
    
    h1 {
        font-size: 1.6rem !important;
        font-weight: 700 !important;
        color: #F0F6FC !important;
        margin-bottom: 0.2rem !important;
        letter-spacing: -0.02em;
    }

    h2, h3 {
        font-size: 1.15rem !important;
        font-weight: 600 !important;
        color: #C9D1D9 !important;
        margin-top: 1rem !important;
        margin-bottom: 0.5rem !important;
    }

    /* Tarjetas de pacientes estilo dashboard minimalista */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #161B22 !important;
        border: 1px solid #30363D !important;
        border-radius: 12px !important;
        padding: 0.8rem 1rem !important;
        margin-bottom: 0.6rem !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.2);
    }

    /* Botones sutiles y modernos */
    div.stButton > button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        padding: 0.35rem 0.2rem !important;
        font-size: 0.82rem !important;
        border: 1px solid #30363D !important;
        background-color: #21262D !important;
        color: #C9D1D9 !important;
        transition: all 0.2s ease-in-out;
    }

    div.stButton > button:hover {
        border-color: #58A6FF !important;
        color: #58A6FF !important;
        background-color: #161B22 !important;
    }

    /* Links tipo píldoras/chips */
    .chip-link {
        display: inline-block;
        background-color: #21262D;
        border: 1px solid #30363D;
        border-radius: 20px;
        padding: 2px 10px;
        font-size: 0.78rem;
        color: #58A6FF !important;
        text-decoration: none;
        margin-right: 6px;
        margin-bottom: 4px;
        font-weight: 500;
    }

    /* Expander limpio */
    .streamlit-expanderHeader {
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        color: #F0F6FC !important;
    }
    
    /* Disminuir margen de columnas móviles */
    [data-testid="column"] {
        width: fit-content !important;
        flex: 1 !important;
        min-width: 0 !important;
    }
    </style>
""", unsafe_allow_html=True)

TARIFA_COMPLETO = 5000
TARIFA_MEDIO = 2500
TURNOS_DISPONIBLES = ["05:40", "08:40"]

DEFAULT_MWF = [
    {"nombre": "Duarte Catalina", "turno": "05:40", "maps_link": "", "contactos": []},
    {"nombre": "Ledesma Tatiana", "turno": "05:40", "maps_link": "", "contactos": []},
    {"nombre": "Mendoza Ramón", "turno": "05:40", "maps_link": "", "contactos": []},
    {"nombre": "Meza Adriana", "turno": "08:40", "maps_link": "", "contactos": []},
    {"nombre": "Paiva Oscar", "turno": "08:40", "maps_link": "", "contactos": []},
]

DEFAULT_TTS = [
    {"nombre": "Cantero Miguel", "turno": "05:40", "maps_link": "", "contactos": []},
    {"nombre": "Martinez José", "turno": "05:40", "maps_link": "", "contactos": []},
    {"nombre": "Píriz Maximiliano", "turno": "05:40", "maps_link": "", "contactos": []},
    {"nombre": "Ramírez Orlando", "turno": "08:40", "maps_link": "", "contactos": []},
    {"nombre": "Villalba Nelida", "turno": "08:40", "maps_link": "", "contactos": []},
    {"nombre": "Cardozo José", "turno": "08:40", "maps_link": "", "contactos": []},
]


@st.cache_data(ttl=600)
def cargar_pacientes_db():
  try:
    response = supabase.table("pacientes").select("*").execute()
    data = response.data
    if not data:
      for p in DEFAULT_MWF:
        supabase.table("pacientes").insert({
            "grupo": "MWF",
            "nombre": p["nombre"],
            "turno": p["turno"],
            "maps_link": p["maps_link"],
            "contactos": p["contactos"],
        }).execute()
      for p in DEFAULT_TTS:
        supabase.table("pacientes").insert({
            "grupo": "TTS",
            "nombre": p["nombre"],
            "turno": p["turno"],
            "maps_link": p["maps_link"],
            "contactos": p["contactos"],
        }).execute()
      response = supabase.table("pacientes").select("*").execute()
      data = response.data
    return data
  except Exception as e:
    st.error(f"Error al conectar con la tabla pacientes: {e}")
    return DEFAULT_MWF + DEFAULT_TTS


def guardar_paciente_db(p_id, grupo, nombre, turno, maps_link, contactos):
  try:
    if p_id:
      supabase.table("pacientes").update({
          "grupo": grupo,
          "nombre": nombre,
          "turno": turno,
          "maps_link": maps_link,
          "contactos": contactos,
      }).eq("id", p_id).execute()
    else:
      supabase.table("pacientes").insert({
          "grupo": grupo,
          "nombre": nombre,
          "turno": turno,
          "maps_link": maps_link,
          "contactos": contactos,
      }).execute()
    st.cache_data.clear()
    return True, None
  except Exception as e:
    return False, str(e)


def eliminar_paciente_db(p_id):
  try:
    if p_id:
      supabase.table("pacientes").delete().eq("id", p_id).execute()
    st.cache_data.clear()
    return True, None
  except Exception as e:
    return False, str(e)


@st.cache_data(ttl=300)
def cargar_registros_db():
  try:
    response = supabase.table("registros").select("*").execute()
    data = response.data
    datos = {}
    if data:
      for row in data:
        f_str = row["fecha"]
        try:
          if datetime.strptime(f_str, "%d/%m/%Y").date() >= date(2026, 10, 14):
            if f_str not in datos:
              datos[f_str] = {}
            datos[f_str][row["nombre"]] = {
                "ida": row["ida"],
                "vuelta": row["vuelta"],
                "listo": row["listo"],
                "estado": row["estado"],
                "valor": row["valor"],
                "turno_dia": row.get("turno_dia", "05:40"),
            }
        except ValueError:
          pass
    return datos
  except Exception:
    return {}


def guardar_registro_db(fecha_str, nombre, info):
  try:
    existing = (
        supabase.table("registros")
        .select("id")
        .eq("fecha", fecha_str)
        .eq("nombre", nombre)
        .execute()
    )
    payload = {
        "fecha": fecha_str,
        "nombre": nombre,
        "ida": info["ida"],
        "vuelta": info["vuelta"],
        "listo": info["listo"],
        "estado": info["estado"],
        "valor": info["valor"],
        "turno_dia": info.get("turno_dia", "05:40"),
    }
    if existing.data:
      supabase.table("registros").update(payload).eq(
          "id", existing.data[0]["id"]
      ).execute()
    else:
      supabase.table("registros").insert(payload).execute()
    st.cache_data.clear()
  except Exception as e:
    st.error(f"Error al guardar registro: {e}")


def eliminar_registro_db(fecha_str, nombre):
  try:
    supabase.table("registros").delete().eq("fecha", fecha_str).eq(
        "nombre", nombre
    ).execute()
    st.cache_data.clear()
  except Exception as e:
    st.error(f"Error al eliminar registro: {e}")


def actualizar_valores_paciente(info):
  reg_ida = info.get("ida", False)
  reg_vuelta = info.get("vuelta", False)
  if reg_ida and reg_vuelta:
    info["estado"] = "Completo (Ida y Vuelta)"
    info["valor"] = TARIFA_COMPLETO
    info["listo"] = True
  elif reg_ida or reg_vuelta:
    info["estado"] = "Medio (Un solo tramo)"
    info["valor"] = TARIFA_MEDIO
  else:
    info["estado"] = "Ausente (No viajó)"
    info["valor"] = 0


def generar_excel_bytes(datos, es_diario=False):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Resumen Diario" if es_diario else "Resumen Mensual"

  encabezados = ["Fecha", "Paciente", "Estado", "Valor ($)"]
  ws.append(encabezados)

  header_fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  header_font = Font(color="FFFFFF", bold=True, size=11)
  thin_border = Border(
      left=Side(style="thin", color="D3D3D3"),
      right=Side(style="thin", color="D3D3D3"),
      top=Side(style="thin", color="D3D3D3"),
      bottom=Side(style="thin", color="D3D3D3"),
  )

  for col_num, header in enumerate(encabezados, 1):
    cell = ws.cell(row=1, column=col_num)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

  fila_actual = 2
  for fecha, pacientes_dia in sorted(
      datos.items(), key=lambda x: datetime.strptime(x[0], "%d/%m/%Y")
  ):
    try:
      if datetime.strptime(fecha, "%d/%m/%Y").date() < date(2026, 10, 14):
        continue
    except ValueError:
      continue

    for nombre, info in pacientes_dia.items():
      ws.cell(row=fila_actual, column=1, value=fecha).border = thin_border
      ws.cell(row=fila_actual, column=2, value=nombre).border = thin_border
      ws.cell(row=fila_actual, column=3, value=info["estado"]).border = (
          thin_border
      )

      valor_num = (
          int(info["valor"])
          if isinstance(info["valor"], (int, float, str))
          and str(info["valor"]).isdigit()
          else 0
      )
      celda_val = ws.cell(row=fila_actual, column=4, value=valor_num)
      celda_val.border = thin_border
      celda_val.number_format = '"$"#,##0'

      fila_actual += 1

  ws.cell(row=fila_actual, column=3, value="TOTAL GENERAL:").font = Font(
      bold=True
  )
  celda_total = ws.cell(
      row=fila_actual, column=4, value=f"=SUM(D2:D{fila_actual-1})"
  )
  celda_total.font = Font(bold=True, color="1F4E78")
  celda_total.border = thin_border
  celda_total.number_format = '"$"#,##0'

  column_widths = {"A": 15, "B": 25, "C": 30, "D": 15}
  for col, width in column_widths.items():
    ws.column_dimensions[col].width = width

  output = io.BytesIO()
  wb.save(output)
  output.seek(0)
  return output.getvalue()


def generar_excel_historial_detallado(historial_detalle):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Historial Detallado"

  encabezados = ["Paciente", "Fecha", "Estado / Tipo de Viaje", "Valor ($)"]
  ws.append(encabezados)

  header_fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  header_font = Font(color="FFFFFF", bold=True, size=11)
  thin_border = Border(
      left=Side(style="thin", color="D3D3D3"),
      right=Side(style="thin", color="D3D3D3"),
      top=Side(style="thin", color="D3D3D3"),
      bottom=Side(style="thin", color="D3D3D3"),
  )

  for col_num, header in enumerate(encabezados, 1):
    cell = ws.cell(row=1, column=col_num)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

  fila_actual = 2
  for item in historial_detalle:
    ws.cell(row=fila_actual, column=1, value=item["nombre"]).border = thin_border
    ws.cell(row=fila_actual, column=2, value=item["fecha"]).border = thin_border
    ws.cell(row=fila_actual, column=3, value=item["estado"]).border = thin_border

    celda_val = ws.cell(row=fila_actual, column=4, value=item["valor"])
    celda_val.border = thin_border
    celda_val.number_format = '"$"#,##0'

    fila_actual += 1

  column_widths = {"A": 25, "B": 15, "C": 30, "D": 15}
  for col, width in column_widths.items():
    ws.column_dimensions[col].width = width

  output = io.BytesIO()
  wb.save(output)
  output.seek(0)
  return output.getvalue()


def generar_pdf_historial_detallado(historial_detalle, titulo_reporte):
  buffer = io.BytesIO()
  doc = SimpleDocTemplate(
      buffer,
      pagesize=letter,
      rightMargin=30,
      leftMargin=30,
      topMargin=30,
      bottomMargin=30,
  )
  elements = []

  styles = getSampleStyleSheet()
  title_style = ParagraphStyle(
      "TitleStyle",
      parent=styles["Heading1"],
      fontName="Helvetica-Bold",
      fontSize=15,
      textColor=colors.HexColor("#1F4E78"),
      spaceAfter=10,
  )
  normal_style = ParagraphStyle(
      "NormalStyle",
      parent=styles["Normal"],
      fontName="Helvetica",
      fontSize=10,
      textColor=colors.HexColor("#333333"),
  )

  elements.append(Paragraph("Gestión de Traslados - Diálisis", title_style))
  elements.append(Paragraph(titulo_reporte, normal_style))
  elements.append(
      Paragraph(
          f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y')}",
          normal_style,
      )
  )
  elements.append(Spacer(1, 15))

  table_data = [["Paciente", "Fecha", "Estado / Tipo de Viaje", "Valor"]]
  for item in historial_detalle:
    table_data.append([
        item["nombre"],
        item["fecha"],
        item["estado"],
        f"${item['valor']:,}",
    ])

  t = Table(table_data, colWidths=[150, 90, 200, 100])
  t.setStyle(
      TableStyle([
          ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
          ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
          ("ALIGN", (0, 0), (-1, -1), "CENTER"),
          ("ALIGN", (0, 1), (0, -1), "LEFT"),
          ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
          ("FONTSIZE", (0, 0), (-1, 0), 10),
          ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
          ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F9F9F9")),
          ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D3D3D3")),
          ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
          ("FONTSIZE", (0, 1), (-1, -1), 9),
          ("TOPPADDING", (0, 1), (-1, -1), 5),
          ("BOTTOMPADDING", (0, 1), (-1, -1), 5),
      ])
  )

  elements.append(t)
  doc.build(elements)
  buffer.seek(0)
  return buffer.getvalue()


def generar_excel_planilla_matricial(datos_registros, todos_pacs, anio_sel, mes_sel):
  wb = openpyxl.Workbook()
  ws = wb.active

  nombres_meses_dict = {
      1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
      5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
      9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
  }
  nombre_mes_str = nombres_meses_dict.get(mes_sel, "Mes")
  ws.title = f"Planilla {nombre_mes_str}"

  header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
  total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
  header_font = Font(color="FFFFFF", bold=True, size=10)
  total_font = Font(bold=True, size=10, color="1F4E78")

  thin_border = Border(
      left=Side(style="thin", color="D3D3D3"),
      right=Side(style="thin", color="D3D3D3"),
      top=Side(style="thin", color="D3D3D3"),
      bottom=Side(style="thin", color="D3D3D3"),
  )

  mwf_nombres = [p["nombre"] for p in todos_pacs if p.get("grupo") == "MWF"]
  tts_nombres = [p["nombre"] for p in todos_pacs if p.get("grupo") == "TTS"]

  dias_mes_total = calendar.monthrange(anio_sel, mes_sel)[1]

  fechas_g1, dias_g1 = [], []
  fechas_g2, dias_g2 = [], []
  letras_dias = {0: "L", 1: "M", 2: "M", 3: "J", 4: "V", 5: "S", 6: "D"}

  for dia in range(1, dias_mes_total + 1):
    dt_obj = date(anio_sel, mes_sel, dia)
    w_day = dt_obj.weekday()
    if w_day in [0, 2, 4]:
      fechas_g1.append(dia)
      dias_g1.append(letras_dias[w_day])
    elif w_day in [1, 3, 5]:
      fechas_g2.append(dia)
      dias_g2.append(letras_dias[w_day])

  # Columnas para Totales: Total Viajes y Total Pesos
  col_total_viajes_g1 = len(fechas_g1) + 3
  col_total_pesos_g1 = len(fechas_g1) + 4

  # --- TABLA 1: GRUPO 1 ---
  titulo_tabla_1 = f"{nombre_mes_str.upper()}-{anio_sel}"
  ws.cell(row=1, column=1, value=titulo_tabla_1).font = Font(bold=True)

  ws.cell(row=1, column=2, value="Fecha").fill = header_fill
  for idx, d in enumerate(dias_g1, start=3):
    cell = ws.cell(row=1, column=idx, value=d)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")
    cell.border = thin_border

  ws.cell(row=2, column=2, value="Nombres").fill = header_fill
  for idx, f in enumerate(fechas_g1, start=3):
    cell = ws.cell(row=2, column=idx, value=f)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")
    cell.border = thin_border

  c_tv1 = ws.cell(row=1, column=col_total_viajes_g1, value="Total Viajes")
  c_tv1.fill = header_fill
  c_tv1.font = header_font
  c_tv1.alignment = Alignment(horizontal="center")
  c_tv1.border = thin_border

  ws.cell(row=2, column=col_total_viajes_g1, value="Viajes").fill = header_fill
  ws.cell(row=2, column=col_total_viajes_g1).font = header_font
  ws.cell(row=2, column=col_total_viajes_g1).alignment = Alignment(horizontal="center")
  ws.cell(row=2, column=col_total_viajes_g1).border = thin_border

  c_tp1 = ws.cell(row=1, column=col_total_pesos_g1, value="Total en Pesos")
  c_tp1.fill = header_fill
  c_tp1.font = header_font
  c_tp1.alignment = Alignment(horizontal="center")
  c_tp1.border = thin_border

  ws.cell(row=2, column=col_total_pesos_g1, value="Pesos ($)").fill = header_fill
  ws.cell(row=2, column=col_total_pesos_g1).font = header_font
  ws.cell(row=2, column=col_total_pesos_g1).alignment = Alignment(horizontal="center")
  ws.cell(row=2, column=col_total_pesos_g1).border = thin_border

  fila = 3
  inicio_g1 = fila
  for nombre in mwf_nombres:
    ws.cell(row=fila, column=2, value=nombre).border = thin_border
    for idx, f in enumerate(fechas_g1, start=3):
      f_str = f"{f:02d}/{mes_sel:02d}/{anio_sel}"
      val_celda = None
      if f_str in datos_registros and nombre in datos_registros[f_str]:
        info_reg = datos_registros[f_str][nombre]
        is_listo = info_reg.get("listo", False)
        reg_ida = info_reg.get("ida", False)
        reg_vuelta = info_reg.get("vuelta", False)

        if (reg_ida and reg_vuelta) or is_listo:
          val_celda = 1
        elif reg_ida or reg_vuelta:
          val_celda = 0.5

      c = ws.cell(row=fila, column=idx, value=val_celda)
      c.border = thin_border
      c.alignment = Alignment(horizontal="center")
      if val_celda is not None:
        c.number_format = '0.0'

    col_letra_inicio = openpyxl.utils.get_column_letter(3)
    col_letra_fin = openpyxl.utils.get_column_letter(len(fechas_g1) + 2)

    celda_tv = ws.cell(row=fila, column=col_total_viajes_g1, value=f"=SUM({col_letra_inicio}{fila}:{col_letra_fin}{fila})")
    celda_tv.border = thin_border
    celda_tv.alignment = Alignment(horizontal="center")
    celda_tv.font = Font(bold=True)
    celda_tv.number_format = '0.0'

    # Fórmula: 1 viaje completo = $5000, 0.5 viajes = $2500 -> Total Viajes * 5000
    celda_tp = ws.cell(row=fila, column=col_total_pesos_g1, value=f"={openpyxl.utils.get_column_letter(col_total_viajes_g1)}{fila}*{TARIFA_COMPLETO}")
    celda_tp.border = thin_border
    celda_tp.alignment = Alignment(horizontal="right")
    celda_tp.font = Font(bold=True)
    celda_tp.number_format = '"$"#,##0'

    fila += 1
  fin_g1 = fila - 1

  # Fila de Totales Generales Grupo 1
  ws.cell(row=fila, column=2, value="TOTAL GRUPO 1:").font = total_font
  ws.cell(row=fila, column=2).fill = total_fill
  ws.cell(row=fila, column=2).border = thin_border

  for idx, f in enumerate(fechas_g1, start=3):
    col_letra = openpyxl.utils.get_column_letter(idx)
    c_tot = ws.cell(row=fila, column=idx, value=f"=SUM({col_letra}{inicio_g1}:{col_letra}{fin_g1})")
    c_tot.fill = total_fill
    c_tot.font = total_font
    c_tot.alignment = Alignment(horizontal="center")
    c_tot.border = thin_border
    c_tot.number_format = '0.0'

  col_tv_letra_g1 = openpyxl.utils.get_column_letter(col_total_viajes_g1)
  c_tot_tv = ws.cell(row=fila, column=col_total_viajes_g1, value=f"=SUM({col_tv_letra_g1}{inicio_g1}:{col_tv_letra_g1}{fin_g1})")
  c_tot_tv.fill = total_fill
  c_tot_tv.font = total_font
  c_tot_tv.alignment = Alignment(horizontal="center")
  c_tot_tv.border = thin_border
  c_tot_tv.number_format = '0.0'

  col_tp_letra_g1 = openpyxl.utils.get_column_letter(col_total_pesos_g1)
  c_tot_tp = ws.cell(row=fila, column=col_total_pesos_g1, value=f"=SUM({col_tp_letra_g1}{inicio_g1}:{col_tp_letra_g1}{fin_g1})")
  c_tot_tp.fill = total_fill
  c_tot_tp.font = total_font
  c_tot_tp.alignment = Alignment(horizontal="right")
  c_tot_tp.border = thin_border
  c_tot_tp.number_format = '"$"#,##0'

  fila += 3

  # --- TABLA 2: GRUPO 2 ---
  col_total_viajes_g2 = len(fechas_g2) + 3
  col_total_pesos_g2 = len(fechas_g2) + 4

  ws.cell(row=fila, column=1, value=titulo_tabla_1).font = Font(bold=True)
  fila_c1 = fila
  fila_c2 = fila + 1

  ws.cell(row=fila_c1, column=2, value="Fecha").fill = header_fill
  for idx, d in enumerate(dias_g2, start=3):
    cell = ws.cell(row=fila_c1, column=idx, value=d)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")
    cell.border = thin_border

  ws.cell(row=fila_c2, column=2, value="Nombres").fill = header_fill
  for idx, f in enumerate(fechas_g2, start=3):
    cell = ws.cell(row=fila_c2, column=idx, value=f)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")
    cell.border = thin_border

  c_tv2 = ws.cell(row=fila_c1, column=col_total_viajes_g2, value="Total Viajes")
  c_tv2.fill = header_fill
  c_tv2.font = header_font
  c_tv2.alignment = Alignment(horizontal="center")
  c_tv2.border = thin_border

  ws.cell(row=fila_c2, column=col_total_viajes_g2, value="Viajes").fill = header_fill
  ws.cell(row=fila_c2, column=col_total_viajes_g2).font = header_font
  ws.cell(row=fila_c2, column=col_total_viajes_g2).alignment = Alignment(horizontal="center")
  ws.cell(row=fila_c2, column=col_total_viajes_g2).border = thin_border

  c_tp2 = ws.cell(row=fila_c1, column=col_total_pesos_g2, value="Total en Pesos")
  c_tp2.fill = header_fill
  c_tp2.font = header_font
  c_tp2.alignment = Alignment(horizontal="center")
  c_tp2.border = thin_border

  ws.cell(row=fila_c2, column=col_total_pesos_g2, value="Pesos ($)").fill = header_fill
  ws.cell(row=fila_c2, column=col_total_pesos_g2).font = header_font
  ws.cell(row=fila_c2, column=col_total_pesos_g2).alignment = Alignment(horizontal="center")
  ws.cell(row=fila_c2, column=col_total_pesos_g2).border = thin_border

  fila_datos = fila_c2 + 1
  inicio_g2 = fila_datos
  for nombre in tts_nombres:
    ws.cell(row=fila_datos, column=2, value=nombre).border = thin_border
    for idx, f in enumerate(fechas_g2, start=3):
      f_str = f"{f:02d}/{mes_sel:02d}/{anio_sel}"
      val_celda = None
      if f_str in datos_registros and nombre in datos_registros[f_str]:
        info_reg = datos_registros[f_str][nombre]
        is_listo = info_reg.get("listo", False)
        reg_ida = info_reg.get("ida", False)
        reg_vuelta = info_reg.get("vuelta", False)

        if (reg_ida and reg_vuelta) or is_listo:
          val_celda = 1
        elif reg_ida or reg_vuelta:
          val_celda = 0.5

      c = ws.cell(row=fila_datos, column=idx, value=val_celda)
      c.border = thin_border
      c.alignment = Alignment(horizontal="center")
      if val_celda is not None:
        c.number_format = '0.0'

    col_letra_inicio_g2 = openpyxl.utils.get_column_letter(3)
    col_letra_fin_g2 = openpyxl.utils.get_column_letter(len(fechas_g2) + 2)

    celda_tv = ws.cell(row=fila_datos, column=col_total_viajes_g2, value=f"=SUM({col_letra_inicio_g2}{fila_datos}:{col_letra_fin_g2}{fila_datos})")
    celda_tv.border = thin_border
    celda_tv.alignment = Alignment(horizontal="center")
    celda_tv.font = Font(bold=True)
    celda_tv.number_format = '0.0'

    celda_tp = ws.cell(row=fila_datos, column=col_total_pesos_g2, value=f"={openpyxl.utils.get_column_letter(col_total_viajes_g2)}{fila_datos}*{TARIFA_COMPLETO}")
    celda_tp.border = thin_border
    celda_tp.alignment = Alignment(horizontal="right")
    celda_tp.font = Font(bold=True)
    celda_tp.number_format = '"$"#,##0'

    fila_datos += 1
  fin_g2 = fila_datos - 1

  # Fila de Totales Generales Grupo 2
  ws.cell(row=fila_datos, column=2, value="TOTAL GRUPO 2:").font = total_font
  ws.cell(row=fila_datos, column=2).fill = total_fill
  ws.cell(row=fila_datos, column=2).border = thin_border

  for idx, f in enumerate(fechas_g2, start=3):
    col_letra = openpyxl.utils.get_column_letter(idx)
    c_tot = ws.cell(row=fila_datos, column=idx, value=f"=SUM({col_letra}{inicio_g2}:{col_letra}{fin_g2})")
    c_tot.fill = total_fill
    c_tot.font = total_font
    c_tot.alignment = Alignment(horizontal="center")
    c_tot.border = thin_border
    c_tot.number_format = '0.0'

  col_tv_letra_g2 = openpyxl.utils.get_column_letter(col_total_viajes_g2)
  c_tot_tv = ws.cell(row=fila_datos, column=col_total_viajes_g2, value=f"=SUM({col_tv_letra_g2}{inicio_g2}:{col_tv_letra_g2}{fin_g2})")
  c_tot_tv.fill = total_fill
  c_tot_tv.font = total_font
  c_tot_tv.alignment = Alignment(horizontal="center")
  c_tot_tv.border = thin_border
  c_tot_tv.number_format = '0.0'

  col_tp_letra_g2 = openpyxl.utils.get_column_letter(col_total_pesos_g2)
  c_tot_tp = ws.cell(row=fila_datos, column=col_total_pesos_g2, value=f"=SUM({col_tp_letra_g2}{inicio_g2}:{col_tp_letra_g2}{fin_g2})")
  c_tot_tp.fill = total_fill
  c_tot_tp.font = total_font
  c_tot_tp.alignment = Alignment(horizontal="right")
  c_tot_tp.border = thin_border
