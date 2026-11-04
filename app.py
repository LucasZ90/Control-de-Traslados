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

# --- ESTILOS VISUALES MINIMALISTAS ---
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

    [data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #161B22 !important;
        border: 1px solid #30363D !important;
        border-radius: 12px !important;
        padding: 0.8rem 1rem !important;
        margin-bottom: 0.6rem !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.2);
    }

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

    .streamlit-expanderHeader {
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        color: #F0F6FC !important;
    }
    
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
    return data if data else []
  except Exception:
    fallback = []
    for i, p in enumerate(DEFAULT_MWF):
      fallback.append({**p, "id": f"def_mwf_{i}", "grupo": "MWF"})
    for i, p in enumerate(DEFAULT_TTS):
      fallback.append({**p, "id": f"def_tts_{i}", "grupo": "TTS"})
    return fallback


def guardar_paciente_db(p_id, grupo, nombre, turno, maps_link, contactos):
  try:
    if p_id and not str(p_id).startswith("def_"):
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
    if p_id and not str(p_id).startswith("def_"):
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
        f_str = row.get("fecha")
        nombre = row.get("nombre")
        if not f_str or not nombre:
          continue
        try:
          if datetime.strptime(f_str, "%d/%m/%Y").date() >= date(2026, 10, 14):
            if f_str not in datos:
              datos[f_str] = {}
            datos[f_str][nombre] = {
                "ida": bool(row.get("ida", False)),
                "vuelta": bool(row.get("vuelta", False)),
                "listo": bool(row.get("listo", False)),
                "estado": row.get("estado", "Ausente (No viajó)"),
                "valor": int(row.get("valor", 0) or 0),
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
    info["listo"] = False
  else:
    info["estado"] = "Ausente (No viajó)"
    info["valor"] = 0
    info["listo"] = False


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
    for nombre, info in pacientes_dia.items():
      ws.cell(row=fila_actual, column=1, value=fecha).border = thin_border
      ws.cell(row=fila_actual, column=2, value=nombre).border = thin_border
      ws.cell(row=fila_actual, column=3, value=info["estado"]).border = thin_border

      valor_num = int(info["valor"] or 0)
      celda_val = ws.cell(row=fila_actual, column=4, value=valor_num)
      celda_val.border = thin_border
      celda_val.number_format = '"$"#,##0'

      fila_actual += 1

  ws.cell(row=fila_actual, column=3, value="TOTAL GENERAL:").font = Font(bold=True)
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
      val_num = None

      if f_str in datos_registros and nombre in datos_registros[f_str]:
        info_reg = datos_registros[f_str][nombre]
        reg_ida = info_reg.get("ida", False)
        reg_vuelta = info_reg.get("vuelta", False)

        if reg_ida and reg_vuelta:
          val_num = 1
        elif reg_ida or reg_vuelta:
          val_num = 0.5

      c = ws.cell(row=fila, column=idx, value=val_num)
      c.border = thin_border
      c.alignment = Alignment(horizontal="center")
      if val_num is not None:
        c.number_format = '0.0'

    col_letra_inicio = openpyxl.utils.get_column_letter(3)
    col_letra_fin = openpyxl.utils.get_column_letter(len(fechas_g1) + 2)
    col_letra_tv = openpyxl.utils.get_column_letter(col_total_viajes_g1)

    celda_tv = ws.cell(row=fila, column=col_total_viajes_g1, value=f"=SUM({col_letra_inicio}{fila}:{col_letra_fin}{fila})")
    celda_tv.border = thin_border
    celda_tv.alignment = Alignment(horizontal="center")
    celda_tv.font = Font(bold=True)
    celda_tv.number_format = '0.0'

    celda_tp = ws.cell(row=fila, column=col_total_pesos_g1, value=f"={col_letra_tv}{fila}*5000")
    celda_tp.border = thin_border
    celda_tp.alignment = Alignment(horizontal="right")
    celda_tp.font = Font(bold=True)
    celda_tp.number_format = '"$"#,##0'

    fila += 1
  fin_g1 = fila - 1

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
      val_num = None

      if f_str in datos_registros and nombre in datos_registros[f_str]:
        info_reg = datos_registros[f_str][nombre]
        reg_ida = info_reg.get("ida", False)
        reg_vuelta = info_reg.get("vuelta", False)

        if reg_ida and reg_vuelta:
          val_num = 1
        elif reg_ida or reg_vuelta:
          val_num = 0.5

      c = ws.cell(row=fila_datos, column=idx, value=val_num)
      c.border = thin_border
      c.alignment = Alignment(horizontal="center")
      if val_num is not None:
        c.number_format = '0.0'

    col_letra_inicio_g2 = openpyxl.utils.get_column_letter(3)
    col_letra_fin_g2 = openpyxl.utils.get_column_letter(len(fechas_g2) + 2)
    col_letra_tv_g2 = openpyxl.utils.get_column_letter(col_total_viajes_g2)

    celda_tv = ws.cell(row=fila_datos, column=col_total_viajes_g2, value=f"=SUM({col_letra_inicio_g2}{fila_datos}:{col_letra_fin_g2}{fila_datos})")
    celda_tv.border = thin_border
    celda_tv.alignment = Alignment(horizontal="center")
    celda_tv.font = Font(bold=True)
    celda_tv.number_format = '0.0'

    celda_tp = ws.cell(row=fila_datos, column=col_total_pesos_g2, value=f"={col_letra_tv_g2}{fila_datos}*5000")
    celda_tp.border = thin_border
    celda_tp.alignment = Alignment(horizontal="right")
    celda_tp.font = Font(bold=True)
    celda_tp.number_format = '"$"#,##0'

    fila_datos += 1
  fin_g2 = fila_datos - 1

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
  c_tot_tp.number_format = '"$"#,##0'

  ws.column_dimensions["A"].width = 5
  ws.column_dimensions["B"].width = 25
  ws.column_dimensions[openpyxl.utils.get_column_letter(col_total_viajes_g1)].width = 15
  ws.column_dimensions[openpyxl.utils.get_column_letter(col_total_pesos_g1)].width = 18

  output = io.BytesIO()
  wb.save(output)
  output.seek(0)
  return output.getvalue()


# Título principal limpio
st.title("Gestión de Traslados")

# --- NAVEGACIÓN SUPERIOR REDUCIDA ---
if "menu_actual" not in st.session_state:
  st.session_state["menu_actual"] = "Registrar / Actualizar Viajes"

col_nav1, col_nav2, col_nav3, col_nav4 = st.columns(4)
with col_nav1:
  if st.button("📋 Registrar", use_container_width=True):
    st.session_state["menu_actual"] = "Registrar / Actualizar Viajes"
with col_nav2:
  if st.button("📊 Resumen", use_container_width=True):
    st.session_state["menu_actual"] = "Resumen Mensual"
with col_nav3:
  if st.button("📂 Historial", use_container_width=True):
    st.session_state["menu_actual"] = "Historial de Pacientes"
with col_nav4:
  if st.button("📅 Planilla", use_container_width=True):
    st.session_state["menu_actual"] = "Planilla Matricial"

menu = st.session_state["menu_actual"]

datos = cargar_registros_db()
todos_pacientes = cargar_pacientes_db()

if menu == "Registrar / Actualizar Viajes":
  with st.expander("➕ Agregar Nuevo Paciente al Sistema"):
    nuevo_nombre = st.text_input(
        "Nombre y Apellido del Paciente", key="nuevo_nombre_paciente"
    )
    nuevo_grupo = st.selectbox(
        "Grupo por Defecto",
        [
            "Grupo 1 (Lunes, Miércoles, Viernes)",
            "Grupo 2 (Martes, Jueves, Sábados)",
        ],
        key="nuevo_grupo_paciente",
    )
    nuevo_turno = st.selectbox(
        "Horario Habitual", TURNOS_DISPONIBLES, key="nuevo_turno_paciente"
    )
    nuevo_maps = st.text_input(
        "Enlace de Google Maps (Opcional)", key="nuevo_maps_paciente"
    )
    nuevo_tel = st.text_input(
        "Número de Teléfono Principal (Opcional)", key="nuevo_tel_paciente"
    )
    nueva_rel_tel = st.text_input(
        "Relación de este teléfono (Ej: Paciente, Hijo/a)",
        value="Paciente",
        key="nuevo_rel_paciente",
    )

    if st.button(
        "💾 Guardar Nuevo Paciente",
        type="primary",
        use_container_width=True,
        key="btn_guardar_nuevo_paciente",
    ):
      if nuevo_nombre.strip() == "":
        st.error("Por favor, escribe el nombre del paciente.")
      else:
        lista_contactos = []
        if nuevo_tel.strip():
          lista_contactos.append({
              "numero": nuevo_tel.strip(),
              "relacion": (
                  nueva_rel_tel.strip()
                  if nueva_rel_tel.strip()
                  else "Paciente"
              ),
          })

        nuevo_grupo_val = "MWF" if "Grupo 1" in nuevo_grupo else "TTS"
        ok, err = guardar_paciente_db(
            None,
            nuevo_grupo_val,
            nuevo_nombre.strip(),
            nuevo_turno,
            nuevo_maps.strip(),
            lista_contactos,
        )
        if ok:
          st.success(f"¡Paciente '{nuevo_nombre}' agregado con éxito!")
          st.rerun()
        else:
          st.error(f"Error al guardar nuevo paciente: {err}")

  fecha_seleccionada = st.date_input("Selecciona la fecha", datetime.now())
  fecha_str = fecha_seleccionada.strftime("%d/%m/%Y")

  if fecha_seleccionada < date(2026, 10, 14):
    st.error(
        "⚠️ Los registros de traslados comienzan a partir del 14 de octubre de"
        " 2026."
    )
  else:
    dia_semana = fecha_seleccionada.weekday()

    if dia_semana == 6:
      st.warning("⚠️ Los domingos no hay servicio programado.")
    else:
      if fecha_str not in datos:
        datos[fecha_str] = {}

      with st.expander(
          "⚙️ Seleccionar pacientes y turnos de hoy", expanded=False
      ):
        st.write(
            "Marca únicamente los pacientes que viajan hoy y ajusta su turno:"
        )
        for p in todos_pacientes:
          p_nombre = p["nombre"]
          p_id = p.get("id", "temp_id")

          ya_marcado_hoy = p_nombre in datos[fecha_str]

          cols_sel = st.columns([3, 2])
          seleccionado = cols_sel[0].checkbox(
              f"{p_nombre} ({p.get('turno','05:40')})",
              value=ya_marcado_hoy,
              key=f"chk_sel_{fecha_str}_{p_id}",
          )

          turno_actual_dia = (
              datos[fecha_str][p_nombre].get(
                  "turno_dia", p.get("turno", "05:40")
              )
              if ya_marcado_hoy
              else p.get("turno", "05:40")
          )
          idx_turno = (
              TURNOS_DISPONIBLES.index(turno_actual_dia)
              if turno_actual_dia in TURNOS_DISPONIBLES
              else 0
          )

          turno_elegido = cols_sel[1].selectbox(
              "Turno hoy",
              TURNOS_DISPONIBLES,
              index=idx_turno,
              key=f"sel_trn_{fecha_str}_{p_id}",
              label_visibility="collapsed",
          )

          if seleccionado:
            if p_nombre not in datos[fecha_str]:
              datos[fecha_str][p_nombre] = {
                  "ida": False,
                  "vuelta": False,
                  "listo": False,
                  "estado": "Ausente (No viajó)",
                  "valor": 0,
                  "turno_dia": turno_elegido,
              }
            else:
              datos[fecha_str][p_nombre]["turno_dia"] = turno_elegido
            guardar_registro_db(fecha_str, p_nombre, datos[fecha_str][p_nombre])
          else:
            if p_nombre in datos[fecha_str]:
              del datos[fecha_str][p_nombre]
              eliminar_registro_db(fecha_str, p_nombre)

      pacientes_hoy_dict = datos.get(fecha_str, {})

      def esta_finalizado(p_nombre):
        info = pacientes_hoy_dict.get(p_nombre, {})
        return (info.get("ida", False) and info.get("vuelta", False)) or info.get(
            "listo", False
        )

      pendientes_totales = sum(
          1 for nom in pacientes_hoy_dict if not esta_finalizado(nom)
      )
      st.metric(label="👥 Pacientes Pendientes Hoy", value=pendientes_totales)

      for turno_turno, titulo_turno in [
          ("05:40", "⏰ Turno 05:40 AM"),
          ("08:40", "⏰ Turno 08:40 AM"),
      ]:
        st.markdown(f"### {titulo_turno}")

        pacientes_turno = []
        for p in todos_pacientes:
          nombre_p = p["nombre"]
          if nombre_p in pacientes_hoy_dict:
            if (
                pacientes_hoy_dict[nombre_p].get("turno_dia", "05:40")
                == turno_turno
            ):
              pacientes_turno.append(p)

        if not pacientes_turno:
          st.write("*(Sin pacientes programados en este turno)*")
        else:
          pacientes_turno.sort(
              key=lambda p: (
                  1
                  if pacientes_hoy_dict[p["nombre"]].get("listo", False)
                  else 0,
                  p["nombre"],
              )
          )

        for p in pacientes_turno:
          nombre = p["nombre"]
          p_id = p.get("id", "temp_id")
          info_paciente = pacientes_hoy_dict[nombre]
          reg_ida = info_paciente.get("ida", False)
          reg_vuelta = info_paciente.get("vuelta", False)
          is_listo = info_paciente.get("listo", False)

          with st.container(border=True):
            estado_icono = (
                "🟢"
                if (reg_ida and reg_vuelta)
                else ("🟡" if (reg_ida or reg_vuelta or is_listo) else "⚪")
            )
            finalizado_txt = " *[Finalizado]*" if is_listo else ""

            with st.expander(
                f"{estado_icono} {nombre}{finalizado_txt}", expanded=False
            ):
              st.markdown(f"**Editando datos de {nombre}:**")
              edit_nombre = st.text_input(
                  "Nombre y Apellido", value=p["nombre"], key=f"edit_nom_{p_id}"
              )
              edit_grupo = st.selectbox(
                  "Grupo Asignado",
                  [
                      "Grupo 1 (Lunes, Miércoles, Viernes)",
                      "Grupo 2 (Martes, Jueves, Sábados)",
                  ],
                  index=0 if p.get("grupo", "MWF") == "MWF" else 1,
                  key=f"edit_grp_{p_id}",
              )
              edit_turno = st.selectbox(
                  "Turno Habitual",
                  TURNOS_DISPONIBLES,
                  index=(
                      TURNOS_DISPONIBLES.index(p["turno"])
                      if p["turno"] in TURNOS_DISPONIBLES
                      else 0
                  ),
                  key=f"edit_trn_{p_id}",
              )
              edit_maps = st.text_input(
                  "Enlace de Google Maps",
                  value=p.get("maps_link", ""),
                  key=f"edit_map_{p_id}",
              )

              st.markdown("**📞 Teléfonos y Contactos:**")
              contactos_actuales = p.get("contactos", [])
              for idx_c, contacto in enumerate(contactos_actuales):
                cols_c = st.columns([3, 3, 1])
                cols_c[0].write(f"📞 {contacto['numero']}")
                cols_c[1].write(f"({contacto['relacion']})")
                if cols_c[2].button(
                    "🗑️", key=f"del_c_{p_id}_{idx_c}", use_container_width=True
                ):
                  contactos_actuales.pop(idx_c)
                  guardar_paciente_db(
                      p_id,
                      p.get("grupo", "MWF"),
                      p["nombre"],
                      p["turno"],
                      p["maps_link"],
                      contactos_actuales,
                  )
                  st.rerun()

              nuevo_num = st.text_input(
                  "Nuevo número de teléfono", key=f"new_num_{p_id}"
              )
              nueva_rel = st.text_input(
                  "Relación (Ej: Paciente, Hijo/a)", key=f"new_rel_{p_id}"
              )

              if st.button(
                  "➕ Añadir Contacto",
                  key=f"add_btn_{p_id}",
                  use_container_width=True,
              ):
                if nuevo_num.strip():
                  contactos_actuales.append({
                      "numero": nuevo_num.strip(),
                      "relacion": (
                          nueva_rel.strip()
                          if nueva_rel.strip()
                          else "Contacto"
                      ),
                  })
                  guardar_paciente_db(
                      p_id,
                      p.get("grupo", "MWF"),
                      p["nombre"],
                      p["turno"],
                      p["maps_link"],
                      contactos_actuales,
                  )
                  st.success("¡Contacto añadido!")
                  st.rerun()
                else:
                  st.error("Ingresa un número.")

              col_save, col_del = st.columns([4, 1])
              if col_save.button(
                  "💾 Guardar Cambios",
                  key=f"save_{p_id}",
                  type="primary",
                  use_container_width=True,
              ):
                nuevo_grupo_val = "MWF" if "Grupo 1" in edit_grupo else "TTS"
                ok, err = guardar_paciente_db(
                    p_id,
                    nuevo_grupo_val,
                    edit_nombre.strip(),
                    edit_turno,
                    edit_maps.strip(),
                    contactos_actuales,
                )
                if ok:
                  st.success("¡Guardado!")
                  st.rerun()
                else:
                  st.error(f"Error: {err}")

              if col_del.button(
                  "🗑️",
                  key=f"del_p_{p_id}",
                  use_container_width=True,
                  help="Eliminar paciente",
              ):
                ok, err = eliminar_paciente_db(p_id)
                if ok:
                  st.rerun()
                else:
                  st.error(f"Error: {err}")

            maps_url = p.get("maps_link", "").strip()
            contactos = [c for c in p.get("contactos", []) if c.get("numero")]

            html_chips = "<div style='margin-bottom: 8px; margin-top: 4px;'>"
            if maps_url:
              html_chips += f"<a href='{maps_url}' target='_blank' class='chip-link'>📍 Ubicación</a>"
            else:
              html_chips += "<span style='font-size: 0.78rem; color: #8B949E; margin-right: 8px;'>📍 Sin ubicación</span>"

            if contactos:
              for c in contactos:
                num_clean = c["numero"].replace(" ", "").replace("-", "")
                html_chips += f"<a href='tel:{num_clean}' class='chip-link'>📞 {c['relacion']}</a>"
            else:
              html_chips += "<span style='font-size: 0.78rem; color: #8B949E;'>📞 Sin contacto</span>"

            html_chips += "</div>"
            st.markdown(html_chips, unsafe_allow_html=True)

            col_b1, col_b2, col_b3 = st.columns([2.2, 2.2, 1.6])

            label_ida = "✅ Ida" if reg_ida else "❌ Ida"
            if col_b1.button(
                label_ida,
                key=f"t_{turno_turno}_i_{p_id}_{fecha_str}",
                use_container_width=True,
                disabled=is_listo,
            ):
              info_paciente["ida"] = not reg_ida
              actualizar_valores_paciente(info_paciente)
              guardar_registro_db(fecha_str, nombre, info_paciente)
              st.rerun()

            label_vuelta = "✅ Vuelta" if reg_vuelta else "❌ Vuelta"
            if col_b2.button(
                label_vuelta,
                key=f"t_{turno_turno}_v_{p_id}_{fecha_str}",
                use_container_width=True,
                disabled=is_listo,
            ):
              info_paciente["vuelta"] = not reg_vuelta
              actualizar_valores_paciente(info_paciente)
              guardar_registro_db(fecha_str, nombre, info_paciente)
              st.rerun()

            if not is_listo:
              if col_b3.button(
                  "🏁 Fin",
                  key=f"t_{turno_turno}_f_{p_id}_{fecha_str}",
                  use_container_width=True,
              ):
                info_paciente["listo"] = True
                guardar_registro_db(fecha_str, nombre, info_paciente)
                st.rerun()
            else:
              if col_b3.button(
                  "🔄 Abrir",
                  key=f"t_{turno_turno}_r_{p_id}_{fecha_str}",
                  use_container_width=True,
              ):
                info_paciente["listo"] = False
                guardar_registro_db(fecha_str, nombre, info_paciente)
                st.rerun()

      if fecha_str in datos and datos[fecha_str]:
        datos_del_dia = {fecha_str: datos[fecha_str]}
        excel_bytes_dia = generar_excel_bytes(datos_del_dia, es_diario=True)
        nombre_archivo_dia = (
            f"resumen_{fecha_seleccionada.strftime('%d_%m_%Y')}.xlsx"
        )
        st.download_button(
            label=f"📊 Descargar Resumen del Día ({fecha_str})",
            data=excel_bytes_dia,
            file_name=nombre_archivo_dia,
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )

elif menu == "Resumen Mensual":
  st.header("📊 Resumen por Meses y Ganancias")

  meses_agrupados = {}
  for f_str, reg in datos.items():
    try:
      f_obj = datetime.strptime(f_str, "%d/%m/%Y")
      if f_obj.date() >= date(2026, 10, 14):
        mes_key_orden = f_obj.strftime("%Y-%m")
        nombres_meses = {
            "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
            "05": "Mayo", "06": "Junio", "07": "Julio", "08": "Agosto",
            "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
        }
        anio, mes_num = mes_key_orden.split("-")
        nombre_mes_amigable = f"{nombres_meses[mes_num]} {anio}"

        if mes_key_orden not in meses_agrupados:
          meses_agrupados[mes_key_orden] = {
              "nombre_amigable": nombre_mes_amigable,
              "registros_dias": {}
          }
        meses_agrupados[mes_key_orden]["registros_dias"][f_str] = reg
    except ValueError:
      pass

  if not meses_agrupados:
    st.info(
        "No hay registros de traslados (comenzando desde el 14/10/2026)."
    )
  else:
    for mes_key in sorted(meses_agrupados.keys(), reverse=True):
      info_mes = meses_agrupados[mes_key]
      nombre_mes = info_mes["nombre_amigable"]
      registros_dias = info_mes["registros_dias"]

      total_mes = 0
      for fecha, pacientes_dia in registros_dias.items():
        for nombre, info in pacientes_dia.items():
          total_mes += int(info.get("valor", 0) or 0)

      with st.expander(f"📅 Mes: {nombre_mes} — Total: ${total_mes:,}", expanded=(mes_key == max(meses_agrupados.keys()))):
        for fecha, pacientes_dia in sorted(
            registros_dias.items(), key=lambda x: datetime.strptime(x[0], "%d/%m/%Y")
        ):
          st.markdown(f"**Día: {fecha}**")
          total_dia = 0
          for nombre, info in pacientes_dia.items():
            val_v = int(info.get("valor", 0) or 0)
            st.write(f"- {nombre}: {info.get('estado', '')} (${val_v:,})")
            total_dia += val_v
          st.markdown(f"_Subtotal día: ${total_dia:,}_")
          st.markdown("---")

        st.markdown(f"### 💰 Total acumulado de {nombre_mes}: ${total_mes:,}")

        excel_bytes_mes = generar_excel_bytes(registros_dias, es_diario=False)
        st.download_button(
            label=f"📊 Descargar reporte de {nombre_mes} (.xlsx)",
            data=excel_bytes_mes,
            file_name=f"control_dialisis_{mes_key}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            type="primary",
            key=f"dl_mes_{mes_key}",
            use_container_width=True,
        )

elif menu == "Historial de Pacientes":
  st.header("📂 Historial Detallado de Pacientes")
  st.write("Consulta los días específicos que viajó cada paciente, sus totales y exporta los informes.")

  historial_detalle = []
  resumen_pacientes = {}

  for p in todos_pacientes:
    nombre_p = p.get("nombre", "")
    if nombre_p:
      resumen_pacientes[nombre_p] = {
          "completos": 0, "medios": 0, "ausencias": 0, "total_viajes": 0, "total_monto": 0
      }

  for fecha, pacientes_dia in sorted(
      datos.items(), key=lambda x: datetime.strptime(x[0], "%d/%m/%Y")
  ):
    for nombre, info in pacientes_dia.items():
      if nombre not in resumen_pacientes:
        resumen_pacientes[nombre] = {
            "completos": 0, "medios": 0, "ausencias": 0, "total_viajes": 0, "total_monto": 0
        }
      
      estado = info.get("estado", "")
      valor = int(info.get("valor", 0) or 0)

      if "Completo" in estado:
        resumen_pacientes[nombre]["completos"] += 1
        resumen_pacientes[nombre]["total_viajes"] += 1
      elif "Medio" in estado:
        resumen_pacientes[nombre]["medios"] += 1
        resumen_pacientes[nombre]["total_viajes"] += 1
      else:
        resumen_pacientes[nombre]["ausencias"] += 1

      resumen_pacientes[nombre]["total_monto"] += valor

      historial_detalle.append({
          "nombre": nombre,
          "fecha": fecha,
          "estado": estado,
          "valor": valor
      })

  modo_descarga = st.radio(
      "Selecciona el alcance del reporte:",
      ["Todos juntos", "Paciente individual"],
      horizontal=True
  )

  paciente_seleccionado = None
  if modo_descarga == "Paciente individual":
    nombres_disponibles = sorted(list(resumen_pacientes.keys()))
    if nombres_disponibles:
      paciente_seleccionado = st.selectbox("Elige al paciente", nombres_disponibles)
    else:
      st.warning("No hay pacientes registrados.")

  st.markdown("---")

  if modo_descarga == "Todos juntos":
    for nombre, datos_p in sorted(resumen_pacientes.items()):
      with st.container(border=True):
        st.markdown(f"### 👤 {nombre}")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("🟢 Completos", datos_p["completos"])
        col2.metric("🟡 Medios", datos_p["medios"])
        col3.metric("🚗 Total Viajes", datos_p["total_viajes"])
        col4.metric("💰 Acumulado", f"${datos_p['total_monto']:,}")

        viajes_paciente = [item for item in historial_detalle if item["nombre"] == nombre]
        if viajes_paciente:
          with st.expander(f"Ver días y fechas de {nombre} ({len(viajes_paciente)} registros)"):
            for v in viajes_paciente:
              st.write(f"• **{v['fecha']}** — {v['estado']} (*${v['valor']:,}*)")
        else:
          st.write("*(Sin registros de viajes)*")

    excel_bytes = generar_excel_historial_detallado(historial_detalle)
    pdf_bytes = generar_pdf_historial_detallado(historial_detalle, "Historial Detallado de Todos los Pacientes")

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
      st.download_button(
          label="📊 Descargar Excel (Todos)",
          data=excel_bytes,
          file_name="historial_detallado_todos.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          type="primary",
          use_container_width=True
      )
    with col_dl2:
      st.download_button(
          label="📄 Descargar PDF (Todos)",
          data=pdf_bytes,
          file_name="historial_detallado_todos.pdf",
          mime="application/pdf",
          type="primary",
          use_container_width=True
      )

  else:
    if paciente_seleccionado:
      datos_p = resumen_pacientes[paciente_seleccionado]
      with st.container(border=True):
        st.markdown(f"### 👤 {paciente_seleccionado}")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("🟢 Días Completos", datos_p["completos"])
        col2.metric("🟡 Días Medios", datos_p["medios"])
        col3.metric("🚗 Total Días Viajados", datos_p["total_viajes"])
        col4.metric("💰 Total Acumulado", f"${datos_p['total_monto']:,}")

        st.markdown("#### 📅 Fechas y detalles de sus viajes:")
        viajes_paciente = [item for item in historial_detalle if item["nombre"] == paciente_seleccionado]
        if viajes_paciente:
          for v in viajes_paciente:
            st.write(f"• **{v['fecha']}** — {v['estado']} (*${v['valor']:,}*)")
        else:
          st.write("*(Sin registros de viajes)*")

      detalle_individual = [item for item in historial_detalle if item["nombre"] == paciente_seleccionado]
      excel_bytes = generar_excel_historial_detallado(detalle_individual)
      pdf_bytes = generar_pdf_historial_detallado(detalle_individual, f"Historial Detallado del Paciente: {paciente_seleccionado}")

      col_dl1, col_dl2 = st.columns(2)
      with col_dl1:
        st.download_button(
            label=f"📊 Descargar Excel ({paciente_seleccionado})",
            data=excel_bytes,
            file_name=f"historial_{paciente_seleccionado.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
      with col_dl2:
        st.download_button(
            label=f"📄 Descargar PDF ({paciente_seleccionado})",
            data=pdf_bytes,
            file_name=f"historial_{paciente_seleccionado.replace(' ', '_')}.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )

elif menu == "Planilla Matricial":
  st.header("📅 Planilla de Control Estilo Matriz")
  st.write("Selecciona el mes y año para generar y descargar la planilla en formato de cuadrícula.")

  col_m1, col_m2 = st.columns(2)
  
  meses_opciones = {
      "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4,
      "Mayo": 5, "Junio": 6, "Julio": 7, "Agosto": 8,
      "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12
  }
  
  with col_m1:
    mes_nombre_elegido = st.selectbox("Mes", list(meses_opciones.keys()), index=10)
  with col_m2:
    anio_elegido = st.selectbox("Año", [2026, 2027, 2028], index=0)

  mes_num_elegido = meses_opciones[mes_nombre_elegido]

  excel_planilla = generar_excel_planilla_matricial(datos, todos_pacientes, anio_elegido, mes_num_elegido)
  
  st.download_button(
      label=f"📥 Descargar Planilla Matricial ({mes_nombre_elegido} {anio_elegido}) en Excel (.xlsx)",
      data=excel_planilla,
      file_name=f"planilla_asistencia_{mes_nombre_elegido.lower()}_{anio_elegido}.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      type="primary",
      use_container_width=True
  )
