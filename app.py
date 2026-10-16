from datetime import datetime, date
import io
import json
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import streamlit as st
from supabase import create_client

# Configuración de la página en español
st.set_page_config(
    page_title="Gestión de Traslados - Diálisis", page_icon="🚗", layout="centered"
)

# --- CONEXIÓN A SUPABASE ---
try:
  supabase_url = st.secrets["SUPABASE_URL"]
  supabase_key = st.secrets["SUPABASE_KEY"]
  supabase = create_client(supabase_url, supabase_key)
except Exception as e:
  st.error(
      "⚠️ Error de configuración: Faltan las credenciales de Supabase en los"
      " Secrets de Streamlit."
  )
  st.stop()

# --- ESTILOS VISUALES ---
st.markdown("""
    <style>
    .stApp {
        background-color: #0E1117;
        color: #FAFAFA;
    }
    h1 {
        color: #FFFFFF;
        font-family: 'Inter', sans-serif;
        font-weight: 700;
        font-size: 1.8rem;
        margin-bottom: 0.5rem;
    }
    h2, h3 {
        color: #E2E8F0;
        font-family: 'Inter', sans-serif;
    }
    div.stButton > button {
        border-radius: 8px;
        font-weight: 600;
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

    mwf = [p for p in data if p["grupo"] == "MWF"]
    tts = [p for p in data if p["grupo"] == "TTS"]
    return {"MWF": mwf, "TTS": tts}
  except Exception:
    return {"MWF": DEFAULT_MWF, "TTS": DEFAULT_TTS}


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
  except Exception as e:
    st.error(f"Error al guardar paciente en la base de datos: {e}")


def eliminar_paciente_db(p_id):
  try:
    if p_id:
      supabase.table("pacientes").delete().eq("id", p_id).execute()
  except Exception as e:
    st.error(f"Error al eliminar paciente: {e}")


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
    if existing.data:
      row_id = existing.data[0]["id"]
      supabase.table("registros").update({
          "ida": info["ida"],
          "vuelta": info["vuelta"],
          "listo": info["listo"],
          "estado": info["estado"],
          "valor": info["valor"],
      }).eq("id", row_id).execute()
    else:
      supabase.table("registros").insert({
          "fecha": fecha_str,
          "nombre": nombre,
          "ida": info["ida"],
          "vuelta": info["vuelta"],
          "listo": info["listo"],
          "estado": info["estado"],
          "valor": info["valor"],
      }).execute()
  except Exception as e:
    st.error(f"Error al guardar registro: {e}")


def actualizar_valores_paciente(info):
  reg_ida = info.get("ida", False)
  reg_vuelta = info.get("vuelta", False)
  if reg_ida and reg_vuelta:
    info["estado"] = "Completo (Ida y Vuelta)"
    info["valor"] = TARIFA_COMPLETO
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
  for fecha, pacientes_dia in sorted(datos.items()):
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


# Título principal limpio
st.title("Gestión de Traslados")

# --- NAVEGACIÓN SUPERIOR COMPACTA (3 BOTONES) ---
if "menu_actual" not in st.session_state:
  st.session_state["menu_actual"] = "Registrar / Actualizar Viajes"

col_nav1, col_nav2, col_nav3 = st.columns(3)
with col_nav1:
  if st.button("📋 Registrar", use_container_width=True):
    st.session_state["menu_actual"] = "Registrar / Actualizar Viajes"
with col_nav2:
  if st.button("👥 Pacientes", use_container_width=True):
    st.session_state["menu_actual"] = "Gestionar Pacientes"
with col_nav3:
  if st.button("📊 Resumen", use_container_width=True):
    st.session_state["menu_actual"] = "Resumen Mensual"

menu = st.session_state["menu_actual"]
st.markdown("---")

datos = cargar_registros_db()
pacientes_db = cargar_pacientes_db()

grupo_MWF = pacientes_db["MWF"]
grupo_TTS = pacientes_db["TTS"]

if menu == "Registrar / Actualizar Viajes":
  st.subheader("📋 Registro Diario")

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
      if dia_semana in [0, 2, 4]:
        pacientes = grupo_MWF
        nombre_dia = [
            "Lunes",
            "Martes",
            "Miércoles",
            "Jueves",
            "Viernes",
            "Sábado",
            "Domingo",
        ][dia_semana]
        st.info(
            f"📅 **{fecha_str}** ({nombre_dia}) — *Grupo 1 (Lunes, Miércoles,"
            " Viernes)*"
        )
      else:
        pacientes = grupo_TTS
        nombre_dia = [
            "Lunes",
            "Martes",
            "Miércoles",
            "Jueves",
            "Viernes",
            "Sábado",
            "Domingo",
        ][dia_semana]
        st.info(
            f"📅 **{fecha_str}** ({nombre_dia}) — *Grupo 2 (Martes, Jueves,"
            " Sábados)*"
        )

      if fecha_str not in datos:
        datos[fecha_str] = {}

      def esta_finalizado(p):
        info = datos.get(fecha_str, {}).get(p["nombre"], {})
        return (info.get("ida", False) and info.get("vuelta", False)) or info.get(
            "listo", False
        )

      pendientes_totales = sum(1 for p in pacientes if not esta_finalizado(p))
      st.metric(label="👥 Pacientes Pendientes Hoy", value=pendientes_totales)

      st.markdown("### ⏰ Turno 05:40 AM")
      pacientes_t1 = [p for p in pacientes if p["turno"] == "05:40"]

      if not pacientes_t1:
        st.write("*(Sin pacientes programados)*")

      for idx_p, p in enumerate(pacientes_t1):
        nombre = p["nombre"]

        if nombre not in datos[fecha_str]:
          datos[fecha_str][nombre] = {
              "ida": False,
              "vuelta": False,
              "listo": False,
              "estado": "Ausente (No viajó)",
              "valor": 0,
          }

        info_paciente = datos[fecha_str][nombre]
        reg_ida = info_paciente.get("ida", False)
        reg_vuelta = info_paciente.get("vuelta", False)
        is_listo = info_paciente.get("listo", False)

        with st.container(border=True):
          estado_icono = (
              "🟢"
              if (reg_ida and reg_vuelta)
              else ("🟡" if (reg_ida or reg_vuelta or is_listo) else "⚪")
          )
          st.markdown(
              f"**{estado_icono} {nombre}**"
              f" {' *[Finalizado]*' if is_listo else ''}"
          )

          maps_url = p.get("maps_link", "").strip()
          contactos = [c for c in p.get("contactos", []) if c.get("numero")]

          partes_info = []
          if maps_url:
            partes_info.append(f"[📍 Ubicación]({maps_url})")
          else:
            partes_info.append("📍 *(Sin ubicación)*")

          if contactos:
            contactos_links = " | ".join([
                f"[📞 {c['relacion']}](tel:{c['numero'].replace(' ', '').replace('-', '')})"
                for c in contactos
            ])
            partes_info.append(contactos_links)
          else:
            partes_info.append("📞 *(Sin contacto)*")

          st.markdown("  •  ".join(partes_info))

          col_a, col_b = st.columns(2)

          label_ida = "✅ Ida Realizada" if reg_ida else "❌ Ida Pendiente"
          if col_a.button(
              label_ida, key=f"t1_i_{idx_p}_{nombre}", use_container_width=True
          ):
            info_paciente["ida"] = not reg_ida
            actualizar_valores_paciente(info_paciente)
            guardar_registro_db(fecha_str, nombre, info_paciente)
            st.rerun()

          label_vuelta = (
              "✅ Vuelta Realizada" if reg_vuelta else "❌ Vuelta Pendiente"
          )
          if col_b.button(
              label_vuelta,
              key=f"t1_v_{idx_p}_{nombre}",
              use_container_width=True,
          ):
            info_paciente["vuelta"] = not reg_vuelta
            actualizar_valores_paciente(info_paciente)
            guardar_registro_db(fecha_str, nombre, info_paciente)
            st.rerun()

          if not is_listo:
            if st.button(
                "🏁 Finalizar Registro",
                key=f"t1_f_{idx_p}_{nombre}",
                use_container_width=True,
            ):
              info_paciente["listo"] = True
              guardar_registro_db(fecha_str, nombre, info_paciente)
              st.rerun()
          else:
            if st.button(
                "🔄 Reabrir Registro",
                key=f"t1_r_{idx_p}_{nombre}",
                use_container_width=True,
            ):
              info_paciente["listo"] = False
              guardar_registro_db(fecha_str, nombre, info_paciente)
              st.rerun()

      st.markdown("### ⏰ Turno 08:40 AM")
      pacientes_t2 = [p for p in pacientes if p["turno"] == "08:40"]

      if not pacientes_t2:
        st.write("*(Sin pacientes programados)*")

      for idx_p, p in enumerate(pacientes_t2):
        nombre = p["nombre"]

        if nombre not in datos[fecha_str]:
          datos[fecha_str][nombre] = {
              "ida": False,
              "vuelta": False,
              "listo": False,
              "estado": "Ausente (No viajó)",
              "valor": 0,
          }

        info_paciente = datos[fecha_str][nombre]
        reg_ida = info_paciente.get("ida", False)
        reg_vuelta = info_paciente.get("vuelta", False)
        is_listo = info_paciente.get("listo", False)

        with st.container(border=True):
          estado_icono = (
              "🟢"
              if (reg_ida and reg_vuelta)
              else ("🟡" if (reg_ida or reg_vuelta or is_listo) else "⚪")
          )
          st.markdown(
              f"**{estado_icono} {nombre}**"
              f" {' *[Finalizado]*' if is_listo else ''}"
          )

          maps_url = p.get("maps_link", "").strip()
          contactos = [c for c in p.get("contactos", []) if c.get("numero")]

          partes_info = []
          if maps_url:
            partes_info.append(f"[📍 Ubicación]({maps_url})")
          else:
            partes_info.append("📍 *(Sin ubicación)*")

          if contactos:
            contactos_links = " | ".join([
                f"[📞 {c['relacion']}](tel:{c['numero'].replace(' ', '').replace('-', '')})"
                for c in contactos
            ])
            partes_info.append(contactos_links)
          else:
            partes_info.append("📞 *(Sin contacto)*")

          st.markdown("  •  ".join(partes_info))

          col_a, col_b = st.columns(2)

          label_ida = "✅ Ida Realizada" if reg_ida else "❌ Ida Pendiente"
          if col_a.button(
              label_ida, key=f"t2_i_{idx_p}_{nombre}", use_container_width=True
          ):
            info_paciente["ida"] = not reg_ida
            actualizar_valores_paciente(info_paciente)
            guardar_registro_db(fecha_str, nombre, info_paciente)
            st.rerun()

          label_vuelta = (
              "✅ Vuelta Realizada" if reg_vuelta else "❌ Vuelta Pendiente"
          )
          if col_b.button(
              label_vuelta,
              key=f"t2_v_{idx_p}_{nombre}",
              use_container_width=True,
          ):
            info_paciente["vuelta"] = not reg_vuelta
            actualizar_valores_paciente(info_paciente)
            guardar_registro_db(fecha_str, nombre, info_paciente)
            st.rerun()

          if not is_listo:
            if st.button(
                "🏁 Finalizar Registro",
                key=f"t2_f_{idx_p}_{nombre}",
                use_container_width=True,
            ):
              info_paciente["listo"] = True
              guardar_registro_db(fecha_str, nombre, info_paciente)
              st.rerun()
          else:
            if st.button(
                "🔄 Reabrir Registro",
                key=f"t2_r_{idx_p}_{nombre}",
                use_container_width=True,
            ):
              info_paciente["listo"] = False
              guardar_registro_db(fecha_str, nombre, info_paciente)
              st.rerun()

      if fecha_str in datos:
        st.markdown("---")
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

elif menu == "Gestionar Pacientes":
  st.header("👥 Administrar Pacientes y Turnos")

  pestana1, pestana2 = st.tabs(
      ["Modificar / Ver Listado", "Agregar Nuevo Paciente"]
  )

  with pestana1:
    st.subheader("Grupo 1 (Lunes, Miércoles, Viernes)")
    for idx_p, p in enumerate(list(grupo_MWF)):
      with st.container(border=True):
        st.markdown(f"**{p['nombre']}** — *Turno: {p['turno']}*")

        with st.expander("✏️ Editar"):
          edit_nombre = st.text_input(
              "Nombre y Apellido", value=p["nombre"], key=f"mwf_nom_{idx_p}"
          )
          edit_grupo = st.selectbox(
              "Grupo Asignado",
              [
                  "Grupo 1 (Lunes, Miércoles, Viernes)",
                  "Grupo 2 (Martes, Jueves, Sábados)",
              ],
              index=0,
              key=f"mwf_grp_{idx_p}",
          )
          edit_turno = st.selectbox(
              "Turno Habitual",
              TURNOS_DISPONIBLES,
              index=TURNOS_DISPONIBLES.index(p["turno"]),
              key=f"mwf_trn_{idx_p}",
          )
          edit_maps = st.text_input(
              "Enlace de Google Maps",
              value=p.get("maps_link", ""),
              key=f"mwf_map_{idx_p}",
          )

          st.markdown("---")
          st.markdown("**📞 Teléfonos y Contactos:**")

          contactos = p.get("contactos", [])
          for idx_c, contacto in enumerate(contactos):
            cols_c = st.columns([3, 3, 1])
            cols_c[0].write(f"📞 {contacto['numero']}")
            cols_c[1].write(f"({contacto['relacion']})")
            if cols_c[2].button(
                "🗑️", key=f"mwf_del_c_{idx_p}_{idx_c}", use_container_width=True
            ):
              contactos.pop(idx_c)
              guardar_paciente_db(
                  p.get("id"),
                  "MWF",
                  p["nombre"],
                  p["turno"],
                  p["maps_link"],
                  contactos,
              )
              st.rerun()

          nuevo_num = st.text_input(
              "Número de teléfono", key=f"mwf_new_num_{idx_p}"
          )
          nueva_rel = st.text_input(
              "Relación (Ej: Paciente, Hijo/a, Cuidador)",
              key=f"mwf_new_rel_{idx_p}",
          )

          if st.button(
              "➕ Añadir este Contacto",
              key=f"mwf_add_btn_{idx_p}",
              use_container_width=True,
          ):
            if nuevo_num.strip():
              contactos.append({
                  "numero": nuevo_num.strip(),
                  "relacion": (
                      nueva_rel.strip() if nueva_rel.strip() else "Contacto"
                  ),
              })
              guardar_paciente_db(
                  p.get("id"),
                  "MWF",
                  p["nombre"],
                  p["turno"],
                  p["maps_link"],
                  contactos,
              )
              st.success("¡Contacto añadido!")
              st.rerun()
            else:
              st.error("Por favor ingresa un número de teléfono.")

          st.markdown("---")
          col_save, col_del = st.columns([5, 1])
          if col_save.button(
              "💾 Guardar Cambios",
              key=f"mwf_save_{idx_p}",
              type="primary",
              use_container_width=True,
          ):
            nuevo_grupo_val = "MWF" if "Grupo 1" in edit_grupo else "TTS"
            guardar_paciente_db(
                p.get("id"),
                nuevo_grupo_val,
                edit_nombre.strip(),
                edit_turno,
                edit_maps.strip(),
                contactos,
            )
            st.success("¡Cambios guardados con éxito!")
            st.rerun()

          if col_del.button(
              "🗑️",
              key=f"mwf_del_p_{idx_p}",
              use_container_width=True,
              help="Eliminar paciente",
          ):
            eliminar_paciente_db(p.get("id"))
            st.rerun()

    st.markdown("---")
    st.subheader("Grupo 2 (Martes, Jueves, Sábados)")
    for idx_p, p in enumerate(list(grupo_TTS)):
      with st.container(border=True):
        st.markdown(f"**{p['nombre']}** — *Turno: {p['turno']}*")

        with st.expander("✏️ Editar"):
          edit_nombre = st.text_input(
              "Nombre y Apellido", value=p["nombre"], key=f"tts_nom_{idx_p}"
          )
          edit_grupo = st.selectbox(
              "Grupo Asignado",
              [
                  "Grupo 1 (Lunes, Miércoles, Viernes)",
                  "Grupo 2 (Martes, Jueves, Sábados)",
              ],
              index=1,
              key=f"tts_grp_{idx_p}",
          )
          edit_turno = st.selectbox(
              "Turno Habitual",
              TURNOS_DISPONIBLES,
              index=TURNOS_DISPONIBLES.index(p["turno"]),
              key=f"tts_trn_{idx_p}",
          )
          edit_maps = st.text_input(
              "Enlace de Google Maps",
              value=p.get("maps_link", ""),
              key=f"tts_map_{idx_p}",
          )

          st.markdown("---")
          st.markdown("**📞 Teléfonos y Contactos:**")

          contactos = p.get("contactos", [])
          for idx_c, contacto in enumerate(contactos):
            cols_c = st.columns([3, 3, 1])
            cols_c[0].write(f"📞 {contacto['numero']}")
            cols_c[1].write(f"({contacto['relacion']})")
            if cols_c[2].button(
                "🗑️", key=f"tts_del_c_{idx_p}_{idx_c}", use_container_width=True
            ):
              contactos.pop(idx_c)
              guardar_paciente_db(
                  p.get("id"),
                  "TTS",
                  p["nombre"],
                  p["turno"],
                  p["maps_link"],
                  contactos,
              )
              st.rerun()

          nuevo_num = st.text_input(
              "Número de teléfono", key=f"tts_new_num_{idx_p}"
          )
          nueva_rel = st.text_input(
              "Relación (Ej: Paciente, Hijo/a, Cuidador)",
              key=f"tts_new_rel_{idx_p}",
          )

          if st.button(
              "➕ Añadir este Contacto",
              key=f"tts_add_btn_{idx_p}",
              use_container_width=True,
          ):
            if nuevo_num.strip():
              contactos.append({
                  "numero": nuevo_num.strip(),
                  "relacion": (
                      nueva_rel.strip() if nueva_rel.strip() else "Contacto"
                  ),
              })
              guardar_paciente_db(
                  p.get("id"),
                  "TTS",
                  p["nombre"],
                  p["turno"],
                  p["maps_link"],
                  contactos,
              )
              st.success("¡Contacto añadido!")
              st.rerun()
            else:
              st.error("Por favor ingresa un número de teléfono.")

          st.markdown("---")
          col_save, col_del = st.columns([5, 1])
          if col_save.button(
              "💾 Guardar Cambios",
              key=f"tts_save_{idx_p}",
              type="primary",
              use_container_width=True,
          ):
            nuevo_grupo_val = "MWF" if "Grupo 1" in edit_grupo else "TTS"
            guardar_paciente_db(
                p.get("id"),
                nuevo_grupo_val,
                edit_nombre.strip(),
                edit_turno,
                edit_maps.strip(),
                contactos,
            )
            st.success("¡Cambios guardados con éxito!")
            st.rerun()

          if col_del.button(
              "🗑️",
              key=f"tts_del_p_{idx_p}",
              use_container_width=True,
              help="Eliminar paciente",
          ):
            eliminar_paciente_db(p.get("id"))
            st.rerun()

  with pestana2:
    st.subheader("Agregar un nuevo paciente")
    nuevo_nombre = st.text_input(
        "Nombre y Apellido del Paciente", key="nuevo_nombre_paciente"
    )
    nuevo_grupo = st.selectbox(
        "Selecciona el Grupo",
        [
            "Grupo 1 (Lunes, Miércoles, Viernes)",
            "Grupo 2 (Martes, Jueves, Sábados)",
        ],
        key="nuevo_grupo_paciente",
    )
    nuevo_turno = st.selectbox(
        "Horario de Turno Habitual",
        TURNOS_DISPONIBLES,
        key="nuevo_turno_paciente",
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

    st.markdown("")
    if st.button(
        "➕ Guardar Nuevo Paciente",
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
                  nueva_rel_tel.strip() if nueva_rel_tel.strip() else "Paciente"
              ),
          })

        nuevo_grupo_val = "MWF" if "Grupo 1" in nuevo_grupo else "TTS"
        guardar_paciente_db(
            None,
            nuevo_grupo_val,
            nuevo_nombre.strip(),
            nuevo_turno,
            nuevo_maps.strip(),
            lista_contactos,
        )
        st.success(f"¡Paciente '{nuevo_nombre}' agregado con éxito!")
        st.rerun()

elif menu == "Resumen Mensual":
  st.header("📊 Resumen Mensual y Total a Cobrar")

  datos_validos = {}
  for f_str, reg in datos.items():
    try:
      if datetime.strptime(f_str, "%d/%m/%Y").date() >= date(2026, 10, 14):
        datos_validos[f_str] = reg
    except ValueError:
      pass

  if not datos_validos:
    st.info(
        "No hay registros de traslados (comenzando estrictamente desde el"
        " 14/10/2026)."
    )
  else:
    total_mes = 0
    for fecha, pacientes_dia in sorted(datos_validos.items()):
      with st.expander(f"📅 Fecha: {fecha}"):
        total_dia = 0
        for nombre, info in pacientes_dia.items():
          st.write(f"- **{nombre}**: {info['estado']} (${info['valor']:,})")
          total_dia += info["valor"]
        st.markdown(f"**Total del día: ${total_dia:,}**")
        total_mes += total_dia

    st.markdown("---")
    st.subheader(f"💰 TOTAL GENERAL ACUMULADO: ${total_mes:,}")

    st.markdown("")
    excel_bytes_mes = generar_excel_bytes(datos_validos, es_diario=False)
    st.download_button(
        label="📊 Descargar Resumen Mensual en Excel (.xlsx)",
        data=excel_bytes_mes,
        file_name="control_dialisis_resumen_mensual.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True,
    )
