from datetime import datetime, date
import io
import json
import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import streamlit as st

# Configuración de la página en español
st.set_page_config(
    page_title="Gestión de Traslados - Diálisis", page_icon="🚗", layout="centered"
)

# --- ESTILOS VISUALES Y PALETA DE COLORES PROFESIONAL ---
st.markdown("""
    <style>
    /* Estilo general limpio y moderno */
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
    /* Botones de navegación superior más estéticos */
    div.stButton > button {
        border-radius: 8px;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

ARCHIVO_DATOS = "registros_dialisis.json"
ARCHIVO_PACIENTES = "pacientes_dialisis.json"

# GRUPOS INICIALES POR DEFECTO
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

TARIFA_COMPLETO = 5000  # Ida y Vuelta
TARIFA_MEDIO = 2500  # Solo un tramo (Ida o Vuelta)
TURNOS_DISPONIBLES = ["05:40", "08:40"]


def cargar_pacientes():
  if os.path.exists(ARCHIVO_PACIENTES):
    with open(ARCHIVO_PACIENTES, "r", encoding="utf-8") as f:
      pacs = json.load(f)
      for grupo in ["MWF", "TTS"]:
        if grupo in pacs:
          for p in pacs[grupo]:
            if "maps_link" not in p:
              p["maps_link"] = ""
            if "contactos" not in p:
              tels_viejos = p.get("telefonos", "")
              if tels_viejos:
                p["contactos"] = [{"numero": tels_viejos, "relacion": "Principal"}]
              else:
                p["contactos"] = []
            if "telefonos" in p:
              del p["telefonos"]
      return pacs
  else:
    pacs = {"MWF": DEFAULT_MWF, "TTS": DEFAULT_TTS}
    guardar_pacientes(pacs)
    return pacs


def guardar_pacientes(pacs):
  with open(ARCHIVO_PACIENTES, "w", encoding="utf-8") as f:
    json.dump(pacs, f, ensure_ascii=False, indent=4)


def cargar_datos():
  if os.path.exists(ARCHIVO_DATOS):
    with open(ARCHIVO_DATOS, "r", encoding="utf-8") as f:
      try:
        datos = json.load(f)
      except json.JSONDecodeError:
        datos = {}
      
      datos_filtrados = {}
      for fecha_str, registros in datos.items():
        try:
          f_dt = datetime.strptime(fecha_str, "%d/%m/%Y").date()
          if f_dt >= date(2026, 10, 14):
            datos_filtrados[fecha_str] = registros
        except ValueError:
          pass
          
      if len(datos_filtrados) != len(datos):
        with open(ARCHIVO_DATOS, "w", encoding="utf-8") as f_out:
          json.dump(datos_filtrados, f_out, ensure_ascii=False, indent=4)
        
      return datos_filtrados
  return {}


def guardar_datos(datos):
  datos_filtrados = {}
  for fecha_str, registros in datos.items():
    try:
      f_dt = datetime.strptime(fecha_str, "%d/%m/%Y").date()
      if f_dt >= date(2026, 10, 14):
        datos_filtrados[fecha_str] = registros
    except ValueError:
      pass
  with open(ARCHIVO_DATOS, "w", encoding="utf-8") as f:
    json.dump(datos_filtrados, f, ensure_ascii=False, indent=4)


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


def generar_excel_bytes(datos):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Resumen Mensual"

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
      ws.cell(row=fila_actual, column=3, value=info["estado"]).border = thin_border

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


# Título serio sin emojis ni dibujos de autos
st.title("Gestión de Traslados")

# --- NAVEGACIÓN SUPERIOR COMPACTA (3 BOTONES UNO AL LADO DEL OTRO) ---
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

datos = cargar_datos()
pacientes_db = cargar_pacientes()

grupo_MWF = pacientes_db["MWF"]
grupo_TTS = pacientes_db["TTS"]

if menu == "Registrar / Actualizar Viajes":
  st.subheader("📋 Registro Diario")

  fecha_seleccionada = st.date_input("Selecciona la fecha", datetime.now())
  fecha_str = fecha_seleccionada.strftime("%d/%m/%Y")

  if fecha_seleccionada < date(2026, 10, 14):
    st.error("⚠️ Los registros de traslados comienzan a partir del 14 de octubre de 2026.")
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
            f"📅 **{fecha_str}** ({nombre_dia}) — *Grupo 1 (Lunes, Miércoles, Viernes)*"
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
            f"📅 **{fecha_str}** ({nombre_dia}) — *Grupo 2 (Martes, Jueves, Sábados)*"
        )

      if fecha_str not in datos:
        datos[fecha_str] = {}

      def esta_finalizado(p):
        info = datos.get(fecha_str, {}).get(p["nombre"], {})
        return (info.get("ida", False) and info.get("vuelta", False)) or info.get("listo", False)

      pendientes_totales = sum(1 for p in pacientes if not esta_finalizado(p))
      st.metric(label="👥 Pacientes Pendientes Hoy", value=pendientes_totales)

      st.markdown("### ⏰ Turno 05:40 AM")
      pacientes_t1 = [p for p in pacientes if p["turno"] == "05:40"]

      if not pacientes_t1:
        st.write("*(Sin pacientes programados)*")

      for p in pacientes_t1:
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
          estado_icono = "🟢" if (reg_ida and reg_vuelta) else ("🟡" if (reg_ida or reg_vuelta or is_listo) else "⚪")
          st.markdown(f"**{estado_icono} {nombre}** {' *[Finalizado]*' if is_listo else ''}")
          
          maps_url = p.get("maps_link", "").strip()
          contactos = [c for c in p.get("contactos", []) if c.get("numero")]

          partes_info = []
          if maps_url:
            partes_info.append(f"[📍 Ubicación]({maps_url})")
          else:
            partes_info.append("📍 *(Sin ubicación)*")

          if contactos:
            contactos_links = " | ".join([f"[📞 {c['relacion']}](tel:{c['numero'].replace(' ', '').replace('-', '')})" for c in contactos])
            partes_info.append(contactos_links)
          else:
            partes_info.append("📞 *(Sin contacto)*")

          st.markdown("  •  ".join(partes_info))

          col_a, col_b = st.columns(2)

          label_ida = "✅ Ida Realizada" if reg_ida else "❌ Ida Pendiente"
          if col_a.button(
              label_ida, key=f"t1_ida_{nombre}", use_container_width=True
          ):
            info_paciente["ida"] = not reg_ida
            actualizar_valores_paciente(info_paciente)
            guardar_datos(datos)
            st.rerun()

          label_vuelta = "✅ Vuelta Realizada" if reg_vuelta else "❌ Vuelta Pendiente"
          if col_b.button(
              label_vuelta, key=f"t1_vuelta_{nombre}", use_container_width=True
          ):
            info_paciente["vuelta"] = not reg_vuelta
            actualizar_valores_paciente(info_paciente)
            guardar_datos(datos)
            st.rerun()

          if not is_listo:
            if st.button("🏁 Finalizar Registro", key=f"t1_fin_{nombre}", use_container_width=True):
              info_paciente["listo"] = True
              guardar_datos(datos)
              st.rerun()
          else:
            if st.button("🔄 Reabrir Registro", key=f"t1_reab_{nombre}", use_container_width=True):
              info_paciente["listo"] = False
              guardar_datos(datos)
              st.rerun()

      st.markdown("### ⏰ Turno 08:40 AM")
      pacientes_t2 = [p for p in pacientes if p["turno"] == "08:40"]

      if not pacientes_t2:
        st.write("*(Sin pacientes programados)*")

      for p in pacientes_t2:
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
          estado_icono = "🟢" if (reg_ida and reg_vuelta) else ("🟡" if (reg_ida or reg_vuelta or is_listo) else "⚪")
          st.markdown(f"**{estado_icono} {nombre}** {' *[Finalizado]*' if is_listo else ''}")
          
          maps_url = p.get("maps_link", "").strip()
          contactos = [c for c in p.get("contactos", []) if c.get("numero")]

          partes_info = []
          if maps_url:
            partes_info.append(f"[📍 Ubicación]({maps_url})")
          else:
            partes_info.append("📍 *(Sin ubicación)*")

          if contactos:
            contactos_links = " | ".join([f"[📞 {c['relacion']}](tel:{c['numero'].replace(' ', '').replace('-', '')})" for c in contactos])
            partes_info.append(contactos_links)
          else:
            partes_info.append("📞 *(Sin contacto)*")

          st.markdown("  •  ".join(partes_info))

          col_a, col_b = st.columns(2)

          label_ida = "✅ Ida Realizada" if reg_ida else "❌ Ida Pendiente"
          if col_a.button(
              label_ida, key=f"t2_ida_{nombre}", use_container_width=True
          ):
            info_paciente["ida"] = not reg_ida
            actualizar_valores_paciente(info_paciente)
            guardar_datos(datos)
            st.rerun()

          label_vuelta = "✅ Vuelta Realizada" if reg_vuelta else "❌ Vuelta Pendiente"
          if col_b.button(
              label_vuelta, key=f"t2_vuelta_{nombre}", use_container_width=True
          ):
            info_paciente["vuelta"] = not reg_vuelta
            actualizar_valores_paciente(info_paciente)
            guardar_datos(datos)
            st.rerun()

          if not is_listo:
            if st.button("🏁 Finalizar Registro", key=f"t2_fin_{nombre}", use_container_width=True):
              info_paciente["listo"] = True
              guardar_datos(datos)
              st.rerun()
          else:
            if st.button("🔄 Reabrir Registro", key=f"t2_reab_{nombre}", use_container_width=True):
              info_paciente["listo"] = False
              guardar_datos(datos)
              st.rerun()

      if datos:
        st.markdown("---")
        excel_bytes = generar_excel_bytes(datos)
        st.download_button(
            label="📊 Descargar Resumen en Excel (.xlsx)",
            data=excel_bytes,
            file_name="control_dialisis_resumen.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
    for p in list(grupo_MWF):
      nombre_key = p['nombre'].replace(" ", "_").replace(".", "_")
      with st.container(border=True):
        st.markdown(f"**{p['nombre']}** — *Turno: {p['turno']}*")
        
        with st.expander("✏️ Editar"):
          edit_nombre = st.text_input("Nombre y Apellido", value=p["nombre"], key=f"edit_nom_mwf_{nombre_key}")
          edit_grupo = st.selectbox(
              "Grupo Asignado",
              ["Grupo 1 (Lunes, Miércoles, Viernes)", "Grupo 2 (Martes, Jueves, Sábados)"],
              index=0,
              key=f"edit_grupo_mwf_{nombre_key}"
          )
          edit_turno = st.selectbox("Turno Habitual", TURNOS_DISPONIBLES, index=TURNOS_DISPONIBLES.index(p["turno"]), key=f"edit_turno_mwf_{nombre_key}")
          edit_maps = st.text_input("Enlace de Google Maps", value=p.get("maps_link", ""), key=f"edit_maps_mwf_{nombre_key}")

          st.markdown("---")
          st.markdown("**📞 Teléfonos y Contactos:**")
          
          contactos = p.get("contactos", [])
          for idx_c, contacto in enumerate(contactos):
            cols_c = st.columns([3, 3, 1])
            cols_c[0].write(f"📞 {contacto['numero']}")
            cols_c[1].write(f"({contacto['relacion']})")
            if cols_c[2].button("🗑️", key=f"del_c_mwf_{nombre_key}_{idx_c}", use_container_width=True):
              contactos.pop(idx_c)
              p["contactos"] = contactos
              guardar_pacientes(pacientes_db)
              st.rerun()

          st.markdown("➕ *Agregar otro contacto:*")
          nuevo_num = st.text_input("Número de teléfono", key=f"new_num_mwf_{nombre_key}")
          nueva_rel = st.text_input("Relación (Ej: Paciente, Hijo/a, Cuidador)", key=f"new_rel_mwf_{nombre_key}")
          
          if st.button("➕ Añadir este Contacto", key=f"add_c_btn_mwf_{nombre_key}", use_container_width=True):
            if nuevo_num.strip():
              contactos.append({"numero": nuevo_num.strip(), "relacion": nueva_rel.strip() if nueva_rel.strip() else "Contacto"})
              p["contactos"] = contactos
              guardar_pacientes(pacientes_db)
              st.success("¡Contacto añadido!")
              st.rerun()
            else:
              st.error("Por favor ingresa un número de teléfono.")

          st.markdown("---")
          col_save, col_del = st.columns([5, 1])
          if col_save.button("💾 Guardar Cambios", key=f"save_mwf_{nombre_key}", type="primary", use_container_width=True):
            p["nombre"] = edit_nombre.strip()
            p["turno"] = edit_turno
            p["maps_link"] = edit_maps.strip()
            
            nuevo_grupo_val = "MWF" if "Grupo 1" in edit_grupo else "TTS"
            if nuevo_grupo_val != "MWF":
              grupo_MWF.remove(p)
              grupo_TTS.append(p)
            
            pacientes_db["MWF"] = grupo_MWF
            pacientes_db["TTS"] = grupo_TTS
            guardar_pacientes(pacientes_db)
            st.success("¡Cambios guardados con éxito!")
            st.rerun()

          if col_del.button("🗑️", key=f"del_mwf_{nombre_key}", use_container_width=True, help="Eliminar paciente"):
            grupo_MWF.remove(p)
            pacientes_db["MWF"] = grupo_MWF
            guardar_pacientes(pacientes_db)
            st.rerun()

    st.markdown("---")
    st.subheader("Grupo 2 (Martes, Jueves, Sábados)")
    for p in list(grupo_TTS):
      nombre_key = p['nombre'].replace(" ", "_").replace(".", "_")
      with st.container(border=True):
        st.markdown(f"**{p['nombre']}** — *Turno: {p['turno']}*")
        
        with st.expander("✏️ Editar"):
          edit_nombre = st.text_input("Nombre y Apellido", value=p["nombre"], key=f"edit_nom_tts_{nombre_key}")
          edit_grupo = st.selectbox(
              "Grupo Asignado",
              ["Grupo 1 (Lunes, Miércoles, Viernes)", "Grupo 2 (Martes, Jueves, Sábados)"],
              index=1,
              key=f"edit_grupo_tts_{nombre_key}"
          )
          edit_turno = st.selectbox("Turno Habitual", TURNOS_DISPONIBLES, index=TURNOS_DISPONIBLES.index(p["turno"]), key=f"edit_turno_tts_{nombre_key}")
          edit_maps = st.text_input("Enlace de Google Maps", value=p.get("maps_link", ""), key=f"edit_maps_tts_{nombre_key}")

          st.markdown("---")
          st.markdown("**📞 Teléfonos y Contactos:**")
          
          contactos = p.get("contactos", [])
          for idx_c, contacto in enumerate(contactos):
            cols_c = st.columns([3, 3, 1])
            cols_c[0].write(f"📞 {contacto['numero']}")
            cols_c[1].write(f"({contacto['relacion']})")
            if cols_c[2].button("🗑️", key=f"del_c_tts_{nombre_key}_{idx_c}", use_container_width=True):
              contactos.pop(idx_c)
              p["contactos"] = contactos
              guardar_pacientes(pacientes_db)
              st.rerun()

          st.markdown("➕ *Agregar otro contacto:*")
          nuevo_num = st.text_input("Número de teléfono", key=f"new_num_tts_{nombre_key}")
          nueva_rel = st.text_input("Relación (Ej: Paciente, Hijo/a, Cuidador)", key=f"new_rel_tts_{nombre_key}")
          
          if st.button("➕ Añadir este Contacto", key=f"add_c_btn_tts_{nombre_key}", use_container_width=True):
            if nuevo_num.strip():
              contactos.append({"numero": nuevo_num.strip(), "relacion": nueva_rel.strip() if nueva_rel.strip() else "Contacto"})
              p["contactos"] = contactos
              guardar_pacientes(pacientes_db)
              st.success("¡Contacto añadido!")
              st.rerun()
            else:
              st.error("Por favor ingresa un número de teléfono.")

          st.markdown("---")
          col_save, col_del = st.columns([5, 1])
          if col_save.button("💾 Guardar Cambios", key=f"save_tts_{nombre_key}", type="primary", use_container_width=True):
            p["nombre"] = edit_nombre.strip()
            p["turno"] = edit_turno
            p["maps_link"] = edit_maps.strip()
            
            nuevo_grupo_val = "MWF" if "Grupo 1" in edit_grupo else "TTS"
            if nuevo_grupo_val != "TTS":
              grupo_TTS.remove(p)
              grupo_MWF.append(p)
            
            pacientes_db["MWF"] = grupo_MWF
            pacientes_db["TTS"] = grupo_TTS
            guardar_pacientes(pacientes_db)
            st.success("¡Cambios guardados con éxito!")
            st.rerun()

          if col_del.button("🗑️", key=f"del_tts_{nombre_key}", use_container_width=True, help="Eliminar paciente"):
            grupo_TTS.remove(p)
            pacientes_db["TTS"] = grupo_TTS
            guardar_pacientes(pacientes_db)
            st.rerun()

  with pestana2:
    st.subheader("Agregar un nuevo paciente")
    nuevo_nombre = st.text_input("Nombre y Apellido del Paciente")
    nuevo_grupo = st.selectbox(
        "Selecciona el Grupo",
        [
            "Grupo 1 (Lunes, Miércoles, Viernes)",
            "Grupo 2 (Martes, Jueves, Sábados)",
        ],
    )
    nuevo_turno = st.selectbox(
        "Horario de Turno Habitual", TURNOS_DISPONIBLES
    )
    nuevo_maps = st.text_input("Enlace de Google Maps (Opcional)")
    nuevo_tel = st.text_input("Número de Teléfono Principal (Opcional)")
    nueva_rel_tel = st.text_input("Relación de este teléfono (Ej: Paciente, Hijo/a)", value="Paciente")

    st.markdown("")
    if st.button(
        "➕ Guardar Nuevo Paciente", type="primary", use_container_width=True
    ):
      if nuevo_nombre.strip() == "":
        st.error("Por favor, escribe el nombre del paciente.")
      else:
        lista_contactos = []
        if nuevo_tel.strip():
          lista_contactos.append({
              "numero": nuevo_tel.strip(),
              "relacion": nueva_rel_tel.strip() if nueva_rel_tel.strip() else "Paciente"
          })

        nuevo_registro = {
            "nombre": nuevo_nombre.strip(),
            "turno": nuevo_turno,
            "maps_link": nuevo_maps.strip(),
            "contactos": lista_contactos
        }
        if "Grupo 1" in nuevo_grupo:
          grupo_MWF.append(nuevo_registro)
          pacientes_db["MWF"] = grupo_MWF
        else:
          grupo_TTS.append(nuevo_registro)
          pacientes_db["TTS"] = grupo_TTS

        guardar_pacientes(pacientes_db)
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
    st.info("No hay registros de traslados (comenzando estrictamente desde el 14/10/2026).")
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
    excel_bytes = generar_excel_bytes(datos_validos)
    st.download_button(
        label="📊 Descargar Resumen en Excel (.xlsx)",
        data=excel_bytes,
        file_name="control_dialisis_resumen.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )