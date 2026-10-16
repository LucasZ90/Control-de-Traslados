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
  supabase_url = st.secrets["SUPABASE_URL"].strip().rstrip("/")
  supabase_key = st.secrets["SUPABASE_KEY"].strip()
  supabase = create_client(supabase_url, supabase_key)
except Exception as e:
  st.error(
      "⚠️ Error de configuración: Faltan las credenciales de Supabase en los"
      f" Secrets de Streamlit. Detalle: {e}"
  )
  st.stop()

# --- ESTILOS VISUALES COMPACTOS ---
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
        border-radius: 6px;
        font-weight: 600;
        padding: 0.3rem 0.1rem;
        font-size: 0.8rem;
        width: 100%;
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

# --- NAVEGACIÓN SUPERIOR REDUCIDA (2 BOTONES) ---
if "menu_actual" not in st.session_state:
  st.session_state["menu_actual"] = "Registrar / Actualizar Viajes"

col_nav1, col_nav2 = st.columns(2)
with col_nav1:
  if st.button("📋 Registrar", use_container_width=True):
    st.session_state["menu_actual"] = "Registrar / Actualizar Viajes"
with col_nav2:
  if st.button("📊 Resumen", use_container_width=True):
    st.session_state["menu_actual"] = "Resumen Mensual"

menu = st.session_state["menu_actual"]
st.markdown("---")

datos = cargar_registros_db()
todos_pacientes = cargar_pacientes_db()

if menu == "Registrar / Actualizar Viajes":
  st.subheader("📋 Registro Diario")

  # Botón superior para agregar nuevo paciente directamente
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
      nombre_dia = [
          "Lunes",
          "Martes",
          "Miércoles",
          "Jueves",
          "Viernes",
          "Sábado",
          "Domingo",
      ][dia_semana]
      grupo_sugerido = "MWF" if dia_semana in [0, 2, 4] else "TTS"
      st.info(
          f"📅 **{fecha_str}** ({nombre_dia}) — *Sugerencia: Grupo"
          f" {'1 (MWF)' if grupo_sugerido == 'MWF' else '2 (TTS)'}*"
      )

      if fecha_str not in datos:
        datos[fecha_str] = {}

      # --- CONFIGURACIÓN RÁPIDA DE PACIENTES PARA ESTE DÍA ---
      with st.expander(
          "⚙️ Seleccionar pacientes y turnos de hoy", expanded=False
      ):
        st.write(
            "Marca los pacientes que viajan hoy y ajusta su turno si es"
            " necesario:"
        )
        for p in todos_pacientes:
          p_nombre = p["nombre"]
          p_id = p.get("id")
          sugerido = (p_nombre in datos[fecha_str]) or (
              p.get("grupo") == grupo_sugerido
          )

          cols_sel = st.columns([3, 2])
          seleccionado = cols_sel[0].checkbox(
              f"{p_nombre} ({p.get('turno','05:40')})",
              value=sugerido,
              key=f"chk_sel_{fecha_str}_{p_id}",
          )

          turno_actual_dia = (
              datos[fecha_str][p_nombre].get(
                  "turno_dia", p.get("turno", "05:40")
              )
              if p_nombre in datos[fecha_str]
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

      # --- RENDERIZAR TURNOS CON TARJETAS COMPACTAS ---
      for turno_turno, titulo_turno in [
          ("05:40", "⏰ Turno 05:40 AM"),
          ("08:40", "⏰ Turno 08:40 AM"),
      ]:
        st.markdown(f"### {titulo_turno}")
        pacientes_turno = [
            p
            for p in todos_pacientes
            if p["nombre"] in pacientes_hoy_dict
            and pacientes_hoy_dict[p["nombre"]].get("turno_dia", "05:40")
            == turno_turno
        ]

        if not pacientes_turno:
          st.write("*(Sin pacientes programados en este turno)*")

        for p in pacientes_turno:
          nombre = p["nombre"]
          p_id = p["id"]
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

            # El nombre del paciente ahora funciona como un expandible desplegable para editar sus datos
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
                  index=0 if p["grupo"] == "MWF" else 1,
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
                      p["grupo"],
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
                      p["grupo"],
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

            # Información de mapas y contactos en línea compacta debajo del nombre
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
            st.markdown(
                f"<small>{'  •  '.join(partes_info)}</small>",
                unsafe_allow_html=True,
            )

            st.markdown("")

            # --- BOTONES IDA, VUELTA Y FIN HORIZONTALES ---
            col_b1, col_b2, col_b3 = st.columns([2.2, 2.2, 1.6])

            label_ida = "✅ Ida" if reg_ida else "❌ Ida"
            if col_b1.button(
                label_ida,
                key=f"t_{turno_turno}_i_{p_id}_{fecha_str}",
                use_container_width=True,
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
