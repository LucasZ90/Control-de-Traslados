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

      # --- FORMULARIO DE CARGA MASIVA CON UN SOLO BOTÓN DE GUARDADO ---
      with st.form(key=f"form_dia_{fecha_str}"):
        st.markdown(
            "### ⚙️ Configuración del Día (Pacientes, Turnos y Viajes)"
        )
        st.write(
            "Selecciona quién viaja, ajusta su turno, marca las idas/vueltas y"
            " presiona Guardar al final."
        )

        form_estados = {}

        for p in todos_pacientes:
          p_nombre = p["nombre"]
          p_id = p.get("id")
          info_existente = datos[fecha_str].get(p_nombre, {})
          ya_registrado = p_nombre in datos[fecha_str]
          sugerido = ya_registrado or (p.get("grupo") == grupo_sugerido)

          with st.container(border=True):
            st.markdown(f"**{p_nombre}** *(Habitual: {p.get('turno','05:40')})*")

            cols_f = st.columns([2, 2, 2, 2])
            chk_viaja = cols_f[0].checkbox(
                "Viaja hoy", value=sugerido, key=f"f_viaja_{fecha_str}_{p_id}"
            )

            turno_def = info_existente.get(
                "turno_dia", p.get("turno", "05:40")
            )
            idx_t = (
                TURNOS_DISPONIBLES.index(turno_def)
                if turno_def in TURNOS_DISPONIBLES
                else 0
            )
            sel_turno = cols_f[1].selectbox(
                "Turno",
                TURNOS_DISPONIBLES,
                index=idx_t,
                key=f"f_turno_{fecha_str}_{p_id}",
            )

            chk_ida = cols_f[2].checkbox(
                "Ida",
                value=info_existente.get("ida", False),
                key=f"f_ida_{fecha_str}_{p_id}",
            )
            chk_vuelta = cols_f[3].checkbox(
                "Vuelta",
                value=info_existente.get("vuelta", False),
                key=f"f_vuelta_{fecha_str}_{p_id}",
            )

            form_estados[p_nombre] = {
                "viaja": chk_viaja,
                "turno_dia": sel_turno,
                "ida": chk_ida,
                "vuelta": chk_vuelta,
            }

          # Expandible opcional de edición de datos fijos del paciente
          with st.expander(f"✏️ Editar datos fijos de {p_nombre}"):
            edit_nombre = st.text_input(
                "Nombre y Apellido",
                value=p["nombre"],
                key=f"f_ed_nom_{p_id}",
            )
            edit_grupo = st.selectbox(
                "Grupo Asignado",
                [
                    "Grupo 1 (Lunes, Miércoles, Viernes)",
                    "Grupo 2 (Martes, Jueves, Sábados)",
                ],
                index=0 if p["grupo"] == "MWF" else 1,
                key=f"f_ed_grp_{p_id}",
            )
            edit_turno = st.selectbox(
                "Turno Habitual",
                TURNOS_DISPONIBLES,
                index=(
                    TURNOS_DISPONIBLES.index(p["turno"])
                    if p["turno"] in TURNOS_DISPONIBLES
                    else 0
                ),
                key=f"f_ed_trn_{p_id}",
            )
            edit_maps = st.text_input(
                "Enlace de Google Maps",
                value=p.get("maps_link", ""),
                key=f"f_ed_map_{p_id}",
            )

            if st.form_submit_button(
                f"💾 Guardar cambios de {p_nombre}", use_container_width=True
            ):
              nuevo_grupo_val = "MWF" if "Grupo 1" in edit_grupo else "TTS"
              ok, err = guardar_paciente_db(
                  p_id,
                  nuevo_grupo_val,
                  edit_nombre.strip(),
                  edit_turno,
                  edit_maps.strip(),
                  p.get("contactos", []),
              )
              if ok:
                st.success(f"¡Datos de {p_nombre} actualizados con éxito!")
                st.rerun()
              else:
                st.error(f"Error: {err}")

        st.markdown("---")
        submitted = st.form_submit_button(
            "💾 GUARDAR CAMBIOS DEL DÍAS",
            type="primary",
            use_container_width=True,
        )

        if submitted:
          try:
            for p_nombre, vals in form_estados.items():
              if vals["viaja"]:
                info_p = {
                    "ida": vals["ida"],
                    "vuelta": vals["vuelta"],
                    "listo": (vals["ida"] and vals["vuelta"]),
                    "turno_dia": vals["turno_dia"],
                }
                actualizar_valores_paciente(info_p)

                # Guardar en Supabase de forma masiva
                existing = (
                    supabase.table("registros")
                    .select("id")
                    .eq("fecha", fecha_str)
                    .eq("nombre", p_nombre)
                    .execute()
                )
                payload = {
                    "fecha": fecha_str,
                    "nombre": p_nombre,
                    "ida": info_p["ida"],
                    "vuelta": info_p["vuelta"],
                    "listo": info_p["listo"],
                    "estado": info_p["estado"],
                    "valor": info_p["valor"],
                    "turno_dia": info_p["turno_dia"],
                }
                if existing.data:
                  supabase.table("registros").update(payload).eq(
                      "id", existing.data[0]["id"]
                  ).execute()
                else:
                  supabase.table("registros").insert(payload).execute()
              else:
                # Si se desmarcó, eliminar el registro de ese día si existía
                supabase.table("registros").delete().eq(
                    "fecha", fecha_str
                ).eq("nombre", p_nombre).execute()

            st.cache_data.clear()
            st.success(
                "¡Todos los cambios del día se guardaron correctamente!"
            )
            st.rerun()
          except Exception as e:
            st.error(f"Error al guardar los registros del día: {e}")

      # --- VISTA RÁPIDA DE CONTROL DEL DÍA ---
      datos_frescos = cargar_registros_db()
      pacientes_hoy_dict = datos_frescos.get(fecha_str, {})

      if pacientes_hoy_dict:
        st.markdown("---")
        st.subheader("📋 Resumen del Día Seleccionado")

        col_r1, col_r2 = st.columns(2)
        with col_r1:
          st.markdown("### ⏰ Turno 05:40 AM")
          t1_list = [
              (nom, inf)
              for nom, inf in pacientes_hoy_dict.items()
              if inf.get("turno_dia") == "05:40"
          ]
          if not t1_list:
            st.write("*(Sin pacientes)*")
          for nom, inf in t1_list:
            st.markdown(
                f"- **{nom}**: {inf['estado']} (${inf['valor']:,})"
            )

        with col_r2:
          st.markdown("### ⏰ Turno 08:40 AM")
          t2_list = [
              (nom, inf)
              for nom, inf in pacientes_hoy_dict.items()
              if inf.get("turno_dia") == "08:40"
          ]
          if not t2_list:
            st.write("*(Sin pacientes)*")
          for nom, inf in t2_list:
            st.markdown(
                f"- **{nom}**: {inf['estado']} (${inf['valor']:,})"
            )

        st.markdown("")
        datos_del_dia = {fecha_str: pacientes_hoy_dict}
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
