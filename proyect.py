import os
import pandas as pd
import streamlit as st

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Gestión de Pacientes", page_icon="", layout="wide"
)

# --- ESTILOS PERSONALIZADOS (AZUL NAVY Y VERDE COMPLEMENTARIO) ---
st.markdown(
    """
    <style>
    /* Fondo principal Azul Navy */
    .stApp {
        background-color: #000000;
        color: #F8FAFC;
    }
    /* Títulos */
    h1, h2, h3 {
        color: #10B981 !important;
    }
    /* Botones principales en Verde */
    .stButton>button, div[data-testid="stFormSubmitButton"]>button {
        background-color: #10B981;
        color: #0B192C;
        font-weight: bold;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1.2rem;
        transition: 0.3s;
    }
    .stButton>button:hover, div[data-testid="stFormSubmitButton"]>button:hover {
        background-color: #059669;
        color: #FFFFFF;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# --- CONFIGURACIÓN DEL ARCHIVO Y LISTAS ---
ARCHIVO_EXCEL = "BdatosD.xlsx"

COLUMNAS_EXACTAS = [
    "Registro",
    "Nombre",
    "Tipo de Documento",
    "Documento",
    "Edad",
    "Género",
    "Síntomas",
    "EPS",
]

LISTA_TIPOS_DOC = [
    "Cédula de Ciudadanía (CC)",
    "Tarjeta de Identidad (TI)",
    "Registro Civil (RC)",
    "Cédula de Extranjería (CE)",
    "Permiso por Protección Temporal (PPT)",
    "Pasaporte",
]

LISTA_GENEROS = ["Masculino", "Femenino", "Otro"]

LISTA_EPS_COLOMBIA = [
    "No tengo",
    "Sura EPS",
    "Nueva EPS",
    "Sanitas",
    "Salud Total",
    "Compensar",
    "Famisanar",
    "Coosalud",
    "Mutual Ser",
    "Aliansalud",
    "Cajacopi",
    "Capital Salud",
    "Comfenalco Valle",
    "Savia Salud",
    "Asmet Salud",
    "Emssanar",
    "Otra",
]


def cargar_y_ajustar_datos():
  df = None
  if os.path.exists(ARCHIVO_EXCEL):
    try:
      df = pd.read_excel(ARCHIVO_EXCEL)
    except Exception:
      # Si el archivo está corrupto (como ocurrió con el EOFError), se crea uno nuevo limpio
      df = None

  if df is None or df.empty:
    df = pd.DataFrame(columns=COLUMNAS_EXACTAS)

  if "Registro" not in df.columns:
    df.insert(0, "Registro", range(1, len(df) + 1))

  for col in COLUMNAS_EXACTAS:
    if col not in df.columns:
      if col == "Edad":
        df[col] = 0
      elif col == "Tipo de Documento":
        df[col] = LISTA_TIPOS_DOC[0]
      elif col == "Género":
        df[col] = LISTA_GENEROS[0]
      elif col == "EPS":
        df[col] = LISTA_EPS_COLOMBIA[0]
      else:
        df[col] = ""

  df = df[COLUMNAS_EXACTAS].copy()
  for col in ["Nombre", "Tipo de Documento", "Documento", "Género", "Síntomas", "EPS"]:
    df[col] = df[col].fillna("").astype(str)
  df["Registro"] = pd.to_numeric(df["Registro"], errors="coerce").fillna(0).astype(int)
  df["Edad"] = pd.to_numeric(df["Edad"], errors="coerce").fillna(0).astype(int)
  try:
    df.to_excel(ARCHIVO_EXCEL, index=False)
  except Exception:
    pass
  return df


def guardar_datos(df):
  df = df[COLUMNAS_EXACTAS]
  df.to_excel(ARCHIVO_EXCEL, index=False)


df = cargar_y_ajustar_datos()

opciones_menu = ["Consultar Registro", "Modificar Registro", "Nuevo Registro"]
if "menu_actual" not in st.session_state:
  st.session_state.menu_actual = "Consultar Registro"


def ir_a_nuevo_registro():
  st.session_state.menu_actual = "Nuevo Registro"


st.title(" Sistema de Gestión y Consulta de Registros")

opcion = st.selectbox(
    "Selecciona la acción que deseas realizar:", opciones_menu, key="menu_actual"
)

st.markdown("---")

# ==========================================
# 1. CONSULTAR REGISTRO
# ==========================================
if opcion == "Consultar Registro":
  st.subheader("Consulta por Número de Registro")
  max_reg = int(df["Registro"].max()) if not df.empty else 1

  num_consulta = st.number_input(
      f"Ingresa el número de registro a consultar (1 al {max_reg}):",
      min_value=1,
      value=1,
      step=1,
  )

  resultado = df[df["Registro"] == num_consulta]

  if not resultado.empty:
    st.success(f"Información del **Registro #{num_consulta}**:")
    fila = resultado.iloc[0]

    col1, col2 = st.columns(2)
    with col1:
      st.info(f"**Nombre:** {fila['Nombre']}")
      st.info(f"**Tipo de Documento:** {fila['Tipo de Documento']}")
      st.info(f"**Documento de Identidad:** {fila['Documento']}")
      st.info(f"**Edad:** {fila['Edad']} años")
    with col2:
      st.info(f"**Género:** {fila['Género']}")
      st.info(f"**EPS:** {fila['EPS']}")
      st.info(f"**Síntomas:** {fila['Síntomas']}")
  else:
    st.warning(f"El registro **#{num_consulta}** aún no existe en la base de datos.")
    st.button("Ir a Nuevo Registro", on_click=ir_a_nuevo_registro)

# ==========================================
# 2. MODIFICAR REGISTRO
# ==========================================
elif opcion == "Modificar Registro":
  st.subheader("Modificar Datos de un Registro")

  num_editar = st.number_input(
      "Ingresa el número del usuario/fila que deseas modificar (ej. 2):",
      min_value=1,
      value=1,
      step=1,
  )

  registro_encontrado = df[df["Registro"] == num_editar]

  if not registro_encontrado.empty:
    datos_actuales = registro_encontrado.iloc[0]
    indice_fila = registro_encontrado.index[0]

    st.success(f"Editando datos del **Registro #{num_editar}**:")

    idx_doc = (
        LISTA_TIPOS_DOC.index(datos_actuales["Tipo de Documento"])
        if datos_actuales["Tipo de Documento"] in LISTA_TIPOS_DOC
        else 0
    )
    idx_gen = (
        LISTA_GENEROS.index(datos_actuales["Género"])
        if datos_actuales["Género"] in LISTA_GENEROS
        else 0
    )
    idx_eps = (
        LISTA_EPS_COLOMBIA.index(datos_actuales["EPS"])
        if datos_actuales["EPS"] in LISTA_EPS_COLOMBIA
        else 0
    )

    edad_actual = (
        int(datos_actuales["Edad"]) if pd.notna(datos_actuales["Edad"]) else 0
    )
    doc_actual = (
        str(datos_actuales["Documento"])
        if pd.notna(datos_actuales["Documento"])
        else ""
    )
    nombre_actual = (
        str(datos_actuales["Nombre"])
        if pd.notna(datos_actuales["Nombre"])
        else ""
    )
    sintomas_actual = (
        str(datos_actuales["Síntomas"])
        if pd.notna(datos_actuales["Síntomas"])
        else ""
    )

    with st.form("form_modificar_registro"):
      nuevo_nombre = st.text_input("Nombre:", value=nombre_actual)
      nuevo_tipo_doc = st.selectbox(
          "Tipo de Documento:", LISTA_TIPOS_DOC, index=idx_doc
      )
      nuevo_documento = st.text_input(
          "Documento de Identidad:", value=doc_actual
      )
      nueva_edad = st.number_input(
          "Edad:", min_value=0, max_value=120, value=edad_actual
      )
      nuevo_genero = st.selectbox("Género:", LISTA_GENEROS, index=idx_gen)
      nuevos_sintomas = st.text_area("Síntomas:", value=sintomas_actual)
      nueva_eps = st.selectbox("EPS:", LISTA_EPS_COLOMBIA, index=idx_eps)

      guardar_edicion = st.form_submit_button(
          f"Guardar cambios del Registro #{num_editar}"
      )

      if guardar_edicion:
        df.at[indice_fila, "Nombre"] = nuevo_nombre
        df.at[indice_fila, "Tipo de Documento"] = nuevo_tipo_doc
        df.at[indice_fila, "Documento"] = nuevo_documento
        df.at[indice_fila, "Edad"] = nueva_edad
        df.at[indice_fila, "Género"] = nuevo_genero
        df.at[indice_fila, "Síntomas"] = nuevos_sintomas
        df.at[indice_fila, "EPS"] = nueva_eps

        guardar_datos(df)
        st.success(
            f"¡Los datos del **Registro #{num_editar}** fueron actualizados"
            " correctamente!"
        )
        st.rerun()
  else:
    st.warning(
        f" El número **{num_editar}** aún no tiene ningún dato registrado en"
        " la base de datos."
    )
    st.write("¿Deseas crear un nuevo reporte en la base de datos?")
    st.button(" Ir a Nuevo Registro", on_click=ir_a_nuevo_registro)

# ==========================================
# 3. NUEVO REGISTRO
# ==========================================
elif opcion == "Nuevo Registro":
  siguiente_numero = int(df["Registro"].max()) + 1 if not df.empty else 1
  st.subheader(
      f" Agregar Nuevo Registro (Se asignará el #{siguiente_numero})"
  )

  with st.form("form_nuevo_registro", clear_on_submit=True):
    nombre = st.text_input("Nombre:")
    tipo_doc = st.selectbox("Tipo de Documento:", LISTA_TIPOS_DOC)
    documento = st.text_input("Documento de Identidad:")
    edad = st.number_input("Edad:", min_value=0, max_value=120, value=18)
    genero = st.selectbox("Género:", LISTA_GENEROS)
    sintomas = st.text_area("Síntomas:")
    eps = st.selectbox("EPS:", LISTA_EPS_COLOMBIA)

    enviado = st.form_submit_button(f" Registrar como #{siguiente_numero}")

    if enviado:
      nuevo_dato = {
          "Registro": siguiente_numero,
          "Nombre": nombre,
          "Tipo de Documento": tipo_doc,
          "Documento": documento,
          "Edad": edad,
          "Género": género if "género" in locals() else genero,
          "Síntomas": sintomas,
          "EPS": eps,
      }
      nuevo_df = pd.DataFrame([nuevo_dato])
      df_actualizado = pd.concat([df, nuevo_df], ignore_index=True)
      guardar_datos(df_actualizado)
      st.success(
          f"¡Dato guardado exitosamente con el número de registro"
          f" **{siguiente_numero}**!"
      )
      st.rerun()