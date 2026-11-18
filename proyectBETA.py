
import os
import re
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
 
# --- VALIDACIONES ---
# Letras (con tildes y ñ), separadas por espacio, apóstrofo o guion
PATRON_NOMBRE = re.compile(
    r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+(?:[ '\-][A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+)*"
)
 
 
def validar_registro(nombre, tipo_doc, documento, sintomas):
  """Devuelve (errores, nombre_limpio, documento_limpio, sintomas_limpios)."""
  errores = []
 
  # Limpiar: quita espacios al inicio/final y colapsa espacios dobles
  nombre = " ".join(nombre.split())
  documento = documento.strip()
  sintomas = " ".join(sintomas.split())
 
  # Nombre: solo letras, mínimo 3 caracteres
  if len(nombre) < 3 or not PATRON_NOMBRE.fullmatch(nombre):
    errores.append(
        "El nombre debe tener al menos 3 caracteres y solo letras"
        " (sin números ni símbolos)."
    )
 
  # Documento: solo números, salvo el pasaporte (que lleva letras)
  if tipo_doc == "Pasaporte":
    if not re.fullmatch(r"[A-Za-z0-9]{5,20}", documento):
      errores.append(
          "El pasaporte debe tener entre 5 y 20 caracteres"
          " (letras y números, sin espacios)."
      )
  else:
    if not re.fullmatch(r"[0-9]{5,15}", documento):
      errores.append(
          "El documento debe contener solo números (entre 5 y 15 dígitos),"
          " sin letras, espacios ni puntos."
      )
 
  # Síntomas: no vacío y con al menos una letra
  if len(sintomas) < 3 or not re.search(r"[^\W\d_]", sintomas):
    errores.append(
        "Describe los síntomas con texto (mínimo 3 caracteres, no solo"
        " espacios o números)."
    )
 
  return errores, nombre, documento, sintomas
 
 
def documento_duplicado(df, tipo_doc, documento, excluir_registro=None):
  """True si ya existe otro registro con el mismo tipo y número de documento."""
  coincide = (df["Documento"] == documento) & (df["Tipo de Documento"] == tipo_doc)
  if excluir_registro is not None:
    coincide = coincide & (df["Registro"] != excluir_registro)
  return bool(coincide.any())
 
 
def cargar_y_ajustar_datos():
  df = None
  if os.path.exists(ARCHIVO_EXCEL):
    try:
      # Las columnas de texto se leen SIEMPRE como str (evita el error de int64)
      df = pd.read_excel(
          ARCHIVO_EXCEL,
          dtype={
              "Nombre": str,
              "Tipo de Documento": str,
              "Documento": str,
              "Género": str,
              "Síntomas": str,
              "EPS": str,
          },
      )
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
 
  # Normalizar tipos
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
      nuevo_nombre = st.text_input(
          "Nombre:", value=nombre_actual, max_chars=100
      )
      nuevo_tipo_doc = st.selectbox(
          "Tipo de Documento:", LISTA_TIPOS_DOC, index=idx_doc
      )
      nuevo_documento = st.text_input(
          "Documento de Identidad:", value=doc_actual, max_chars=20
      )
      nueva_edad = st.number_input(
          "Edad:", min_value=0, max_value=120, value=edad_actual
      )
      nuevo_genero = st.selectbox("Género:", LISTA_GENEROS, index=idx_gen)
      nuevos_sintomas = st.text_area(
          "Síntomas:", value=sintomas_actual, max_chars=500
      )
      nueva_eps = st.selectbox("EPS:", LISTA_EPS_COLOMBIA, index=idx_eps)
 
      guardar_edicion = st.form_submit_button(
          f"Guardar cambios del Registro #{num_editar}"
      )
      # Validación y guardado de cambios, incluyendo verificación de duplicados. 
      # Lo que busca este fragmento es asegurarse de que los datos ingresados sean correctos y que no haya conflictos con otros registros existentes antes de registrar en la base de datos y mostrar.
      if guardar_edicion:
        errores, nombre_l, doc_l, sint_l = validar_registro(
            nuevo_nombre, nuevo_tipo_doc, nuevo_documento, nuevos_sintomas
        )
        if not errores and documento_duplicado(
            df, nuevo_tipo_doc, doc_l, excluir_registro=num_editar
        ):
          errores.append("Ya existe otro registro con ese documento.")
 
        if errores:
          for e in errores:
            st.error(e)
        else:
          df.at[indice_fila, "Nombre"] = nombre_l
          df.at[indice_fila, "Tipo de Documento"] = nuevo_tipo_doc
          df.at[indice_fila, "Documento"] = doc_l
          df.at[indice_fila, "Edad"] = nueva_edad
          df.at[indice_fila, "Género"] = nuevo_genero
          df.at[indice_fila, "Síntomas"] = sint_l
          df.at[indice_fila, "EPS"] = nueva_eps
 
          guardar_datos(df)
          st.success(
              f"¡Los datos del **Registro #{num_editar}** fueron actualizados"
              " correctamente!"
          )
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
 
  # Contador para reiniciar el formulario solo cuando el guardado es exitoso
  if "form_n" not in st.session_state:
    st.session_state.form_n = 0
  if "mensaje_ok" in st.session_state:
    st.success(st.session_state.pop("mensaje_ok"))
 
  with st.form(f"form_nuevo_registro_{st.session_state.form_n}"):
    nombre = st.text_input("Nombre:", max_chars=100)
    tipo_doc = st.selectbox("Tipo de Documento:", LISTA_TIPOS_DOC)
    documento = st.text_input("Documento de Identidad:", max_chars=20)
    edad = st.number_input("Edad:", min_value=0, max_value=120, value=18)
    genero = st.selectbox("Género:", LISTA_GENEROS)
    sintomas = st.text_area("Síntomas:", max_chars=500)
    eps = st.selectbox("EPS:", LISTA_EPS_COLOMBIA)
 
    enviado = st.form_submit_button(f" Registrar como #{siguiente_numero}")
 
  if enviado:
    errores, nombre_l, doc_l, sint_l = validar_registro(
        nombre, tipo_doc, documento, sintomas
    )
    if not errores and documento_duplicado(df, tipo_doc, doc_l):
      errores.append("Ya existe un registro con ese documento.")
 
    if errores:
      for e in errores:
        st.error(e)
    else:
      nuevo_dato = {
          "Registro": siguiente_numero,
          "Nombre": nombre_l,
          "Tipo de Documento": tipo_doc,
          "Documento": doc_l,
          "Edad": edad,
          "Género": genero,
          "Síntomas": sint_l,
          "EPS": eps,
      }
      nuevo_df = pd.DataFrame([nuevo_dato])
      df_actualizado = pd.concat([df, nuevo_df], ignore_index=True)
      guardar_datos(df_actualizado)
      st.session_state.mensaje_ok = (
          f"¡Dato guardado exitosamente con el número de registro"
          f" **{siguiente_numero}**!"
      )
      st.session_state.form_n += 1  # formulario nuevo y vacío
      st.rerun()