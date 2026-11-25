import os
import re
import matplotlib.pyplot as plt
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Dengue Colombia", layout="wide")

if "pagina" not in st.session_state:
  st.session_state.pagina = "inicio"

if "seccion" not in st.session_state:
  st.session_state.seccion = "Inicio"

if "usuario_actual" not in st.session_state:
  st.session_state.usuario_actual = ""

codigo = ["234", "456", "675"]
ARCHIVO_USUARIOS = "usuarios.xlsx"


# --- FUNCIONES DE USUARIOS (OPENPYXL) ---
def inicializar_usuarios():
  if not os.path.exists(ARCHIVO_USUARIOS):
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.append(["Nombre", "Cargo", "Usuario", "Contraseña"])
    libro.save(ARCHIVO_USUARIOS)


def guardar_usuario(nombre, cargo, usuario, contraseña):
  inicializar_usuarios()
  libro = openpyxl.load_workbook(ARCHIVO_USUARIOS)
  hoja = libro.active
  hoja.append([nombre, cargo, usuario, contraseña])
  libro.save(ARCHIVO_USUARIOS)


def verificar_usuario(usuario, contraseña):
  inicializar_usuarios()
  libro = openpyxl.load_workbook(ARCHIVO_USUARIOS)
  hoja = libro.active
  for fila in hoja.iter_rows(min_row=2, values_only=True):
    if str(fila[2]) == usuario and str(fila[3]) == contraseña:
      return True
  return False


def existe_usuario(usuario):
  inicializar_usuarios()
  libro = openpyxl.load_workbook(ARCHIVO_USUARIOS)
  hoja = libro.active
  for fila in hoja.iter_rows(min_row=2, values_only=True):
    if str(fila[2]) == usuario:
      return True
  return False


# ==========================================
# VISTA 1: INICIO DE SESIÓN
# ==========================================
if st.session_state.pagina == "inicio":
  st.title("Dengue Colombia")
  st.write("Sistema de vigilancia epidemiológica en Colombia")

  usuario = st.text_input("Usuario")
  contraseña = st.text_input("Contraseña", type="password")

  if st.button("Iniciar sesión"):
    if verificar_usuario(usuario, contraseña):
      st.session_state.pagina = "principal"
      st.session_state.usuario_actual = usuario
      st.rerun()
    else:
      st.error("Usuario o contraseña incorrectos")

  if st.button("Registrarse"):
    st.session_state.pagina = "registro"
    st.rerun()


# ==========================================
# VISTA 2: REGISTRO DE USUARIOS
# ==========================================
elif st.session_state.pagina == "registro":
  st.title("Crear una cuenta")

  nombre = st.text_input("Nombre completo")
  puesto = st.text_input("Cargo")
  codigo_del_hospital = st.text_input("Código del hospital")
  nuevo_usuario = st.text_input("Nombre de usuario")
  contraseña = st.text_input("Crear contraseña", type="password")

  if st.button("Crear cuenta"):
    if codigo_del_hospital in codigo:
      if (
          nombre != ""
          and puesto != ""
          and nuevo_usuario != ""
          and contraseña != ""
      ):
        if existe_usuario(nuevo_usuario):
          st.error("Este usuario ya existe")
        else:
          guardar_usuario(nombre, puesto, nuevo_usuario, contraseña)
          st.success("Cuenta creada correctamente")
      else:
        st.error("Completa todos los campos")
    else:
      st.error("Código del hospital incorrecto")

  if st.button("Volver"):
    st.session_state.pagina = "inicio"
    st.rerun()


# ==========================================
# VISTA 3: SISTEMA PRINCIPAL DEL HOSPITAL
# ==========================================
elif st.session_state.pagina == "principal":
  st.markdown(
      """
        <style>
        .stApp {
            background-color: #000000;
            color: #F8FAFC;
        }
        h1, h2, h3 {
            color: #10B981 !important;
        }
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
  PATRON_NOMBRE = re.compile(
      r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+(?:[ '\-][A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+)*"
  )

  def validar_registro(nombre, tipo_doc, documento, sintomas):
    errores = []
    nombre = " ".join(nombre.split())
    documento = documento.strip()
    sintomas = " ".join(sintomas.split())

    if len(nombre) < 3 or not PATRON_NOMBRE.fullmatch(nombre):
      errores.append(
          "El nombre debe tener al menos 3 caracteres y solo letras (sin"
          " números ni símbolos)."
      )

    if tipo_doc == "Pasaporte":
      if not re.fullmatch(r"[A-Za-z0-9]{5,20}", documento):
        errores.append(
            "El pasaporte debe tener entre 5 y 20 caracteres (letras y números,"
            " sin espacios)."
        )
    else:
      if not re.fullmatch(r"[0-9]{5,15}", documento):
        errores.append(
            "El documento debe contener solo números (entre 5 y 15 dígitos),"
            " sin letras, espacios ni puntos."
        )

    if len(sintomas) < 3 or not re.search(r"[^\W\d_]", sintomas):
      errores.append(
          "Describe los síntomas con texto (mínimo 3 caracteres, no solo"
          " espacios o números)."
      )

    return errores, nombre, documento, sintomas

  def documento_duplicado(df, tipo_doc, documento, excluir_registro=None):
    coincide = (df["Documento"] == documento) & (
        df["Tipo de Documento"] == tipo_doc
    )
    if excluir_registro is not None:
      coincide = coincide & (df["Registro"] != excluir_registro)
    return bool(coincide.any())

  def cargar_y_ajustar_datos():
    df = None
    if os.path.exists(ARCHIVO_EXCEL):
      try:
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

    for col in [
        "Nombre",
        "Tipo de Documento",
        "Documento",
        "Género",
        "Síntomas",
        "EPS",
    ]:
      df[col] = df[col].fillna("").astype(str)
    df["Registro"] = (
        pd.to_numeric(df["Registro"], errors="coerce").fillna(0).astype(int)
    )
    df["Edad"] = (
        pd.to_numeric(df["Edad"], errors="coerce").fillna(0).astype(int)
    )

    try:
      df.to_excel(ARCHIVO_EXCEL, index=False)
    except Exception:
      pass
    return df

  def guardar_datos(df):
    df = df[COLUMNAS_EXACTAS]
    df.to_excel(ARCHIVO_EXCEL, index=False)

  df = cargar_y_ajustar_datos()

  # Encabezado y botón para cerrar sesión
  col_titulo, col_salir = st.columns([6, 1])
  with col_titulo:
    st.title("Sistema de Gestión y Consulta de Registros")
    st.write(f"Usuario activo: {st.session_state.usuario_actual}")
  with col_salir:
    if st.button("Cerrar sesión"):
      st.session_state.pagina = "inicio"
      st.session_state.usuario_actual = ""
      st.rerun()

  opciones_menu = [
      "Consultar Registro",
      "Modificar Registro",
      "Nuevo Registro",
      "Estadísticas",
  ]
  if "menu_actual" not in st.session_state:
    st.session_state.menu_actual = "Consultar Registro"

  def ir_a_nuevo_registro():
    st.session_state.menu_actual = "Nuevo Registro"

  opcion = st.selectbox(
      "Selecciona la acción que deseas realizar:",
      opciones_menu,
      key="menu_actual",
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
      st.success(f"Información del Registro #{num_consulta}:")
      fila = resultado.iloc[0]

      col1, col2 = st.columns(2)
      with col1:
        st.info(f"Nombre: {fila['Nombre']}")
        st.info(f"Tipo de Documento: {fila['Tipo de Documento']}")
        st.info(f"Documento de Identidad: {fila['Documento']}")
        st.info(f"Edad: {fila['Edad']} años")
      with col2:
        st.info(f"Género: {fila['Género']}")
        st.info(f"EPS: {fila['EPS']}")
        st.info(f"Síntomas: {fila['Síntomas']}")
    else:
      st.warning(
          f"El registro #{num_consulta} aún no existe en la base de datos."
      )
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

      st.success(f"Editando datos del Registro #{num_editar}:")

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
                f"Los datos del Registro #{num_editar} fueron actualizados"
                " correctamente."
            )
    else:
      st.warning(
          f"El número {num_editar} aún no tiene ningún dato registrado en la"
          " base de datos."
      )
      st.write("¿Deseas crear un nuevo reporte en la base de datos?")
      st.button("Ir a Nuevo Registro", on_click=ir_a_nuevo_registro)

  # ==========================================
  # 3. NUEVO REGISTRO
  # ==========================================
  elif opcion == "Nuevo Registro":
    siguiente_numero = int(df["Registro"].max()) + 1 if not df.empty else 1
    st.subheader(f"Agregar Nuevo Registro (Se asignará el #{siguiente_numero})")

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

      enviado = st.form_submit_button(f"Registrar como #{siguiente_numero}")

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
            f"Dato guardado exitosamente con el número de registro"
            f" #{siguiente_numero}."
        )
        st.session_state.form_n += 1
        st.rerun()

  # ==========================================
  # 4. ESTADÍSTICAS
  # ==========================================
  elif opcion == "Estadísticas":
    st.subheader("Estadísticas de Pacientes")

    if df.empty:
      st.warning("No hay datos registrados para generar gráficas.")
    else:
      st.write(f"Total de pacientes registrados: {len(df)}")
      st.write(f"Promedio de edad: {df['Edad'].mean():.1f} años")

      # Estilo oscuro sencillo para matplotlib
      plt.style.use("dark_background")

      col1, col2 = st.columns(2)

      # Gráfica 1: Género
      with col1:
        st.write("Distribución por Género")
        conteo_genero = df["Género"].value_counts().reindex(LISTA_GENEROS, fill_value=0)
        fig1, ax1 = plt.subplots()
        ax1.bar(
            conteo_genero.index,
            conteo_genero.values,
            color="#1D1F5C",
        )
        ax1.set_ylabel("Cantidad de pacientes")
        ax1.set_xlabel("Género")
        st.pyplot(fig1)
        

      # Gráfica 2: Edad
      with col2:
        st.write("Distribución por Edad")
        fig2, ax2 = plt.subplots()
        ax2.hist(df["Edad"], bins=10, color="#12426B", edgecolor="black")
        ax2.set_ylabel("Cantidad de pacientes")
        ax2.set_xlabel("Edad")
        st.pyplot(fig2)

      # Gráfica 3: EPS
      st.write("Distribución por EPS")
      conteo_eps = df["EPS"].value_counts()
      fig3, ax3 = plt.subplots(figsize=(10, 4))
      ax3.barh(conteo_eps.index, conteo_eps.values, color="#10B981")
      ax3.set_xlabel("Cantidad de pacientes")
      ax3.set_ylabel("EPS")
      ax3.invert_yaxis()
      st.pyplot(fig3)