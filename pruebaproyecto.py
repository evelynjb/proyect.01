import streamlit as st #importamos la libreria 

st.set_page_config(page_title="Dengue colombia") #Titulo de la pagina web


if "pagina" not in st.session_state: #Crea la página de inicio 
    st.session_state.pagina = "inicio"

if "usuarios" not in st.session_state:
    st.session_state.usuarios = { "pauli": "1234" }

codigo = ["234", "456", "675"] #Guardo en una lista los codigos de los hospitales para que la persona pueda crear la cuenta 

if st.session_state.pagina == "inicio": #Página de inicio de sesión 

    st.title("Dengue Colombia") #El titulo de la bandeja 
    st.write("Sistema de vigilancia epidemiológica en Colombia") #Es como poner un subtitulo

    usuario = st.text_input("Usuario") #Creamos una entrada que sea el usuario 
    contraseña = st.text_input("Contraseña", type="password") #Creamos una entrada que sera la contraseña 

    if st.button("Iniciar sesión"): #Creamos el boton para que inicie sesion 

        if usuario in st.session_state.usuarios and st.session_state.usuarios[usuario] == contraseña: #Si el usuario esta en el diccionario 
            st.success("Bienvenido") 
        else:
            st.error("Usuario incorrecto")

    if st.button("Registrarse"):  #Creamos el boton para registrarse 
        st.session_state.pagina = "registro"
        st.rerun()

elif st.session_state.pagina == "registro": #Cuando se oprime este boton se va a la pagina de registro 

    st.title("Crear una cuenta") 

    nombre = st.text_input("Nombre completo")
    puesto = st.text_input("Cargo") 
    codigo_del_hospital = st.text_input("Código del hospital")
    nuevo_usuario = st.text_input("Nombre de usuario")
    contraseña = st.text_input("Crear contraseña", type="password")

    if st.button("Crear cuenta"):

        if codigo_del_hospital in codigo:
            if nombre != "" and nuevo_usuario != "" and contraseña != "":

                if nuevo_usuario in st.state_session.usuarios:
                    st.error("Ese nombre de usuario ya existe, elige otro.")
                else:
                    st.session_state.usuarios[nuevo_usuario] = contraseña
                    st.success("Cuenta creada, da clic a 'Volver' para iniciar sesión.")       
            
        else:
            st.error("Código incorrecto")

    if st.button("Volver"):
        st.session_state.pagina = "inicio"
        st.rerun()
