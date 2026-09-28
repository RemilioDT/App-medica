import warnings
warnings.filterwarnings("ignore")

import streamlit as st
import os
import json
import random
import re
import time
import pdfplumber
from pptx import Presentation
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

# Motor de IA + búsqueda en la teoría (archivo motor_ia.py, en la misma carpeta)
from motor_ia import (
    generar_respuesta_con_fallback,
    extraer_json,
    buscar_fragmentos,
    explicar_con_teoria,
)

# ==========================================
# 1. CONFIGURACIÓN DE RUTAS Y BASE DE DATOS
# ==========================================
st.set_page_config(page_title="Tutor Médico AI - Multimotor", layout="wide")

ARCHIVO_BD = "usuarios_bd.json"
CARPETA_NUBE = "nube_medica"

if not os.path.exists(ARCHIVO_BD):
    with open(ARCHIVO_BD, "w") as f:
        json.dump({"admin": {"password": "admin", "role": "admin"}}, f)

def asegurar_carpetas_usuario(usuario):
    rutas = [
        os.path.join(CARPETA_NUBE, usuario, "bancos"),
        os.path.join(CARPETA_NUBE, usuario, "teoria")
    ]
    for r in rutas:
        os.makedirs(r, exist_ok=True)
    return rutas

asegurar_carpetas_usuario("admin")

def cargar_usuarios():
    with open(ARCHIVO_BD, "r") as f:
        return json.load(f)

def guardar_usuarios(db):
    with open(ARCHIVO_BD, "w") as f:
        json.dump(db, f, indent=4)

# ==========================================
# 2. MOTOR DE EXTRACCIÓN Y LIMPIEZA INTELIGENTE
# ==========================================
def extraer_texto_plano_con_paginas(ruta_archivo):
    texto_total = ""
    if ruta_archivo.lower().endswith(".pdf"):
        with pdfplumber.open(ruta_archivo) as pdf:
            for i, pagina in enumerate(pdf.pages):
                t = pagina.extract_text()
                if t:
                    texto_total += f"\n[Página interna del PDF: {i+1}] {t}\n"
    elif ruta_archivo.lower().endswith((".pptx", ".ppt")):
        prs = Presentation(ruta_archivo)
        for i, slide in enumerate(prs.slides):
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    texto_total += f"\n[Diapositiva {i+1}] {shape.text}\n"
    elif ruta_archivo.lower().endswith((".png", ".jpg", ".jpeg")):
        try:
            from PIL import Image
            import google.generativeai as genai
            
            api_key = st.secrets.get("GOOGLE_API_KEY", "")
            genai.configure(api_key=api_key)
            modelo_vision = genai.GenerativeModel('gemini-1.5-flash')
            
            img = Image.open(ruta_archivo)
            prompt_ocr = "Extrae todo el texto, tablas, esquemas y datos médicos de esta imagen con sumo detalle. Si hay tablas o asociaciones de enfermedades con síntomas, descríbelas claramente. Es para una base de datos teórica médica."
            respuesta = modelo_vision.generate_content([prompt_ocr, img])
            
            texto_total += f"\n[Contenido extraído de Imagen: {os.path.basename(ruta_archivo)}]\n{respuesta.text}\n"
        except Exception as e:
            texto_total += f"\n[Error extrayendo imagen: {str(e)}]\n"
            
    return texto_total

def extraer_preguntas_local_pdf(ruta_archivo):
    texto_total = ""
    if ruta_archivo.lower().endswith(".pdf"):
        with pdfplumber.open(ruta_archivo) as pdf:
            for i, pagina in enumerate(pdf.pages):
                t = pagina.extract_text()
                if t: 
                    texto_total += f"\n[Página {i+1}]\n{t}\n"
    elif ruta_archivo.lower().endswith((".pptx", ".ppt", ".png", ".jpg", ".jpeg")):
        texto_total = extraer_texto_plano_con_paginas(ruta_archivo)

    # PARSER ROBUSTO POR BLOQUES: Evita que las preguntas o respuestas se corten a la mitad
    tamano_chunk = 8000
    chunks = [texto_total[i:i+tamano_chunk] for i in range(0, len(texto_total), tamano_chunk)]
    banco_preguntas = []

    for chunk in chunks:
        prompt_parser = f"""
        Eres un procesador experto de exámenes médicos (ENAM / Residentado).
        Analiza el siguiente fragmento de texto de un banco de preguntas y extrae TODAS las preguntas de opción múltiple.
        REGLAS CRÍTICAS:
        1. NO recortes ni dejes a medias los enunciados ni las opciones. Deben estar COMPLETOS.
        2. Cada pregunta debe tener sus 5 opciones (A, B, C, D, E) completas.
        3. Identifica la clave de respuesta si aparece en el texto (ej. "Rpta: B" o "Clave: A"). Si no aparece, pon "N/A".
        4. Clasifica la especialidad médica (Medicina Interna, Cirugía General, Pediatría, Gineco-Obstetricia, etc.).

        FRAGMENTO DE TEXTO:
        {chunk}

        Devuelve la respuesta ESTRICTAMENTE en formato JSON como una lista de objetos, sin texto adicional:
        [
          {{
            "pregunta": "Enunciado completo...",
            "opciones": ["A) ...", "B) ...", "C) ...", "D) ...", "E) ..."],
            "correcta": "Letra o N/A",
            "especialidad": "Especialidad"
          }}
        ]
        """
        try:
            texto_resp, _ = generar_respuesta_con_fallback(prompt_parser)
            t_limpio = texto_resp.strip()
            if t_limpio.startswith("```json"):
                t_limpio = t_limpio[7:]
            elif t_limpio.startswith("```"):
                t_limpio = t_limpio[3:]
            if t_limpio.endswith("```"):
                t_limpio = t_limpio[:-3]
            
            i, j = t_limpio.find("["), t_limpio.rfind("]")
            if i != -1 and j != -1:
                parcial = json.loads(t_limpio[i:j+1])
                if isinstance(parcial, list):
                    banco_preguntas.extend(parcial)
        except Exception:
            continue

    return banco_preguntas

# ==========================================
# 3. SESIÓN Y AUTENTICACIÓN
# ==========================================
if 'usuario_actual' not in st.session_state:
    query_params = st.query_params
    if "usr" in query_params:
        db = cargar_usuarios()
        usr_guardado = query_params["usr"]
        if usr_guardado in db:
            st.session_state.usuario_actual = usr_guardado
            st.session_state.rol_actual = db[usr_guardado]["role"]
        else:
            st.session_state.usuario_actual = None
            st.session_state.rol_actual = None
    else:
        st.session_state.usuario_actual = None
        st.session_state.rol_actual = None

def login_register_page():
    st.title("🩺 Tutor Médico AI - Acceso")
    tab1, tab2 = st.tabs(["Iniciar Sesión", "Registrarse"])

    with tab1:
        usuario_login = st.text_input("Usuario", key="log_user").strip()
        pass_login = st.text_input("Contraseña", type="password", key="log_pass")
        if st.button("Ingresar", type="primary"):
            db = cargar_usuarios()
            if usuario_login in db and db[usuario_login]["password"] == pass_login:
                st.session_state.usuario_actual = usuario_login
                st.session_state.rol_actual = db[usuario_login]["role"]
                st.query_params["usr"] = usuario_login
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")

    with tab2:
        usuario_reg = st.text_input("Nuevo Usuario", key="reg_user").strip()
        pass_reg = st.text_input("Nueva Contraseña", type="password", key="reg_pass")
        if st.button("Registrarse"):
            db = cargar_usuarios()
            if usuario_reg in db:
                st.error("Usuario existente.")
            elif len(usuario_reg) < 3 or len(pass_reg) < 3:
                st.warning("Debe tener al menos 3 caracteres.")
            else:
                db[usuario_reg] = {"password": pass_reg, "role": "estudiante"}
                guardar_usuarios(db)
                asegurar_carpetas_usuario(usuario_reg)
                st.success("¡Cuenta creada!")

# ==========================================
# 4. INTERFAZ PRINCIPAL
# ==========================================
def main_app():
    st.sidebar.title(f"Hola, {st.session_state.usuario_actual}")
    st.sidebar.markdown(f"**Rol:** {st.session_state.rol_actual.capitalize()}")
    st.sidebar.divider()

    menu = ["Simulacro Médico", "Generador ENAM", "Mis Documentos", "Mi Perfil"]
    if st.session_state.rol_actual == "admin": menu.append("Panel de Admin")
    opcion_menu = st.sidebar.radio("Navegación:", menu)
    st.sidebar.divider()

    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.usuario_actual = None
        if "usr" in st.query_params: del st.query_params["usr"]
        st.rerun()

    carpetas_usuario = asegurar_carpetas_usuario(st.session_state.usuario_actual)
    carpetas_admin = asegurar_carpetas_usuario("admin")

    if opcion_menu == "Mi Perfil":
        st.header("👤 Mi Perfil")
        st.info("Opciones de perfil listas.")

    elif opcion_menu in ["Mis Documentos", "Panel de Admin"]:
        st.header(f"📂 {opcion_menu}")
        carpeta_base = carpetas_admin if opcion_menu == "Panel de Admin" else carpetas_usuario

        tipo_subida = st.radio("¿Qué tipo de material vas a subir?", ["Banco de Preguntas", "Material Teórico (Cerebro AI)"], horizontal=True)
        archivos_subidos = st.file_uploader("Selecciona documentos", type=["pdf", "pptx", "ppt", "jpg", "jpeg", "png"], accept_multiple_files=True)

        if archivos_subidos:
            if st.button("Guardar y Procesar Archivos", type="primary"):
                destino = carpeta_base[0] if tipo_subida == "Banco de Preguntas" else carpeta_base[1]

                for archivo in archivos_subidos:
                    ruta_raw = os.path.join(destino, archivo.name)
                    barra = st.progress(0, text=f"Subiendo {archivo.name}...")
                    with open(ruta_raw, "wb") as f:
                        f.write(archivo.getbuffer())

                    barra.progress(30, text=f"Extrayendo texto y páginas de {archivo.name}...")

                    if tipo_subida == "Banco de Preguntas":
                        preguntas = extraer_preguntas_local_pdf(ruta_raw)
                        with open(ruta_raw + ".json", "w", encoding="utf-8") as f:
                            json.dump(preguntas, f, ensure_ascii=False, indent=4)
                    else:
                        texto_teoria = extraer_texto_plano_con_paginas(ruta_raw)
                        with open(ruta_raw + ".txt", "w", encoding="utf-8") as f:
                            f.write(texto_teoria)
                        if len(texto_teoria.strip()) < 200 and not ruta_raw.lower().endswith((".png", ".jpg", ".jpeg")):
                            st.warning(f"⚠️ De '{archivo.name}' casi no se pudo extraer texto. Si es un PDF escaneado, la IA no podrá leerlo.")

                    barra.progress(100, text=f"¡{archivo.name} procesado!")
                    time.sleep(1)
                    barra.empty()

                st.success(f"¡Procesados {len(archivos_subidos)} archivo(s) en '{tipo_subida}'!")
                time.sleep(1)
                st.rerun()

        st.divider()
        col_sec1, col_sec2 = st.columns(2)

        with col_sec1:
            st.subheader("📝 Bancos de Preguntas")
            archivos_banco = [f for f in os.listdir(carpeta_base[0]) if not f.endswith(".json")]

            if archivos_banco:
                bancos_sel = st.multiselect("Seleccionar bancos para borrar:", archivos_banco, key="ms_bancos")
                c_del1, c_del2 = st.columns(2)
                with c_del1:
                    if bancos_sel and st.button("🗑️ Borrar selección", key="btn_del_sel_b"):
                        for arch in bancos_sel:
                            r = os.path.join(carpeta_base[0], arch)
                            if os.path.exists(r): os.remove(r)
                            if os.path.exists(r + ".json"): os.remove(r + ".json")
                        st.success(f"Se eliminaron {len(bancos_sel)} bancos.")
                        st.rerun()
                with c_del2:
                    if st.button("⚠️ Borrar TODOS", key="btn_del_all_b"):
                        for arch in archivos_banco:
                            r = os.path.join(carpeta_base[0], arch)
                            if os.path.exists(r): os.remove(r)
                            if os.path.exists(r + ".json"): os.remove(r + ".json")
                        st.success("Se eliminaron todos los bancos.")
                        st.rerun()

                st.divider()
                for arch in archivos_banco:
                    c1, c2 = st.columns([4, 1])
                    c1.write(f"📄 {arch}")
                    if c2.button("Eliminar", key=f"del_b_{arch}"):
                        r = os.path.join(carpeta_base[0], arch)
                        os.remove(r)
                        if os.path.exists(r + ".json"): os.remove(r + ".json")
                        st.rerun()
            else:
                st.info("No hay bancos cargados.")

        with col_sec2:
            st.subheader("🧠 Material Teórico")
            archivos_teoria = [f for f in os.listdir(carpeta_base[1]) if not f.endswith(".txt")]

            if archivos_teoria:
                teoria_sel = st.multiselect("Seleccionar teoría para borrar:", archivos_teoria, key="ms_teoria")
                c_del3, c_del4 = st.columns(2)
                with c_del3:
                    if teoria_sel and st.button("🗑️ Borrar selección", key="btn_del_sel_t"):
                        for arch in teoria_sel:
                            r = os.path.join(carpeta_base[1], arch)
                            if os.path.exists(r): os.remove(r)
                            if os.path.exists(r + ".txt"): os.remove(r + ".txt")
                        st.success(f"Se eliminaron {len(teoria_sel)} archivos teóricos.")
                        st.rerun()
                with c_del4:
                    if st.button("⚠️ Borrar TODA la teoría", key="btn_del_all_t"):
                        for arch in archivos_teoria:
                            r = os.path.join(carpeta_base[1], arch)
                            if os.path.exists(r): os.remove(r)
                            if os.path.exists(r + ".txt"): os.remove(r + ".txt")
                        st.success("Se eliminó toda la teoría.")
                        st.rerun()

                st.divider()
                for arch in archivos_teoria:
                    c1, c2 = st.columns([4, 1])
                    c1.write(f"📘 {arch}")
                    if c2.button("Eliminar", key=f"del_t_{arch}"):
                        r = os.path.join(carpeta_base[1], arch)
                        os.remove(r)
                        if os.path.exists(r + ".txt"): os.remove(r + ".txt")
                        st.rerun()
            else:
                st.info("No hay material teórico.")

    elif opcion_menu == "Generador ENAM":
        st.header("⚡ Generador de Casos Clínicos (Multimotor)")
        st.write("La IA leerá tu material teórico y construirá preguntas inéditas.")

        especialidad = st.selectbox("Selecciona la especialidad a practicar:",
                                    ["Medicina Interna", "Cirugía General", "Pediatría", "Gineco-Obstetricia", "Salud Pública / Ciencias Básicas"])

        if st.button("🧬 Generar Pregunta Inédita", type="primary"):
            with st.spinner("Creando caso clínico... (Buscando el mejor motor disponible)"):
                carpetas_teoria = [carpetas_admin[1], carpetas_usuario[1]]

                try:
                    frags = buscar_fragmentos(carpetas_teoria, especialidad, k=12)
                    muestra = random.sample(frags, min(4, len(frags)))
                    memoria_limitada = "\n\n".join(
                        f"[FUENTE: {c['archivo']} | Página {c['pagina']}]\n{c['texto']}" for c in muestra
                    ) or "(no hay teoría cargada sobre este tema)"

                    prompt_gen = f"""
                    Eres un catedrático evaluador del ENAM.
                    === BASE TEÓRICA ===
                    {memoria_limitada}
                    =====================
                    Crea UNA pregunta de opción múltiple INÉDITA de la especialidad: {especialidad}.
                    Basa la pregunta en la BASE TEÓRICA. En "fuente" usa el archivo y la página de las etiquetas [FUENTE: ...].
                    Devuelve SOLO y ESTRICTAMENTE un JSON con esta estructura:
                    {{"pregunta": "Texto...", "opciones": ["A)...", "B)...", "C)...", "D)...", "E)..."], "correcta": "Letra correcta", "explicacion": "Por qué es correcta...", "fuente": "Nombre archivo y página"}}
                    """

                    texto_resp, motor_usado = generar_respuesta_con_fallback(prompt_gen)
                    datos_gen = extraer_json(texto_resp)

                    st.session_state.gen_activa = True
                    st.session_state.gen_datos = datos_gen
                    st.session_state.gen_motor = motor_usado
                    st.session_state.gen_evaluado = False
                    st.session_state.gen_mostrar_resp = False

                except Exception as e:
                    st.error(f"{e}")

        if st.session_state.get("gen_activa", False):
            datos = st.session_state.gen_datos
            st.divider()
            st.markdown(f"*(Generado por: {st.session_state.gen_motor})*")
            st.markdown(f"### {datos['pregunta']}")

            resp_usuario = st.radio("Selecciona tu respuesta:", datos['opciones'], index=None, key="radio_gen")

            if not st.session_state.get("gen_evaluado", False):
                if st.button("Evaluar respuesta"):
                    if resp_usuario:
                        st.session_state.gen_evaluado = True
                        st.session_state.gen_letra_marcada = resp_usuario[0].upper()
                        st.rerun()
                    else:
                        st.warning("Selecciona una alternativa.")

            if st.session_state.get("gen_evaluado", False):
                letra_real = datos['correcta'][0].upper()
                if st.session_state.gen_letra_marcada == letra_real:
                    st.success(f"¡Correcto! La clave es la {letra_real}.")
                    st.info(f"💡 **Explicación:** {datos['explicacion']}\n\n📚 **Fuente:** {datos['fuente']}")
                else:
                    st.error(f"Incorrecto. Marcaste la {st.session_state.gen_letra_marcada}.")
                    if not st.session_state.gen_mostrar_resp:
                        if st.button("👁️ Mostrar respuesta y explicación"):
                            st.session_state.gen_mostrar_resp = True
                            st.rerun()
                    else:
                        st.success(f"La respuesta correcta era la {letra_real}.")
                        st.info(f"💡 **Explicación:** {datos['explicacion']}\n\n📚 **Fuente:** {datos['fuente']}")

    elif opcion_menu == "Simulacro Médico":
        st.header("🧠 Simulador de Alta Velocidad (Multimotor RAG)")

        fuente_docs = st.radio("Selecciona origen de los Bancos:", ["Material Global (Admin)", "Mis Documentos Personales"], horizontal=True)
        carpeta_bancos = carpetas_admin[0] if fuente_docs == "Material Global (Admin)" else carpetas_usuario[0]
        archivos_disponibles = [f for f in os.listdir(carpeta_bancos) if not f.endswith(".json")]

        if not archivos_disponibles:
            st.warning("⚠️ No hay bancos de preguntas disponibles. Sube tus documentos primero.")
        else:
            doc_seleccionado = st.selectbox("Elige el documento:", archivos_disponibles)
            ruta_completa = os.path.join(carpeta_bancos, doc_seleccionado)

            if 'banco_cargado' not in st.session_state or st.session_state.get('doc_activo') != doc_seleccionado:
                st.session_state.banco_cargado = False
                st.session_state.examen_terminado = False

            if not st.session_state.banco_cargado:
                if st.button("🚀 Iniciar Simulacro de inmediato", type="primary"):
                    ruta_json = ruta_completa + ".json"
                    preguntas_extraidas = []

                    if os.path.exists(ruta_json):
                        with open(ruta_json, "r", encoding="utf-8") as f:
                            preguntas_extraidas = json.load(f)
                    else:
                        with st.spinner("Procesando banco con IA (leyendo completo)..."):
                            preguntas_extraidas = extraer_preguntas_local_pdf(ruta_completa)
                            with open(ruta_json, "w", encoding="utf-8") as f:
                                json.dump(preguntas_extraidas, f, ensure_ascii=False, indent=4)

                    if len(preguntas_extraidas) > 0:
                        st.session_state.lista_preguntas = preguntas_extraidas
                        random.shuffle(st.session_state.lista_preguntas)
                        st.session_state.banco_cargado = True
                        st.session_state.doc_activo = doc_seleccionado
                        st.session_state.indice_actual = 0
                        st.session_state.evaluado_rapido = False
                        st.session_state.mostrar_respuesta = False
                        st.session_state.correctas = 0
                        st.session_state.incorrectas = 0
                        st.session_state.errores_especialidad = {}
                        st.session_state.examen_terminado = False
                        st.rerun()
                    else:
                        st.error("No se detectaron preguntas estructuradas.")

            if st.session_state.get('examen_terminado', False):
                st.divider()
                st.subheader("📊 Resultados de tu Simulacro")
                total_respondidas = st.session_state.correctas + st.session_state.incorrectas

                if total_respondidas == 0:
                    st.info("No respondiste ninguna pregunta.")
                else:
                    col_grafico, col_datos = st.columns(2)
                    with col_grafico:
                        fig, ax = plt.subplots(figsize=(4, 4))
                        fig.patch.set_facecolor('none')
                        ax.pie([st.session_state.correctas, st.session_state.incorrectas], labels=['Correctas', 'Incorrectas'], autopct='%1.1f%%', startangle=90, colors=['#27AE60', '#E74C3C'], textprops={'color':"w", 'weight':'bold'})
                        ax.axis('equal')
                        st.pyplot(fig)

                    with col_datos:
                        st.write(f"**Total respondidas:** {total_respondidas}")
                        st.write(f"✅ **Aciertos:** {st.session_state.correctas}")
                        st.write(f"❌ **Fallos:** {st.session_state.incorrectas}")

                        st.divider()
                        st.subheader("⚠️ Áreas a reforzar")
                        if st.session_state.incorrectas > 0:
                            errores = st.session_state.errores_especialidad
                            especialidades_ordenadas = sorted(errores.items(), key=lambda x: x[1], reverse=True)
                            peor_esp, fallos = especialidades_ordenadas[0]
                            st.error(f"Urge repasar: **{peor_esp}** ({fallos} errores)")
                            for esp, f in especialidades_ordenadas[1:]: st.write(f"- {esp}: {f} errores")

                st.divider()
                if st.button("🔄 Volver al simulacro", type="primary"):
                    st.session_state.examen_terminado = False
                    st.session_state.correctas = 0
                    st.session_state.incorrectas = 0
                    st.session_state.errores_especialidad = {}
                    random.shuffle(st.session_state.lista_preguntas)
                    st.session_state.indice_actual = 0
                    st.session_state.evaluado_rapido = False
                    st.rerun()

            elif st.session_state.get('banco_cargado', False):
                preguntas = st.session_state.lista_preguntas
                idx = st.session_state.indice_actual
                pregunta_actual = preguntas[idx]

                st.divider()
                col_titulo, col_contador = st.columns([3, 1])
                with col_titulo: st.subheader(f"Estudiando: {doc_seleccionado}")
                with col_contador: st.markdown(f"<h3 style='text-align: right; color: #F1C40F; border: 2px solid #F1C40F; padding: 5px; border-radius: 5px;'>Preg. {idx + 1} / {len(preguntas)}</h3>", unsafe_allow_html=True)

                st.markdown(f"<h5 style='color: #9B59B6; font-style: italic;'>🏷️ Especialidad: {pregunta_actual['especialidad']}</h5>", unsafe_allow_html=True)
                st.markdown(f"### {pregunta_actual['pregunta']}")

                respuesta_usuario = st.radio("Selecciona tu respuesta:", pregunta_actual['opciones'], index=None, key=f"radio_local_{idx}")

                if not st.session_state.evaluado_rapido:
                    if st.button("Evaluar respuesta"):
                        if respuesta_usuario:
                            letra_marcada = respuesta_usuario[0].upper()
                            clave_pdf = pregunta_actual.get("correcta", "N/A")

                            st.session_state.clave_dinamica = clave_pdf
                            st.session_state.explicacion_dinamica = ""
                            st.session_state.motor_dinamico = ""

                            st.session_state.evaluado_rapido = True
                            st.session_state.letra_marcada_rapida = letra_marcada
                            st.session_state.mostrar_respuesta = False

                            if st.session_state.letra_marcada_rapida == st.session_state.clave_dinamica:
                                st.session_state.correctas += 1
                            else:
                                st.session_state.incorrectas += 1
                                esp_actual = pregunta_actual['especialidad']
                                st.session_state.errores_especialidad[esp_actual] = st.session_state.errores_especialidad.get(esp_actual, 0) + 1
                            st.rerun()
                        else:
                            st.warning("Selecciona una alternativa.")

                if st.session_state.evaluado_rapido:
                    letra_correcta = st.session_state.get("clave_dinamica", "N/A")

                    if st.session_state.letra_marcada_rapida == letra_correcta:
                        st.success(f"¡Correcto! La clave es la {letra_correcta}.")
                    else:
                        st.error(f"Incorrecto. Marcaste la {st.session_state.letra_marcada_rapida}.")

                        if not st.session_state.mostrar_respuesta:
                            if st.button("👁️ Mostrar respuesta y justificación teórica"):
                                with st.spinner("🧠 Buscando en tu teoría y analizando..."):
                                    try:
                                        datos_ai, motor_usado, frags = explicar_con_teoria(
                                            pregunta_actual,
                                            st.session_state.letra_marcada_rapida,
                                            letra_correcta,
                                            [carpetas_admin[1], carpetas_usuario[1]],
                                        )
                                        if letra_correcta == "N/A":
                                            st.session_state.clave_dinamica = datos_ai["correcta"][0].upper()

                                        usados = ", ".join(sorted({f"{c['archivo']} p.{c['pagina']}" for c in frags})) or "ninguno"
                                        st.session_state.explicacion_dinamica = (
                                            f"🧠 {datos_ai['explicacion']}\n\n"
                                            f"📚 **Fuente citada:** {datos_ai.get('fuente_cita', '—')}\n\n"
                                            f"🔎 **Fragmentos consultados:** {usados}"
                                        )
                                        st.session_state.motor_dinamico = motor_usado

                                    except Exception as e:
                                        st.session_state.explicacion_dinamica = f"Error: {str(e)}"

                                st.session_state.mostrar_respuesta = True
                                st.rerun()
                        else:
                            letra_final = st.session_state.get("clave_dinamica", "N/A")
                            if letra_final != "N/A":
                                st.success(f"La respuesta correcta es la {letra_final}.")
                            st.info(f"💡 **Feedback Teórico:**\n\n{st.session_state.get('explicacion_dinamica', '')}")
                            st.caption(f"⚡ Procesado por: {st.session_state.get('motor_dinamico', 'AI')}")

                    col_vacia, col_entregar, col_siguiente = st.columns([2, 1, 1])
                    with col_entregar:
                        if st.button("📊 Entregar examen"):
                            st.session_state.examen_terminado = True
                            st.rerun()
                    with col_siguiente:
                        if st.button("Siguiente pregunta ➡️", type="primary"):
                            if st.session_state.indice_actual < len(preguntas) - 1:
                                st.session_state.indice_actual += 1
                            else:
                                st.session_state.indice_actual = 0
                                random.shuffle(st.session_state.lista_preguntas)
                            st.session_state.evaluado_rapido = False
                            st.session_state.mostrar_respuesta = False
                            st.rerun()

# ==========================================
# 5. CONTROLADOR
# ==========================================
if st.session_state.usuario_actual is None:
    login_register_page()
else:
    main_app()
