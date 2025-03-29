import pandas as pd
import streamlit as st
import time
import random
import itertools
import functools
#import ssl  # Para conexión segura con el servidor SMTP
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO
#from email.message import EmailMessage

# Lee el archivo Excel
file_path = 'Base preguntas.xlsx'

# Usa un diccionario para almacenar cada hoja en un dataframe diferente
sheets_dict = pd.read_excel(file_path, sheet_name=None)

sheets_dict2 = pd.DataFrame(sheets_dict.items())
for i in range(len(sheets_dict2[1])):
    sheets_dict2[1][i] = sheets_dict2[1][i].dropna(axis=1, how="all").dropna()

for i in range(len(sheets_dict2[1])):
    sheets_dict2[1][i].columns = ['preguntas']

def construir_diccionarios(indice): # Indice de la sección
    ejemplo_pregunta = sheets_dict2[1][indice]
    ejemplo_pregunta = ejemplo_pregunta.reset_index(drop=True)
    ejemplo_pregunta['Grupo'] = ejemplo_pregunta.index // 8
    ejemplo_pregunta = ejemplo_pregunta[ejemplo_pregunta['Grupo'].isin((ejemplo_pregunta['Grupo'].value_counts() == 8)\
        .reset_index()[(ejemplo_pregunta['Grupo'].value_counts() == 8).reset_index()['count'] == True]['Grupo'].tolist())]

    dict_preg_resp = []
    dict_preg_resp_correct_expli = []
    for i in ejemplo_pregunta['Grupo'].unique():
        try:
            dict_preg_resp.append({ejemplo_pregunta[ejemplo_pregunta['Grupo']==i].reset_index(drop=True)['preguntas'][0]:\
                                {item[0]: item[1] for item in ejemplo_pregunta[ejemplo_pregunta['Grupo']==i].reset_index(drop=True)['preguntas'][1:6]\
                                    .apply(lambda x: x.split(') '))}})
        except:
            dict_preg_resp.append({ejemplo_pregunta[ejemplo_pregunta['Grupo']==i].reset_index(drop=True)['preguntas'][0]:\
                            {item[0]: item[1] for item in ejemplo_pregunta[ejemplo_pregunta['Grupo']==i].reset_index(drop=True)['preguntas'][1:6]\
                                .apply(lambda x: x.split('. '))}})
        dict_preg_resp_correct_expli.append({ejemplo_pregunta[ejemplo_pregunta['Grupo']==i].reset_index(drop=True)['preguntas'][0]: {ejemplo_pregunta\
        [ejemplo_pregunta['Grupo']==i].reset_index(drop=True)['preguntas'][6]: ejemplo_pregunta[ejemplo_pregunta['Grupo']==i].reset_index(drop=True)['preguntas'][7]}})
    resultado = {}
    for j in range(len(dict_preg_resp)):
        for pregunta, opciones in dict_preg_resp[j].items():
            if pregunta in dict_preg_resp_correct_expli[j]:
                respuesta_y_explicacion = list(dict_preg_resp_correct_expli[j][pregunta].items())[0]
                respuesta_correcta = respuesta_y_explicacion[0].split(': ')[1]  # Extrae el texto de la respuesta
                explicacion = respuesta_y_explicacion[1]
                # Crear la estructura combinada
                resultado[pregunta] = {
                    'Opciones': opciones,
                    'Respuesta correcta': respuesta_correcta,
                    'Explicación': explicacion
                }
    return resultado
sheets_dict2['diccionarios'] = ""
resultado = []
for i in range(len(sheets_dict2[1])):
    resultado.append(construir_diccionarios(i))

for i in range(len(sheets_dict2)):
    sheets_dict2['diccionarios'][i] = resultado[i]
sheets_dict2['nivel'] = ""
sheets_dict2.loc[sheets_dict2[0].isin(['AC Witzel 02', 'AC Witzel 03', 'AC Witzel 04',
                                       'George 01','George 02', 'George 03',
                                       'George 04']), 'nivel'] = 'Nivel 1: Los orígenes de la administración'
sheets_dict2.loc[sheets_dict2[0].isin(['George 05 Ad Ci', 'George 06 Ad Ci',
                                       'Ad Ci Witzel 05',
                                       'Ad Ci Davila']), 'nivel'] = 'Nivel 2: Administración científica'
sheets_dict2.loc[sheets_dict2[0].isin(['Fay Davila', 'Fayol Chiav',
                                       'Fayol Aktouf']), 'nivel'] = 'Nivel 3: Concepción clásica de la administración'
sheets_dict2.loc[sheets_dict2[0].isin(['RH Witzel 07', 'RH Davila',
                                       'RH Chiav']), 'nivel'] = 'Nivel 4: Relaciones Humanas'
sheets_dict2.loc[sheets_dict2[0].isin(['Weber Davila',
                                       'Weber Chiav']), 'nivel'] = 'Nivel 5: La Burocracia de Max Weber'
sheets_dict2['seccion'] = sheets_dict2['nivel'] + '\n\n' + sheets_dict2[0]
# Función para fusionar diccionarios
def merge_dicts(dict_list):
    return functools.reduce(lambda d1, d2: {**d1, **d2}, dict_list)
sheets_dict2 = sheets_dict2.groupby('nivel', as_index=False).agg({'diccionarios': lambda x: merge_dicts(x)})
sheets_dict2 = {row['nivel']: row["diccionarios"] for _, row in sheets_dict2.iterrows()}
sheets_dict2 = dict(itertools.islice(sheets_dict2.items(), 1)) #seleccionar las secciones

def enviar_resultados_por_correo():
    """Genera un archivo Excel con los resultados y lo envía al correo del profesor."""
    # Crear un DataFrame con las respuestas
    df = pd.DataFrame(st.session_state.respuestas)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Resultados")
    output.seek(0)  # Mover el puntero al inicio del archivo en memoria

    # Crear el correo
    email_from = "jpsaenzp@ut.edu.co"
    #with open("pswd.txt", "r") as archivo:
    #    pswd = archivo.readline()
    pswd = st.secrets["pswd"]
    email_to = "jpsaenzp@ut.edu.co"
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    msg = MIMEMultipart()
    msg['From'] = email_from
    msg['To'] = email_to
    msg['Subject'] = f"Resultados del Quiz - {st.session_state.nombre_usuario}"
    msg.attach(MIMEText(f"Hola, adjunto los resultados del quiz de {st.session_state.nombre_usuario}.", 'plain'))

    # Adjuntar archivo Excel desde memoria
    attachment = MIMEBase('application', 'octet-stream')
    attachment.set_payload(output.read())
    encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', 'attachment', filename="resultados_quiz.xlsx")
    msg.attach(attachment)

    # Enviar el correo
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(email_from, pswd)
    server.sendmail(email_from, email_to, msg.as_string())
    server.quit()

    # Confirmar envío en Streamlit
    st.success("📧 ¡Resultados enviados con éxito al profesor!")

# Solicitar nombre al inicio
if 'nombre_usuario' not in st.session_state:
    st.session_state.nombre_usuario = ""

st.session_state.nombre_usuario = st.text_input("Ingrese su nombre:", st.session_state.nombre_usuario)
if not st.session_state.nombre_usuario:
    st.stop()

# Seleccionar 5 preguntas por sección
if 'preguntas_mostradas' not in st.session_state:
    st.session_state.preguntas_mostradas = {
        seccion: random.sample(list(preguntas.keys()), 5) ##¿cuántas preguntas se requieren por seccion?
        for seccion, preguntas in sheets_dict2.items()
    }

# Variables de sesión
if 'seccion_actual' not in st.session_state:
    st.session_state.seccion_actual = list(sheets_dict2.keys())[0]
if 'pregunta_actual' not in st.session_state:
    st.session_state.pregunta_actual = st.session_state.preguntas_mostradas[st.session_state.seccion_actual][0]
if 'respuestas' not in st.session_state:
    st.session_state.respuestas = []
if 'puntajes' not in st.session_state:
    st.session_state.puntajes = {seccion: 0 for seccion in sheets_dict2.keys()}  # Inicializar puntajes
if 'respuesta_validada' not in st.session_state:
    st.session_state.respuesta_validada = False
if 'quiz_finalizado' not in st.session_state:
    st.session_state.quiz_finalizado = False

if 'quiz_finalizado2' not in st.session_state:
    st.session_state.quiz_finalizado2 = False   

# Muestra la pantalla final si el quiz ha terminado
if st.session_state.quiz_finalizado:
    st.empty()
    st.header("🎉 ¡Quiz finalizado! 🎉")
    st.subheader("Gracias por participar, has completado todas las preguntas.")
    if st.button("Enviar Resultados al Profesor"):
        enviar_resultados_por_correo()

    # Botón de descarga de resultados
    if st.button("Descargar Resultados"):
        df = pd.DataFrame(st.session_state.respuestas)
        df.to_excel("resultados_quiz.xlsx", index=False)
        st.success("Resultados guardados como 'resultados_quiz.xlsx'")
    # Mostrar puntajes
    for seccion, puntaje in st.session_state.puntajes.items():
        st.write(f"**{seccion}:** {puntaje}/5 - Puntaje: {puntaje / 5:.2f}")

    st.stop()

# Obtener sección y pregunta actual
seccion_nombre = st.session_state.seccion_actual
preguntas = sheets_dict2[seccion_nombre]
pregunta_actual_key = st.session_state.pregunta_actual
pregunta_actual = preguntas[pregunta_actual_key]

st.header(f'Sección {seccion_nombre}')
st.subheader(f'Pregunta {pregunta_actual_key}')

opciones = pregunta_actual['Opciones']
opciones_formateadas = [f"{clave}) {valor}" for clave, valor in opciones.items()]
respuesta_seleccionada = st.radio(
    "Seleccione una opción:", 
    opciones_formateadas, 
    key=f"{seccion_nombre}_{pregunta_actual_key}", 
    disabled=st.session_state.respuesta_validada
)

if 'run_button' in st.session_state and st.session_state.run_button == True:
    st.session_state.running = True
else:
    st.session_state.running = False

if st.session_state.quiz_finalizado2:
    st.empty()
    st.header("🎉 ¡Quiz finalizado! 🎉")
    st.subheader(f"No puedes continuar, tu puntaje en {seccion_nombre} fue {st.session_state.puntajes[seccion_nombre]}/5.")
    if st.button("Enviar Resultados al Profesor"):
        enviar_resultados_por_correo()
    
    # Botón de descarga de resultados
    if st.button("Descargar Resultados"):
        df = pd.DataFrame(st.session_state.respuestas)
        df.to_excel("resultados_quiz.xlsx", index=False)
        st.success("Resultados guardados como 'resultados_quiz.xlsx'")
    st.stop()

if respuesta_seleccionada:
    st.session_state.respuesta_validada = True
    respuesta_clave, respuesta_valor = respuesta_seleccionada.split(') ', 1)
    respuesta_usuario = f"{respuesta_clave}) {respuesta_valor}".rstrip('.')

    if st.button("Validar Respuesta",disabled=st.session_state.running, key='run_button'):
        correcta = pregunta_actual['Respuesta correcta'].rstrip('.')
        if respuesta_usuario == correcta:
            st.success("✅ Respuesta correcta!")
            st.session_state.puntajes[seccion_nombre] += 1  # Sumar puntaje
        else:
            st.error("❌ Respuesta incorrecta")
            st.warning(f"Respuesta correcta: {correcta}")

        st.warning(pregunta_actual['Explicación'])
        time.sleep(3)
        st.session_state.respuesta_validada = False
        st.session_state.respuestas.append({
            'Nombre': st.session_state.nombre_usuario,
            'Sección': seccion_nombre,
            'Pregunta': pregunta_actual_key,
            'Respuesta Usuario': respuesta_usuario[0],
            'Correcta': correcta,
            'Resultado': 'Correcto' if respuesta_usuario[0] == correcta[0] else 'Incorrecto'
        })

        # Lógica de navegación
        index_actual = st.session_state.preguntas_mostradas[seccion_nombre].index(pregunta_actual_key)
        if index_actual < len(st.session_state.preguntas_mostradas[seccion_nombre]) - 1:  # Si aún quedan preguntas en la sección
            st.session_state.pregunta_actual = st.session_state.preguntas_mostradas[seccion_nombre][index_actual + 1]
        else:  # Si finalizó la sección, validar si puede continuar
            if st.session_state.puntajes[seccion_nombre] >= 3:
                seccion_keys = list(sheets_dict2.keys())
                index_seccion = seccion_keys.index(seccion_nombre)
                if index_seccion < len(seccion_keys) - 1:
                    st.session_state.seccion_actual = seccion_keys[index_seccion + 1]
                    st.session_state.pregunta_actual = st.session_state.preguntas_mostradas[st.session_state.seccion_actual][0]
                else:
                    st.session_state.quiz_finalizado = True
            else:
                st.session_state.quiz_finalizado2 = True
        st.rerun()