# SIM → Word | Convertidor de Metadatos INEGI

App web para convertir los JSON exportados del Sistema Integrador de Metadatos (SIM)
a documentos Word con formato institucional.

---

## Estructura del repositorio

```
sim_app/
├── app.py              # App Streamlit (interfaz web)
├── converter.py        # Lógica de conversión JSON → Word
├── requirements.txt    # Dependencias Python
├── .streamlit/
│   └── config.toml     # Tema visual institucional
└── README.md
```

---

## Cómo publicar en Streamlit Community Cloud (paso a paso)

### 1. Subir el código a GitHub

1. Ve a [github.com](https://github.com) e inicia sesión.
2. Crea un repositorio nuevo (puede ser **privado**). Nómbralo por ejemplo `sim-metadatos-word`.
3. Sube todos los archivos de esta carpeta al repositorio
   (puedes arrastrarlos desde la interfaz web de GitHub).

### 2. Crear cuenta en Streamlit Cloud

1. Ve a [share.streamlit.io](https://share.streamlit.io).
2. Haz clic en **"Sign up"** e inicia sesión con tu cuenta de **GitHub**.

### 3. Publicar la app

1. En Streamlit Cloud haz clic en **"New app"**.
2. Selecciona tu repositorio (`sim-metadatos-word`).
3. En **"Main file path"** escribe: `app.py`
4. Haz clic en **"Deploy"**. En 1-2 minutos tendrás tu URL pública.

### 4. Restringir el acceso (solo personas autorizadas)

1. En Streamlit Cloud, entra a la configuración de tu app (ícono ⚙️).
2. Ve a la sección **"Sharing"**.
3. Cambia de "Public" a **"Private"**.
4. En el campo **"Invite viewers by email"**, agrega los correos de cada
   persona que debe tener acceso (uno por línea).
5. Guarda los cambios.

Cada persona recibirá un correo de invitación. Para acceder deberá
iniciar sesión con Google o GitHub usando ese mismo correo.

---

## Uso de la app

1. Entra a la URL de la app.
2. Inicia sesión con tu correo autorizado.
3. Sube el archivo `.json` exportado del SIM.
4. La app muestra el nombre del Proceso de Producción identificado.
5. Haz clic en **"Descargar documento Word"**.

---

## Actualizar el convertidor

Si en el futuro necesitas ajustar el formato del Word, solo modifica
`converter.py` y sube el cambio a GitHub. Streamlit Cloud detecta el
cambio y actualiza la app automáticamente en segundos.

---

## Dependencias

| Paquete | Versión mínima | Para qué sirve |
|---|---|---|
| streamlit | 1.35.0 | Interfaz web |
| python-docx | 1.1.0 | Generar el Word |
