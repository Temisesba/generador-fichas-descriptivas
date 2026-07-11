# Generador de Fichas Descriptivas

App para llenar fichas descriptivas individuales (por campo formativo) de un grupo escolar, usando Google Sheets como base de datos.

## Uso

1. Abre `GeneradorFichas.html` (o la versión publicada en GitHub Pages).
2. Ve a la pestaña **Configuración** y pega la URL de tu Apps Script (termina en `/exec`). Se guarda una sola vez en tu navegador.
3. Llena fichas en la pestaña **Fichas**, revisa el avance en **Trabajadas**, y edita el catálogo de frases en **Catálogo de Frases**.

## Backend (Google Sheets + Apps Script)

`Code.gs` es el backend. Instrucciones de instalación completas al inicio de ese archivo:

1. Crea un Google Sheet (o usa uno vacío).
2. Extensiones → Apps Script → pega `Code.gs` → Guardar.
3. Corre la función `setup()` una vez (crea las pestañas Config, Alumnos, Frases, Fichas con datos de ejemplo).
4. Implementar → Web app → acceso "Cualquier usuario" → copia la URL `/exec`.
5. Pega esa URL en la app (paso 2 de arriba).

La URL de Apps Script **no** va escrita en el código a propósito: es la llave de acceso a tus datos, y este repositorio es público. Cada persona que use la app la pega una sola vez en su propio navegador.
