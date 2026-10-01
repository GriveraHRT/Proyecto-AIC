# Directrices de Desarrollo - Proyecto AIC (LabControl)

## Flujo de Despliegue para Google Apps Script (Backend)

Al modificar archivos de backend (`codigo.gs` o `appscript/Código.js`):

1. **Sincronización de Archivos Locales**:
   - Mantener siempre sincronizados tanto `codigo.gs` (raíz) como `appscript/Código.js`.

2. **Ejecución de Clasp en Windows**:
   - Ejecutar los comandos de clasp usando `cmd /c clasp` desde la carpeta `appscript/` para evitar restricciones de PowerShell (`PSSecurityException`).

3. **Protocolo Obligatorio de Despliegue (Web App)**:
   - `clasp push` solo actualiza el código en `@HEAD`.
   - **SIEMPRE** crear y vincular una nueva versión al despliegue activo:
     1. Crear versión: `cmd /c clasp deploy`
     2. Identificar el Deployment ID activo de la API (ver `API_URL` en `app.js`).
     3. Redesplegar la versión activa:
        `cmd /c clasp deploy -i <deploymentId> -V <versionNumber> -d "<descripcion>"`

4. **Verificación de Endpoint**:
   - Comprobar la respuesta del endpoint desplegado con Node.js (`node -e "fetch('...').then(...)"`) verificando que reconozca los nuevos campos y hojas antes de dar la tarea por finalizada.
