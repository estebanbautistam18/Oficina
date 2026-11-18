# Naciones | Oficina

Aplicación web adaptable a computadora y celular, construida con Google Apps Script + Google Sheets. El código se versiona en GitHub y Apps Script sirve la interfaz y la API.

## Qué incluye esta primera versión

- **Inicio / Top 3:** tres objetivos semanales con detalle, guardados por semana.
- **Recordatorio semanal:** correo programado para los sábados alrededor de las 12:00 p. m. para preparar el Top 3 siguiente.
- **11 Sur:** tareas y mensajes por enviar, inventario, organigrama y retroalimentación semanal.
- **Retro 11 Sur:** guarda el registro y envía un correo con asunto `Retro 11 sur - dd/mmm/yyyy` y las secciones solicitadas.
- **Hombría al Máximo:** checklist semanal generado automáticamente (Puebla, Huejotzingo, agenda y vehículo), calendarios por campus y asistencia leída desde las celdas especificadas.
- **Checklist general:** consulta y actualiza los mismos registros de tareas que aparecen en los módulos. No hay copias por sección.
- **Eventos / Otras áreas:** permite administrar tareas de áreas adicionales.
- **Insumos y Equipos:** módulos separados para inventario y organigrama de personas.

## Arquitectura

- **Google Sheets:** fuente única de datos. El script independiente agrega siete pestañas `Oficina - ...` a la hoja de Hombría existente.
- **Google Apps Script:** API, reglas de negocio, envío de correo y programación del recordatorio.
- **GitHub:** historial y control de versiones del proyecto.
- No se utiliza `localStorage` como base de datos, por lo que los cambios se comparten entre los dispositivos que abran la misma implementación.

## Instalación sin modificar el Apps Script existente

1. Abre [Google Apps Script](https://script.google.com/) con la misma cuenta que tiene acceso de edición a la hoja de Hombría.
2. Crea un **proyecto independiente** llamado `Naciones | Oficina`. No lo crees desde Extensiones dentro de la hoja; así no se modifica ni se reemplaza el código que ya utiliza Hombría.
3. En ese proyecto, reemplaza/crea los archivos con el contenido de este repositorio:
   - `Code.gs`
   - `Index.html`
   - `Styles.html`
   - `App.html`
4. En **Configuración del proyecto**, activa la opción para mostrar el archivo de manifiesto `appsscript.json` y reemplázalo con el incluido aquí.
5. Guarda el proyecto y ejecuta una vez `setupOfficePlatform` y `authorizeOfficePlatform` desde el editor. Acepta los permisos de Google Sheets, correo y activadores.
6. En el editor, abre **Implementar → Nueva implementación → Aplicación web**.
   - Ejecutar como: **tu cuenta**.
   - Quién tiene acceso: **solo tú** para el primer lanzamiento. Más adelante se puede abrir a usuarios de tu organización.
7. Abre la URL de la aplicación web.
8. En **Configuración**, verifica tu correo y guarda para instalar el recordatorio del sábado.

> El ID de la hoja ya está definido en `Code.gs`. El script independiente abre el documento con `SpreadsheetApp.openById()`, así no depende de un proyecto Apps Script vinculado ni toca el código que ya tengas.

## Integración de Hombría

La integración usa el ID de la hoja compartida y los nombres de pestaña y rangos que indicaste. El script necesita que la cuenta que lo ejecuta tenga permiso de edición sobre ese documento:

- `Calendario Oradores / Host`: A:I para Puebla y K:S para Huejotzingo.
- `Asistencias`: B11:K11 / B12:K12 para fechas y asistencias de Puebla; B27:K27 / B28:K28 para Huejotzingo.

La tabla de oradores se lee desde la fila 2 hasta la última fila disponible. Las fechas se muestran tal como aparecen en la hoja. El total de inscritos todavía se configura manualmente en **Configuración**, porque no se proporcionó la celda o rango donde vive ese dato.

## Modelo de datos

El script crea estas pestañas adicionales, sin alterar las dos pestañas fuente de Hombría:

- `Oficina - Tareas`
- `Oficina - Top 3`
- `Oficina - Retros 11 Sur`
- `Oficina - Insumos`
- `Oficina - Equipos`
- `Oficina - Áreas`
- `Oficina - Configuración`

Cada tarea se almacena una sola vez con `area`, `categoria`, `campus`, `semana` y `completada`. Al marcarla, editarla o eliminarla desde cualquier módulo, el cambio se verá en todos los paneles después de actualizar los datos.

## Recordatorio de los sábados

Se crea un activador semanal de Apps Script para el sábado a las 12:00 p. m. con zona horaria `America/Mexico_City`. Google ejecuta los activadores de tiempo cerca de la hora solicitada, no como un reloj exacto al segundo. La aplicación programa el envío a tu correo configurado.

## Notas antes de producción

- La primera versión está pensada para un administrador principal. Si después entrarán coordinadores o equipos, habrá que diseñar permisos por rol y campus antes de compartir la aplicación.
- Las tareas, Top 3, insumos, equipo y retros se guardan en la hoja vinculada. Evita renombrar manualmente las pestañas `Oficina - ...`.
- Los datos históricos de asistencia y el calendario original se leen; no se editan desde la aplicación.
- El correo de la retro se envía desde la cuenta propietaria del script después de autorizar `MailApp`.


## Versionado con GitHub (opcional con clasp)

Puedes mantener el proyecto en un repositorio independiente y sincronizarlo con Apps Script:

```bash
npm install -g @google/clasp
clasp login
clasp create --type standalone --title "Naciones | Oficina"
clasp push
clasp deploy --description "Naciones Oficina v1"
```

`clasp create` genera el archivo local `.clasp.json` con el ID privado del proyecto. No lo publiques. Después de crear el proyecto, ejecuta `clasp push`, abre el proyecto con `clasp open` y ejecuta desde el editor `setupOfficePlatform` y `authorizeOfficePlatform`. Luego implementa como aplicación web siguiendo los pasos anteriores. Para cada actualización del repositorio, ejecuta `clasp push` y publica una nueva versión desde **Implementar → Gestionar implementaciones**.
