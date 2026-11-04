# Control Bus Prize – Conexión con OneDrive

```
Excel en OneDrive  ⇄  Microsoft Graph  ⇄  Servidor en Vercel (/api)  ⇄  App en el celular
```

- **USUARIOS, CAPACIDADES y ASIGNACIONES** se leen de `BASE_MAESTRA_ASIGNACIONES.xlsx`. Un cambio en el Excel llega a los celulares en **1 minuto o menos**.
- **Los registros** se escriben en la tabla REGISTROS de `REGISTROS_EMBARQUE.xlsx` apenas se validan.
- **Las contraseñas** nunca salen del servidor.

## 1. Registrar la app en Microsoft (lo hace TI / administrador de Microsoft 365)
1. Ir a **portal.azure.com → Microsoft Entra ID → Registros de aplicaciones → Nuevo registro**.
   - **Nombre:** `Control Bus Prize`
   - **Tipo de cuenta:** "Solo este directorio organizativo".
   - Clic en **Registrar**.
2. Copiar el **Id. de aplicación (cliente)** y el **Id. de directorio (inquilino)**.
3. En **Certificados y secretos → Nuevo secreto de cliente** (vence en 24 meses), copiar el **Valor**.
4. Agregar el permiso en **Permisos de API → Agregar un permiso → Microsoft Graph → Permisos de aplicación → `Files.ReadWrite.All` → Agregar**.
5. Clic en **Conceder consentimiento de administrador**.

## 2. Variables en Vercel
Entrar a **Proyecto → Settings → Environment Variables** y agregar:

| Variable | Valor |
|---|---|
| `TENANT_ID` | Id. de directorio (inquilino) |
| `CLIENT_ID` | Id. de aplicación (cliente) |
| `CLIENT_SECRET` | Valor del secreto |
| `DRIVE_USER` | correo del dueño del OneDrive, ej. `jcontreras@aquanqa.pe` (debe ser la cuenta real de Microsoft 365; si no, Graph responde "User not found") |
| `BASE_PATH` | `Gi Cosecha y Operaciones/App_Sheet/Registros Control_Bus` |
| `BASE_FILE` | `BASE_MAESTRA_ASIGNACIONES.xlsx` |
| `REG_FILE` | `REGISTROS_EMBARQUE.xlsx` |
| `APP_SECRET` | una clave larga inventada (40+ caracteres) |

Después: **Deployments → … → Redeploy**.

> **Si la carpeta "Gi Cosecha y Operaciones" es un acceso directo a una carpeta compartida**, `DRIVE_USER` no la encuentra. En ese caso se debe usar el sitio de SharePoint original:
> - Poner en `DRIVE_ID` el id de esa biblioteca (TI lo obtiene en Graph Explorer) y borrar `DRIVE_USER`.
> - Ajustar `BASE_PATH` a la ruta dentro de esa biblioteca.

## 3. Probar
- Abrir `https://TU-APP.vercel.app/api/login?placa=PLACA` con una placa real de USUARIOS. Debe responder su rol, ej. `{"rol":"CONDUCTOR"}`.
- Si sale un error, el mensaje dice qué falta: permisos, ruta o nombre de tabla.

## Reglas del Excel
- **Nombres de tablas y hojas:** no cambiar los nombres de las tablas de Excel (USUARIOS, CAPACIDADES, ASIGNACIONES, REGISTROS) ni de las columnas.
- **Filas nuevas:** agregarlas dentro de la tabla, pegando justo debajo de la última fila para que la tabla se expanda sola.
- **ID REGISTRO:** la columna `ID REGISTRO` de REGISTROS la llena la app. Evita duplicados; no editarla.
