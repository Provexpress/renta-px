# Renta PX

Dashboard web estático para visualizar arriendos de equipos desde un Excel alojado en SharePoint.

## Publicar en GitHub Pages

1. Sube este repositorio a GitHub con el nombre `renta-px`.
2. En GitHub, entra a `Settings > Pages`.
3. En `Build and deployment`, selecciona `Deploy from a branch`.
4. Usa la rama principal y la carpeta `/root`.
5. La URL quedará parecida a `https://USUARIO.github.io/renta-px/`.

## Crear App Registration en Microsoft Entra ID

1. Entra a Microsoft Entra ID.
2. Ve a `App registrations > New registration`.
3. Nombre sugerido: `renta-px`.
4. En tipos de cuenta, usa la opción de tu organización.
5. Crea la aplicación y copia:
   - `Application (client) ID`
   - `Directory (tenant) ID`
6. En `Authentication`, agrega plataforma `Single-page application`.
7. Configura la Redirect URI:

```txt
https://USUARIO.github.io/renta-px/
```

No uses `client secret`: esta app es una SPA pública.

## Permisos Microsoft Graph

Configura estos permisos delegados:

- `User.Read`
- `Files.Read.All`
- `Sites.Read.All`

Según la política del tenant, un administrador puede tener que aprobar consentimiento.

## Ubicar el Excel en SharePoint

Necesitas estos datos:

- Hostname, por ejemplo `empresa.sharepoint.com`
- Sitio, por ejemplo `/sites/NOMBRE_DEL_SITIO`
- Biblioteca, por ejemplo `Documents` o `Documentos`
- Ruta del archivo, por ejemplo `/Ruta/Archivo/Base.xlsx`
- Hoja del Excel, por ejemplo `Hoja1`

La app lee `sheetName` si existe; si no coincide, usa la primera hoja del archivo.

## Configurar `js/config.js`

Copia `js/config.example.js` sobre `js/config.js` o edita `js/config.js` directamente:

```js
const APP_CONFIG = {
  msal: {
    clientId: "CLIENT_ID_REAL",
    tenantId: "TENANT_ID_REAL",
    redirectUri: "https://USUARIO.github.io/renta-px/"
  },
  graph: {
    scopes: ["User.Read", "Files.Read.All", "Sites.Read.All"],
    sharePointFile: {
      siteHostname: "empresa.sharepoint.com",
      sitePath: "/sites/NOMBRE_DEL_SITIO",
      driveName: "Documents",
      driveNameAliases: ["Documentos", "Documentos compartidos"],
      filePath: "/Ruta/Archivo/Base.xlsx",
      sheetName: "Hoja1"
    }
  }
};
```

## Mapear usuarios y roles

Edita `js/permissions.js` y agrega los correos reales en `USER_ACCESS`.

Roles soportados:

- `gerencia`: ve consolidado completo.
- `finanzas`: ve rentabilidad global.
- `operaciones`: ve inventario global y puede ocultar columnas financieras.
- `comercial`: ve solo las filas donde `row.comercial` coincide con su comercial asignado.

Para comerciales puedes usar `comercialAliases`. Esto ayuda cuando el correo tiene el nombre completo, pero en el Excel aparece un nombre corto como `OSCAR BELTRAN` o `ASTRID JIMENEZ`.

Para probar roles, cambia temporalmente el correo de tu usuario Microsoft en `USER_ACCESS`.

## Limitación de seguridad

Esta versión es 100% estática. El filtro por comercial ocurre en el navegador, después de descargar el Excel. Para seguridad fuerte, el Excel completo no debería enviarse al navegador de comerciales.

En una fase futura conviene usar backend, Azure Function, Power Automate/API, o separar archivos y permisos por comercial en SharePoint.

## Validación esperada con la base de ejemplo

La lectura del Excel debe acercarse a estos valores, sin quemarlos en el código:

- 850 equipos
- 804 portátiles
- 46 PC
- Canon mensual aproximado: 130.259.406 COP
- Costo mensual aproximado: 98.819.377 COP
- Utilidad mensual aproximada: 31.440.027 COP
- Margen aproximado: 24,1%
- 31 clientes
- 18 comerciales

La consola del navegador imprime un resumen de validación cuando los datos cargan.
