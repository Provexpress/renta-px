# Acta de Proyecto - Renta PX

## 1. Informacion general

**Proyecto:** Renta PX  
**Nombre visible:** Renta PX - Dashboard de Arriendos  
**Fecha del acta:** 11 de mayo de 2026  
**Organizacion:** Provexpress SAS  
**Repositorio:** https://github.com/Provexpress/renta-px  
**Sitio publicado:** https://provexpress.github.io/renta-px/  
**Tipo de solucion:** Aplicacion web estatica publicada en GitHub Pages

## 2. Objetivo

Implementar una solucion web basica para visualizar, filtrar y exportar informacion de arriendos y subrentas de equipos, tomando como fuente un archivo Excel alojado en SharePoint.

El objetivo principal es entregar a Gerencia y al equipo comercial una vista consolidada, controlada por roles, que permita consultar indicadores, clientes, comerciales, equipos, alertas de calidad de dato y exportaciones para seguimiento operativo y comercial.

## 3. Alcance implementado

Se desarrollo una aplicacion web estatica con autenticacion Microsoft, lectura de Excel desde SharePoint mediante Microsoft Graph, procesamiento de datos en navegador, dashboard gerencial y visualizacion por roles.

Funcionalidades entregadas:

- Login con cuenta corporativa Microsoft.
- Obtencion del perfil del usuario autenticado.
- Consulta del archivo Excel en SharePoint mediante Microsoft Graph.
- Lectura de la hoja principal `PROVEXPRESS`.
- Lectura de la hoja de subrenta `PC COM`.
- Normalizacion de nombres de columnas.
- Calculo de indicadores operativos y financieros.
- Filtro de datos por rol.
- Vista global para Gerencia.
- Vista individual para comerciales.
- Vista de Subrenta para Gerencia y comerciales autorizados.
- Filtro gerencial por comercial.
- Graficas con Chart.js.
- Tabla con busqueda, filtros, ordenamiento y paginacion.
- Exportacion de tabla a Excel con formato visual corporativo.
- Panel de calidad del dato.
- Favicon y estilo visual alineado con Provex One.

## 4. Fuente de datos

La aplicacion consume un archivo Excel en SharePoint:

**Sitio:** `https://provexpress.sharepoint.com/sites/ProvexpressIntranet/servicios`  
**Biblioteca:** `Documentos` / `Documentos compartidos`  
**Ruta:** `Coordinador de servicios/Renta/RENTA_2024_T_1.xlsx`

Hojas configuradas:

- `PROVEXPRESS`: base principal de Renta.
- `PC COM`: base de Subrenta.

## 5. Tecnologia utilizada

La solucion se construyo bajo las restricciones definidas para una aplicacion estatica:

- HTML.
- CSS.
- JavaScript vanilla.
- MSAL.js por CDN.
- Microsoft Graph API.
- SheetJS / xlsx-js-style por CDN.
- Chart.js por CDN.
- GitHub Pages.

No se utilizo:

- Backend.
- Base de datos.
- React.
- Next.js.
- Vue.
- Angular.
- Framework CSS.
- Servidor propio.

## 6. Roles y permisos funcionales

### Gerencia

Puede consultar:

- Vista Renta.
- Vista Subrenta.
- Consolidado global.
- Filtro por comercial.
- KPIs financieros y operativos.
- Graficas globales.
- Tabla completa.
- Exportacion de datos filtrados.

Correos configurados:

- `oscar.perez@provexpress.com.co`
- `especialista.preventa@provexpress.com.co`

### Comercial

Puede consultar unicamente su cartera, filtrada por el campo `COMERCIAL` del Excel.

Comerciales configurados para Renta:

- Jasbleidy Mojica.
- Astrid Jimenez.
- Oscar Beltran.
- Lington Linares.
- Angela Torres.
- Yeison Urrego.
- Dilma Cuesta.
- Tatiana Parra.
- Johanna Jaime.
- Javier Cortes.
- Daniel Galindo.
- Julieth Galindo.
- Claudia Triana.
- Fernando Quinonez.
- Camilo Hernandez.
- Diana Castro.
- Juan Velasquez.
- Andres Pena.

Comerciales configurados para Subrenta:

- Astrid Jimenez.
- Angela Torres.
- Karent Carrillo.
- Dilma Cuesta.

### Finanzas

Rol previsto para consultar rentabilidad global.

### Operaciones

Rol previsto para consultar inventario global, con menor exposicion financiera.

## 7. Calidad del dato

La aplicacion calcula alertas por fila y clasifica el estado del registro:

- **Completo:** registro sin alertas relevantes.
- **Revisar:** falta placa, fecha de entrega o informacion de accesorios.
- **Error:** falta valor de arriendo, costo de renta o existe inconsistencia entre utilidad y la resta de valor menos costo.

Indicadores de calidad disponibles:

- Registros sin placa.
- Registros sin fecha de entrega.
- Registros sin valor de arriendo.
- Registros sin costo de renta.
- Registros con utilidad inconsistente.
- Registros con accesorios sin dato.

## 8. Exportacion

La tabla permite exportar los datos filtrados a Excel.

La exportacion:

- Respeta el rol del usuario.
- Respeta el filtro de Gerencia por comercial.
- Respeta busqueda y filtros de tabla.
- Exporta columnas de inventario y arriendo.
- Genera archivo `.xlsx` real.
- Aplica estilo visual corporativo.
- Usa nombre del cliente cuando el filtro contiene un solo cliente.

## 9. Decisiones tomadas

- Mantener la solucion 100% estatica para publicacion en GitHub Pages.
- Usar Microsoft Graph con permisos delegados.
- No usar client secret por tratarse de una SPA publica.
- Usar mapeo local temporal de usuarios y roles.
- Filtrar por comercial en frontend.
- Leer Subrenta desde la hoja `PC COM` del mismo archivo Excel.
- Mantener una sola fuente SharePoint para facilitar la operacion inicial.

## 10. Riesgos y limitaciones

Al ser una aplicacion estatica, el navegador descarga el archivo Excel completo antes de aplicar filtros por comercial. Esto significa que el filtro por rol es funcional y operativo, pero no representa una capa de seguridad fuerte.

Riesgos principales:

- Un usuario con acceso al archivo en SharePoint podria acceder a mas informacion de la que visualmente se le muestra.
- Los permisos reales dependen de SharePoint y Microsoft Graph.
- Si cambia la estructura del Excel, podria requerirse ajuste en normalizacion de columnas.
- Si cambian nombres de comerciales, deben actualizarse alias en `js/permissions.js`.

## 11. Recomendaciones

Para una fase futura se recomienda:

- Implementar backend o Azure Function para filtrar datos antes de enviarlos al navegador.
- Separar archivos por comercial o aplicar permisos mas granulares en SharePoint.
- Migrar roles a una fuente administrable, como SharePoint List, Dataverse o Azure Table.
- Agregar logs de acceso y auditoria.
- Crear un proceso de validacion de calidad de dato antes de publicar el Excel.
- Automatizar actualizacion y control de versiones del archivo fuente.

## 12. Estado del proyecto

El proyecto se encuentra publicado y funcional en GitHub Pages. La aplicacion ya permite autenticacion, lectura desde SharePoint, visualizacion por roles, vista de Renta, vista de Subrenta para usuarios autorizados, dashboard gerencial, tabla filtrable y exportacion a Excel.

## 13. Pendientes sugeridos

- Validar con cada comercial el acceso real al archivo en SharePoint.
- Confirmar que todos los nombres del campo `COMERCIAL` coincidan con los alias configurados.
- Revisar periodicamente el archivo `RENTA_2024_T_1.xlsx`.
- Definir si Finanzas y Operaciones requieren usuarios finales adicionales.
- Aprobar formalmente la estrategia de seguridad para una siguiente fase.
