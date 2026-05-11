# Informe Gerencial - Proyecto Renta PX

## Resumen ejecutivo

Renta PX es una aplicacion web estatica desarrollada para visualizar la base de arriendos y subrentas de equipos de Provexpress SAS, utilizando como fuente un archivo Excel alojado en SharePoint.

La solucion permite a Gerencia consultar indicadores consolidados, filtrar por comercial, revisar calidad de datos, visualizar graficas y exportar informacion en Excel con formato corporativo. Adicionalmente, los comerciales autorizados pueden consultar su cartera de Renta y, cuando aplique, su cartera de Subrenta.

El proyecto ya se encuentra publicado en GitHub Pages y opera con autenticacion Microsoft mediante MSAL y Microsoft Graph.

## Objetivo de negocio

Centralizar la visualizacion de la base de arriendos y subrentas para facilitar el seguimiento gerencial, comercial y operativo de los equipos arrendados.

La herramienta busca reducir consultas manuales sobre el archivo fuente, entregar informacion filtrada por rol y generar una base visual para decisiones sobre canon, costos, utilidad, margen, clientes, comerciales y calidad del dato.

## Alcance funcional

La aplicacion cubre:

- Autenticacion corporativa Microsoft.
- Lectura de Excel desde SharePoint.
- Dashboard de indicadores.
- Graficas por cliente, comercial, tipo y marca.
- Tabla detallada con filtros.
- Exportacion a Excel.
- Calidad del dato.
- Vista de Renta.
- Vista de Subrenta.
- Control visual por rol.

## Sitio y repositorio

**Repositorio GitHub:**  
https://github.com/Provexpress/renta-px

**Sitio publicado:**  
https://provexpress.github.io/renta-px/

## Fuente de informacion

La fuente principal es el archivo:

`RENTA_2024_T_1.xlsx`

Ubicado en SharePoint:

`https://provexpress.sharepoint.com/sites/ProvexpressIntranet/servicios`

Ruta funcional:

`Documentos compartidos / Coordinador de servicios / Renta`

Hojas utilizadas:

- `PROVEXPRESS`: informacion principal de Renta.
- `PC COM`: informacion de Subrenta.

## Indicadores disponibles

El dashboard calcula automaticamente:

- Total de equipos.
- Clientes activos.
- Canon mensual.
- Costo mensual.
- Utilidad mensual.
- Margen.
- Alertas de calidad.
- Distribucion por tipo.
- Distribucion por marca.
- Top clientes por canon.
- Top comerciales por canon, para roles autorizados.

## Vista Gerencia

Gerencia cuenta con la vista mas completa de la solucion.

Capacidades:

- Consultar Renta.
- Consultar Subrenta.
- Ver consolidado global.
- Filtrar por comercial.
- Visualizar KPIs y graficas.
- Consultar tabla completa.
- Exportar informacion filtrada.
- Revisar alertas de calidad del dato.

La vista gerencial permite comparar comportamientos por cartera comercial y revisar los datos de forma agregada o segmentada.

## Vista Comercial

El comercial visualiza unicamente la informacion asociada a su nombre comercial configurado.

Capacidades:

- Consultar su cartera de Renta.
- Consultar su cartera de Subrenta si esta autorizado.
- Ver sus propios KPIs.
- Ver sus clientes.
- Filtrar su tabla.
- Exportar su cartera a Excel.

Comerciales con Subrenta habilitada:

- Astrid Jimenez.
- Angela Torres.
- Karent Carrillo.
- Dilma Cuesta.

## Vista Finanzas

El rol Finanzas esta previsto para consultar informacion de rentabilidad global:

- Canon mensual.
- Costo mensual.
- Utilidad mensual.
- Margen.
- Rentabilidad por cliente.
- Rentabilidad por comercial.
- Rentabilidad por tipo de equipo.

## Vista Operaciones

El rol Operaciones esta previsto para consultar informacion de inventario:

- Tipo.
- Marca.
- Modelo.
- Serial.
- Placa.
- Procesador.
- Memoria.
- Disco.
- Garantia.
- Accesorios.
- Fecha de entrega.
- Alertas de datos.

## Tabla y exportacion

La tabla permite:

- Buscar registros.
- Filtrar por cliente.
- Filtrar por marca.
- Filtrar por tipo.
- Filtrar por estado.
- Ordenar columnas.
- Paginar resultados.
- Exportar a Excel.

La exportacion genera un archivo `.xlsx` con estilo corporativo, encabezado visual, filas alternadas y formatos de moneda. El nombre del archivo usa el cliente cuando la seleccion corresponde a un unico cliente.

Columnas incluidas en tabla y exportacion:

- Cliente.
- Comercial.
- Tipo.
- Marca.
- Modelo.
- Serial.
- Placa.
- Memoria.
- Tamano disco.
- Garantia.
- Office.
- Morral.
- Guaya.
- Mouse.
- Teclado.
- Monitor.
- Accesorios.
- Fecha entrega.
- Valor arriendo.

Segun el rol, tambien se pueden visualizar columnas financieras adicionales como costo, utilidad y margen.

## Calidad del dato

La aplicacion incluye un panel de calidad con alertas automaticas:

- Registros sin placa.
- Registros sin fecha de entrega.
- Registros sin valor de arriendo.
- Registros sin costo de renta.
- Registros con utilidad inconsistente.
- Registros con accesorios sin dato.

Estados utilizados:

- **Completo:** registro sin observaciones relevantes.
- **Revisar:** faltan datos operativos o accesorios.
- **Error:** faltan datos financieros o existe inconsistencia de utilidad.

## Seguridad y permisos

La aplicacion utiliza Microsoft Graph con permisos delegados. Esto implica que el usuario debe iniciar sesion con su cuenta Microsoft y debe tener permisos reales sobre el archivo en SharePoint.

Permisos Microsoft Graph configurados:

- `User.Read`
- `Files.Read.All`
- `Sites.Read.All`

La aplicacion no usa client secret, debido a que es una aplicacion SPA publica.

## Limitacion relevante

El control por comercial se realiza en el navegador despues de descargar el archivo Excel desde SharePoint. Por esta razon, la solucion actual debe considerarse una capa operativa de visualizacion y no una barrera de seguridad fuerte.

Para una seguridad robusta, el archivo completo no deberia llegar al navegador de un comercial. En una fase futura se recomienda filtrar la informacion desde un backend, Azure Function, Power Automate/API o separar archivos y permisos por comercial.

## Beneficios para Gerencia

- Acceso rapido al consolidado de arriendos.
- Visibilidad por comercial y cliente.
- Indicadores financieros calculados desde datos reales.
- Identificacion de alertas de calidad.
- Exportacion facil de informacion para seguimiento.
- Separacion entre Renta y Subrenta.
- Reduccion de dependencia de consultas manuales al Excel.

## Recomendaciones gerenciales

1. Mantener el Excel fuente bajo control de responsables definidos.
2. Validar periodicamente calidad de placa, fecha, costos y utilidad.
3. Revisar permisos de SharePoint para evitar accesos no deseados.
4. Evaluar una fase 2 con backend para seguridad por comercial.
5. Definir administradores funcionales del mapeo de usuarios y roles.
6. Formalizar reglas de actualizacion de las hojas `PROVEXPRESS` y `PC COM`.

## Estado actual

La solucion se encuentra funcional y publicada. Actualmente permite consulta de Renta y Subrenta, autenticacion Microsoft, segmentacion por roles, filtro gerencial por comercial, exportacion en Excel y visualizacion de alertas de calidad.

## Conclusion

Renta PX entrega una primera version funcional y de rapida adopcion para seguimiento de arriendos y subrentas. La solucion responde a la necesidad de visibilidad gerencial y comercial, manteniendo una arquitectura ligera y compatible con GitHub Pages.

Para consolidarla como herramienta corporativa de mayor seguridad y gobierno de datos, se recomienda planear una segunda fase con filtrado del lado servidor, administracion centralizada de roles y auditoria de accesos.
