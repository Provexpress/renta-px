const APP_CONFIG = {
  msal: {
    clientId: "REEMPLAZAR_CLIENT_ID",
    tenantId: "REEMPLAZAR_TENANT_ID",
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
      sheetName: "Hoja1",
      subRentSheetName: "PC COM"
    }
  }
};
