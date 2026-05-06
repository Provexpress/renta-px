const APP_CONFIG = {
  msal: {
    clientId: "cc4ed0ed-1eb1-45a3-82bc-25a8df0f3f03",
    tenantId: "e6805558-f5bb-444c-8af2-5f3a4d6dd3fc",
    redirectUri: "https://provexpress.github.io/renta-px/"
  },
  graph: {
    scopes: ["User.Read", "Files.Read.All", "Sites.Read.All"],
    sharePointFile: {
      siteHostname: "empresa.sharepoint.com",
      sitePath: "/sites/NOMBRE_DEL_SITIO",
      driveName: "Documents",
      filePath: "/Ruta/Archivo/Base.xlsx",
      sheetName: "Hoja1"
    }
  }
};
