const APP_CONFIG = {
  msal: {
    clientId: "cc4ed0ed-1eb1-45a3-82bc-25a8df0f3f03",
    tenantId: "e6805558-f5bb-444c-8af2-5f3a4d6dd3fc",
    redirectUri: "https://provexpress.github.io/renta-px/"
  },
  graph: {
    scopes: ["User.Read", "Files.Read.All", "Sites.Read.All"],
    sharePointFile: {
      siteHostname: "provexpress.sharepoint.com",
      sitePath: "/sites/ProvexpressIntranet/servicios",
      driveName: "Documentos",
      driveNameAliases: ["Documentos compartidos", "Documents"],
      filePath: "/Coordinador de servicios/Renta/RENTA_2024_T_1.xlsx",
      sheetName: "PROVEXPRESS",
      subRentSheetName: "PC COM",
      accessoriesSheetName: "ACCESORIOS"
    },
    availableInventoryFile: {
      siteHostname: "provexpress.sharepoint.com",
      sitePath: "/sites/ProvexpressIntranet/servicios",
      driveName: "Documentos",
      driveNameAliases: ["Documentos compartidos", "Documents", "Shared Documents"],
      filePath: "/Técnico de soporte/Inventario Renta Bodega/Equipos.xlsx",
      filePathAliases: [
        "/Técnico de soporte/Inventario Renta Bodega/Equipos.xlsx",
        "/Tecnico de soporte/Inventario Renta Bodega/Equipos.xlsx",
        "/Inventario Renta Bodega/Equipos.xlsx",
        "Equipos.xlsx"
      ],
      itemId: "1FA0AA63-5AED-4ADC-B947-E9540DA81880",
      sheetName: "Disponible para Rentar"
    }
  }
};
