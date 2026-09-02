(function () {
  const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

  async function graphFetch(endpoint, options = {}) {
    const token = await AuthService.getAccessToken();
    const response = await fetch(`${GRAPH_BASE}${endpoint}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(options.headers || {})
      }
    });

    if (!response.ok) {
      const body = await response.text();
      const parsedError = parseGraphError(body);
      const error = new Error(parsedError.message || body || response.statusText);
      error.status = response.status;
      error.graphCode = parsedError.code;
      error.rawBody = body;
      throw error;
    }

    if (options.responseType === "arrayBuffer") {
      return response.arrayBuffer();
    }

    if (response.status === 204) {
      return null;
    }

    return response.json();
  }

  function parseGraphError(body) {
    try {
      const payload = JSON.parse(body);
      return {
        code: payload.error && payload.error.code,
        message: payload.error && payload.error.message
      };
    } catch (error) {
      return {
        code: "",
        message: body
      };
    }
  }

  async function getSiteId(fileConfig = APP_CONFIG.graph.sharePointFile) {
    const site = await graphFetch(`/sites/${fileConfig.siteHostname}:${fileConfig.sitePath}`);
    return site.id;
  }

  async function getDriveId(siteId, fileConfig = APP_CONFIG.graph.sharePointFile) {
    const drives = await graphFetch(`/sites/${siteId}/drives`);
    const allowedNames = [fileConfig.driveName, ...(fileConfig.driveNameAliases || [])]
      .filter(Boolean)
      .map((name) => name.toLowerCase());
    const drive = (drives.value || []).find((item) => allowedNames.includes(String(item.name || "").toLowerCase()));

    if (!drive) {
      const available = (drives.value || []).map((item) => item.name).join(", ");
      const error = new Error(`No se encontró la biblioteca "${fileConfig.driveName}" en SharePoint. Bibliotecas disponibles: ${available}`);
      error.status = 404;
      throw error;
    }

    return drive.id;
  }

  async function getFileItem(siteId, driveId, fileConfig = APP_CONFIG.graph.sharePointFile) {
    // 1. Intentar directamente por itemId / GUID si está disponible
    if (fileConfig.itemId) {
      try {
        const item = await graphFetch(`/sites/${siteId}/drives/${driveId}/items/${fileConfig.itemId}`);
        if (item && item.id) return item;
      } catch (err) {
        try {
          const item = await graphFetch(`/sites/${siteId}/drive/items/${fileConfig.itemId}`);
          if (item && item.id) {
            item._driveId = item.parentReference && item.parentReference.driveId;
            return item;
          }
        } catch (e) {}
      }
    }

    // 2. Intentar por rutas codificadas
    const pathsToTry = [
      fileConfig.filePath,
      ...(fileConfig.filePathAliases || [])
    ].filter(Boolean);

    for (const rawPath of pathsToTry) {
      const cleanPath = rawPath.startsWith("/") ? rawPath : "/" + rawPath;
      const encodedSegments = cleanPath
        .split("/")
        .filter(Boolean)
        .map((segment) => encodeURIComponent(segment))
        .join("/");

      try {
        const item = await graphFetch(`/sites/${siteId}/drives/${driveId}/root:/${encodedSegments}`);
        if (item && item.id) return item;
      } catch (e) {}

      try {
        const item = await graphFetch(`/sites/${siteId}/drives/${driveId}/root:${cleanPath}`);
        if (item && item.id) return item;
      } catch (e) {}

      try {
        const item = await graphFetch(`/sites/${siteId}/drives/${driveId}/root:${encodeURI(cleanPath)}`);
        if (item && item.id) return item;
      } catch (e) {}
    }

    // 3. Estrategia de búsqueda por nombre de archivo en la biblioteca
    const fileName = (fileConfig.filePath || "").split("/").pop() || "Equipos.xlsx";
    try {
      const searchRes = await graphFetch(`/sites/${siteId}/drives/${driveId}/root/search(q='${encodeURIComponent(fileName)}')`);
      if (searchRes && searchRes.value && searchRes.value.length > 0) {
        const matched = searchRes.value.find((i) => (i.name || "").toLowerCase() === fileName.toLowerCase()) || searchRes.value[0];
        if (matched) return matched;
      }
    } catch (e) {}

    // 4. Búsqueda en todas las bibliotecas del sitio
    try {
      const drivesRes = await graphFetch(`/sites/${siteId}/drives`);
      for (const d of (drivesRes.value || [])) {
        if (d.id === driveId) continue;
        try {
          const searchRes = await graphFetch(`/sites/${siteId}/drives/${d.id}/root/search(q='${encodeURIComponent(fileName)}')`);
          if (searchRes && searchRes.value && searchRes.value.length > 0) {
            const matched = searchRes.value.find((i) => (i.name || "").toLowerCase() === fileName.toLowerCase()) || searchRes.value[0];
            if (matched) {
              matched._driveId = d.id;
              return matched;
            }
          }
        } catch (e) {}
      }
    } catch (e) {}

    const error = new Error(`No se encontró el archivo "${fileConfig.filePath}" en SharePoint.`);
    error.status = 404;
    error.graphCode = "itemNotFound";
    throw error;
  }

  async function downloadExcelFile(fileConfig = APP_CONFIG.graph.sharePointFile) {
    const siteId = await getSiteId(fileConfig);
    const driveId = await getDriveId(siteId, fileConfig);
    const item = await getFileItem(siteId, driveId, fileConfig);
    const targetDriveId = item._driveId || driveId;

    return graphFetch(`/sites/${siteId}/drives/${targetDriveId}/items/${item.id}/content`, {
      responseType: "arrayBuffer"
    });
  }

  window.GraphService = {
    graphFetch,
    getSiteId,
    getDriveId,
    getFileItem,
    downloadExcelFile
  };
})();
