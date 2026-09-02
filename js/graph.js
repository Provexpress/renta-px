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
    const pathsToTry = [
      fileConfig.filePath,
      ...(fileConfig.filePathAliases || [])
    ].filter(Boolean);

    let lastError = null;
    for (const filePath of pathsToTry) {
      try {
        const encodedPath = filePath
          .split("/")
          .map((segment) => encodeURIComponent(segment))
          .join("/");
        return await graphFetch(`/sites/${siteId}/drives/${driveId}/root:${encodedPath}`);
      } catch (err) {
        lastError = err;
        try {
          return await graphFetch(`/sites/${siteId}/drives/${driveId}/root:${encodeURI(filePath)}`);
        } catch (err2) {
          lastError = err2;
        }
      }
    }
    throw lastError;
  }

  async function downloadExcelFile(fileConfig = APP_CONFIG.graph.sharePointFile) {
    const siteId = await getSiteId(fileConfig);
    const driveId = await getDriveId(siteId, fileConfig);
    const item = await getFileItem(siteId, driveId, fileConfig);

    return graphFetch(`/sites/${siteId}/drives/${driveId}/items/${item.id}/content`, {
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
