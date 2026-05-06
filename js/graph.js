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

  async function getSiteId() {
    const file = APP_CONFIG.graph.sharePointFile;
    const site = await graphFetch(`/sites/${file.siteHostname}:${file.sitePath}`);
    return site.id;
  }

  async function getDriveId(siteId) {
    const file = APP_CONFIG.graph.sharePointFile;
    const drives = await graphFetch(`/sites/${siteId}/drives`);
    const allowedNames = [file.driveName, ...(file.driveNameAliases || [])]
      .filter(Boolean)
      .map((name) => name.toLowerCase());
    const drive = (drives.value || []).find((item) => allowedNames.includes(String(item.name || "").toLowerCase()));

    if (!drive) {
      const available = (drives.value || []).map((item) => item.name).join(", ");
      const error = new Error(`No se encontró la biblioteca "${file.driveName}" en SharePoint. Bibliotecas disponibles: ${available}`);
      error.status = 404;
      throw error;
    }

    return drive.id;
  }

  async function getFileItem(siteId, driveId) {
    const file = APP_CONFIG.graph.sharePointFile;
    return graphFetch(`/sites/${siteId}/drives/${driveId}/root:${encodeURI(file.filePath)}`);
  }

  async function downloadExcelFile() {
    const siteId = await getSiteId();
    const driveId = await getDriveId(siteId);
    const item = await getFileItem(siteId, driveId);

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
