(function () {
  const app = document.getElementById("app");

  document.addEventListener("DOMContentLoaded", start);

  async function start() {
    try {
      validateConfig();
      renderLoading("Autenticando...");
      const account = await AuthService.initializeMsal();

      if (!account) {
        renderLogin();
        return;
      }

      await loadDashboard();
    } catch (error) {
      renderError(getFriendlyError(error), error);
    }
  }

  async function loadDashboard() {
    try {
      renderLoading("Cargando datos desde SharePoint...");
      const profile = await AuthService.getUserProfile();
      const user = PermissionService.getCurrentUserAccess(profile);

      if (user.role === "sin_permiso") {
        renderError("No tienes permisos para consultar este archivo.");
        return;
      }

      const arrayBuffer = await GraphService.downloadExcelFile();
      const rawRows = ExcelService.readExcelArrayBuffer(arrayBuffer, APP_CONFIG.graph.sharePointFile.sheetName);
      const rows = ExcelService.normalizeRows(rawRows);
      const scopedRows = PermissionService.getScopedRows(user, rows);
      const subRentRows = getSubRentRows(user, arrayBuffer);

      if (!scopedRows.length && PermissionService.isCommercial(user)) {
        renderError("No hay registros asignados a este comercial.");
        return;
      }

      DashboardView.renderDashboard({ profile, user, rows, scopedRows, subRentRows });
      if (PermissionService.canViewGlobalDashboard(user)) {
        validateExpectedReading(rows);
      }
    } catch (error) {
      renderError(getFriendlyError(error), error);
    }
  }

  function getSubRentRows(user, arrayBuffer) {
    const sheetName = APP_CONFIG.graph.sharePointFile.subRentSheetName;
    if (user.role !== "gerencia" || !sheetName) {
      return [];
    }

    try {
      const rawRows = ExcelService.readExcelArrayBuffer(arrayBuffer, sheetName);
      return ExcelService.normalizeRows(rawRows);
    } catch (error) {
      console.warn("No se pudo leer la hoja de subrenta.", error);
      return [];
    }
  }

  function renderLogin() {
    app.innerHTML = `
      <section class="state-screen">
        <div class="state-card">
          <img class="brand-mark" src="assets/provex_icon_64.png" alt="Provexpress">
          <h1>Renta PX</h1>
          <p>Inicia sesión con Microsoft para continuar.</p>
          <button class="primary-button" id="loginButton">Iniciar sesión con Microsoft</button>
        </div>
      </section>
    `;

    document.getElementById("loginButton").addEventListener("click", async () => {
      try {
        renderLoading("Autenticando...");
        await AuthService.login();
        await loadDashboard();
      } catch (error) {
        renderError(getFriendlyError(error), error);
      }
    });
  }

  function renderLoading(message) {
    app.innerHTML = `
      <section class="state-screen">
        <div class="state-card">
          <img class="brand-mark" src="assets/provex_icon_64.png" alt="Provexpress">
          <h1>Renta PX</h1>
          <p>${DashboardView ? DashboardView.escapeHtml(message) : message}</p>
        </div>
      </section>
    `;
  }

  function renderError(message, error) {
    const detail = getErrorDetail(error);
    app.innerHTML = `
      <section class="state-screen">
        <div class="state-card">
          <img class="brand-mark" src="assets/provex_icon_64.png" alt="Provexpress">
          <h1>Renta PX</h1>
          <p>${DashboardView.escapeHtml(message)}</p>
          <button class="primary-button" id="retryButton">Reintentar</button>
          ${detail ? `<div class="error-detail">${DashboardView.escapeHtml(detail)}</div>` : ""}
        </div>
      </section>
    `;

    document.getElementById("retryButton").addEventListener("click", start);
  }

  function getFriendlyError(error) {
    if (!error) return "Ocurrió un error inesperado.";
    if (error.status === 403 || error.status === 401 || error.graphCode === "accessDenied") {
      return "No tienes permisos para consultar este archivo.";
    }
    if (error.status === 404 && error.graphCode === "itemNotFound") {
      return "No se encontró el archivo en SharePoint o tu usuario no tiene acceso de lectura.";
    }
    if (error.status === 404) return "No se encontró el archivo en SharePoint.";
    if ((error.message || "").includes("No se pudo leer el Excel")) return "No se pudo leer el Excel.";
    return error.friendlyMessage || "Ocurrió un error cargando el dashboard.";
  }

  function getErrorDetail(error) {
    if (!error) return "";
    if (error.graphCode === "accessDenied" || (error.message || "").includes("accessDenied")) {
      return "Revisa consentimiento de Microsoft Graph y permiso de lectura del usuario en SharePoint.";
    }
    if (error.graphCode === "itemNotFound" || (error.message || "").includes("itemNotFound")) {
      return "Revisa la ruta del Excel y el permiso de lectura del usuario en SharePoint.";
    }
    return error.message || String(error);
  }

  function validateConfig() {
    if (typeof APP_CONFIG === "undefined") {
      throw new Error("No existe js/config.js.");
    }

    if (APP_CONFIG.msal.clientId.includes("REEMPLAZAR") || APP_CONFIG.msal.tenantId.includes("REEMPLAZAR")) {
      throw new Error("Configura clientId y tenantId en js/config.js antes de iniciar sesión.");
    }
  }

  function validateExpectedReading(rows) {
    const metrics = DashboardView.getMetrics(rows);
    const typeCounts = rows.reduce((acc, row) => {
      const key = ExcelService.comparableText(row.tipo);
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});

    console.info("Validación de lectura Renta PX", {
      equipos: metrics.equipos,
      portatiles: typeCounts.portatil || 0,
      pc: typeCounts.pc || 0,
      canonMensual: metrics.valorArriendo,
      costoMensual: metrics.costoRenta,
      utilidadMensual: metrics.utilidadRenta,
      margen: metrics.margen,
      clientes: metrics.clientes,
      comerciales: metrics.comerciales
    });
  }
})();
