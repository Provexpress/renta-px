(function () {
  let chartInstances = [];
  let activeCommercialFilter = "";
  let activeDataset = "renta";

  function renderDashboard({ profile, user, rows, scopedRows, subRentRows = [] }) {
    destroyCharts();
    const datasetRows = getDatasetRows(user, scopedRows, subRentRows);
    const visibleRows = getVisibleRows(user, datasetRows);

    const app = document.getElementById("app");
    app.innerHTML = `
      <main class="dashboard-page">
        ${renderHeader(profile, user)}
        <p class="status-line">Datos cargados desde SharePoint</p>
        ${renderDatasetTabs(user, subRentRows)}
        ${renderDashboardControls(user, datasetRows, visibleRows)}
        ${renderKpis(user, visibleRows)}
        <section class="content-grid">
          ${renderQualityPanel(visibleRows)}
          ${renderCharts(user)}
          ${TableView.renderTableShell(user, getDatasetLabel())}
        </section>
      </main>
    `;

    bindDashboardEvents({ profile, user, rows, scopedRows, subRentRows });
    renderChartInstances(user, visibleRows);
    TableView.initTable(user, visibleRows, { datasetLabel: getDatasetLabel() });
  }

  function renderDatasetTabs(user, subRentRows) {
    if (user.role !== "gerencia") {
      return "";
    }

    return `
      <nav class="dataset-tabs" aria-label="Vistas de arriendo">
        <button class="dataset-tab ${activeDataset === "renta" ? "active" : ""}" type="button" data-dataset="renta">Renta</button>
        <button class="dataset-tab ${activeDataset === "subrenta" ? "active" : ""}" type="button" data-dataset="subrenta"${subRentRows.length ? "" : " disabled"}>Subrenta</button>
      </nav>
    `;
  }

  function renderHeader(profile, user) {
    const displayName = escapeHtml(profile.displayName || "Usuario Microsoft");
    const email = escapeHtml(profile.mail || profile.userPrincipalName || user.email || "");
    const role = escapeHtml(getRoleLabel(user.role));

    return `
      <header class="topbar">
        <div class="brand-row">
          <img class="brand-mark" src="assets/provex_icon_64.png" alt="Provexpress">
          <div>
            <p class="eyebrow">Provex One</p>
            <h1 class="brand-title">Renta PX</h1>
            <p class="brand-subtitle">Dashboard de Arriendos</p>
          </div>
        </div>
        <div class="user-panel">
          <div>
            <div class="user-name">${displayName}</div>
            <div class="user-meta">${email} · Rol: ${role}</div>
          </div>
          <button class="secondary-button" id="logoutButton">Cerrar sesión</button>
        </div>
      </header>
    `;
  }

  function renderDashboardControls(user, rows, visibleRows) {
    if (user.role !== "gerencia") {
      return "";
    }

    const options = getCommercialOptions(rows);
    const selectedLabel = activeCommercialFilter || "Todos los comerciales";

    return `
      <section class="dashboard-controls" aria-label="Filtros de gerencia">
        <div>
          <p class="eyebrow">Filtro gerencial</p>
          <h2 class="control-title">${escapeHtml(selectedLabel)}</h2>
          <p class="control-copy">${formatNumber(visibleRows.length)} registros visibles</p>
        </div>
        <label class="control-field">
          <span>Comercial</span>
          <select id="dashboardCommercialFilter">
            <option value="">Todos los comerciales</option>
            ${options.map((name) => `
              <option value="${escapeHtml(name)}"${name === activeCommercialFilter ? " selected" : ""}>${escapeHtml(name)}</option>
            `).join("")}
          </select>
        </label>
      </section>
    `;
  }

  function bindDashboardEvents(context) {
    const logoutButton = document.getElementById("logoutButton");
    if (logoutButton) {
      logoutButton.addEventListener("click", () => AuthService.logout());
    }

    const commercialFilter = document.getElementById("dashboardCommercialFilter");
    if (commercialFilter) {
      commercialFilter.addEventListener("change", () => {
        activeCommercialFilter = commercialFilter.value;
        renderDashboard(context);
      });
    }

    document.querySelectorAll("[data-dataset]").forEach((button) => {
      button.addEventListener("click", () => {
        activeDataset = button.getAttribute("data-dataset");
        activeCommercialFilter = "";
        renderDashboard(context);
      });
    });
  }

  function getDatasetRows(user, scopedRows, subRentRows) {
    if (user.role === "gerencia" && activeDataset === "subrenta") {
      return subRentRows;
    }

    return scopedRows;
  }

  function getDatasetLabel() {
    return activeDataset === "subrenta" ? "Subrenta" : "Renta";
  }

  function getVisibleRows(user, rows) {
    if (user.role !== "gerencia" || !activeCommercialFilter) {
      return rows;
    }

    return rows.filter((row) => ExcelService.comparableText(row.comercial) === ExcelService.comparableText(activeCommercialFilter));
  }

  function getCommercialOptions(rows) {
    const byKey = new Map();

    rows.forEach((row) => {
      const name = row.comercial;
      const key = ExcelService.comparableText(name);
      if (name && key && !byKey.has(key)) {
        byKey.set(key, name);
      }
    });

    return Array.from(byKey.values()).sort((a, b) => a.localeCompare(b, "es"));
  }

  function renderKpis(user, rows) {
    const metrics = getMetrics(rows);
    let cards;

    if (user.role === "operaciones") {
      const quality = getQualityCounts(rows);
      cards = [
        { label: "Equipos", value: formatNumber(metrics.equipos) },
        { label: "Clientes", value: formatNumber(metrics.clientes) },
        { label: "Comerciales", value: formatNumber(metrics.comerciales) },
        { label: "Alertas", value: formatNumber(metrics.alertas), help: "Datos por revisar" },
        { label: "Sin placa", value: formatNumber(quality.sinPlaca) },
        { label: "Sin fecha", value: formatNumber(quality.sinFechaEntrega) },
        { label: "Completos", value: formatNumber(rows.filter((row) => row.dataQualityStatus === "ok").length) }
      ];
    } else if (PermissionService.isCommercial(user)) {
      cards = [
        { label: "Mis equipos", value: formatNumber(metrics.equipos) },
        { label: "Mis clientes", value: formatNumber(metrics.clientes) },
        { label: "Mi canon mensual", value: formatCurrency(metrics.valorArriendo) },
        { label: "Mi costo mensual", value: formatCurrency(metrics.costoRenta) },
        { label: "Mi utilidad mensual", value: formatCurrency(metrics.utilidadRenta) },
        { label: "Mi margen", value: formatPercent(metrics.margen) },
        { label: "Mis alertas", value: formatNumber(metrics.alertas), help: "Datos por revisar" }
      ];
    } else {
      cards = [
        { label: "Equipos", value: formatNumber(metrics.equipos) },
        { label: "Clientes", value: formatNumber(metrics.clientes) },
        { label: "Canon mensual", value: formatCurrency(metrics.valorArriendo) },
        { label: "Costo mensual", value: formatCurrency(metrics.costoRenta) },
        { label: "Utilidad mensual", value: formatCurrency(metrics.utilidadRenta) },
        { label: "Margen global", value: formatPercent(metrics.margen) },
        { label: "Alertas", value: formatNumber(metrics.alertas), help: "Datos por revisar" }
      ];
    }

    return `
      <section class="kpi-grid">
        ${cards.map((card) => `
          <article class="kpi-card">
            <div class="kpi-label">${escapeHtml(card.label)}</div>
            <div class="kpi-value">${card.value}</div>
            ${card.help ? `<div class="kpi-help">${escapeHtml(card.help)}</div>` : ""}
          </article>
        `).join("")}
      </section>
    `;
  }

  function renderQualityPanel(rows) {
    const quality = getQualityCounts(rows);
    const items = [
      ["Registros sin placa", quality.sinPlaca],
      ["Registros sin fecha de entrega", quality.sinFechaEntrega],
      ["Registros sin valor de arriendo", quality.sinValorArriendo],
      ["Registros sin costo de renta", quality.sinCostoRenta],
      ["Registros con utilidad inconsistente", quality.utilidadInconsistente],
      ["Registros con accesorios sin dato", quality.accesoriosSinDato]
    ];

    return `
      <section class="panel">
        <h2 class="panel-title">Calidad del dato</h2>
        <div class="quality-grid">
          ${items.map(([label, value]) => `
            <div class="quality-item">
              <div class="quality-number">${formatNumber(value)}</div>
              <div class="quality-label">${escapeHtml(label)}</div>
            </div>
          `).join("")}
        </div>
      </section>
    `;
  }

  function renderCharts(user) {
    const showCommercialRanking = PermissionService.canViewCommercialRanking(user);
    const clientTitle = PermissionService.isCommercial(user)
      ? "Mis clientes por canon"
      : PermissionService.canViewFinancials(user)
        ? "Top clientes por canon"
        : "Equipos por cliente";

    return `
      <section class="chart-grid">
        <article class="chart-panel">
          <h2 class="panel-title">${clientTitle}</h2>
          <div class="chart-box"><canvas id="clientsChart"></canvas></div>
        </article>
        ${showCommercialRanking ? `
          <article class="chart-panel">
            <h2 class="panel-title">Top comerciales por canon</h2>
            <div class="chart-box"><canvas id="commercialChart"></canvas></div>
          </article>
        ` : ""}
        <article class="chart-panel">
          <h2 class="panel-title">Equipos por tipo</h2>
          <div class="chart-box"><canvas id="typeChart"></canvas></div>
        </article>
        <article class="chart-panel">
          <h2 class="panel-title">Equipos por marca</h2>
          <div class="chart-box"><canvas id="brandChart"></canvas></div>
        </article>
        ${PermissionService.canViewFinancials(user) ? `
          <article class="chart-panel">
            <h2 class="panel-title">Canon vs costo vs utilidad</h2>
            <div class="chart-box"><canvas id="financialChart"></canvas></div>
          </article>
        ` : ""}
      </section>
    `;
  }

  function renderChartInstances(user, rows) {
    if (PermissionService.canViewFinancials(user)) {
      createBarChart("clientsChart", groupSum(rows, "cliente", "valorArriendo", 8), "Canon");
    } else {
      createBarChart("clientsChart", groupCount(rows, "cliente", 8), "Equipos");
    }

    if (PermissionService.canViewCommercialRanking(user)) {
      createBarChart("commercialChart", groupSum(rows, "comercial", "valorArriendo", 8), "Canon");
    }

    createDoughnutChart("typeChart", groupCount(rows, "tipo", 8));
    createDoughnutChart("brandChart", groupCount(rows, "marca", 8));

    if (PermissionService.canViewFinancials(user)) {
      const metrics = getMetrics(rows);
      createBarChart("financialChart", [
        { label: "Canon", value: metrics.valorArriendo },
        { label: "Costo", value: metrics.costoRenta },
        { label: "Utilidad", value: metrics.utilidadRenta }
      ], "COP");
    }
  }

  function createBarChart(canvasId, data, label) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    chartInstances.push(new Chart(canvas, {
      type: "bar",
      data: {
        labels: data.map((item) => item.label),
        datasets: [{
          label,
          data: data.map((item) => item.value),
          backgroundColor: "#1565c0",
          borderRadius: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false }
        },
        scales: {
          x: { ticks: { color: "#64748b" }, grid: { display: false } },
          y: { ticks: { color: "#64748b" } }
        }
      }
    }));
  }

  function createDoughnutChart(canvasId, data) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    chartInstances.push(new Chart(canvas, {
      type: "doughnut",
      data: {
        labels: data.map((item) => item.label),
        datasets: [{
          data: data.map((item) => item.value),
          backgroundColor: ["#1a2b6b", "#1565c0", "#6a3fa0", "#2abfdf", "#16a34a", "#d97706", "#c62828", "#677592"]
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: "bottom" }
        }
      }
    }));
  }

  function destroyCharts() {
    chartInstances.forEach((chart) => chart.destroy());
    chartInstances = [];
  }

  function getMetrics(rows) {
    const valorArriendo = sum(rows, "valorArriendo");
    const costoRenta = sum(rows, "costoRenta");
    const utilidadRenta = sum(rows, "utilidadRenta");

    return {
      equipos: rows.length,
      clientes: countUnique(rows, "cliente"),
      comerciales: countUnique(rows, "comercial"),
      valorArriendo,
      costoRenta,
      utilidadRenta,
      margen: valorArriendo ? utilidadRenta / valorArriendo : 0,
      alertas: rows.filter((row) => row.dataQualityStatus !== "ok").length
    };
  }

  function getQualityCounts(rows) {
    return rows.reduce((acc, row) => {
      row.validations.forEach((issue) => {
        acc[issue] = (acc[issue] || 0) + 1;
      });
      return acc;
    }, {
      sinPlaca: 0,
      sinFechaEntrega: 0,
      sinValorArriendo: 0,
      sinCostoRenta: 0,
      utilidadInconsistente: 0,
      accesoriosSinDato: 0
    });
  }

  function sum(rows, key) {
    return rows.reduce((total, row) => total + (Number(row[key]) || 0), 0);
  }

  function countUnique(rows, key) {
    return new Set(rows.map((row) => ExcelService.comparableText(row[key])).filter(Boolean)).size;
  }

  function groupCount(rows, key, limit) {
    const groups = rows.reduce((acc, row) => {
      const label = row[key] || "Sin dato";
      acc[label] = (acc[label] || 0) + 1;
      return acc;
    }, {});

    return sortGroups(groups, limit);
  }

  function groupSum(rows, labelKey, valueKey, limit) {
    const groups = rows.reduce((acc, row) => {
      const label = row[labelKey] || "Sin dato";
      acc[label] = (acc[label] || 0) + (Number(row[valueKey]) || 0);
      return acc;
    }, {});

    return sortGroups(groups, limit);
  }

  function sortGroups(groups, limit) {
    return Object.entries(groups)
      .map(([label, value]) => ({ label, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, limit);
  }

  function formatCurrency(value) {
    return new Intl.NumberFormat("es-CO", {
      style: "currency",
      currency: "COP",
      maximumFractionDigits: 0
    }).format(value || 0);
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("es-CO").format(value || 0);
  }

  function formatPercent(value) {
    return new Intl.NumberFormat("es-CO", {
      style: "percent",
      maximumFractionDigits: 1
    }).format(value || 0);
  }

  function getRoleLabel(role) {
    const labels = {
      gerencia: "Gerencia",
      finanzas: "Finanzas",
      operaciones: "Operaciones",
      comercial: "Comercial",
      sin_permiso: "Sin permiso"
    };
    return labels[role] || role;
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  window.DashboardView = {
    renderDashboard,
    getMetrics,
    formatCurrency,
    formatNumber,
    formatPercent,
    escapeHtml
  };
})();
