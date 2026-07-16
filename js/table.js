(function () {
  const PAGE_SIZE = 25;
  const INVENTORY_COLUMNS = [
    { key: "memoria", label: "Memoria" },
    { key: "tamanoDisco", label: "Tamaño disco" },
    { key: "garantia", label: "Garantía" },
    { key: "office", label: "Office" },
    { key: "morral", label: "Morral" },
    { key: "guaya", label: "Guaya" },
    { key: "mouse", label: "Mouse" },
    { key: "teclado", label: "Teclado" },
    { key: "monitor", label: "Monitor" },
    { key: "accesorios", label: "Accesorios" }
  ];
  const ACCESSORIES_COLUMNS = [
    { key: "tipo", label: "Tipo" },
    { key: "marca", label: "Marca" },
    { key: "modelo", label: "Modelo" },
    { key: "serial", label: "Serial" },
    { key: "placa", label: "Placa" },
    { key: "comercial", label: "Comercial" },
    { key: "cliente", label: "Cliente" },
    { key: "valorArriendo", label: "Valor arriendo", type: "currency" },
    { key: "costoRenta", label: "Costo de renta", type: "currency" },
    { key: "utilidadRenta", label: "Utilidad de renta", type: "currency" },
    { key: "margen", label: "Margen", type: "percent" }
  ];
  let state = {
    rows: [],
    filteredRows: [],
    user: null,
    page: 1,
    sortKey: "cliente",
    sortDirection: "asc",
    datasetLabel: "Renta",
    datasetKey: "renta"
  };

  function renderTableShell(user, datasetLabel = "Renta") {
    const title = PermissionService.isCommercial(user)
      ? "Mi tabla de equipos"
      : `Tabla de equipos - ${datasetLabel}`;

    return `
      <section class="panel">
        <h2 class="panel-title">${title}</h2>
        <div class="table-toolbar">
          <input id="tableSearch" type="search" placeholder="Buscar">
          <select id="clientFilter"><option value="">Cliente</option></select>
          <select id="brandFilter"><option value="">Marca</option></select>
          <select id="typeFilter"><option value="">Tipo</option></select>
          <select id="statusFilter">
            <option value="">Estado</option>
            <option value="ok">Completo</option>
            <option value="warning">Advertencia</option>
            <option value="error">Error</option>
          </select>
          <button class="primary-button table-export-button" id="exportTableButton" type="button">Exportar Excel</button>
        </div>
        <div id="tableContainer"></div>
      </section>
    `;
  }

  function initTable(user, rows, options = {}) {
    state = {
      rows,
      filteredRows: rows,
      user,
      page: 1,
      sortKey: "cliente",
      sortDirection: "asc",
      datasetLabel: options.datasetLabel || "Renta",
      datasetKey: options.datasetKey || "renta"
    };

    fillSelect("clientFilter", uniqueValues(rows, "cliente"));
    fillSelect("brandFilter", uniqueValues(rows, "marca"));
    fillSelect("typeFilter", uniqueValues(rows, "tipo"));
    bindFilters();
    bindExport();
    applyFilters();
  }

  function bindFilters() {
    ["tableSearch", "clientFilter", "brandFilter", "typeFilter", "statusFilter"].forEach((id) => {
      document.getElementById(id).addEventListener("input", () => {
        state.page = 1;
        applyFilters();
      });
    });
  }

  function bindExport() {
    document.getElementById("exportTableButton").addEventListener("click", exportFilteredRows);
  }

  function applyFilters() {
    const search = ExcelService.comparableText(document.getElementById("tableSearch").value);
    const client = document.getElementById("clientFilter").value;
    const brand = document.getElementById("brandFilter").value;
    const type = document.getElementById("typeFilter").value;
    const status = document.getElementById("statusFilter").value;

    state.filteredRows = state.rows.filter((row) => {
      const searchText = ExcelService.comparableText([
        row.cliente,
        row.comercial,
        row.tipo,
        row.marca,
        row.modelo,
        row.serial,
        row.placa,
        row.memoria,
        row.tamanoDisco,
        row.garantia,
        row.office,
        row.morral,
        row.guaya,
        row.mouse,
        row.teclado,
        row.monitor,
        row.accesorios
      ].join(" "));

      return (!search || searchText.includes(search))
        && (!client || row.cliente === client)
        && (!brand || row.marca === brand)
        && (!type || row.tipo === type)
        && (!status || row.dataQualityStatus === status);
    });

    sortRows();
    renderTable();
  }

  function sortRows() {
    const direction = state.sortDirection === "asc" ? 1 : -1;
    state.filteredRows.sort((a, b) => {
      const left = a[state.sortKey];
      const right = b[state.sortKey];

      if (typeof left === "number" || typeof right === "number") {
        return ((Number(left) || 0) - (Number(right) || 0)) * direction;
      }

      return String(left || "").localeCompare(String(right || ""), "es") * direction;
    });
  }

  function renderTable() {
    const container = document.getElementById("tableContainer");
    const columns = getColumns(state.user);
    const totalPages = Math.max(1, Math.ceil(state.filteredRows.length / PAGE_SIZE));
    state.page = Math.min(state.page, totalPages);

    const start = (state.page - 1) * PAGE_SIZE;
    const pageRows = state.filteredRows.slice(start, start + PAGE_SIZE);

    if (!state.filteredRows.length) {
      container.innerHTML = `<div class="empty-state">No hay registros para los filtros seleccionados.</div>`;
      return;
    }

    container.innerHTML = `
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              ${columns.map((column) => `<th class="sortable" data-sort="${column.key}">${column.label}${getSortMark(column.key)}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${pageRows.map((row) => `
              <tr>
                ${columns.map((column) => `<td>${formatCell(row, column)}</td>`).join("")}
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
      <div class="pagination">
        <span>${DashboardView.formatNumber(state.filteredRows.length)} registros · Página ${state.page} de ${totalPages}</span>
        <div class="pagination-actions">
          <button class="secondary-button" id="prevPage">Anterior</button>
          <button class="secondary-button" id="nextPage">Siguiente</button>
        </div>
      </div>
    `;

    container.querySelectorAll("th[data-sort]").forEach((th) => {
      th.addEventListener("click", () => {
        const key = th.getAttribute("data-sort");
        if (state.sortKey === key) {
          state.sortDirection = state.sortDirection === "asc" ? "desc" : "asc";
        } else {
          state.sortKey = key;
          state.sortDirection = "asc";
        }
        applyFilters();
      });
    });

    document.getElementById("prevPage").addEventListener("click", () => {
      if (state.page > 1) {
        state.page -= 1;
        renderTable();
      }
    });

    document.getElementById("nextPage").addEventListener("click", () => {
      if (state.page < totalPages) {
        state.page += 1;
        renderTable();
      }
    });
  }

  function getColumns(user) {
    if (state.datasetKey === "accesorios") {
      return ACCESSORIES_COLUMNS;
    }

    const columns = [
      { key: "cliente", label: "Cliente" }
    ];

    if (PermissionService.isAdmin(user)) {
      columns.push({ key: "comercial", label: "Comercial" });
    }

    columns.push(
      { key: "tipo", label: "Tipo" },
      { key: "marca", label: "Marca" },
      { key: "modelo", label: "Modelo" },
      { key: "serial", label: "Serial" },
      { key: "placa", label: "Placa" },
      ...INVENTORY_COLUMNS,
      { key: "fechaEntrega", label: "Fecha entrega" }
    );

    if (PermissionService.canViewFinancials(user)) {
      columns.push(
        { key: "valorArriendo", label: "Valor arriendo", type: "currency" },
        { key: "costoRenta", label: "Costo", type: "currency" },
        { key: "utilidadRenta", label: "Utilidad", type: "currency" },
        { key: "margen", label: "Margen", type: "percent" }
      );
    }

    columns.push({ key: "dataQualityStatus", label: "Estado", type: "status" });
    return columns;
  }

  function formatCell(row, column) {
    const value = row[column.key];
    if (column.type === "currency") return DashboardView.formatCurrency(value);
    if (column.type === "percent") return DashboardView.formatPercent(value);
    if (column.type === "status") return statusBadge(value);
    return DashboardView.escapeHtml(value || "");
  }

  function exportFilteredRows() {
    if (!state.filteredRows.length) return;

    const columns = getExportColumns();
    const worksheet = buildStyledWorksheet(columns, state.filteredRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Renta PX");
    XLSX.writeFile(workbook, getExportFileName());
  }

  function getExportColumns() {
    if (state.datasetKey === "accesorios") {
      return ACCESSORIES_COLUMNS;
    }

    return getCommercialExportColumns();
  }

  function getCommercialExportColumns() {
    return [
      { key: "cliente", label: "Cliente" },
      { key: "comercial", label: "Comercial" },
      { key: "tipo", label: "Tipo" },
      { key: "marca", label: "Marca" },
      { key: "modelo", label: "Modelo" },
      { key: "serial", label: "Serial" },
      { key: "placa", label: "Placa" },
      ...INVENTORY_COLUMNS,
      { key: "fechaEntrega", label: "Fecha entrega" },
      { key: "valorArriendo", label: "Valor arriendo", type: "currency" }
    ];
  }

  function buildStyledWorksheet(columns, rows) {
    const title = PermissionService.isCommercial(state.user)
      ? `Cartera comercial - ${state.user.comercial}`
      : `${state.datasetLabel} Provexpress - Equipos en arriendo`;
    const generatedAt = new Date().toLocaleDateString("es-CO");

    const data = [
      [title],
      [`Generado: ${generatedAt} | Registros: ${rows.length}`],
      [],
      columns.map((column) => column.label),
      ...rows.map((row) => columns.map((column) => getExportValue(row, column)))
    ];
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const lastColumn = columns.length - 1;
    const lastRow = rows.length + 3;

    worksheet["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: lastColumn } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: lastColumn } }
    ];
    worksheet["!cols"] = columns.map((column) => ({ wch: getExportColumnWidth(column) }));
    worksheet["!autofilter"] = {
      ref: XLSX.utils.encode_range({ s: { r: 3, c: 0 }, e: { r: lastRow, c: lastColumn } })
    };

    styleExportWorksheet(worksheet, columns, rows.length);
    return worksheet;
  }

  function getExportValue(row, column) {
    if (column.type === "currency") {
      return Number(row[column.key]) || 0;
    }

    return row[column.key] || "";
  }

  function styleExportWorksheet(worksheet, columns, rowCount) {
    const lastColumn = columns.length - 1;
    const titleCell = worksheet.A1;
    const subtitleCell = worksheet.A2;
    if (titleCell) {
      titleCell.s = {
        font: { bold: true, color: { rgb: "1A2B6B" }, sz: 18 },
        fill: { fgColor: { rgb: "EEF3FF" } },
        alignment: { horizontal: "center", vertical: "center" }
      };
    }
    if (subtitleCell) {
      subtitleCell.s = {
        font: { color: { rgb: "677592" }, sz: 11 },
        alignment: { horizontal: "center", vertical: "center" }
      };
    }

    for (let columnIndex = 0; columnIndex <= lastColumn; columnIndex += 1) {
      const headerCell = worksheet[XLSX.utils.encode_cell({ r: 3, c: columnIndex })];
      if (headerCell) {
        headerCell.s = {
          font: { bold: true, color: { rgb: "FFFFFF" } },
          fill: { fgColor: { rgb: "1A2B6B" } },
          alignment: { horizontal: "center", vertical: "center" },
          border: getExportBorder()
        };
      }
    }

    for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
      const excelRow = rowIndex + 4;
      const fillColor = rowIndex % 2 ? "F4F7FF" : "FFFFFF";
      columns.forEach((column, columnIndex) => {
        const cell = worksheet[XLSX.utils.encode_cell({ r: excelRow, c: columnIndex })];
        if (!cell) return;
        cell.s = {
          font: { color: { rgb: column.type === "currency" ? "1565C0" : "1A2B6B" }, bold: column.type === "currency" },
          fill: { fgColor: { rgb: fillColor } },
          alignment: { horizontal: column.type === "currency" ? "right" : "left", vertical: "center" },
          border: getExportBorder()
        };
        if (column.type === "currency") {
          cell.t = "n";
          cell.z = '"$"#,##0';
        }
      });
    }
  }

  function getExportColumnWidth(column) {
    const widths = {
      cliente: 30,
      comercial: 26,
      tipo: 14,
      marca: 16,
      modelo: 24,
      serial: 24,
      placa: 16,
      memoria: 14,
      tamanoDisco: 18,
      garantia: 16,
      office: 12,
      morral: 12,
      guaya: 12,
      mouse: 12,
      teclado: 12,
      monitor: 12,
      accesorios: 24,
      fechaEntrega: 16,
      valorArriendo: 18,
      costoRenta: 18,
      utilidadRenta: 18,
      margen: 12
    };
    return widths[column.key] || 16;
  }

  function getExportBorder() {
    return {
      top: { style: "thin", color: { rgb: "D7E0F0" } },
      right: { style: "thin", color: { rgb: "D7E0F0" } },
      bottom: { style: "thin", color: { rgb: "D7E0F0" } },
      left: { style: "thin", color: { rgb: "D7E0F0" } }
    };
  }

  function getExportFileName() {
    const date = new Date().toISOString().slice(0, 10);
    const client = getExportClientName();
    return `renta-px-${client}-${date}.xlsx`;
  }

  function getExportClientName() {
    const clients = Array.from(new Set(state.filteredRows.map((row) => row.cliente).filter(Boolean)));
    if (clients.length === 1) {
      return slugifyFileName(clients[0]);
    }

    const selectedClient = document.getElementById("clientFilter") && document.getElementById("clientFilter").value;
    if (selectedClient) {
      return slugifyFileName(selectedClient);
    }

    return "varios-clientes";
  }

  function slugifyFileName(value) {
    const normalized = ExcelService.comparableText(value).replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
    return normalized || "cliente";
  }

  function statusBadge(status) {
    const map = {
      ok: ["Completo", "badge-ok"],
      warning: ["Revisar", "badge-warning"],
      error: ["Error", "badge-error"]
    };
    const item = map[status] || map.warning;
    return `<span class="badge ${item[1]}">${item[0]}</span>`;
  }

  function fillSelect(id, values) {
    const select = document.getElementById(id);
    const first = select.options[0].outerHTML;
    select.innerHTML = first + values.map((value) => `<option value="${DashboardView.escapeHtml(value)}">${DashboardView.escapeHtml(value)}</option>`).join("");
  }

  function uniqueValues(rows, key) {
    return Array.from(new Set(rows.map((row) => row[key]).filter(Boolean))).sort((a, b) => a.localeCompare(b, "es"));
  }

  function getSortMark(key) {
    if (state.sortKey !== key) return "";
    return state.sortDirection === "asc" ? " ↑" : " ↓";
  }

  window.TableView = {
    renderTableShell,
    initTable
  };
})();
