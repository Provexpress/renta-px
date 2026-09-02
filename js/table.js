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
  const ACCESSORY_EXPORT_ITEMS = [
    { key: "office", label: "Office" },
    { key: "morral", label: "Morral" },
    { key: "guaya", label: "Guaya" },
    { key: "mouse", label: "Mouse" },
    { key: "teclado", label: "Teclado" },
    { key: "monitor", label: "Monitor" }
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
  const AVAILABLE_INVENTORY_COLUMNS = [
    { key: "tipo", label: "Tipo" },
    { key: "marca", label: "Marca" },
    { key: "modelo", label: "Modelo" },
    { key: "serial", label: "Serial" },
    { key: "placa", label: "Placa" },
    { key: "procesador", label: "Procesador" },
    { key: "memoria", label: "Memoria (GB)" },
    { key: "tamanoDisco", label: "Tamaño Disco (GB)" },
    { key: "valorRentaCliente", label: "Valor Renta a Cliente", type: "currency" }
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
    const isAvailable = datasetLabel === "Disponible para Rentar";
    const title = isAvailable
      ? "Inventario de Equipos Disponibles para Ofrecer a Clientes"
      : PermissionService.isCommercial(user)
        ? "Mi tabla de equipos"
        : `Tabla de equipos - ${datasetLabel}`;

    if (isAvailable) {
      return `
        <section class="panel">
          <div class="panel-header-row">
            <h2 class="panel-title">${title}</h2>
            <span class="state-chip">Catálogo Comercial</span>
          </div>
          <div class="table-toolbar table-toolbar-available">
            <input id="tableSearch" type="search" placeholder="Buscar por procesador, modelo, serial, etc.">
            <select id="typeFilter"><option value="">Tipo</option></select>
            <select id="brandFilter"><option value="">Marca</option></select>
            <select id="processorFilter"><option value="">Procesador</option></select>
            <select id="memoryFilter"><option value="">Memoria</option></select>
            <button class="primary-button table-export-button" id="exportTableButton" type="button">Exportar Catálogo Excel</button>
          </div>
          <div id="tableContainer"></div>
        </section>
      `;
    }

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
    const isAvailable = (options.datasetKey === "disponible") || (options.datasetLabel === "Disponible para Rentar");

    state = {
      rows,
      filteredRows: rows,
      user,
      page: 1,
      sortKey: isAvailable ? "tipo" : "cliente",
      sortDirection: "asc",
      datasetLabel: options.datasetLabel || "Renta",
      datasetKey: options.datasetKey || "renta"
    };

    if (isAvailable) {
      fillSelect("typeFilter", uniqueValues(rows, "tipo"));
      fillSelect("brandFilter", uniqueValues(rows, "marca"));
      fillSelect("processorFilter", uniqueValues(rows, "procesador"));
      fillSelect("memoryFilter", uniqueValues(rows, "memoria"));
    } else {
      fillSelect("clientFilter", uniqueValues(rows, "cliente"));
      fillSelect("brandFilter", uniqueValues(rows, "marca"));
      fillSelect("typeFilter", uniqueValues(rows, "tipo"));
    }

    bindFilters(isAvailable);
    bindExport();
    applyFilters();
  }

  function bindFilters(isAvailable) {
    const filterIds = isAvailable
      ? ["tableSearch", "typeFilter", "brandFilter", "processorFilter", "memoryFilter"]
      : ["tableSearch", "clientFilter", "brandFilter", "typeFilter", "statusFilter"];

    filterIds.forEach((id) => {
      const el = document.getElementById(id);
      if (el) {
        el.addEventListener("input", () => {
          state.page = 1;
          applyFilters();
        });
      }
    });
  }

  function bindExport() {
    const button = document.getElementById("exportTableButton");
    if (button) {
      button.addEventListener("click", exportFilteredRows);
    }
  }

  function applyFilters() {
    const isAvailable = state.datasetKey === "disponible";
    const searchEl = document.getElementById("tableSearch");
    const search = ExcelService.comparableText(searchEl ? searchEl.value : "");

    if (isAvailable) {
      const typeEl = document.getElementById("typeFilter");
      const brandEl = document.getElementById("brandFilter");
      const procEl = document.getElementById("processorFilter");
      const memEl = document.getElementById("memoryFilter");

      const type = typeEl ? typeEl.value : "";
      const brand = brandEl ? brandEl.value : "";
      const processor = procEl ? procEl.value : "";
      const memory = memEl ? memEl.value : "";

      state.filteredRows = state.rows.filter((row) => {
        const searchText = ExcelService.comparableText([
          row.tipo,
          row.marca,
          row.modelo,
          row.serial,
          row.placa,
          row.procesador,
          row.memoria,
          row.tamanoDisco,
          row.valorRentaCliente,
          row.valorArriendo
        ].join(" "));

        return (!search || searchText.includes(search))
          && (!type || row.tipo === type)
          && (!brand || row.marca === brand)
          && (!processor || row.procesador === processor)
          && (!memory || row.memoria === memory);
      });
    } else {
      const clientEl = document.getElementById("clientFilter");
      const brandEl = document.getElementById("brandFilter");
      const typeEl = document.getElementById("typeFilter");
      const statusEl = document.getElementById("statusFilter");

      const client = clientEl ? clientEl.value : "";
      const brand = brandEl ? brandEl.value : "";
      const type = typeEl ? typeEl.value : "";
      const status = statusEl ? statusEl.value : "";

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
    }

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
    if (state.datasetKey === "disponible") {
      return AVAILABLE_INVENTORY_COLUMNS;
    }

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
    let value = row[column.key];
    if (column.key === "valorRentaCliente" && (value === undefined || value === 0)) {
      value = row.valorArriendo;
    }
    if (column.type === "currency") return DashboardView.formatCurrency(value);
    if (column.type === "percent") return DashboardView.formatPercent(value);
    if (column.type === "status") return statusBadge(value);
    return DashboardView.escapeHtml(value || "");
  }

  function exportFilteredRows() {
    if (!state.filteredRows.length) return;

    const columns = getCompactExportColumns(getExportColumns(), state.filteredRows);
    const worksheet = buildStyledWorksheet(columns, state.filteredRows);
    const workbook = XLSX.utils.book_new();
    const sheetTitle = state.datasetKey === "disponible" ? "Equipos Disponibles" : "Renta PX";
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetTitle);
    XLSX.writeFile(workbook, getExportFileName());
  }

  function getCompactExportColumns(columns, rows) {
    const compactColumns = columns.filter((column) => hasExportColumnData(column, rows));
    return compactColumns.length ? compactColumns : columns;
  }

  function hasExportColumnData(column, rows) {
    return rows.some((row) => isMeaningfulExportValue(getCleanExportValue(row, column)));
  }

  function isMeaningfulExportValue(value) {
    if (value === null || value === undefined) return false;
    if (typeof value === "number") return true;
    return !["", "no"].includes(ExcelService.comparableText(value));
  }

  function getExportColumns() {
    if (state.datasetKey === "disponible") {
      return AVAILABLE_INVENTORY_COLUMNS;
    }

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
      { key: "memoria", label: "Memoria" },
      { key: "tamanoDisco", label: "Tamaño disco" },
      { key: "garantia", label: "Garantía" },
      { key: "accessorySummary", label: "Accesorios", getValue: getAccessorySummary },
      { key: "fechaEntrega", label: "Fecha entrega" },
      { key: "valorArriendo", label: "Valor arriendo", type: "currency" }
    ];
  }

  function buildStyledWorksheet(columns, rows) {
    let title;
    if (state.datasetKey === "disponible") {
      title = "Provexpress - Catálogo de Equipos Disponibles para Renta";
    } else if (PermissionService.isCommercial(state.user)) {
      title = `Cartera comercial - ${state.user.comercial}`;
    } else {
      title = `${state.datasetLabel} Provexpress - Equipos en arriendo`;
    }
    const generatedAt = new Date().toLocaleDateString("es-CO");

    const data = [
      [title],
      [`Generado: ${generatedAt} | Registros: ${rows.length}`],
      [],
      columns.map((column) => column.label),
      ...rows.map((row) => columns.map((column) => getCleanExportValue(row, column)))
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
    if (typeof column.getValue === "function") {
      return column.getValue(row);
    }

    if (column.type === "currency") {
      return Number(row[column.key]) || 0;
    }

    return row[column.key] || "";
  }

  function getAccessorySummary(row) {
    const items = ACCESSORY_EXPORT_ITEMS
      .filter((item) => ExcelService.comparableText(row[item.key]) === "si")
      .map((item) => item.label);
    const extra = row.accesorios || "";
    const extraKey = ExcelService.comparableText(extra);

    if (extraKey && !["si", "no"].includes(extraKey) && !items.includes(extra)) {
      items.push(extra);
    }

    return items.join(", ");
  }

  function getCleanExportValue(row, column) {
    const value = getExportValue(row, column);
    if (typeof value === "number") return value;
    return ExcelService.comparableText(value) === "no" ? "" : value;
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
      tipo: 16,
      marca: 16,
      modelo: 24,
      serial: 24,
      placa: 16,
      procesador: 24,
      memoria: 16,
      tamanoDisco: 20,
      valorRentaCliente: 22,
      garantia: 16,
      office: 12,
      morral: 12,
      guaya: 12,
      mouse: 12,
      teclado: 12,
      monitor: 12,
      accesorios: 24,
      accessorySummary: 28,
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
    if (state.datasetKey === "disponible") {
      return `catalogo-equipos-disponibles-provexpress-${date}.xlsx`;
    }
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
