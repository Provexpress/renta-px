(function () {
  const PAGE_SIZE = 25;
  let state = {
    rows: [],
    filteredRows: [],
    user: null,
    page: 1,
    sortKey: "cliente",
    sortDirection: "asc"
  };

  function renderTableShell(user) {
    return `
      <section class="panel">
        <h2 class="panel-title">${PermissionService.isCommercial(user) ? "Mi tabla de equipos" : "Tabla de equipos"}</h2>
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

  function initTable(user, rows) {
    state = {
      rows,
      filteredRows: rows,
      user,
      page: 1,
      sortKey: "cliente",
      sortDirection: "asc"
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
        row.placa
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

    const columns = getCommercialExportColumns();
    const html = buildStyledExcelHtml(columns, state.filteredRows);
    const blob = new Blob([html], {
      type: "application/vnd.ms-excel;charset=utf-8"
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = getExportFileName();
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
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
      { key: "fechaEntrega", label: "Fecha entrega" },
      { key: "valorArriendo", label: "Valor arriendo", type: "currency" }
    ];
  }

  function buildStyledExcelHtml(columns, rows) {
    const title = PermissionService.isCommercial(state.user)
      ? `Cartera comercial - ${state.user.comercial}`
      : "Renta PX - Equipos filtrados";
    const generatedAt = new Date().toLocaleDateString("es-CO");
    const headerCells = columns.map((column) => `<th>${escapeExcelHtml(column.label)}</th>`).join("");
    const bodyRows = rows.map((row, index) => {
      const cells = columns.map((column) => {
        const value = getExportDisplayValue(row, column);
        const className = column.type === "currency" ? "money" : "";
        return `<td class="${className}">${escapeExcelHtml(value)}</td>`;
      }).join("");
      return `<tr class="${index % 2 ? "even" : "odd"}">${cells}</tr>`;
    }).join("");

    return `
      <html>
        <head>
          <meta charset="utf-8">
          <style>
            body { font-family: Segoe UI, Arial, sans-serif; color: #1A2B6B; }
            .title { font-size: 20px; font-weight: 700; color: #1A2B6B; }
            .subtitle { font-size: 12px; color: #677592; }
            table { border-collapse: collapse; width: 100%; }
            th {
              background: #1A2B6B;
              color: #FFFFFF;
              font-weight: 700;
              border: 1px solid #1A2B6B;
              padding: 8px;
              text-align: left;
            }
            td {
              border: 1px solid #D7E0F0;
              padding: 7px;
              color: #1A2B6B;
              mso-number-format: "\\@";
            }
            tr.odd td { background: #FFFFFF; }
            tr.even td { background: #F4F7FF; }
            td.money {
              color: #1565C0;
              font-weight: 700;
              text-align: right;
            }
          </style>
        </head>
        <body>
          <table>
            <tr><td colspan="${columns.length}" class="title">${escapeExcelHtml(title)}</td></tr>
            <tr><td colspan="${columns.length}" class="subtitle">Generado: ${escapeExcelHtml(generatedAt)} | Registros: ${rows.length}</td></tr>
            <tr><td colspan="${columns.length}"></td></tr>
            <tr>${headerCells}</tr>
            ${bodyRows}
          </table>
        </body>
      </html>
    `;
  }

  function getExportDisplayValue(row, column) {
    if (column.type === "currency") {
      return DashboardView.formatCurrency(row[column.key]);
    }

    return row[column.key] || "";
  }

  function escapeExcelHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function getExportFileName() {
    const date = new Date().toISOString().slice(0, 10);
    const role = state.user.role === "comercial" ? ExcelService.comparableText(state.user.comercial).replace(/\s+/g, "-") : state.user.role;
    return `renta-px-${role}-${date}.xls`;
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
