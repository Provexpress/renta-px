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

    const columns = getColumns(state.user).filter((column) => column.type !== "status");
    const exportRows = state.filteredRows.map((row) => {
      return columns.reduce((record, column) => {
        record[column.label] = formatExportValue(row, column);
        return record;
      }, {});
    });

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    worksheet["!cols"] = columns.map((column) => ({
      wch: getColumnWidth(column)
    }));
    worksheet["!autofilter"] = {
      ref: XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: exportRows.length, c: columns.length - 1 }
      })
    };
    worksheet["!tables"] = [{
      name: "TablaRentaPX",
      ref: worksheet["!autofilter"].ref,
      headerRow: true,
      totalsRow: false,
      style: {
        theme: "TableStyleMedium2",
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false
      }
    }];
    worksheet["!freeze"] = { xSplit: 0, ySplit: 1 };
    applyWorksheetFormats(worksheet, columns, exportRows.length);

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Renta PX");
    XLSX.writeFile(workbook, getExportFileName());
  }

  function formatExportValue(row, column) {
    const value = row[column.key];
    if (column.type === "percent") return Number(value || 0);
    if (column.type === "currency") return Number(value || 0);
    return value || "";
  }

  function getColumnWidth(column) {
    const widths = {
      cliente: 30,
      comercial: 26,
      tipo: 16,
      marca: 16,
      modelo: 24,
      serial: 24,
      placa: 16,
      fechaEntrega: 16,
      valorArriendo: 18,
      costoRenta: 18,
      utilidadRenta: 18,
      margen: 12
    };

    return widths[column.key] || Math.max(14, column.label.length + 4);
  }

  function applyWorksheetFormats(worksheet, columns, rowCount) {
    columns.forEach((column, index) => {
      const headerCell = XLSX.utils.encode_cell({ r: 0, c: index });
      if (worksheet[headerCell]) {
        worksheet[headerCell].s = {
          font: { bold: true, color: { rgb: "FFFFFF" } },
          fill: { fgColor: { rgb: "1A2B6B" } },
          alignment: { horizontal: "center", vertical: "center" },
          border: getThinBorder()
        };
      }

      for (let rowIndex = 1; rowIndex <= rowCount; rowIndex += 1) {
        const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: index });
        const cell = worksheet[cellAddress];
        if (!cell) continue;

        cell.s = {
          alignment: { vertical: "center" },
          border: getThinBorder()
        };

        if (column.type === "currency") {
          cell.t = "n";
          cell.z = '"$"#,##0';
        }

        if (column.type === "percent") {
          cell.t = "n";
          cell.z = "0.0%";
        }
      }
    });
  }

  function getThinBorder() {
    return {
      top: { style: "thin", color: { rgb: "D7E0F0" } },
      right: { style: "thin", color: { rgb: "D7E0F0" } },
      bottom: { style: "thin", color: { rgb: "D7E0F0" } },
      left: { style: "thin", color: { rgb: "D7E0F0" } }
    };
  }

  function getExportFileName() {
    const date = new Date().toISOString().slice(0, 10);
    const role = state.user.role === "comercial" ? ExcelService.comparableText(state.user.comercial).replace(/\s+/g, "-") : state.user.role;
    return `renta-px-${role}-${date}.xlsx`;
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
