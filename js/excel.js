(function () {
  const COLUMN_ALIASES = {
    tipo: "tipo",
    marca: "marca",
    marc: "marca",
    modelo: "modelo",
    serial: "serial",
    placa: "placa",
    procesador: "procesador",
    procesado: "procesador",
    proc: "procesador",
    cpu: "procesador",
    memoria: "memoria",
    memoriagb: "memoria",
    memoriaengb: "memoria",
    ram: "memoria",
    ramgb: "memoria",
    tamanodisco: "tamanoDisco",
    tamanodedisco: "tamanoDisco",
    tamanodiscogb: "tamanoDisco",
    tamanodediscogb: "tamanoDisco",
    tamanodiscoengb: "tamanoDisco",
    discoduro: "tamanoDisco",
    discodurogb: "tamanoDisco",
    disco: "tamanoDisco",
    discogb: "tamanoDisco",
    almacenamiento: "tamanoDisco",
    almacenamientogb: "tamanoDisco",
    garantia: "garantia",
    office: "office",
    morral: "morral",
    guaya: "guaya",
    mouse: "mouse",
    teclado: "teclado",
    monitor: "monitor",
    accesorios: "accesorios",
    comercial: "comercial",
    cliente: "cliente",
    valorarriendo: "valorArriendo",
    canon: "valorArriendo",
    canonmensual: "valorArriendo",
    costorenta: "costoRenta",
    costoderenta: "costoRenta",
    utilidadrenta: "utilidadRenta",
    utilidadderenta: "utilidadRenta",
    margen: "margen",
    mesesarrendado: "mesesArrendado",
    mesesarrendados: "mesesArrendado",
    fechaentrega: "fechaEntrega",
    fechadeentrega: "fechaEntrega",
    valorrentaacliente: "valorRentaCliente",
    valorrentacliente: "valorRentaCliente",
    valorrenta: "valorRentaCliente",
    valorarriendoacliente: "valorRentaCliente",
    valorarriendocliente: "valorRentaCliente",
    rentacliente: "valorRentaCliente",
    rentaacliente: "valorRentaCliente",
    canoncliente: "valorRentaCliente",
    canonacliente: "valorRentaCliente",
    preciorenta: "valorRentaCliente",
    preciocliente: "valorRentaCliente",
    preciodecliente: "valorRentaCliente"
  };

  const ACCESSORY_KEYS = ["office", "morral", "guaya", "mouse", "teclado", "monitor"];

  function readExcelArrayBuffer(arrayBuffer, requestedSheetName) {
    try {
      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
        raw: true
      });
      const configuredSheet = requestedSheetName || (APP_CONFIG.graph && APP_CONFIG.graph.sharePointFile && APP_CONFIG.graph.sharePointFile.sheetName);
      const selectedSheetName = getWorkbookSheetName(workbook, configuredSheet) || workbook.SheetNames[0];

      const sheet = workbook.Sheets[selectedSheetName];

      if (!sheet) {
        throw new Error(`No se encontró la hoja "${selectedSheetName}" en el Excel.`);
      }

      // Obtener matriz bidimensional de celdas para encontrar dinámicamente la fila de encabezados
      const rawRowsAoa = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: true,
        blankrows: false
      });

      if (!rawRowsAoa || !rawRowsAoa.length) {
        return [];
      }

      const headerRowIndex = findHeaderRowIndex(rawRowsAoa);
      const headerRow = rawRowsAoa[headerRowIndex] || [];
      const dataRows = rawRowsAoa.slice(headerRowIndex + 1);

      return dataRows
        .filter((rowArr) => Array.isArray(rowArr) && rowArr.some((cell) => cell !== null && cell !== undefined && String(cell).trim() !== ""))
        .map((rowArr) => {
          const rowObj = {};
          headerRow.forEach((colName, colIdx) => {
            const rawColName = String(colName || "").trim();
            if (rawColName) {
              rowObj[rawColName] = rowArr[colIdx] !== undefined ? rowArr[colIdx] : "";
            } else if (rowArr[colIdx] !== undefined && rowArr[colIdx] !== "") {
              rowObj[`__EMPTY_${colIdx}`] = rowArr[colIdx];
            }
          });
          return rowObj;
        });
    } catch (error) {
      const readError = new Error("No se pudo leer el Excel.");
      readError.cause = error;
      throw readError;
    }
  }

  function findHeaderRowIndex(rowsAoa) {
    const knownHeaderTokens = [
      "tipo", "marca", "marc", "modelo", "serial", "placa", "procesador",
      "memoria", "disco", "tamano", "cliente", "comercial", "canon", "valor",
      "renta", "arriendo", "costo", "utilidad", "garantia", "fecha"
    ];

    let bestIndex = 0;
    let maxMatches = 0;

    for (let i = 0; i < Math.min(rowsAoa.length, 30); i++) {
      const row = rowsAoa[i];
      if (!Array.isArray(row)) continue;

      let matchCount = 0;
      row.forEach((cell) => {
        const text = comparableText(cell);
        if (text && knownHeaderTokens.some((token) => text.includes(token))) {
          matchCount += 1;
        }
      });

      if (matchCount > maxMatches && matchCount >= 2) {
        maxMatches = matchCount;
        bestIndex = i;
      }
    }

    return bestIndex;
  }

  function getWorkbookSheetName(workbook, sheetName) {
    if (!sheetName) return workbook.SheetNames[0] || "";
    if (workbook.SheetNames.includes(sheetName)) return sheetName;

    const requestedName = comparableText(sheetName);
    const exact = workbook.SheetNames.find((name) => comparableText(name) === requestedName);
    if (exact) return exact;

    const partial = workbook.SheetNames.find((name) => {
      const c = comparableText(name);
      return c.includes(requestedName) || requestedName.includes(c);
    });
    if (partial) return partial;

    return workbook.SheetNames[0] || "";
  }

  function normalizeRows(rows) {
    return rows
      .map((row, index) => ({ row, index }))
      .filter(({ row }) => hasExcelData(row))
      .map(({ row, index }) => normalizeRow(row, index))
      .filter((row) => hasRealInformation(row));
  }

  function hasExcelData(row) {
    return Object.entries(row).some(([key, value]) => {
      if (key.startsWith("__EMPTY") && (value === null || value === undefined || normalizeText(value) === "")) {
        return false;
      }
      if (value === null || value === undefined) return false;
      if (value instanceof Date) return !Number.isNaN(value.getTime());
      return normalizeText(value) !== "";
    });
  }

  function hasRealInformation(row) {
    const mainFields = [
      row.cliente,
      row.comercial,
      row.serial,
      row.placa,
      row.tipo,
      row.marca,
      row.modelo,
      row.procesador,
      row.accesorios
    ];

    const hasIdentifyingText = mainFields.some((text) => {
      const normalized = normalizeText(text);
      if (!normalized) return false;
      const lower = comparableText(normalized);
      if (lower.startsWith("total") || lower === "grand total" || lower === "suma") {
        return false;
      }
      return true;
    });

    if (hasIdentifyingText) {
      return true;
    }

    return Boolean(row.valorArriendo > 0 || row.costoRenta > 0 || row.utilidadRenta > 0);
  }

  function normalizeRow(source, index) {
    const row = { rowNumber: index + 2 };

    Object.entries(source).forEach(([key, value]) => {
      row[normalizeKey(key)] = value;
    });

    row.tipo = normalizeText(row.tipo);
    row.marca = normalizeText(row.marca);
    row.modelo = normalizeText(row.modelo);
    row.serial = normalizeText(row.serial);
    row.placa = normalizeText(row.placa);
    row.procesador = normalizeText(row.procesador);
    row.memoria = normalizeText(row.memoria);
    row.tamanoDisco = normalizeText(row.tamanoDisco);
    row.garantia = normalizeText(row.garantia);
    row.office = normalizeText(row.office);
    row.morral = normalizeText(row.morral);
    row.guaya = normalizeText(row.guaya);
    row.mouse = normalizeText(row.mouse);
    row.teclado = normalizeText(row.teclado);
    row.monitor = normalizeText(row.monitor);
    row.accesorios = normalizeText(row.accesorios);
    row.comercial = normalizeText(row.comercial);
    row.cliente = normalizeText(row.cliente);
    row.valorArriendo = parseMoney(row.valorArriendo);
    row.costoRenta = parseMoney(row.costoRenta);
    row.utilidadRenta = parseMoney(row.utilidadRenta);
    row.mesesArrendado = parseNumber(row.mesesArrendado);
    row.fechaEntrega = parseExcelDate(row.fechaEntrega);

    if (!row.utilidadRenta && row.valorArriendo && row.costoRenta) {
      row.utilidadRenta = row.valorArriendo - row.costoRenta;
    }

    row.margen = calculateMargin(row);
    row.validations = validateRow(row);
    row.dataQualityStatus = getDataQualityStatus(row);
    return row;
  }

  function normalizeKey(key) {
    const compact = removeAccents(String(key || ""))
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "");

    if (COLUMN_ALIASES[compact]) {
      return COLUMN_ALIASES[compact];
    }

    return compact.replace(/_([a-z])/g, (_, letter) => letter.toUpperCase());
  }

  function removeAccents(value) {
    return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  }

  function normalizeText(value) {
    if (value === null || value === undefined) return "";
    return String(value).replace(/\s+/g, " ").trim();
  }

  function comparableText(value) {
    return removeAccents(normalizeText(value)).toLowerCase();
  }

  function parseMoney(value) {
    if (typeof value === "number") return Number.isFinite(value) ? value : 0;
    const text = normalizeText(value);
    if (!text) return 0;

    const cleaned = text
      .replace(/\$/g, "")
      .replace(/COP/gi, "")
      .replace(/\s/g, "");

    if (cleaned.includes(",") && cleaned.includes(".")) {
      return Number(cleaned.replace(/\./g, "").replace(",", ".")) || 0;
    }

    if (cleaned.includes(",")) {
      const parts = cleaned.split(",");
      if (parts.length > 1 && parts[parts.length - 1].length === 3) {
        return Number(cleaned.replace(/,/g, "")) || 0;
      }
      return Number(cleaned.replace(",", ".")) || 0;
    }

    if (cleaned.includes(".")) {
      const parts = cleaned.split(".");
      if (parts.length > 1 && parts[parts.length - 1].length === 3) {
        return Number(cleaned.replace(/\./g, "")) || 0;
      }
    }

    return Number(cleaned.replace(/[^\d.-]/g, "")) || 0;
  }

  function parseNumber(value) {
    if (typeof value === "number") return Number.isFinite(value) ? value : 0;
    return Number(normalizeText(value).replace(",", ".")) || 0;
  }

  function parseExcelDate(value) {
    if (!value) return "";
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return value.toISOString().slice(0, 10);
    }

    if (typeof value === "number") {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (!parsed) return "";
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }

    const text = normalizeText(value);
    const date = new Date(text);
    if (!Number.isNaN(date.getTime())) {
      return date.toISOString().slice(0, 10);
    }

    return text;
  }

  function calculateMargin(row) {
    if (!row.valorArriendo) return 0;
    return row.utilidadRenta / row.valorArriendo;
  }

  function validateRow(row) {
    const issues = [];
    if (!row.placa) issues.push("sinPlaca");
    if (!row.fechaEntrega) issues.push("sinFechaEntrega");
    if (!row.valorArriendo) issues.push("sinValorArriendo");
    if (!row.costoRenta) issues.push("sinCostoRenta");

    const expectedProfit = row.valorArriendo - row.costoRenta;
    if (row.valorArriendo && row.costoRenta && Math.abs(row.utilidadRenta - expectedProfit) > 100) {
      issues.push("utilidadInconsistente");
    }

    if (ACCESSORY_KEYS.some((key) => !row[key])) {
      issues.push("accesoriosSinDato");
    }

    return issues;
  }

  function getDataQualityStatus(row) {
    if (row.validations.includes("sinValorArriendo") || row.validations.includes("sinCostoRenta") || row.validations.includes("utilidadInconsistente")) {
      return "error";
    }

    if (row.validations.includes("sinPlaca") || row.validations.includes("sinFechaEntrega") || row.validations.includes("accesoriosSinDato")) {
      return "warning";
    }

    return "ok";
  }

  function normalizeAvailableInventoryRows(rows) {
    return rows
      .map((row, index) => ({ row, index }))
      .filter(({ row }) => hasExcelData(row))
      .map(({ row, index }) => normalizeAvailableInventoryRow(row, index))
      .filter((row) => hasAvailableInventoryData(row));
  }

  function normalizeAvailableInventoryRow(source, index) {
    const row = { rowNumber: index + 2 };

    Object.entries(source).forEach(([key, value]) => {
      row[normalizeKey(key)] = value;
    });

    row.tipo = normalizeText(row.tipo);
    row.marca = normalizeText(row.marca);
    row.modelo = normalizeText(row.modelo);
    row.serial = normalizeText(row.serial);
    row.placa = normalizeText(row.placa);
    row.procesador = normalizeText(row.procesador);
    row.memoria = normalizeText(row.memoria);
    row.tamanoDisco = normalizeText(row.tamanoDisco);
    
    const moneyVal = row.valorRentaCliente !== undefined && row.valorRentaCliente !== ""
      ? row.valorRentaCliente
      : row.valorArriendo;
    row.valorRentaCliente = parseMoney(moneyVal);
    row.valorArriendo = row.valorRentaCliente;

    return row;
  }

  function hasAvailableInventoryData(row) {
    const fields = [row.tipo, row.marca, row.modelo, row.serial, row.placa, row.procesador];
    return fields.some((text) => {
      const normalized = normalizeText(text);
      if (!normalized) return false;
      const lower = comparableText(normalized);
      if (lower.startsWith("total") || lower === "grand total" || lower === "suma") return false;
      return true;
    }) || Boolean(row.valorRentaCliente > 0);
  }

  window.ExcelService = {
    readExcelArrayBuffer,
    normalizeRows,
    normalizeAvailableInventoryRows,
    normalizeKey,
    normalizeText,
    comparableText,
    parseMoney,
    parseExcelDate,
    calculateMargin,
    validateRow,
    getDataQualityStatus,
    hasRealInformation
  };
})();
