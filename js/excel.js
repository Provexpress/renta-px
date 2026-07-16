(function () {
  const COLUMN_ALIASES = {
    tipo: "tipo",
    marca: "marca",
    modelo: "modelo",
    serial: "serial",
    placa: "placa",
    procesador: "procesador",
    memoria: "memoria",
    tamanodisco: "tamanoDisco",
    tamanodedisco: "tamanoDisco",
    discoduro: "tamanoDisco",
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
    fechadeentrega: "fechaEntrega"
  };

  const ACCESSORY_KEYS = ["office", "morral", "guaya", "mouse", "teclado", "monitor"];

  function readExcelArrayBuffer(arrayBuffer, requestedSheetName) {
    try {
      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
        raw: true
      });
      const configuredSheet = requestedSheetName || APP_CONFIG.graph.sharePointFile.sheetName;
      if (configuredSheet && !workbook.SheetNames.includes(configuredSheet)) {
        throw new Error(`No se encontró la hoja "${configuredSheet}" en el Excel.`);
      }

      const sheetName = configuredSheet || workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      if (!sheet) {
        throw new Error("No se encontró una hoja legible en el Excel.");
      }

      return XLSX.utils.sheet_to_json(sheet, {
        defval: "",
        raw: true
      });
    } catch (error) {
      const readError = new Error("No se pudo leer el Excel.");
      readError.cause = error;
      throw readError;
    }
  }

  function normalizeRows(rows) {
    return rows
      .map((row, index) => ({ row, index }))
      .filter(({ row }) => hasExcelData(row))
      .map(({ row, index }) => normalizeRow(row, index));
  }

  function hasExcelData(row) {
    return Object.values(row).some((value) => {
      if (value === null || value === undefined) return false;
      if (value instanceof Date) return !Number.isNaN(value.getTime());
      return normalizeText(value) !== "";
    });
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

  window.ExcelService = {
    readExcelArrayBuffer,
    normalizeRows,
    normalizeKey,
    normalizeText,
    comparableText,
    parseMoney,
    parseExcelDate,
    calculateMargin,
    validateRow,
    getDataQualityStatus
  };
})();
