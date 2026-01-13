const state = {
  nomenclature: [],
  categoryIds: new Map(),
  firstSales: new Map(),
  shelfDates: new Map(),
  salesMonths: {},
  settings: {
    excludedCategories: [],
    excludedSkus: [],
  },
};

const elements = {
  nomenclatureFile: document.getElementById("nomenclatureFile"),
  categoryIdFile: document.getElementById("categoryIdFile"),
  firstSalesFile: document.getElementById("firstSalesFile"),
  shelfDatesFile: document.getElementById("shelfDatesFile"),
  salesFile: document.getElementById("salesFile"),
  salesMonth: document.getElementById("salesMonth"),
  addSales: document.getElementById("addSales"),
  importStatus: document.getElementById("importStatus"),
  excludedCategories: document.getElementById("excludedCategories"),
  excludedSkus: document.getElementById("excludedSkus"),
  saveSettings: document.getElementById("saveSettings"),
  nomenclatureTable: document.getElementById("nomenclatureTable"),
  nomenclatureSearch: document.getElementById("nomenclatureSearch"),
  nomenclatureCount: document.getElementById("nomenclatureCount"),
  abcTable: document.getElementById("abcTable"),
  abcSearch: document.getElementById("abcSearch"),
  abcCount: document.getElementById("abcCount"),
  monthSummaryTable: document.getElementById("monthSummaryTable"),
  resetData: document.getElementById("resetData"),
};

const DATA_KEY = "analysis-app-state";

const formatNumber = (value, fractionDigits = 2) => {
  const num = Number(value);
  if (Number.isNaN(num)) return "";
  return num.toLocaleString("uk-UA", {
    minimumFractionDigits: fractionDigits,
    maximumFractionDigits: fractionDigits,
  });
};

const formatPercent = (value, fractionDigits = 2) => {
  if (value === "" || value === null || typeof value === "undefined") return "";
  const num = Number(value);
  if (Number.isNaN(num)) return "";
  return `${(num * 100).toFixed(fractionDigits)}%`;
};

const parseDate = (value) => {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const raw = String(value).trim();
  if (!raw) return null;
  const parts = raw.split(/[.\/\-]/).map((p) => p.trim());
  if (parts.length === 3 && parts[0].length <= 2) {
    const [day, month, year] = parts.map(Number);
    const date = new Date(year, month - 1, day);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const daysBetween = (start, end) => {
  if (!start || !end) return "";
  const diff = end.getTime() - start.getTime();
  return Math.floor(diff / (1000 * 60 * 60 * 24));
};

const cleanCategory = (value) => {
  if (!value) return "";
  return String(value)
    .replace(/^[\d.\)]*\s*/g, "")
    .replace(/:$/g, "")
    .trim();
};

const splitLines = (value) =>
  value
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

const detectDelimiter = (text) => {
  const sample = text.split(/\r?\n/).slice(0, 5).join("\n");
  const commaCount = (sample.match(/,/g) || []).length;
  const semicolonCount = (sample.match(/;/g) || []).length;
  const tabCount = (sample.match(/\t/g) || []).length;
  if (tabCount > commaCount && tabCount > semicolonCount) return "\t";
  if (semicolonCount > commaCount) return ";";
  return ",";
};

const parseCSV = (text) => {
  const delimiter = detectDelimiter(text);
  const rows = [];
  let current = [];
  let field = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];
    if (char === '"') {
      if (inQuotes && next === '"') {
        field += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === delimiter && !inQuotes) {
      current.push(field);
      field = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") i += 1;
      current.push(field);
      if (current.length > 1 || current[0] !== "") rows.push(current);
      current = [];
      field = "";
    } else {
      field += char;
    }
  }
  if (field.length || current.length) {
    current.push(field);
    if (current.length > 1 || current[0] !== "") rows.push(current);
  }

  const [headerRow, ...dataRows] = rows;
  const headers = headerRow.map((h) => h.trim());
  const data = dataRows.map((row) => {
    const entry = {};
    headers.forEach((key, index) => {
      entry[key] = row[index] ? row[index].trim() : "";
    });
    return entry;
  });
  return { headers, data };
};

const storeState = () => {
  const payload = {
    nomenclature: state.nomenclature,
    categoryIds: Array.from(state.categoryIds.entries()),
    firstSales: Array.from(state.firstSales.entries()),
    shelfDates: Array.from(state.shelfDates.entries()),
    salesMonths: state.salesMonths,
    settings: state.settings,
  };
  localStorage.setItem(DATA_KEY, JSON.stringify(payload));
};

const restoreState = () => {
  const raw = localStorage.getItem(DATA_KEY);
  if (!raw) return;
  const payload = JSON.parse(raw);
  state.nomenclature = payload.nomenclature || [];
  state.categoryIds = new Map(payload.categoryIds || []);
  state.firstSales = new Map(payload.firstSales || []);
  state.shelfDates = new Map(payload.shelfDates || []);
  state.salesMonths = payload.salesMonths || {};
  state.settings = payload.settings || state.settings;
  elements.excludedCategories.value = state.settings.excludedCategories.join("\n");
  elements.excludedSkus.value = state.settings.excludedSkus.join("\n");
};

const updateStatus = (message, tone = "ok") => {
  elements.importStatus.textContent = message;
  elements.importStatus.className = `status ${tone}`;
};

const readFileAsText = (file) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsText(file);
  });

const readFileAsArrayBuffer = (file) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });

const parseRowsToObjects = (rows) => {
  if (!rows.length) return { headers: [], data: [] };
  const headers = rows[0].map((cell) => String(cell || "").trim());
  const data = rows.slice(1).map((row) => {
    const entry = {};
    headers.forEach((key, index) => {
      entry[key] = row[index] ? String(row[index]).trim() : "";
    });
    return entry;
  });
  return { headers, data };
};

const parseXmlRows = (text) => {
  const parser = new DOMParser();
  const xml = parser.parseFromString(text, "text/xml");
  const rowNodes = Array.from(xml.getElementsByTagName("Row"));
  if (!rowNodes.length) {
    const genericRows = Array.from(xml.getElementsByTagName("row"));
    return genericRows.map((row) =>
      Array.from(row.getElementsByTagName("cell")).map((cell) => cell.textContent || "")
    );
  }
  return rowNodes.map((row) => {
    const cells = Array.from(row.getElementsByTagName("Cell"));
    if (cells.length) {
      return cells.map((cell) => {
        const dataNode = cell.getElementsByTagName("Data")[0];
        return dataNode ? dataNode.textContent || "" : cell.textContent || "";
      });
    }
    return Array.from(row.children).map((cell) => cell.textContent || "");
  });
};

const parseFile = async (file) => {
  const ext = file.name.split(".").pop().toLowerCase();
  if (ext === "xlsx" || ext === "xls") {
    const buffer = await readFileAsArrayBuffer(file);
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
    return parseRowsToObjects(rows);
  }
  if (ext === "xml") {
    const text = await readFileAsText(file);
    const rows = parseXmlRows(text);
    return parseRowsToObjects(rows);
  }
  const text = await readFileAsText(file);
  return parseCSV(text);
};

const filterNomenclature = (items) => {
  const excludedCats = new Set(state.settings.excludedCategories.map(cleanCategory));
  const excludedSkus = new Set(state.settings.excludedSkus.map((sku) => sku.trim()));
  return items.filter((item) => {
    if (excludedSkus.has(item.sku)) return false;
    const cleanedCategory = cleanCategory(item.category);
    if (excludedCats.has(cleanedCategory)) return false;
    return true;
  });
};

const normalizeNomenclature = (rows) => {
  const normalized = rows.map((row) => {
    const sku = row["Артикул"] || row["SKU"] || row["Артикул "] || row["Артикул товара"] || "";
    return {
      sku: String(sku).trim(),
      code: row["Код"] || "",
      name: row["Ім'я"] || row["Имя"] || row["Назва"] || "",
      category: row["Категорія"] || row["Категория"] || "",
      purchasePrice: Number(String(row["Ціна закупівлі"] || row["Цена закупки"] || 0).replace(",", ".")) || 0,
      salePrice: Number(String(row["Ціна продажу"] || row["Цена продажи"] || 0).replace(",", ".")) || 0,
      stockQty: Number(String(row["Кількість"] || row["Количество"] || 0).replace(",", ".")) || 0,
      barcode: row["Штрихкод"] || "",
    };
  });
  return filterNomenclature(normalized).filter((item) => item.sku);
};

const computeSalesBySku = () => {
  const skuMap = new Map();
  Object.entries(state.salesMonths).forEach(([month, rows]) => {
    rows.forEach((row) => {
      const sku = row.sku;
      if (!sku) return;
      if (!skuMap.has(sku)) {
        skuMap.set(sku, { qty: 0, revenue: 0, months: {} });
      }
      const entry = skuMap.get(sku);
      entry.qty += row.qty;
      entry.revenue += row.revenue;
      entry.months[month] = (entry.months[month] || 0) + row.qty;
    });
  });
  return skuMap;
};

const computeNomenclatureMetrics = () => {
  const salesBySku = computeSalesBySku();
  const monthsCount = Object.keys(state.salesMonths).length;
  const weeks = monthsCount ? monthsCount * 4.333 : 0;
  const today = new Date();

  return state.nomenclature.map((item) => {
    const cleanCat = cleanCategory(item.category);
    const idCategory = state.categoryIds.get(cleanCat) || "ID не знайдено";
    const firstSale = state.firstSales.get(item.sku) || null;
    const shelfDate = state.shelfDates.get(item.sku) || null;
    const ageDays = shelfDate ? daysBetween(shelfDate, today) : "";
    const daysToFirstSale = shelfDate && firstSale ? daysBetween(shelfDate, firstSale) : "";
    let agingBucket = "";
    if (ageDays !== "") {
      if (ageDays <= 30) agingBucket = "≤30";
      else if (ageDays <= 90) agingBucket = "31-90";
      else if (ageDays <= 180) agingBucket = "91-180";
      else agingBucket = ">180";
    }
    const sales = salesBySku.get(item.sku) || { qty: 0, revenue: 0 };
    const weeklySales = weeks ? sales.qty / weeks : 0;
    const stockWeeks = weeklySales ? item.stockQty / weeklySales : "";
    const deadStock = ageDays !== "" && ageDays > 180 && !firstSale ? "⚠️" : "";
    const endDate = firstSale || today;
    let activeWeeks = shelfDate ? Math.round(daysBetween(shelfDate, endDate) / 7) : "";
    if (activeWeeks !== "" && activeWeeks < 1) activeWeeks = 1;

    return {
      ...item,
      cleanCategory: cleanCat,
      idCategory,
      firstSale,
      shelfDate,
      ageDays,
      daysToFirstSale,
      agingBucket,
      weeklySales,
      stockWeeks,
      deadStock,
      activeWeeks,
    };
  });
};

const computeAbcXyz = () => {
  const salesBySku = computeSalesBySku();
  const rows = state.nomenclature.map((item) => {
    const sales = salesBySku.get(item.sku) || { qty: 0, revenue: 0, months: {} };
    const margin = sales.revenue - item.purchasePrice * sales.qty;
    const marginPct = item.salePrice ? (item.salePrice - item.purchasePrice) / item.salePrice : 0;
    const markup = item.purchasePrice ? (item.salePrice - item.purchasePrice) / item.purchasePrice : 0;
    return {
      sku: item.sku,
      name: item.name,
      category: cleanCategory(item.category),
      purchasePrice: item.purchasePrice,
      salePrice: item.salePrice,
      stockQty: item.stockQty,
      revenue: sales.revenue,
      qty: sales.qty,
      margin,
      marginPct,
      markup,
      months: sales.months || {},
    };
  });

  const totals = rows.reduce(
    (acc, row) => {
      acc.revenue += row.revenue;
      acc.margin += row.margin;
      acc.qty += row.qty;
      return acc;
    },
    { revenue: 0, margin: 0, qty: 0 }
  );

  const rankBy = (key, totalKey, label) => {
    const sorted = [...rows].sort((a, b) => b[key] - a[key]);
    let cumulative = 0;
    sorted.forEach((row) => {
      const value = row[key];
      const share = totals[totalKey] ? value / totals[totalKey] : 0;
      cumulative += share;
      let cls = "C";
      if (cumulative <= 0.8) cls = "A";
      else if (cumulative <= 0.95) cls = "B";
      row[label] = cls;
    });
  };

  rankBy("revenue", "revenue", "abcRevenue");
  rankBy("margin", "margin", "abcMargin");
  rankBy("qty", "qty", "abcQty");

  rows.forEach((row) => {
    const monthValues = Object.values(row.months || {});
    const avgQty = monthValues.length
      ? monthValues.reduce((sum, value) => sum + value, 0) / monthValues.length
      : 0;
    const variance =
      monthValues.length
        ? monthValues.reduce((sum, value) => sum + Math.pow(value - avgQty, 2), 0) /
          monthValues.length
        : 0;
    const stdDev = Math.sqrt(variance);
    const cv = avgQty > 0 ? stdDev / avgQty : 0;
    let xyz = "Z";
    if (row.qty === 0) {
      xyz = "Z";
    } else if (cv <= 0.35) {
      xyz = "X";
    } else if (cv <= 0.7) {
      xyz = "Y";
    }
    row.xyz = xyz;
    row.abcXyzRevenue = `${row.abcRevenue}${xyz}`;
    row.abcXyzMargin = `${row.abcMargin}${xyz}`;
    row.turnover = row.stockQty ? row.qty / row.stockQty : 0;
    row.gmroi = row.stockQty ? (row.margin / (row.stockQty * row.purchasePrice || 1)) : 0;
    row.asp = row.qty ? row.revenue / row.qty : 0;
  });

  return { rows, totals };
};

const computeMonthSummary = () => {
  const summary = [];
  Object.entries(state.salesMonths).forEach(([month, rows]) => {
    const agg = rows.reduce(
      (acc, row) => {
        acc.qty += row.qty;
        acc.revenue += row.revenue;
        const item = state.nomenclature.find((nom) => nom.sku === row.sku);
        const purchase = item ? item.purchasePrice : 0;
        acc.margin += row.revenue - purchase * row.qty;
        return acc;
      },
      { qty: 0, revenue: 0, margin: 0 }
    );
    summary.push({ month, ...agg });
  });
  return summary.sort((a, b) => a.month.localeCompare(b.month));
};

const renderTable = (table, headers, rows) => {
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  const headRow = document.createElement("tr");
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header.label;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    headers.forEach((header) => {
      const td = document.createElement("td");
      td.innerHTML = header.render(row);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
};

const renderNomenclature = () => {
  const metrics = computeNomenclatureMetrics();
  const query = elements.nomenclatureSearch.value.trim().toLowerCase();
  const filtered = metrics.filter((item) => {
    if (!query) return true;
    return (
      item.sku.toLowerCase().includes(query) ||
      item.name.toLowerCase().includes(query) ||
      item.cleanCategory.toLowerCase().includes(query)
    );
  });

  const headers = [
    { label: "Артикул", render: (row) => row.sku },
    { label: "Назва", render: (row) => row.name },
    { label: "Категорія", render: (row) => row.cleanCategory },
    { label: "ID категорії", render: (row) => row.idCategory },
    { label: "Ціна закупівлі", render: (row) => formatNumber(row.purchasePrice) },
    { label: "Ціна продажу", render: (row) => formatNumber(row.salePrice) },
    { label: "Кількість", render: (row) => formatNumber(row.stockQty, 0) },
    { label: "Перший продаж", render: (row) => (row.firstSale ? row.firstSale.toLocaleDateString("uk-UA") : "") },
    { label: "Дата полиці", render: (row) => (row.shelfDate ? row.shelfDate.toLocaleDateString("uk-UA") : "") },
    { label: "Вік, днів", render: (row) => (row.ageDays !== "" ? row.ageDays : "") },
    { label: "Днів до 1-ї продажі", render: (row) => (row.daysToFirstSale !== "" ? row.daysToFirstSale : "") },
    { label: "Aging bucket", render: (row) => row.agingBucket },
    { label: "СТ (шт/тиж)", render: (row) => formatNumber(row.weeklySales) },
    { label: "Запас, тижнів", render: (row) => (row.stockWeeks !== "" ? formatNumber(row.stockWeeks) : "") },
    { label: "Dead stock", render: (row) => (row.deadStock ? `<span class="badge warning">${row.deadStock}</span>` : "") },
    { label: "Активні тижні", render: (row) => (row.activeWeeks !== "" ? row.activeWeeks : "") },
  ];

  renderTable(elements.nomenclatureTable, headers, filtered);
  elements.nomenclatureCount.textContent = `Записів: ${filtered.length}`;
};

const renderAbcTable = () => {
  const { rows, totals } = computeAbcXyz();
  const query = elements.abcSearch.value.trim().toLowerCase();
  const filtered = rows.filter((row) => {
    if (!query) return true;
    return row.sku.toLowerCase().includes(query) || row.name.toLowerCase().includes(query);
  });

  const headers = [
    { label: "SKU", render: (row) => row.sku },
    { label: "Назва", render: (row) => row.name },
    { label: "Категорія", render: (row) => row.category },
    { label: "К-сть", render: (row) => formatNumber(row.qty, 0) },
    { label: "Виторг", render: (row) => formatNumber(row.revenue) },
    { label: "Маржа", render: (row) => formatNumber(row.margin) },
    { label: "ABC (в)", render: (row) => row.abcRevenue },
    { label: "ABC (м)", render: (row) => row.abcMargin },
    { label: "ABC (к)", render: (row) => row.abcQty },
    { label: "XYZ", render: (row) => row.xyz },
    { label: "ABC(в)-XYZ", render: (row) => row.abcXyzRevenue },
    { label: "ABC(м)-XYZ", render: (row) => row.abcXyzMargin },
    { label: "Оборотність", render: (row) => formatNumber(row.turnover) },
    { label: "GMROI", render: (row) => formatNumber(row.gmroi) },
    { label: "Маржинальність", render: (row) => formatPercent(row.marginPct) },
    { label: "Націнка", render: (row) => formatPercent(row.markup) },
    { label: "Сер. ціна", render: (row) => formatNumber(row.asp) },
  ];

  renderTable(elements.abcTable, headers, filtered);
  elements.abcCount.textContent = `SKU: ${filtered.length} | Виторг: ${formatNumber(totals.revenue)} | Маржа: ${formatNumber(totals.margin)}`;
};

const renderMonthSummary = () => {
  const summary = computeMonthSummary();
  const headers = [
    { label: "Місяць", render: (row) => row.month },
    { label: "Кількість", render: (row) => formatNumber(row.qty, 0) },
    { label: "Виторг", render: (row) => formatNumber(row.revenue) },
    { label: "Маржа", render: (row) => formatNumber(row.margin) },
  ];
  renderTable(elements.monthSummaryTable, headers, summary);
};

const refreshAll = () => {
  renderNomenclature();
  renderAbcTable();
  renderMonthSummary();
  storeState();
};

const handleNomenclatureUpload = async (file) => {
  const { data } = await parseFile(file);
  state.nomenclature = normalizeNomenclature(data);
  updateStatus(`Номенклатура імпортована: ${state.nomenclature.length} позицій.`, "ok");
  refreshAll();
};

const handleCategoryIdUpload = async (file) => {
  const { data } = await parseFile(file);
  state.categoryIds = new Map();
  data.forEach((row) => {
    const category = cleanCategory(row["Категорія"] || row["Категория"] || "");
    const id = row["ID"] || row["Id"] || row["ID Категорії"] || "";
    if (category && id) state.categoryIds.set(category, id);
  });
  updateStatus(`ID категорій імпортовано: ${state.categoryIds.size}.`, "ok");
  refreshAll();
};

const handleFirstSalesUpload = async (file) => {
  const { data } = await parseFile(file);
  state.firstSales = new Map();
  data.forEach((row) => {
    const sku = row["SKU"] || row["Артикул"] || "";
    const date = parseDate(row["Перша_продаж"] || row["Первая_продажа"] || row["Дата"] || "");
    if (sku && date) state.firstSales.set(String(sku).trim(), date);
  });
  updateStatus(`Перші продажі імпортовано: ${state.firstSales.size}.`, "ok");
  refreshAll();
};

const handleShelfDatesUpload = async (file) => {
  const { data } = await parseFile(file);
  state.shelfDates = new Map();
  data.forEach((row) => {
    const sku = row["Артикул"] || row["SKU"] || "";
    const date = parseDate(row["Дата_полки"] || row["Дата полки"] || row["Дата"] || "");
    if (sku && date) state.shelfDates.set(String(sku).trim(), date);
  });
  updateStatus(`ShelfDates імпортовано: ${state.shelfDates.size}.`, "ok");
  refreshAll();
};

const normalizeSales = (rows) =>
  rows
    .map((row) => {
      const sku = row["SKU"] || row["Артикул"] || row["Код"] || "";
      return {
        sku: String(sku).trim(),
        qty: Number(String(row["Кількість"] || row["Количество"] || row["Qty"] || 0).replace(",", ".")) || 0,
        revenue: Number(String(row["Виторг"] || row["Выручка"] || row["Revenue"] || 0).replace(",", ".")) || 0,
      };
    })
    .filter((row) => row.sku);

const handleSalesUpload = async () => {
  const file = elements.salesFile.files[0];
  const month = elements.salesMonth.value.trim();
  if (!file || !month) {
    updateStatus("Додайте файл продажів та назву місяця.", "warn");
    return;
  }
  const { data } = await parseFile(file);
  state.salesMonths[month] = normalizeSales(data);
  updateStatus(`Продажі за ${month} додано (${state.salesMonths[month].length} рядків).`, "ok");
  elements.salesFile.value = "";
  elements.salesMonth.value = "";
  refreshAll();
};

const handleSaveSettings = () => {
  state.settings.excludedCategories = splitLines(elements.excludedCategories.value);
  state.settings.excludedSkus = splitLines(elements.excludedSkus.value);
  updateStatus("Налаштування збережено.", "ok");
  state.nomenclature = filterNomenclature(state.nomenclature);
  refreshAll();
};

const handleReset = () => {
  localStorage.removeItem(DATA_KEY);
  window.location.reload();
};

const init = () => {
  restoreState();
  refreshAll();

  elements.nomenclatureFile.addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (file) handleNomenclatureUpload(file);
  });

  elements.categoryIdFile.addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (file) handleCategoryIdUpload(file);
  });

  elements.firstSalesFile.addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (file) handleFirstSalesUpload(file);
  });

  elements.shelfDatesFile.addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (file) handleShelfDatesUpload(file);
  });

  elements.addSales.addEventListener("click", handleSalesUpload);
  elements.saveSettings.addEventListener("click", handleSaveSettings);
  elements.nomenclatureSearch.addEventListener("input", renderNomenclature);
  elements.abcSearch.addEventListener("input", renderAbcTable);
  elements.resetData.addEventListener("click", handleReset);
};

init();
