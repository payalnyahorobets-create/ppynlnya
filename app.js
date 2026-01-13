const state = {
  nomenclature: [],
  categoryIds: new Map(),
  firstSales: new Map(),
  shelfDates: new Map(),
  salesMonths: {},
  attributes: [],
  settings: {
    excludedCategories: [],
    excludedSkus: [],
  },
  meta: {
    lastUpdate: null,
    theme: "light",
  },
};

const branches = [
  { key: "shevchenko", label: "Шевченко" },
  { key: "nahorka", label: "Нагорка" },
  { key: "appollo", label: "Апполо" },
  { key: "horodok", label: "Городок" },
];

const elements = {
  themeToggle: document.getElementById("themeToggle"),
  resetData: document.getElementById("resetData"),
  tabs: document.querySelectorAll(".tab"),
  panels: document.querySelectorAll(".tab-panel"),
  openImport: document.getElementById("openImport"),
  importDialog: document.getElementById("importDialog"),
  importScope: document.getElementById("importScope"),
  importNomenclature: document.getElementById("importNomenclature"),
  nomenclatureFile: document.getElementById("nomenclatureFile"),
  categoryIdFile: document.getElementById("categoryIdFile"),
  firstSalesFile: document.getElementById("firstSalesFile"),
  shelfDatesFile: document.getElementById("shelfDatesFile"),
  excludedCategories: document.getElementById("excludedCategories"),
  excludedSkus: document.getElementById("excludedSkus"),
  saveSettings: document.getElementById("saveSettings"),
  importStatus: document.getElementById("importStatus"),
  lastUpdate: document.getElementById("lastUpdate"),
  nomenclatureTable: document.getElementById("nomenclatureTable"),
  nomenclatureSearch: document.getElementById("nomenclatureSearch"),
  nomenclatureCount: document.getElementById("nomenclatureCount"),
  categoryFilter: document.getElementById("categoryFilter"),
  clearFilters: document.getElementById("clearFilters"),
  attributesTable: document.getElementById("attributesTable"),
  attributeCategoryFilter: document.getElementById("attributeCategoryFilter"),
  attributeSearch: document.getElementById("attributeSearch"),
  addAttribute: document.getElementById("addAttribute"),
  saveAttributes: document.getElementById("saveAttributes"),
  purchaseTable: document.getElementById("purchaseTable"),
  monthSummaryTable: document.getElementById("monthSummaryTable"),
  salesDialog: document.getElementById("salesDialog"),
  openSalesImport: document.getElementById("openSalesImport"),
  salesFile: document.getElementById("salesFile"),
  salesMonth: document.getElementById("salesMonth"),
  addSales: document.getElementById("addSales"),
  salesCategoryFilter: document.getElementById("salesCategoryFilter"),
  salesSearch: document.getElementById("salesSearch"),
  abcTable: document.getElementById("abcTable"),
  abcSearch: document.getElementById("abcSearch"),
  abcCount: document.getElementById("abcCount"),
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

const storeState = () => {
  const payload = {
    nomenclature: state.nomenclature,
    categoryIds: Array.from(state.categoryIds.entries()),
    firstSales: Array.from(state.firstSales.entries()),
    shelfDates: Array.from(state.shelfDates.entries()),
    salesMonths: state.salesMonths,
    settings: state.settings,
    attributes: state.attributes,
    meta: state.meta,
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
  state.attributes = payload.attributes || [];
  state.meta = payload.meta || state.meta;
  elements.excludedCategories.value = state.settings.excludedCategories.join("\n");
  elements.excludedSkus.value = state.settings.excludedSkus.join("\n");
};

const updateStatus = (message, tone = "ok") => {
  elements.importStatus.textContent = message;
  elements.importStatus.className = `status ${tone}`;
};

const updateLastUpdate = () => {
  elements.lastUpdate.textContent = state.meta.lastUpdate
    ? new Date(state.meta.lastUpdate).toLocaleString("uk-UA")
    : "—";
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

const mergeNomenclature = (items, scope) => {
  const bySku = new Map(state.nomenclature.map((item) => [item.sku, item]));
  items.forEach((item) => {
    const existing = bySku.get(item.sku) || {
      sku: item.sku,
      code: item.code,
      name: item.name,
      category: item.category,
      purchasePrice: item.purchasePrice,
      salePrice: item.salePrice,
      barcode: item.barcode,
      stocks: {
        global: 0,
        shevchenko: 0,
        nahorka: 0,
        appollo: 0,
        horodok: 0,
      },
    };

    if (scope === "global") {
      existing.code = item.code;
      existing.name = item.name;
      existing.category = item.category;
      existing.purchasePrice = item.purchasePrice;
      existing.salePrice = item.salePrice;
      existing.barcode = item.barcode;
      existing.stocks.global = item.stockQty;
    } else {
      existing.stocks[scope] = item.stockQty;
    }

    bySku.set(item.sku, existing);
  });

  state.nomenclature = Array.from(bySku.values());
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
    const stockWeeks = weeklySales ? item.stocks.global / weeklySales : "";
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
      stockQty: item.stocks.global,
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
    row.gmroi = row.stockQty ? row.margin / (row.stockQty * row.purchasePrice || 1) : 0;
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

const buildCategoryOptions = (select, categories) => {
  select.innerHTML = "";
  categories.forEach((category) => {
    const option = document.createElement("option");
    option.value = category;
    option.textContent = category;
    select.appendChild(option);
  });
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
  const selectedCategories = Array.from(elements.categoryFilter.selectedOptions).map((opt) => opt.value);

  const filtered = metrics.filter((item) => {
    if (query) {
      const match =
        item.sku.toLowerCase().includes(query) ||
        item.name.toLowerCase().includes(query) ||
        item.cleanCategory.toLowerCase().includes(query);
      if (!match) return false;
    }
    if (selectedCategories.length && !selectedCategories.includes(item.cleanCategory)) return false;
    return true;
  });

  const headers = [
    { label: "Артикул", render: (row) => row.sku },
    { label: "Назва", render: (row) => row.name },
    { label: "Категорія", render: (row) => row.cleanCategory },
    { label: "ID категорії", render: (row) => row.idCategory },
    { label: "Ціна закупівлі", render: (row) => formatNumber(row.purchasePrice) },
    { label: "Ціна продажу", render: (row) => formatNumber(row.salePrice) },
    { label: "Загальний склад", render: (row) => formatNumber(row.stocks.global, 0) },
    ...branches.map((branch) => ({
      label: branch.label,
      render: (row) => formatNumber(row.stocks[branch.key], 0),
    })),
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

const renderAttributes = () => {
  const query = elements.attributeSearch.value.trim().toLowerCase();
  const selectedCategories = Array.from(elements.attributeCategoryFilter.selectedOptions).map(
    (opt) => opt.value
  );
  const filtered = state.attributes.filter((attr) => {
    if (query && !attr.name.toLowerCase().includes(query)) return false;
    if (selectedCategories.length && !selectedCategories.includes(attr.category)) return false;
    return true;
  });

  const headers = [
    { label: "Категорія", render: (row) => row.category },
    { label: "Атрибут", render: (row) => row.name },
    { label: "Значення", render: (row) => row.value },
    { label: "Примітка", render: (row) => row.note || "" },
  ];

  renderTable(elements.attributesTable, headers, filtered);
};

const renderPurchases = () => {
  const rows = state.nomenclature.map((item) => ({
    sku: item.sku,
    name: item.name,
    category: cleanCategory(item.category),
    global: item.stocks.global,
    shevchenko: item.stocks.shevchenko,
    nahorka: item.stocks.nahorka,
    appollo: item.stocks.appollo,
    horodok: item.stocks.horodok,
  }));

  const headers = [
    { label: "Артикул", render: (row) => row.sku },
    { label: "Назва", render: (row) => row.name },
    { label: "Категорія", render: (row) => row.category },
    { label: "Загальний склад", render: (row) => formatNumber(row.global, 0) },
    { label: "Шевченко", render: (row) => formatNumber(row.shevchenko, 0) },
    { label: "Нагорка", render: (row) => formatNumber(row.nahorka, 0) },
    { label: "Апполо", render: (row) => formatNumber(row.appollo, 0) },
    { label: "Городок", render: (row) => formatNumber(row.horodok, 0) },
  ];

  renderTable(elements.purchaseTable, headers, rows);
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
  elements.abcCount.textContent = `SKU: ${filtered.length} | Виторг: ${formatNumber(
    totals.revenue
  )} | Маржа: ${formatNumber(totals.margin)}`;
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
  const categories = Array.from(
    new Set(state.nomenclature.map((item) => cleanCategory(item.category)).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));
  buildCategoryOptions(elements.categoryFilter, categories);
  buildCategoryOptions(elements.attributeCategoryFilter, categories);
  buildCategoryOptions(elements.salesCategoryFilter, categories);
  renderNomenclature();
  renderAttributes();
  renderPurchases();
  renderAbcTable();
  renderMonthSummary();
  updateLastUpdate();
  storeState();
};

const handleNomenclatureImport = async () => {
  const file = elements.nomenclatureFile.files[0];
  if (!file) {
    updateStatus("Додайте файл номенклатури.", "warn");
    return;
  }
  const scope = elements.importScope.value;
  const { data } = await parseFile(file);
  const normalized = normalizeNomenclature(data);
  mergeNomenclature(normalized, scope);
  state.meta.lastUpdate = new Date().toISOString();
  updateStatus(`Номенклатура імпортована: ${normalized.length} позицій.`, "ok");
  refreshAll();
};

const handleCategoryIdUpload = async (file) => {
  if (!file) return;
  const { data } = await parseFile(file);
  state.categoryIds = new Map();
  data.forEach((row) => {
    const category = cleanCategory(row["Категорія"] || row["Категория"] || "");
    const id = row["ID"] || row["Id"] || row["ID Категорії"] || "";
    if (category && id) state.categoryIds.set(category, id);
  });
};

const handleFirstSalesUpload = async (file) => {
  if (!file) return;
  const { data } = await parseFile(file);
  state.firstSales = new Map();
  data.forEach((row) => {
    const sku = row["SKU"] || row["Артикул"] || "";
    const date = parseDate(row["Перша_продаж"] || row["Первая_продажа"] || row["Дата"] || "");
    if (sku && date) state.firstSales.set(String(sku).trim(), date);
  });
};

const handleShelfDatesUpload = async (file) => {
  if (!file) return;
  const { data } = await parseFile(file);
  state.shelfDates = new Map();
  data.forEach((row) => {
    const sku = row["Артикул"] || row["SKU"] || "";
    const date = parseDate(row["Дата_полки"] || row["Дата полки"] || row["Дата"] || "");
    if (sku && date) state.shelfDates.set(String(sku).trim(), date);
  });
};

const normalizeSales = (rows) =>
  rows
    .map((row) => {
      const sku = row["SKU"] || row["Артикул"] || row["Код"] || "";
      return {
        sku: String(sku).trim(),
        qty: Number(String(row["Кількість"] || row["Количество"] || row["Qty"] || 0).replace(",", ".")) ||
          0,
        revenue: Number(String(row["Виторг"] || row["Выручка"] || row["Revenue"] || 0).replace(",", ".")) ||
          0,
      };
    })
    .filter((row) => row.sku);

const handleSalesUpload = async () => {
  const file = elements.salesFile.files[0];
  const month = elements.salesMonth.value.trim();
  if (!file || !month) {
    return;
  }
  const { data } = await parseFile(file);
  state.salesMonths[month] = normalizeSales(data);
  elements.salesFile.value = "";
  elements.salesMonth.value = "";
  refreshAll();
};

const handleSaveSettings = () => {
  state.settings.excludedCategories = splitLines(elements.excludedCategories.value);
  state.settings.excludedSkus = splitLines(elements.excludedSkus.value);
  refreshAll();
};

const handleAddAttribute = () => {
  const category = prompt("Категорія:");
  if (!category) return;
  const name = prompt("Назва атрибуту:");
  if (!name) return;
  const value = prompt("Значення:") || "";
  state.attributes.push({ category, name, value, note: "" });
  refreshAll();
};

const switchTab = (tabId) => {
  elements.tabs.forEach((tab) => tab.classList.toggle("active", tab.dataset.tab === tabId));
  elements.panels.forEach((panel) =>
    panel.classList.toggle("active", panel.dataset.panel === tabId)
  );
};

const toggleTheme = () => {
  state.meta.theme = state.meta.theme === "dark" ? "light" : "dark";
  document.body.setAttribute("data-theme", state.meta.theme);
  storeState();
};

const init = () => {
  restoreState();
  document.body.setAttribute("data-theme", state.meta.theme);
  updateLastUpdate();
  refreshAll();

  elements.tabs.forEach((tab) => {
    tab.addEventListener("click", () => switchTab(tab.dataset.tab));
  });

  elements.themeToggle.addEventListener("click", toggleTheme);
  elements.resetData.addEventListener("click", () => {
    localStorage.removeItem(DATA_KEY);
    window.location.reload();
  });

  elements.openImport.addEventListener("click", () => elements.importDialog.showModal());
  elements.openSalesImport.addEventListener("click", () => elements.salesDialog.showModal());

  elements.importNomenclature.addEventListener("click", async (event) => {
    event.preventDefault();
    await handleCategoryIdUpload(elements.categoryIdFile.files[0]);
    await handleFirstSalesUpload(elements.firstSalesFile.files[0]);
    await handleShelfDatesUpload(elements.shelfDatesFile.files[0]);
    handleSaveSettings();
    await handleNomenclatureImport();
    elements.importDialog.close();
  });

  elements.addSales.addEventListener("click", async (event) => {
    event.preventDefault();
    await handleSalesUpload();
    elements.salesDialog.close();
  });

  elements.saveSettings.addEventListener("click", (event) => {
    event.preventDefault();
    handleSaveSettings();
    elements.importDialog.close();
  });

  elements.nomenclatureSearch.addEventListener("input", renderNomenclature);
  elements.categoryFilter.addEventListener("change", renderNomenclature);
  elements.clearFilters.addEventListener("click", () => {
    elements.categoryFilter.selectedIndex = -1;
    elements.nomenclatureSearch.value = "";
    renderNomenclature();
  });

  elements.attributeSearch.addEventListener("input", renderAttributes);
  elements.attributeCategoryFilter.addEventListener("change", renderAttributes);
  elements.addAttribute.addEventListener("click", handleAddAttribute);

  elements.salesSearch.addEventListener("input", renderMonthSummary);
  elements.salesCategoryFilter.addEventListener("change", renderMonthSummary);

  elements.abcSearch.addEventListener("input", renderAbcTable);
};

init();
