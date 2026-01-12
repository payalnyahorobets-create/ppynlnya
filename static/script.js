const summaryUrl = "/api/summary";
const productsUrl = "/api/products";
const analysisUrl = "/api/analysis";
const monthlyUrl = "/api/monthly";

const productsTable = document.getElementById("products-table");
const analysisTable = document.getElementById("analysis-table");
const monthTable = document.getElementById("month-table");
const monthSelect = document.getElementById("month-select");

function renderTable(container, data) {
  if (!data || !data.columns || data.columns.length === 0) {
    container.innerHTML = '<div class="empty">Дані недоступні.</div>';
    return;
  }

  const headerCells = data.columns
    .map((col) => `<th>${col}</th>`)
    .join("");
  const rows = data.rows
    .map((row) => {
      const cells = data.columns
        .map((col) => `<td>${row[col] ?? ""}</td>`)
        .join("");
      return `<tr>${cells}</tr>`;
    })
    .join("");

  container.innerHTML = `
    <table class="table">
      <thead><tr>${headerCells}</tr></thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

async function loadSummary() {
  const response = await fetch(summaryUrl);
  const data = await response.json();
  document.getElementById("products-count").textContent = data.products_count ?? "—";
  document.getElementById("analysis-count").textContent = data.analysis_count ?? "—";
  document.getElementById("months-count").textContent = data.month_count ?? "—";

  monthSelect.innerHTML = "";
  (data.month_sheets || []).forEach((sheet) => {
    const option = document.createElement("option");
    option.value = sheet;
    option.textContent = sheet;
    monthSelect.appendChild(option);
  });
}

async function loadProducts() {
  const response = await fetch(`${productsUrl}?limit=200`);
  const data = await response.json();
  renderTable(productsTable, data);
}

async function loadAnalysis() {
  const response = await fetch(`${analysisUrl}?limit=200`);
  const data = await response.json();
  renderTable(analysisTable, data);
}

async function loadMonth() {
  const sheet = monthSelect.value;
  if (!sheet) {
    monthTable.innerHTML = '<div class="empty">Оберіть місяць.</div>';
    return;
  }
  const response = await fetch(`${monthlyUrl}?limit=200&sheet=${encodeURIComponent(sheet)}`);
  const data = await response.json();
  renderTable(monthTable, data);
}

async function initialize() {
  await loadSummary();
  productsTable.innerHTML = '<div class="empty">Натисніть «Завантажити», щоб переглянути дані.</div>';
  analysisTable.innerHTML = '<div class="empty">Натисніть «Завантажити», щоб переглянути дані.</div>';
  monthTable.innerHTML = '<div class="empty">Оберіть місяць та натисніть «Показати».</div>';
}

document.getElementById("refresh").addEventListener("click", initialize);
document.getElementById("load-products").addEventListener("click", loadProducts);
document.getElementById("load-analysis").addEventListener("click", loadAnalysis);
document.getElementById("load-month").addEventListener("click", loadMonth);

initialize();
