const statusMessage = document.querySelector("#status");
const reportInput = document.querySelector("#report");
const uploadForm = document.querySelector("#upload-form");
const catalogInput = document.querySelector("#catalog");
const catalogForm = document.querySelector("#catalog-form");
const trainButton = document.querySelector("#train-button");
const forecastButton = document.querySelector("#forecast-button");
const forecastDays = document.querySelector("#days");
const forecastCutoff = document.querySelector("#forecast-cutoff");
const clearForecastCutoff = document.querySelector("#clear-forecast-cutoff");
const MIN_HISTORICAL_TRAINING_DAYS = 48;
const forecastDate = document.querySelector("#forecast-date");
const forecastDateControl = document.querySelector("#forecast-date-control");
const downloadForecastButton = document.querySelector("#download-forecast-button");
const inventoryInput = document.querySelector("#inventory");
const inventoryForm = document.querySelector("#inventory-form");
const runPosttestButton = document.querySelector("#run-posttest-button");
const downloadPosttestButton = document.querySelector("#download-posttest-button");
const posttestResults = document.querySelector("#posttest-results");

let currentForecast = null;
let currentPosttest = null;
let currentDashboardSummary = null;

function selectedForecastCutoff() {
  return forecastCutoff.value || null;
}

function shiftUtcDate(value, days) {
  if (!value) return "";
  const date = new Date(`${value}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) return "";
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

function earliestHistoricalCutoff(firstSaleDate) {
  return shiftUtcDate(firstSaleDate, MIN_HISTORICAL_TRAINING_DAYS - 1);
}

function posttestCutoffBounds(summary = currentDashboardSummary) {
  const days = Number(forecastDays.value);
  const min = earliestHistoricalCutoff(summary?.first_sale_date);
  const max = Number.isFinite(days) && days > 0
    ? shiftUtcDate(summary?.last_sale_date, -days)
    : "";
  return {
    days,
    min,
    max,
    available: Boolean(min && max && min <= max)
  };
}

function updateHistoricalCutoffControl(summary = currentDashboardSummary) {
  const bounds = posttestCutoffBounds(summary);
  const note = document.querySelector("#forecast-cutoff-note");
  const selected = selectedForecastCutoff();
  const selectionIsOutsideRange = Boolean(selected && (!bounds.available || selected < bounds.min || selected > bounds.max));

  forecastCutoff.min = bounds.min;
  forecastCutoff.max = bounds.max;
  forecastCutoff.disabled = !bounds.available;
  if (selectionIsOutsideRange) forecastCutoff.value = "";

  if (!summary?.first_sale_date || !summary?.last_sale_date) {
    note.textContent = "Carga ventas históricas para habilitar una prueba histórica.";
  } else if (!bounds.available) {
    note.textContent = `La prueba histórica de ${numberFormat.format(bounds.days)} días estará disponible cuando haya al menos ${MIN_HISTORICAL_TRAINING_DAYS} días para entrenar y ${numberFormat.format(bounds.days)} días reales posteriores. Para una compra futura, deja el corte vacío: se usará todo el historial.`;
  } else {
    note.textContent = `Para un Postest de ${numberFormat.format(bounds.days)} días, elige un corte entre ${formatDate(bounds.min)} y ${formatDate(bounds.max)}. El modelo se entrena con al menos ${MIN_HISTORICAL_TRAINING_DAYS} días hasta el corte y se compara con las ventas reales de los ${numberFormat.format(bounds.days)} días siguientes. Para una compra futura, déjala vacía: se usará todo el historial.`;
  }
  return { ...bounds, selectionIsOutsideRange };
}

function hasNumericValue(value) {
  return value !== null
    && value !== undefined
    && !(typeof value === "string" && value.trim() === "")
    && Number.isFinite(Number(value));
}

function formatMetricNumber(value) {
  return hasNumericValue(value) ? numberFormat.format(Number(value)) : "—";
}

const numberFormat = new Intl.NumberFormat("es-PE", { maximumFractionDigits: 2 });
const dateFormat = new Intl.DateTimeFormat("es-PE", { day: "2-digit", month: "short", year: "numeric", timeZone: "UTC" });
const dateTimeFormat = new Intl.DateTimeFormat("es-PE", { dateStyle: "medium", timeStyle: "short" });

function formatDate(value) {
  return value ? dateFormat.format(new Date(`${value}T00:00:00Z`)) : "Sin historial";
}

function formatDateTime(value) {
  return value ? `Historial actualizado: ${dateTimeFormat.format(new Date(value))}` : "Historial aún no actualizado";
}

function formatPercent(value) {
  return hasNumericValue(value) ? `${numberFormat.format(Number(value))}%` : "—";
}

function escapeHtml(value) {
  return String(value).replace(/[&<>'"]/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", "\"": "&quot;"
  }[character]));
}

function setStatus(message = "", kind = "") {
  statusMessage.textContent = message;
  statusMessage.dataset.kind = kind;
}

function setButtonLoading(button, active, label) {
  button.disabled = active;
  button.dataset.originalLabel ||= button.innerHTML;
  button.innerHTML = active ? `${label} <span aria-hidden="true">…</span>` : button.dataset.originalLabel;
}

function hideForecastResults() {
  const results = document.querySelector("#forecast-results");
  results.hidden = true;
  document.querySelector("#forecast-period").textContent = "";
  document.querySelector("#prediction-list").innerHTML = "";
  forecastDateControl.hidden = true;
  forecastDate.innerHTML = "";
  downloadForecastButton.hidden = true;
  downloadForecastButton.onclick = null;
  currentForecast = null;
}

function hidePosttestResults() {
  posttestResults.hidden = true;
  document.querySelector("#posttest-period").textContent = "—";
  document.querySelector("#posttest-summary").textContent = "—";
  document.querySelector("#posttest-metrics").innerHTML = "";
  document.querySelector("#posttest-comparison").textContent = "";
  document.querySelector("#posttest-category-rows").innerHTML = "";
  document.querySelector("#inventory-result-note").textContent = "Carga un inventario para calcular ISI y TQS en este período.";
  document.querySelector("#inventory-metric-rows").innerHTML = "";
  downloadPosttestButton.onclick = null;
  currentPosttest = null;
}

async function request(url, options = {}) {
  const response = await fetch(url, options);
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || "No se pudo completar la operación.");
  return data;
}

function renderDashboard(data) {
  const { summary, categories, recent, catalog } = data;
  currentDashboardSummary = summary;
  document.querySelector("#records").textContent = numberFormat.format(summary.records);
  document.querySelector("#categories-count").textContent = numberFormat.format(summary.categories);
  document.querySelector("#total-quantity").textContent = numberFormat.format(summary.total_quantity);
  document.querySelector("#last-sale-date").textContent = formatDate(summary.last_sale_date);
  document.querySelector("#history-updated-at").textContent = formatDateTime(summary.history_updated_at);
  updateHistoricalCutoffControl(summary);
  document.querySelector("#category-history-explanation").textContent = summary.records
    ? `Acumulado histórico: unidades y porcentaje de participación de cada categoría en todos los reportes, desde ${formatDate(summary.first_sale_date)} hasta ${formatDate(summary.last_sale_date)}. Se actualiza al cargar un reporte. No es un pronóstico.`
    : "Aún no hay reportes cargados para calcular el acumulado histórico.";
  document.querySelector("#catalog-products").textContent = numberFormat.format(catalog.products);
  document.querySelector("#catalog-families").textContent = numberFormat.format(catalog.families);
  document.querySelector("#catalog-categories").textContent = numberFormat.format(catalog.categories);
  const bars = document.querySelector("#category-bars");
  if (!categories.length) {
    bars.innerHTML = '<p class="empty-state">Aún no hay demanda cargada. Empieza procesando un reporte.</p>';
  } else {
    bars.innerHTML = categories.map((item, index) => `
      <div class="bar-row bar-row-tone-${index % 5}">
        <span class="bar-rank" aria-hidden="true">${String(index + 1).padStart(2, "0")}</span>
        <span class="bar-label" title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</span>
        <div class="bar-track" aria-label="${escapeHtml(item.name)}: ${numberFormat.format(item.percentage)}% del total histórico"><div class="bar-value" style="width: ${Math.max(0, Math.min(100, Number(item.percentage) || 0))}%"></div></div>
        <span class="bar-metrics"><span class="bar-number">${numberFormat.format(item.quantity)} unid.</span><span class="bar-percentage">${numberFormat.format(item.percentage)}%</span></span>
      </div>`).join("");
  }

  const table = document.querySelector("#recent-table");
  table.innerHTML = recent.length
    ? recent.map((item) => `<tr><td><time datetime="${escapeHtml(item.date)}">${formatDate(item.date)}</time></td><td title="${escapeHtml(item.category)}"><span class="recent-category">${escapeHtml(item.category)}</span></td><td><strong class="recent-units">${numberFormat.format(item.quantity)}</strong></td></tr>`).join("")
    : '<tr><td colspan="3" class="empty-state">No hay movimientos para mostrar.</td></tr>';
}

async function loadDashboard() {
  try {
    renderDashboard(await request("/api/dashboard"));
  } catch (error) {
    setStatus(error.message, "error");
  }
}

function renderInventoryStatus(data) {
  const note = document.querySelector("#inventory-file-note");
  if (!data?.available) {
    note.textContent = data?.message || "Aún no hay un inventario cargado.";
    document.querySelector("#inventory-preview").hidden = true;
    return;
  }
  const inventory = data.inventory || {};
  const isProjectTestFile = inventory.is_demonstration === true;
  const sourceMessage = isProjectTestFile
    ? "Es un archivo de demostración del proyecto; carga el kardex real para resultados oficiales."
    : inventory.source_kind === "archivo cargado por el usuario"
      ? "Archivo cargado por el usuario; confirma que sea el kardex oficial de la tienda."
      : "Archivo de inventario activo.";
  const reconciliation = inventory.reconciliation || data.reconciliation || {};
  const hasCoverageWarning = Number(reconciliation.rows_with_difference) > 0
    || (hasNumericValue(reconciliation.difference) && Math.abs(Number(reconciliation.difference)) > 0.000001);
  const coverageDetail = hasNumericValue(reconciliation.rows_with_difference)
    ? `${formatMetricNumber(reconciliation.rows_with_difference)} registros difieren de las ventas consolidadas`
    : "la demanda del inventario difiere de las ventas consolidadas";
  const warnings = Array.isArray(data.warnings) ? data.warnings.filter(Boolean) : [];
  const reconciliationMessage = warnings.length
    ? ` Advertencia: ${warnings.join(" ")}`
    : hasCoverageWarning
      ? ` Advertencia de conciliación: ${coverageDetail}; revisa el kardex antes de usar los indicadores como evidencia oficial.`
      : "";
  note.textContent = `${inventory.name || "Inventario activo"}: ${formatMetricNumber(inventory.records)} registros, ${formatMetricNumber(inventory.categories)} categorías, del ${formatDate(inventory.start_date)} al ${formatDate(inventory.end_date)}. ${sourceMessage}${reconciliationMessage}`;
  renderInventoryPreview(data);
}

function renderInventoryPreview(data) {
  const preview = document.querySelector("#inventory-preview");
  const inventory = data.inventory || {};
  const metrics = Array.isArray(data.metrics) ? data.metrics : [];
  if (!data?.available || !metrics.length) {
    preview.hidden = true;
    return;
  }
  document.querySelector("#inventory-preview-period").textContent = `${formatDate(inventory.start_date)} al ${formatDate(inventory.end_date)}`;
  document.querySelector("#inventory-preview-rows").innerHTML = metrics.map((metric) => `<tr>
    <td title="${escapeHtml(metric.category)}">${escapeHtml(metric.category)}</td>
    <td>${formatMetricNumber(metric.CD)}</td>
    <td>${formatPercent(metric.ISI)}</td>
    <td>${formatMetricNumber(metric.DQS)}</td>
    <td>${formatPercent(metric.TQS)}</td>
  </tr>`).join("");
  preview.hidden = false;
}

async function loadInventorySummary() {
  try {
    renderInventoryStatus(await request("/api/inventory"));
  } catch (error) {
    document.querySelector("#inventory-file-note").textContent = error.message;
  }
}

reportInput.addEventListener("change", () => {
  document.querySelector("#file-name").textContent = reportInput.files[0]?.name || "Seleccionar archivo";
});

catalogInput.addEventListener("change", () => {
  document.querySelector("#catalog-file-name").textContent = catalogInput.files[0]?.name || "Nuevo catálogo .xlsx";
});

inventoryInput.addEventListener("change", () => {
  document.querySelector("#inventory-file-name").textContent = inventoryInput.files[0]?.name || "Seleccionar inventario .xlsx";
});

forecastDays.addEventListener("change", () => {
  const bounds = updateHistoricalCutoffControl();
  hideForecastResults();
  hidePosttestResults();
  setStatus(
    bounds.selectionIsOutsideRange
      ? `El corte anterior no tiene ${numberFormat.format(bounds.days)} días reales posteriores. Se quitó para evitar una prueba incompleta. Para una compra futura, déjalo vacío.`
      : "Selecciona «Generar pronóstico» para ver la estimación del nuevo periodo.",
    ""
  );
});

forecastCutoff.addEventListener("change", () => {
  document.querySelector("#metrics").hidden = true;
  hideForecastResults();
  hidePosttestResults();
  setStatus(
    selectedForecastCutoff()
      ? `Prueba histórica configurada hasta el ${formatDate(selectedForecastCutoff())}. Entrena el modelo antes de pronosticar.`
      : "Se usará todo el historial disponible. Entrena el modelo antes de pronosticar.",
    ""
  );
});

clearForecastCutoff.addEventListener("click", () => {
  forecastCutoff.value = "";
  forecastCutoff.dispatchEvent(new Event("change"));
});

forecastDate.addEventListener("change", () => {
  if (currentForecast) renderForecast(currentForecast, forecastDate.value, false);
});

uploadForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!reportInput.files[0]) return;
  const button = uploadForm.querySelector("button");
  setButtonLoading(button, true, "Procesando");
  setStatus("Procesando el reporte y consolidando la demanda…", "working");
  try {
    const formData = new FormData(uploadForm);
    const data = await request("/api/reports", { method: "POST", body: formData });
    renderDashboard(data.dashboard);
    const summary = data.save_summary;
    const catalogSummary = data.catalog_summary;
    const catalogMessage = catalogSummary.unmatched_products.length
      ? ` Productos sin catálogo: ${catalogSummary.unmatched_products.join(", ")}.`
      : " Todos los productos del reporte se asociaron al catálogo.";
    // La prueba histórica es opcional. Un reporte reciente no debe convertirse
    // automáticamente en corte, porque podría dejar muy poco historial para
    // entrenar el modelo.
    forecastCutoff.value = "";
    setStatus(
      `${data.message} Nuevos: ${numberFormat.format(summary.added)} · actualizados: ${numberFormat.format(summary.updated)} · sin cambios: ${numberFormat.format(summary.unchanged)}.${catalogMessage} Se usará todo el historial disponible. Entrena nuevamente antes de pronosticar.`,
      "success"
    );
    document.querySelector("#metrics").hidden = true;
    hideForecastResults();
    hidePosttestResults();
    uploadForm.reset();
    document.querySelector("#file-name").textContent = "Seleccionar archivo";
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setButtonLoading(button, false);
  }
});

catalogForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!catalogInput.files[0]) {
    setStatus("Selecciona el nuevo catálogo .xlsx antes de actualizarlo.", "error");
    return;
  }
  const button = catalogForm.querySelector("button");
  setButtonLoading(button, true, "Validando");
  setStatus("Validando y actualizando el catálogo maestro…", "working");
  try {
    const data = await request("/api/catalog", { method: "POST", body: new FormData(catalogForm) });
    renderDashboard(data.dashboard);
    hideForecastResults();
    hidePosttestResults();
    const unresolved = data.dashboard.catalog.unmapped_historical_categories;
    const unresolvedMessage = unresolved.length
      ? ` Quedan ${numberFormat.format(unresolved.length)} productos históricos sin coincidencia.`
      : " Todos los productos históricos tienen clasificación.";
    setStatus(`${data.message}${unresolvedMessage}`, "success");
    catalogForm.reset();
    document.querySelector("#catalog-file-name").textContent = "Nuevo catálogo .xlsx";
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setButtonLoading(button, false);
  }
});

inventoryForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!inventoryInput.files[0]) {
    setStatus("Selecciona el archivo de inventario .xlsx antes de cargarlo.", "error");
    return;
  }
  const button = document.querySelector("#inventory-upload-button");
  setButtonLoading(button, true, "Validando");
  setStatus("Validando el inventario y calculando sus indicadores…", "working");
  try {
    const data = await request("/api/inventory", { method: "POST", body: new FormData(inventoryForm) });
    renderInventoryStatus(data);
    hidePosttestResults();
    inventoryForm.reset();
    document.querySelector("#inventory-file-name").textContent = "Seleccionar inventario .xlsx";
    setStatus(data.message, "success");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setButtonLoading(button, false);
  }
});

runPosttestButton.addEventListener("click", async () => {
  const cutoffDate = selectedForecastCutoff();
  const days = Number(forecastDays.value);
  const bounds = posttestCutoffBounds();
  if (!cutoffDate) {
    setStatus("Para ejecutar el Postest, selecciona arriba el corte de prueba: el día anterior al período que deseas evaluar.", "error");
    forecastCutoff.focus();
    return;
  }
  if (!bounds.available || cutoffDate < bounds.min || cutoffDate > bounds.max) {
    setStatus(
      bounds.available
        ? `El corte debe estar entre ${formatDate(bounds.min)} y ${formatDate(bounds.max)} para evaluar ${numberFormat.format(bounds.days)} días con ventas reales.`
        : `Aún no hay datos suficientes para un Postest de ${numberFormat.format(bounds.days)} días.`,
      "error"
    );
    forecastCutoff.focus();
    return;
  }
  setButtonLoading(runPosttestButton, true, "Evaluando");
  setStatus(`Ejecutando el Postest desde el día posterior al ${formatDate(cutoffDate)}…`, "working");
  try {
    const data = await request("/api/posttest", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ cutoff_date: cutoffDate, days })
    });
    renderPosttest(data);
    setStatus(data.message, "success");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setButtonLoading(runPosttestButton, false);
  }
});

trainButton.addEventListener("click", async () => {
  const cutoffDate = selectedForecastCutoff();
  setButtonLoading(trainButton, true, "Entrenando");
  setStatus(
    cutoffDate
      ? `Entrenando con ventas hasta el ${formatDate(cutoffDate)}…`
      : "Entrenando el modelo con todo el historial disponible…",
    "working"
  );
  try {
    const data = await request("/api/train", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ cutoff_date: cutoffDate })
    });
    document.querySelector("#metrics").hidden = false;
    document.querySelector("#metric-wape").textContent = `${numberFormat.format(data.metrics.wape)}%`;
    document.querySelector("#metric-rmse").textContent = numberFormat.format(data.metrics.rmse);
    document.querySelector("#metric-mae").textContent = numberFormat.format(data.metrics.mae);
    hideForecastResults();
    setStatus(data.message, "success");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setButtonLoading(trainButton, false);
  }
});

forecastButton.addEventListener("click", async () => {
  const days = Number(forecastDays.value);
  const cutoffDate = selectedForecastCutoff();
  setButtonLoading(forecastButton, true, "Calculando");
  setStatus(
    cutoffDate
      ? `Calculando el pronóstico desde el día posterior al ${formatDate(cutoffDate)}…`
      : "Calculando el pronóstico por categoría…",
    "working"
  );
  try {
    const data = await request("/api/forecast", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ days, cutoff_date: cutoffDate })
    });
    renderForecast(data);
    setStatus(data.message, "success");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setButtonLoading(forecastButton, false);
  }
});

function consolidateProductForecasts(dailyForecasts) {
  const products = new Map();
  for (const forecast of dailyForecasts) {
    for (const product of forecast.products || []) {
      const current = products.get(product.name) || { ...product, quantity: 0 };
      current.quantity += Number(product.quantity) || 0;
      products.set(product.name, current);
    }
  }
  const consolidated = Array.from(products.values()).map((product) => ({
    ...product,
    quantity: Math.round(product.quantity * 100) / 100
  }));
  const totalQuantity = consolidated.reduce((total, product) => total + product.quantity, 0);
  return consolidated.map((product) => ({
    ...product,
    allocation_share: totalQuantity ? Math.round((product.quantity / totalQuantity) * 10000) / 100 : 0
  }));
}

function allocateWholePackages(products) {
  const plan = products
    .map((product) => {
      const quantity = Math.max(0, Number(product.quantity) || 0);
      return {
        ...product,
        quantity,
        suggested_packages: Math.floor(quantity),
        remainder: quantity % 1
      };
    })
    .filter((product) => product.quantity > 0);
  const targetPackages = Math.ceil(plan.reduce((total, product) => total + product.quantity, 0));
  let packagesPending = targetPackages - plan.reduce((total, product) => total + product.suggested_packages, 0);
  const priority = [...plan].sort((first, second) => (
    second.remainder - first.remainder
    || second.quantity - first.quantity
    || first.name.localeCompare(second.name, "es")
  ));

  for (let index = 0; packagesPending > 0 && priority.length; index += 1, packagesPending -= 1) {
    priority[index % priority.length].suggested_packages += 1;
  }

  return plan
    .filter((product) => product.suggested_packages > 0)
    .sort((first, second) => second.suggested_packages - first.suggested_packages || first.name.localeCompare(second.name, "es"));
}

function buildForecastExportPayload(data, viewToShow, categoryScope = null) {
  const entries = Object.entries(data.predictions || {}).filter(([category]) => (
    !categoryScope || category === categoryScope
  ));
  const availableDates = entries[0]?.[1].map((demand) => demand.date) || [];
  const isPeriodView = viewToShow === "period";
  const dateToShow = isPeriodView ? null : viewToShow;
  const rows = [];

  for (const [category, demands] of entries) {
    const dailyForecasts = (data.product_forecasts_by_category || {})[category] || demands.map((demand) => ({
      date: demand.date,
      category_quantity: demand.quantity,
      products: []
    }));
    const selectedForecast = dailyForecasts.find((forecast) => forecast.date === dateToShow)
      || dailyForecasts[0];
    const products = isPeriodView
      ? consolidateProductForecasts(dailyForecasts)
      : (selectedForecast?.products || []);
    const purchasePlan = allocateWholePackages(products);

    for (const product of purchasePlan) {
      rows.push({
        category,
        product: product.name,
        quantity: Math.round((Number(product.quantity) || 0) * 100) / 100,
        suggested_packages: product.suggested_packages,
        allocation_share: Number(product.allocation_share ?? product.historical_share) || 0
      });
    }
  }

  const selectedDate = isPeriodView ? null : dateToShow;
  return {
    scope: {
      view: isPeriodView ? "period" : "date",
      days: data.days,
      start_date: selectedDate || availableDates[0],
      end_date: selectedDate || availableDates.at(-1),
      cutoff_date: data.cutoff_date || null,
      category: categoryScope || null
    },
    rows
  };
}

function categoryExportFilename(categoryScope, payload) {
  if (!categoryScope) return null;
  const category = categoryScope
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
  const { start_date: startDate, end_date: endDate } = payload.scope;
  const period = startDate === endDate ? startDate : `${startDate}_a_${endDate}`;
  return `pronostico_compra_${category}_${period}.xlsx`;
}

async function downloadForecastExcel(data, viewToShow, categoryScope = null, button = downloadForecastButton) {
  const payload = buildForecastExportPayload(data, viewToShow, categoryScope);
  if (!payload.rows.length) {
    setStatus(categoryScope
      ? `No hay productos pronosticados para descargar en ${categoryScope}.`
      : "No hay productos pronosticados para descargar.", "error");
    return;
  }

  setButtonLoading(button, true, "Preparando");
  setStatus(categoryScope
    ? `Preparando el Excel de ${categoryScope}…`
    : "Preparando el Excel del pronóstico…", "working");
  try {
    const response = await fetch("/api/forecast/export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    if (!response.ok) {
      const error = await response.json().catch(() => ({}));
      throw new Error(error.error || "No se pudo preparar el archivo Excel.");
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const filename = response.headers.get("Content-Disposition")?.match(/filename=([^;]+)/i)?.[1];
    link.href = url;
    link.download = categoryExportFilename(categoryScope, payload)
      || filename?.replaceAll('"', "")
      || "pronostico_compra.xlsx";
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setStatus(categoryScope
      ? `El Excel de ${categoryScope} se descargó correctamente.`
      : "El archivo Excel del pronóstico se descargó correctamente.", "success");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setButtonLoading(button, false);
  }
}

function posttestMetricCard(title, metrics, modifier = "") {
  const entries = [
    ["WAPE", formatPercent(metrics?.wape)],
    ["MAPE", formatPercent(metrics?.mape)],
    ["MAE", formatMetricNumber(metrics?.mae)],
    ["RMSE", formatMetricNumber(metrics?.rmse)]
  ];
  return `
    <article class="posttest-metric-card ${modifier}">
      <h4>${escapeHtml(title)}</h4>
      <dl>${entries.map(([label, value]) => `<div><dt>${label}</dt><dd>${escapeHtml(value)}</dd></div>`).join("")}</dl>
    </article>`;
}

function posttestWinnerLabel(modelWape, baselineWape) {
  if (!hasNumericValue(modelWape) || !hasNumericValue(baselineWape)) {
    return { label: "No disponible", className: "" };
  }
  const model = Number(modelWape);
  const baseline = Number(baselineWape);
  if (Math.abs(model - baseline) < 0.000001) return { label: "Empate", className: "" };
  return model < baseline
    ? { label: "Modelo", className: "is-model" }
    : { label: "PMS-7", className: "is-baseline" };
}

function renderPosttest(data) {
  currentPosttest = data;
  const model = data.model || {};
  const baseline = data.baseline || {};
  const period = data.period || {};
  const modelMetrics = model.metrics || {};
  const baselineMetrics = baseline.metrics || {};
  const categories = Array.from(new Set([
    ...Object.keys(model.by_category || {}),
    ...Object.keys(baseline.by_category || {})
  ])).sort((first, second) => first.localeCompare(second, "es"));
  const observationCount = formatMetricNumber(data.observation_count);
  const excludedCategorySource = data.training?.excluded_categories;
  const excludedCategories = Array.isArray(excludedCategorySource)
    ? excludedCategorySource.filter((category) => typeof category === "string" && category.trim())
    : excludedCategorySource && typeof excludedCategorySource === "object"
      ? Object.keys(excludedCategorySource).filter((category) => category.trim())
      : [];
  const exclusionNote = excludedCategories.length
    ? ` Las métricas globales cubren solo las categorías evaluadas; se excluyeron: ${excludedCategories.join(", ")}.`
    : "";

  document.querySelector("#posttest-period").textContent = `${formatDate(period.start_date)} al ${formatDate(period.end_date)} · ${numberFormat.format(data.horizon_days)} días`;
  document.querySelector("#posttest-summary").textContent = `Corte: ${formatDate(data.cutoff_date)}. Se contrastaron ${observationCount} observaciones de ${numberFormat.format(categories.length)} categorías con ventas reales posteriores.${exclusionNote}`;
  document.querySelector("#posttest-metrics").innerHTML = [
    posttestMetricCard(model.name || "Modelo predictivo", modelMetrics),
    posttestMetricCard(baseline.name || "PMS-7", baselineMetrics, "is-baseline")
  ].join("");

  const winner = hasNumericValue(modelMetrics.wape) && hasNumericValue(baselineMetrics.wape)
    ? data.comparison?.winner_by_metric?.wape
    : null;
  const wapeImprovement = data.comparison?.improvement_vs_baseline?.wape?.percent;
  const comparison = document.querySelector("#posttest-comparison");
  if (winner === "model") {
    const reduction = hasNumericValue(wapeImprovement)
      ? ` La reducción frente a PMS-7 fue de ${escapeHtml(formatPercent(wapeImprovement))}.`
      : " No se reportó el porcentaje de reducción.";
    comparison.innerHTML = `<strong>Resultado WAPE:</strong> el modelo predictivo obtuvo menor error que PMS-7.${reduction}`;
  } else if (winner === "pms_7") {
    comparison.innerHTML = `<strong>Resultado WAPE:</strong> PMS-7 obtuvo menor error en este período. El Postest queda guardado para revisar y mejorar el modelo con más datos.`;
  } else if (winner === "tie") {
    comparison.innerHTML = `<strong>Resultado WAPE:</strong> el modelo predictivo y PMS-7 tuvieron el mismo error en este período.`;
  } else {
    comparison.textContent = "No fue posible comparar el WAPE de ambos métodos en este período.";
  }

  document.querySelector("#posttest-category-rows").innerHTML = categories.length
    ? categories.map((category) => {
      const categoryModel = model.by_category?.[category] || {};
      const categoryBaseline = baseline.by_category?.[category] || {};
      const winnerByCategory = posttestWinnerLabel(categoryModel.wape, categoryBaseline.wape);
      return `<tr>
        <td title="${escapeHtml(category)}">${escapeHtml(category)}</td>
        <td>${formatMetricNumber(categoryModel.actual_total)}</td>
        <td>${formatPercent(categoryModel.wape)}</td>
        <td>${formatPercent(categoryBaseline.wape)}</td>
        <td><span class="posttest-winner ${winnerByCategory.className}">${escapeHtml(winnerByCategory.label)}</span></td>
      </tr>`;
    }).join("")
    : '<tr><td colspan="5" class="empty-state">No hay categorías para mostrar.</td></tr>';

  const inventory = data.inventory || {};
  const inventoryNote = document.querySelector("#inventory-result-note");
  const inventoryRows = document.querySelector("#inventory-metric-rows");
  if (inventory.available && Array.isArray(inventory.metrics) && inventory.metrics.length) {
    const inventoryFile = inventory.inventory?.name ? ` Archivo: ${inventory.inventory.name}.` : "";
    const inventoryWarnings = Array.isArray(inventory.warnings) ? inventory.warnings.filter(Boolean) : [];
    const demonstrationNote = inventory.inventory?.is_demonstration
      ? " Este archivo es demostrativo y no debe usarse como evidencia de tesis."
      : "";
    const warningsNote = inventoryWarnings.length ? ` Advertencia: ${inventoryWarnings.join(" ")}` : "";
    inventoryNote.textContent = `${inventory.message || "Indicadores calculados"} ISI = salida frente a stock disponible; TQS = días con quiebre frente a días evaluados.${inventoryFile}${demonstrationNote}${warningsNote}`;
    inventoryRows.innerHTML = inventory.metrics.map((metric) => `<tr>
      <td title="${escapeHtml(metric.category)}">${escapeHtml(metric.category)}</td>
      <td>${formatMetricNumber(metric.CD)}</td>
      <td>${formatPercent(metric.ISI)}</td>
      <td>${formatMetricNumber(metric.DQS)}</td>
      <td>${formatPercent(metric.TQS)}</td>
    </tr>`).join("");
  } else {
    inventoryNote.textContent = inventory.message || "No hay datos de inventario para este período.";
    inventoryRows.innerHTML = '<tr><td colspan="5" class="empty-state">Sin indicadores de inventario para mostrar.</td></tr>';
  }

  downloadPosttestButton.onclick = () => downloadPosttestExcel(data.run_id);
  posttestResults.hidden = false;
  posttestResults.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

async function downloadPosttestExcel(runId) {
  if (!runId) {
    setStatus("No se encontró el identificador del Postest para descargarlo.", "error");
    return;
  }
  setButtonLoading(downloadPosttestButton, true, "Preparando");
  setStatus("Preparando el Excel del Postest…", "working");
  try {
    const response = await fetch("/api/posttest/export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ run_id: runId })
    });
    if (!response.ok) {
      const error = await response.json().catch(() => ({}));
      throw new Error(error.error || "No se pudo preparar el Excel del Postest.");
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const filename = response.headers.get("Content-Disposition")?.match(/filename=([^;]+)/i)?.[1];
    link.href = url;
    link.download = filename?.replaceAll('"', "") || "postest_pronostico.xlsx";
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setStatus("El Excel del Postest se descargó correctamente.", "success");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setButtonLoading(downloadPosttestButton, false);
  }
}

function renderForecast(data, selectedDate = null, shouldScroll = true) {
  const results = document.querySelector("#forecast-results");
  const list = document.querySelector("#prediction-list");
  const entries = Object.entries(data.predictions);
  const productForecastsByCategory = data.product_forecasts_by_category || {};
  const validationByCategory = data.validation_by_category || {};
  const allocationMethod = data.product_allocation_method || "Participación histórica disponible.";
  const availableDates = entries[0]?.[1].map((demand) => demand.date) || [];
  const viewToShow = selectedDate === "period" || availableDates.includes(selectedDate)
    ? selectedDate
    : "period";
  const isPeriodView = viewToShow === "period";
  const dateToShow = isPeriodView ? null : viewToShow;
  const firstDate = availableDates[0];
  const lastDate = availableDates.at(-1);

  currentForecast = data;
  document.querySelector("#forecast-period").textContent = data.cutoff_date
    ? `Prueba histórica · ${data.days} días`
    : `${data.days} días`;
  forecastDateControl.hidden = !availableDates.length;
  forecastDate.innerHTML = [
    `<option value="period">Período completo: ${data.days} días</option>`,
    ...availableDates.map((date) => `<option value="${escapeHtml(date)}">${formatDate(date)}</option>`)
  ].join("");
  forecastDate.value = viewToShow;
  results.hidden = false;
  downloadForecastButton.hidden = !entries.length;
  downloadForecastButton.onclick = () => downloadForecastExcel(data, viewToShow);
  list.innerHTML = entries.length
    ? entries.map(([category, demands]) => {
      const dailyProductForecasts = productForecastsByCategory[category] || demands.map((demand) => ({
        date: demand.date,
        category_quantity: demand.quantity,
        products: []
      }));
      const selectedForecast = dailyProductForecasts.find((forecast) => forecast.date === dateToShow)
        || dailyProductForecasts[0];
      const products = isPeriodView
        ? consolidateProductForecasts(dailyProductForecasts)
        : (selectedForecast?.products || []);
      const purchasePlan = allocateWholePackages(products);
      const consultedPeriod = firstDate && lastDate
        ? `${formatDate(firstDate)} al ${formatDate(lastDate)}`
        : "período seleccionado";
      const purchaseDescription = isPeriodView
        ? `Compra sugerida para los próximos ${data.days} días (${consultedPeriod}): ${numberFormat.format(purchasePlan.length)} productos. Los paquetes enteros se distribuyen sin exceder el total estimado de la categoría.`
        : `Compra sugerida para ${formatDate(selectedForecast?.date || dateToShow)}: ${numberFormat.format(purchasePlan.length)} productos. Los paquetes enteros se asignan sin exceder el total estimado de la categoría.`;
      const productRows = purchasePlan.length
        ? purchasePlan.map((product) => {
          return `
            <li>
              <span class="forecast-product-name" title="${escapeHtml(product.name)}">${escapeHtml(product.name)}</span>
              <span class="forecast-product-category">${escapeHtml(category)}</span>
              <strong>Comprar ${numberFormat.format(product.suggested_packages)} paquetes</strong>
            </li>`;
        }).join("")
        : '<li><span class="forecast-product-name">Sin detalle disponible</span><span class="forecast-product-category">—</span><strong>—</strong></li>';

      return `
        <article class="prediction-card">
          <div class="prediction-card-heading" style="display:flex; align-items:flex-start; justify-content:space-between; flex-wrap:wrap; gap:0.55rem;">
            <h3 style="margin:0 0 0.55rem;">${escapeHtml(category)}</h3>
            <button class="download-forecast-button download-category-forecast-button" style="flex:0 0 auto;" type="button" data-category="${escapeHtml(category)}">Descargar Excel <span aria-hidden="true">↓</span></button>
          </div>
          <p class="prediction-category-label">Pronóstico de demanda por categoría con distribución por producto. ${escapeHtml(allocationMethod)}</p>
          <p class="prediction-validation">Validación histórica: WAPE ${numberFormat.format(validationByCategory[category]?.wape ?? 0)}% · ${escapeHtml(validationByCategory[category]?.method || "método validado")}</p>
          <div class="forecast-selected-day">
            <span><small>${isPeriodView ? "PERÍODO CONSULTADO" : "FECHA SELECCIONADA"}</small>${isPeriodView ? consultedPeriod : formatDate(selectedForecast?.date || dateToShow)}</span>
          </div>
          <div class="period-product-section">
            <p>${purchaseDescription}</p>
            <div class="forecast-product-head"><span>PRODUCTO</span><span>CATEGORÍA</span><span>COMPRA SUGERIDA</span></div>
            <ul class="forecast-product-breakdown">${productRows}</ul>
          </div>
        </article>`;
    }).join("")
    : '<p class="empty-state">No se pudo generar un pronóstico para las categorías disponibles.</p>';
  list.querySelectorAll(".download-category-forecast-button").forEach((button) => {
    button.addEventListener("click", () => {
      downloadForecastExcel(data, viewToShow, button.dataset.category, button);
    });
  });
  if (shouldScroll) results.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

loadDashboard();
loadInventorySummary();
