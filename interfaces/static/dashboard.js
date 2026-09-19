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

let currentForecast = null;
function selectedForecastCutoff() {
  return forecastCutoff.value || null;
}

function earliestHistoricalCutoff(firstSaleDate) {
  if (!firstSaleDate) return "";
  const firstDate = new Date(`${firstSaleDate}T00:00:00Z`);
  firstDate.setUTCDate(firstDate.getUTCDate() + MIN_HISTORICAL_TRAINING_DAYS - 1);
  return firstDate.toISOString().slice(0, 10);
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

async function request(url, options = {}) {
  const response = await fetch(url, options);
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || "No se pudo completar la operación.");
  return data;
}

function renderDashboard(data) {
  const { summary, categories, recent, catalog } = data;
  document.querySelector("#records").textContent = numberFormat.format(summary.records);
  document.querySelector("#categories-count").textContent = numberFormat.format(summary.categories);
  document.querySelector("#total-quantity").textContent = numberFormat.format(summary.total_quantity);
  document.querySelector("#last-sale-date").textContent = formatDate(summary.last_sale_date);
  document.querySelector("#history-updated-at").textContent = formatDateTime(summary.history_updated_at);
  const firstValidCutoff = earliestHistoricalCutoff(summary.first_sale_date);
  forecastCutoff.min = firstValidCutoff;
  forecastCutoff.max = summary.last_sale_date || "";
  const hasHistoricalDates = Boolean(firstValidCutoff && summary.last_sale_date);
  forecastCutoff.disabled = !hasHistoricalDates || firstValidCutoff > summary.last_sale_date;
  document.querySelector("#forecast-cutoff-note").textContent = !hasHistoricalDates
    ? "Carga ventas históricas para habilitar una prueba histórica."
    : forecastCutoff.disabled
      ? `La prueba histórica estará disponible al completar ${MIN_HISTORICAL_TRAINING_DAYS} días de ventas. Mientras tanto se usará todo el historial.`
      : `Usa esta fecha solo para comprobar qué tan bien predice el modelo. Para una compra futura, déjala vacía. Para comparar con la venta real de un día pasado, selecciona el día anterior: por ejemplo, para revisar el 15/09 elige 14/09. Tus reportes se conservan. Disponible desde ${formatDate(firstValidCutoff)}.`;
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

reportInput.addEventListener("change", () => {
  document.querySelector("#file-name").textContent = reportInput.files[0]?.name || "Seleccionar archivo";
});

catalogInput.addEventListener("change", () => {
  document.querySelector("#catalog-file-name").textContent = catalogInput.files[0]?.name || "Nuevo catálogo .xlsx";
});

forecastDays.addEventListener("change", () => {
  hideForecastResults();
  setStatus("Selecciona «Generar pronóstico» para ver la estimación del nuevo periodo.", "");
});

forecastCutoff.addEventListener("change", () => {
  document.querySelector("#metrics").hidden = true;
  hideForecastResults();
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

function buildForecastExportPayload(data, viewToShow) {
  const entries = Object.entries(data.predictions || {});
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
      cutoff_date: data.cutoff_date || null
    },
    rows
  };
}

async function downloadForecastExcel(data, viewToShow) {
  const payload = buildForecastExportPayload(data, viewToShow);
  if (!payload.rows.length) {
    setStatus("No hay productos pronosticados para descargar.", "error");
    return;
  }

  setButtonLoading(downloadForecastButton, true, "Preparando");
  setStatus("Preparando el Excel del pronóstico…", "working");
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
    link.download = filename?.replaceAll('"', "") || "pronostico_compra.xlsx";
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setStatus("El archivo Excel del pronóstico se descargó correctamente.", "success");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setButtonLoading(downloadForecastButton, false);
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
          <h3>${escapeHtml(category)}</h3>
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
  if (shouldScroll) results.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

loadDashboard();
