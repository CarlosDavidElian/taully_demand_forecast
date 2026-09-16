const statusMessage = document.querySelector("#status");
const reportInput = document.querySelector("#report");
const uploadForm = document.querySelector("#upload-form");
const catalogInput = document.querySelector("#catalog");
const catalogForm = document.querySelector("#catalog-form");
const trainButton = document.querySelector("#train-button");
const forecastButton = document.querySelector("#forecast-button");
const forecastDays = document.querySelector("#days");
const forecastDate = document.querySelector("#forecast-date");
const forecastDateControl = document.querySelector("#forecast-date-control");

let currentForecast = null;

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
  document.querySelector("#category-history-explanation").textContent = summary.records
    ? `Acumulado histórico: suma las unidades vendidas en todos los reportes, desde ${formatDate(summary.first_sale_date)} hasta ${formatDate(summary.last_sale_date)}. No es un pronóstico.`
    : "Aún no hay reportes cargados para calcular el acumulado histórico.";
  document.querySelector("#catalog-products").textContent = numberFormat.format(catalog.products);
  document.querySelector("#catalog-families").textContent = numberFormat.format(catalog.families);
  document.querySelector("#catalog-categories").textContent = numberFormat.format(catalog.categories);
  const bars = document.querySelector("#category-bars");
  if (!categories.length) {
    bars.innerHTML = '<p class="empty-state">Aún no hay demanda cargada. Empieza procesando un reporte.</p>';
  } else {
    const maximum = Math.max(...categories.map((item) => item.quantity), 1);
    bars.innerHTML = categories.slice(0, 8).map((item) => `
      <div class="bar-row">
        <span class="bar-label" title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</span>
        <div class="bar-track"><div class="bar-value" style="width: ${(item.quantity / maximum) * 100}%"></div></div>
        <span class="bar-number">${numberFormat.format(item.quantity)}</span>
      </div>`).join("");
  }

  const table = document.querySelector("#recent-table");
  table.innerHTML = recent.length
    ? recent.map((item) => `<tr><td>${formatDate(item.date)}</td><td title="${escapeHtml(item.category)}">${escapeHtml(item.category)}</td><td>${numberFormat.format(item.quantity)}</td></tr>`).join("")
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
    setStatus(
      `${data.message} Nuevos: ${numberFormat.format(summary.added)} · actualizados: ${numberFormat.format(summary.updated)} · sin cambios: ${numberFormat.format(summary.unchanged)}.${catalogMessage} Entrena nuevamente antes de pronosticar.`,
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
  setButtonLoading(trainButton, true, "Entrenando");
  setStatus("Entrenando el modelo con el historial disponible…", "working");
  try {
    const data = await request("/api/train", { method: "POST" });
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
  setButtonLoading(forecastButton, true, "Calculando");
  setStatus("Calculando el pronóstico por categoría…", "working");
  try {
    const data = await request("/api/forecast", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ days })
    });
    renderForecast(data);
    setStatus(data.message, "success");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setButtonLoading(forecastButton, false);
  }
});

function renderForecast(data, selectedDate = null, shouldScroll = true) {
  const results = document.querySelector("#forecast-results");
  const list = document.querySelector("#prediction-list");
  const entries = Object.entries(data.predictions);
  const productForecastsByCategory = data.product_forecasts_by_category || {};
  const validationByCategory = data.validation_by_category || {};
  const availableDates = entries[0]?.[1].map((demand) => demand.date) || [];
  const dateToShow = availableDates.includes(selectedDate)
    ? selectedDate
    : (availableDates.includes(forecastDate.value) ? forecastDate.value : availableDates[0]);

  currentForecast = data;
  document.querySelector("#forecast-period").textContent = `${data.days} días`;
  forecastDateControl.hidden = !dateToShow;
  forecastDate.innerHTML = availableDates.map((date) => `
    <option value="${escapeHtml(date)}">${formatDate(date)}</option>`).join("");
  forecastDate.value = dateToShow || "";
  results.hidden = false;
  list.innerHTML = entries.length
    ? entries.map(([category, demands]) => {
      const dailyProductForecasts = productForecastsByCategory[category] || demands.map((demand) => ({
        date: demand.date,
        category_quantity: demand.quantity,
        products: []
      }));
      const selectedForecast = dailyProductForecasts.find((forecast) => forecast.date === dateToShow)
        || dailyProductForecasts[0];
      const categoryQuantity = selectedForecast?.category_quantity
        ?? demands.find((demand) => demand.date === dateToShow)?.quantity
        ?? 0;
      const products = selectedForecast?.products || [];
      const productRows = products.length
        ? products.map((product) => `
            <li>
              <span class="forecast-product-name" title="${escapeHtml(product.name)}">${escapeHtml(product.name)}</span>
              <span class="forecast-product-category">${escapeHtml(category)}</span>
              <strong>${numberFormat.format(product.quantity)}</strong>
            </li>`).join("")
        : '<li><span class="forecast-product-name">Sin detalle disponible</span><span class="forecast-product-category">—</span><strong>—</strong></li>';

      return `
        <article class="prediction-card">
          <h3>${escapeHtml(category)}</h3>
          <p class="prediction-category-label">Pronóstico de demanda por producto dentro de la categoría ${escapeHtml(category)} para la fecha seleccionada.</p>
          <p class="prediction-validation">Validación histórica: WAPE ${numberFormat.format(validationByCategory[category]?.wape ?? 0)}% · ${escapeHtml(validationByCategory[category]?.method || "método validado")}</p>
          <div class="forecast-selected-day">
            <span><small>FECHA SELECCIONADA</small>${formatDate(selectedForecast?.date || dateToShow)}</span>
            <strong><small>DEMANDA ESTIMADA · ${escapeHtml(category)}</small>${numberFormat.format(categoryQuantity)} unidades</strong>
          </div>
          <div class="period-product-section">
            <p>Productos estimados para ${formatDate(selectedForecast?.date || dateToShow)}: ${numberFormat.format(products.length)} productos.</p>
            <div class="forecast-product-head"><span>PRODUCTO</span><span>CATEGORÍA</span><span>DEMANDA ESTIMADA</span></div>
            <ul class="forecast-product-breakdown">${productRows}</ul>
          </div>
        </article>`;
    }).join("")
    : '<p class="empty-state">No se pudo generar un pronóstico para las categorías disponibles.</p>';
  if (shouldScroll) results.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

loadDashboard();
