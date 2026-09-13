const statusMessage = document.querySelector("#status");
const reportInput = document.querySelector("#report");
const uploadForm = document.querySelector("#upload-form");
const trainButton = document.querySelector("#train-button");
const forecastButton = document.querySelector("#forecast-button");

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

async function request(url, options = {}) {
  const response = await fetch(url, options);
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || "No se pudo completar la operación.");
  return data;
}

function renderDashboard(data) {
  const { summary, products, recent } = data;
  document.querySelector("#records").textContent = numberFormat.format(summary.records);
  document.querySelector("#products-count").textContent = numberFormat.format(summary.products);
  document.querySelector("#total-quantity").textContent = numberFormat.format(summary.total_quantity);
  document.querySelector("#last-sale-date").textContent = formatDate(summary.last_sale_date);
  document.querySelector("#history-updated-at").textContent = formatDateTime(summary.history_updated_at);

  const bars = document.querySelector("#category-bars");
  if (!products.length) {
    bars.innerHTML = '<p class="empty-state">Aún no hay demanda cargada. Empieza procesando un reporte.</p>';
  } else {
    const maximum = Math.max(...products.map((item) => item.quantity), 1);
    bars.innerHTML = products.slice(0, 8).map((item) => `
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
    setStatus(
      `${data.message} Nuevos: ${numberFormat.format(summary.added)} · actualizados: ${numberFormat.format(summary.updated)} · sin cambios: ${numberFormat.format(summary.unchanged)}. Entrena nuevamente antes de pronosticar.`,
      "success"
    );
    document.querySelector("#metrics").hidden = true;
    document.querySelector("#forecast-results").hidden = true;
    uploadForm.reset();
    document.querySelector("#file-name").textContent = "Seleccionar archivo";
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
    document.querySelector("#metric-mape").textContent = `${numberFormat.format(data.metrics.mape)}%`;
    document.querySelector("#metric-rmse").textContent = numberFormat.format(data.metrics.rmse);
    document.querySelector("#metric-mae").textContent = numberFormat.format(data.metrics.mae);
    setStatus(data.message, "success");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setButtonLoading(trainButton, false);
  }
});

forecastButton.addEventListener("click", async () => {
  const days = Number(document.querySelector("#days").value);
  setButtonLoading(forecastButton, true, "Calculando");
  setStatus("Calculando el pronóstico por producto…", "working");
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

function renderForecast(data) {
  const results = document.querySelector("#forecast-results");
  const list = document.querySelector("#prediction-list");
  const entries = Object.entries(data.predictions);
  document.querySelector("#forecast-period").textContent = `${data.days} días`;
  results.hidden = false;
  list.innerHTML = entries.length
    ? entries.map(([category, demands]) => `
        <article class="prediction-card">
          <h3>${escapeHtml(category)}</h3>
          <ul>${demands.map((demand) => `<li><span>${formatDate(demand.date)}</span><strong>${numberFormat.format(demand.quantity)}</strong></li>`).join("")}</ul>
        </article>`).join("")
    : '<p class="empty-state">No se pudo generar un pronóstico para las categorías disponibles.</p>';
  results.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

loadDashboard();
