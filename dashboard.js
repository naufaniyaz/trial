const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
let dashboardState = [];
let activeDepartment = "all";
let activeStatus = "all";
let chartInstances = [];

document.getElementById("file-input").addEventListener("change", handleExcelUpload);
loadPreparedData();

async function loadPreparedData() {
  try {
    const response = await fetch("./data/spi-dashboard-data.json", { cache: "no-store" });
    if (!response.ok) throw new Error("Prepared SPI data was not found.");
    const payload = await response.json();
    buildDashboardFromRows(
      payload.sheets["Monthly Data Entry"] || [],
      payload.sheets["SPI Master"] || []
    );
    document.getElementById("file-name").textContent = payload.workbookName || "Prepared SPI data";
    document.getElementById("loaded-date").textContent = `Loaded: ${formatLoadDate(payload.generatedAt)}`;
  } catch (error) {
    console.warn(error);
    document.getElementById("loaded-date").textContent = "Load the SPI workbook";
    document.getElementById("dashboard").innerHTML = '<div class="empty-state">Use Load Excel or run the VBA export to prepare dashboard data.</div>';
  }
}

function handleExcelUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (loadEvent) => {
    const workbook = XLSX.read(new Uint8Array(loadEvent.target.result), { type: "array" });
    buildDashboardFromRows(sheetToJson(workbook, "Monthly Data Entry"), sheetToJson(workbook, "SPI Master"));
    document.getElementById("file-name").textContent = file.name;
    document.getElementById("loaded-date").textContent = `Loaded: ${new Date().toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" })}`;
  };
  reader.readAsArrayBuffer(file);
}

function sheetToJson(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  return sheet ? XLSX.utils.sheet_to_json(sheet, { defval: null }) : [];
}

function buildDashboardFromRows(monthlyRows, masterRows) {
  const masterByCode = new Map();
  masterRows.forEach((row) => {
    const code = row["SPI Code"];
    if (!code) return;
    masterByCode.set(code, {
      code,
      department: row.Department || "Unmapped",
      riskArea: row["Risk Area"] || "",
      section: row.Section || "",
      name: row["SPI Name"] || code,
      measureType: row["Measure Type"] || "",
      direction: row.Direction || "",
      denominatorBasis: row["Denominator Basis"] || "",
      alert: numberOrNull(row.Alert),
      target: numberOrNull(row.Target),
      source: row["Data Source"] || "Not specified"
    });
  });

  const monthlyByCode = new Map();
  monthlyRows.forEach((row) => {
    const code = row["SPI Code"];
    if (!code) return;
    if (!monthlyByCode.has(code)) monthlyByCode.set(code, []);
    monthlyByCode.get(code).push({
      month: row.Month,
      monthNo: Number(row["Month No"]) || 0,
      value: numberOrNull(row["Status Value"]),
      events: numberOrNull(row.Events),
      originalValue: numberOrNull(row["Original Value"]),
      status: row.Status || "No Target"
    });

    if (!masterByCode.has(code)) {
      masterByCode.set(code, {
        code,
        department: row.Department || "Unmapped",
        riskArea: "",
        section: "",
        name: row["SPI Name"] || code,
        measureType: row["Measure Type"] || "",
        direction: row.Direction || "",
        denominatorBasis: row["Denominator Basis"] || "",
        alert: numberOrNull(row.Alert),
        target: numberOrNull(row.Target),
        source: "Not specified"
      });
    }
  });

  const rows = Array.from(masterByCode.values()).map((spi) => {
    const months = (monthlyByCode.get(spi.code) || [])
      .filter((month) => month.monthNo)
      .sort((a, b) => a.monthNo - b.monthNo);
    return { ...spi, months, status: currentStatus(months) };
  });

  buildDashboard(rows);
}

function numberOrNull(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function currentStatus(months) {
  const month = [...months].reverse().find((item) => item.value !== null && item.value !== undefined);
  return month ? month.status : "No Target";
}

function statusClass(status) {
  const value = String(status || "").toLowerCase();
  if (value === "breach") return "danger";
  if (value === "watch") return "watch";
  if (value === "on target") return "good";
  return "nodata";
}

function buildDashboard(rows) {
  dashboardState = rows;
  activeDepartment = "all";
  activeStatus = "all";
  renderSummary(rows);
  renderMissingData(rows);
  renderFilters(rows);
  renderCards(rows);
}

function renderSummary(rows) {
  const counts = rows.reduce((acc, row) => {
    acc[statusClass(row.status)] += 1;
    return acc;
  }, { danger: 0, watch: 0, good: 0, nodata: 0 });

  document.getElementById("summary-bar").innerHTML = `
    <div class="summary-card"><span class="summary-number">${rows.length}</span><span class="summary-label">Total SPIs</span></div>
    <div class="summary-card danger"><span class="summary-number">${counts.danger}</span><span class="summary-label">Breach</span></div>
    <div class="summary-card watch"><span class="summary-number">${counts.watch}</span><span class="summary-label">Watch</span></div>
    <div class="summary-card good"><span class="summary-number">${counts.good}</span><span class="summary-label">On Target</span></div>
    <div class="summary-card nodata"><span class="summary-number">${counts.nodata}</span><span class="summary-label">No Target / No Data</span></div>
  `;
}

function renderMissingData(rows) {
  const dueCount = currentDueMonthCount();
  const dueMonths = MONTHS.slice(0, dueCount);
  const trackedRows = rows.filter((row) => row.name !== "Date");
  const missing = trackedRows.map((row) => {
    const missingMonths = [];
    for (let monthNo = 1; monthNo <= dueCount; monthNo += 1) {
      const monthRows = row.months.filter((month) => month.monthNo === monthNo);
      const received = monthRows.some(hasEnteredMonthlyValue);
      if (!received) missingMonths.push(MONTHS[monthNo - 1]);
    }
    return { ...row, missingMonths, missingCount: missingMonths.length };
  }).filter((row) => row.missingCount > 0);

  const missingCells = missing.reduce((sum, row) => sum + row.missingCount, 0);
  const totalCells = trackedRows.length * dueCount;
  const complete = totalCells ? Math.round(((totalCells - missingCells) / totalCells) * 100) : 100;

  document.getElementById("missing-subtitle").textContent = `Due months: ${dueMonths.length ? dueMonths.join(", ") : "No months due yet"}. Current and future months are not counted as missing.`;
  document.getElementById("missing-metrics").innerHTML = `
    <div class="metric-card warning"><span class="metric-number">${missingCells}</span><span class="metric-label">Missing Cells</span></div>
    <div class="metric-card warning"><span class="metric-number">${missing.length}</span><span class="metric-label">SPIs Missing</span></div>
    <div class="metric-card good"><span class="metric-number">${complete}%</span><span class="metric-label">Complete</span></div>
  `;

  document.getElementById("missing-table").innerHTML = missing.length ? missing.slice(0, 28).map((row) => `
    <tr>
      <td><strong>${escapeHtml(row.code)}</strong></td>
      <td>${escapeHtml(row.name)}</td>
      <td class="muted">${escapeHtml(row.department)}</td>
      <td class="muted">${escapeHtml(row.source)}</td>
      <td class="count">${row.missingCount}</td>
      <td>${escapeHtml(row.missingMonths.join(", "))}</td>
    </tr>
  `).join("") : `<tr><td colspan="6">All due monthly SPI data has been received.</td></tr>`;
}

function hasEnteredMonthlyValue(month) {
  return month.events !== null && month.events !== undefined
    || month.originalValue !== null && month.originalValue !== undefined
    || month.value !== null && month.value !== undefined;
}

function currentDueMonthCount() {
  const now = new Date();
  const year = now.getFullYear();
  if (year < 2026) return 0;
  if (year > 2026) return 12;
  return Math.max(0, Math.min(12, now.getMonth()));
}

function renderFilters(rows) {
  const departments = [...new Set(rows.map((row) => row.department))].sort();
  document.getElementById("filter-bar").innerHTML = `
    <span class="filter-label">Dept:</span>
    <button class="filter-btn active" data-filter-type="department" data-filter="all">All</button>
    ${departments.map((department) => `<button class="filter-btn" data-filter-type="department" data-filter="${escapeAttribute(department)}">${escapeHtml(department)}</button>`).join("")}
    <span class="filter-split"></span>
    <span class="filter-label">Status:</span>
    <button class="filter-btn active" data-filter-type="status" data-filter="all">All</button>
    <button class="filter-btn" data-filter-type="status" data-filter="danger"><span class="dot danger"></span>Breach</button>
    <button class="filter-btn" data-filter-type="status" data-filter="watch"><span class="dot watch"></span>Watch</button>
    <button class="filter-btn" data-filter-type="status" data-filter="good"><span class="dot good"></span>On Target</button>
    <button class="filter-btn" data-filter-type="status" data-filter="nodata"><span class="dot nodata"></span>No Target</button>
  `;

  document.querySelectorAll(".filter-btn").forEach((button) => {
    button.addEventListener("click", () => {
      const type = button.dataset.filterType;
      if (type === "department") activeDepartment = button.dataset.filter;
      if (type === "status") activeStatus = button.dataset.filter;
      document.querySelectorAll(`.filter-btn[data-filter-type="${type}"]`).forEach((item) => item.classList.remove("active"));
      button.classList.add("active");
      renderCards(dashboardState);
    });
  });
}

function renderCards(rows) {
  const dashboard = document.getElementById("dashboard");
  chartInstances.forEach((chart) => chart.destroy());
  chartInstances = [];

  const filtered = rows.filter((row) => {
    const deptOk = activeDepartment === "all" || row.department === activeDepartment;
    const statusOk = activeStatus === "all" || statusClass(row.status) === activeStatus;
    return deptOk && statusOk;
  });

  const byDepartment = groupBy(filtered, "department");
  dashboard.innerHTML = Object.entries(byDepartment).sort(([a], [b]) => a.localeCompare(b)).map(([department, items]) => `
    <section class="dept-section">
      <div class="dept-header">
        <span class="dept-title">${escapeHtml(department)}</span>
        ${departmentBadges(items)}
      </div>
      <div class="kpi-grid">
        ${items.map(renderCard).join("")}
      </div>
    </section>
  `).join("");

  requestAnimationFrame(() => {
    filtered.forEach(renderChart);
  });
}

function renderCard(row) {
  const cls = statusClass(row.status);
  const hasData = row.months.some((month) => month.value !== null && month.value !== undefined);
  return `
    <article class="kpi-card ${cls}">
      <div class="kpi-top">
        <span class="kpi-code">${escapeHtml(row.code)}</span>
        <span class="status-chip ${cls}">${escapeHtml(row.status)}</span>
      </div>
      <div class="kpi-name">${escapeHtml(row.name)}</div>
      <div class="kpi-tags">
        ${row.measureType ? `<span class="tag">${escapeHtml(row.measureType)}</span>` : ""}
        ${row.denominatorBasis ? `<span class="tag">${escapeHtml(row.denominatorBasis)}</span>` : ""}
        ${row.direction ? `<span class="tag">${escapeHtml(row.direction)}</span>` : ""}
      </div>
      <div class="chart-wrap">
        ${hasData ? `<canvas id="chart-${safeId(row.code)}"></canvas>` : `<div class="no-chart">No data recorded</div>`}
      </div>
    </article>
  `;
}

function renderChart(row) {
  if (typeof Chart === "undefined") return;
  const canvas = document.getElementById(`chart-${safeId(row.code)}`);
  if (!canvas) return;

  const chartMonths = row.months.filter((month) => month.value !== null && month.value !== undefined);
  if (!chartMonths.length) return;

  const chart = new Chart(canvas, {
    type: "line",
    data: {
      labels: chartMonths.map((month) => month.month),
      datasets: [
        {
          label: "Status Value",
          data: chartMonths.map((month) => month.value),
          borderColor: "#e39a64",
          backgroundColor: "rgba(227, 154, 100, 0.08)",
          borderWidth: 2,
          pointRadius: 3,
          pointBackgroundColor: chartMonths.map((month) => {
            const cls = statusClass(month.status);
            if (cls === "danger") return "#e45454";
            if (cls === "watch") return "#d49a28";
            if (cls === "good") return "#28b85d";
            return "#8a91a5";
          }),
          tension: 0.28
        },
        thresholdDataset("Alert", row.alert, "#e45454", [4, 3], chartMonths.length),
        thresholdDataset("Target", row.target, "#28b85d", [6, 4], chartMonths.length)
      ].filter(Boolean)
    },
    options: {
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: "rgba(31, 36, 44, 0.96)",
          borderColor: "rgba(227, 154, 100, 0.3)",
          borderWidth: 1,
          titleColor: "#e39a64",
          bodyColor: "#f1f0ef"
        }
      },
      scales: {
        x: {
          grid: { color: "rgba(241,240,239,0.045)" },
          ticks: { color: "rgba(241,240,239,0.45)", font: { size: 9 }, maxRotation: 0 }
        },
        y: {
          grid: { color: "rgba(241,240,239,0.045)" },
          ticks: { color: "rgba(241,240,239,0.45)", font: { size: 9 }, maxTicksLimit: 4 }
        }
      }
    }
  });
  chartInstances.push(chart);
}

function thresholdDataset(label, value, color, dash, count) {
  if (value === null || value === undefined || Number.isNaN(Number(value))) return null;
  return {
    label,
    data: Array(count).fill(Number(value)),
    borderColor: color,
    borderWidth: 1.4,
    borderDash: dash,
    pointRadius: 0,
    fill: false
  };
}

function departmentBadges(items) {
  const counts = items.reduce((acc, row) => {
    acc[statusClass(row.status)] += 1;
    return acc;
  }, { danger: 0, watch: 0, good: 0, nodata: 0 });

  const labels = {
    danger: "Breach",
    watch: "Watch",
    good: "On Target",
    nodata: "No Target"
  };

  return Object.entries(counts)
    .filter(([, count]) => count)
    .map(([cls, count]) => `<span class="dept-badge status-chip ${cls}">${count} ${labels[cls]}</span>`)
    .join("");
}

function groupBy(rows, key) {
  return rows.reduce((acc, row) => {
    const value = row[key] || "Unmapped";
    if (!acc[value]) acc[value] = [];
    acc[value].push(row);
    return acc;
  }, {});
}

function formatLoadDate(value) {
  const date = value ? new Date(value) : new Date();
  if (Number.isNaN(date.getTime())) return "prepared SPI data";
  return date.toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });
}

function safeId(value) {
  return String(value).replace(/[^a-z0-9_-]/gi, "_");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replace(/`/g, "&#096;");
}
