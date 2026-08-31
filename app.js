const state = {
  rows: [],
  weeklyPlan: [],
  charts: {}
};

const $ = (id) => document.getElementById(id);

function normalizeHeader(value) {
  return String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

function cleanText(value) {
  return String(value ?? "").trim();
}

function parseDate(value) {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date && !isNaN(value)) return value;
  if (typeof value === "number" && window.XLSX?.SSF) {
    const d = XLSX.SSF.parse_date_code(value);
    if (d) return new Date(d.y, d.m - 1, d.d);
  }
  const s = String(value).trim();
  const iso = new Date(s);
  if (!isNaN(iso)) return iso;
  const m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  if (m) {
    const day = Number(m[1]), month = Number(m[2]), year = Number(m[3]);
    const d = new Date(year, month - 1, day);
    if (!isNaN(d)) return d;
  }
  return null;
}

function dateText(date) {
  if (!date) return "—";
  return date.toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "numeric" });
}

function dateKey(date) {
  return date ? new Date(date.getFullYear(), date.getMonth(), date.getDate()).toISOString().slice(0, 10) : null;
}

function statusOf(row) {
  if (row.revision) return "Revised";
  if (row.completion) return "Completed";
  if (row.start) return "In Progress";
  return "Not Started";
}

function statusClass(status) {
  return "status-" + status.toLowerCase().replace(/\s+/g, "-");
}

function destroyCharts() {
  Object.values(state.charts).forEach(c => c?.destroy());
  state.charts = {};
}

function makeRow(subject, raw, index) {
  const normalized = {};
  Object.entries(raw).forEach(([k, v]) => normalized[normalizeHeader(k)] = v);

  const pick = (...keys) => {
    for (const key of keys) {
      const v = normalized[normalizeHeader(key)];
      if (v !== undefined) return v;
    }
    return "";
  };

  const number = cleanText(pick("Chapter Number", "Chapter No", "Chapter #", "No"));
  const name = cleanText(pick("Chapter Name", "Chapter", "Topic"));
  if (!number && !name) return null;

  return {
    subject,
    number,
    name,
    start: parseDate(pick("Start Date", "Started", "Start")),
    completion: parseDate(pick("Completion Date", "Completed Date", "Completion", "Completed")),
    revision: parseDate(pick("Preparation Revision Date", "Revision Date", "Preparation Revision", "Revision")),
    sourceRow: index + 2
  };
}

function workbookToRows(workbook) {
  const all = [];
  state.weeklyPlan = [];
  workbook.SheetNames.forEach(subject => {
    const sheet = workbook.Sheets[subject];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });

    if (normalizeHeader(subject) === "weekly plan") {
      data.forEach((raw, i) => {
        const normalized = {};
        Object.entries(raw).forEach(([k, v]) => normalized[normalizeHeader(k)] = v);
        const pick = (...keys) => {
          for (const key of keys) {
            const v = normalized[normalizeHeader(key)];
            if (v !== undefined) return v;
          }
          return "";
        };
        const weekStart = parseDate(pick("Week Start", "Start Date"));
        const weekEnd = parseDate(pick("Week End", "End Date"));
        const subjectName = cleanText(pick("Subject"));
        const targetChapters = Number(pick("Target Chapters", "Chapters Target", "Target"));
        const targetRevisions = Number(pick("Target Revisions", "Revisions Target", "Revision Target"));
        if (weekStart && (subjectName || targetChapters || targetRevisions)) {
          state.weeklyPlan.push({ weekStart, weekEnd, subject: subjectName || "All subjects",
            targetChapters: Number.isFinite(targetChapters) ? targetChapters : 0,
            targetRevisions: Number.isFinite(targetRevisions) ? targetRevisions : 0 });
        }
      });
      return;
    }

    data.forEach((raw, i) => {
      const row = makeRow(subject, raw, i);
      if (row) all.push(row);
    });
  });
  return all;
}

function loadWorkbook(workbook, fileName = "Workbook") {
  const rows = workbookToRows(workbook);
  if (!rows.length) {
    alert("No chapter rows were found. Check the header names and make sure each subject sheet has Chapter Number and Chapter Name.");
    return;
  }
  state.rows = rows;
  $("emptyState").classList.add("hidden");
  $("dashboard").classList.remove("hidden");
  $("fileName").textContent = fileName;
  $("lastUpdated").textContent = `Loaded ${new Date().toLocaleString()}`;
  render();
}

async function handleFile(file) {
  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    loadWorkbook(workbook, file.name);
  } catch (err) {
    console.error(err);
    alert("Could not read this workbook. Please use .xlsx, .xls or .csv and follow the template.");
  }
}

function demoWorkbook() {
  const demoRows = [
    ["Mathematics", 1, "Commercial Mathematics", "2026-06-01", "2026-06-05", "2026-06-08"],
    ["Mathematics", 2, "Algebra", "2026-06-06", "2026-06-13", "2026-06-16"],
    ["Mathematics", 3, "Geometry", "2026-06-14", "2026-06-22", "2026-06-25"],
    ["Mathematics", 4, "Mensuration", "2026-06-23", "", ""],
    ["Mathematics", 5, "Statistics", "", "", ""],
    ["Mathematics", 6, "Probability", "", "", ""],

    ["Physics", 1, "Force", "2026-06-03", "2026-06-09", "2026-06-12"],
    ["Physics", 2, "Work, Energy and Power", "2026-06-10", "2026-06-18", ""],
    ["Physics", 3, "Machines", "2026-06-19", "", ""],
    ["Physics", 4, "Light", "", "", ""],
    ["Physics", 5, "Sound", "", "", ""],

    ["Chemistry", 1, "Periodic Table", "2026-06-02", "2026-06-07", "2026-06-10"],
    ["Chemistry", 2, "Chemical Bonding", "2026-06-08", "2026-06-15", "2026-06-18"],
    ["Chemistry", 3, "Acids, Bases and Salts", "2026-06-16", "2026-06-24", ""],
    ["Chemistry", 4, "Analytical Chemistry", "", "", ""],
    ["Chemistry", 5, "Organic Chemistry", "", "", ""],

    ["Economics", 1, "The Productive System", "2026-06-04", "2026-06-09", "2026-06-13"],
    ["Economics", 2, "Demand and Supply", "2026-06-10", "2026-06-17", ""],
    ["Economics", 3, "Market", "", "", ""],
    ["Economics", 4, "Public Finance", "", "", ""],

    ["History", 1, "The First War of Independence", "2026-06-01", "2026-06-08", "2026-06-12"],
    ["History", 2, "Growth of Nationalism", "2026-06-09", "2026-06-16", "2026-06-20"],
    ["History", 3, "The Contemporary World", "2026-06-17", "", ""],
    ["History", 4, "Towards Independence", "", "", ""]
  ];

  state.rows = demoRows.map(row => ({
    subject: row[0],
    number: String(row[1]),
    name: row[2],
    start: parseDate(row[3]),
    completion: parseDate(row[4]),
    revision: parseDate(row[5]),
    sourceRow: null
  }));

  state.weeklyPlan = [
    {
      weekStart: parseDate("2026-06-01"),
      weekEnd: parseDate("2026-06-07"),
      subject: "Mathematics",
      targetChapters: 1,
      targetRevisions: 1
    },
    {
      weekStart: parseDate("2026-06-08"),
      weekEnd: parseDate("2026-06-14"),
      subject: "Mathematics",
      targetChapters: 1,
      targetRevisions: 1
    },
    {
      weekStart: parseDate("2026-06-15"),
      weekEnd: parseDate("2026-06-21"),
      subject: "Physics",
      targetChapters: 1,
      targetRevisions: 1
    },
    {
      weekStart: parseDate("2026-06-22"),
      weekEnd: parseDate("2026-06-28"),
      subject: "Chemistry",
      targetChapters: 1,
      targetRevisions: 1
    },
    {
      weekStart: parseDate("2026-06-29"),
      weekEnd: parseDate("2026-07-05"),
      subject: "Economics",
      targetChapters: 1,
      targetRevisions: 1
    }
  ];

  $("emptyState").classList.add("hidden");
  $("dashboard").classList.remove("hidden");

  $("fileName").textContent = "Demo ICSE Preparation Workbook";
  $("lastUpdated").textContent = `Demo loaded ${new Date().toLocaleString()}`;

  render();
}

function render() {
  destroyCharts();
  renderStats();
  renderJourneyChart();
  renderSubjectChart();
  renderStatusChart();
  renderSubjectCards();
  renderTable();
  renderReview();
  renderWeeklyPlan();
  renderForecast();
}

function renderStats() {
  const rows = state.rows;
  const total = rows.length;
  const started = rows.filter(r => r.start).length;
  const completed = rows.filter(r => r.completion).length;
  const revised = rows.filter(r => r.revision).length;
  const subjects = [...new Set(rows.map(r => r.subject))];
  const pct = total ? Math.round(completed / total * 100) : 0;
  const revPct = total ? Math.round(revised / total * 100) : 0;

  $("overallPercent").textContent = pct + "%";
  $("overallDetail").textContent = `${completed} / ${total} chapters completed`;
  $("overallBar").style.width = pct + "%";
  $("chapterTotal").textContent = total;
  $("chapterBreakdown").textContent = `${started} started • ${completed} completed • ${revised} revised`;
  $("subjectTotal").textContent = subjects.length;
  $("subjectBreakdown").textContent = subjects.slice(0, 3).join(" • ") + (subjects.length > 3 ? " • …" : "");
  $("revisionPercent").textContent = revPct + "%";
  $("revisionDetail").textContent = `${revised} / ${total} chapters revised`;
}

function renderJourneyChart() {
  const completedDates = state.rows.filter(r => r.completion).map(r => r.completion).sort((a,b) => a-b);
  const ctx = $("journeyChart");
  if (!completedDates.length) {
    state.charts.journey = new Chart(ctx, {
      type: "line",
      data: { labels: ["No completion dates yet"], datasets: [{ label: "Completed chapters", data: [0], borderWidth: 2, tension: .25 }] },
      options: chartOptions("Completed chapters")
    });
    return;
  }

  const counts = {};
  completedDates.forEach(d => { const k = dateKey(d); counts[k] = (counts[k] || 0) + 1; });
  let running = 0;
  const labels = Object.keys(counts).sort();
  const values = labels.map(k => running += counts[k]);
  const total = state.rows.length;

  state.charts.journey = new Chart(ctx, {
    type: "line",
    data: {
      labels: labels.map(k => new Date(k).toLocaleDateString(undefined, { day: "2-digit", month: "short" })),
      datasets: [
        { label: "Completed chapters", data: values, borderWidth: 3, pointRadius: 4, tension: .25 },
        { label: "Total chapters", data: labels.map(() => total), borderDash: [6,6], borderWidth: 1.5, pointRadius: 0, tension: 0 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { intersect: false, mode: "index" },
      scales: {
        y: { beginAtZero: true, ticks: { precision: 0 } },
        x: { grid: { display: false } }
      },
      plugins: { legend: { position: "bottom" } }
    }
  });
}

function subjectStats() {
  const map = {};
  state.rows.forEach(r => {
    if (!map[r.subject]) map[r.subject] = { total: 0, completed: 0, revised: 0, started: 0 };
    map[r.subject].total++;
    if (r.start) map[r.subject].started++;
    if (r.completion) map[r.subject].completed++;
    if (r.revision) map[r.subject].revised++;
  });
  return map;
}

function renderSubjectChart() {
  const stats = subjectStats();
  const labels = Object.keys(stats);
  state.charts.subject = new Chart($("subjectChart"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        { label: "Completed", data: labels.map(s => stats[s].completed), borderWidth: 1 },
        { label: "Revised", data: labels.map(s => stats[s].revised), borderWidth: 1 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: { y: { beginAtZero: true, ticks: { precision: 0 } }, x: { grid: { display: false } } },
      plugins: { legend: { position: "bottom" } }
    }
  });
}

function renderStatusChart() {
  const statuses = ["Not Started", "In Progress", "Completed", "Revised"];
  const values = statuses.map(s => state.rows.filter(r => statusOf(r) === s).length);
  state.charts.status = new Chart($("statusChart"), {
    type: "doughnut",
    data: { labels: statuses, datasets: [{ data: values, borderWidth: 0 }] },
    options: { responsive: true, maintainAspectRatio: false, cutout: "68%", plugins: { legend: { position: "bottom" } } }
  });
}

function renderSubjectCards() {
  const stats = subjectStats();
  const container = $("subjectCards");
  container.innerHTML = "";
  const filter = $("subjectFilter");
  const previous = filter.value;
  filter.innerHTML = `<option value="ALL">All subjects</option>`;
  Object.keys(stats).sort().forEach(subject => {
    const s = stats[subject];
    const pct = s.total ? Math.round(s.completed / s.total * 100) : 0;
    const opt = document.createElement("option");
    opt.value = subject; opt.textContent = subject; filter.appendChild(opt);

    const card = document.createElement("div");
    card.className = "subject-card";
    card.dataset.subject = subject;
    card.innerHTML = `
      <div class="subject-head"><span class="subject-name">${escapeHtml(subject)}</span><span class="subject-percent">${pct}%</span></div>
      <div class="subject-meta">${s.completed}/${s.total} completed • ${s.revised}/${s.total} revised</div>
      <div class="mini-track"><div class="mini-fill" style="width:${pct}%"></div></div>`;
    card.addEventListener("click", () => {
      filter.value = subject;
      renderTable();
      document.querySelector(".weekly-panel").scrollIntoView({ behavior: "smooth", block: "start" });
    });
    container.appendChild(card);
  });
  if ([...filter.options].some(o => o.value === previous)) filter.value = previous;
}

function renderTable() {
  const subject = $("subjectFilter").value;
  const status = $("statusFilter").value;
  const search = $("searchBox").value.trim().toLowerCase();

  const rows = state.rows.filter(r => {
    if (subject !== "ALL" && r.subject !== subject) return false;
    if (status !== "ALL" && statusOf(r) !== status) return false;
    const hay = `${r.subject} ${r.number} ${r.name}`.toLowerCase();
    return !search || hay.includes(search);
  });

  $("chapterTable").innerHTML = rows.map(r => {
    const s = statusOf(r);
    return `<tr>
      <td><strong>${escapeHtml(r.subject)}</strong></td>
      <td>${escapeHtml(r.number)}</td>
      <td>${escapeHtml(r.name)}</td>
      <td>${dateText(r.start)}</td>
      <td>${dateText(r.completion)}</td>
      <td>${dateText(r.revision)}</td>
      <td><span class="status-badge ${statusClass(s)}">${s}</span></td>
    </tr>`;
  }).join("") || `<tr><td colspan="7">No chapters match the current filters.</td></tr>`;
}


function weekKey(date) {
  if (!date) return "";
  const d = new Date(date);
  const day = d.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + diff);
  return dateKey(d);
}

function actualsForPeriod(start, end, subject) {
  const endDate = end || new Date(start.getTime() + 6*86400000);
  const rows = state.rows.filter(r => {
    if (subject && subject !== "All subjects" && r.subject !== subject) return false;
    return (r.completion && r.completion >= start && r.completion <= endDate) ||
           (r.revision && r.revision >= start && r.revision <= endDate);
  });
  return {
    completed: rows.filter(r => r.completion && r.completion >= start && r.completion <= endDate).length,
    revised: rows.filter(r => r.revision && r.revision >= start && r.revision <= endDate).length
  };
}

function renderWeeklyPlan() {
  const wrap = $("weeklyTableWrap");
  if (!state.weeklyPlan.length) {
    wrap.innerHTML = `<div class="review-item">No Weekly Plan sheet found. Add one to compare weekly targets with actual results.</div>`;
    return;
  }
  const rows = [...state.weeklyPlan].sort((a,b) => a.weekStart-b.weekStart);
  wrap.innerHTML = `<table><thead><tr>
    <th>Week</th><th>Subject</th><th>Target chapters</th><th>Actual completed</th>
    <th>Target revisions</th><th>Actual revised</th><th>Status</th>
  </tr></thead><tbody>${rows.map(r => {
    const a = actualsForPeriod(r.weekStart, r.weekEnd, r.subject);
    const chapterOK = a.completed >= r.targetChapters;
    const revisionOK = a.revised >= r.targetRevisions;
    const status = chapterOK && revisionOK ? "On track" : (a.completed + a.revised > 0 ? "Needs attention" : "Behind");
    const cls = status === "On track" ? "on-track" : status === "Behind" ? "off-track" : "at-risk";
    return `<tr>
      <td>${dateText(r.weekStart)}${r.weekEnd ? " – " + dateText(r.weekEnd) : ""}</td>
      <td><strong>${escapeHtml(r.subject)}</strong></td>
      <td>${r.targetChapters}</td><td>${a.completed}</td>
      <td>${r.targetRevisions}</td><td>${a.revised}</td>
      <td><span class="status-badge ${cls}">${status}</span></td>
    </tr>`;
  }).join("")}</tbody></table>`;
}

function getExamDate() {
  const value = $("examDate").value;
  return value ? new Date(value + "T23:59:59") : null;
}

function renderForecast() {
  const box = $("forecastBox");
  const completions = state.rows.filter(r => r.completion).map(r => r.completion).sort((a,b)=>a-b);
  const total = state.rows.length;
  const completed = completions.length;
  const remaining = total - completed;

  if (!completed) {
    box.innerHTML = `<div class="forecast-note">Enter completion dates in Excel. Once at least one chapter is completed, the dashboard can estimate the student's pace.</div>`;
    return;
  }

  let forecast = null, rate = 0;
  if (completed >= 2) {
    const first = completions[0], last = completions[completions.length - 1];
    const days = Math.max(1, (last-first)/86400000);
    rate = (completed-1) / days;
    forecast = new Date(last.getTime() + (remaining / Math.max(rate, 0.00001))*86400000);
  } else {
    const first = completions[0];
    rate = 1 / Math.max(1, (new Date()-first)/86400000);
  }

  const exam = getExamDate();
  let status = "Set an exam date";
  let statusClassName = "at-risk";
  let note = "Set the board exam target above to compare the forecast with the deadline.";
  if (forecast && exam) {
    const daysEarly = Math.round((exam - forecast)/86400000);
    if (daysEarly >= 21) { status = "Comfortably on track"; statusClassName = "on-track"; }
    else if (daysEarly >= 0) { status = "On track, but watch pace"; statusClassName = "at-risk"; }
    else { status = "Behind the current pace"; statusClassName = "off-track"; }
    note = daysEarly >= 0
      ? `At the current completion pace, the syllabus is forecast to finish about ${daysEarly} day${daysEarly===1?"":"s"} before the exam target.`
      : `At the current completion pace, the syllabus is forecast to finish about ${Math.abs(daysEarly)} day${Math.abs(daysEarly)===1?"":"s"} after the exam target.`;
  }

  const daysLeft = exam ? Math.max(0, Math.ceil((exam - new Date())/86400000)) : null;
  box.innerHTML = `
    <div class="forecast-item"><span class="label">Completed</span><strong>${completed}/${total}</strong></div>
    <div class="forecast-item"><span class="label">Current pace</span><strong>${rate.toFixed(2)} / day</strong></div>
    <div class="forecast-item"><span class="label">Forecast finish</span><strong>${forecast ? dateText(forecast) : "—"}</strong></div>
    <div class="forecast-item"><span class="label">Exam countdown</span><strong>${daysLeft === null ? "—" : daysLeft + " days"}</strong></div>
    <div class="forecast-note ${statusClassName}"><strong>${status}</strong><br>${note}</div>`;
}

function renderReview() {
  const total = state.rows.length;
  const completed = state.rows.filter(r => r.completion).length;
  const revised = state.rows.filter(r => r.revision).length;
  const inProgress = state.rows.filter(r => statusOf(r) === "In Progress").length;
  const notStarted = state.rows.filter(r => statusOf(r) === "Not Started").length;
  const remaining = total - completed;

  $("reviewSnapshot").innerHTML = `
    <div class="review-item"><span>Completed</span><strong>${completed}</strong><span>chapters</span></div>
    <div class="review-item"><span>Still to complete</span><strong>${remaining}</strong><span>chapters</span></div>
    <div class="review-item"><span>In progress</span><strong>${inProgress}</strong><span>chapters</span></div>
    <div class="review-item"><span>Not started</span><strong>${notStarted}</strong><span>chapters</span></div>
    <div class="review-item"><span>Revised</span><strong>${revised}</strong><span>chapters</span></div>
  `;
}

function chartOptions(yLabel) {
  return {
    responsive: true, maintainAspectRatio: false,
    scales: { y: { beginAtZero: true, ticks: { precision: 0 }, title: { display: !!yLabel, text: yLabel } }, x: { grid: { display: false } } },
    plugins: { legend: { position: "bottom" } }
  };
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, c => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#039;" }[c]));
}

$("excelFile").addEventListener("change", e => {
  const file = e.target.files?.[0];
  if (file) handleFile(file);
});
$("demoBtn").addEventListener("click", demoWorkbook);
$("demoBtn2").addEventListener("click", demoWorkbook);
$("subjectFilter").addEventListener("change", renderTable);
$("statusFilter").addEventListener("change", renderTable);
$("searchBox").addEventListener("input", renderTable);
const savedExamDate = localStorage.getItem("icseExamDate");
if (savedExamDate) $("examDate").value = savedExamDate;
$("examDate").addEventListener("change", () => {
  localStorage.setItem("icseExamDate", $("examDate").value);
  renderForecast();
});
