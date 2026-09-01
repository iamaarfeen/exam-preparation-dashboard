const state = {
  rows: [],
  weeklyPlan: [],
  charts: {}
};

const $ = (id) => document.getElementById(id);


/* =========================================================
   BASIC HELPERS
   ========================================================= */

function normalizeHeader(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function cleanText(value) {
  return String(value ?? "").trim();
}

function escapeHtml(value) {
  return String(value ?? "").replace(
    /[&<>"']/g,
    c => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#039;"
    }[c])
  );
}


/* =========================================================
   DATE HELPERS
   ========================================================= */

/*
 * The workbook should contain REAL Excel date cells.
 *
 * SheetJS is loaded with:
 *     cellDates: true
 *
 * Therefore normal Excel dates arrive here as JavaScript Date
 * objects.
 *
 * Numeric Excel serial dates are also supported as a safety
 * measure.
 */

function parseDate(value) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return null;
  }

  /*
   * 1. Real JavaScript Date
   *
   * This is the normal result when SheetJS reads
   * a genuine Excel date with cellDates: true.
   */
  if (
    value instanceof Date &&
    !isNaN(value.getTime())
  ) {
    return new Date(
      value.getFullYear(),
      value.getMonth(),
      value.getDate()
    );
  }

  /*
   * 2. Excel serial date
   *
   * Handles numeric Excel dates safely.
   */
  if (typeof value === "number") {

    const excelEpoch =
      new Date(
        Date.UTC(
          1899,
          11,
          30
        )
      );

    const date =
      new Date(
        excelEpoch.getTime() +
        value * 86400000
      );

    if (!isNaN(date.getTime())) {

      return new Date(
        date.getUTCFullYear(),
        date.getUTCMonth(),
        date.getUTCDate()
      );
    }

    return null;
  }

  /*
   * 3. Valid date text fallback
   *
   * The workbook should preferably contain real Excel
   * dates, but accepting recognizable date text makes
   * the dashboard robust to Excel/SheetJS variations.
   *
   * This does NOT change how your dashboard works.
   */
  const text =
    String(value).trim();

  /*
   * YYYY-MM-DD
   */
  let match =
    text.match(
      /^(\d{4})-(\d{1,2})-(\d{1,2})$/
    );

  if (match) {

    const year =
      Number(match[1]);

    const month =
      Number(match[2]) - 1;

    const day =
      Number(match[3]);

    const date =
      new Date(
        year,
        month,
        day
      );

    if (
      date.getFullYear() === year &&
      date.getMonth() === month &&
      date.getDate() === day
    ) {
      return date;
    }
  }

  /*
   * DD/MM/YYYY
   * DD-MM-YYYY
   * DD.MM.YYYY
   *
   * Important for Indian date format.
   */
  match =
    text.match(
      /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/
    );

  if (match) {

    const day =
      Number(match[1]);

    const month =
      Number(match[2]) - 1;

    const year =
      Number(match[3]);

    const date =
      new Date(
        year,
        month,
        day
      );

    if (
      date.getFullYear() === year &&
      date.getMonth() === month &&
      date.getDate() === day
    ) {
      return date;
    }
  }

  /*
   * Other recognizable date formats.
   */
  const parsed =
    new Date(text);

  if (
    !isNaN(parsed.getTime())
  ) {

    return new Date(
      parsed.getFullYear(),
      parsed.getMonth(),
      parsed.getDate()
    );
  }

  return null;
}

function startOfDay(date) {
  if (!date) {
    return null;
  }

  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate()
  );
}

function daysBetween(startDate, endDate) {
  if (!startDate || !endDate) {
    return 0;
  }

  const start =
    startOfDay(startDate);

  const end =
    startOfDay(endDate);

  return Math.round(
    (
      end.getTime() -
      start.getTime()
    ) / 86400000
  );
}

function dateText(date) {
  if (!date) {
    return "—";
  }

  return date.toLocaleDateString(
    undefined,
    {
      day: "2-digit",
      month: "short",
      year: "numeric"
    }
  );
}

function dateKey(date) {
  if (!date) {
    return null;
  }

  return [
    date.getFullYear(),
    String(
      date.getMonth() + 1
    ).padStart(2, "0"),
    String(
      date.getDate()
    ).padStart(2, "0")
  ].join("-");
}


/* =========================================================
   STATUS
   ========================================================= */

/*
 * IMPORTANT:
 *
 * Overall completion is NOT based on this function.
 *
 * Completion percentage:
 *     completion date exists
 *
 * Revision percentage:
 *     revision date exists
 *
 * Current status:
 *
 *     Revised
 *         ↓
 *     Completed
 *         ↓
 *     In Progress
 *         ↓
 *     Not Started
 */

function statusOf(row) {

  if (row.revision) {
    return "Revised";
  }

  if (row.completion) {
    return "Completed";
  }

  if (row.start) {
    return "In Progress";
  }

  return "Not Started";
}

function statusClass(status) {
  return (
    "status-" +
    status
      .toLowerCase()
      .replace(/\s+/g, "-")
  );
}


/* =========================================================
   EXCEL ROW PARSING
   ========================================================= */

function makeRow(
  subject,
  raw,
  index
) {

  const normalized = {};

  Object.entries(raw).forEach(
    ([key, value]) => {
      normalized[
        normalizeHeader(key)
      ] = value;
    }
  );

  const pick = (...keys) => {

    for (const key of keys) {

      const value =
        normalized[
          normalizeHeader(key)
        ];

      if (
        value !== undefined
      ) {
        return value;
      }
    }

    return "";
  };

  const number =
    cleanText(
      pick(
        "Chapter Number",
        "Chapter No",
        "Chapter #",
        "No"
      )
    );

  const name =
    cleanText(
      pick(
        "Chapter Name",
        "Chapter",
        "Topic"
      )
    );

  if (
    !number &&
    !name
  ) {
    return null;
  }

  return {

    subject,

    number,

    name,

    start:
      parseDate(
        pick(
          "Start Date",
          "Started",
          "Start"
        )
      ),

    completion:
      parseDate(
        pick(
          "Completion Date",
          "Completed Date",
          "Completion",
          "Completed"
        )
      ),

    revision:
      parseDate(
        pick(
          "Preparation Revision Date",
          "Revision Date",
          "Preparation Revision",
          "Revision"
        )
      ),

    sourceRow:
      index + 2
  };
}


/* =========================================================
   WORKBOOK → ROWS
   ========================================================= */

function workbookToRows(workbook) {

  const all = [];

  state.weeklyPlan = [];

  workbook.SheetNames.forEach(
    subject => {

      const sheet =
        workbook.Sheets[subject];

      const data =
        XLSX.utils.sheet_to_json(
          sheet,
          {
            defval: "",
            raw: true
          }
        );

      /*
       * Weekly Plan sheet
       */

      if (
        normalizeHeader(subject) ===
        "weekly plan"
      ) {

        data.forEach(
          raw => {

            const normalized = {};

            Object.entries(raw).forEach(
              ([key, value]) => {

                normalized[
                  normalizeHeader(key)
                ] = value;

              }
            );

            const pick = (...keys) => {

              for (
                const key of keys
              ) {

                const value =
                  normalized[
                    normalizeHeader(key)
                  ];

                if (
                  value !== undefined
                ) {
                  return value;
                }
              }

              return "";
            };

            const weekStart =
              parseDate(
                pick(
                  "Week Start",
                  "Start Date"
                )
              );

            const weekEnd =
              parseDate(
                pick(
                  "Week End",
                  "End Date"
                )
              );

            const subjectName =
              cleanText(
                pick("Subject")
              ) ||
              "All subjects";

            const targetChapters =
              Number(
                pick(
                  "Target Chapters",
                  "Chapters Target",
                  "Target"
                )
              );

            const targetRevisions =
              Number(
                pick(
                  "Target Revisions",
                  "Revisions Target",
                  "Revision Target"
                )
              );

            if (
              weekStart
            ) {

              state.weeklyPlan.push({

                weekStart,

                weekEnd,

                subject:
                  subjectName,

                targetChapters:
                  Number.isFinite(
                    targetChapters
                  )
                    ? targetChapters
                    : 0,

                targetRevisions:
                  Number.isFinite(
                    targetRevisions
                  )
                    ? targetRevisions
                    : 0
              });
            }
          }
        );

        return;
      }

      /*
       * Subject sheet
       */

      data.forEach(
        (raw, index) => {

          const row =
            makeRow(
              subject,
              raw,
              index
            );

          if (row) {
            all.push(row);
          }
        }
      );
    }
  );

  return all;
}


/* =========================================================
   LOAD WORKBOOK
   ========================================================= */

function loadWorkbook(
  workbook,
  fileName = "Workbook"
) {

  const rows =
    workbookToRows(workbook);

  if (!rows.length) {

    alert(
      "No chapter rows were found. " +
      "Check the workbook headers. " +
      "Each subject sheet should contain " +
      "Chapter Number and Chapter Name."
    );

    return;
  }

  state.rows = rows;

  $("emptyState")
    .classList
    .add("hidden");

  $("dashboard")
    .classList
    .remove("hidden");

  $("fileName")
    .textContent =
    fileName;

  $("lastUpdated")
    .textContent =
    `Loaded ${new Date().toLocaleString()}`;

  render();
}


/* =========================================================
   FILE UPLOAD
   ========================================================= */

async function handleFile(file) {

  if (!window.XLSX) {

    alert(
      "Excel reader could not be loaded. " +
      "Please refresh the page and try again."
    );

    return;
  }

  try {

    const buffer =
      await file.arrayBuffer();

    const workbook =
      XLSX.read(
        buffer,
        {
          type: "array",
          cellDates: true
        }
      );

    loadWorkbook(
      workbook,
      file.name
    );

  }
  catch (error) {

    console.error(
      "Workbook loading error:",
      error
    );

    alert(
      "Could not read this workbook. " +
      "Please make sure it is a valid Excel workbook " +
      "and that the date columns contain real Excel dates."
    );
  }
}


/* =========================================================
   DEMO WORKBOOK
   ========================================================= */

function demoWorkbook() {

  const demoRows = [

    [
      "Mathematics",
      1,
      "Commercial Mathematics",
      "2026-06-01",
      "2026-06-05",
      "2026-06-08"
    ],

    [
      "Mathematics",
      2,
      "Algebra",
      "2026-06-06",
      "2026-06-13",
      "2026-06-16"
    ],

    [
      "Mathematics",
      3,
      "Geometry",
      "2026-06-14",
      "2026-06-22",
      "2026-06-25"
    ],

    [
      "Mathematics",
      4,
      "Mensuration",
      "2026-06-23",
      "",
      ""
    ],

    [
      "Mathematics",
      5,
      "Statistics",
      "",
      "",
      ""
    ],

    [
      "Mathematics",
      6,
      "Probability",
      "",
      "",
      ""
    ],

    [
      "Physics",
      1,
      "Force",
      "2026-06-03",
      "2026-06-09",
      "2026-06-12"
    ],

    [
      "Physics",
      2,
      "Work, Energy and Power",
      "2026-06-10",
      "2026-06-18",
      ""
    ],

    [
      "Physics",
      3,
      "Machines",
      "2026-06-19",
      "",
      ""
    ],

    [
      "Physics",
      4,
      "Light",
      "",
      "",
      ""
    ],

    [
      "Physics",
      5,
      "Sound",
      "",
      "",
      ""
    ],

    [
      "Chemistry",
      1,
      "Periodic Table",
      "2026-06-02",
      "2026-06-07",
      "2026-06-10"
    ],

    [
      "Chemistry",
      2,
      "Chemical Bonding",
      "2026-06-08",
      "2026-06-15",
      "2026-06-18"
    ],

    [
      "Chemistry",
      3,
      "Acids, Bases and Salts",
      "2026-06-16",
      "2026-06-24",
      ""
    ],

    [
      "Chemistry",
      4,
      "Analytical Chemistry",
      "",
      "",
      ""
    ],

    [
      "Chemistry",
      5,
      "Organic Chemistry",
      "",
      "",
      ""
    ],

    [
      "Economics",
      1,
      "The Productive System",
      "2026-06-04",
      "2026-06-09",
      "2026-06-13"
    ],

    [
      "Economics",
      2,
      "Demand and Supply",
      "2026-06-10",
      "2026-06-17",
      ""
    ],

    [
      "Economics",
      3,
      "Market",
      "",
      "",
      ""
    ],

    [
      "Economics",
      4,
      "Public Finance",
      "",
      "",
      ""
    ],

    [
      "History",
      1,
      "The First War of Independence",
      "2026-06-01",
      "2026-06-08",
      "2026-06-12"
    ],

    [
      "History",
      2,
      "Growth of Nationalism",
      "2026-06-09",
      "2026-06-16",
      "2026-06-20"
    ],

    [
      "History",
      3,
      "The Contemporary World",
      "2026-06-17",
      "",
      ""
    ],

    [
      "History",
      4,
      "Towards Independence",
      "",
      "",
      ""
    ]

  ];

  state.rows =
    demoRows.map(
      row => ({

        subject:
          row[0],

        number:
          String(row[1]),

        name:
          row[2],

        start:
          parseDemoDate(row[3]),

        completion:
          parseDemoDate(row[4]),

        revision:
          parseDemoDate(row[5]),

        sourceRow:
          null

      })
    );

  state.weeklyPlan = [

    {
      weekStart:
        parseDemoDate(
          "2026-06-01"
        ),

      weekEnd:
        parseDemoDate(
          "2026-06-07"
        ),

      subject:
        "Mathematics",

      targetChapters: 1,

      targetRevisions: 1
    },

    {
      weekStart:
        parseDemoDate(
          "2026-06-08"
        ),

      weekEnd:
        parseDemoDate(
          "2026-06-14"
        ),

      subject:
        "Mathematics",

      targetChapters: 1,

      targetRevisions: 1
    },

    {
      weekStart:
        parseDemoDate(
          "2026-06-15"
        ),

      weekEnd:
        parseDemoDate(
          "2026-06-21"
        ),

      subject:
        "Physics",

      targetChapters: 1,

      targetRevisions: 1
    },

    {
      weekStart:
        parseDemoDate(
          "2026-06-22"
        ),

      weekEnd:
        parseDemoDate(
          "2026-06-28"
        ),

      subject:
        "Chemistry",

      targetChapters: 1,

      targetRevisions: 1
    },

    {
      weekStart:
        parseDemoDate(
          "2026-06-29"
        ),

      weekEnd:
        parseDemoDate(
          "2026-07-05"
        ),

      subject:
        "Economics",

      targetChapters: 1,

      targetRevisions: 1
    }

  ];

  $("emptyState")
    .classList
    .add("hidden");

  $("dashboard")
    .classList
    .remove("hidden");

  $("fileName")
    .textContent =
    "🛑 Demo ICSE Preparation Workbook";

  $("lastUpdated")
    .textContent =
    `Demo loaded ${new Date().toLocaleString()}`;

  render();
}

function parseDemoDate(value) {

  if (!value) {
    return null;
  }

  const parts =
    value.split("-");

  return new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2])
  );
}


/* =========================================================
   MAIN RENDER
   ========================================================= */

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


/* =========================================================
   STATISTICS
   ========================================================= */

function renderStats() {

  const rows =
    state.rows;

  const total =
    rows.length;

  const started =
    rows.filter(
      row => row.start
    ).length;

  /*
   * Completion is based ONLY on Completion Date.
   *
   * A revised chapter is still completed.
   */

  const completed =
    rows.filter(
      row => row.completion
    ).length;

  /*
   * Revision is based ONLY on Revision Date.
   */

  const revised =
    rows.filter(
      row => row.revision
    ).length;

  const subjects =
    [
      ...new Set(
        rows.map(
          row => row.subject
        )
      )
    ];

  const completionPercent =
    total
      ? Math.round(
          completed /
          total *
          100
        )
      : 0;

  const revisionPercent =
    total
      ? Math.round(
          revised /
          total *
          100
        )
      : 0;

  $("overallPercent")
    .textContent =
    completionPercent +
    "%";

  $("overallDetail")
    .textContent =
    `${completed} / ${total} chapters completed`;

  $("overallBar")
    .style.width =
    completionPercent +
    "%";

  $("chapterTotal")
    .textContent =
    total;

  $("chapterBreakdown")
    .textContent =
    `${started} started • ` +
    `${completed} completed • ` +
    `${revised} revised`;

  $("subjectTotal")
    .textContent =
    subjects.length;

  $("subjectBreakdown")
    .textContent =
    subjects
      .slice(0, 3)
      .join(" • ") +
    (
      subjects.length > 3
        ? " • …"
        : ""
    );

  $("revisionPercent")
    .textContent =
    revisionPercent +
    "%";

  $("revisionDetail")
    .textContent =
    `${revised} / ${total} chapters revised`;
}


/* =========================================================
   JOURNEY CHART
   ========================================================= */

function renderJourneyChart() {

  if (!window.Chart) {
    return;
  }

  const canvas =
    $("journeyChart");

  const completedDates =
    state.rows
      .filter(
        row => row.completion
      )
      .map(
        row =>
          row.completion
      )
      .sort(
        (a, b) =>
          a - b
      );

  if (
    !completedDates.length
  ) {

    state.charts.journey =
      new Chart(
        canvas,
        {
          type: "line",

          data: {

            labels: [
              "No completion dates yet"
            ],

            datasets: [

              {
                label:
                  "Completed chapters",

                data: [0],

                borderWidth: 2,

                tension: .25
              }

            ]
          },

          options:
            chartOptions(
              "Completed chapters"
            )
        }
      );

    return;
  }

  const counts = {};

  completedDates.forEach(
    date => {

      const key =
        dateKey(date);

      counts[key] =
        (counts[key] || 0) +
        1;
    }
  );

  let running =
    0;

  const labels =
    Object.keys(counts)
      .sort();

  const values =
    labels.map(
      key => {

        running +=
          counts[key];

        return running;
      }
    );

  const total =
    state.rows.length;

  state.charts.journey =
    new Chart(
      canvas,
      {

        type: "line",

        data: {

          labels:
            labels.map(
              key => {

                const date =
                  parseDate(key);

                return date.toLocaleDateString(
                  undefined,
                  {
                    day: "2-digit",
                    month: "short"
                  }
                );
              }
            ),

          datasets: [

            {
              label:
                "Completed chapters",

              data:
                values,

              borderWidth:
                3,

              pointRadius:
                4,

              tension:
                .25
            },

            {
              label:
                "Total chapters",

              data:
                labels.map(
                  () => total
                ),

              borderDash:
                [6, 6],

              borderWidth:
                1.5,

              pointRadius:
                0,

              tension:
                0
            }

          ]
        },

        options: {

          responsive:
            true,

          maintainAspectRatio:
            false,

          interaction: {

            intersect:
              false,

            mode:
              "index"
          },

          scales: {

            y: {

              beginAtZero:
                true,

              ticks: {

                precision:
                  0
              }
            },

            x: {

              grid: {

                display:
                  false
              }
            }

          },

          plugins: {

            legend: {

              position:
                "bottom"
            }
          }
        }
      }
    );
}


/* =========================================================
   SUBJECT STATISTICS
   ========================================================= */

function subjectStats() {

  const map = {};

  state.rows.forEach(
    row => {

      if (
        !map[row.subject]
      ) {

        map[row.subject] = {

          total:
            0,

          completed:
            0,

          revised:
            0,

          started:
            0
        };
      }

      map[
        row.subject
      ].total++;

      if (row.start) {

        map[
          row.subject
        ].started++;
      }

      if (row.completion) {

        map[
          row.subject
        ].completed++;
      }

      if (row.revision) {

        map[
          row.subject
        ].revised++;
      }
    }
  );

  return map;
}


/* =========================================================
   SUBJECT CHART
   ========================================================= */

function renderSubjectChart() {

  if (!window.Chart) {
    return;
  }

  const stats =
    subjectStats();

  const labels =
    Object.keys(stats)
      .sort();

  state.charts.subject =
    new Chart(
      $("subjectChart"),
      {

        type:
          "bar",

        data: {

          labels,

          datasets: [

            {
              label:
                "Completed",

              data:
                labels.map(
                  subject =>
                    stats[
                      subject
                    ].completed
                ),

              borderWidth:
                1
            },

            {
              label:
                "Revised",

              data:
                labels.map(
                  subject =>
                    stats[
                      subject
                    ].revised
                ),

              borderWidth:
                1
            }

          ]
        },

        options: {

          responsive:
            true,

          maintainAspectRatio:
            false,

          scales: {

            y: {

              beginAtZero:
                true,

              ticks: {

                precision:
                  0
              }
            },

            x: {

              grid: {

                display:
                  false
              }
            }
          },

          plugins: {

            legend: {

              position:
                "bottom"
            }
          }
        }
      }
    );
}


/* =========================================================
   CURRENT STATUS CHART
   ========================================================= */

function renderStatusChart() {

  if (!window.Chart) {
    return;
  }

  const statuses = [
    "Not Started",
    "In Progress",
    "Completed",
    "Revised"
  ];

  const values =
    statuses.map(
      status =>
        state.rows.filter(
          row =>
            statusOf(row) ===
            status
        ).length
    );

  state.charts.status =
    new Chart(
      $("statusChart"),
      {

        type:
          "doughnut",

        data: {

          labels:
            statuses,

          datasets: [

            {
              data:
                values,

              borderWidth:
                0
            }

          ]
        },

        options: {

          responsive:
            true,

          maintainAspectRatio:
            false,

          cutout:
            "68%",

          plugins: {

            legend: {

              position:
                "bottom"
            }
          }
        }
      }
    );
}


/* =========================================================
   SUBJECT CARDS
   ========================================================= */

function renderSubjectCards() {

  const stats =
    subjectStats();

  const container =
    $("subjectCards");

  const filter =
    $("subjectFilter");

  const previous =
    filter.value;

  container.innerHTML = "";

  filter.innerHTML =
    `<option value="ALL">All subjects</option>`;

  Object.keys(stats)
    .sort()
    .forEach(
      subject => {

        const s =
          stats[subject];

        const percent =
          s.total
            ? Math.round(
                s.completed /
                s.total *
                100
              )
            : 0;

        const option =
          document.createElement(
            "option"
          );

        option.value =
          subject;

        option.textContent =
          subject;

        filter.appendChild(
          option
        );

        const card =
          document.createElement(
            "div"
          );

        card.className =
          "subject-card";

        card.dataset.subject =
          subject;

        card.innerHTML = `

          <div class="subject-head">

            <span class="subject-name">
              ${escapeHtml(subject)}
            </span>

            <span class="subject-percent">
              ${percent}%
            </span>

          </div>

          <div class="subject-meta">
            ${s.completed}/${s.total}
            completed •
            ${s.revised}/${s.total}
            revised
          </div>

          <div class="mini-track">

            <div
              class="mini-fill"
              style="width:${percent}%">
            </div>

          </div>
        `;

        card.addEventListener(
          "click",
          () => {

            filter.value =
              subject;

            renderTable();

            const panel =
              document.querySelector(
                ".weekly-panel"
              );

            if (panel) {

              panel.scrollIntoView({
                behavior:
                  "smooth",

                block:
                  "start"
              });
            }
          }
        );

        container.appendChild(
          card
        );
      }
    );

  if (
    [
      ...filter.options
    ].some(
      option =>
        option.value ===
        previous
    )
  ) {

    filter.value =
      previous;
  }
}


/* =========================================================
   CHAPTER TABLE
   ========================================================= */

function renderTable() {

  const subject =
    $("subjectFilter")
      .value;

  const status =
    $("statusFilter")
      .value;

  const search =
    $("searchBox")
      .value
      .trim()
      .toLowerCase();

  const rows =
    state.rows.filter(
      row => {

        if (
          subject !== "ALL" &&
          row.subject !== subject
        ) {
          return false;
        }

        if (
          status !== "ALL" &&
          statusOf(row) !== status
        ) {
          return false;
        }

        const haystack =
          `${row.subject} ` +
          `${row.number} ` +
          `${row.name}`
            .toLowerCase();

        return (
          !search ||
          haystack.includes(search)
        );
      }
    );

  $("chapterTable")
    .innerHTML =
    rows.map(
      row => {

        const status =
          statusOf(row);

        return `

          <tr>

            <td>
              <strong>
                ${escapeHtml(
                  row.subject
                )}
              </strong>
            </td>

            <td>
              ${escapeHtml(
                row.number
              )}
            </td>

            <td>
              ${escapeHtml(
                row.name
              )}
            </td>

            <td>
              ${dateText(
                row.start
              )}
            </td>

            <td>
              ${dateText(
                row.completion
              )}
            </td>

            <td>
              ${dateText(
                row.revision
              )}
            </td>

            <td>

              <span
                class="status-badge ${statusClass(status)}">

                ${status}

              </span>

            </td>

          </tr>

        `;
      }
    ).join("") ||

    `
      <tr>
        <td colspan="7">
          No chapters match the current filters.
        </td>
      </tr>
    `;
}


/* =========================================================
   WEEKLY PLAN
   ========================================================= */

function actualsForPeriod(
  start,
  end,
  subject
) {

  const periodStart =
    startOfDay(start);

  const periodEnd =
    end
      ? startOfDay(end)
      : new Date(
          periodStart.getTime() +
          6 * 86400000
        );

  const rows =
    state.rows.filter(
      row => {

        if (
          subject &&
          subject !==
            "All subjects" &&
          row.subject !==
            subject
        ) {
          return false;
        }

        return true;
      }
    );

  const completed =
    rows.filter(
      row =>
        row.completion &&
        row.completion >=
          periodStart &&
        row.completion <=
          periodEnd
    ).length;

  const revised =
    rows.filter(
      row =>
        row.revision &&
        row.revision >=
          periodStart &&
        row.revision <=
          periodEnd
    ).length;

  return {
    completed,
    revised
  };
}

function renderWeeklyPlan() {

  const wrap =
    $("weeklyTableWrap");

  if (
    !state.weeklyPlan.length
  ) {

    wrap.innerHTML = `

      <div class="review-item">

        <strong>
          No Weekly Plan sheet found.
        </strong>

        <br>

        Add a sheet named
        <strong>Weekly Plan</strong>
        to compare targets with actual results.

      </div>
    `;

    return;
  }

  const rows =
    [
      ...state.weeklyPlan
    ].sort(
      (a, b) =>
        a.weekStart -
        b.weekStart
    );

  wrap.innerHTML = `

    <table>

      <thead>

        <tr>

          <th>Week</th>

          <th>Subject</th>

          <th>Target chapters</th>

          <th>Actual completed</th>

          <th>Target revisions</th>

          <th>Actual revised</th>

          <th>Status</th>

        </tr>

      </thead>

      <tbody>

        ${
          rows.map(
            row => {

              const actual =
                actualsForPeriod(
                  row.weekStart,
                  row.weekEnd,
                  row.subject
                );

              const chapterOK =
                actual.completed >=
                row.targetChapters;

              const revisionOK =
                actual.revised >=
                row.targetRevisions;

              let status;
              let className;

              if (
                chapterOK &&
                revisionOK
              ) {

                status =
                  "On track";

                className =
                  "on-track";

              }
              else if (
                actual.completed > 0 ||
                actual.revised > 0
              ) {

                status =
                  "Needs attention";

                className =
                  "at-risk";

              }
              else {

                status =
                  "Behind";

                className =
                  "off-track";
              }

              return `

                <tr>

                  <td>

                    ${dateText(
                      row.weekStart
                    )}

                    ${
                      row.weekEnd
                        ? " – " +
                          dateText(
                            row.weekEnd
                          )
                        : ""
                    }

                  </td>

                  <td>

                    <strong>
                      ${escapeHtml(
                        row.subject
                      )}
                    </strong>

                  </td>

                  <td>
                    ${row.targetChapters}
                  </td>

                  <td>
                    ${actual.completed}
                  </td>

                  <td>
                    ${row.targetRevisions}
                  </td>

                  <td>
                    ${actual.revised}
                  </td>

                  <td>

                    <span
                      class="status-badge ${className}">

                      ${status}

                    </span>

                  </td>

                </tr>

              `;
            }
          ).join("")
        }

      </tbody>

    </table>
  `;
}


/* =========================================================
   EXAM DATE
   ========================================================= */

function getExamDate() {

  const value =
    $("examDate").value;

  if (!value) {
    return null;
  }

  const parts =
    value.split("-");

  if (
    parts.length !== 3
  ) {
    return null;
  }

  const year =
    Number(parts[0]);

  const month =
    Number(parts[1]) - 1;

  const day =
    Number(parts[2]);

  const date =
    new Date(
      year,
      month,
      day
    );

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month ||
    date.getDate() !== day
  ) {
    return null;
  }

  return date;
}


/* =========================================================
   COMPLETION PACE
   ========================================================= */

function calculateCompletionPace(
  completions
) {

  if (
    completions.length < 2
  ) {
    return null;
  }

  const latest =
    completions[
      completions.length - 1
    ];

  /*
   * Recent 28-day activity.
   */

  const recentStart =
    new Date(
      latest.getTime() -
      28 * 86400000
    );

  const recent =
    completions.filter(
      date =>
        date >= recentStart
    );

  /*
   * If there are at least two recent
   * completion events, use recent pace.
   */

  if (
    recent.length >= 2
  ) {

    const firstRecent =
      recent[0];

    const elapsedDays =
      Math.max(
        7,
        daysBetween(
          firstRecent,
          latest
        )
      );

    return {

      chaptersPerWeek:
        recent.length /
        (elapsedDays / 7),

      basis:
        "recent 28-day activity"
    };
  }

  /*
   * Historical fallback.
   */

  const first =
    completions[0];

  const elapsedDays =
    Math.max(
      7,
      daysBetween(
        first,
        latest
      )
    );

  return {

    chaptersPerWeek:
      (completions.length - 1) /
      (elapsedDays / 7),

    basis:
      "recorded completion history"
  };
}


/* =========================================================
   FORECAST
   ========================================================= */

function renderForecast() {

  const box =
    $("forecastBox");

  const total =
    state.rows.length;

  const completedRows =
    state.rows
      .filter(
        row => row.completion
      )
      .sort(
        (a, b) =>
          a.completion -
          b.completion
      );

  const completions =
    completedRows.map(
      row =>
        row.completion
    );

  const completed =
    completions.length;

  const remaining =
    total -
    completed;

  /*
   * Nothing completed.
   */

  if (
    completed === 0
  ) {

    box.innerHTML = `

      <div class="forecast-note at-risk">

        <strong>
          Forecast unavailable
        </strong>

        <br>

        Enter completion dates in Excel
        before using the completion forecast.

      </div>

    `;

    return;
  }

  const exam =
    getExamDate();

  /*
   * Entire syllabus completed.
   */

  if (
    remaining === 0
  ) {

    const latest =
      completions[
        completions.length - 1
      ];

    let countdown =
      "—";

    if (exam) {

      countdown =
        Math.max(
          0,
          daysBetween(
            new Date(),
            exam
          )
        ) +
        " days";
    }

    box.innerHTML = `

      <div class="forecast-item">

        <span class="label">
          Completed
        </span>

        <strong>
          ${completed}/${total}
        </strong>

      </div>

      <div class="forecast-item">

        <span class="label">
          Remaining
        </span>

        <strong>
          0
        </strong>

      </div>

      <div class="forecast-item">

        <span class="label">
          Syllabus finished
        </span>

        <strong>
          ${dateText(latest)}
        </strong>

      </div>

      <div class="forecast-item">

        <span class="label">
          Exam countdown
        </span>

        <strong>
          ${countdown}
        </strong>

      </div>

      <div class="forecast-note on-track">

        <strong>
          Syllabus completed
        </strong>

        <br>

        All chapters have completion dates.
        The focus can now shift to revision,
        tests and past papers.

      </div>

    `;

    return;
  }

  /*
   * One completed chapter isn't enough
   * to estimate a meaningful pace.
   */

  if (
    completed === 1
  ) {

    let countdown =
      "—";

    if (exam) {

      countdown =
        Math.max(
          0,
          daysBetween(
            new Date(),
            exam
          )
        ) +
        " days";
    }

    box.innerHTML = `

      <div class="forecast-item">

        <span class="label">
          Completed
        </span>

        <strong>
          ${completed}/${total}
        </strong>

      </div>

      <div class="forecast-item">

        <span class="label">
          Remaining
        </span>

        <strong>
          ${remaining}
        </strong>

      </div>

      <div class="forecast-item">

        <span class="label">
          Current pace
        </span>

        <strong>
          Not enough data
        </strong>

      </div>

      <div class="forecast-item">

        <span class="label">
          Exam countdown
        </span>

        <strong>
          ${countdown}
        </strong>

      </div>

      <div class="forecast-note at-risk">

        <strong>
          Forecast needs more data
        </strong>

        <br>

        Complete at least two chapters
        on different dates before the
        dashboard estimates the completion date.

      </div>

    `;

    return;
  }

  const pace =
    calculateCompletionPace(
      completions
    );

  if (
    !pace ||
    !pace.chaptersPerWeek ||
    pace.chaptersPerWeek <= 0
  ) {

    box.innerHTML = `

      <div class="forecast-note at-risk">

        <strong>
          Forecast unavailable
        </strong>

        <br>

        There isn't enough completion history
        to calculate a reliable pace.

      </div>

    `;

    return;
  }

  /*
   * Forecast remaining chapters.
   */

  const weeksRequired =
    remaining /
    pace.chaptersPerWeek;

  const latestCompletion =
    completions[
      completions.length - 1
    ];

  const forecastDate =
    new Date(
      latestCompletion.getTime() +
      weeksRequired *
      7 *
      86400000
    );

  let status =
    "Set an exam date";

  let statusClassName =
    "at-risk";

  let note =
    `Forecast uses ${pace.basis}.`;

  let requiredPerWeek =
    null;

  let paceGap =
    null;

  /*
   * Exam comparison.
   */

  if (exam) {

    const today =
      startOfDay(
        new Date()
      );

    const daysToExam =
      daysBetween(
        today,
        exam
      );

    const daysEarly =
      daysBetween(
        forecastDate,
        exam
      );

    /*
     * Exam date has already passed.
     */

    if (
      daysToExam <= 0
    ) {

      status =
        "Exam target has passed";

      statusClassName =
        "off-track";

      requiredPerWeek =
        null;

      paceGap =
        null;

      note =
        `The exam target has already passed. ` +
        `At the current completion pace, the syllabus ` +
        `is forecast to finish about ` +
        `${Math.abs(daysEarly)} day` +
        `${Math.abs(daysEarly) === 1 ? "" : "s"} ` +
        `after the exam target.`;
    }

    /*
     * Exam is still in the future.
     */

    else {

      const weeksToExam =
        daysToExam /
        7;

      requiredPerWeek =
        remaining /
        weeksToExam;

      paceGap =
        pace.chaptersPerWeek -
        requiredPerWeek;

      if (
        daysEarly >= 30
      ) {

        status =
          "Comfortably on track";

        statusClassName =
          "on-track";

      }
      else if (
        daysEarly >= 0
      ) {

        status =
          "On track, but watch pace";

        statusClassName =
          "at-risk";

      }
      else {

        status =
          "Behind the current pace";

        statusClassName =
          "off-track";
      }

      if (
        daysEarly >= 0
      ) {

        note =
          `At the current pace, the syllabus ` +
          `is forecast to finish about ` +
          `${daysEarly} day` +
          `${daysEarly === 1 ? "" : "s"} ` +
          `before the exam target.`;

      }
      else {

        note =
          `At the current pace, the syllabus ` +
          `is forecast to finish about ` +
          `${Math.abs(daysEarly)} day` +
          `${Math.abs(daysEarly) === 1 ? "" : "s"} ` +
          `after the exam target.`;
      }
    }
  }

  /*
   * Countdown.
   */

  let countdown =
    "—";

  if (exam) {

    countdown =
      Math.max(
        0,
        daysBetween(
          new Date(),
          exam
        )
      ) +
      " days";
  }

  box.innerHTML = `

    <div class="forecast-item">

      <span class="label">
        Completed
      </span>

      <strong>
        ${completed}/${total}
      </strong>

    </div>

    <div class="forecast-item">

      <span class="label">
        Remaining
      </span>

      <strong>
        ${remaining}
      </strong>

    </div>

    <div class="forecast-item">

      <span class="label">
        Current pace
      </span>

      <strong>
        ${pace.chaptersPerWeek.toFixed(2)}/week
      </strong>

    </div>

    <div class="forecast-item">

      <span class="label">
        Forecast finish
      </span>

      <strong>
        ${dateText(forecastDate)}
      </strong>

    </div>

    <div class="forecast-item">

      <span class="label">
        Exam countdown
      </span>

      <strong>
        ${countdown}
      </strong>

    </div>

    <div class="forecast-note ${statusClassName}">

      <strong>
        ${status}
      </strong>

      <br>

      ${note}

      ${
        requiredPerWeek !== null
          ? `

            <br><br>

            <strong>
              Required pace:
            </strong>

            ${requiredPerWeek.toFixed(2)}
            chapters/week

            &nbsp; • &nbsp;

            <strong>
              Pace gap:
            </strong>

            ${
              paceGap >= 0
                ? "+"
                : ""
            }

            ${paceGap.toFixed(2)}
            chapters/week

          `
          : ""
      }

    </div>

  `;
}


/* =========================================================
   SUNDAY REVIEW
   ========================================================= */

function renderReview() {

  const total =
    state.rows.length;

  const completed =
    state.rows.filter(
      row =>
        row.completion
    ).length;

  const revised =
    state.rows.filter(
      row =>
        row.revision
    ).length;

  const inProgress =
    state.rows.filter(
      row =>
        statusOf(row) ===
        "In Progress"
    ).length;

  const notStarted =
    state.rows.filter(
      row =>
        statusOf(row) ===
        "Not Started"
    ).length;

  $("reviewSnapshot")
    .innerHTML = `

      <div class="review-item">

        <span>
          Completed
        </span>

        <strong>
          ${completed}
        </strong>

        <span>
          chapters
        </span>

      </div>

      <div class="review-item">

        <span>
          Still to complete
        </span>

        <strong>
          ${total - completed}
        </strong>

        <span>
          chapters
        </span>

      </div>

      <div class="review-item">

        <span>
          In progress
        </span>

        <strong>
          ${inProgress}
        </strong>

        <span>
          chapters
        </span>

      </div>

      <div class="review-item">

        <span>
          Not started
        </span>

        <strong>
          ${notStarted}
        </strong>

        <span>
          chapters
        </span>

      </div>

      <div class="review-item">

        <span>
          Revised
        </span>

        <strong>
          ${revised}
        </strong>

        <span>
          chapters
        </span>

      </div>

    `;
}


/* =========================================================
   CHART OPTIONS
   ========================================================= */

function chartOptions(
  yLabel
) {

  return {

    responsive:
      true,

    maintainAspectRatio:
      false,

    scales: {

      y: {

        beginAtZero:
          true,

        ticks: {

          precision:
            0
        },

        title: {

          display:
            Boolean(yLabel),

          text:
            yLabel
        }
      },

      x: {

        grid: {

          display:
            false
        }
      }
    },

    plugins: {

      legend: {

        position:
          "bottom"
      }
    }
  };
}


/* =========================================================
   DESTROY CHARTS
   ========================================================= */

function destroyCharts() {

  Object.values(
    state.charts
  ).forEach(
    chart => {

      if (chart) {
        chart.destroy();
      }

    }
  );

  state.charts = {};
}


/* =========================================================
   EVENT LISTENERS
   ========================================================= */

$("excelFile")
  .addEventListener(
    "change",
    event => {

      const file =
        event.target
          .files?.[0];

      if (file) {
        handleFile(file);
      }
    }
  );


$("demoBtn")
  .addEventListener(
    "click",
    demoWorkbook
  );


$("demoBtn2")
  .addEventListener(
    "click",
    demoWorkbook
  );


$("subjectFilter")
  .addEventListener(
    "change",
    renderTable
  );


$("statusFilter")
  .addEventListener(
    "change",
    renderTable
  );


$("searchBox")
  .addEventListener(
    "input",
    renderTable
  );


/* =========================================================
   EXAM DATE — LOCAL STORAGE
   ========================================================= */

const savedExamDate =
  localStorage.getItem(
    "icseExamDate"
  );

if (savedExamDate) {

  $("examDate").value =
    savedExamDate;
}


$("examDate")
  .addEventListener(
    "change",
    () => {

      const selectedDate =
        $("examDate").value;

      if (
        selectedDate
      ) {

        localStorage.setItem(
          "icseExamDate",
          selectedDate
        );

      }
      else {

        localStorage.removeItem(
          "icseExamDate"
        );
      }

      /*
       * Only the forecast depends directly
       * on the exam date.
       */
      renderForecast();
    }
  );
