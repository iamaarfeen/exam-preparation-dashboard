
<script>

/* =========================================================
   HOMEWORK DATA
   ========================================================= */

let homeworkRows = [];

let homeworkCharts = {};


/* =========================================================
   HELPERS
   ========================================================= */

const $ = id =>
  document.getElementById(id);


function clean(value) {

  return String(value ?? "").trim();

}


function normalize(value) {

  return clean(value)
    .toLowerCase()
    .replace(/\s+/g, " ");

}


function parseStoredDate(value) {

  if (!value) return null;

  const date = new Date(value);

  return isNaN(date.getTime())
    ? null
    : date;

}


function dateText(date) {

  if (!date) return "—";

  return new Date(date)
    .toLocaleDateString(undefined, {
      day: "2-digit",
      month: "short",
      year: "numeric"
    });

}


function dateOnly(date) {

  if (!date) return null;

  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate()
  );

}


function escapeHtml(value) {

  return String(value ?? "")
    .replace(/[&<>"']/g, c => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#039;"
    }[c]));

}


/* =========================================================
   STATUS
   ========================================================= */

function homeworkStatus(row) {

  if (row.submitted) {

    if (
      row.assignDate &&
      row.submittedDate &&
      row.submittedDate > row.assignDate
    ) {

      return "Submitted";

    }

    return "Submitted";

  }

  return "Pending";

}


/* =========================================================
   LOAD DATA FROM INDEX PAGE
   ========================================================= */

function loadHomeworkData() {

  const stored =
    localStorage.getItem("icseHomeworkData");

  if (!stored) {

    showEmpty();

    return;

  }


  try {

    homeworkRows =
      JSON.parse(stored)
        .map(row => ({

          assignDate:
            parseStoredDate(row.assignDate),

          submissionDate:
            parseStoredDate(row.submissionDate),

          subject:
            clean(row.subject),

          submittedDate:
            parseStoredDate(row.submittedDate),

          points:
            row.points === null ||
            row.points === undefined ||
            row.points === ""
              ? null
              : Number(row.points),

          remark:
            clean(row.remark)

        }));


    homeworkRows =
      homeworkRows.filter(row =>
        row.subject ||
        row.assignDate ||
        row.submissionDate ||
        row.submittedDate
      );


    if (!homeworkRows.length) {

      showEmpty();

      return;

    }


    $("homeworkEmpty")
      .classList
      .add("hidden");

    $("homeworkDashboard")
      .classList
      .remove("hidden");


    $("homeworkUpdated").textContent =
      "Data loaded from latest workbook";


    renderHomework();

  }

  catch (error) {

    console.error(error);

    showEmpty();

  }

}


/* =========================================================
   EMPTY STATE
   ========================================================= */

function showEmpty() {

  $("homeworkEmpty")
    .classList
    .remove("hidden");

  $("homeworkDashboard")
    .classList
    .add("hidden");

}


/* =========================================================
   DESTROY CHARTS
   ========================================================= */

function destroyHomeworkCharts() {

  Object.values(homeworkCharts)
    .forEach(chart => {

      if (chart) chart.destroy();

    });

  homeworkCharts = {};

}


/* =========================================================
   MAIN RENDER
   ========================================================= */

function renderHomework() {

  destroyHomeworkCharts();

  renderHomeworkStats();

  renderSubmissionChart();

  renderSubjectChart();

  renderMarksChart();

  renderSubjectTable();

  renderHomeworkTable();

  renderInsights();

}


/* =========================================================
   STATS
   ========================================================= */

function renderHomeworkStats() {

  const total =
    homeworkRows.length;


  const submitted =
    homeworkRows.filter(row =>
      row.submittedDate
    ).length;


  const pending =
    total - submitted;


  const points =
    homeworkRows
      .filter(row =>
        typeof row.points === "number" &&
        Number.isFinite(row.points)
      )
      .map(row => row.points);


  const average =
    points.length
      ? points.reduce((a, b) => a + b, 0) / points.length
      : null;


  const onTime =
    homeworkRows.filter(row => {

      if (!row.submissionDate ||
          !row.submittedDate) {

        return false;

      }

      return (
        dateOnly(row.submittedDate) <=
        dateOnly(row.submissionDate)
      );

    }).length;


  const late =
    homeworkRows.filter(row => {

      if (!row.submissionDate ||
          !row.submittedDate) {

        return false;

      }

      return (
        dateOnly(row.submittedDate) >
        dateOnly(row.submissionDate)
      );

    }).length;


  const excellent =
    homeworkRows.filter(row =>
      typeof row.points === "number" &&
      row.points >= 9
    ).length;


  const needsImprovement =
    homeworkRows.filter(row =>
      typeof row.points === "number" &&
      row.points <= 5
    ).length;


  const submissionRate =
    total
      ? Math.round(submitted / total * 100)
      : 0;


  const onTimeRate =
    submitted
      ? Math.round(onTime / submitted * 100)
      : 0;


  $("hwTotal").textContent =
    total;

  $("hwTotalDetail").textContent =
    `${total} assignment${total === 1 ? "" : "s"}`;


  $("hwSubmitted").textContent =
    submitted;

  $("hwSubmissionRate").textContent =
    `${submissionRate}% submission rate`;


  $("hwPending").textContent =
    pending;


  $("hwAverage").textContent =
    average === null
      ? "—"
      : average.toFixed(1);


  $("hwOnTime").textContent =
    onTime;

  $("hwOnTimeRate").textContent =
    `${onTimeRate}% of submitted work`;


  $("hwLate").textContent =
    late;


  $("hwExcellent").textContent =
    excellent;


  $("hwNeedsImprovement").textContent =
    needsImprovement;


  $("submissionProgress").style.width =
    `${submissionRate}%`;


  $("submissionProgressText").textContent =
    `${submissionRate}% of assigned homework submitted`;

}


/* =========================================================
   SUBMISSION CHART
   ========================================================= */

function renderSubmissionChart() {

  const submitted =
    homeworkRows.filter(row =>
      row.submittedDate
    ).length;


  const pending =
    homeworkRows.length - submitted;


  homeworkCharts.submission =
    new Chart(
      $("submissionChart"),
      {
        type: "doughnut",

        data: {

          labels: [
            "Submitted",
            "Pending"
          ],

          datasets: [{

            data: [
              submitted,
              pending
            ],

            borderWidth: 0

          }]

        },

        options: {

          responsive: true,

          maintainAspectRatio: false,

          cutout: "68%",

          plugins: {

            legend: {
              position: "bottom"
            }

          }

        }

      }
    );

}


/* =========================================================
   SUBJECT STATS
   ========================================================= */

function getSubjectStats() {

  const stats = {};


  homeworkRows.forEach(row => {

    const subject =
      row.subject || "Unknown";


    if (!stats[subject]) {

      stats[subject] = {

        total: 0,

        submitted: 0,

        pending: 0,

        onTime: 0,

        late: 0,

        points: []

      };

    }


    const s =
      stats[subject];


    s.total++;


    if (row.submittedDate) {

      s.submitted++;

    }
    else {

      s.pending++;

    }


    if (
      row.submissionDate &&
      row.submittedDate
    ) {

      if (
        dateOnly(row.submittedDate) <=
        dateOnly(row.submissionDate)
      ) {

        s.onTime++;

      }
      else {

        s.late++;

      }

    }


    if (
      typeof row.points === "number" &&
      Number.isFinite(row.points)
    ) {

      s.points.push(row.points);

    }

  });


  return stats;

}


/* =========================================================
   SUBJECT CHART
   ========================================================= */

function renderSubjectChart() {

  const stats =
    getSubjectStats();


  const subjects =
    Object.keys(stats).sort();


  const averages =
    subjects.map(subject => {

      const points =
        stats[subject].points;

      if (!points.length) return 0;

      return (
        points.reduce((a, b) => a + b, 0)
        / points.length
      );

    });


  homeworkCharts.subject =
    new Chart(
      $("subjectHomeworkChart"),
      {

        type: "bar",

        data: {

          labels: subjects,

          datasets: [{

            label: "Average Points",

            data: averages,

            borderWidth: 1

          }]

        },

        options: {

          responsive: true,

          maintainAspectRatio: false,

          scales: {

            y: {

              beginAtZero: true,

              max: 10

            },

            x: {

              grid: {
                display: false
              }

            }

          },

          plugins: {

            legend: {
              display: false
            }

          }

        }

      }
    );

}


/* =========================================================
   MARKS DISTRIBUTION
   ========================================================= */

function renderMarksChart() {

  const ranges = {

    "9–10": 0,

    "7–8": 0,

    "6": 0,

    "0–5": 0

  };


  homeworkRows.forEach(row => {

    if (
      typeof row.points !== "number" ||
      !Number.isFinite(row.points)
    ) {

      return;

    }


    if (row.points >= 9) {

      ranges["9–10"]++;

    }
    else if (row.points >= 7) {

      ranges["7–8"]++;

    }
    else if (row.points >= 6) {

      ranges["6"]++;

    }
    else {

      ranges["0–5"]++;

    }

  });


  homeworkCharts.marks =
    new Chart(
      $("marksChart"),
      {

        type: "bar",

        data: {

          labels:
            Object.keys(ranges),

          datasets: [{

            label: "Assignments",

            data:
              Object.values(ranges),

            borderWidth: 1

          }]

        },

        options: {

          responsive: true,

          maintainAspectRatio: false,

          scales: {

            y: {

              beginAtZero: true,

              ticks: {
                precision: 0
              }

            },

            x: {

              grid: {
                display: false
              }

            }

          },

          plugins: {

            legend: {
              display: false
            }

          }

        }

      }
    );

}


/* =========================================================
   SUBJECT TABLE
   ========================================================= */

function renderSubjectTable() {

  const stats =
    getSubjectStats();


  const rows =
    Object.entries(stats)
      .sort(([a], [b]) =>
        a.localeCompare(b)
      );


  $("subjectTable").innerHTML =
    rows.map(([subject, s]) => {

      const average =
        s.points.length
          ? s.points.reduce(
              (a, b) => a + b,
              0
            ) / s.points.length
          : null;


      let performance = "No marks";


      if (average !== null) {

        if (average >= 9) {

          performance = "Excellent";

        }
        else if (average >= 7) {

          performance = "Good";

        }
        else if (average >= 6) {

          performance = "Average";

        }
        else {

          performance = "Needs improvement";

        }

      }


      return `

        <tr>

          <td>
            <strong>
              ${escapeHtml(subject)}
            </strong>
          </td>

          <td>${s.total}</td>

          <td>${s.submitted}</td>

          <td>${s.pending}</td>

          <td>${s.onTime}</td>

          <td>${s.late}</td>

          <td>
            ${
              average === null
                ? "—"
                : average.toFixed(1)
            }
          </td>

          <td>
            <span class="status-badge status-completed">
              ${escapeHtml(performance)}
            </span>
          </td>

        </tr>

      `;

    }).join("");


}


/* =========================================================
   HOMEWORK TABLE
   ========================================================= */

function renderHomeworkTable() {

  const selected =
    $("homeworkSubjectFilter").value;


  const rows =
    homeworkRows.filter(row => {

      if (
        selected !== "ALL" &&
        row.subject !== selected
      ) {

        return false;

      }

      return true;

    });


  $("homeworkTable").innerHTML =
    rows.map(row => {

      const submitted =
        !!row.submittedDate;


      const status =
        submitted
          ? "Submitted"
          : "Pending";


      const statusClass =
        submitted
          ? "status-completed"
          : "status-in-progress";


      return `

        <tr>

          <td>
            ${dateText(row.assignDate)}
          </td>

          <td>
            ${dateText(row.submissionDate)}
          </td>

          <td>
            <strong>
              ${escapeHtml(row.subject)}
            </strong>
          </td>

          <td>
            ${dateText(row.submittedDate)}
          </td>

          <td>
            ${
              row.points === null
                ? "—"
                : `${row.points}/10`
            }
          </td>

          <td>
            ${escapeHtml(row.remark) || "—"}
          </td>

          <td>

            <span
              class="status-badge ${statusClass}">

              ${status}

            </span>

          </td>

        </tr>

      `;

    }).join("")
    ||
    `
      <tr>
        <td colspan="7">
          No homework records found.
        </td>
      </tr>
    `;

}


/* =========================================================
   SUBJECT FILTER
   ========================================================= */

function populateSubjectFilter() {

  const subjects =
    [...new Set(
      homeworkRows
        .map(row => row.subject)
        .filter(Boolean)
    )]
    .sort();


  const filter =
    $("homeworkSubjectFilter");


  subjects.forEach(subject => {

    const option =
      document.createElement("option");

    option.value =
      subject;

    option.textContent =
      subject;

    filter.appendChild(option);

  });

}


/* =========================================================
   INSIGHTS
   ========================================================= */

function renderInsights() {

  const total =
    homeworkRows.length;


  const submitted =
    homeworkRows.filter(row =>
      row.submittedDate
    ).length;


  const pending =
    total - submitted;


  const late =
    homeworkRows.filter(row => {

      if (
        !row.submissionDate ||
        !row.submittedDate
      ) {

        return false;

      }

      return (
        dateOnly(row.submittedDate) >
        dateOnly(row.submissionDate)
      );

    }).length;


  const points =
    homeworkRows
      .filter(row =>
        typeof row.points === "number" &&
        Number.isFinite(row.points)
      )
      .map(row => row.points);


  const average =
    points.length
      ? points.reduce((a, b) => a + b, 0)
        / points.length
      : null;


  const insights = [];


  if (pending > 0) {

    insights.push(
      `There are ${pending} homework assignment${pending === 1 ? "" : "s"} still pending.`
    );

  }
  else if (total > 0) {

    insights.push(
      "All recorded homework assignments have been submitted."
    );

  }


  if (late > 0) {

    insights.push(
      `${late} submission${late === 1 ? "" : "s"} ${late === 1 ? "was" : "were"} submitted after the expected submission date.`
    );

  }
  else if (submitted > 0) {

    insights.push(
      "No late submissions are recorded."
    );

  }


  if (average !== null) {

    if (average >= 9) {

      insights.push(
        `The overall homework average is ${average.toFixed(1)}/10 — excellent performance.`
      );

    }
    else if (average >= 7) {

      insights.push(
        `The overall homework average is ${average.toFixed(1)}/10 — good performance with room to improve.`
      );

    }
    else if (average >= 6) {

      insights.push(
        `The overall homework average is ${average.toFixed(1)}/10 — average performance.`
      );

    }
    else {

      insights.push(
        `The overall homework average is ${average.toFixed(1)}/10 — improvement is needed.`
      );

    }

  }


  const stats =
    getSubjectStats();


  const subjectEntries =
    Object.entries(stats)
      .filter(([, s]) =>
        s.points.length
      );


  if (subjectEntries.length) {

    const best =
      subjectEntries.reduce(
        (best, current) => {

          const bestAvg =
            best[1].points.reduce(
              (a, b) => a + b,
              0
            ) / best[1].points.length;

          const currentAvg =
            current[1].points.reduce(
              (a, b) => a + b,
              0
            ) / current[1].points.length;

          return currentAvg > bestAvg
            ? current
            : best;

        }
      );


    const bestAverage =
      best[1].points.reduce(
        (a, b) => a + b,
        0
      ) / best[1].points.length;


    insights.push(
      `${best[0]} currently has the highest homework average at ${bestAverage.toFixed(1)}/10.`
    );

  }


  $("homeworkInsights").innerHTML =
    insights.map(text => `

      <div class="review-item">

        <span>
          Observation
        </span>

        <strong style="font-size:14px; line-height:1.45;">
          ${escapeHtml(text)}
        </strong>

      </div>

    `).join("");

}


/* =========================================================
   FILTER EVENT
   ========================================================= */

$("homeworkSubjectFilter")
  .addEventListener(
    "change",
    renderHomeworkTable
  );


/* =========================================================
   START
   ========================================================= */

populateSubjectFilter();

loadHomeworkData();

</script>
