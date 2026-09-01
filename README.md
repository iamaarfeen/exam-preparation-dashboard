# EXAM Preparation Dashboard

A lightweight static dashboard for tracking a student's syllabus preparation.

## Tech stack

- HTML
- CSS
- Vanilla JavaScript
- SheetJS Community Edition (loaded from CDN)
- Chart.js (loaded from CDN)
- Netlify for hosting

No backend and no database are required.

## Excel format

Create one sheet per subject. Example sheet names:

- Mathematics
- Physics
- Chemistry
- Biology
- History
- Geography
- Economics
- English

The first row of each sheet should contain:

1. Chapter Number
2. Chapter Name
3. Start Date
4. Completion Date
5. Preparation Revision Date

Example:

| Chapter Number | Chapter Name | Start Date | Completion Date | Preparation Revision Date |
|---|---|---|---|---|
| 1 | Algebra | 01/06/2026 | 07/06/2026 | 10/06/2026 |
| 2 | Geometry | 08/06/2026 | 15/06/2026 | 18/06/2026 |
| 3 | Mensuration | 16/06/2026 | | |

Keep all chapters in the workbook from the beginning. Fill dates as the student progresses.

## How it works

Status is calculated automatically:

- No Start Date = Not Started
- Start Date only = In Progress
- Completion Date = Completed
- Preparation Revision Date = Revised

The dashboard calculates:

- Overall syllabus completion
- Revision coverage
- Subject-wise progress
- Chapter status
- Completion journey over time
- Sunday review snapshot
- Search and filters

## Local use

Simply open `index.html` in a modern browser.

If your browser blocks local scripts, use a tiny local server or deploy to Netlify.

## Netlify deployment

The project is static. You can deploy the folder directly with Netlify Drop.

1. Log in to Netlify.
2. Open Netlify Drop: https://app.netlify.com/drop
3. Drag the `exam-preparation-dashboard` folder onto the drop area.
4. Netlify will publish it with a `netlify.app` URL.
5. Open the site and upload your Excel workbook.

There is no build command.

## Updating the dashboard data

The Excel workbook is the source of truth. The current version uses an **Upload Excel** button, so the workbook does not have to be stored on the server.

Whenever you change the Excel file:

1. Save the workbook.
2. Open the dashboard.
3. Upload the updated workbook.
4. The charts and tables recalculate immediately.

## Important note about data privacy

The Excel workbook is processed in the browser. This version does not send the workbook to a custom backend or store it in a database.

The app loads SheetJS and Chart.js from their CDNs. For a fully self-contained deployment, download/vendor those library files locally and change the `<script>` tags in `index.html`.

## Future upgrades

Possible next versions:

- Automatic weekly target planning
- Sunday review report
- Planned vs actual completion
- Days ahead/behind schedule
- Exam countdown
- PDF weekly report
- Teacher/parent login
- Cloud-saved workbook
- Automatic fetching of a workbook from cloud storage
- Multiple students


## Weekly Plan sheet

For planned-vs-actual tracking, add a sheet named `Weekly Plan` with:

| Week Start | Week End | Subject | Target Chapters | Target Revisions |
|---|---|---|---:|---:|
| 01/06/2026 | 07/06/2026 | Mathematics | 1 | 1 |
| 08/06/2026 | 14/06/2026 | Physics | 2 | 1 |

The dashboard compares those targets against completion/revision dates in the subject sheets.

## Board completion forecast

Set the expected first board exam date in the dashboard. The app estimates the syllabus finish date from the student's actual chapter-completion pace and labels the trajectory as:

- Comfortably on track
- On track, but watch pace
- Behind the current pace

The forecast is an estimate, not an official Exam timetable.

The selected exam date is saved only in the browser's local storage.
