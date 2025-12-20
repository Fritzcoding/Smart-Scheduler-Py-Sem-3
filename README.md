# Schedule Creator

Schedule Creator is a Python tool for generating fair, constraint-aware activity schedules for groups (originally built for a summer camp / student association setting). It includes a small PyQt GUI for configuration and exports schedules in multiple formats (Word, JSON, CSV, PNG) with visualizations.

## The Problem

- Organizers need to create weekly activity schedules for multiple groups while keeping assignments fair and diverse.
- Manually balancing constraints (no simultaneous duplicate activities across groups, limited repeats per week, spacing repeated activities) is time-consuming and error-prone.

Example use case: Indonesian Student Association events where each group should experience events fairly and not repeatedly at the same time.

## Objective

- Provide a small, practical tool to automatically produce weekly group schedules that respect user-configurable constraints and export them in common formats for distribution.
- Make exports flexible so a user can request just one output (Word, CSV, JSON, PNG, pie chart) without running the full generator when desired.

## Schedule Rules

Each schedule follows a set of rules intended to keep group assignments fair and varied. The scheduler attempts to satisfy these rules whenever possible; when constraints make a strict solution impossible it will either fail with a helpful message or attempt safe adjustments (for example increasing per-activity usage limits).

- **No group may have the same activity more than 3 times per week (2 if possible).**
  - Rationale: encourage variety within each group's week.
  - Enforced by: the `max_activity_uses` limit checked in `valid()`; `try_auto_adjust_and_solve()` may relax this when necessary for feasibility.

- **No group may have the same activity more than once on the same day.**
  - Rationale: avoid repetition within a single day.
  - Enforced by: the solver fills per-day slots and `valid()` prevents assigning the same activity twice to the same group on a single day.

- **No two groups can have the same activity at the same period/day.**
  - Rationale: many activities use shared resources (instructors, courts) and must not overlap.
  - Enforced by: `valid()` checks other groups' assignments for the same (period, day) and disallows duplicates.

- **If an activity occurs twice in a week for the same group, there should be at least one day gap.**
  - Rationale: spread repeated experiences apart to avoid clustering.
  - Enforced by: the solver tracks existing placements and disallows adjacent-day repeats when possible; this is implemented as part of the validity checks and placement heuristics.

- **If an activity occurs more than twice in a week it should alternate the part of the day (morning vs afternoon).
  - Rationale: balance daily distribution and avoid back-to-back same-part-of-day repeats.
  - Enforced by: placement rules in `valid()` that consider period indices and prefer opposite-day-part slots for repeated activities.

Notes on strictness: these rules are treated as hard constraints where necessary (e.g., no simultaneous duplicate across groups) and soft constraints elsewhere (e.g., the 2-vs-3-per-week preference). The solver uses backtracking and a small auto-adjust helper to find feasible schedules under tight capacity.

## Details

Each week staff create schedules for the groups expected the following week. The primary goal is diversity: each group should experience as many different activities as possible. Doing this manually for many groups (often up to 10) is time-consuming and combinatorially complex — similar to solving a constrained Sudoku where many cross-group and within-group constraints must be balanced.

## Limitations

- Some combinations of constraints are impossible. Example: 10 groups but only 8 distinct activities available for the same period — because activities cannot run simultaneously across groups, the schedule cannot be generated.
- The solver may auto-relax some soft limits (e.g., per-activity usage) to find a feasible solution; if it still fails, the user should add activities or reduce periods.

## System Architecture / Overview

- UI layer: PyQt5 GUI (`bin/ui.py`) — collects configuration (groups, activities, periods per day, week selection) and exposes export buttons.
- Scheduling core: backtracking solver implemented in the UI module — fills a 3D matrix [group][period][day] while enforcing constraints.
- Export layer: `bin/word.py` and UI handlers — write Word documents (`python-docx`), JSON, CSV, Pillow-based PNG, and Matplotlib charts.

Data flow:
- User config → build matrix template → run solver (optional) → export routines (Word/JSON/CSV/Image/Pie)

## Code Structural Flow & Output Explanation

- Entry point: `bin/ui.py` (run the GUI from the `bin` folder with `python ui.py`).
- When generating a schedule: `generate_weekly_matrix()` creates an empty matrix sized by groups × periods × days. The solver (`solve()` / `valid()`) fills the matrix.
- Exports:
  - Word: `word.make_word_doc()` (full export) or `word.make_word_doc_only()` (Word-only).
  - JSON/CSV: written with `json` and `csv` modules; structure is Group → Period → [days].
  - Image (Pillow): text-based schedule image.
  - Matplotlib visualizations: bar chart (activity frequency) and pie chart (activity distribution). Pie chart is now available independently and falls back to the available activity list if no assignments exist.

Generated files are saved to the `Generated Schedules/` folder with names like `Week 1 Schedules.docx`, `Week 1_schedule.json`, `Week 1_pie.png`, etc.

## Data Structures Design

- Matrix representation: list[list[list[str]]] — outer list = groups, middle list = periods, inner list = days (strings for activity names). Example: matrix[group_idx][period_idx][day_idx] = "Soccer".
- Activities: simple list of strings.
- Groups: dictionary mapping group name to participant list for UI display; scheduling uses the group count primarily.

## Core Logic & Algorithm

- The core is a recursive backtracking solver:
  1. `find_empty()` finds the next empty slot (group, period, day).
  2. For each shuffled activity choice, `valid()` checks constraints:
     - Activity not used by another group at the same (period, day).
     - Activity use count for the group does not exceed `max_activity_uses`.
  3. If valid, set activity and recurse; otherwise backtrack.

- The UI contains a `try_auto_adjust_and_solve()` helper that increases `max_activity_uses` (when feasible) to allow the solver to succeed in tight capacities.

## Core Features & Technologies

- GUI: PyQt5 (user inputs groups, activities, number of periods per day).
- Scheduling algorithm: backtracking with simple heuristics (randomized choices to reduce bias).
- Exports and libraries:
  - `python-docx` (`docx`) — Word document creation and table filling.
  - `Pillow` — schedule image generation.
  - `Matplotlib` — bar charts and pie charts (pie chart improved and exportable independently).
  - `OpenCV` (optional) — image processing utilities (included but not required for basic use).
  - `NumPy` — basic statistics in analysis screen.
  - `json`, `csv` — native exports.

## Format-Specific Design Decisions

- Word (.docx): Uses a template (`2019 Template Schedules.docx`) with a pre-formatted table for Group 1; additional groups copy that table and the script fills table cells. A Word-only exporter (`make_word_doc_only`) avoids generating other formats when the user requests a single output.
- JSON: Exports a Group → Period → Days hierarchical structure for programmatic consumption.
- CSV: Flattened rows per Group/Period, with day columns for simple spreadsheet viewing.
- PNG (Pillow): Simple, readable text-based image per group (good for quick sharing).
- Charts: Matplotlib bar and pie charts provide quick visual summaries. The pie exporter will create a chart even when no schedule has been generated by falling back to the configured activity list.

## Usage

1. Open a terminal and change into the `bin` folder:

```powershell
cd Schedule-Creator-master/bin
pip install -r requirements.txt
python ui.py
```

2. Configure activities and groups in the UI.
3. Set `Periods per Day` (capped at 6) and select the week.
4. Click `Generate Schedule` to produce all outputs, or use individual export buttons (`Word`, `JSON`, `CSV`, `Image`, `Pie`) to export just one file.

## Notes & Troubleshooting

- If you need to export a single file without generating the full schedule, use the corresponding export button — the app will create a blank template matrix from current groups and periods and export using that template.
- If the solver fails, try adding more activities or reducing periods; the UI may auto-adjust activity usage limits to attempt to find a solution.

## Contributing

PRs welcome. Please keep UI changes isolated to `bin/ui.py` and export logic in `bin/word.py`.

---
*Updated: Dec 2025*
