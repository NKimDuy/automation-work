# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Vietnamese educational automation system that synchronizes student enrollment data from **LSA** (learning management info system) to **LMS** (Moodle-based system at OU) using Selenium and MySQL.

## Running the Workflows

**Setup:**
```bash
pip install -r requirements.txt
```

Configure `config/settings.py` with credentials before running (see Configuration section).

**Run student enrollment sync:**
```bash
python -m automation.add_students_to_lms.run
```

**Run LMS performance report:**
```bash
python -m automation.report_perform_lms.run
```

Both run files instantiate their class at the bottom and call the main method directly — no CLI flags needed.

## Architecture

### Data Flow

```
LSA (web scraping) ──┐
                      ├──> MySQL (automatic_work.students_lms)
API (api.ou.edu.vn) ─┘
         │
         └──> PHDT (web scraping) ──> LMS (Selenium enrollment update)
```

### Key Modules

- **`modules/init_selenium.py` → `InitSelenium`** — Central Selenium wrapper. All browser automation starts here.
  - `init_selenium()` — Creates Chrome headless driver with anti-detection settings
  - `login_selenium(url)` — SSO login flow for both LMS and PHDT
  - `process_get_detail_lsa(semester, url_lsa)` — Scrapes subject table from LSA, returns list of dicts with keys: `lms_id`, `subject_id`, `subject_name`, `teacher_id`, `teacher_name`, `group_id`, etc.
  - `process_get_detail_phdt()` — Scrapes PHDT timetable, returns dict keyed by `"subject_id-group_id"`

- **`utils/api.py` → `APIHandler`** — REST client for `api.ou.edu.vn`
  - `get_unit()` — Returns `[[MaDP, TenDP], ...]`
  - `get_subject_from_api(semester, unit_id)` — Returns list of subject dicts

- **`db/connect_mysql.py` → `get_connection()`** — Returns a `mysql.connector` connection. Caller is responsible for closing it.

- **`utils/logger.py` → `setup_logger(file_path)`** — Creates a UTF-8 file logger. Each run overwrites the previous log file.

### Workflows

**`automation/add_students_to_lms/run.py` → `UpdateStudents`**
1. Calls `process_get_detail_lsa()` to get subject list
2. Logs into LMS via `login_selenium()`
3. For each subject, navigates to the enrollment update URL and reads the enrollment count text (Vietnamese: "Đã ghi danh X/Y sinh viên")
4. Extracts numbers with regex and inserts into `automatic_work.students_lms`

**`automation/report_perform_lms/run.py` → `ReportPerformLMS`**
1. Calls `get_unit()` then `get_subject_from_api()` per unit
2. Filters subjects by date range (`TUNGAYTKB`)
3. Currently calls `process_get_detail_phdt()` in the `test()` method

## Configuration

`config/settings.py` holds all credentials (gitignored). Variables used across the project:

| Variable | Used In | Description |
|----------|---------|-------------|
| `user_lms`, `password_lms`, `captcha` | `init_selenium.py` | LMS SSO login |
| `host`, `user_db`, `password_db`, `database` | `db/connect_mysql.py` | MySQL connection |
| `authorization` | `utils/api.py` | Bearer token for `api.ou.edu.vn` |
| `semester` | multiple | Current semester code e.g. `"251"` (year+term) |

The `semester` value format: `"251"` maps to the LMS URL path and API param `nhhk=20251`.

## Database

Table: `automatic_work.students_lms`

```sql
CREATE TABLE automatic_work.students_lms (
    id INT AUTO_INCREMENT PRIMARY KEY,
    lms_id VARCHAR(50),
    group_id VARCHAR(20),
    subject_id VARCHAR(50),
    subject_name VARCHAR(255),
    teacher_id VARCHAR(50),
    teacher_name VARCHAR(255),
    date_update DATETIME,
    number_student_update INT,
    semester VARCHAR(20)
);
```

## Conventions

- Code and comments are in **Vietnamese**
- Module-level instantiation pattern: each `run.py` creates an instance and calls the method at the bottom of the file (not inside `if __name__ == "__main__"`)
- Selenium waits use `WebDriverWait` with 10–40 second timeouts; errors are caught per-step and printed/logged without stopping execution
- Log files are at `automation/<workflow>/` and overwritten each run
