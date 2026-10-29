# Cathedral Sub App — PowerSchool Plugin (`v2.0.0`)

This directory contains the PowerSchool plugin package that exposes the two PowerQuery endpoints used by the Cathedral Substitute Management App:

1. **`com.cathedral.subapp.masterschedule`**
   - **Endpoint**: `POST /ws/schema/query/com.cathedral.subapp.masterschedule?pagesize=0`
   - **Payload**: `{"target_term": "3601"}`
   - **Purpose**: Pulls distinct teacher, period, room, and combined co-seated course names for the active term from `CC`, `COURSES`, `TEACHERS`, and `SECTIONS`.

2. **`com.cathedral.subapp.bellschedule`**
   - **Endpoint**: `POST /ws/schema/query/com.cathedral.subapp.bellschedule?pagesize=0`
   - **Payload**: `{"target_term": "3601"}`
   - **Purpose**: Pulls all Monday–Friday calendar days for the full school year (`YEARID = FLOOR(:target_term / 100)`, `ISYEARREC = 1`) from `CALENDAR_DAY`, `TERMS`, `CYCLE_DAY`, and `BELL_SCHEDULE`, mapping each date to `Green`, `Blue`, `Gold`, or `Red` (No School).

---

## Directory Structure

```text
powerschool-plugin/
├── Cathedral-Sub-App-PowerSchool-Plugin.zip   # Ready-to-upload PowerSchool plugin bundle
├── plugin.xml                                 # Plugin manifest & Data Access Request declarations
├── queries_root/
│   └── subapp.named_queries.xml               # PowerQuery definitions
└── README.md
```

---

## Installing or Updating in PowerSchool

1. Sign in to the **PowerSchool Admin Portal**.
2. Navigate to **System** (or **Setup**) > **System Settings** > **Plugin Management Configuration**.
3.- **If installing for the first time**:
    - Click **Install**, select `Cathedral-Sub-App-PowerSchool-Plugin.zip`, and click **Install**.
    - Check the **Enable/Disable** box next to **Cathedral Sub App Integration** and confirm the data access permissions.
  - **If updating an existing plugin**:
    - Click on your existing plugin name, click **Update**, upload `Cathedral-Sub-App-PowerSchool-Plugin.zip`, and approve the new table access permissions (`CALENDAR_DAY`, `CYCLE_DAY`, `BELL_SCHEDULE`, `TERMS`).
4. Click **Data Provider Configuration** on the plugin page to verify or copy the **Client ID** and **Client Secret** configured in your Sub App **Settings > General Settings** (`PS_URL`, `PS_CLIENT_ID`, `PS_CLIENT_SECRET`).

---

## Rebuilding the `.zip` Archive

PowerSchool requires `plugin.xml` and `queries_root/` to be at the root of the `.zip` archive (with no parent folder wrapper and no macOS `__MACOSX` metadata). To rebuild:

```bash
cd powerschool-plugin
zip -r -X Cathedral-Sub-App-PowerSchool-Plugin.zip plugin.xml queries_root/subapp.named_queries.xml
```
