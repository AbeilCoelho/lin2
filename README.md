# 📚 Enterprise Data Lineage Manager - Architecture & Function Documentation

## 🏗️ System Overview
The application is a Flask-based web server using Pandas for heavy data manipulation. It utilizes a **Breadth-First Search (BFS) Background Engine** to automatically trace data lineage across multiple Excel files. It relies on strict thread-locking (`threading.RLock`) to allow parallel asynchronous processing while preventing file corruption or race conditions.

---

## 🐍 Backend Functions (`app.py`)

### 1. State Management (Thread-Safe JSON)
These functions manage the `auto_resolve_state.json` file, which tracks the exact status of every CDE (Pending, Processing, Needs Review, Completed) and stores the active branch queues.
* **`make_json_serializable(obj)`**: Recursively cleans complex Pandas/Numpy data types (like `NaT`, `Timestamp`, or sets) into safe Python dictionaries/lists so the state can be saved as a string without crashing.
* **`load_state(project_name)`**: Thread-safe read of the project's state JSON. Returns an empty dictionary if the file doesn't exist or fails to parse.
* **`save_state(project_name, state)`**: Thread-safe write of the current memory dictionary into the JSON file.
* **`update_state_for_key(project_name, key, data_dict)`**: Atomic read-modify-write operation. Locks the file, loads state, updates a specific CDE instance, and saves. Prevents race conditions.
* **`delete_state_for_key(project_name, key)`**: Safely removes a CDE instance from the state file (used when a user clicks "Retrace").

### 2. Smart File Memory
* **`track_file_usage(project_name, file_path)`**: Increments a counter in the JSON state tracking how many times a specific Excel file has been successfully mapped.
* **`get_file_memory(project_name)`**: Retrieves the dictionary of historical file usage to apply scoring weights during future candidate searches.

### 3. Core Search Engine
* **`flexible_normalize(val)`**: Strips special characters and date strings (e.g., `YYYYMMDD`, `YYYY-MM-DD`) from table/column names to allow "Flexible Matching".
* **`search_dataframe(df, app_code, table, col)`**: The hierarchical search logic.
  1. Checks for an **Exact Match**.
  2. If none, strips dates/symbols and checks for a **Flexible Match**.
  3. If none, uses Python's `difflib` to find a **Fuzzy Match** (misspellings) with a >30% similarity threshold. Returns the matching rows and the "Match Type".
* **`execute_search_api_logic(project_name, cde_name, app_code, table, col, target_apps)`**: Opens the `primary_lineage` and `global_lineage` files. Groups matches by file and origin, calculates a "Score" based on Exact/Fuzzy matching, checks if it hits a Target App, and adds a massive score boost using `file_memory`. Returns a sorted list of candidate groups.

### 4. Auto-Resolve (Breadth-First Search Engine)
* **`process_instance_background(project_name, cde_name, target_key, initial_raw_data)`**: The core automation worker running on a separate thread.
  * **Goal:** Trace a single instance all the way to a target application without human intervention.
  * **Logic:** Uses a BFS queue (`stacks_to_process`). It searches for the current node. 
    * If it finds an **Exact Match** in a **single file**, it auto-resolves, adds the new node(s) to the queue, and loops again. (If the match splits into 2 sources, it pushes *both* paths into the queue).
    * **Safety Stops:** If it hits a *Dead End*, a *Fuzzy Match*, or *Multiple Conflicting Files*, it immediately aborts, saves the active stack + pending queues to JSON, and flags the UI as `NEEDS_REVIEW`.
* **`write_lineage_to_files(project_name, cde_name, stack, is_dead_end)`**: Takes a successfully completed branch stack, translates it into "Current" and "Source" pairs, and safely appends the rows to `outputs/[CDE_NAME]_Consolidation_Workbook.xlsx` and `reporting_layers.xlsx` "Done" sheet.

### 5. API & Flask Routes
* **`set_project`**: Switches the active workspace directory in the Flask session.
* **`process_upload`**: Saves target, primary, and global Excel files into the active project's `uploads/` folder.
* **`dashboard`**: Calculates high-level metrics (Total, Completed, Pending, Processing) by combining the Target CDEs with the current JSON state.
* **`api_retrace` / `edit_trace`**: Removes a completed trace from the "Done" Excel sheet and moves it back into the active queue. `edit_trace` pushes the saved trace stack into the UI so the user can literally "undo" past the mistake.
* **`save_midway`**: Grabs the active stack and pending queues from the frontend and saves them to JSON as `NEEDS_REVIEW` so the user can safely close the app and resume tomorrow.

---

## 🌐 Frontend Functions (Alpine.js & Vanilla JS)

### 1. Dashboard (`dashboard.html`)
* **`startPolling()` / `refreshUI()`**: Uses DOM Swapping. Every 1.5 seconds, it fetches the HTML of the dashboard in the background, extracts the Table Body and Metrics widgets, and injects them into the current screen. Allows real-time progress viewing without screen flashing. Stops polling when `Processing == 0`.

### 2. Manual Tracing UI (`trace.html`)
Managed by the Alpine.js `tracingApp()` component:
* **`initTrace()`**: Uses Jinja's `| tojson` to safely inject the Python stack, pending branches, and completed paths into the Javascript memory.
* **`searchCandidates()`**: Calls the backend search API and renders the Candidate Cards.
* **`getStitchedValues(cand)`**: Calculates the Transformation Output (Keep Searched, Keep Found, or Concatenate) based on the user's radio-button selection during a Fuzzy Match.
* **`selectCandidateGroup(cand)`**: User confirmed a manual step. If the candidate has multiple sources (a branch), it pushes the extra paths into the `pendingBranches` queue, moves the active path forward, and fetches the next hop.
* **`submitManualSource()`**: Appends a user's typed App/Table/Col into the stack when they hit a Dead End.
* **`finishLineage()`**: Moves the active stack into `completedPaths`. If `pendingBranches` has data, it pulls the next branch out and continues tracing. If all queues are empty, it posts everything to the backend to generate the Excel file.

### 3. Graph Visualizer (`view_lineage.html`)
* **`initGraph(data)`**: Renders the Vis.js canvas. **The Magic Trick:** It initializes using a strict `Right-to-Left (RL)` hierarchical layout so the Reporting Node is always on the far right. Once drawn, it captures the X/Y coordinates, applies them permanently, and *turns off the physics engine*. This allows the user to drag nodes freely anywhere on the screen while maintaining the horizontal tree shape.
* **`createAppSVG(meta)`**: A dynamic SVG generator. Reads backend metadata to draw custom App Nodes, injecting the "Green Man" (Manual), "Yellow Circle" (Transform), or "Red Trapezoid" (Filter) directly onto the graph canvas.


### Input data

| Report name | CDE name | Business Term in Collibra | CDE Definition in CDE Memo | Business Term Domain | Product | Line of Business | Tier | Destination to (App Code) | Destination to (App Name) | Lineage Flow (App Codes) | Current app code | Current app name | Current App Ownership (RBC Owned or Vendor Managed) | Current App Type (System / EUC / System) | Cluster 7 ID# | Created/Sourced | Manually entered/Derived | Screen Name (UI) | Screen Field Name (UI) | Sub-DE | Database System/File System (Host Name/Instance Name) | Database/Directory | Schema | Current table/file name | Current column/field name | Data Element Field Type | Field Description | Frequency | Source app code | Source app name | App Ownership (RBC Owned or Vendor Managed) | Source Database System/File System (Host Name/Instance Name) | Field Database/Directory | Source Schema | Source table/file name | Source column/field name | Source Field Type | Source Field Definition | Transformed/Passed Through | Transformation Description | Transformation Logic | Filtered | Filtration Description | Filtration Logic | BUSINESS_TERM Collibra ID | BUSINESS_PROCESS Collibra IDs | BUSINESS_PROCESS Domains | file | column | group_id | current_key | source_key | modification_date | FFIEC 002, FR Y-9C, FFIEC 009, FR Y-15 | Source Booking Transit ID |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| `REPORT_001` | `CDE_001` | `Business Term A` | `Sample definition for critical data element A` | `Domain A` | `Product A` | `Line of Business A` | `Tier 1` | `APP100` | `Application Alpha` | `APP100 -> APP200` | `APP200` | `Application Beta` | `RBC Owned` | `System` | `CL-1001` | `Sourced` | `N/A` | `Screen A` | `Field A` | `SubDE_A` | `dbhost01.example.com` | `/data/source_a` | `SCHEMA_A` | `SOURCE_TABLE_A` | `SOURCE_FIELD_A` | `Data File Field` | `Sample source field description` | `Daily` | `APP300` | `Application Gamma` | `Vendor Managed` | `sourcehost01.example.com` | `/incoming/data_a` | `SRC_SCHEMA_A` | `SOURCE_FILE_A_YYYYMMDD.csv` | `FIELD_A` | `String` | `Sample source field definition` | `Passed Through` | `N/A` | `N/A` | `Not Filtered` | `N/A` | `N/A` | `BT-100001` | `BP-200001` | `Domain A` | `SOURCE_FILE_A` | `FIELD_A` | `GRP-001` | `CUR-001` | `SRC-001` | `2026-05-01 11:09:57` | `REPORT_A` | `TRANSIT_001` |
| `REPORT_002` | `CDE_002` | `Business Term B` | `Sample definition for critical data element B` | `Domain B` | `Product B` | `Line of Business B` | `Tier 1` | `APP110` | `Application Bravo` | `APP110 -> APP210` | `APP210` | `Application Charlie` | `RBC Owned` | `System` | `CL-1002` | `Sourced` | `N/A` | `Screen B` | `Field B` | `SubDE_B` | `dbhost02.example.com` | `/data/source_b` | `SCHEMA_B` | `SOURCE_TABLE_B` | `SOURCE_FIELD_B` | `Data File Field` | `Sample source field description` | `Daily` | `APP310` | `Application Delta` | `Vendor Managed` | `sourcehost02.example.com` | `/incoming/data_b` | `SRC_SCHEMA_B` | `SOURCE_FILE_B_YYYYMMDD.csv` | `FIELD_B` | `String` | `Sample source field definition` | `Passed Through` | `N/A` | `N/A` | `Filtered` | `Records matching predefined criteria` | `FIELD_B = 'VALUE_B'` | `BT-100002` | `BP-200002` | `Domain B` | `SOURCE_FILE_B` | `FIELD_B` | `GRP-002` | `CUR-002` | `SRC-002` | `2026-05-02 12:15:43` | `REPORT_B` | `TRANSIT_002` |
| `REPORT_003` | `CDE_003` | `Business Term C` | `Sample definition for critical data element C` | `Domain C` | `Product C` | `Line of Business C` | `Tier 1` | `APP120` | `Application Echo` | `APP120 -> APP220` | `APP220` | `Application Foxtrot` | `RBC Owned` | `System` | `CL-1003` | `Sourced` | `N/A` | `N/A` | `N/A` | `N/A` | `dbhost03.example.com` | `/data/source_c` | `SCHEMA_C` | `SOURCE_TABLE_C` | `SOURCE_FIELD_C` | `Data File Field` | `Business field containing product information` | `Daily` | `APP320` | `Application Golf` | `RBC Owned` | `sourcehost03.example.com` | `/incoming/data_c` | `SRC_SCHEMA_C` | `SOURCE_FILE_C_YYYYMMDD.csv` | `FIELD_C` | `String` | `Business product classification value` | `Passed Through` | `No transformation applied` | `SOURCE_FIELD_C` | `Not Filtered` | `N/A` | `N/A` | `BT-100003` | `BP-200003` | `Domain C` | `SOURCE_FILE_C` | `FIELD_C` | `GRP-003` | `CUR-003` | `SRC-003` | `2026-05-03 09:42:18` | `REPORT_C` | `TRANSIT_003` |
| `REPORT_004` | `CDE_004` | `Business Term D` | `Sample definition for critical data element D` | `Domain D` | `Product D` | `Line of Business D` | `Tier 2` | `APP130` | `Application Hotel` | `APP130 -> APP230` | `APP230` | `Application India` | `RBC Owned` | `System` | `CL-1004` | `Sourced` | `N/A` | `N/A` | `N/A` | `N/A` | `dbhost04.example.com` | `/data/source_d` | `SCHEMA_D` | `SOURCE_TABLE_D` | `SOURCE_FIELD_D` | `Data File Field` | `Product identifier used for classification` | `Daily` | `APP330` | `Application Juliet` | `RBC Owned` | `sourcehost04.example.com` | `/incoming/data_d` | `SRC_SCHEMA_D` | `SOURCE_FILE_D_YYYYMMDD.csv` | `FIELD_D` | `String` | `Source product identifier` | `Passed Through` | `Value passed directly to destination` | `SOURCE_FIELD_D` | `Filtered` | `Only eligible source records are retained` | `SOURCE_FIELD_D IN ('A','B')` | `BT-100004` | `BP-200004` | `Domain D` | `SOURCE_FILE_D` | `FIELD_D` | `GRP-004` | `CUR-004` | `SRC-004` | `2026-05-04 14:22:31` | `REPORT_D` | `TRANSIT_004` |
| `REPORT_005` | `CDE_005` | `Business Term E` | `Sample definition for critical data element E` | `Domain E` | `Product E` | `Line of Business E` | `Tier 1` | `APP140` | `Application Kilo` | `APP140 -> APP240` | `APP240` | `Application Lima` | `Vendor Managed` | `System` | `CL-1005` | `Sourced` | `Derived` | `N/A` | `N/A` | `N/A` | `dbhost05.example.com` | `/data/source_e` | `SCHEMA_E` | `SOURCE_TABLE_E` | `SOURCE_FIELD_E` | `Data File Field` | `Calculated business indicator` | `Daily` | `APP340` | `Application Mike` | `Vendor Managed` | `sourcehost05.example.com` | `/incoming/data_e` | `SRC_SCHEMA_E` | `SOURCE_FILE_E_YYYYMMDD.csv` | `FIELD_E` | `String` | `Source value used to derive indicator` | `Transformed` | `Source value converted to standardized code` | `CASE WHEN FIELD_E = 'X' THEN 'Y' ELSE 'Z' END` | `Not Filtered` | `N/A` | `N/A` | `BT-100005` | `BP-200005` | `Domain E` | `SOURCE_FILE_E` | `FIELD_E` | `GRP-005` | `CUR-005` | `SRC-005` | `2026-05-05 16:08:52` | `REPORT_E` | `TRANSIT_005` |
| `REPORT_006` | `CDE_006` | `Business Term F` | `Sample definition for critical data element F` | `Domain F` | `Product F` | `Line of Business F` | `Tier 2` | `APP150` | `Application November` | `APP150 -> APP250` | `APP250` | `Application Oscar` | `RBC Owned` | `EUC` | `CL-1006` | `Created` | `Manually entered` | `Input Screen F` | `Input Field F` | `SubDE_F` | `dbhost06.example.com` | `/data/source_f` | `SCHEMA_F` | `SOURCE_TABLE_F` | `SOURCE_FIELD_F` | `Data File Field` | `Manually maintained business attribute` | `Weekly` | `APP350` | `Application Papa` | `RBC Owned` | `sourcehost06.example.com` | `/incoming/data_f` | `SRC_SCHEMA_F` | `SOURCE_FILE_F_YYYYMMDD.csv` | `FIELD_F` | `String` | `Manually maintained source value` | `Passed Through` | `No transformation applied` | `SOURCE_FIELD_F` | `Filtered` | `Only active records are included` | `STATUS = 'ACTIVE'` | `BT-100006` | `BP-200006` | `Domain F` | `SOURCE_FILE_F` | `FIELD_F` | `GRP-006` | `CUR-006` | `SRC-006` | `2026-05-06 10:17:06` | `REPORT_F` | `TRANSIT_006` |
| `REPORT_007` | `CDE_007` | `Business Term G` | `Sample definition for critical data element G` | `Domain G` | `Product G` | `Line of Business G` | `Tier 1` | `APP160` | `Application Quebec` | `APP160 -> APP260` | `APP260` | `Application Romeo` | `RBC Owned` | `System` | `CL-1007` | `Sourced` | `Derived` | `N/A` | `N/A` | `N/A` | `dbhost07.example.com` | `/data/source_g` | `SCHEMA_G` | `SOURCE_TABLE_G` | `SOURCE_FIELD_G` | `Data File Field` | `Derived product subtype attribute` | `Daily` | `APP360` | `Application Sierra` | `RBC Owned` | `sourcehost07.example.com` | `/incoming/data_g` | `SRC_SCHEMA_G` | `SOURCE_FILE_G_YYYYMMDD.csv` | `FIELD_G` | `String` | `Source product subtype` | `Transformed` | `Source subtype mapped to standardized classification` | `MAP(FIELD_G)` | `Filtered` | `Excluded records based on business rules` | `FIELD_G NOT IN ('TYPE_A','TYPE_B')` | `BT-100007` | `BP-200007` | `Domain G` | `SOURCE_FILE_G` | `FIELD_G` | `GRP-007` | `CUR-007` | `SRC-007` | `2026-05-07 13:51:29` | `REPORT_G` | `TRANSIT_007` |
| `REPORT_008` | `CDE_008` | `Business Term H` | `Sample definition for critical data element H` | `Domain H` | `Product H` | `Line of Business H` | `Tier 1` | `APP170` | `Application Tango` | `APP170 -> APP270` | `APP270` | `Application Uniform` | `Vendor Managed` | `System` | `CL-1008` | `Sourced` | `Derived` | `N/A` | `N/A` | `N/A` | `dbhost08.example.com` | `/data/source_h` | `SCHEMA_H` | `SOURCE_TABLE_H` | `SOURCE_FIELD_H` | `Data File Field` | `Business product subtype` | `Daily` | `APP370` | `Application Victor` | `Vendor Managed` | `sourcehost08.example.com` | `/incoming/data_h` | `SRC_SCHEMA_H` | `SOURCE_FILE_H_YYYYMMDD.csv` | `FIELD_H` | `String` | `Source product subtype definition` | `Passed Through` | `No transformation applied` | `SOURCE_FIELD_H` | `Filtered` | `Records failing eligibility rules are excluded` | `FIELD_H IS NOT NULL` | `BT-100008` | `BP-200008` | `Domain H` | `SOURCE_FILE_H` | `FIELD_H` | `GRP-008` | `CUR-008` | `SRC-008` | `2026-05-08 15:36:44` | `REPORT_H` | `TRANSIT_008` |



---
---

### 🚀 HOW TO RESUME IN A NEW CHAT:
When you are ready to keep coding in a few hours/days, open a new chat window and paste this exact prompt. It will instantly give me all the context I need:

> **System Context Prompt:**
> "We are building an Enterprise Data Lineage Manager using Flask, Pandas, Alpine.js, and Vis.js. 
> **Architecture Rules:**
> 1. Multi-project support: Files live in `workspace/<project_name>/uploads/` and `outputs/`.
> 2. Thread-safety: A `threading.RLock()` protects all reads/writes to `auto_resolve_state.json` and Excel files to prevent Pandas race conditions.
> 3. State Management: The JSON file tracks active stacks, pending queues (for multi-branch scenarios), and completed paths.
> 4. Auto-Resolve: Runs in the background using a Breadth-First Search (BFS) queue. It automatically processes Exact/Flexible matches but safely halts on Fuzzy Matches or Dead Ends, flagging them as `NEEDS_REVIEW` for the user.
> 5. UI: Uses DOM swapping for live dashboard updates, Alpine.js for trace queue management, and Vis.js with a physics-freeze trick for free-draggability.
> 
> I have the latest code working perfectly. as you can see on the files below. My next feature request is: [INSERT YOUR REQUEST HERE]
>
> [FILES STATE]