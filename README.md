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

Report name
CDE name
Business Term in Collibra
CDE Definition in CDE Memo
Business Term Domain
Product
Line of Business
Tier
Destination to (App Code)
Destination to (App Name)
Lineage Flow (App Codes)
Current app code
Current app name
Current App Ownership (RBC Owned or Vendor Managed)
Current App Type (System / EUC / System)
Cluster 7 ID#
Created/Sourced
Manually entered/Derived
Screen Name (UI)
Screen Field Name (UI)
Sub-DE
Database System/File System (Host Name/Instance Name)
Database/Directory
Schema
Current table/file name
Current column/field name
Data Element Field Type
Field Description
Frequency
Source app code
Source app name
App Ownership (RBC Owned or Vendor Managed)
Source Database System/File System (Host Name/Instance Name)
Field Database/Directory
Source Schema
Source table/file name
Source column/field name
Source Field Type
Source Field Definition
Transformed/Passed Through
Transformation Description
Transformation Logic
Filtered
Filtration Description
Filtration Logic
BUSINESS_TERM Collibra ID
BUSINESS_PROCESS Collibra IDs
BUSINESS_PROCESS Domains
file
column
group_id
current_key
source_key
modification_date
FFIEC 002, FR Y-9C, FFIEC 009, FR Y-15
Source Booking Transit ID









9200
URF











PH3.VURFAD_ADDRESS 
AD_UNIT_NO






















./FFIEC 002\URF - 9200\2022.DLQ.FRB.FFIEC002-FR2420.9200.xlsm
E


||
2023-11-29 14:06:21
FR 2052a
Maturity Bucket









9200
URF



Sourced
N/A






PH3.VURFAD_ADDRESS 
AD_UNIT_NO



WPN0
FiBRS




UR_NODE_ADDR
UR_UNIT_NO


Passed through


Not Filtered





./FFIEC 002\URF - 9200\2023.DLQ.FRB.FR - 2052a - URF.xlsm
E
E-1

WPN0|UR_NODE_ADDR|UR_UNIT_NO
2023-11-29 14:06:21
NCCF
Internal Indicator
Counterparty Type
riskInternalFlag attribute indicates whether the position is to be considered as internal or external for reporting
Client
N/A
N/A
Tier 1
8L00 -> 8l00
XTA0
Operational Trade Information Service
(OTIS)
8L00
RIMMS - Toronto
RBC Owned
System
N/A
Sourced
N/A
N/A
N/A
N/A
strplvadaf0002.fg.rbc.com
/XTA0/otis2/PROD/incoming/RIMMS
N/A
RIMMS.OTIS.EOD_Positions.TOR.yyyymmdd.csv
InstrumentProductType
Data File Field
business product type (hardcoded value "LD")
Daily
8L00
RIMMS - Toronto >> 8L00 (IIPM#368)
RBC Owned
8L00_Nas
\SCX9.$RIMM1.RIMMDATA
N/A
CONTFILE
ID.TYP
DataFile Field
Hardcode 'LD'
Passed Through
N/A
N/A
Filtered
the contract file (CONTFILE) that match the CNLD-REC layout,
specifically where the field CNLD-REC.ID.TYP
has a value of "LD." "FX" "AN"
Predefined Variables fixed in the string, it is fed into LD , that’s hardcoded,only used when compile the program
ID.TYPE---IT CAN BE either LD,FX,AN (TYPE OF RECORD) IT IS FIXED
The data includes end-of-day position information for New Toronto.
File used:
contract file (CONTFILE, record layout is CNLD-REC)
RIMMS.OTIS.EOD_Positions.TOR.YYYYMMDD.csv
includes only CNLD-REC record
 where CNLD-REC.ID.TYP = "LD"
2a2ebef2-8d2f-4c68-96d8-c0912c1774c6
072cddd5-b74c-4588-bbcd-2bb2a9f101fe
Corporate Treasury: SIRR
 Liquidity and Securitization
./FFIEC 002\OTIS - XTA0\1821_2024.DLQ.LCR.NCCF.Wholesale.RIMMS Toronto 8L00 to OTIS Hop1,2,3,4.xlsx
N/A
3
8L00|RIMMS.OTIS.EOD_POSITIONS.TOR.YYYYMMDD.CSV|INSTRUMENTPRODUCTTYPE
8L00|CONTFILE|ID.TYP
2026-05-01 11:09:57
NCCF
Internal Indicator
Counterparty Type
riskInternalFlag attribute indicates whether the position is to be considered as internal or external for reporting
Client
N/A
N/A
Tier 1
8L00 -> 8l00
XTA0
Operational Trade Information Service
(OTIS)
8L00
RIMMS - Toronto
RBC Owned
System
N/A
Sourced
N/A
N/A
N/A
N/A
strplvadaf0002.fg.rbc.com
/XTA0/otis2/PROD/incoming/RIMMS
N/A
RIMMS.OTIS.EOD_UnsettledPositions.TOR.yyyymmdd.csv
InstrumentProductType
Data File Field
business product type (hardcoded value "LD")
Daily
8L00
RIMMS - Toronto >> 8L00 (IIPM#368)
RBC Owned
8L00_Nas
\SCX9.$RIMM1.RIMMDATA
N/A
CONTFILE
ID.TYP
DataFile Field
Hardcode 'LD'
Passed Through
N/A
N/A
Filtered
the contract file (CONTFILE) that match the CNLD-REC layout,
specifically where the field CNLD-REC.ID.TYP
has a value of "LD."
The data includes end-of-day position information for New Toronto.
File used:
contract file (CONTFILE, record layout is CNLD-REC)
RIMMS.OTIS.EOD_Positions.TOR.YYYYMMDD.csv
includes only CNLD-REC record
 where CNLD-REC.ID.TYP = "LD"
2a2ebef2-8d2f-4c68-96d8-c0912c1774c6
072cddd5-b74c-4588-bbcd-2bb2a9f101fe
Corporate Treasury: SIRR
 Liquidity and Securitization
./FFIEC 002\OTIS - XTA0\1821_2024.DLQ.LCR.NCCF.Wholesale.RIMMS Toronto 8L00 to OTIS Hop1,2,3,4.xlsx
N/A
4
8L00|RIMMS.OTIS.EOD_UNSETTLEDPOSITIONS.TOR.YYYYMMDD.CSV|INSTRUMENTPRODUCTTYPE
8L00|CONTFILE|ID.TYP
2026-05-01 11:09:57
NCCF
 LRM Product Code
Product Identifier
The LRM product code is the basis of product classification in LRM. Identifies and differentiates the various product categories. Rules are built on top of the LRM product code and other attributes to arrive at the requisite granularity required for liquidity reporting.
Security & Derivatives Master
N/A
N/A
Tier 1
8L00 -> 8l00
XTA0
Operational Trade Information Service
(OTIS)
8L00
RIMMS - Toronto
RBC Owned
System
N/A
Sourced
N/A
N/A
N/A
N/A
strplvadaf0002.fg.rbc.com
/XTA0/otis2/PROD/incoming/RIMMS
N/A
RIMMS.OTIS.EOD_Positions.TOR.yyyymmdd.csv
InstrumentProductType
Data File Field
business product type (hardcoded value "LD")
Daily
8L00
RIMMS - Toronto >> 8L00 (IIPM#368)
RBC Owned
8L00_Nas
\SCX9.$RIMM1.RIMMDATA
N/A
CONTFILE
ID.TYP
DataFile Field
Hardcode 'LD'
Passed Through
N/A
N/A
Filtered
the contract file (CONTFILE) that match the CNLD-REC layout,
specifically where the field CNLD-REC.ID.TYP
has a value of "LD." "FX" "AN"
Predefined Variables fixed in the string, it is fed into LD , that’s hardcoded,only used when compile the program
ID.TYPE---IT CAN BE either LD,FX,AN (TYPE OF RECORD) IT IS FIXED
The data includes end-of-day position information for New Toronto.
File used:
contract file (CONTFILE, record layout is CNLD-REC)
RIMMS.OTIS.EOD_Positions.TOR.YYYYMMDD.csv
includes only CNLD-REC record
 where CNLD-REC.ID.TYP = "LD"
f74798e8-fe7a-420a-91b8-aa48912857d4
072cddd5-b74c-4588-bbcd-2bb2a9f101fe
Corporate Treasury: SIRR
 Liquidity and Securitization
./FFIEC 002\OTIS - XTA0\1821_2024.DLQ.LCR.NCCF.Wholesale.RIMMS Toronto 8L00 to OTIS Hop1,2,3,4.xlsx
N/A
5
8L00|RIMMS.OTIS.EOD_POSITIONS.TOR.YYYYMMDD.CSV|INSTRUMENTPRODUCTTYPE
8L00|CONTFILE|ID.TYP
2026-05-01 11:09:57
NCCF
 LRM Product Code
Product Identifier
The LRM product code is the basis of product classification in LRM. Identifies and differentiates the various product categories. Rules are built on top of the LRM product code and other attributes to arrive at the requisite granularity required for liquidity reporting.
Security & Derivatives Master
N/A
N/A
Tier 1
8L00 -> 8l00
XTA0
Operational Trade Information Service
(OTIS)
8L00
RIMMS - Toronto
RBC Owned
System
N/A
Sourced
N/A
N/A
N/A
N/A
strplvadaf0002.fg.rbc.com
/XTA0/otis2/PROD/incoming/RIMMS
N/A
RIMMS.OTIS.EOD_UnsettledPositions.TOR.yyyymmdd.csv
InstrumentProductType
Data File Field
business product type (hardcoded value "LD")
Daily
8L00
RIMMS - Toronto >> 8L00 (IIPM#368)
RBC Owned
8L00_Nas
\SCX9.$RIMM1.RIMMDATA
N/A
CONTFILE
ID.TYP
DataFile Field
Hardcode 'LD'
Passed Through
N/A
N/A
Filtered
the contract file (CONTFILE) that match the CNLD-REC layout,
specifically where the field CNLD-REC.ID.TYP
has a value of "LD."
The data includes end-of-day position information for New Toronto.
File used:
contract file (CONTFILE, record layout is CNLD-REC)
RIMMS.OTIS.EOD_Positions.TOR.YYYYMMDD.csv
includes only CNLD-REC record
 where CNLD-REC.ID.TYP = "LD"
f74798e8-fe7a-420a-91b8-aa48912857d4
072cddd5-b74c-4588-bbcd-2bb2a9f101fe
Corporate Treasury: SIRR
 Liquidity and Securitization
./FFIEC 002\OTIS - XTA0\1821_2024.DLQ.LCR.NCCF.Wholesale.RIMMS Toronto 8L00 to OTIS Hop1,2,3,4.xlsx
N/A
6
8L00|RIMMS.OTIS.EOD_UNSETTLEDPOSITIONS.TOR.YYYYMMDD.CSV|INSTRUMENTPRODUCTTYPE
8L00|CONTFILE|ID.TYP
2026-05-01 11:09:57
NCCF
Internal Indicator
Counterparty Type
riskInternalFlag attribute indicates whether the position is to be considered as internal or external for reporting
Client
N/A
N/A
Tier 1
8L00 -> 8l00
XTA0
Operational Trade Information Service
(OTIS)
8L00
RIMMS - Toronto
RBC Owned
System
N/A
Sourced
N/A
N/A
N/A
N/A
strplvadaf0002.fg.rbc.com
/XTA0/otis2/PROD/incoming/RIMMS
N/A
RIMMS.OTIS.EOD_Positions.TOR.yyyymmdd.csv
InstrumentProductSubType
Data File Field
business product sub type
Daily
8L00
RIMMS - Toronto >> 8L00 (IIPM#368)
RBC Owned
8L00_Nas
 \SCX9.$RIMM1.RIMMDATA
N/A
CONFFILE
ID.CREC-TYPE
DataFile Field
record type
Passed Through
N/A
N/A
Filtered
For each CNLD-REC record,
check the corresponding BFNC-REC record in CONFFILE
where specific conditions match (CREC-TYPE = "BFNC", CONT-TYPE = "LD", and CONT-STYPE = CNLD-REC.SUBTYPE).
If BFNC-REC.GENERIC-TYPE is not 1 or 2,
or if CNLD-REC.SUBTYPE is one of the listed values (LFASSA LFPARA LFMANA LFOTHA CDMIAM SWAPA SWAPL TSWPF),
the CNLD-REC record is excluded from the file "RIMMS.OTIS.EOD_Positions.TOR.yyyymmdd.csv".
File used:
contract file (CONTFILE, record layout is CNLD-REC),
configuration file (CONFFILE,
record layout is BFNC-REC,
business function record)
For each CNLD-REC record,
check the BFNC-REC record in CONFFILE,
where BFNC-REC.ID.CREC-TYPE = "BFNC" and BFNC-REC.ID.CONT-TYPE = "LD" and BFNC-REC.ID.CONT-STYPE = CNLD-REC.SUBTYPE, if BFNC-REC.GENERIC-TYPE is not 1 or 2,
RIMMS.OTIS.EOD_Positions.TOR.yyyymmdd.csv excludes this CNLD-REC record.

The file RIMMS.OTIS.EOD_Positions.TOR.yyyymmdd.csv also excludes CNLD-REC records where CNLD-REC.SUBTYPE is one of below:
LFASSA LFPARA LFMANA LFOTHA CDMIAM SWAPA SWAPL TSWPF
2a2ebef2-8d2f-4c68-96d8-c0912c1774c6
072cddd5-b74c-4588-bbcd-2bb2a9f101fe
Corporate Treasury: SIRR
 Liquidity and Securitization
./FFIEC 002\OTIS - XTA0\1821_2024.DLQ.LCR.NCCF.Wholesale.RIMMS Toronto 8L00 to OTIS Hop1,2,3,4.xlsx
N/A
7
8L00|RIMMS.OTIS.EOD_POSITIONS.TOR.YYYYMMDD.CSV|INSTRUMENTPRODUCTSUBTYPE
8L00|CONFFILE|ID.CREC-TYPE
2026-05-01 11:09:57
NCCF
Internal Indicator
Counterparty Type
riskInternalFlag attribute indicates whether the position is to be considered as internal or external for reporting
Client
N/A
N/A
Tier 1
8L00 -> 8l00
XTA0
Operational Trade Information Service
(OTIS)
8L00
RIMMS - Toronto
RBC Owned
System
N/A
Sourced
N/A
N/A
N/A
N/A
strplvadaf0002.fg.rbc.com
/XTA0/otis2/PROD/incoming/RIMMS
N/A
RIMMS.OTIS.EOD_Positions.TOR.yyyymmdd.csv
InstrumentProductSubType
Data File Field
business product sub type
Daily
8L00
RIMMS - Toronto >> 8L00 (IIPM#368)
RBC Owned
8L00_Nas
 \SCX9.$RIMM1.RIMMDATA
N/A
CONFFILE
ID.CONT-TYPE
DataFile Field
product type
Passed Through
N/A
N/A
Filtered
For each CNLD-REC record,
check the corresponding BFNC-REC record in CONFFILE
where specific conditions match (CREC-TYPE = "BFNC", CONT-TYPE = "LD", and CONT-STYPE = CNLD-REC.SUBTYPE).
If BFNC-REC.GENERIC-TYPE is not 1 or 2,
or if CNLD-REC.SUBTYPE is one of the listed values (LFASSA LFPARA LFMANA LFOTHA CDMIAM SWAPA SWAPL TSWPF),
the CNLD-REC record is excluded from the file "RIMMS.OTIS.EOD_Positions.TOR.yyyymmdd.csv".
File used:
contract file (CONTFILE, record layout is CNLD-REC),
configuration file (CONFFILE,
record layout is BFNC-REC,
business function record)
For each CNLD-REC record,
check the BFNC-REC record in CONFFILE,
where BFNC-REC.ID.CREC-TYPE = "BFNC" and BFNC-REC.ID.CONT-TYPE = "LD" and BFNC-REC.ID.CONT-STYPE = CNLD-REC.SUBTYPE, if BFNC-REC.GENERIC-TYPE is not 1 or 2,
RIMMS.OTIS.EOD_Positions.TOR.yyyymmdd.csv excludes this CNLD-REC record.

The file RIMMS.OTIS.EOD_Positions.TOR.yyyymmdd.csv also excludes CNLD-REC records where CNLD-REC.SUBTYPE is one of below:
LFASSA LFPARA LFMANA LFOTHA CDMIAM SWAPA SWAPL TSWPF
2a2ebef2-8d2f-4c68-96d8-c0912c1774c6
072cddd5-b74c-4588-bbcd-2bb2a9f101fe
Corporate Treasury: SIRR
 Liquidity and Securitization
./FFIEC 002\OTIS - XTA0\1821_2024.DLQ.LCR.NCCF.Wholesale.RIMMS Toronto 8L00 to OTIS Hop1,2,3,4.xlsx
N/A
8
8L00|RIMMS.OTIS.EOD_POSITIONS.TOR.YYYYMMDD.CSV|INSTRUMENTPRODUCTSUBTYPE
8L00|CONFFILE|ID.CONT-TYPE
2026-05-01 11:09:57


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