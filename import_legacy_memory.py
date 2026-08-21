import json
import os
import sys

LEGACY_FILE = "lineage_rank_history.json"
WORKSPACE_DIR = "workspace"

def main():
    print("=" * 60)
    print("🧠 Legacy File Memory Import Tool")
    print("=" * 60)

    # 1. Check if legacy file exists
    if not os.path.exists(LEGACY_FILE):
        print(f"[X] ERROR: Could not find '{LEGACY_FILE}' in the current directory.")
        print("    Please place the file in the same folder as this script and try again.")
        sys.exit(1)

    # 2. Load legacy data
    print(f"[*] Loading legacy history from {LEGACY_FILE}...")
    try:
        with open(LEGACY_FILE, "r", encoding="utf-8") as f:
            legacy_data = json.load(f)
    except Exception as e:
        print(f"[X] ERROR parsing JSON: {e}")
        sys.exit(1)

    print(f"    -> Found {len(legacy_data)} file records.")

    # 3. Get target project
    project_name = input("\n[?] Enter the target Project Name (exactly as it appears in your workspace): ").strip()
    
    project_dir = os.path.join(WORKSPACE_DIR, project_name)
    if not os.path.exists(project_dir):
        print(f"[X] ERROR: Project folder '{project_dir}' does not exist.")
        sys.exit(1)

    state_file = os.path.join(project_dir, "uploads", "auto_resolve_state.json")

    # 4. Load current state
    state = {}
    if os.path.exists(state_file):
        try:
            with open(state_file, "r", encoding="utf-8") as f:
                state = json.load(f)
            print(f"[*] Found existing state for project '{project_name}'.")
        except json.JSONDecodeError:
            print("[!] Warning: auto_resolve_state.json is empty or corrupt. Starting fresh.")

    # 5. Merge Data
    if "file_memory" not in state:
        state["file_memory"] = {}

    merged_count = 0
    total_boost = 0

    for file_path, count in legacy_data.items():
        if isinstance(count, int) or str(count).isdigit():
            val = int(count)
            current_count = state["file_memory"].get(file_path, 0)
            state["file_memory"][file_path] = current_count + val
            
            merged_count += 1
            total_boost += val

    # 6. Save updated state
    print("[*] Saving merged memory to project state...")
    os.makedirs(os.path.dirname(state_file), exist_ok=True)
    with open(state_file, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=4)

    print("=" * 60)
    print(f"✅ SUCCESS! Imported {merged_count} unique files.")
    print(f"📈 Total Historical Uses Added: {total_boost}")
    print("=" * 60)

if __name__ == "__main__":
    main()