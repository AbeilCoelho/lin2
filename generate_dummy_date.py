import pandas as pd
import numpy as np

SCHEMA_COLUMNS = ["Report name", "CDE name", "Business Term in Collibra", "CDE Definition in CDE Memo", "Business Term Domain", "Product", "Line of Business", "Tier", "Destination to (App Code)", "Destination to (App Name)", "Lineage Flow (App Codes)", "Current app code", "Current app name", "Current App Ownership (RBC Owned or Vendor Managed)", "Current App Type (System / EUC / System)", "Cluster 7 ID#", "Created/Sourced", "Manually entered/Derived", "Screen Name (UI)", "Screen Field Name (UI)", "Sub-DE", "Database System/File System (Host Name/Instance Name)", "Database/Directory", "Schema", "Current table/file name", "Current column/field name", "Data Element Field Type", "Field Description", "Frequency", "Source app code", "Source app name", "App Ownership (RBC Owned or Vendor Managed)", "Source Database System/File System (Host Name/Instance Name)", "Field Database/Directory", "Source Schema", "Source table/file name", "Source column/field name", "Source Field Type", "Source Field Definition", "Transformed/Passed Through", "Transformation Description", "Transformation Logic", "Filtered", "Filtration Description", "Filtration Logic", "BUSINESS_TERM Collibra ID", "BUSINESS_PROCESS Collibra IDs", "BUSINESS_PROCESS Domains", "file", "column", "group_id", "current_key", "source_key", "modification_date"]

def create_row(**kwargs):
    row = {col: "" for col in SCHEMA_COLUMNS}
    row.update(kwargs)
    return row

def generate_data():
    cde_targets = [
        {"Report name": "FRY9C", "CDE name": "Auto-Branching Loan", "Target App Codes": "DW00, GL00", "Current app code": "RPT1", "Current app name": "Regulatory Reporting Engine", "Current table/file name": "Consol_Loan", "Current column/field name": "com_loan_bal"},
        {"Report name": "FRY9C", "CDE name": "Fuzzy Exposure", "Target App Codes": "DW00", "Current app code": "RPT1", "Current app name": "Regulatory Reporting Engine", "Current table/file name": "Basel_Report", "Current column/field name": "exp_amt"},
        {"Report name": "FRY9C", "CDE name": "Dead End Metric", "Target App Codes": "DW00", "Current app code": "RPT1", "Current app name": "Regulatory Reporting Engine", "Current table/file name": "Metric_Report", "Current column/field name": "val"}
    ]
    
    with pd.ExcelWriter("reporting_layers.xlsx", engine="openpyxl") as writer:
        pd.DataFrame(cde_targets).to_excel(writer, sheet_name="Sheet1", index=False)
        pd.DataFrame(columns=["Report name", "CDE name", "Current app code", "Current table/file name", "Current column/field name"]).to_excel(writer, sheet_name="Done", index=False)

    primary_rows = [
        create_row(**{"Report name": "FRY9C", "CDE name": "Auto-Branching Loan", "Current app code": "RPT1", "Current app name": "Regulatory Reporting Engine", "Current table/file name": "Consol_Loan", "Current column/field name": "com_loan_bal", 
                      "Source app code": "CCI0", "Source app name": "Commercial Credit Hub", "Source table/file name": "Risk_Hub", "Source column/field name": "r_bal"}),
        create_row(**{"Report name": "FRY9C", "CDE name": "Auto-Branching Loan", "Current app code": "RPT1", "Current app name": "Regulatory Reporting Engine", "Current table/file name": "Consol_Loan", "Current column/field name": "com_loan_bal", 
                      "Source app code": "FIN1", "Source app name": "Finance Ledger System", "Source table/file name": "Fin_Hub", "Source column/field name": "f_bal"}),
        create_row(**{"Report name": "FRY9C", "CDE name": "Fuzzy Exposure", "Current app code": "RPT1", "Current app name": "Regulatory Reporting Engine", "Current table/file name": "Basel_Report", "Current column/field name": "exp_amt", 
                      "Source app code": "WOD0", "Source app name": "Wholesale Data Store", "Source table/file name": "Rsk_Sumary_Tbl", "Source column/field name": "exp_amt"}),
        create_row(**{"Report name": "FRY9C", "CDE name": "Dead End Metric", "Current app code": "RPT1", "Current app name": "Regulatory Reporting Engine", "Current table/file name": "Metric_Report", "Current column/field name": "val", 
                      "Source app code": "EXT1", "Source app name": "External Vendor App", "Source table/file name": "Vendor_File", "Source column/field name": "val"})
    ]
    pd.DataFrame(primary_rows).to_excel("primary_lineage.xlsx", index=False)

    global_rows = [
        create_row(**{"Report name": "Gen", "CDE name": "Auto-Branching Loan", "Current app code": "CCI0", "Current app name": "Commercial Credit Hub", "Current table/file name": "Risk_Hub", "Current column/field name": "r_bal", 
                      "Source app code": "DW00", "Source app name": "Enterprise Data Warehouse", "Source table/file name": "EDW", "Source column/field name": "DW_BAL"}),
        
        # FIN1: Has ALL THREE indicators to test the UI (Manual, Filtered, Transformed)
        create_row(**{"Report name": "Gen", "CDE name": "Auto-Branching Loan", "Current app code": "FIN1", "Current app name": "Finance Ledger System", "Current table/file name": "Fin_Hub", "Current column/field name": "f_bal", 
                      "Source app code": "GL00", "Source app name": "General Ledger", "Source table/file name": "GL", "Source column/field name": "GL_BAL",
                      "Created/Sourced": "Manually entered", "Transformed/Passed Through": "Transformed", "Transformation Logic": "SUM(GL_BAL)", "Filtered": "Filtered", "Filtration Logic": "WHERE GL_BAL > 0"}),

        # WOD0: Has ONLY a Filter applied
        create_row(**{"Report name": "Gen", "CDE name": "Fuzzy Exposure", "Current app code": "WOD0", "Current app name": "Wholesale Data Store", "Current table/file name": "Risk_Summary_Tbl", "Current column/field name": "exp_amt",
                      "Source app code": "DW00", "Source app name": "Enterprise Data Warehouse", "Source table/file name": "EDW", "Source column/field name": "DW_EXP",
                      "Filtered": "Filtered", "Filtration Logic": "WHERE STATUS = 'ACTIVE'"}),

        # EXT1: Dead end with ONLY Manually entered data
        create_row(**{"Report name": "Gen", "CDE name": "Dead End Metric", "Current app code": "EXT1", "Current app name": "External Vendor App", "Current table/file name": "Vendor_File", "Current column/field name": "val", 
                      "Source app code": np.nan, "Source table/file name": np.nan, "Source column/field name": np.nan,
                      "Manually entered/Derived": "Manually entered"})
    ]
    pd.DataFrame(global_rows).to_excel("global_lineage.xlsx", index=False)

if __name__ == "__main__":
    generate_data()