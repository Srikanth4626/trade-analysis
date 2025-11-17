#!/usr/bin/env python3
"""
generate_workbook.py
Generates Trade_Analysis_Full_Workbook.xlsx from sample data (CSV or XLSX).
Produces Raw Data, Cleaned Data (formula-driven), Lookup Tables, Summaries, Charts.
"""
import sys
import os
import pandas as pd
import numpy as np

def load_input(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv",):
        df = pd.read_csv(path)
    elif ext in (".xlsx", ".xls"):
        df = pd.read_excel(path, sheet_name=0)
    else:
        raise ValueError("Unsupported input file type: " + ext)
    return df

def build_workbook(df, out_path):
    writer = pd.ExcelWriter(out_path, engine='xlsxwriter', datetime_format='yyyy-mm-dd')
    workbook = writer.book

    # Raw Data sheet
    raw_sheet = 'Raw Data'
    df.to_excel(writer, sheet_name=raw_sheet, index=False)
    ws_raw = writer.sheets[raw_sheet]

    # Lookup Tables
    lookup = [
        ('HS Code','HSN Description','Main Category'),
        ('73239990','Household articles of iron or steel','Steel'),
        ('73239900','Table, kitchen household articles','Steel'),
        ('73211900','Cooking appliances and plate warmers','Steel'),
        ('73239300','Kitchen or tableware','Steel'),
    ]
    ws_lookup = workbook.add_worksheet('Lookup Tables')
    writer.sheets['Lookup Tables'] = ws_lookup
    for r,row in enumerate(lookup):
        for c,val in enumerate(row):
            ws_lookup.write(r, c, val)

    # Cleaned Data headers
    clean_headers = ["Date","Port Code","IEC","HS Code","HSN Description","Goods Description",
                     "Main Category","Sub Category","Model Name","Model Number","Capacity",
                     "Quantity","Unit","Unit Price INR","Total Value INR","Unit Price USD",
                     "Total Value USD","Duty Paid INR","Grand Total INR","Year"]
    ws_clean = workbook.add_worksheet('Cleaned Data')
    writer.sheets['Cleaned Data'] = ws_clean
    for c,h in enumerate(clean_headers):
        ws_clean.write(0,c,h)

    # Map raw columns to excel columns (raw sheet starts at row 1 header)
    raw_cols = df.columns.tolist()
    raw_map = {name: idx for idx,name in enumerate(raw_cols)}  # 0-based

    # helper to get column letter
    def col_letter(idx):
        letters = ""
        while idx >= 0:
            letters = chr(ord('A') + (idx % 26)) + letters
            idx = idx // 26 - 1
        return letters

    nrows = len(df)
    for i in range(nrows):
        excel_row = i + 2  # Excel row index (1-based header)
        def raw_cell(col_name, r):
            col_idx = raw_map[col_name]
            return f"'{raw_sheet}'!{col_letter(col_idx)}{r}"

        # Many formulas rely on the raw column headers present in Sample Data.
        # Adjust these references if your input has different headers.
        # Date
        if 'DATE' in raw_map:
            ws_clean.write_formula(i+1, 0, f"={raw_cell('DATE', excel_row)}")
        else:
            ws_clean.write_blank(i+1, 0, None)
        # Port Code
        if 'PORT CODE' in raw_map:
            ws_clean.write_formula(i+1, 1, f"={raw_cell('PORT CODE', excel_row)}")
        # IEC
        if 'IEC' in raw_map:
            ws_clean.write_formula(i+1, 2, f"={raw_cell('IEC', excel_row)}")
        # HS CODE
        if 'HS CODE' in raw_map:
            ws_clean.write_formula(i+1, 3, f"={raw_cell('HS CODE', excel_row)}")
        # HSN Description lookup
        ws_clean.write_formula(i+1, 4, f"=IFERROR(VLOOKUP(D{excel_row},'Lookup Tables'!$A:$B,2,FALSE),\"Unknown\")")
        # Goods Description
        if 'GOODS DESCRIPTION' in raw_map:
            ws_clean.write_formula(i+1, 5, f"={raw_cell('GOODS DESCRIPTION', excel_row)}")
        # Main Category
        ws_clean.write_formula(i+1, 6, f"=IFERROR(VLOOKUP(D{excel_row},'Lookup Tables'!$A:$C,3,FALSE),IF(ISNUMBER(SEARCH(\"STEEL\",F{excel_row})),\"Steel\",\"Others\"))")
        # Sub Category (keywords)
        ws_clean.write_formula(i+1, 7, f"=IF(ISNUMBER(SEARCH(\"scrubber\",F{excel_row})) , \"Scrubber\", IF(ISNUMBER(SEARCH(\"container\",F{excel_row})),\"Container\", IF(ISNUMBER(SEARCH(\"basket\",F{excel_row})),\"Basket\", IF(ISNUMBER(SEARCH(\"lunch\",F{excel_row})),\"Lunch Box\", IF(ISNUMBER(SEARCH(\"cutlery\",F{excel_row})),\"Cutlery\",\"Other\")))))")
        # Model Name prefer raw 'Model Name'
        if 'Model Name' in raw_map:
            ws_clean.write_formula(i+1, 8, f"=IF({raw_cell('Model Name', excel_row)}<>\"\",{raw_cell('Model Name', excel_row)},IFERROR(MID(F{excel_row},SEARCH(\"MODEL\",F{excel_row})+6,20),\"\"))")
        else:
            ws_clean.write_formula(i+1, 8, f"=IFERROR(MID(F{excel_row},SEARCH(\"MODEL\",F{excel_row})+6,20),\"\")")
        # Model Number
        if 'Model Number' in raw_map:
            ws_clean.write_formula(i+1, 9, f"={raw_cell('Model Number', excel_row)}")
        # Capacity
        if 'Capacity' in raw_map:
            ws_clean.write_formula(i+1, 10, f"={raw_cell('Capacity', excel_row)}")
        # Quantity prefer 'QUANTITY' else parse QTY from description
        if 'QUANTITY' in raw_map:
            ws_clean.write_formula(i+1, 11, f"=IF({raw_cell('QUANTITY', excel_row)}<>0,{raw_cell('QUANTITY', excel_row)},IFERROR(VALUE(MID(F{excel_row},SEARCH(\"QTY\",F{excel_row})+4,IFERROR(FIND(\" \",F{excel_row},SEARCH(\"QTY\",F{excel_row})+4)-(SEARCH(\"QTY\",F{excel_row})+4),3))),\"\") )")
        else:
            ws_clean.write_formula(i+1, 11, f"=IFERROR(VALUE(MID(F{excel_row},SEARCH(\"QTY\",F{excel_row})+4,IFERROR(FIND(\" \",F{excel_row},SEARCH(\"QTY\",F{excel_row})+4)-(SEARCH(\"QTY\",F{excel_row})+4),3))),\"\")")
        # Unit
        if 'UNIT' in raw_map:
            ws_clean.write_formula(i+1, 12, f"={raw_cell('UNIT', excel_row)}")
        elif 'Unit of measure' in raw_map:
            ws_clean.write_formula(i+1, 12, f"={raw_cell('Unit of measure', excel_row)}")
        # Unit Price INR
        if 'UNIT PRICE_INR' in raw_map:
            ws_clean.write_formula(i+1, 13, f"={raw_cell('UNIT PRICE_INR', excel_row)}")
        # Total Value INR
        if 'TOTAL VALUE_INR' in raw_map:
            ws_clean.write_formula(i+1, 14, f"={raw_cell('TOTAL VALUE_INR', excel_row)}")
        # Unit Price USD
        if 'UNIT PRICE_USD' in raw_map:
            ws_clean.write_formula(i+1, 15, f"={raw_cell('UNIT PRICE_USD', excel_row)}")
        # Total Value USD
        if 'TOTAL VALUE_USD' in raw_map:
            ws_clean.write_formula(i+1, 16, f"={raw_cell('TOTAL VALUE_USD', excel_row)}")
        # Duty Paid INR
        if 'DUTY PAID_INR' in raw_map:
            ws_clean.write_formula(i+1, 17, f"={raw_cell('DUTY PAID_INR', excel_row)}")

        # Grand Total = Total Value INR + Duty Paid INR
        ws_clean.write_formula(i+1, 18, f"=IFERROR(O{excel_row},0)+IFERROR(R{excel_row},0)")
        # Year
        ws_clean.write_formula(i+1, 19, f"=IF(A{excel_row}=\"\",\"\",YEAR(A{excel_row}))")

    # Auto column widths
    for i, col in enumerate(df.columns.tolist()):
        ws_raw.set_column(i, i, max(12, min(40, len(str(col))+2)))
    for i, col in enumerate(clean_headers):
        ws_clean.set_column(i, i, max(12, min(40, len(str(col))+2)))

    # Year Summary sheet (SUMIFS)
    years = sorted(pd.to_datetime(df['DATE'], errors='coerce').dropna().dt.year.unique().tolist())
    ws_year = workbook.add_worksheet('Year Summary')
    writer.sheets['Year Summary'] = ws_year
    ws_year.write_row(0,0, ['Year','Total Value INR','Duty Paid INR','Grand Total INR','YoY Growth %'])
    for r, yr in enumerate(years, start=1):
        ws_year.write(r,0, yr)
        ws_year.write_formula(r,1, f"=SUMIFS('Cleaned Data'!O:O,'Cleaned Data'!T:T,{yr})")
        ws_year.write_formula(r,2, f"=SUMIFS('Cleaned Data'!R:R,'Cleaned Data'!T:T,{yr})")
        ws_year.write_formula(r,3, f"=SUMIFS('Cleaned Data'!S:S,'Cleaned Data'!T:T,{yr})")
    # YoY formula
    for r in range(2, len(years)+1):
        ws_year.write_formula(r,4, f"=IFERROR((D{r+1}-D{r})/D{r},\"\")")

    # HSN summary sheet
    hs_codes = sorted(df['HS CODE'].dropna().astype(str).unique().tolist())
    ws_hsn = workbook.add_worksheet('HSN Summary')
    writer.sheets['HSN Summary'] = ws_hsn
    ws_hsn.write_row(0,0, ['HS Code','HSN Description','Total Value INR','Duty Paid INR','Grand Total INR','% Contribution'])
    for r, hs in enumerate(hs_codes, start=1):
        ws_hsn.write(r,0, hs)
        ws_hsn.write_formula(r,1, f"=IFERROR(VLOOKUP(A{r+1},'Lookup Tables'!$A:$B,2,FALSE),\"Unknown\")")
        ws_hsn.write_formula(r,2, f"=SUMIFS('Cleaned Data'!O:O,'Cleaned Data'!D:D,A{r+1})")
        ws_hsn.write_formula(r,3, f"=SUMIFS('Cleaned Data'!R:R,'Cleaned Data'!D:D,A{r+1})")
        ws_hsn.write_formula(r,4, f"=SUMIFS('Cleaned Data'!S:S,'Cleaned Data'!D:D,A{r+1})")
    total_row = len(hs_codes)+2
    ws_hsn.write_formula(total_row,4, f"=SUM(E2:E{len(hs_codes)+1})")
    for r in range(1, len(hs_codes)+1):
        ws_hsn.write_formula(r,5, f"=IF($E${total_row+1}=0,0,E{r+1}/$E${total_row+1})")

    # Model summary
    models = df.get('Model Name', pd.Series([''] * len(df))).fillna('').astype(str)
    unique_models = sorted([m for m in models.unique() if str(m).strip()!=''])
    ws_model = workbook.add_worksheet('Model Summary')
    writer.sheets['Model Summary'] = ws_model
    ws_model.write_row(0,0, ['Model Name','Total Qty','Total Value INR','Avg Unit Price USD','Avg Unit Price INR','% of Total Value'])
    for r,m in enumerate(unique_models, start=1):
        ws_model.write(r,0,m)
        ws_model.write_formula(r,1, f"=SUMIFS('Cleaned Data'!L:L,'Cleaned Data'!I:I,\"{m}\")")
        ws_model.write_formula(r,2, f"=SUMIFS('Cleaned Data'!O:O,'Cleaned Data'!I:I,\"{m}\")")
        ws_model.write_formula(r,3, f"=IFERROR(AVERAGEIFS('Cleaned Data'!P:P,'Cleaned Data'!I:I,\"{m}\"),\"\")")
        ws_model.write_formula(r,4, f"=IFERROR(AVERAGEIFS('Cleaned Data'!N:N,'Cleaned Data'!I:I,\"{m}\"),\"\")")
    ws_model.write_formula(len(unique_models)+2,2, f"=SUM(C2:C{len(unique_models)+1})")
    for r in range(1, len(unique_models)+1):
        ws_model.write_formula(r,5, f"=IF($C${len(unique_models)+3}=0,0, C{r+1}/$C${len(unique_models)+3})")

    # Supplier summary
    suppliers = df.get('IEC', pd.Series([''] * len(df))).fillna('').astype(str)
    unique_suppliers = sorted([s for s in suppliers.unique() if str(s).strip()!=''])
    ws_sup = workbook.add_worksheet('Supplier Summary')
    writer.sheets['Supplier Summary'] = ws_sup
    ws_sup.write_row(0,0, ['IEC','Total Value INR','Total Qty','% Contribution'])
    for r,s in enumerate(unique_suppliers, start=1):
        ws_sup.write(r,0,s)
        ws_sup.write_formula(r,1, f"=SUMIFS('Cleaned Data'!O:O,'Cleaned Data'!C:C,\"{s}\")")
        ws_sup.write_formula(r,2, f"=SUMIFS('Cleaned Data'!L:L,'Cleaned Data'!C:C,\"{s}\")")
    ws_sup.write_formula(len(unique_suppliers)+2,1, f"=SUM(B2:B{len(unique_suppliers)+1})")
    for r in range(1, len(unique_suppliers)+1):
        ws_sup.write_formula(r,3, f"=IF($B${len(unique_suppliers)+3}=0,0, B{r+1}/$B${len(unique_suppliers)+3})")

    # Notes
    ws_notes = workbook.add_worksheet('Notes')
    writer.sheets['Notes'] = ws_notes
    notes = [
        "Notes:",
        "- Cleaned Data contains formulas referencing Raw Data.",
        "- Use Insert > PivotTable in Excel (tbl or range) if you prefer pivot objects.",
        "- All calculations use Excel functions (SUMIFS, AVERAGEIFS, VLOOKUP, IF, MID, SEARCH).",
    ]
    for r,line in enumerate(notes):
        ws_notes.write(r,0,line)

    writer.close()
    print("Workbook written to", out_path)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python scripts/generate_workbook.py <input_csv_or_xlsx> <output_xlsx>")
        print("Example: python scripts/generate_workbook.py sample_data/'Sample Data 2.xlsx' outputs/Trade_Analysis_Full_Workbook.xlsx")
        sys.exit(1)
    input_path = sys.argv[1]
    out_path = sys.argv[2]
    df = load_input(input_path)
    build_workbook(df, out_path)
