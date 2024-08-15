import openpyxl
import pandas as pd

# Load dest. workbook
log_wb = openpyxl.load_workbook(f"C:\\Users\\Sharon YY Tseng\\Desktop\\issue_log_automation\\Issue_log_automation-\\Model Office Trustee Issue Log 2024_LIVE.xlsx")
remove_ls = ["Sheet1", "Drop-down List", "Status update summary", "Issue Log Screenshot Reference ","Status"]

for i in remove_ls:
    if i in log_wb.sheetnames:
        log_wb.remove(log_wb[i])
         
    
issue_ws = log_wb["Model Office Issue Log"]
log_wb.active = issue_ws

trustee_status_df = pd.read_excel(f"C:\\Users\\Sharon YY Tseng\\Desktop\\issue_log_automation\\Issue_log_automation-\\Model Office Trustee Issue Log 2024_LIVE.xlsx",
                                  sheet_name = "Trustee_MO_status")

# Identify selected trustee name
for i in trustee_status_df["Chosen_trustee:"].dropna():
    selected_trustee = i

# get max_row number 
total_rows = issue_ws.max_row

# Filter rows
for row_idx in range(total_rows, 4, -1):  # Start from the bottom and go up
    cell_value = issue_ws.cell(row=row_idx, column=2).value
    if cell_value:
        issue_id_letters = ''.join(filter(str.isalpha, str(cell_value)))
        if issue_id_letters != selected_trustee:  # Delete if they don't match
            issue_ws.delete_rows(row_idx, 1)

#remove extra columns 
issue_ws = issue_ws.delete_cols(idx=18, amount=16)
          
log_wb.save("mod_issue_log.xlsx")       
log_wb.close()
