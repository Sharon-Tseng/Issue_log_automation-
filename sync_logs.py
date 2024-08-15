import pandas as pd 
import openpyxl
import datetime


trustee_xlsx_path = "C:\\Users\\Sharon YY Tseng\\Desktop\\issue_log_automation\\Issue_log_automation-\\Model Office Issue Log_Cycle 2 Batch 2_BEAT(VS)_20240812.xlsx"
detailed_tl_path = "C:\\Users\\Sharon YY Tseng\\Desktop\\issue_log_automation\\BEATVS_Model_Office_Trustee_Issue_Log.xlsx"

trustee_latest_df = pd.read_excel(trustee_xlsx_path, sheet_name = "Model Office Issue Log", index_col=False, header = None)
detailed_tl_df = pd.read_excel(detailed_tl_path)


# Data Table Cleaning
trustee_latest_df = trustee_latest_df[3:]
trustee_latest_df.columns = trustee_latest_df.iloc[0]     # reset first row as colum name 
trustee_latest_df = trustee_latest_df[1:]
trustee_latest_df.reset_index(drop = True, inplace = True)

# trustee_latest_df drop columns before "Issue ID (TR_No.)"
columns_to_drop = []
drop = False
for col in trustee_latest_df.columns:
    if col == "Issue ID (TR_No.)":
        drop = True
    if not drop:
        columns_to_drop.append(col)
trustee_latest_df.drop(columns=columns_to_drop, inplace=True)


# trustee_latest_df drop columns after "Simulation Re-run Result"
col_index = trustee_latest_df.columns.get_loc("Simulation Re-run Result")
trustee_latest_df = trustee_latest_df.iloc[:, :col_index+1]

# function to save missing columns 
def save_missing_col(to_be_changed_df, accordance_df, col_name, pending_df = None):
    if pending_df is None:
        pending_df = to_be_changed_df[["Issue ID (TR_No.)",col_name]]
        to_be_changed_df = to_be_changed_df.drop(columns = col_name)
    else:
        pending_df[col_name] = to_be_changed_df[col_name]
        to_be_changed_df = to_be_changed_df.drop(columns=col_name)
    return to_be_changed_df, pending_df

detailed_tl_df, pending_df = save_missing_col(detailed_tl_df, trustee_latest_df, "Sub-issue ID (ORG_No.)")
detailed_tl_df, pending_df = save_missing_col(detailed_tl_df,trustee_latest_df,"Cycle ", pending_df= pending_df)


# Function to get updates_df
def get_updates(to_be_changed_df, accordance_df):
    to_be_changed_df = to_be_changed_df.set_index("Issue ID (TR_No.)")
    accordance_df = accordance_df.set_index("Issue ID (TR_No.)")
    
    updates_df = pd.DataFrame(columns=["Issue ID","Column Name","Old Value","New Value"])
    for issue_id in accordance_df.index:   
        if issue_id in to_be_changed_df.index:   
            for col in to_be_changed_df.columns:
                if to_be_changed_df.at[issue_id, col] != accordance_df.at[issue_id,col]:
                    new_row = {"Issue ID":issue_id, "Column Name":col,
                               "Old Value": to_be_changed_df.at[issue_id, col],
                               "New Value": accordance_df.at[issue_id, col]}
                    updates_df.loc[len(updates_df)] = new_row
            to_be_changed_df.loc[issue_id] = accordance_df.loc[issue_id]

        else:
            new_row = accordance_df.loc[issue_id]
            for col in new_row.index:
                to_be_changed_df.at[issue_id, col] = new_row[col]
                to_be_changed_df.loc[issue_id] = accordance_df.loc[issue_id]
                
    return to_be_changed_df, accordance_df, updates_df    
     
result = get_updates(detailed_tl_df, trustee_latest_df)

result[0].to_excel("test_detailed_tl.xlsx")
result[1].to_excel("test_trustee.xlsx")
                
updates_df = result[2].dropna(subset = ["Old Value", "New Value"], how="all")

# updates_df= updates_df[updates_df['Column Name'].isin(["Issue Description Note: Please be reminded to exclude personal information in description."
#                                                         ,"Status (Drop down list)"])]

# # Export to Excel
updates_df.to_excel("updates.xlsx", index=False)
