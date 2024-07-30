import pandas as pd 
import openpyxl

# using BCOMM as trustee_wb - testing 
trustee_df = pd.read_excel(f"C:\\Users\\Sharon YY Tseng\\Desktop\\issue_log_automation\\Issue_log_automation-\\Model Office Trustee Issue Log 2024_LIVE.xlsx",sheet_name="Model Office Issue Log")
detailed_tl_df = pd.read_excel(f"C:\\Users\\Sharon YY Tseng\\Desktop\\issue_log_automation\\BCOMM_Model_Office_Trustee_Issue_Log_formatted.xlsx")


updates_dict = {"Row_updated":[],
                "Column_updated":[],
                "Updated_value_from":[],
                "Updated_value_to":[]
}

# Synchronize data
new_trustee_df = trustee_df.iloc[:,5:19][4:]
new_trustee_df.columns  = ["Issue ID", "Raised by", "Severity", "Issue Type", 
                           "Process", "Simulation Scenario No", "Issue Description",
                           "Status","Potential Impact","Creation Date", "Follow up by",
                           "Simulation Re-run Date", "Proposed Workaround/Clarification Response",
                           "Simulation Re-run Result"]

detailed_tl_df.columns = ["Issue ID", "Raised by", "Sub-Issue ID","Severity", "Issue Type", 
                           "Process", "Simulation Scenario No", "Issue Description",
                           "Status","Potential Impact","Creation Date", "Follow up by",
                           "Simulation Re-run Date", "Proposed Workaround/Clarification Response",
                           "Simulation Re-run Result"]

# Save "Sub-Issue ID" for future merge
sub_issue_id = detailed_tl_df["Sub-Issue ID"]

detailed_tl_df = detailed_tl_df.drop("Sub-Issue ID", axis = "columns")
print(detailed_tl_df)


#updates_df = pd.DataFrame(updates_dict)
#updates_df.to_excel(f"test_update_outcome.xlsx", index = False)



