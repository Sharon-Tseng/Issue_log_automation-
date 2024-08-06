import pandas as pd 
import openpyxl
import datetime


# using BCOMM as testing 
# trustee_latest_df: to be updated, should be the latest indi trustee updates.xlsx
trustee_latest_df = pd.read_excel(f"C:\\Users\\Sharon YY Tseng\\Desktop\\issue_log_automation\\Issue_log_automation-\\Model Office BComm Issue Log_20240729 (3).xlsx", 
                                  index_col=False, header = None)

trustee_latest_df.drop (columns = trustee_latest_df.columns[:5], axis = 1, inplace = True)
trustee_latest_df.drop(columns = trustee_latest_df.columns[14:], axis=1, inplace = True)
trustee_latest_df = trustee_latest_df[3:]

# set new header to trustee_latest_df
new_header = trustee_latest_df.iloc[0] 
trustee_latest_df = trustee_latest_df[1:] 
trustee_latest_df.columns = new_header
trustee_latest_df = trustee_latest_df.reset_index().drop(columns = "index")
trustee_latest_df.columns = ['Issue ID (TR_No.)', 'Raised by (Drop down list)','Severity (Drop down list)', 'Issue Type (Drop down list)',
                          'Process (Drop down list)', 'Simulation Scenario No. ', 'Issue Description Note: Please be reminded to exclude personal information in description.',
                            'Status (Drop down list)', 'Potential Impact (Optional)','Creation Date (DD-MM-YYYY)', 'Follow up by (i..e Owner)',
                            'Simulation Re-run Date (DD-MM-YYYY)','Proposed Workaround (if applicable) / Clarification Response','Simulation Re-run Result']

detailed_tl_df = pd.read_excel(f"C:\\Users\\Sharon YY Tseng\\Desktop\\issue_log_automation\\BCOMM_Model_Office_Trustee_Issue_Log_formatted.xlsx",
                               header=None)
# Name columns 
detailed_tl_df.columns = ['Issue ID (TR_No.)', 'Raised by (Drop down list)', 'Sub-Issue ID','Severity (Drop down list)', 'Issue Type (Drop down list)',
                          'Process (Drop down list)', 'Simulation Scenario No. ','Issue Description Note: Please be reminded to exclude personal information in description.',
                          'Status (Drop down list)', 'Potential Impact (Optional)','Creation Date (DD-MM-YYYY)', 'Follow up by (i..e Owner)','Simulation Re-run Date (DD-MM-YYYY)',
                          'Proposed Workaround (if applicable) / Clarification Response','Simulation Re-run Result']


# drop "Sub-Issue ID" 
detailed_tl_df = detailed_tl_df.drop("Sub-Issue ID", axis = "columns")


def get_updates(to_be_changed_df, accordance_df):
    to_be_changed_df = to_be_changed_df.set_index("Issue ID (TR_No.)")
    accordance_df = accordance_df.set_index("Issue ID (TR_No.)")
    
    
    updates_df = pd.DataFrame(columns=["Issue ID","Column Name","Old Value","New Value"])
    for issue_id in accordance_df.index:   #BCOMM_001~BCOMM200
        if issue_id in to_be_changed_df.index:   #BCOMM_001~BCOMM143
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

updates_df= updates_df[updates_df['Column Name'].isin(["Issue Description Note: Please be reminded to exclude personal information in description."
                                                        ,"Status (Drop down list)"])]

print(updates_df)
updates_df.to_excel("updates.xlsx", index=False)

# # reset index as Issue ID (TR_No.)
# detailed_tl_df = detailed_tl_df.set_index("Issue ID (TR_No.)")
# trustee_latest_df = trustee_latest_df.set_index("Issue ID (TR_No.)")


# updates_df = pd.DataFrame(columns=["Issue ID","Column Name","Old Value","New Value"])
# for issue_id in trustee_latest_df.index:   #BCOMM_001~BCOMM200
#     if issue_id in detailed_tl_df.index:   #BCOMM_001~BCOMM143
#         for col in detailed_tl_df.columns:
#             if detailed_tl_df.at[issue_id, col] != trustee_latest_df.at[issue_id,col]:
#                 new_row = {"Issue ID":issue_id, "Column Name":col,
#                            "Old Value": detailed_tl_df.at[issue_id, col],
#                            "New Value": trustee_latest_df.at[issue_id, col]}
#                 updates_df.loc[len(updates_df)] = new_row
#         detailed_tl_df.loc[issue_id] = trustee_latest_df.loc[issue_id]


#     else:
#         new_row = trustee_latest_df.loc[issue_id]
#         for col in new_row.index:
#             detailed_tl_df.at[issue_id, col] = new_row[col]
#             detailed_tl_df.loc[issue_id] = trustee_latest_df.loc[issue_id]    
