import pandas as pd

issue_log_df = pd.read_excel(r"C:\\Users\\Sharon YY Tseng\Desktop\\Model Office Trustee Issue Log 2024_LIVE.xlsx", 
                             sheet_name = "Model Office Issue Log")

mod_log_df = issue_log_df.copy()

# remove first 5 columns, and the last few columns 
mod_log_df.drop (columns = mod_log_df.columns[1:5], axis = 1, inplace = True)
mod_log_df.drop(columns = mod_log_df.columns[16:], axis=1, inplace = True)

# Extract individual trustee issue
trustee_ls = list(mod_log_df.iloc[3:,2].dropna().unique())

def extract_indi_log(input_df, trustee):
    output_df = input_df[input_df.iloc[:,2] == trustee]
    
    return output_df


for i in trustee_ls:
    output_df = extract_indi_log(mod_log_df, i)
    
    # Export to Excel 
    output_df.to_excel(f"{i}_Model_Office_Trustee_Issue_Log.xlsx")
    print(f"Export Completed:{i}")
    
    