import pandas as pd

# File paths
control_plan_path = r"C:\Users\jukk\OneDrive - COWI\Documents\Revision check\Arch MASTER Document Control Plan2.xlsx"
revision_summary_path = r"C:\Users\jukk\OneDrive - COWI\Documents\Revision check\Revision_Summary.xlsx"
output_path = r"C:\Users\jukk\OneDrive - COWI\Documents\Revision check\Updated_Arch_MASTER_Document_Control_Plan.xlsx"

# Read the Excel files into pandas DataFrames
control_plan_df = pd.read_excel(control_plan_path)
revision_summary_df = pd.read_excel(revision_summary_path)

# Get the drawing numbers from the 'NUMBER' column in the control plan and revision summary
control_plan_drawings = control_plan_df["NUMBER"].tolist()

# Loop through each row of the Revision Summary to match drawing numbers and fill in revisions
for index, row in revision_summary_df.iterrows():
    drawing_number = row["NUMBER"]
    
    # Check if the drawing number exists in the control plan
    if drawing_number in control_plan_drawings:
        # Loop through the grayed-out columns in the control plan (e.g., columns with dates)
        for column in control_plan_df.columns:
            if "C Revision" in column:  # Look for columns with 'C Revision'
                # Match and populate the C Revision data
                control_plan_df.loc[control_plan_df["NUMBER"] == drawing_number, column] = row[column]
            elif "P Revision" in column:  # Look for columns with 'P Revision'
                # Match and populate the P Revision data
                control_plan_df.loc[control_plan_df["NUMBER"] == drawing_number, column] = row[column]

# Save the updated DataFrame to a new Excel file
control_plan_df.to_excel(output_path, index=False)

print(f"✅ Updated Excel file saved to: {output_path}")
