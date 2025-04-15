import pandas as pd

# Load Excel file
df = pd.read_excel("Yola_tasks.xlsx")

df.rename(columns={"Desidantion": "Designation"}, inplace=True)
df = df[df["Hours_task"] > 0]
df["Priority"].fillna(99, inplace=True)
df = df.sort_values(by=["Type", "Priority"])

with pd.ExcelWriter("Project_Timesheet_By_Technician.xlsx", engine="openpyxl") as writer:
    for tech_type, group in df.groupby("Type"):
        group.to_excel(writer, sheet_name=tech_type, index=False)
        total = group["Hours_task"].sum()
        total_df = pd.DataFrame([["", "TOTAL", total, "", ""]], columns=group.columns)
        total_df.to_excel(writer, sheet_name=tech_type, index=False, startrow=len(group)+2, header=False)
