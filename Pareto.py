import pandas as pd
import matplotlib.pyplot as plt

# Read the Excel file with the information about the rejection: number of rejected parts, Article number, Article denomination, cost per Nonconformity, date of appearance
# Read the excel sheet
file_path = "NC-Journal.xlsx"
sheet_name = "Product-Journal"
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Filter data based on date range (replace with your desired start and end dates)
start_date = "2024-01-01"
end_date = "2024-03-31"
filtered_df = df[(df["Date of appearance"] >= start_date) & (df["Date of appearance"] <= end_date)]

# Group by article number and calculate total costs for each cost component
grouped_df = filtered_df.groupby("Article").agg(
    {"NC handling costs": "sum", "Dispositions Costs": "sum", "Scrap Costs": "sum", "Repair Cost": "sum", "Article denomination": "first"}
).reset_index()

# Calculate total costs
grouped_df["Total Costs"] = grouped_df["NC handling costs"] + grouped_df["Dispositions Costs"] + \
                             grouped_df["Scrap Costs"] + grouped_df["Repair Cost"]

# Sort by total costs in descending order
sorted_df = grouped_df.sort_values(by="Total Costs", ascending=False).head(15)

# Create the bar chart with stacked bars
plt.figure(figsize=(10, 6))

# Iterate over each cost component and plot stacked bars
bottom = None
for cost_component in ["NC handling costs", "Dispositions Costs", "Scrap Costs", "Repair Cost"]:
    plt.bar([f"{desc} - {int(art)}" for desc, art in zip(sorted_df["Article denomination"], sorted_df["Article"])],
            sorted_df[cost_component],
            label=cost_component,
            bottom=bottom)
    if bottom is None:
        bottom = sorted_df[cost_component]
    else:
        bottom += sorted_df[cost_component]

# Add total cost labels above each bar
for i, total_cost in enumerate(sorted_df["Total Costs"]):
    plt.text(i, total_cost + 10, f' {total_cost:.0f}', ha='center', va='bottom')

plt.xlabel("Article Denomination - Article Number")
plt.ylabel("Total Costs")
plt.title("Total Costs of Nonquality by Article Number")
plt.xticks(rotation=45, ha="right")
plt.legend()
plt.tight_layout()

# Save or display the chart (adjust as needed)
plt.savefig("nonquality_costs_chart_stacked_with_total_labels.png")
plt.show()
