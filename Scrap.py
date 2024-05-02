import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter

# Read the Excel file with the information about the rejection: number of rejected parts, Article number, Article denomination, cost per Nonconformity, date of appearance
# Read the excel sheet
file_path_Produkt_Journal = "Journal.xlsx"
sheet_name = "Product-Journal"
#df_Produkt_Journal = pd.read_excel(file_path_Produkt_Journal, sheet_name=sheet_name)
df_Produkt_Journal = pd.read_excel(file_path_Produkt_Journal, sheet_name=sheet_name, parse_dates=['Auftrittsdatum'])

# Load the "Used_Materials" Excel file coming from the ERP system
file_path_Used_Materials = "Used_Materials.xlsx"
#df_Used_Materials = pd.read_excel(file_path_Used_Materials)
df_Used_Materials = pd.read_excel(file_path_Used_Materials, parse_dates=['BookingDate'])


# Define your desired date range
start_date = '2024-01-01'
end_date = '2024-03-31'

# Filter the "Produkt-Journal" data by date range
filtered_Product_Journal = df_Product_Journal[(df_Product_Journal['Date of appearance'] >= start_date) & (df_Product_Journal['Date of appearance'] <= end_date)]

# Convert "Rejected quantity" column to numeric (if needed)
#filtered_Product_Journal["Rejected quantity"] = pd.to_numeric(filtered_Product_Journal["Rejected quantity"], errors="coerce")

# Filter relevant entries based on "Decision"
relevant_decisions = ["Scrap", "Rework", "Return to Supplier", "Sort"]
filtered_Produkt_Journal = filtered_Product_Journal[filtered_Product_Journal["Decision"].isin(relevant_decisions)]

#Exclude specific article numbers (e.g., "12345" and "67890")
excluded_article_numbers = [602231, 601217, 602240, 601900, 601754, 600970]
filtered_Product_Journal = filtered_Product_Journal[~filtered_Product_Journal["Article"].isin(excluded_article_numbers)]

print(filtered_Product_Journal["Date of appearance"].size)

# Group "Product-Journal" data by article number and sum the "Rejected quantity"
consolidated_Product_Journal = filtered_Product_Journal.groupby("Article")["Rejected quantity"].sum()

# Filter the "Used_Materials" data by User_ID (based on Benutzer-ID), if needed
user_id = "XXX"
filtered_Rejected quantity = df_Rejected quantity[df_vRejected quantity["Used-ID"] == user_id]

# Filter the "Rejected quantity" data by date range
filtered_Rejected quantity = df_Rejected quantity[(df_Rejected quantity['BookingDate'] >= start_date) & (df_Rejected quantity['BookingDate'] <= end_date)]

# Convert "Quantity" column to positive values
filtered_verbrauch["Quantity"] *= -1


# Group "Rejected quantity" data by article number and sum the "Quantity"
consolidated_Rejected quantity = filtered_Rejected quantity.groupby("Articlenr.")["Quantity"].sum()

# Merge the consolidated DataFrames
article_summary = pd.concat([consolidated_Product_Journal, consolidated_Rejected quantity], axis=1, join="inner")
article_summary.columns = ["Rejected quantity", "Quantity"]

print(article_summary)

# Calculate scrap rate as a percentage
article_summary["Scrap Rate (%)"] = (article_summary["Rejected quantity"] / article_summary["Quantity"]) * 100

# Sort by scrap rate in descending order
article_summary.sort_values(by="Scrap Rate (%)", ascending=False, inplace=True)

# Display the result
print(article_summary)

print("type(article_summary)")
print(type(article_summary))

# Sort by scrap rate in descending order
article_summary.sort_values(by="Scrap Rate (%)", ascending=False, inplace=True)

# Select the top 15 article numbers
top_15_article_summary = article_summary.head(20)

article_numbers = top_15_article_summary.index.tolist()
scrap_rates = top_15_article_summary["Scrap Rate (%)"].tolist()

# Remove decimal points and convert to strings
article_numbers = [str(int(num)) for num in article_numbers]

print(article_numbers)

# Create a bar chart
plt.figure(figsize=(10, 6))
bars=plt.bar(article_numbers, scrap_rates, color="C0")

# Add value labels to each bar
for bar in bars:
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width() / 2, height, f"{height:.2f}%", ha="center", va="bottom")

plt.xlabel("Article Number")
plt.ylabel("Scrap Rate (%)")
plt.title("Top 20 Article Numbers by Scrap Rate")
plt.xticks(article_numbers, rotation=45)  # Explicitly set the X-axis tick labels
plt.tight_layout()
plt.show()

