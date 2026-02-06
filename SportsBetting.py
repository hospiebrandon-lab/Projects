import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter, MultipleLocator

# load the excel file
file_path = "/Users/brandon/downloads/SportsBettingByMonth.xlsx"
df1 = pd.read_excel(file_path, sheet_name = "Total Bets by Month")
df2 = pd.read_excel(file_path, sheet_name = "Total Wages by Month")

# convert data to time
df1["date"] = pd.to_datetime(df1["date"])
df2["date"] = pd.to_datetime(df2["date"])

# plot findings for sheet 1
plt.figure()
plt.plot(df1["date"],df1["count"], marker="o")
plt.axhline(y=20000, color='red', linestyle='--', linewidth=2)
plt.xlabel("Date")
plt.ylabel("Total Bets Placed")
plt.title("Total Bets by Month")
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# plot findings for sheet 2
plt.figure(figsize=(12,8))
plt.bar(df2["date"], df2["totalwages"], color="green", width=10)
plt.xlabel("Date")
plt.ylabel("Total Wages")
plt.title("Total Wages by Month")
plt.xticks(rotation=45)

million = 1_000_000
plt.gca().yaxis.set_major_locator(MultipleLocator(million))
plt.gca().yaxis.set_major_formatter(FuncFormatter(lambda x, _: f'{int(x):,}'))
plt.gca().grid(True, axis="y", linestyle = "--", linewidth = 2, color = "black")

plt.tight_layout()
plt.show()

