#%% ###
import pandas as pd
import os 
from src.parse_grouped_report import parse_grouped_report
import matplotlib.pyplot as plt

# df = pd.read_excel(r'E:\Personal\Ivan Yankov 2025\Издадени документи - детайлна 2000.01.01-2025.08.11.xls')
# %%
input = r'E:\Printconsult\data\input_job.xls'
sheet1 = 'Sheet1'
sheet1_1 = 'Sheet1 1'
df1 = parse_grouped_report(input, sheet=sheet1, keep_bulgarian_headers=False)
df1_1 = parse_grouped_report(input, sheet=sheet1_1, keep_bulgarian_headers=False)

# %%
df1
# %%
df1_1
# %%
df_combined = pd.concat([df1, df1_1], ignore_index=True)

# %%
df_combined
# %%
df_combined['Document No'].nunique()
df_combined['Doc Type'].nunique()
df_combined['Partner'].nunique()
df_combined['Company ID'].nunique()
df_combined['Item'].nunique()
df_combined['Payment Type'].nunique()
df_combined['User'].nunique()
# %%
# Extract date parts
df_combined["year"] = df_combined["Date"].dt.year
df_combined["month"] = df_combined["Date"].dt.month
df_combined["day"] = df_combined["Date"].dt.day

# Group by year, month, day and sum
# df_grouped = df_combined.groupby(["year", "month", "day"], as_index=False)["Line Total"].sum()
df_grouped = df_combined.groupby(["year"], as_index=False)["Line Total"].sum()
# df_grouped = df_combined.groupby(["month"], as_index=False)["Line Total"].sum()

# %%
df_grouped = df_combined.groupby(["month"], as_index=False)["Line Total"].sum()
plt.plot(df_grouped.index, df_grouped["Line Total"])
plt.xticks(df_grouped.index, df_grouped[["month"]].astype(str).agg("-".join, axis=1), rotation=45)
plt.ylabel("Line Total")
plt.title("Daily Line Total Over Time")
plt.show()
# %%
df_grouped = df_combined.groupby(["year"], as_index=False)["Line Total"].sum()
plt.plot(df_grouped.index, df_grouped["Line Total"])
plt.xticks(df_grouped.index, df_grouped[["year"]].astype(str).agg("-".join, axis=1), rotation=45)
plt.ylabel("Line Total")
plt.title("Daily Line Total Over Time")
plt.show()
# %%
df_grouped = df_combined.groupby(["year","month"], as_index=False)["Line Total"].sum()
plt.plot(df_grouped.index, df_grouped["Line Total"])
plt.xticks(df_grouped.index, df_grouped[["year","month"]].astype(str).agg("-".join, axis=1), rotation=45)
plt.ylabel("Line Total")
plt.title("Daily Line Total Over Time")
plt.show()
# %%

import matplotlib.cm as cm
import numpy as np

# Group by year-month and sum
df_grouped = df_combined.groupby(["year", "month"], as_index=False)["Line Total"].sum()

plt.figure(figsize=(10, 6))

# Get unique years and a viridis color map
years = sorted(df_grouped["year"].unique())
colors = cm.viridis(np.linspace(0, 1, len(years)))

# Plot each year's monthly trend with viridis colors
for color, year in zip(colors, years):
    yearly_data = df_grouped[df_grouped["year"] == year]
    plt.plot(yearly_data["month"], yearly_data["Line Total"], 
             marker='o', color=color, label=str(year))

plt.xticks(range(1, 13), 
           ["Jan", "Feb", "Mar", "Apr", "May", "Jun", 
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
plt.ylabel("Line Total")
plt.xlabel("Month")
plt.title("Seasonality Overlay: Monthly Line Total by Year")
plt.legend(title="Year")
plt.grid(True, linestyle="--", alpha=0.6)
plt.tight_layout()
plt.show()
#### investigate seasonality in more detail
# %%



df_combined.columns
# %%

# Group by Partner and sum the total
partner_totals = (
    df_combined
    .groupby("Partner", as_index=False)["Line Total"]
    .sum()
    .sort_values("Line Total", ascending=False)
)

# Calculate cumulative sum and percentage
partner_totals["Cumulative Total"] = partner_totals["Line Total"].cumsum()
partner_totals["Cumulative %"] = 100 * partner_totals["Cumulative Total"] / partner_totals["Line Total"].sum()

# Plot
plt.figure(figsize=(10, 6))
plt.bar(partner_totals["Partner"], partner_totals["Line Total"], color="skyblue", label="Partner Contribution")
plt.plot(partner_totals["Partner"], partner_totals["Cumulative %"], color="orange", marker="o", label="Cumulative %")

plt.xticks(rotation=90)
plt.ylabel("Line Total")
plt.title("Cumulative Contribution by Partner")
plt.grid(axis="y", linestyle="--", alpha=0.6)
plt.legend()
plt.tight_layout()
plt.show()

partner_totals.head(10)  # Show top 10 partners in the table
# %%

# --- CONFIG ---
top_n = 10   # show top N partners by total spend (change as needed)

# 1) Aggregate to yearly totals per Partner
df_year_partner = (
    df_combined
    .groupby(["Partner", "year"], as_index=False)["Line Total"].sum()
    .rename(columns={"Line Total": "year_total"})
)

# Ensure years are sorted and all partners have all years (missing -> 0)
years = sorted(df_year_partner["year"].unique())
pvt = (
    df_year_partner
    .pivot(index="Partner", columns="year", values="year_total")
    .reindex(columns=years)
    .fillna(0)
)

# Pick top partners by overall spend
top_partners = pvt.sum(axis=1).nlargest(top_n).index
pvt_top = pvt.loc[top_partners]

# 2) CUMULATIVE over years (left-to-right cumulative sum)
pvt_cum = pvt_top.cumsum(axis=1)

# --- Plot A: Cumulative spend per Partner over years ---
plt.figure(figsize=(10, 6))
colors = cm.viridis(np.linspace(0, 1, len(top_partners)))

for color, partner in zip(colors, pvt_cum.index):
    plt.plot(years, pvt_cum.loc[partner, years].values, marker='o', label=str(partner), color=color)

plt.title("Cumulative Orders by Partner (from start to each year)")
plt.xlabel("Year")
plt.ylabel("Cumulative Line Total")
plt.grid(True, linestyle="--", alpha=0.5)
plt.xticks(years, rotation=0)
plt.legend(title="Partner", ncols=2, fontsize=8)
plt.tight_layout()
plt.show()

# --- Plot B: Year-by-year totals per Partner (no cumulative) ---
plt.figure(figsize=(10, 6))
colors = cm.viridis(np.linspace(0, 1, len(top_partners)))

for color, partner in zip(colors, pvt_top.index):
    plt.plot(years, pvt_top.loc[partner, years].values, marker='o', label=str(partner), color=color)

plt.title("Year-by-Year Orders by Partner")
plt.xlabel("Year")
plt.ylabel("Annual Line Total")
plt.grid(True, linestyle="--", alpha=0.5)
plt.xticks(years, rotation=0)
plt.legend(title="Partner", ncols=2, fontsize=8)
plt.tight_layout()
plt.show()
# %%


rank_start = 270  # inclusive
rank_end = 300    # inclusive
rank_start = 20  # inclusive
rank_end = 40    # inclusive
# 1) Aggregate yearly totals per Partner
df_year_partner = (
    df_combined
    .groupby(["Partner", "year"], as_index=False)["Line Total"].sum()
    .rename(columns={"Line Total": "year_total"})
)

# Ensure years are sorted and all partners have all years (missing -> 0)
years = sorted(df_year_partner["year"].unique())
pvt = (
    df_year_partner
    .pivot(index="Partner", columns="year", values="year_total")
    .reindex(columns=years)
    .fillna(0)
)

# 2) Select partners by rank range (ranked by total spend)
total_spend_sorted = pvt.sum(axis=1).sort_values(ascending=False)
partners_in_range = total_spend_sorted.iloc[rank_start-1 : rank_end].index
pvt_range = pvt.loc[partners_in_range]

# 3) CUMULATIVE over years
pvt_cum = pvt_range.cumsum(axis=1)

# --- Plot A: Cumulative spend per Partner over years ---
plt.figure(figsize=(10, 6))
colors = cm.viridis(np.linspace(0, 1, len(partners_in_range)))

for color, partner in zip(colors, pvt_cum.index):
    plt.plot(years, pvt_cum.loc[partner, years].values, 
             marker='o', label=str(partner), color=color)

plt.title(f"Cumulative Orders by Partner (Ranks {rank_start}–{rank_end})")
plt.xlabel("Year")
plt.ylabel("Cumulative Line Total")
plt.grid(True, linestyle="--", alpha=0.5)
plt.xticks(years, rotation=0)
plt.legend(title="Partner", ncols=2, fontsize=8)
plt.tight_layout()
plt.show()

# --- Plot B: Year-by-year totals per Partner ---
plt.figure(figsize=(10, 6))
colors = cm.viridis(np.linspace(0, 1, len(partners_in_range)))

for color, partner in zip(colors, pvt_range.index):
    plt.plot(years, pvt_range.loc[partner, years].values, 
             marker='o', label=str(partner), color=color)

plt.title(f"Year-by-Year Orders by Partner (Ranks {rank_start}–{rank_end})")
plt.xlabel("Year")
plt.ylabel("Annual Line Total")
plt.grid(True, linestyle="--", alpha=0.5)
plt.xticks(years, rotation=0)
plt.legend(title="Partner", ncols=2, fontsize=8)
plt.tight_layout()
plt.show()
# %%
df_combined
# %%

# --- CONFIG ---
rank_start = 1   # inclusive (e.g., 20)
rank_end   = 30   # inclusive (e.g., 30)

# 1) Aggregate yearly totals per Item
df_year_item = (
    df_combined
    .groupby(["Item", "year"], as_index=False)["Line Total"].sum()
    .rename(columns={"Line Total": "year_total"})
)

# Ensure years are sorted and all items have all years (missing -> 0)
years = sorted(df_year_item["year"].unique())
pvt = (
    df_year_item
    .pivot(index="Item", columns="year", values="year_total")
    .reindex(columns=years)
    .fillna(0)
)

# 2) Select items by rank range (ranked by total spend across all years)
total_spend_sorted = pvt.sum(axis=1).sort_values(ascending=False)
items_in_range = total_spend_sorted.iloc[rank_start-1:rank_end].index
pvt_range = pvt.loc[items_in_range]

# 3) CUMULATIVE over years
pvt_cum = pvt_range.cumsum(axis=1)

# --- Plot A: Cumulative spend per Item over years ---
plt.figure(figsize=(10, 6))
colors = cm.viridis(np.linspace(0, 1, len(items_in_range)))

for color, item in zip(colors, pvt_cum.index):
    plt.plot(years, pvt_cum.loc[item, years].values,
             marker='o', label=str(item), color=color)

plt.title(f"Cumulative Orders by Item (Ranks {rank_start}–{rank_end})")
plt.xlabel("Year")
plt.ylabel("Cumulative Line Total")
plt.grid(True, linestyle="--", alpha=0.5)
plt.xticks(years)
plt.legend(title="Item", ncols=2, fontsize=8)
plt.tight_layout()
plt.show()

# --- Plot B: Year-by-year totals per Item ---
plt.figure(figsize=(10, 6))
colors = cm.viridis(np.linspace(0, 1, len(items_in_range)))

for color, item in zip(colors, pvt_range.index):
    plt.plot(years, pvt_range.loc[item, years].values,
             marker='o', label=str(item), color=color)

plt.title(f"Year-by-Year Orders by Item (Ranks {rank_start}–{rank_end})")
plt.xlabel("Year")
plt.ylabel("Annual Line Total")
plt.grid(True, linestyle="--", alpha=0.5)
plt.xticks(years)
plt.legend(title="Item", ncols=2, fontsize=8)
plt.tight_layout()
plt.show()
# %%

df_combined[df_combined.Item=='марка състав 1+1 /ribkoff/']

# %%
