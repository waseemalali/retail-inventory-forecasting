# ==========================================================
# RETAIL STORE INVENTORY ANALYSIS
# Cleaned / Fixed Version
# ==========================================================

import time
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# ----------------------------------------------------------
# Display Settings
# ----------------------------------------------------------

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 1200)

plt.style.use("ggplot")

# ----------------------------------------------------------
# File Paths (relative to this script)
# ----------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
DATA_PATH = BASE_DIR / "data" / "retail_store_inventory.csv"
OUTPUT_PATH = BASE_DIR / "outputs" / "Retail_Inventory_Final.csv"
CHARTS_DIR = BASE_DIR / "charts"  # visualizations are also saved here as PNGs

SHOW_PLOTS = False  # True = also pop up each chart in a window (you must close each one for the script to continue)

# ----------------------------------------------------------
# Replenishment / Procurement Planning Assumptions
# The dataset has no supplier lead time or cost fields, so these
# are business assumptions. EDIT THEM to match your real terms.
# ----------------------------------------------------------

LEAD_TIME_DAYS = 7          # supplier lead time in days
SERVICE_LEVEL_Z = 1.65       # Z-score for target service level (1.65 = ~95%)
ORDERING_COST = 50           # fixed cost per purchase order, in $
HOLDING_COST_RATE = 0.20     # annual holding cost as a % of unit price
DEMAND_WINDOW_DAYS = 90      # how many most-recent days to use for demand statistics


# ==========================================================
# Load Dataset
# ==========================================================

def load_data():

    print("=" * 70)
    print("Loading Dataset...")
    print("=" * 70)

    if not DATA_PATH.exists():
        raise FileNotFoundError(
            f"Could not find the CSV file at: {DATA_PATH.resolve()}\n"
            f"-> Check DATA_PATH at the top of the script and make sure "
            f"the file really exists at that exact location."
        )

    df = pd.read_csv(DATA_PATH)

    print(f"Dataset Shape : {df.shape}")
    print(f"Rows          : {len(df):,}")
    print(f"Columns       : {len(df.columns)}")

    return df


# ==========================================================
# Data Cleaning
# ==========================================================

def clean_data(df):

    print("\nCleaning Dataset...\n")

    print("Missing Values")
    print(df.isnull().sum())

    print("\nDuplicate Rows :", df.duplicated().sum())

    # Remove duplicate rows
    df = df.drop_duplicates()

    # Remove missing values only from important columns
    critical_columns = [
        "Inventory Level",
        "Units Sold",
        "Units Ordered",
        "Demand Forecast",
        "Price",
    ]

    df = df.dropna(subset=critical_columns)

    # Convert Date column
    df["Date"] = pd.to_datetime(df["Date"])

    print("\nDataset Shape After Cleaning :", df.shape)

    return df


# ==========================================================
# Feature Engineering
# ==========================================================

def feature_engineering(df):

    print("\nCreating New Features...\n")

    # Date Features
    df["Year"] = df["Date"].dt.year
    df["Quarter"] = df["Date"].dt.quarter
    df["Month"] = df["Date"].dt.month
    df["Month Name"] = df["Date"].dt.month_name()
    df["Day"] = df["Date"].dt.day

    # Business Features
    df["Revenue"] = df["Units Sold"] * df["Price"]

    df["Revenue After Discount"] = df["Revenue"] * (1 - df["Discount"] / 100)

    df["Forecast Error"] = df["Units Sold"] - df["Demand Forecast"]

    df["Inventory Difference"] = df["Inventory Level"] - df["Demand Forecast"]

    df["Inventory Turnover"] = np.where(
        df["Inventory Level"] == 0,
        0,
        df["Units Sold"] / df["Inventory Level"],
    )

    df["Procurement Efficiency"] = np.where(
        df["Units Ordered"] == 0,
        0,
        df["Units Sold"] / df["Units Ordered"],
    )

    # Inventory Status
    conditions = [
        df["Inventory Difference"] < 0,
        df["Inventory Difference"] > df["Demand Forecast"],
    ]

    choices = ["Low Stock", "Over Stock"]

    df["Inventory Status"] = np.select(conditions, choices, default="Normal")

    print("Feature Engineering Completed.")

    return df


# ==========================================================
# Save Dataset
# ==========================================================

def save_dataset(df):

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUTPUT_PATH, index=False)

    print("\nClean Dataset Saved Successfully.")
    print(f"Location : {OUTPUT_PATH.resolve()}")


# ==========================================================
# Save Excel Report
# (requires: pip install openpyxl)
# ==========================================================

def save_to_excel(df, eda, kpis, correlation, abc, insights, plan):

    excel_path = OUTPUT_PATH.with_suffix(".xlsx")

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:

        # Full cleaned dataset
        df.to_excel(writer, sheet_name="Cleaned Data", index=False)

        # KPIs
        kpi_df = pd.DataFrame(list(kpis.items()), columns=["KPI", "Value"])
        kpi_df.to_excel(writer, sheet_name="KPIs", index=False)

        # Business Insights
        insights_df = pd.DataFrame(list(insights.items()), columns=["Insight", "Value"])
        insights_df.to_excel(writer, sheet_name="Business Insights", index=False)

        # EDA Summary Tables
        eda["category_sales"].to_frame("Units Sold").to_excel(writer, sheet_name="Sales by Category")
        eda["region_sales"].to_frame("Units Sold").to_excel(writer, sheet_name="Sales by Region")
        eda["monthly_sales"].to_frame("Units Sold").to_excel(writer, sheet_name="Monthly Sales")
        eda["category_revenue"].to_frame("Revenue").to_excel(writer, sheet_name="Revenue by Category")
        eda["region_revenue"].to_frame("Revenue").to_excel(writer, sheet_name="Revenue by Region")
        eda["top_products"].to_frame("Revenue").to_excel(writer, sheet_name="Top 10 Products")
        eda["inventory_status"].to_frame("Count").to_excel(writer, sheet_name="Inventory Status")

        # ABC Analysis
        abc.to_excel(writer, sheet_name="ABC Analysis", index=False)

        # Correlation Matrix
        correlation.to_excel(writer, sheet_name="Correlation Matrix")

        # Replenishment / Procurement Plan
        plan.to_excel(writer, sheet_name="Replenishment Plan", index=False)

        critical_only = plan[plan["Replenishment Status"] == "Critical - Order Now"]
        critical_only.to_excel(writer, sheet_name="Critical - Order Now", index=False)

    print("\nExcel Report Saved Successfully.")
    print(f"Location : {excel_path.resolve()}")


# ==========================================================
# Exploratory Data Analysis (EDA)
# ==========================================================

def perform_eda(df):

    print("\n" + "=" * 70)
    print("EXPLORATORY DATA ANALYSIS")
    print("=" * 70)

    print("\nDataset Information")
    print("-" * 70)
    print(df.info())

    print("\nNumerical Statistics")
    print(df.describe().round(2))

    print("\nMissing Values")
    print(df.isnull().sum())

    print("\nOverall Business KPIs")
    print("-" * 70)

    print(f"Total Revenue              : ${df['Revenue'].sum():,.2f}")
    print(f"Revenue After Discount     : ${df['Revenue After Discount'].sum():,.2f}")
    print(f"Total Units Sold           : {df['Units Sold'].sum():,}")
    print(f"Total Inventory            : {df['Inventory Level'].sum():,}")
    print(f"Total Orders               : {df['Units Ordered'].sum():,}")
    print(f"Average Selling Price      : ${df['Price'].mean():.2f}")
    print(f"Average Discount           : {df['Discount'].mean():.2f}%")

    category_sales = df.groupby("Category")["Units Sold"].sum().sort_values(ascending=False)
    region_sales = df.groupby("Region")["Units Sold"].sum().sort_values(ascending=False)
    store_sales = df.groupby("Store ID")["Units Sold"].sum().sort_values(ascending=False)
    season_sales = df.groupby("Seasonality")["Units Sold"].mean().sort_values(ascending=False)
    weather_sales = df.groupby("Weather Condition")["Units Sold"].mean().sort_values(ascending=False)
    promotion_sales = df.groupby("Holiday/Promotion")["Units Sold"].mean()

    month_order = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December",
    ]

    monthly_sales = (
        df.groupby("Month Name")["Units Sold"].sum().reindex(month_order)
    )

    category_revenue = df.groupby("Category")["Revenue"].sum().sort_values(ascending=False)
    region_revenue = df.groupby("Region")["Revenue"].sum().sort_values(ascending=False)

    inventory_status = df["Inventory Status"].value_counts()

    top_products = (
        df.groupby("Product ID")["Revenue"].sum().sort_values(ascending=False).head(10)
    )

    return {
        "category_sales": category_sales,
        "region_sales": region_sales,
        "store_sales": store_sales,
        "season_sales": season_sales,
        "weather_sales": weather_sales,
        "promotion_sales": promotion_sales,
        "monthly_sales": monthly_sales,
        "category_revenue": category_revenue,
        "region_revenue": region_revenue,
        "inventory_status": inventory_status,
        "top_products": top_products,
    }


# ==========================================================
# KPI ANALYSIS
# ==========================================================

def calculate_kpis(df):

    print("\n" + "=" * 70)
    print("KEY PERFORMANCE INDICATORS")
    print("=" * 70)

    kpis = {}

    kpis["Total Revenue"] = df["Revenue"].sum()
    kpis["Revenue After Discount"] = df["Revenue After Discount"].sum()
    kpis["Average Revenue"] = df["Revenue"].mean()

    kpis["Total Units Sold"] = df["Units Sold"].sum()
    kpis["Average Units Sold"] = df["Units Sold"].mean()
    kpis["Total Orders"] = df["Units Ordered"].sum()

    kpis["Average Inventory"] = df["Inventory Level"].mean()
    kpis["Average Inventory Turnover"] = df["Inventory Turnover"].mean()
    kpis["Average Procurement Efficiency"] = df["Procurement Efficiency"].mean()

    forecast_accuracy = (
        1 - (abs(df["Forecast Error"]).sum() / df["Demand Forecast"].sum())
    ) * 100
    kpis["Forecast Accuracy"] = forecast_accuracy

    total_rows = len(df)
    low_stock = (df["Inventory Status"] == "Low Stock").sum()
    over_stock = (df["Inventory Status"] == "Over Stock").sum()
    normal_stock = (df["Inventory Status"] == "Normal").sum()

    kpis["Low Stock %"] = (low_stock / total_rows) * 100
    kpis["Over Stock %"] = (over_stock / total_rows) * 100
    kpis["Normal Stock %"] = (normal_stock / total_rows) * 100

    kpis["Average Discount"] = df["Discount"].mean()
    kpis["Maximum Discount"] = df["Discount"].max()

    kpis["Average Price"] = df["Price"].mean()
    kpis["Average Competitor Price"] = df["Competitor Pricing"].mean()

    print()
    for key, value in kpis.items():
        if "%" in key:
            print(f"{key:<35}: {value:.2f}%")
        elif "Revenue" in key:
            print(f"{key:<35}: ${value:,.2f}")
        else:
            print(f"{key:<35}: {value:,.2f}")

    return kpis


# ==========================================================
# CORRELATION & BUSINESS INSIGHTS
# ==========================================================

def correlation_analysis(df):

    print("\n" + "=" * 70)
    print("CORRELATION ANALYSIS")
    print("=" * 70)

    numeric_columns = [
        "Inventory Level",
        "Units Sold",
        "Units Ordered",
        "Demand Forecast",
        "Price",
        "Discount",
        "Competitor Pricing",
        "Revenue",
        "Revenue After Discount",
        "Forecast Error",
        "Inventory Difference",
        "Inventory Turnover",
        "Procurement Efficiency",
    ]

    # keep only columns that actually exist, in case feature engineering
    # was changed upstream - avoids a hard KeyError crash
    numeric_columns = [c for c in numeric_columns if c in df.columns]

    correlation = df[numeric_columns].corr().round(2)

    strongest_positive = (
        correlation["Revenue"].drop("Revenue").sort_values(ascending=False).head(5)
    )
    strongest_negative = (
        correlation["Revenue"].drop("Revenue").sort_values().head(5)
    )

    print("\nTop Positive Correlations With Revenue")
    print(strongest_positive)

    print("\nTop Negative Correlations With Revenue")
    print(strongest_negative)

    return correlation


# ==========================================================
# ABC ANALYSIS
# ==========================================================

def abc_analysis(df):

    print("\n" + "=" * 70)
    print("ABC INVENTORY ANALYSIS")
    print("=" * 70)

    abc = (
        df.groupby("Product ID", as_index=False)["Revenue"]
        .sum()
        .sort_values("Revenue", ascending=False)
    )

    abc["Cumulative Revenue"] = abc["Revenue"].cumsum()

    total_revenue = abc["Revenue"].sum()
    abc["Cumulative %"] = abc["Cumulative Revenue"] / total_revenue

    abc["ABC Class"] = np.where(
        abc["Cumulative %"] <= 0.80,
        "A",
        np.where(abc["Cumulative %"] <= 0.95, "B", "C"),
    )

    print("\nABC Distribution")
    print(abc["ABC Class"].value_counts().sort_index())

    return abc


# ==========================================================
# REPLENISHMENT & PROCUREMENT PLAN
# (sourcing/procurement recommendations driven by current
#  inventory levels + historical demand + demand forecast)
# ==========================================================

def build_replenishment_plan(df, abc):

    print("\n" + "=" * 70)
    print("REPLENISHMENT & PROCUREMENT PLAN")
    print("=" * 70)

    print(
        f"Assumptions -> Lead Time: {LEAD_TIME_DAYS} days | "
        f"Service Level Z: {SERVICE_LEVEL_Z} | "
        f"Ordering Cost: ${ORDERING_COST} | "
        f"Holding Cost Rate: {HOLDING_COST_RATE:.0%}/year | "
        f"Demand Window: last {DEMAND_WINDOW_DAYS} days"
    )
    print("(Edit these constants near the top of the script to match your real supplier terms)")

    # --- Demand statistics from the most recent window, per Store + Product ---
    cutoff_date = df["Date"].max() - pd.Timedelta(days=DEMAND_WINDOW_DAYS)
    recent = df[df["Date"] >= cutoff_date]

    demand_stats = (
        recent.groupby(["Store ID", "Product ID"])["Units Sold"]
        .agg(["mean", "std"])
        .rename(columns={"mean": "Avg Daily Demand", "std": "Demand Std Dev"})
    )
    demand_stats["Demand Std Dev"] = demand_stats["Demand Std Dev"].fillna(0)

    # --- Current snapshot (latest date on record) per Store + Product ---
    latest = (
        df.sort_values("Date")
        .groupby(["Store ID", "Product ID"])
        .tail(1)
        .set_index(["Store ID", "Product ID"])
    )
    latest = latest[["Date", "Category", "Region", "Inventory Level", "Demand Forecast", "Price"]]

    plan = latest.join(demand_stats).reset_index()

    # --- Safety Stock & Reorder Point (classic inventory formulas) ---
    plan["Safety Stock"] = SERVICE_LEVEL_Z * plan["Demand Std Dev"] * np.sqrt(LEAD_TIME_DAYS)

    plan["Reorder Point"] = (plan["Avg Daily Demand"] * LEAD_TIME_DAYS) + plan["Safety Stock"]

    # --- Economic Order Quantity ---
    annual_demand = plan["Avg Daily Demand"] * 365
    holding_cost_per_unit = (plan["Price"] * HOLDING_COST_RATE).replace(0, np.nan)

    plan["EOQ"] = np.sqrt(
        (2 * annual_demand * ORDERING_COST) / holding_cost_per_unit
    ).fillna(0)

    # --- Days of Supply remaining at current sales pace ---
    plan["Days of Supply"] = np.where(
        plan["Avg Daily Demand"] > 0,
        plan["Inventory Level"] / plan["Avg Daily Demand"],
        np.nan,
    )

    # --- Replenishment status ---
    conditions = [
        plan["Inventory Level"] < plan["Safety Stock"],
        plan["Inventory Level"] < plan["Reorder Point"],
    ]
    choices = ["Critical - Order Now", "Reorder Soon"]
    plan["Replenishment Status"] = np.select(conditions, choices, default="Adequate Stock")

    # --- Suggested order quantity / cost ---
    plan["Suggested Order Qty"] = np.where(
        plan["Replenishment Status"] != "Adequate Stock",
        np.maximum(plan["Reorder Point"] + plan["EOQ"] - plan["Inventory Level"], 0),
        0,
    )
    plan["Estimated Order Cost"] = plan["Suggested Order Qty"] * plan["Price"]

    # --- Bring in ABC class so sourcing can prioritize high-value products ---
    abc_lookup = abc.set_index("Product ID")["ABC Class"]
    plan["ABC Class"] = plan["Product ID"].map(abc_lookup)

    status_rank = {"Critical - Order Now": 0, "Reorder Soon": 1, "Adequate Stock": 2}
    abc_rank = {"A": 0, "B": 1, "C": 2}
    plan["_status_rank"] = plan["Replenishment Status"].map(status_rank)
    plan["_abc_rank"] = plan["ABC Class"].map(abc_rank)
    plan = plan.sort_values(["_status_rank", "_abc_rank", "Days of Supply"])
    plan = plan.drop(columns=["_status_rank", "_abc_rank"]).reset_index(drop=True)

    round_cols = [
        "Avg Daily Demand", "Demand Std Dev", "Safety Stock", "Reorder Point",
        "EOQ", "Days of Supply", "Suggested Order Qty", "Estimated Order Cost",
    ]
    plan[round_cols] = plan[round_cols].round(1)

    # --- Console summary ---
    status_counts = plan["Replenishment Status"].value_counts()
    total_cost = plan["Estimated Order Cost"].sum()
    critical = plan[plan["Replenishment Status"] == "Critical - Order Now"]

    print("\nReplenishment Status Breakdown (Store x Product combinations)")
    print(status_counts)
    print(f"\nTotal Estimated Procurement Cost Needed : ${total_cost:,.2f}")
    print(f"Critical Items Needing Immediate Order  : {len(critical)}")

    if len(critical) > 0:
        print("\nTop 10 Most Urgent Items (lowest Days of Supply first)")
        print(
            critical[
                [
                    "Store ID", "Product ID", "Category", "Region",
                    "Inventory Level", "Days of Supply",
                    "Suggested Order Qty", "Estimated Order Cost", "ABC Class",
                ]
            ].head(10).to_string(index=False)
        )

    return plan


def visualize_replenishment(plan):

    plt.figure(figsize=(7, 5))
    plan["Replenishment Status"].value_counts().plot(
        kind="bar", color=["crimson", "orange", "seagreen"]
    )
    plt.title("Replenishment Status Distribution")
    plt.xlabel("Status")
    plt.ylabel("Store x Product Combinations")
    plt.xticks(rotation=15)
    _save_and_show("14_replenishment_status.png")

    plt.figure(figsize=(10, 6))
    cost_by_category = (
        plan.groupby("Category")["Estimated Order Cost"].sum().sort_values(ascending=False)
    )
    cost_by_category.plot(kind="bar", color="teal")
    plt.title("Estimated Procurement Cost by Category")
    plt.xlabel("Category")
    plt.ylabel("Estimated Cost ($)")
    plt.xticks(rotation=45)
    _save_and_show("15_procurement_cost_by_category.png")

    plt.figure(figsize=(10, 6))
    sns.histplot(data=plan, x="Days of Supply", bins=30, kde=True)
    plt.title("Days of Supply Distribution (Current Stock)")
    _save_and_show("16_days_of_supply_distribution.png")


# ==========================================================
# BUSINESS INSIGHTS
# ==========================================================

def business_insights(df):

    print("\n" + "=" * 70)
    print("BUSINESS INSIGHTS")
    print("=" * 70)

    best_region = df.groupby("Region")["Revenue"].sum().idxmax()
    best_category = df.groupby("Category")["Revenue"].sum().idxmax()
    highest_season = df.groupby("Seasonality")["Revenue"].mean().idxmax()
    best_weather = df.groupby("Weather Condition")["Revenue"].mean().idxmax()
    average_discount = df["Discount"].mean()

    print(f"Best Performing Region      : {best_region}")
    print(f"Best Selling Category       : {best_category}")
    print(f"Best Season                 : {highest_season}")
    print(f"Best Weather Condition      : {best_weather}")
    print(f"Average Discount            : {average_discount:.2f}%")

    return {
        "Best Region": best_region,
        "Best Category": best_category,
        "Best Season": highest_season,
        "Best Weather": best_weather,
        "Average Discount": average_discount,
    }


# ==========================================================
# DATA VISUALIZATION
# (single, de-duplicated version - the old script defined this
#  function twice; the first copy referenced `df` without ever
#  receiving it as a parameter, which would raise a NameError
#  if it had ever been the one actually called)
# ==========================================================

def _save_and_show(filename):
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)
    plt.tight_layout()
    plt.savefig(CHARTS_DIR / filename, dpi=150)
    if SHOW_PLOTS:
        plt.show()
    plt.close()


def create_visualizations(df, eda, correlation, abc):

    print("\n" + "=" * 70)
    print("CREATING VISUALIZATIONS")
    print("=" * 70)

    plt.figure(figsize=(12, 6))
    eda["category_sales"].plot(kind="bar", color="steelblue")
    plt.title("Units Sold by Category")
    plt.xlabel("Category")
    plt.ylabel("Units Sold")
    plt.xticks(rotation=45)
    _save_and_show("01_units_sold_by_category.png")

    plt.figure(figsize=(12, 6))
    eda["monthly_sales"].plot(marker="o", linewidth=3)
    plt.title("Monthly Sales Trend")
    plt.xlabel("Month")
    plt.ylabel("Units Sold")
    plt.grid(True)
    _save_and_show("02_monthly_sales_trend.png")

    plt.figure(figsize=(10, 6))
    eda["region_revenue"].plot(kind="bar", color="darkorange")
    plt.title("Revenue by Region")
    plt.xlabel("Region")
    plt.ylabel("Revenue")
    plt.xticks(rotation=30)
    _save_and_show("03_revenue_by_region.png")

    plt.figure(figsize=(7, 7))
    eda["inventory_status"].plot(kind="pie", autopct="%1.1f%%", startangle=90)
    plt.title("Inventory Status Distribution")
    plt.ylabel("")
    _save_and_show("04_inventory_status.png")

    plt.figure(figsize=(12, 6))
    eda["top_products"].plot(kind="bar", color="forestgreen")
    plt.title("Top 10 Products by Revenue")
    plt.xlabel("Product ID")
    plt.ylabel("Revenue")
    _save_and_show("05_top_products.png")

    plt.figure(figsize=(10, 6))
    sns.scatterplot(data=df, x="Price", y="Units Sold", alpha=0.5)
    plt.title("Price vs Units Sold")
    _save_and_show("06_price_vs_units_sold.png")

    plt.figure(figsize=(10, 6))
    sns.histplot(data=df, x="Revenue", bins=30, kde=True)
    plt.title("Revenue Distribution")
    plt.xlabel("Revenue")
    _save_and_show("07_revenue_distribution.png")

    plt.figure(figsize=(11, 6))
    sns.boxplot(data=df, x="Category", y="Revenue")
    plt.title("Revenue by Category")
    plt.xticks(rotation=45)
    _save_and_show("08_revenue_by_category_box.png")

    plt.figure(figsize=(10, 6))
    sns.histplot(data=df, x="Inventory Turnover", bins=30, kde=True)
    plt.title("Inventory Turnover Distribution")
    _save_and_show("09_inventory_turnover_distribution.png")

    plt.figure(figsize=(12, 9))
    sns.heatmap(correlation, annot=True, cmap="coolwarm", fmt=".2f", linewidths=0.5)
    plt.title("Correlation Matrix")
    _save_and_show("10_correlation_matrix.png")

    plt.figure(figsize=(7, 5))
    abc["ABC Class"].value_counts().sort_index().plot(
        kind="bar", color=["green", "orange", "red"]
    )
    plt.title("ABC Classification")
    plt.xlabel("Class")
    plt.ylabel("Number of Products")
    _save_and_show("11_abc_classification.png")

    plt.figure(figsize=(10, 6))
    sns.histplot(data=df, x="Forecast Error", bins=30, kde=True)
    plt.title("Forecast Error Distribution")
    _save_and_show("12_forecast_error_distribution.png")

    plt.figure(figsize=(10, 6))
    sns.scatterplot(data=df, x="Discount", y="Revenue", alpha=0.5)
    plt.title("Discount vs Revenue")
    _save_and_show("13_discount_vs_revenue.png")

    print(f"\nAll visualizations generated successfully and saved to: {CHARTS_DIR.resolve()}")


# ==========================================================
# MAIN FUNCTION
# ==========================================================

def main():

    start_time = time.time()

    try:
        df = load_data()
        df = clean_data(df)
        df = feature_engineering(df)

        eda = perform_eda(df)
        kpis = calculate_kpis(df)
        correlation = correlation_analysis(df)
        abc = abc_analysis(df)
        insights = business_insights(df)
        plan = build_replenishment_plan(df, abc)

        create_visualizations(df, eda, correlation, abc)
        visualize_replenishment(plan)

        save_dataset(df)
        save_to_excel(df, eda, kpis, correlation, abc, insights, plan)

        # also save the replenishment plan as its own standalone CSV
        replenishment_csv = OUTPUT_PATH.parent / "Replenishment_Plan.csv"
        plan.to_csv(replenishment_csv, index=False)
        print(f"\nReplenishment Plan CSV Saved : {replenishment_csv.resolve()}")

        execution_time = time.time() - start_time

        print("\n" + "=" * 70)
        print("PROJECT SUMMARY")
        print("=" * 70)

        print(f"Dataset Shape            : {df.shape}")
        print(f"Total Revenue            : ${kpis['Total Revenue']:,.2f}")
        print(f"Forecast Accuracy        : {kpis['Forecast Accuracy']:.2f}%")
        print(f"Best Region              : {insights['Best Region']}")
        print(f"Best Category            : {insights['Best Category']}")
        print(f"Critical Reorder Items   : {(plan['Replenishment Status'] == 'Critical - Order Now').sum()}")
        print(f"Total Procurement Needed : ${plan['Estimated Order Cost'].sum():,.2f}")
        print(f"Execution Time           : {execution_time:.2f} Seconds")

        print("\nProject Completed Successfully.")

    except FileNotFoundError as e:
        print("\nFile Not Found")
        print(e)

    except KeyError as e:
        print("\nA Column Was Missing From The Dataset")
        print(f"Missing column: {e}")
        print("-> Open the CSV and check the exact column names/spelling.")

    except Exception as e:
        print("\nAn Error Occurred During Execution")
        print(f"{type(e).__name__}: {e}")


# ==========================================================
# PROGRAM ENTRY
# ==========================================================

if __name__ == "__main__":
    main()