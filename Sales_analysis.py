import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# READ MAIN EXCEL FILE (READ ONLY)
data = pd.read_excel(r"C:\Desktop\data analysis\Superstore_Main.xlsx", sheet_name="Sheet1")

# Clean column names
data.columns = data.columns.str.strip().str.replace(" ", "_")

# Convert Sales to numeric
data["Sales"] = pd.to_numeric(data["Sales"], errors="coerce")

#  ANALYSIS

# Region Sales
region_sales = data.groupby("Region")["Sales"].sum()

# Top Products
top_products = data.groupby("Product_Name")["Sales"].sum().sort_values(ascending=False)

# Category Sales
category_sales = data.groupby("Category")["Sales"].sum()

#  PRINT OUTPUT (TERMINAL)
print("\nRegion Sales:\n", region_sales)
print("\nTop Products:\n", top_products.head(10))
print("\nCategory Sales:\n", category_sales)
print("\nTop Product:", top_products.idxmax())


# GRAPHS
# Region Graph
plt.figure(figsize=(8, 5))
sns.barplot(x=region_sales.index, y=region_sales.values)
plt.title("Sales by Region")
plt.xlabel("Region")
plt.ylabel("Sales")
plt.show()

# Top Products Graph
plt.figure(figsize=(10, 5))
sns.barplot(x=top_products.head(10).index, y=top_products.head(10).values)
plt.title("Top 10 Products by Sales")
plt.xlabel("Product")
plt.ylabel("Sales")
plt.xticks(rotation=45)
plt.show()

# Category Graph
plt.figure(figsize=(8, 5))
sns.barplot(x=category_sales.index, y=category_sales.values)
plt.title("Sales by Category")
plt.xlabel("Category")
plt.ylabel("Sales")
plt.show()

# SAVE OUTPUT (SAFE FILE)

with pd.ExcelWriter(r"C:\Desktop\data analysis\Superstore_Output.xlsx") as writer:
    region_sales.to_excel(writer, sheet_name="Region_Sales")
    top_products.head(10).to_excel(writer, sheet_name="Top_Products")
    category_sales.to_excel(writer, sheet_name="Category_Sales")