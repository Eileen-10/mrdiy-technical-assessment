import pandas as pd
from pathlib import Path

DATA_DIR = Path.cwd() / "data"
OUTPUT_FILE = DATA_DIR / "aggregated_sales_report.xlsx"

SALES_CSV = DATA_DIR / "sales.csv"
PRODUCT_CSV = DATA_DIR / "product.csv"
STORE_CSV = DATA_DIR / "store.csv"

# 1. Data Retrieval
def load_data():
    sales = pd.read_csv(SALES_CSV)
    product = pd.read_csv(PRODUCT_CSV)
    store = pd.read_csv(STORE_CSV)
    return sales, product, store


# 2. Data Processing
def merge_data(sales: pd.DataFrame, product: pd.DataFrame, store: pd.DataFrame) -> pd.DataFrame:
    df = sales.merge(product, on="product_code", how="left")
    df = df.merge(store, on="store_code", how="left")
    return df

def filter_data(df: pd.DataFrame, region: str = None, product_category: str = None) -> pd.DataFrame:
    if region:
        df = df[df["store_region"] == region]
    if product_category:
        df = df[df["product_category"] == product_category]
    return df


# 3. Sales Aggregation
def aggregate_by_region(df: pd.DataFrame) -> pd.DataFrame:
    df["sales_amount"] = df["sales_qty"] * df["price"]
    df["sales_cost"] = df["sales_qty"] * df["cost"]
    agg = df.groupby("store_region").agg(
        sales_qty=pd.NamedAgg(column="sales_qty", aggfunc="sum"),
        sales_amount=pd.NamedAgg(column="sales_amount", aggfunc="sum"),
        sales_cost=pd.NamedAgg(column="sales_cost", aggfunc="sum")
    ).reset_index()
    agg["profit"] = agg["sales_amount"] - agg["sales_cost"]
    return agg

def aggregate_by_category(df: pd.DataFrame) -> pd.DataFrame:
    df["sales_amount"] = df["sales_qty"] * df["price"]
    df["sales_cost"] = df["sales_qty"] * df["cost"]
    agg = df.groupby("product_category").agg(
        sales_qty=pd.NamedAgg(column="sales_qty", aggfunc="sum"),
        sales_amount=pd.NamedAgg(column="sales_amount", aggfunc="sum"),
        sales_cost=pd.NamedAgg(column="sales_cost", aggfunc="sum")
    ).reset_index()
    agg["profit"] = agg["sales_amount"] - agg["sales_cost"]
    return agg


# 4. Excel Export
def export_to_excel(region_df: pd.DataFrame, category_df: pd.DataFrame, output_file: Path = OUTPUT_FILE):
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        region_df.to_excel(writer, sheet_name="By_Region", index=False)
        category_df.to_excel(writer, sheet_name="By_Category", index=False)
        workbook = writer.book
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
            for col in ws.columns:
                max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max_length + 2


### Main Function ###
def main(region: str = None, product_category: str = None):
    sales, product, store = load_data()

    df = merge_data(sales, product, store)
    df = filter_data(df, region, product_category)

    region_report = aggregate_by_region(df)
    category_report = aggregate_by_category(df)

    export_to_excel(region_report, category_report)
    print(f"Aggregated report saved to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()