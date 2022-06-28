from pathlib import Path  # libreria estandar Python 

import pandas as pd  
import plotly.express as px  
import xlwings as xw  

import plotly.io as pio
pio.templates.default = "plotly_white"

excel_path = Path.cwd() / "data" / "data.xlsx"
df = pd.read_excel(excel_path, sheet_name="Orders")

df.head()

df.info()

df["Month"] = df["Order Date"].dt.month
df["Profit Margin"] = df["Profit"] / df["Sales"]
df.head()

grouped_by_subcategory = df.groupby(by="Sub-Category", as_index=False).sum()
grouped_by_subcategory

sales_profit_bar = px.bar(grouped_by_subcategory,x="Sub-Category",y="Sales",color="Profit",color_continuous_scale=["red", "yellow", "green"],
    title="<b>Ventas & ganacias por Sub Category</b>",
)
sales_profit_bar.show()

sales_profit_scatter = px.scatter( df, x="Sales", y="Profit", color="Discount",title="<b>Descuento en Ventas/Ganacias</b>",)
sales_profit_scatter.show()

df_descuento = df.groupby("Sub-Category").agg({"Discount": "mean", "Profit": "sum"})
df_descuento

profit_discount_bar = px.bar(df_descuento, x=df_descuento.index, y="Discount", color="Profit",color_continuous_scale=["red", "yellow", "green"],
    title="<b>Significado de Descuento por Sub Category</b>",
)
profit_discount_bar.show()

output_dir_analysis = Path().cwd() / "OUTPUT" / "Analysis"
output_dir_charts = Path().cwd() / "OUTPUT" / "Charts"
output_dir_cities = Path().cwd() / "OUTPUT" / "Cities"

output_dir_analysis.mkdir(parents=True, exist_ok=True)
output_dir_charts.mkdir(parents=True, exist_ok=True)
output_dir_cities.mkdir(parents=True, exist_ok=True)

sales_profit_bar.write_html(str(output_dir_charts / "sales_profit_bar.html"))
sales_profit_scatter.write_html(str(output_dir_charts / "sales_profit_scatter.html"))
profit_discount_bar.write_html(str(output_dir_charts / "profit_discount_bar.html"))

# AUTOMATION EXAMPLE: Save each city in a separate workbook
for unique_value in df["City"].unique():
    df_output = df.query("City == @unique_value")
    df_output.to_excel(
        output_dir_cities / f"{unique_value}.xlsx",
        sheet_name=unique_value[:31],
        index=False,
    )

from pathlib import Path
import pandas as pd

excel_file = Path.cwd() / "bonus" / "Sales_Data.xlsx"
df = pd.read_excel(excel_file)
df

df = pd.read_excel(
    io=excel_file,
    engine='openpyxl',
    sheet_name='Overview',
    skiprows=3,
    usecols='B:L',
    nrows=105,
)
df.head()