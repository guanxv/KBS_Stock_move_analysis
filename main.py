import sys

print(sys.version)
print(sys.executable)

import pandas as pd
import openpyxl

pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", None)
pd.set_option("display.width", None)
pd.set_option("display.max_colwidth", None)

rdata1 = pd.read_excel(
    "Resources/Ageing Analysis 2020-02-24_Raw.xlsx", "Product Ageing Report"
)
rdata2 = pd.read_excel(
    "Resources/Ageing Analysis 2020-07-28_Raw.xlsx", "Product Ageing Report"
)

# General Enquiry of the DF
# print(rdata1.head())

data_name_1 = "Feb"
data_name_2 = "Jul"

# new_column_names_1 =  ['0-1' + data_name_1, '1-2' + data_name_1, '2-3' + data_name_1, '3-4' + data_name_1, '4+' + data_name_1, 'Total' + data_name_1,]

# join 2 df as one

rdata1.columns = [
    "Product",
    "Knum",
    "0-1" + data_name_1,
    "1-2" + data_name_1,
    "2-3" + data_name_1,
    "3-4" + data_name_1,
    "4+" + data_name_1,
    "Total" + data_name_1,
    "LastRun" + data_name_1,
    "UnitCost" + data_name_1,
    "ExtCost" + data_name_1,
]
rdata2.columns = [
    "Product",
    "Knum",
    "0-1" + data_name_2,
    "1-2" + data_name_2,
    "2-3" + data_name_2,
    "3-4" + data_name_2,
    "4+" + data_name_2,
    "Total" + data_name_2,
    "LastRun" + data_name_2,
    "UnitCost" + data_name_2,
    "ExtCost" + data_name_2,
]

# rdatamelted = pd.melt(rdata1, id_vars = [''])

# print(rdata1.head())
# print(rdata2.head())


# -------- Prepare Product info DF --------------

# generate Product info Dataframe, which contain Knum, Product , unit price for Feb and Jul
product_info_Feb = rdata1[["Product", "Knum", "UnitCost" + data_name_1]]
product_info_Jul = rdata2[["Product", "Knum", "UnitCost" + data_name_2]]

product_info = pd.merge(
    product_info_Feb, product_info_Jul, left_on="Knum", right_on="Knum", how="outer"
)

# print(product_info.head())

# check if the K number is same
product_info["Product_Name_Check"] = product_info.apply(
    lambda row: 0 if row["Product_x"] == row["Product_y"] else 1, axis=1
)

# looks like lots of k number is different.
# lets move the k number from old to New one, if new one is Nan.
product_info["Product_y"] = product_info.apply(
    lambda row: row["Product_x"]
    if not isinstance(row["Product_y"], str)
    else row["Product_y"],
    axis=1,
)

product_info = product_info.drop(columns=["Product_x", "Product_Name_Check"])
product_info = product_info.rename(columns={"Product_y": "Product"})

# re-arrange the product info columns sequence
product_info = product_info[["Knum", "Product", "UnitCostFeb", "UnitCostJul"]]

# print(product_info.head(1000))

# --------End of Prepare Product info DF --------------


# --------prepare the main dataframe --------------------

rdata1 = rdata1.drop(columns=["ExtCostFeb", "LastRunFeb", "UnitCostFeb", "Product"])
rdata2 = rdata2.drop(columns=["ExtCostJul", "LastRunJul", "UnitCostJul", "Product"])

data1melted = pd.melt(
    rdata1,
    id_vars=["Knum"],
    value_vars=[
        "0-1" + data_name_1,
        "1-2" + data_name_1,
        "2-3" + data_name_1,
        "3-4" + data_name_1,
        "4+" + data_name_1,
        "Total" + data_name_1,
    ],
    value_name="Qty",
    var_name="Qty_Type",
)

data2melted = pd.melt(
    rdata2,
    id_vars=["Knum"],
    value_vars=[
        "0-1" + data_name_2,
        "1-2" + data_name_2,
        "2-3" + data_name_2,
        "3-4" + data_name_2,
        "4+" + data_name_2,
        "Total" + data_name_2,
    ],
    value_name="Qty",
    var_name="Qty_Type",
)

df = pd.concat([data1melted, data2melted], ignore_index=True)

# drop all 0 value, drop Nan

df = df[df.Qty != 0]
df.reset_index(drop=True, inplace=True)

df = df.dropna(subset=["Knum"])

# get correct unit price
# Defind func for find unitcost

"""
def findunitcost(knum , costmonth):
    if costmonth == 'Feb':
        return product_info[product_info.Knum == knum].UnitCostFeb.iloc[0]
    elif costmonth == 'Jul':
        return product_info[product_info.Knum == knum].UnitCostJul.iloc[0]

"""
# better way to define findunitcost function


def findunitcost(knum, costmonth):
    col = "UnitCost" + costmonth
    return product_info.loc[product_info["Knum"] == knum][col].iloc[0]


writeunitcost = lambda row: findunitcost(row["Knum"], row["Qty_Type"][-3:])

# apply find unit cost function.
df["Unit_Cost"] = df.apply(writeunitcost, axis=1)

# creat ext cost column
df["Ext_Cost"] = df["Qty"] * df["Unit_Cost"]

summary = df.groupby("Qty_Type").Ext_Cost.sum()

# summary.columns = ['Date_Range','Total Ext Cost']

df01feb = df[df.Qty_Type == "0-1Feb"]
df01jul = df[df.Qty_Type == "0-1Jul"]
df12feb = df[df.Qty_Type == "1-2Feb"]
df12jul = df[df.Qty_Type == "1-2Jul"]
df23feb = df[df.Qty_Type == "2-3Feb"]
df23jul = df[df.Qty_Type == "2-3Jul"]
df34feb = df[df.Qty_Type == "3-4Feb"]
df34jul = df[df.Qty_Type == "3-4Jul"]
df4pfeb = df[df.Qty_Type == "4+Feb"]
df4pjul = df[df.Qty_Type == "4+Jul"]
dftotfeb = df[df.Qty_Type == "TotalFeb"]
dftotjul = df[df.Qty_Type == "TotalJul"]


# print(product_info.head(10))
# print(df.head(10))
# print(summary.head(30))


def comparesameage(df1, df2):

    # merge the dfs.
    comparedf = pd.merge(
        df1,
        df2,
        left_on="Knum",
        right_on="Knum",
        how="outer",
        suffixes=["_Feb", "_Jul"],
    )

    # find first Non Nan index
    Feb_First_Non_Nan = comparedf["Qty_Type_Feb"].first_valid_index()
    Jul_First_Non_Nan = comparedf["Qty_Type_Jul"].first_valid_index()

    # fill nan for qty type column
    comparedf = comparedf.fillna(
        value={
            "Qty_Type_Feb": comparedf.iloc[Feb_First_Non_Nan]["Qty_Type_Feb"],
            "Qty_Type_Jul": comparedf.iloc[Jul_First_Non_Nan]["Qty_Type_Jul"],
        }
    )
    # fill the rest with 0 (qty , unitcost, exit cost)
    comparedf = comparedf.fillna(0)
    # workout the diff for Qty and Exit Cost
    comparedf["Qty_Diff"] = comparedf["Qty_Jul"] - comparedf["Qty_Feb"]
    comparedf["Ext_Cost_Diff"] = comparedf["Ext_Cost_Jul"] - comparedf["Ext_Cost_Feb"]

    # drop useless columns
    comparedf = comparedf.drop(columns=["Qty_Type_Feb", "Qty_Type_Jul"])

    # sort by ext_cost_diff
    comparedf = comparedf.sort_values(by=["Ext_Cost_Diff"], ascending=True)

    # bring in the product description
    comparedf = pd.merge(
        comparedf, product_info, left_on="Knum", right_on="Knum", how="inner",
    )

    comparedf.drop(columns=["UnitCostFeb", "UnitCostJul"])

    # re-arrange the columns
    comparedf = comparedf[
        [
            "Product",
            "Knum",
            "Qty_Feb",
            "Qty_Jul",
            "Qty_Diff",
            "Unit_Cost_Feb",
            "Unit_Cost_Jul",
            "Ext_Cost_Feb",
            "Ext_Cost_Jul",
            "Ext_Cost_Diff",
        ]
    ]

    return comparedf


# genarate Result DF
df01 = comparesameage(df01feb, df01jul)
df12 = comparesameage(df12feb, df12jul)
df23 = comparesameage(df23feb, df23jul)
df34 = comparesameage(df34feb, df34jul)
df4p = comparesameage(df4pfeb, df4pjul)
dftot = comparesameage(dftotfeb, dftotjul)


def trackgetold(df1, df2):

    # track things cross age group
    dfcompare = pd.merge(
        df1,
        df2,
        left_on="Knum",
        right_on="Knum",
        how="inner",
        suffixes=["_Feb", "_Jul"],
    )

    dfcompare = dfcompare.sort_values(by=["Ext_Cost_Jul"], ascending=False)

    # bring in the product description
    dfcompare = pd.merge(
        dfcompare, product_info, left_on="Knum", right_on="Knum", how="inner",
    )

    dfcompare.drop(columns=["UnitCostFeb", "UnitCostJul"])

    # re-arrange the columns
    dfcompare = dfcompare[
        [
            "Product",
            "Knum",
            "Qty_Type_Feb",
            "Qty_Feb",
            "Unit_Cost_Feb",
            "Ext_Cost_Feb",
            "Qty_Type_Jul",
            "Qty_Jul",
            "Unit_Cost_Jul",
            "Ext_Cost_Jul",
        ]
    ]

    return dfcompare


df01_12 = trackgetold(df01feb, df12jul)
df12_23 = trackgetold(df12feb, df23jul)
df23_34 = trackgetold(df23feb, df34jul)
df34_4p = trackgetold(df34feb, df4pjul)


# track qty reduced k num
dftotqtyreduce = pd.merge(
    dftotfeb,
    dftotjul,
    left_on="Knum",
    right_on="Knum",
    how="inner",
    suffixes=["_Feb", "_Jul"],
)

dftotqtyreduce = dftotqtyreduce[dftotqtyreduce.Qty_Jul < dftotqtyreduce.Qty_Feb]
dftotqtyreduce = dftotqtyreduce.sort_values(by=["Ext_Cost_Jul"], ascending=False)

# bring in the product description
dftotqtyreduce = pd.merge(
    dftotqtyreduce, product_info, left_on="Knum", right_on="Knum", how="inner",
)

dftotqtyreduce.drop(columns=["UnitCostFeb", "UnitCostJul"])

# re-arrange the columns
dftotqtyreduce = dftotqtyreduce[
    [
        "Product",
        "Knum",
        "Qty_Type_Feb",
        "Qty_Feb",
        "Unit_Cost_Feb",
        "Ext_Cost_Feb",
        "Qty_Type_Jul",
        "Qty_Jul",
        "Unit_Cost_Jul",
        "Ext_Cost_Jul",
    ]
]


# write result to DF

with pd.ExcelWriter("output.xlsx") as writer:
    summary.to_excel(writer, sheet_name="Summary")
    df01.to_excel(writer, sheet_name="01_Feb_Jul")
    df12.to_excel(writer, sheet_name="12_Feb_Jul")
    df23.to_excel(writer, sheet_name="23_Feb_Jul")
    df34.to_excel(writer, sheet_name="34_Feb_Jul")
    df4p.to_excel(writer, sheet_name="4+_Feb_Jul")
    dftot.to_excel(writer, sheet_name="Tot_Feb_Jul")
    df01_12.to_excel(writer, sheet_name="01Feb==>12Jul")
    df12_23.to_excel(writer, sheet_name="12Feb==>23Jul")
    df23_34.to_excel(writer, sheet_name="23Feb==>34Jul")
    df34_4p.to_excel(writer, sheet_name="34Jul==>4+Jul")
    dftotqtyreduce.to_excel(writer, sheet_name="Feb-Jul_Qty_Reduced")

"""
print(df01.head(50))
print(df12.head(50))
print(df23.head(50))
print(df34.head(50))
print(df4p.head(50))
print(dftot.head(50))

print(df01_12.head(50))
print(df01_12.count())
print(df01_12.Knum.nunique())
print(df01_12.Ext_Cost_Jul.sum())

print(dftotqtyreduce.head(50))
"""

