import pandas as pd

df = pd.read_excel("data.xlsx", engine="openpyxl")
print(df.head())

df["address"] = df["address"].astype("category")
df["user"] = df["user"].astype("category")
df["exchange"] = df["exchange"].astype("category")
df["type"] = df["type"].astype("category")

print("************")

#print(pd.pivot_table(df, index=["address", "exchange", "user", "type"], aggfunc=[len], margins=True))

print("************")
print("************")
df_pivot =pd.pivot_table(df, index=["address", "exchange", "user", "type"], aggfunc=[len])
print(df_pivot)
print("************")
#print(pd.pivot_table(df_pivot, index=["index"], values=["len"],aggfunc=[np.sum]))
print(df_pivot.info())

print("************")
Gt = df_pivot["len"].sum()
print(Gt)


print("************")
#create pivot table with margins
print(df_pivot.groupby(["len"]).sum())



print("************")

print(df_pivot)

# df_unstacked = df_pivot.set_index(["address", "exchange", "user", "type"]).unstack("len")

# print(df_unstacked)

# df_pivot["sum"] = df_pivot["len"] + df_pivot["len"]
# print(df_pivot)


address_sum = df_pivot.groupby(level='address')['len'].sum()

df_pivot['address_len_sum'] = df_pivot.index.get_level_values('address').map(address_sum)

print(df_pivot)

print("************")
df_sorted = df_pivot.sort_values(by='address_len_sum', ascending=False)
print(df_sorted)

df_sorted_reset = df_sorted.reset_index()


# Save the sorted DataFrame to an Excel file with table name 'pivot'
excel_filename = 'sorted_pivot_table.xlsx'
with pd.ExcelWriter(excel_filename, engine='xlsxwriter') as writer:
    # Write the DataFrame to Excel with header
    df_sorted_reset.to_excel(writer, sheet_name='Sheet3', index=False)

    # Get the xlsxwriter workbook and worksheet objects.
    # sheet_name = "pivot_sheet"
    workbook = writer.book
    worksheet = writer.sheets['Sheet3']

    # workbook = Workbook()
    # worksheet = workbook.create_sheet(title=sheet_name)



    # Get the dimensions of the DataFrame
    (max_row, max_col) = df_sorted_reset.shape

    # Add a table that covers the entire DataFrame range
    worksheet.add_table(0, 0, max_row, max_col - 1, {
        'name': 'pivot',
        'columns': [{'header': col} for col in df_sorted_reset.columns],
        'autofilter': False  # Disable autofilter
    })


print(f"\nDataFrame has been saved to {excel_filename}")


# with pd.ExcelWriter('sorted_pivot_table.xlsx') as writer:
#     df_sorted_reset.to_excel(writer, sheet_name="pivot_table22", index=False)

