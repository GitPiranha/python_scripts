import pandas as pd


def create_pivot_table(file_name="data.xlsx"):
    df = pd.read_excel(file_name, engine="openpyxl")
    print(df.head())

    df["address"] = df["address"].astype("category")
    df["user"] = df["user"].astype("category")
    df["exchange"] = df["exchange"].astype("category")
    df["type"] = df["type"].astype("category")

    df_pivot = pd.pivot_table(df, index=["address", "exchange", "user", "type"], observed=True, aggfunc=[len])

    print(df_pivot.info())

    address_sum = df_pivot.groupby(level='address', observed=True)['len'].sum()

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
        df_sorted_reset.to_excel(writer, sheet_name='Sheet1', index=False)

        # Access the xlsxwriter workbook and worksheet objects
        worksheet = writer.sheets['Sheet1']

        # Get the dimensions of the DataFrame
        (max_row, max_col) = df_sorted_reset.shape

        # Add a table that covers the entire DataFrame range
        worksheet.add_table(0, 0, max_row, max_col - 1, {
            'name': 'pivot',
            'columns': [{'header': col} for col in df_sorted_reset.columns],
            'autofilter': False  # Disable autofilter
        })

    print(f"\nDataFrame has been saved to {excel_filename}")


if __name__ == '__main__':
    create_pivot_table()
# with pd.ExcelWriter('sorted_pivot_table.xlsx') as writer:
#     df_sorted_reset.to_excel(writer, sheet_name="pivot_table22", index=False)

