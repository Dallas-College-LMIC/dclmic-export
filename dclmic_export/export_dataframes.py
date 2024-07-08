import pandas as pd
import xlsxwriter
import gspread as gs
from typing import Iterable


def friendlize(s: str) -> str:
    split = s.lower().split("_")
    return " ".join([x.capitalize() for x in split])


def save_dfs_as_xl(
    list_of_frames: list[pd.DataFrame],
    col_format: dict = {},
    path: str = "./",
    file_name: str = "file",
    sheet_names: list[str] = [],
    tab_names: dict = {},
    friendly_names: bool = True,
) -> None:
    """
    Save a list of pandas DataFrames as sheets in an Excel file.

    Parameters:
    - list_of_frames: a list of pandas DataFrames to save as sheets. Each DataFrame will be saved as a separate sheet in the same file.
    - path: the path to the directory where the Excel file will be saved.
    - file_name: the name of the Excel file (without the .xlsx extension).
    - sheet_names: (optional) a list of names for the sheets.
        - If not provided, the default sheet names will be used.
    - tab_names: (optional) a dictionary mapping sheet names to custom tab names, for sheets where the tab name should be different from the title.
    - friendly_names: (optional) a boolean indicating whether to format the column names as friendly names (e.g., "column_name" -> "Column Name").
    Returns:
    None
    """

    # loop through

    # create fullname
    name = path + file_name + ".xlsx"

    print(f"Saving to {name}...")
    # set up writer object

    # send each sheet to excel
    # send to excel
    writer = pd.ExcelWriter(name, engine="xlsxwriter")
    workbook = writer.book

    STYLEBOOK = {
        "thousands": workbook.add_format({"num_format": "#,###"}),
        "currency": workbook.add_format({"num_format": "$#,##0.00"}),
        "currency_int": workbook.add_format({"num_format": "$#,##0"}),
        "decimal": workbook.add_format({"num_format": "#,##0.00"}),
        "percent": workbook.add_format({"num_format": "0.00%"}),
        "percent_int": workbook.add_format({"num_format": "0%"}),
    }

    for i, bdf in enumerate(list_of_frames):
        if sheet_names:
            bdf.name = sheet_names[i]
            if bdf.name in tab_names:
                sheet_name = tab_names[bdf.name][:31]
            else:
                sheet_name = sheet_names[i][:31]
        else:
            bdf.name = f"Sheet{i}"
            sheet_name = bdf.name

        # excel ui breaks if sheet names are longer than 31 characters

        if friendly_names:
            sheet_name = friendlize(sheet_name)
        worksheet = workbook.add_worksheet(sheet_name)
        bdf.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            startrow=3,
            startcol=1,
            header=False,
        )

        # add column formats
        cell_format = workbook.add_format()
        cell_format.set_font_name("Calibri Light")
        cell_format.set_font_size(12)
        lastcol = xlsxwriter.utility.xl_col_to_name(len(bdf.columns))
        worksheet.set_column(f"A:{lastcol}", None, cell_format)

        # add header
        header_format = workbook.add_format(
            {
                "bold": True,
                "font_name": "Calibri Light",
                "font_size": 12,
                "valign": "bottom",
                "fg_color": "#ece7f2",
                "border": 1,
            }
        )

        # apply header format to each column
        for col_num, value in enumerate(bdf.columns.values):
            worksheet.write(2, col_num + 1, friendlize(value) if friendly_names else value, header_format)

        # add title merged cell
        merge_format = workbook.add_format(
            {
                "bold": True,
                "font_name": "Calibri Light",
                "font_size": 14,
                "valign": "bottom",
                "align": "center",
                "fg_color": "#a6bddb",
                "border": 1,
            }
        )

        lastcol = xlsxwriter.utility.xl_col_to_name(len(bdf.columns))
        title = f"{bdf.name}"
        if len(bdf.columns) > 1:
            worksheet.merge_range(f"B2:{lastcol}2", title, merge_format)
        else:
            worksheet.write(1, 1, title, merge_format)

        # TODO: add formatted source.

        # add source:
        # source_format = workbook.add_format({
        #     'italic':True,
        #     'text_wrap':True,
        #     'font_name':'Calibri Light',
        #     'font_size':11,
        #     'valign': 'top',
        #     'align': 'left',
        #     'fg_color': '#ffffff',
        #     'border': 1})
        # scol1 = xlsxwriter.utility.xl_col_to_name(len(bdf.columns)+3)
        # scol2 = xlsxwriter.utility.xl_col_to_name(len(bdf.columns)+5)

        # worksheet.merge_range(f'{scol1}2:{scol2}3', "Source: ESRI Business Analyst", source_format)

        # adjust widths
        # TODO: width based on title?
        for idx, col_name in enumerate(bdf.columns.values):
            series = bdf[col_name]

            max_len = (
                max(
                    (
                        series.astype(str).map(len).max(),  # len of largest item
                        len(str(series.name)),  # len of column name/header
                    )
                )
                + 6
            )  # adding a little extra space

            if bdf.name in col_format and col_name in col_format[bdf.name]:
                style = col_format[bdf.name][col_name]
                if style in STYLEBOOK:
                    num_format = STYLEBOOK[col_format[bdf.name][col_name]]
                else:
                    num_format = workbook.add_format({"num_format": style})
                # This basically checks if the format passed in is one in the stylebook. If not, it assumes that it's a custom format code and uses that.
                # This way we can keep the stylebook to the most used styles and just apply custom formats to columns we want to customize as needed, instead of bloating the stylebook.
            elif '%' in col_name or 'percent' in col_name.lower():
                num_format = STYLEBOOK["percent"]
            elif 'income' in col_name.lower() or 'salary' in col_name.lower() or 'wage' in col_name.lower():
                num_format = STYLEBOOK["currency"]
            elif series.dtype == "int64":
                num_format = STYLEBOOK["thousands"]
            elif series.dtype == "float64":
                if (series.fillna(-9999) % 1  == 0).all():
                    num_format = STYLEBOOK["thousands"]
                else:
                    num_format = STYLEBOOK["decimal"]
            else:
                num_format = None

            worksheet.set_column(
                first_col=idx + 1,
                last_col=idx + 1,
                width=max_len,
                cell_format=num_format,
            )

        # add filter
        worksheet.autofilter(2, 1, bdf.shape[0], bdf.shape[1])

        # save
    writer.close()

    print("Saved!")


def upload_to_google_sheets(
    dataframe,
    bookname="newsheet",
    sheet=0,
    new_book=True,
    clear=True,
    authfile=r"P:\RESEARCH\LMI\GIS Data Resources\XX - No Project\XX - notes\gsheets_key.json",
):
    print(
        f"Loading in df with {len(dataframe.columns)} columns and {len(dataframe)} rows"
    )
    # connect to service account
    try:
        # gc equals account connection
        gc = gs.service_account(filename=authfile)
    except Exception as e:
        print("Could not connect,check filename")
        print(f"Exception: {e}")
        return

    # make new book if needed
    if new_book:
        print(f"Creating new workbook '{bookname}'")
        sh = gc.create(bookname)
        sh.share("cgilchriest.dcccd@gmail.com", perm_type="user", role="writer")

    try:
        print("Opening workbook...")
        wks = gc.open(bookname)
    except Exception as e:
        print("Could not open workbook, check workbook name.")
        print(f"Exception: {e}")
        return

    # select worksheet
    print("Selecting worksheet...")
    if isinstance(sheet, int):
        worksheet = wks.get_worksheet(sheet)
    else:
        try:
            worksheet = wks.worksheet(sheet)
        except Exception as e:
            if sheet != 0:
                title = sheet
            else:
                title = "NewSheet"
            worksheet = wks.add_worksheet(title=title, rows=100000, cols=200)
            print(f"Exception: {e}")

    if not new_book and clear:
        worksheet.clear()

    try:
        print("Updating workbook...")
        returns = worksheet.update(
            [dataframe.columns.values.tolist()] + dataframe.fillna("").values.tolist()
        )
        print(
            f"Success!\nSpreadsheet ID: {returns['spreadsheetId']}\nRows added: {returns['updatedRows']}"
        )
    except Exception as e:
        print("Could not upload...")
        print(f"Exception: {e}")
    print("Done")


def upload_to_sql(
    crsr, con, df, table_name="test__", schema="dbo", drop=True, chunk_print_size=50
):
    """
    chunk_print_size = number of rows to print count when uploading
    """

    pd.options.mode.chained_assignment = None

    print("Writing DataFrame into sql table.")
    print(f"Table name: LMDW.{schema}.{table_name}")

    # rename columns
    df.columns = [
        y.replace(" ", "_").replace(":", "").replace("(", "").replace(")", "").lower()
        for y in df.columns
    ]

    # default all to varchars
    columns = df.columns.tolist()
    format_columns = [f"[{x}][VARCHAR](Max) NOT NULL" for x in columns]

    # adjust formatting
    for x in columns:
        df[x] = df[x].astype(str)
        # adjust datetime
        if x in df.select_dtypes(include=["datetime64[ns, UTC]"]).columns.tolist():
            df[x] = df[x].dt.strftime("%Y-%m-%d")

    # if the drop parameter is True, then drop the existing table
    if drop == True:
        try:
            print("Trying to drop old table.")
            crsr.execute(
                f"""IF EXISTS(SELECT * FROM sys.tables WHERE name like '%{table_name}%')  
            DROP TABLE [lmdw].[{schema}].[{table_name}]"""
            )
            con.commit()
            print("Dropped!")
        except Exception as e:
            print("No existing table to drop.")
            print(f"Exception: {e}")

        try:
            print("Creating new sql table...")
            crsr.execute(
                f"""IF NOT EXISTS(SELECT * FROM sys.tables WHERE name like '{table_name}') 
            CREATE TABLE [lmdw].[{schema}].[{table_name}](
                {",".join(format_columns)})"""
            )
            con.commit()
            print("Done.")
        except Exception as e:
            print("Could not create new table.")
            print(f"Exception: {e}")

    print(f"Writing in {len(df)} rows.")

    print("Uploading...")
    for index, row in df.reset_index().iterrows():
        index, row
        try:
            if index % chunk_print_size == 0:
                print(f"{index} out of {len(df)}")
            crsr.execute(
                f"""INSERT INTO [lmdw].[{schema}].[{table_name}] 
            ({",".join(columns)}) values ({",".join(["?"] * len(columns))})""",
                row[1:].tolist(),
            )
        except Exception as e:
            print(f"Could not upload {index}")
            print(f"Exception: {e}")

    con.commit()
    print("Done!")


def make_table_spatial(
    crsr,
    con,
    table_name="test__",
    schema="dbo",
    wkt_geom_col="geom_wkt",
    destination_crs="4326",
):
    print(f"Making [LMDW].[{schema}].[{table_name}] spatial and adding spatial index.")

    try:
        print("Create a geometry column.")
        crsr.execute(
            f"""ALTER TABLE [LMDW].[{schema}].[{table_name}] ADD geom as geometry::STGeomFromText({wkt_geom_col},{destination_crs}).MakeValid() PERSISTED;"""
        )
        con.commit()
    except Exception as e:
        print("Error on geometry column.")
        print(f"Exception: {e}")

    try:
        print("Create a spatial index.")
        crsr.execute(
            f"""ALTER TABLE [LMDW].[{schema}].[{table_name}] ADD PKEY_IDX INT IDENTITY;"""
        )
        crsr.execute(
            f"""ALTER TABLE [LMDW].[{schema}].[{table_name}] ADD CONSTRAINT {table_name}_constraint_sidx PRIMARY KEY CLUSTERED ([PKEY_IDX]);"""
        )
        con.commit()

        print("Setting spatial index...")
        crsr.execute(
            f"""CREATE SPATIAL INDEX spatial_idx_{table_name.lower()} ON [LMDW].[{schema}].[{table_name}](geom)
                            WITH 
                            ( 
                                    BOUNDING_BOX= (xmin=-99, ymin=32, xmax=-96, ymax=33) 
                            );"""
        )
    except Exception as e:
        print("Error on spatial index.")
        print(f"Exception: {e}")
    print("Done!")
    con.commit()
