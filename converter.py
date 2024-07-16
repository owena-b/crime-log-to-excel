import tabula
import pandas as pd

# optional adjustment to prevent dataframe print from truncating
pd.set_option('display.max_rows', 999)
pd.set_option('display.max_columns', 999)
pd.set_option('display.width', 999)

pdf_path = 'https://www.american.edu/finance/publicsafety/upload/2024-daily-crime-log.pdf'

dfs = tabula.read_pdf(pdf_path, pages='all')

# pdf only has headers on first page. 'dfs' is a list of dataframes, where each df is a page of data. each df reads
# the first line of each page's table as the header, so this block of code moves the pseudo 'header' to the top of the
# dataframe and reassigns the header from the first page's data.
if len(dfs) > 1:  # if there is more than one page ...
    for i in dfs[1:]:  # iterates over all pages except the first
        i.loc[-1] = i.columns  # appends the fake header as a row
        i.index = i.index + 1  # adjusts the df index
        i.sort_index(inplace=True)  # re-sorts the rows so that the fake header's new row is back at the top
        i.columns = dfs[0].columns  # assigns the first df's headers to this df

# concatenates the resulting dataframes, removing extra headers
df = pd.concat(dfs, axis=0).iloc[1:]

print(df)

# writes and saves the Excel sheet
xl = pd.ExcelWriter('Crime Log.xlsx')
df.to_excel(xl, index=False)
xl._save()
