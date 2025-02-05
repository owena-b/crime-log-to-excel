import tabula
import pandas as pd
import requests
from bs4 import BeautifulSoup
from lxml import etree

# optional adjustment to prevent dataframe print from truncating
pd.set_option('display.max_rows', 999)
pd.set_option('display.max_columns', 999)
pd.set_option('display.width', 999)

r = requests.get('https://www.american.edu/police/index.cfm')

soup = BeautifulSoup(r.text, 'html.parser')

dom = etree.HTML(str(soup))

crime_log = dom.xpath('//*[@id="cta-F191C0865A22F644ECAB328F8E8E36C00048582448C260127530B4581F3D5D97"]/a[1]')

pdf_path = 'https://www.american.edu'+crime_log[0].attrib.get('href')

dfs = tabula.read_pdf(pdf_path, pages='all')

# pdf only has headers on first page. 'dfs' is a list of dataframes, where each df is a page of data. each df reads
# the first line of each page's table as the header, so this block of code moves the pseudo 'header' to the top of the
# dataframe and reassigns the header from the first page's data.
if len(dfs) > 1:  # if there is more than one page ...
    for df in dfs[1:]:  # iterates over all pages except the first
        df.loc[-1] = df.columns  # appends the fake header as a row
        df.index = df.index + 1  # adjusts the df index
        df.sort_index(inplace=True)  # re-sorts the rows so that the fake header's new row is back at the top
        df.columns = dfs[0].columns  # assigns the first df's headers to this df

# concatenates the resulting dataframes, removing extra headers
new_df = pd.concat(dfs, axis=0).iloc[1:]

print(new_df)

# writes and saves the Excel sheet
xl = pd.ExcelWriter('Crime Log.xlsx')
new_df.to_excel(xl, index=False)
xl._save()
