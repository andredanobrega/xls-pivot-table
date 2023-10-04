import win32com.client
import os
import pandas as pd
import xlrd
import datetime as dt
import yaml

with open('catalog.yaml', 'r', encoding = 'UTF-8') as file:
    catalog = yaml.safe_load(file)

xlApp = win32com.client.Dispatch('Excel.Application')
class PivotTableParser():
    def __init__(self, catalog, xlApp):
        self.catalog = catalog
        path = self.catalog['file_parameters']['path']
        filename = self.catalog['file_parameters']['filename']
        file_format = self.catalog['file_parameters']['format']
        self.file = filename + file_format
        self.worksheet = self.catalog['file_parameters']['worksheet']
        self.fullpath = os.path.join(path,filename)
        self.xlApp = xlApp
        
    def get_pivot_table_data(self, pivot_table):
        wb = self.xlApp.Workbooks.Open(self.fullpath)
        pivot_table = wb.Worksheets(self.worksheet).PivotTables(pivot_table)
        pivot_cache = pivot_table.PivotCache()
        new_ws = wb.Worksheets.Add()
        new_pivot_table = new_ws.PivotTables().Add(PivotCache=pivot_cache, TableDestination=new_ws.Range('A1'))
        xls_parameters = self.catalog['xls_parameters']
        rows = self.catalog['file_parameters']['rows']
        columns = self.catalog['file_parameters']['columns']

        # Add fields to Rows area
        for r in rows:
            new_pivot_table.PivotFields(r).Orientation = xls_parameters['row']
            new_pivot_table.PivotFields(r).Position = rows.index(r) + 1
            new_pivot_table.PivotFields(r).Subtotals = [False]*12

        for c in columns:
            new_pivot_table.PivotFields(c).Orientation = xls_parameters['values']  # Column Field
            new_pivot_table.PivotFields(c).Subtotals = [False]*12

        # Save the workbook
        wb.SaveAs(f'{pivot_table}.xls')

        # Close the workbook
        wb.Close(SaveChanges=False)

    def pivot_table_parsing_xls(self, pivot_table):
        wb = xlrd.open_workbook(pivot_table + '.xls')
        sheet = wb.sheet_by_index(0)
        data = []
        column_names = [sheet.cell_value(0, col_num) for col_num in range(sheet.ncols)]
        data_cache = {}
        # Loop through rows to populate 'data'
        for row_num in range(1, sheet.nrows):  # Start from 1 to skip the header row
            row_data = {}
            for col_num in range(sheet.ncols):
                # Caches the info to make a 'ffill' method while parsing, which gives the complete DataFrame
                if sheet.cell_type(row_num, col_num) != 0:
                    data_cache[column_names[col_num]] = sheet.cell_value(row_num, col_num)
                elif col_num == sheet.ncols-1:
                    data_cache[column_names[col_num]] = 0
                cell_value = data_cache[column_names[col_num]]
                row_data[column_names[col_num]] = cell_value
            data.append(row_data)
        df = pd.DataFrame(data)
        return df
    
    def pivot_table_parsing_xlsx(self, pivot_table):
        df = pd.read_excel(pivot_table + '.xls')
        df['Total'] = df['Total'].fillna(0)
        df = df.ffill()
        return df

    def data_treating(self, df):
        df.columns = df.columns.str.lower().str.replace('í','i')
        
        # Rename columns
        columns_rename = self.catalog['maps']['columns']
        df = df.rename(columns = columns_rename)

        # Year column
        df.year = df.year.astype('str')
        df = df.loc[~(df.year.str.startswith('Total Soma'))].copy()
        df.year = df.year.str[:4]
        df.year = df.year.astype('int')

        # Month column
        df['month'] = df.dados.str[-3:]

        # Creating a map for converting months into numbers
        # This works because excel pivot tables sort months in an ascending pattern
        month_map = {}
        for i in range (len(df.month.unique())):
            month_map[df.month.unique()[i]] = i+1
        df.month = df.month.replace(month_map)

        df['year_month'] = df.apply(lambda row: dt.datetime(row.year,row.month,1), axis = 1)
        # UF column
        uf_map = self.catalog['maps']['uf']
        df.uf = df.uf.replace(uf_map)

        # Timestamp
        df['created_at'] = dt.datetime.now()


        return df[['year_month','uf','product','unit','volume','created_at']]
    
#    def consistency_check(self, pivot_table):

    def pipeline_out(self):
        pivot_tables = self.catalog['file_parameters']['pivot_tables']
        for pivot_table in pivot_tables:
            try:
                df = parser.pivot_table_parsing_xls(pivot_table)
            except:
                df = parser.pivot_table_parsing_xlsx(pivot_table)
            df = parser.data_treating(df)
            df.to_parquet('{}.parquet'.format(parser.catalog['out_file_name'][pivot_table]), index = False)

PivotTableParser(catalog, xlApp).get_pivot_table_data('Tabela dinâmica1')
xlApp.Quit()

