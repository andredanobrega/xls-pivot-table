import subprocess
import os
import pandas as pd
import datetime as dt
import yaml


class PivotTableParser():
    def __init__(self, catalog):
        self.catalog = catalog
        self.path = self.catalog['file_parameters']['xls_path']
        filename = self.catalog['file_parameters']['filename']
        self.file = filename
        self.worksheet = self.catalog['file_parameters']['worksheet']
        self.fullpath = os.path.join(self.path,filename)

    def convert_xls_to_xlsx(self):
        output_folder = os.path.abspath(self.catalog['file_parameters']['xlsx_path'])
        command = f'soffice --headless --convert-to xlsx --outdir {output_folder} {self.fullpath}'
        subprocess.run(command, shell=True)
    
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
with open('catalog.yaml', 'r', encoding = 'UTF-8') as file:
    catalog = yaml.safe_load(file)
parser = PivotTableParser(catalog)
parser.convert_xls_to_xlsx()
