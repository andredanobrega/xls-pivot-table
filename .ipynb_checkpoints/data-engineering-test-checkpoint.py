import subprocess
import os
def convert_xls_to_xlsx(file_path, output_folder = 'files_xlsx'):
    output_folder = os.path.abspath(output_folder)
    command = f'soffice --headless --convert-to xlsx --outdir {output_folder} {file_path}'
    subprocess.run(command, shell=True)

# Usage
convert_xls_to_xlsx("files/vendas-combustiveis-m3.xls")
