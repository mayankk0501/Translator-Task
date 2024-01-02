import pandas as pd
import traceback
from mtranslate import translate
import concurrent.futures
import time

#############################################################################

start_time = time.time()

#############################################################################

def translate_cell(cell):
    try:
        if isinstance(cell, str):
            return translate(cell, 'en')
        else:
            return cell
    except Exception as e:
        print(traceback.format_exc())
        return 'errrrrrrrrrrrrrrrrrrrrrrrror'

#############################################################################

def translate_column(df, col):
    with concurrent.futures.ThreadPoolExecutor() as executor:
        translated_column = list(executor.map(translate_cell, df[col]))
    return translated_column

#############################################################################

file_path = 'Order Export.xls'
df = pd.read_excel(file_path)
translated_dataframe = pd.DataFrame()

try:
    with concurrent.futures.ThreadPoolExecutor() as executor:
        for col in df.columns:
            translated_column = translate_column(df, col)
            translated_dataframe[col] = translated_column
except Exception as e:
    print(traceback.format_exc())

translated_dataframe.columns = [translate(col, 'en') for col in translated_dataframe.columns]

output_file_path = 'output.xlsx'
translated_dataframe.to_excel(output_file_path, index=False)

#############################################################################

end_time = time.time()
elapsed_time = end_time - start_time
print(f"Time taken to execute the code: {elapsed_time} seconds")