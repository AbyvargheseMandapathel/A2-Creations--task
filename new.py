import pandas as pd
from deep_translator import GoogleTranslator
import time

def translate_to_english(value):
    if pd.isna(value):
        return ""
    try:
        return GoogleTranslator(source='auto', target='en').translate(value)
    except Exception as e:
        return value

def translate_df_to_english(df):
    translated_columns = [translate_to_english(col) for col in df.columns]
    df.columns = translated_columns

    total_rows = len(df)
    for i, (index, row) in enumerate(df.iterrows(), 1):
        print(f"Row {i}/{total_rows} translated")
        df.loc[index] = row.apply(translate_to_english)
    return df

def main():
    input_file_path = 'Order Export.xls'
    output_file_path = 'translated_file.xlsx'

    start = time.time()

    print("Reading Excel file...")
    df = pd.read_excel(input_file_path, dtype=str)

    print("Translating cells data and column headers to English...")
    df = translate_df_to_english(df)

    elapsed_time = time.time() - start
    print(f"Translation completed in {elapsed_time:.2f} seconds.")

    print("Saving to a new Excel file...")
    df.to_excel(output_file_path, index=False)

    print(f"New Excel file created at: {output_file_path}")

if __name__ == '__main__':
    main()
