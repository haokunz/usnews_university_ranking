import pandas as pd

# Function to load translations from translations.txt
def load_translations(filename):
    translations = {}
    with open(filename, 'r', encoding='utf-8') as file:
        for line in file:
            if ': ' in line:
                key, value = line.strip().split(': ', 1)
                translations[key] = value
    return translations

# Function to translate a column using the provided translations
def translate_column(column, translations):
    return column.apply(lambda x: translations.get(x, x) if pd.notnull(x) else x)

# Load translations
translations = load_translations("/学校排名数据/数据/大学排名/翻译件/translations.txt")

# Load the original Excel file
original_excel_path = "/学校排名数据/数据/大学排名/原件/USNEWS/us_news_US_national_ranking.xlsx"
wb = pd.ExcelFile(original_excel_path)

# Create a writer for the new Excel file
new_excel_path = "/学校排名数据/数据/大学排名/翻译件/us_news_national_ranking_Translation.xlsx"
writer = pd.ExcelWriter(new_excel_path, engine='openpyxl')

# Iterate through each sheet and apply translations
for sheet_name in wb.sheet_names:
    df = pd.read_excel(original_excel_path, sheet_name=sheet_name)
    
    # Translate each column in the sheet
    for column in df.columns:
        df[column] = translate_column(df[column], translations)
    
    # Save the translated sheet to the new Excel file
    df.to_excel(writer, sheet_name=sheet_name, index=False)

# Save the new Excel file
writer.close()

print("Translation completed and saved to", new_excel_path)
