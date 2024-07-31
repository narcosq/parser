import pandas as pd
from fuzzywuzzy import fuzz

file1 = 'housekg.xlsx'
file2 = 'gosstroy.xlsx'

df1 = pd.read_excel(file1)
df2 = pd.read_excel(file2)

key_columns = ["Наименование объекта", "Адрес объекта", "Количество этажей", "Заказчик"]

df1[key_columns] = df1[key_columns].astype(str)
df2[key_columns] = df2[key_columns].astype(str)

similarity_threshold = 50

def is_similar(row1, row2, key_columns, threshold):
    total_similarity = 0
    for col in key_columns:
        similarity = fuzz.token_sort_ratio(str(row1[col]), str(row2[col]))
        total_similarity += similarity
    average_similarity = total_similarity / len(key_columns)
    return average_similarity >= threshold

common_rows = []
for index1, row1 in df1.iterrows():
    for index2, row2 in df2.iterrows():
        if is_similar(row1, row2, key_columns, similarity_threshold):
            common_rows.append(row1)

common_rows_df = pd.DataFrame(common_rows, columns=df1.columns)

common_rows_df.to_excel('comparison.xlsx', index=False)

print("Сравнение завершено. Результаты сохранены в 'common_rows.xlsx'.")
