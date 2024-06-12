#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import os

def load_data(file_path):
    return pd.read_excel(file_path)

def filter_fixations(df):
    df_fixation = df[df['Eye movement type'] == 'Fixation'].copy()
    df_fixation['LookZone'] = 'Out'
    return df_fixation

def get_aoi_columns():
    return [
        'AOI hit [1 - sci aru]', 'AOI hit [1 - sci aru_eng]', 'AOI hit [1 - Pic]',
        'AOI hit [2 - def]', 'AOI hit [3 - B]', 'AOI hit [3 - C]', 'AOI hit [3 - D]',
        'AOI hit [3 - intro]', 'AOI hit [3 - framework3]', 'AOI hit [3 - Q]', 'AOI hit [3 - R]',
        'AOI hit [3 - W]', 'AOI hit [4 - B]', 'AOI hit [4 - C]', 'AOI hit [4 - D]', 'AOI hit [4 - def_C]',
        'AOI hit [4 - framework4]', 'AOI hit [4 - Q]', 'AOI hit [4 - R]', 'AOI hit [4 - W]',
        'AOI hit [5 - B]', 'AOI hit [5 - C]', 'AOI hit [5 - D]', 'AOI hit [5 - def_D]',
        'AOI hit [5 - framework5]', 'AOI hit [5 - Q]', 'AOI hit [5 - R]', 'AOI hit [5 - W]',
        'AOI hit [6 - B]', 'AOI hit [6 - C]', 'AOI hit [6 - D]', 'AOI hit [6 - def_W]',
        'AOI hit [6 - framework6]', 'AOI hit [6 - Q]', 'AOI hit [6 - R]', 'AOI hit [6 - W]',
        'AOI hit [7 - B]', 'AOI hit [7 - C]', 'AOI hit [7 - D]', 'AOI hit [7 - def_B]',
        'AOI hit [7 - framework7]', 'AOI hit [7 - Q]', 'AOI hit [7 - R]', 'AOI hit [7 - W]',
        'AOI hit [8 - B]', 'AOI hit [8 - C]', 'AOI hit [8 - D]', 'AOI hit [8 - def_Q]',
        'AOI hit [8 - framework8]', 'AOI hit [8 - Q]', 'AOI hit [8 - R]', 'AOI hit [8 - W]',
        'AOI hit [9 - B]', 'AOI hit [9 - C]', 'AOI hit [9 - D]', 'AOI hit [9 - def_R9]',
        'AOI hit [9 - framework9]', 'AOI hit [9 - Q]', 'AOI hit [9 - R]', 'AOI hit [9 - W]',
        'AOI hit [10 - def]', 'AOI hit [2_1 - def]', 'AOI hit [3_1 - B]', 'AOI hit [3_1 - C]',
        'AOI hit [3_1 - D]', 'AOI hit [3_1 - intro]', 'AOI hit [3_1 - framework3]', 'AOI hit [3_1 - Q]',
        'AOI hit [3_1 - R]', 'AOI hit [3_1 - W]', 'AOI hit [4_1 - B]', 'AOI hit [4_1 - C]',
        'AOI hit [4_1 - D]', 'AOI hit [4_1 - def_C]', 'AOI hit [4_1 - framework4]', 'AOI hit [4_1 - Q]',
        'AOI hit [4_1 - R]', 'AOI hit [4_1 - W]', 'AOI hit [5_1 - B]', 'AOI hit [5_1 - C]',
        'AOI hit [5_1 - D]', 'AOI hit [5_1 - def_D]', 'AOI hit [5_1 - framework5]', 'AOI hit [5_1 - Q]',
        'AOI hit [5_1 - R]', 'AOI hit [5_1 - W]', 'AOI hit [6_1 - B]', 'AOI hit [6_1 - C]',
        'AOI hit [6_1 - D]', 'AOI hit [6_1 - def_W]', 'AOI hit [6_1 - framework6]', 'AOI hit [6_1 - Q]',
        'AOI hit [6_1 - R]', 'AOI hit [6_1 - W]', 'AOI hit [7_1 - B]', 'AOI hit [7_1 - C]',
        'AOI hit [7_1 - D]', 'AOI hit [7_1 - def_B]', 'AOI hit [7_1 - framework7]', 'AOI hit [7_1 - Q]',
        'AOI hit [7_1 - R]', 'AOI hit [7_1 - W]', 'AOI hit [8_1 - B]', 'AOI hit [8_1 - C]',
        'AOI hit [8_1 - D]', 'AOI hit [8_1 - def_Q8]', 'AOI hit [8_1 - framework8]', 'AOI hit [8_1 - Q]',
        'AOI hit [8_1 - R]', 'AOI hit [8_1 - W]', 'AOI hit [9_1 - B]', 'AOI hit [9_1 - C]',
        'AOI hit [9_1 - D]', 'AOI hit [9_1 - def_R9]', 'AOI hit [9_1 - framework9]', 'AOI hit [9_1 - Q]',
        'AOI hit [9_1 - R]', 'AOI hit [9_1 - W]', 'AOI hit [12 (1) - 12-B1]', 'AOI hit [12 (1) - 12-C1]',
        'AOI hit [12 (1) - 12-D1]', 'AOI hit [12 (1) - 12-D2]', 'AOI hit [12 (1) - 12_article]',
        'AOI hit [12 (1) - 12-def]', 'AOI hit [12 (1) - 12-def_ b]', 'AOI hit [12 (1) - 12-def_ c]',
        'AOI hit [12 (1) - 12-def_ d]', 'AOI hit [12 (1) - 12-def_ q]', 'AOI hit [12 (1) - 12-def_ r]',
        'AOI hit [12 (1) - 12-def_ w]', 'AOI hit [12 (1) - 12-def_f]', 'AOI hit [12 (1) - 12-title]',
        'AOI hit [12 (2) - 12-B2]', 'AOI hit [12 (2) - 12-D3]', 'AOI hit [12 (2) - 12-W1]',
        'AOI hit [12 (2) - 12_article]', 'AOI hit [12 (2) - 12-def]', 'AOI hit [12 (2) - 12-def_ b]',
        'AOI hit [12 (2) - 12-def_ c]', 'AOI hit [12 (2) - 12-def_ d]', 'AOI hit [12 (2) - 12-def_ q]',
        'AOI hit [12 (2) - 12-def_ r]', 'AOI hit [12 (2) - 12-def_ w]', 'AOI hit [12 (2) - 12-def_f]',
        'AOI hit [14 - 14-B1]', 'AOI hit [14 - 14-C1]', 'AOI hit [14 - 14-C2]', 'AOI hit [14 - 14-D1]',
        'AOI hit [14 - 14-D2]', 'AOI hit [14 - 14_article]', 'AOI hit [14 - 14-def]', 'AOI hit [14 - 14-def_ b]',
        'AOI hit [14 - 14-def_ c]', 'AOI hit [14 - 14-def_ d]', 'AOI hit [14 - 14-def_ q]', 'AOI hit [14 - 14-def_ r]',
        'AOI hit [14 - 14-def_ w]', 'AOI hit [14 - 14-def_f]', 'AOI hit [14 - 14-title]', 'AOI hit [14 (3) - 14-B2]',
        'AOI hit [14 (3) - 14-D3]', 'AOI hit [14 (3) - 14-W1]', 'AOI hit [14 (3) - 14_article]', 'AOI hit [14 (3) - 14-def]',
        'AOI hit [14 (3) - 14-def_ b]', 'AOI hit [14 (3) - 14-def_ c]', 'AOI hit [14 (3) - 14-def_ d]', 'AOI hit [14 (3) - 14-def_ q]',
        'AOI hit [14 (3) - 14-def_ r]', 'AOI hit [14 (3) - 14-def_ w]', 'AOI hit [14 (3) - 14-def_f]', 'AOI hit [15 - 15]'
    ]

def classify_lookzone(row, aoi_columns):
    zones_with_one = [col for col in aoi_columns if row[col] == 1]
    non_framework_zones = [zone for zone in zones_with_one if 'framework' not in zone]
    framework_zones = [zone for zone in zones_with_one if 'framework' in zone]
    non_article_zones = [zone for zone in zones_with_one if 'article' not in zone]
    article_zones = [zone for zone in zones_with_one if 'article' in zone]

    if len(zones_with_one) >= 2:
        return non_framework_zones[0] if non_framework_zones else framework_zones[0]
        return non_article_zones[0] if non_article_zones else article_zones[0]
    elif len(zones_with_one) == 1:
        return zones_with_one[0]
    else:
        return 'Out'

def process_fixations(df_fixation, aoi_columns):
    for index, row in df_fixation.iterrows():
        df_fixation.at[index, 'LookZone'] = classify_lookzone(row, aoi_columns)
    return df_fixation

def remove_duplicates(df):
    return df.drop_duplicates(subset=['Fixation point X', 'Fixation point Y', 'Recording start time', 'Gaze event duration', 'LookZone'])

def save_to_excel(df, file_path):
    new_df = df[['Fixation point X', 'Fixation point Y', 'Recording start time', 'Gaze event duration', 'LookZone']]
    new_df.reset_index(drop=True, inplace=True)
    new_df.index = new_df.index + 1
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl') as writer:
        new_df.to_excel(writer, sheet_name='Summary1', index=True, index_label='Number')

def process_excel_files(directory):
    aoi_columns = get_aoi_columns()
    for filename in os.listdir(directory):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(directory, filename)
            df = load_data(file_path)
            df_fixation = filter_fixations(df)
            df_fixation = process_fixations(df_fixation, aoi_columns)
            df_fixation = remove_duplicates(df_fixation)
            save_to_excel(df_fixation, file_path)
            print(f"Processed {filename}")

if __name__ == "__main__":
    directory = '/Users/yu-tingguo/Downloads/Data Export - TAP Bird 23 (4)'  # 更改為包含 Excel 文件的目錄
    process_excel_files(directory)


# In[ ]:




