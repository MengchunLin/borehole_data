import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
import shutil

def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    return file_path

def classify_soil_type(Ic):
    Ic= round(Ic, 2)
    if Ic <= 2.05:
        return 1
    elif 2.05 < Ic <= 2.3:
        return 2
    elif 2.3 < Ic < 2.6:
        return 3
    elif 2.6 < Ic < 2.95:
        return 4
    elif Ic >= 2.95:
        return 5
    
def mark_positions_in_excel(Soil_Type):
    # 讀取數據長度
    n = len(Soil_Type)
    # 初始化標記行，長度與數據相同，初始為空白
    mark_line = [''] * n

    # 遍歷數據
    i = 0
    while i < n:
        current_num = Soil_Type[i]
        count = 1
        
        # 計算當前數字的連續長度
        while i + count < n and Soil_Type[i + count] == current_num:
            count += 1

        # 檢查是否滿足條件
        if count <= 5:
            # 將這些位置在標記行中設置為 '*'
            for j in range(i, i + count):
                mark_line[j] = '*'

        # 移動到下一段不同的數字
        i += count

    # 將標記行添加到 DataFrame 作為下一行
    df_copy['Mark']=mark_line

# make an array include the soil type, thickness and Ic_avg
# data_array[0] is the soil type, data_array[1] is the thickness, data_array[2] is the Ic_avg
def data_array(Soil_Type, Ic):
    # Initialize lists to store the results
    layer = []
    thickness = []
    ic_avg = []

    current_soil_type = None
    current_count = 0

    for i in range(len(Soil_Type)):
        soil_type = Soil_Type[i]
        
        # If the current soil type is different from the previous one
        if soil_type != current_soil_type and current_soil_type is not None:
            # Append data for the previous layer
            layer.append(current_soil_type)
            thickness.append(current_count)
            ic_avg.append(np.mean(Ic[i - current_count:i]))

            # Reset the current count
            current_count = 0

        # Update the current soil type and count
        current_soil_type = soil_type
        current_count += 1

    # Handle the last layer after the loop ends
    if current_soil_type is not None:
        layer.append(current_soil_type)
        thickness.append(current_count)
        ic_avg.append(np.mean(Ic[len(Soil_Type) - current_count:]))

    # a soil type with a tickness and Ic_avg
    data_array=[layer, thickness, ic_avg]

    return data_array


def merge_layer(soil_data):
    while True:
        merged = False  # 用來檢查本次迭代是否進行了合併
        for i in range(len(soil_data) - 1, 0, -1):  # 從最後一行開始迴圈，避免越界
            if soil_data.iloc[i, 1] <= 5:  # 如果厚度小於等於5，進行合併
                merged = True  # 表示進行了合併操作
                
                if i == len(soil_data) - 1:  # 對於最後一行，與前一行合併
                    # 將厚度加到上一行
                    soil_data.iloc[i - 1, 1] += soil_data.iloc[i, 1]
                    soil_data.iloc[i, 0] = soil_data.iloc[i - 1, 0]  # 更新為上一行的類型
                    soil_data = soil_data.drop(i)  # 刪除當前行
                else:
                    # 檢查前後的類型是否相同
                    if soil_data.iloc[i + 1, 0] == soil_data.iloc[i - 1, 0]:
                        # 更新類型並將當前行的厚度加到上一行的厚度上
                        soil_data.iloc[i - 1, 1] += soil_data.iloc[i, 1]
                        soil_data = soil_data.drop(i)  # 刪除當前行
                    elif soil_data.iloc[i - 1, 0] != soil_data.iloc[i + 1, 0]:
                        # 比較 Ic_avg 並更新類型
                        if abs(soil_data.iloc[i - 1, 2] - soil_data.iloc[i, 2]) > abs(soil_data.iloc[i, 2] - soil_data.iloc[i + 1, 2]):
                            # 將厚度加到下一行
                            soil_data.iloc[i + 1, 1] += soil_data.iloc[i, 1]
                            soil_data.iloc[i, 0] = soil_data.iloc[i + 1, 0]  # 更新為下一行的類型
                        else:
                            # 將厚度加到上一行
                            soil_data.iloc[i - 1, 1] += soil_data.iloc[i, 1]
                            soil_data.iloc[i, 0] = soil_data.iloc[i - 1, 0]  # 更新為上一行的類型
                        soil_data = soil_data.drop(i)  # 刪除當前行
        
        # 重置索引，以避免刪除行後的混亂
        soil_data = soil_data.reset_index(drop=True)

        # 如果本次迭代沒有進行合併，則退出循環
        if not merged:
            break
    
    return soil_data
            

def write_merged_data(soil_data):
    data_input = []
    
    for i in range(len(soil_data)):
        soil_type = soil_data.iloc[i, 0]  # 提取土壤類型
        thickness = int(soil_data.iloc[i, 1])  # 提取厚度並轉為整數
        
        # 根據厚度重複輸入類型
        for _ in range(thickness):
            data_input.append(soil_type)
    return data_input



# 讀取 Excel 檔案
selected_file = select_file()
df = pd.read_excel(selected_file, header=0)
# 複製 DataFrame
df_copy = df.copy()

# 資料處理
Ic = df_copy['Ic'].interpolate(method='linear')
Ic = Ic.round(2)
df_copy['Ic'] = Ic
# if soil type is nan and previous = next, then soil type = previous
df_copy['Soil Type'] = df_copy['Soil Type'].fillna(method='ffill') 
Soil_Type_CECI=df_copy['Soil Type']
Soil_type=df_copy['Soil Type']
Soil_Type_5 = df_copy['Ic'].apply(classify_soil_type)
df_copy['Soil Type 5 type'] = Soil_Type_5

# mark the positions in the excel
mark_positions_in_excel(Soil_Type_5)
layers, thicknesses, ic_avgs = data_array(Soil_Type_5, Ic)
result_df = pd.DataFrame({'Soil Type': layers,'Thickness': thicknesses,'Ic_avg': ic_avgs})
result_array = merge_layer(result_df)
print(result_array)

# 取data_input的值，並將其轉換為dataframe
data_input=write_merged_data(result_array)
df_copy['5cm']=data_input




# 將處理後的資料存入新的 Excel 檔案
df_copy.to_excel('output.xlsx', index=False)
print('資料處理完成')