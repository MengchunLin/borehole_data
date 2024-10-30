import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, simpledialog

process_data=[]
# 選擇檔案
def select_file():
    root = tk.Tk()
    root.withdraw()
    file= filedialog.askopenfilename(title="選擇檔案")
    return file

# 分類土壤類型
def classify_soil_type(Ic):
    Ic = round(Ic, 2)
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

# 標記差異
def mark(previous_data, current_data):
    mark_list = [''] * len(current_data)
    for i in range(len(current_data)):
        if i < len(previous_data) and previous_data[i] != current_data[i]:
            mark_list[i] = '*'
    return mark_list

# 數據組織 (土壤類型、厚度、Ic 平均)
def data_array(Soil_Type, Ic):
    layer, thickness, ic_avg = [], [], []
    current_soil_type = None
    current_count = 0

    for i in range(len(Soil_Type)):
        soil_type = Soil_Type[i]
        if soil_type != current_soil_type and current_soil_type is not None:
            layer.append(current_soil_type)
            thickness.append(current_count)
            ic_avg.append(np.mean(Ic[i - current_count:i]))
            current_count = 0

        current_soil_type = soil_type
        current_count += 1

    if current_soil_type is not None:
        layer.append(current_soil_type)
        thickness.append(current_count)
        ic_avg.append(np.mean(Ic[len(Soil_Type) - current_count:]))

    return [layer, thickness, ic_avg]

# 合併層
def merge_layer(soil_data, thickness_threshold):
    i = 0  # Start index at 0 and manually control the iteration
    while i < len(soil_data):
        if soil_data.iloc[i, 1] <= thickness_threshold:
            if i == len(soil_data) - 1:  # If it's the last row
                soil_data.iloc[i - 1, 1] += soil_data.iloc[i, 1]  # Merge with the previous row
                soil_data.iloc[i, 0] = soil_data.iloc[i - 1, 0]  # Adjust layer designation
            else:
                if soil_data.iloc[i + 1, 0] == soil_data.iloc[i - 1, 0]:  # Merge with surrounding layers
                    soil_data.iloc[i - 1, 1] += soil_data.iloc[i, 1]
                elif soil_data.iloc[i - 1, 0] != soil_data.iloc[i + 1, 0]:
                    if abs(soil_data.iloc[i - 1, 2] - soil_data.iloc[i, 2]) > abs(soil_data.iloc[i, 2] - soil_data.iloc[i + 1, 2]):
                        soil_data.iloc[i + 1, 1] += soil_data.iloc[i, 1]  # Merge with the next row
                    else:
                        soil_data.iloc[i - 1, 1] += soil_data.iloc[i, 1]  # Merge with the previous row
            
            # Drop the current row after merging
            soil_data = soil_data.drop(i).reset_index(drop=True)
            i -= 1  # Move back one index to compensate for the dropped row
        
        i += 1  # Move to the next index

    return soil_data

# 簡化合併soil data
def merge_processed_data(soil_data):
    i = 0
    while i < len(soil_data) - 1:
        # 如果相邻两行的土壤类型相同
        if soil_data.iloc[i, 0] == soil_data.iloc[i + 1, 0]:
            # 合并厚度
            soil_data.iloc[i, 1] += soil_data.iloc[i + 1, 1]
            
            # 删除合并后的行
            soil_data = soil_data.drop(i + 1).reset_index(drop=True)
        else:
            # 只有在没有合并的情况下才增加索引
            i += 1
    
    return soil_data



# 生成最終數據
def write_merged_data(soil_data):
    data_input = []
    for i in range(len(soil_data)):
        soil_type = soil_data.iloc[i, 0]
        thickness = int(soil_data.iloc[i, 1])
        data_input.extend([soil_type] * thickness)
    return data_input

def process_file(file):
    df = pd.read_excel(file, header=0)
    df_copy = df.copy()

    # 資料處理
    df_copy['Ic'] = df_copy['Ic'].interpolate(method='linear')
    df_copy['Soil Type'] = df_copy['Soil Type'].ffill()

    # 分類土壤類型
    Soil_Type_5 = df_copy['Ic'].apply(classify_soil_type)
    df_copy['Soil Type 5 type'] = Soil_Type_5
    df_copy['Mark1'] = ''

    # 計算層數、厚度和 Ic 平均值
    layers, thicknesses, ic_avgs = data_array(Soil_Type_5, df_copy['Ic'])
    result_df = pd.DataFrame({'Soil Type': layers, 'Thickness': thicknesses, 'Ic_avg': ic_avgs})

    # 第一次合併（合併厚度 <= 5cm）
    result_array1 = merge_layer(result_df, 5)

    # 寫入第一次處理後的數據
    data_input1 = write_merged_data(result_array1)
    df_copy['10cm'] = data_input1
    df_copy['Mark2']=''
    result_array1 = merge_processed_data(result_array1)
    print(result_array1)
    print('第一次合併完成')
    
    # 確保數據長度匹配
    if len(data_input1) > len(df_copy):
        data_input1 = data_input1[:len(df_copy)]  # 截斷數據以匹配長度
    elif len(data_input1) < len(df_copy):
        data_input1.extend([''] * (len(df_copy) - len(data_input1)))  # 填充空值以匹配長度


    #對比soil type 5 和 5cm
    mark_array = mark(Soil_Type_5, data_input1)

    # 標記第一次合併後的數據
    df_copy['Mark1'] = mark_array

    # 提示用戶輸入合併厚度
    root = tk.Tk()
    root.withdraw()
    thickness_threshold = (simpledialog.askinteger("合併層厚度", "請輸入合併厚度的閾值 (cm):")) / 2
    if thickness_threshold is None:
        thickness_threshold = 5  # 默認為5cm
    root.destroy()

    # 第二次合併（基於用戶輸入的厚度閾值）
    result_array2 = merge_layer(result_array1, thickness_threshold)

    # 寫入第二次處理後的數據
    data_input = write_merged_data(result_array2)

    # 確保數據長度匹配
    if len(data_input) > len(df_copy):
        data_input = data_input[:len(df_copy)]  # 截斷數據以匹配長度
    elif len(data_input) < len(df_copy):
        data_input.extend([''] * (len(df_copy) - len(data_input)))  # 填充空值以匹配長度

    df_copy['合併後'] = data_input


    # 標記第二次合併後的數據
    mark_array = mark(data_input1, data_input)
    df_copy['Mark2'] = mark_array  # 標記第二次合併的變化

    # 將處理後的資料存入新的 Excel 檔案
    processed_file = file.replace('.xlsx', '_processed.xlsx')
    process_data.append(processed_file)
    df_copy.to_excel(processed_file, index=False)
    print('資料處理完成')

# 主函數，讀取 Excel，處理數據，並導出結果
def main():
    file = select_file()
    process_file(file)
    file = select_file()
    process_file(file)
    print(process_data)

if __name__ == "__main__":
    main()