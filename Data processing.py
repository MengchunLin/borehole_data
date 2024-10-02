import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog

# 1. 選擇檔案
def select_file():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename()

# 2. 分類土壤類型
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

# 3. 標記差異
def mark(previous_data, current_data):
    mark_list = [''] * len(current_data)

    # 如果長度不一致，報錯並返回空標記列表
    if len(previous_data) != len(current_data):
        print("Error: Lengths of previous_data and current_data do not match!")
        return mark_list

    for i in range(len(current_data)):
        if previous_data[i] != current_data[i]:
            print(f'不同：{previous_data[i]} != {current_data[i]}')
            mark_list[i] = '*'  # 更新標記

    return mark_list

# 4. 數據組織 (土壤類型、厚度、Ic 平均)
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

# 5. 合併層
def merge_layer(soil_data, threshold):
    while True:
        merged = False
        for i in range(len(soil_data) - 1, 0, -1):
            if soil_data.iloc[i, 1] <= threshold:
                merged = True
                if i == len(soil_data) - 1:
                    soil_data.iloc[i - 1, 1] += soil_data.iloc[i, 1]
                    soil_data.iloc[i, 0] = soil_data.iloc[i - 1, 0]
                else:
                    if soil_data.iloc[i + 1, 0] == soil_data.iloc[i - 1, 0]:
                        soil_data.iloc[i - 1, 1] += soil_data.iloc[i, 1]
                    elif soil_data.iloc[i - 1, 0] != soil_data.iloc[i + 1, 0]:
                        if abs(soil_data.iloc[i - 1, 2] - soil_data.iloc[i, 2]) > abs(soil_data.iloc[i, 2] - soil_data.iloc[i + 1, 2]):
                            soil_data.iloc[i + 1, 1] += soil_data.iloc[i, 1]
                        else:
                            soil_data.iloc[i - 1, 1] += soil_data.iloc[i, 1]

                soil_data = soil_data.drop(i).reset_index(drop=True)

        if not merged:
            break

    return soil_data

# 6. 寫入處理後的數據
def write_merged_data(soil_data):
    data_input = []
    for i in range(len(soil_data)):
        soil_type = soil_data.iloc[i, 0]
        thickness = int(soil_data.iloc[i, 1])
        data_input.extend([soil_type] * thickness)
    return data_input

# 7. 建立一個視窗，輸入要合併的厚度
def merge_layer_window(soil_data):
    window = tk.Tk()
    window.title('合併層')
    window.geometry('300x200')

    label = tk.Label(window, text='輸入要合併的厚度')
    label.pack()

    entry = tk.Entry(window)
    entry.pack()

    threshold = []

    def merge_layers():
        threshold_value = int(entry.get())
        threshold.append(threshold_value)  # 保存輸入的數值
        window.destroy()

    button = tk.Button(window, text='確定', command=merge_layers)
    button.pack()

    window.mainloop()
    return threshold[0] if threshold else None  # 確保返回值有效

# 8. 主函數，讀取 Excel，處理數據，並導出結果
def main():
    # 讀取 Excel 檔案
    selected_file = select_file()
    df = pd.read_excel(selected_file, header=0)
    df_copy = df.copy()

    # 資料處理
    df_copy['Ic'] = df_copy['Ic'].interpolate(method='linear').round(2)
    df_copy['Soil Type'] = df_copy['Soil Type'].fillna(method='ffill')
    
    Soil_Type_CECI = df_copy['Soil Type']
    Soil_Type_5 = df_copy['Ic'].apply(classify_soil_type)
    df_copy['Soil Type 5 type'] = Soil_Type_5
    df_copy['Mark'] = ''

    # 計算層數、厚度和 Ic 平均值
    layers, thicknesses, ic_avgs = data_array(Soil_Type_5, df_copy['Ic'])
    result_df = pd.DataFrame({'Soil Type': layers, 'Thickness': thicknesses, 'Ic_avg': ic_avgs})

    # 顯示合併層窗口並取得用戶輸入的厚度
    threshold = merge_layer_window(result_df)

    # 合併層並生成新的數據
    if threshold is not None:
        result_array = merge_layer(result_df, threshold)
        data_input = write_merged_data(result_array)
        df_copy['5cm'] = data_input

        # 標記差異
        mark_array = mark(Soil_Type_5, data_input)
        if len(mark_array) == len(df_copy):  # 檢查長度是否一致
            df_copy['Mark'] = mark_array
        else:
            print("標記失敗：長度不一致")

        # 將處理後的資料存入新的 Excel 檔案
        df_copy.to_excel('output.xlsx', index=False)
        print('資料處理完成')
    else:
        print("未輸入有效的合併厚度")

if __name__ == "__main__":
    main()
