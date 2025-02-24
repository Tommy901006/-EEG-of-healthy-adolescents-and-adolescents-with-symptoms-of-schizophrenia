import os
import glob
import pandas as pd
import numpy as np
from scipy.stats import kurtosis

# 固定通道順序
CHANNEL_ORDER = ["F7", "F3", "F4", "F8",
                 "T3", "C3", "Cz", "C4",
                 "T4", "T5", "P3", "Pz",
                 "P4", "T6", "O1", "O2"]

def compute_metrics_per_channel(df):
    """
    針對讀入的 CSV 資料（DataFrame），計算每個通道的標準差、RMS 與峰度 (Fisher)。
    回傳一個字典，鍵為通道名稱，值為該通道各指標的字典。
    """
    metrics = {}
    for channel in CHANNEL_ORDER:
        if channel in df.columns:
            data = df[channel]
            std_val = data.std()
            rms_val = np.sqrt(np.mean(data**2))
            kurt_val = kurtosis(data, fisher=True)
            metrics[channel] = {"標準差": std_val, "RMS": rms_val, "峰度 (Fisher)": kurt_val}
        else:
            metrics[channel] = {"標準差": np.nan, "RMS": np.nan, "峰度 (Fisher)": np.nan}
    return metrics

def process_all_files(input_folder, output_excel):
    """
    遍歷 input_folder 中所有 CSV 檔案，
    針對每個檔案計算各通道的 EEG 時域參數，
    並將相同通道的結果匯整到同一工作表中。
    
    每個工作表的欄位為：檔名、標準差、RMS、峰度 (Fisher)
    工作表名稱即為各通道：F7, F3, F4, …, O2。
    """
    csv_files = glob.glob(os.path.join(input_folder, '*.csv'))
    if not csv_files:
        print(f"在 {input_folder} 中找不到任何 CSV 檔案")
        return

    # 建立一個字典，key 為通道，value 為記錄列表
    channel_records = {channel: [] for channel in CHANNEL_ORDER}

    for csv_file in csv_files:
        try:
            df = pd.read_csv(csv_file)
            metrics = compute_metrics_per_channel(df)
            file_name = os.path.basename(csv_file)
            for channel in CHANNEL_ORDER:
                record = {"檔名": file_name,
                          "標準差": metrics[channel]["標準差"],
                          "RMS": metrics[channel]["RMS"],
                          "峰度 (Fisher)": metrics[channel]["峰度 (Fisher)"]}
                channel_records[channel].append(record)
            print(f"處理完成：{csv_file}")
        except Exception as e:
            print(f"處理 {csv_file} 時發生錯誤：{e}")

    # 將每個通道的資料輸出到 Excel 的各個工作表中
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        for channel in CHANNEL_ORDER:
            df_channel = pd.DataFrame(channel_records[channel])
            # 調整欄位順序
            df_channel = df_channel[["檔名", "標準差", "RMS", "峰度 (Fisher)"]]
            df_channel.to_excel(writer, sheet_name=channel, index=False)
    print(f"所有檔案的結果已儲存至 {output_excel}")

def main():
    # 修改以下路徑
    input_folder = 't1'  # 請設定存放 CSV 檔案的資料夾
    output_excel = 'all_metrics.xlsx'   # 輸出的 Excel 檔案名稱
    process_all_files(input_folder, output_excel)

if __name__ == '__main__':
    main()
