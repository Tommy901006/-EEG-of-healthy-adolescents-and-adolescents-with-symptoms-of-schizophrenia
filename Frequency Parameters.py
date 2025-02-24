import os
import glob
import pandas as pd
import numpy as np
from scipy.signal import welch

# 設定取樣頻率 (Hz) 與 Welch 參數
fs = 128
nperseg = 256

# 定義各頻段範圍 (Hz)
bands = {
    'Delta (0.5-4 Hz)': (0.5, 4),
    'Theta (4-8 Hz)': (4, 8),
    'Alpha (8-13 Hz)': (8, 13),
    'Beta (13-30 Hz)': (13, 30)
}

# 固定通道順序
CHANNEL_ORDER = ["F7", "F3", "F4", "F8",
                 "T3", "C3", "Cz", "C4",
                 "T4", "T5", "P3", "Pz",
                 "P4", "T6", "O1", "O2"]

def compute_band_power(signal):
    """
    利用 Welch 方法計算 PSD，並依據 bands 定義的頻段計算功率（積分）。
    同時計算各頻段的相對功率（頻段功率除以所有頻段功率總和）。
    回傳一個字典，包含各頻段的功率及相對功率。
    """
    freqs, psd = welch(signal, fs=fs, nperseg=nperseg)
    band_powers = {}
    # 計算各頻段功率
    for band_name, (low, high) in bands.items():
        idx = np.logical_and(freqs >= low, freqs <= high)
        power = np.trapz(psd[idx], freqs[idx])
        band_powers[band_name] = power

    # 計算總功率
    total_power = sum(band_powers.values())
    
    # 計算相對功率，避免除以零
    if total_power > 0:
        band_powers_rel = {f"{band_name} Relative": power / total_power 
                           for band_name, power in band_powers.items()}
    else:
        band_powers_rel = {f"{band_name} Relative": np.nan for band_name in band_powers.keys()}
    
    # 合併結果
    band_powers.update(band_powers_rel)
    return band_powers

def process_file(input_file):
    """
    讀取 CSV 檔案，假設檔案中各欄位為不同 EEG 通道（例如 F7, F3, F4, ...）。
    針對每個通道計算各頻段功率與相對功率，回傳一個列表，
    每一項為一個字典，字典內包含 Channel 與各頻段的數值。
    """
    df = pd.read_csv(input_file)
    results = []
    for channel in df.columns:
        # 只處理在 CHANNEL_ORDER 中定義的通道
        if channel in CHANNEL_ORDER:
            signal = df[channel].values
            band_power = compute_band_power(signal)
            result = {"Channel": channel}
            result.update(band_power)
            results.append(result)
    return results

def process_all_files(input_folder):
    """
    遍歷 input_folder 中所有 CSV 檔案，
    針對每個檔案計算各通道的頻帶功率與相對功率，
    並依照通道彙整結果。
    
    回傳一個字典，其 key 為通道名稱，
    value 為該通道所有檔案的結果列表，每筆記錄包含「檔名」及各頻帶參數。
    """
    csv_files = glob.glob(os.path.join(input_folder, '*.csv'))
    if not csv_files:
        print(f"在 {input_folder} 中找不到任何 CSV 檔案")
        return None

    # 建立一個字典，鍵為各通道
    channel_records = {channel: [] for channel in CHANNEL_ORDER}
    
    for csv_file in csv_files:
        try:
            file_results = process_file(csv_file)
            file_name = os.path.basename(csv_file)
            for record in file_results:
                record["檔名"] = file_name
                channel = record["Channel"]
                channel_records[channel].append(record)
            print(f"處理完成：{csv_file}")
        except Exception as e:
            print(f"處理 {csv_file} 時發生錯誤：{e}")
    
    return channel_records

def main():
    # 請修改以下路徑：資料夾中應包含欲處理的 CSV 檔案
    input_folder = r't1'   # 例如 r'C:\EEGData'
    # 輸出 Excel 檔案名稱
    output_excel = 'frequency_band_power.xlsx'
    
    # 處理所有檔案，依通道彙整結果
    channel_records = process_all_files(input_folder)
    if channel_records is None:
        return

    # 使用 ExcelWriter 將每個通道的結果輸出到不同的工作表中
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        for channel in CHANNEL_ORDER:
            df_channel = pd.DataFrame(channel_records[channel])
            # 調整欄位順序
            desired_columns = [
                "檔名",
                "Delta (0.5-4 Hz)",
                "Delta (0.5-4 Hz) Relative",
                "Theta (4-8 Hz)",
                "Theta (4-8 Hz) Relative",
                "Alpha (8-13 Hz)",
                "Alpha (8-13 Hz) Relative",
                "Beta (13-30 Hz)",
                "Beta (13-30 Hz) Relative"
            ]
            # 若某些欄位不存在，則忽略（以防萬一）
            df_channel = df_channel[[col for col in desired_columns if col in df_channel.columns]]
            df_channel.to_excel(writer, sheet_name=channel, index=False)
    
    print(f"所有檔案的頻帶功率結果已儲存至 {output_excel}")

if __name__ == '__main__':
    main()
