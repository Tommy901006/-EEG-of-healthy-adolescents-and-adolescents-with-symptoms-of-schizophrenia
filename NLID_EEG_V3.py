# -*- coding: utf-8 -*-
"""
Created on Sat Nov 23 13:00:00 2024

@author: User
"""

import os
import numpy as np
import pandas as pd
from scipy.signal import butter, lfilter
from NLIDOOP3 import RecurrenceAnalysis

# 頻帶範圍設定
BAND_FREQS = {
    "Delta": (0.5, 4),
    "Theta": (4, 8),
    "Alpha": (8, 12),
    "Beta": (12, 30),
    "ALL": (0.5, 30),
}

# 頭皮通道選項
CHANNELS = ['F7', 'F3', 'F4', 'F8', 
            'T3', 'C3', 'Cz', 'C4', 'T4', 'T5', 'P3', 'Pz', 'P4', 'T6', 'O1', 'O2']

# 讓使用者自行設定參考通道與欲處理的目標通道
reference_channel = input(f"請輸入參考通道 (選項: {CHANNELS}): ").strip()
if reference_channel not in CHANNELS:
    raise ValueError("參考通道輸入錯誤。")

target_channel = input(f"請輸入欲處理的通道 (選項: {CHANNELS}，但不可與參考通道相同): ").strip()
if target_channel not in CHANNELS:
    raise ValueError("處理通道輸入錯誤。")
if target_channel == reference_channel:
    raise ValueError("目標通道與參考通道不能相同。")

# 將目標通道設定為單一通道列表
target_channels = [target_channel]

# 參數設定
sampling_rate = 128  # EEG 訊號取樣率 (Hz)
window_duration = 60  # 窗口持續時間 (秒)
overlap = 0.2        # 窗口重疊比例
window_size = int(sampling_rate * window_duration)  # 每個窗口的資料點數
step_size = int(window_size * (1 - overlap))        # 窗口移動的步長

# 資料夾路徑（請根據您的環境修改）
input_folder = r"C:\Users\User\Desktop\4\test\t23\t1"
output_folder = r"C:\Users\User\Desktop\4\test\t23\t1"
os.makedirs(output_folder, exist_ok=True)

# 濾波函數
def butter_bandpass(lowcut, highcut, fs, order=4):
    nyquist = 0.5 * fs
    low = lowcut / nyquist
    high = highcut / nyquist
    b, a = butter(order, [low, high], btype='band')
    return b, a

def butter_bandpass_filter(data, lowcut, highcut, fs, order=4):
    b, a = butter_bandpass(lowcut, highcut, fs, order=order)
    return lfilter(b, a, data)

def process_file(file_path, target_channels, reference_channel):
    """
    處理單一 CSV 檔案：
      - 對於每個目標通道 (target_channels) 以及每個頻帶（Delta、Theta、Alpha、Beta），
        利用參考通道 (reference_channel) 與該目標通道的信號計算 NLID。
      - 針對每個頻帶以滑動窗口計算 NLID_XY 與 NLID_YX，並取平均作為該頻帶的結果。
    回傳一個列表，每筆記錄包含：
      'Filename', 'Channel', 'Band', 'NLID_XY_Avg', 'NLID_YX_Avg'
    """
    file_name = os.path.basename(file_path)
    try:
        df = pd.read_csv(file_path)
    except Exception as e:
        print(f"無法讀取檔案 {file_name}: {e}")
        return []
    
    if reference_channel not in df.columns:
        print(f"檔案 {file_name} 缺少參考通道 {reference_channel}，跳過。")
        return []
    
    results = []
    for channel in target_channels:
        if channel not in df.columns:
            print(f"檔案 {file_name} 缺少通道 {channel}，跳過此通道。")
            continue
        x = df[channel].values
        y = df[reference_channel].values
        for band, (low, high) in BAND_FREQS.items():
            x_filtered = butter_bandpass_filter(x, low, high, sampling_rate)
            y_filtered = butter_bandpass_filter(y, low, high, sampling_rate)
            
            num_windows = (len(x_filtered) - window_size) // step_size + 1
            nlid_xy_list = []
            nlid_yx_list = []
            for i in range(num_windows):
                start_idx = i * step_size
                end_idx = start_idx + window_size
                x_window = x_filtered[start_idx:end_idx]
                y_window = y_filtered[start_idx:end_idx]
                
                ra_x = RecurrenceAnalysis(x_window, m=3, tau=1)
                phase_space_x = ra_x.reconstruct_phase_space()
                
                ra_y = RecurrenceAnalysis(y_window, m=3, tau=1)
                phase_space_y = ra_y.reconstruct_phase_space()
                
                AR_HR_BW = RecurrenceAnalysis.compute_reconstruction_matrix(
                    phase_space_x, threshold=0.1, threshold_type="dynamic"
                )
                AR_RP_BW = RecurrenceAnalysis.compute_reconstruction_matrix(
                    phase_space_y, threshold=0.1, threshold_type="dynamic"
                )
                
                NLID_XY_avg, NLID_YX_avg = RecurrenceAnalysis.calculate_nlid(AR_HR_BW, AR_RP_BW)
                nlid_xy_list.append(NLID_XY_avg)
                nlid_yx_list.append(NLID_YX_avg)
            avg_nlid_xy = np.mean(nlid_xy_list) if nlid_xy_list else np.nan
            avg_nlid_yx = np.mean(nlid_yx_list) if nlid_yx_list else np.nan
            
            results.append({
                'Filename': file_name,
                'Band': band,
                'NLID_XY_Avg': avg_nlid_xy,
                'NLID_YX_Avg': avg_nlid_yx
            })
    return results

# 處理所有檔案並彙整結果
all_results = []
for file_name in os.listdir(input_folder):
    if file_name.endswith(".csv"):
        file_path = os.path.join(input_folder, file_name)
        print(f"處理檔案: {file_name}")
        file_results = process_file(file_path, target_channels, reference_channel)
        all_results.extend(file_results)

summary_df = pd.DataFrame(all_results)

# 輸出到 Excel：以頻帶 (Delta, Theta, Alpha, Beta) 為工作表名稱
output_excel = os.path.join(output_folder, 'NLID_Summary_By_Band.xlsx')
with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
    for band in BAND_FREQS.keys():
        band_df = summary_df[summary_df['Band'] == band]
        # 調整欄位順序：檔名, Channel, NLID_XY_Avg, NLID_YX_Avg
        band_df = band_df[['Filename',  'NLID_XY_Avg', 'NLID_YX_Avg']]
        band_df.to_excel(writer, sheet_name=band, index=False)

print(f"所有檔案的 NLID 結果已彙整完成，總結儲存至 {output_excel}")
