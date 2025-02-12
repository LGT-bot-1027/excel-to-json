import os
import sys
import subprocess

def ensure_latest_pip():
    subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], check=True)

ensure_latest_pip()

# 確保安裝所需的 Python 模組
required_packages = ["pandas", "openpyxl"]
for package in required_packages:
    try:
        __import__(package)
    except ImportError:
        print(f"正在安裝 {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

import pandas as pd
import json
import re

# 載入 Excel 文件
def load_excel(file_path):
    return pd.read_excel(file_path)

# 篩選需要的列和數據
def filter_data(df):
    filtered_df = df[["Name", "pos_x", "pos_y", "scale_x", "scale_y"]]
    return filtered_df

# 根據 "Name" 提取頁面索引和語言對應
def parse_name(name):
    match = re.match(r"(\d+)_([\u4e00-\u9fa5]+)_(\d+)", name)
    if not match:
        return None, None, None
    return match.group(1), match.group(2), match.group(3)

# 將語言碼轉換為網站語言
def map_language_code(language):
    lang_map = {"英": "EN", "越": "VN", "泰": "TH", "孟": "BD"}
    return lang_map.get(language, None)

def create_json_structure(filtered_df):
    result = {"Pages": []}
    pages = {}

    for _, row in filtered_df.iterrows():
        name, pos_x, pos_y, scale_x, scale_y = row["Name"], row["pos_x"], row["pos_y"], row["scale_x"], row["scale_y"]
        page_num, language, element_num = parse_name(name)
        if not page_num:
            continue

        # 修正浮點數的精度問題
        pos_x = round(pos_x, 3)  # 保留小數點後 3 位
        pos_y = round(pos_y, 3)

        # PageIndex 和 BackgroundName
        if page_num not in pages:
            pages[page_num] = {
                "PageIndex": len(pages),
                "BackgroundName": f"{page_num}_底",
                "Context_CN": [],
            }

        if language == "中":
            pages[page_num]["Context_CN"].append({
                "Name": name,
                "pos_x": pos_x,
                "pos_y": pos_y,
                "scale_x": scale_x,
                "scale_y": scale_y
            })
        else:
            lang_code = map_language_code(language)
            if lang_code:
                context_key = f"Context_{lang_code}"
                if context_key not in pages[page_num]:
                    pages[page_num][context_key] = []
                pages[page_num][context_key].append({
                    "Name": name,
                    "pos_x": pos_x,
                    "pos_y": pos_y,
                    "scale_x": scale_x,
                    "scale_y": scale_y
                })
            else:
                if "Context_外語" not in pages[page_num]:
                    pages[page_num]["Context_外語"] = []
                pages[page_num]["Context_外語"].append({
                    "Name": name,
                    "pos_x": pos_x,
                    "pos_y": pos_y,
                    "scale_x": scale_x,
                    "scale_y": scale_y
                })

    # 處理 "外語" 替補邏輯，並確保排序
    for page_num, page_data in pages.items():
        if "Context_外語" in page_data:
            external_elements = sorted(page_data["Context_外語"], key=lambda x: x["Name"])
            for key in list(page_data.keys()):
                if key.startswith("Context_") and key != "Context_CN":
                    if len(page_data[key]) < len(page_data["Context_CN"]):
                        page_data[key] = sorted(external_elements[:len(page_data["Context_CN"])] + page_data[key], key=lambda x: x["Name"])
            del page_data["Context_外語"]  # 刪除原始的 "Context_外語"

    result["Pages"] = list(pages.values())
    return result

# 儲存為 JSON 文件
def save_to_json(data, output_path):
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# 主程式
if __name__ == "__main__":
    folder_path = os.getcwd()
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]

    for excel_file in excel_files:
        file_path = os.path.join(folder_path, excel_file)
        output_json = os.path.splitext(excel_file)[0] + "Localization" + ".txt"
        
        df = load_excel(file_path)
        filtered_df = filter_data(df)
        json_data = create_json_structure(filtered_df)
        save_to_json(json_data, output_json)
        
        print(f"JSON 文件已成功輸出到 {output_json}")

    input("按下 Enter/回車鍵 繼續")
