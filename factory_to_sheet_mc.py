from pdf2image import convert_from_path
import fitz
import pandas as pd
import pytesseract
import json
import os
import re
from concurrent.futures import ProcessPoolExecutor
import multiprocessing

'''
    這段程式碼會讀取一個資料夾內所有的PDF
    並依正則表達式(Regular Expressions,REs)
    取得其發文字號、工廠編號
    並造冊成Excel

    修改段落：主函式內
    資料夾名稱、輸出表單名稱.xlsx

'''


def ensure_tesseract_path():
    tesseract_path = os.environ.get('TESSERACT_PATH', '').strip()
    if tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path

# 圖像預處理（可選）
def preprocess_image(image):
    # 將圖片轉為灰度
    gray_image = image.convert('L')
    # 二值化處理（可選，根據需要調整閾值）
    binary_image = gray_image.point(lambda x: 0 if x < 175 else 255, '1')
    return binary_image


# 讀取 PDF 或進行 OCR
def pdf_to_text(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = ""

    settings = ' --oem 1 --psm 6'

    for i, page in enumerate(doc):
        text = page.get_text("text")
        if not text.strip():
            images = convert_from_path(pdf_path, dpi=300)
            text = pytesseract.image_to_string(images[i], lang="chi_tra", config = settings)
        full_text += f"--- 第 {i + 1} 頁 ---\n{text}\n"
    return full_text


def extract_pdf_data(pdf_path, exclude_path="exclude_numbers.txt"):
    text = pdf_to_text(pdf_path)

    # 讀取排除清單
    try:
        with open(exclude_path, "r", encoding="utf-8") as f:
            exclude_set = set(line.strip() for line in f if re.fullmatch(r"\d{8}", line.strip()))
    except FileNotFoundError:
        exclude_set = set()

    dispatch_number_pattern = r"(?<!\d)(\d{10})(?!\d)"
    factory_number_pattern = r"(?<!\d)(\d{8})(?!\d)|(?<!\w)(S\d{7})(?!\d)"

    # 擷取發文字號
    dispatch_number_match = re.search(dispatch_number_pattern, text)
    dispatch_number = f"府經工行字第{dispatch_number_match.group(1)}號" if dispatch_number_match else "未匹配"

    # 擷取工廠編號：findall 會回傳 tuple
    matches = re.findall(factory_number_pattern, text)
    factory_numbers = [m1 if m1 else m2 for m1, m2 in matches]

    # 排除出現在排除清單的 8 碼數字
    filtered_factory_numbers = [
        num for num in factory_numbers
        if not (re.fullmatch(r"\d{8}", num) and num in exclude_set)
    ]

    # 去重、排序並組成字串
    unique_factory_numbers = sorted(set(filtered_factory_numbers))
    factory_numbers_result = ", ".join(unique_factory_numbers) if unique_factory_numbers else "無"

    return {
        "發文字號": dispatch_number,
        "工廠編號": factory_numbers_result,
    }


def process_single_pdf(pdf_file):
    """處理單個 PDF 文件並返回提取的數據。"""
    #ensure_tesseract_path()
    data = extract_pdf_data(pdf_file)
    data["檔名"] = os.path.basename(pdf_file)
    return data

def process_folder_multiprocessing(folder_path, output_excel='extracted_data_factory.xlsx', max_processes=None):
    ensure_tesseract_path()

    cpu_count = multiprocessing.cpu_count()
    if max_processes is None:
        max_processes = max(1, cpu_count - 1)
        
    if not os.path.exists(folder_path):
        raise FileNotFoundError(f"❌ 找不到指定資料夾：{folder_path}")
    
    pdf_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".pdf")]
    extracted_data = []

    with ProcessPoolExecutor(max_workers=max_processes) as executor:
        futures = [executor.submit(process_single_pdf, pdf_file) for pdf_file in pdf_files]
        for future in futures:
            extracted_data.append(future.result())
            print(f"Processed: {extracted_data[-1]['檔名']}")

    df = pd.DataFrame(extracted_data)
    df["編號"] = df["檔名"].apply(lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else None)
    columns_order = ["編號", "檔名"] + [col for col in df.columns if col not in ["檔名", "編號"]]
    df = df[columns_order]
    df.to_excel(output_excel, index=False)
    print(f"\n提取結果已保存至：{output_excel}")

# 定義函數：處理資料夾內的所有 PDF 並輸出到 Excel
def process_folder(folder_path, output_excel = 'extracted_data_factory.xlsx'):
    # 搜索資料夾內的所有 PDF 文件
    pdf_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".pdf")]

    extracted_data = []  # 存放所有提取結果

    for index, pdf_file in enumerate(pdf_files, start=1):
        # 提取單個 PDF 的數據
        data = extract_pdf_data(pdf_file)
        # 加入檔名到提取結果中
        data["檔名"] = os.path.basename(pdf_file)
        extracted_data.append(data)

        # 顯示進度訊息
        print(f"[{index}/{len(pdf_files)}] 已處理：{os.path.basename(pdf_file)}")

    # 將結果保存為 Excel 文件
    df = pd.DataFrame(extracted_data)
    
    # 擷取檔名中的數字
    df["編號"] = df["檔名"].apply(lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else None)
    # 調整欄位順序，把 "數字編號" 放最前面
    columns_order = ["編號", "檔名"] + [col for col in df.columns if col not in ["檔名", "編號"]]
    df = df[columns_order]

    df.to_excel(output_excel, index=False)
    print(f"\n提取結果已保存至：{output_excel}")

# 正則匹配的發文字號和工廠編號
DISPATCH_NUMBER_PATTERN = r"^\d{10}$"  # 發文字號：10 位數字
FACTORY_NUMBER_PATTERN = r"^\d{8}$|^S\d{7}$"  # 工廠編號：8 位數字或 S 開頭 + 7 位數字    
#print(pytesseract.pytesseract.tesseract_cmd)
# 主程序：設置資料夾路徑與輸出 Excel 路徑
if __name__ == "__main__":

    tesseract_cmd = os.environ.get('TESSERACT_PATH')
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    #folder_path = r"D:\users\user\Desktop\image_decode\split_pdf"  # 替換為存放 PDF 的資料夾路徑
    folder_path = "split_pdf"  # 替換為存放 PDF 的資料夾路徑
    output_excel = "extracted_data_factory.xlsx"  # 輸出的 Excel 文件名稱

    process_folder_multiprocessing(folder_path, output_excel)