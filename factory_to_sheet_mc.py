from pdf2image import convert_from_path
import fitz
import pandas as pd
import pytesseract
import json
import os
import re
from concurrent.futures import ProcessPoolExecutor
import multiprocessing
from functools import partial


def load_config(config_path="config.json"):
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def ensure_tesseract_path(config):
    tesseract_path = config.get("tesseract_path", "").strip()
    if tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path


def preprocess_image(image):
    gray_image = image.convert('L')
    binary_image = gray_image.point(lambda x: 0 if x < 175 else 255, '1')
    return binary_image


def pdf_to_text(pdf_path, config):
    doc = fitz.open(pdf_path)
    full_text = ""

    settings = config.get("tesseract_config", "")
    lang = config.get("tesseract_lang", "chi_tra")
    dpi = config.get("dpi", 300)

    for i, page in enumerate(doc):
        text = page.get_text("text")
        if not text.strip():
            images = convert_from_path(pdf_path, dpi=dpi,poppler_path=config["poppler_path"])
            text = pytesseract.image_to_string(images[i], lang=lang, config=settings)
        full_text += f"--- 第 {i + 1} 頁 ---\n{text}\n"
    return full_text


def extract_pdf_data(pdf_path, config):
    text = pdf_to_text(pdf_path, config)
    exclude_path = config.get("exclude_path", "exclude_numbers.txt")

    try:
        with open(exclude_path, "r", encoding="utf-8") as f:
            exclude_set = set(line.strip() for line in f if re.fullmatch(r"\d{8}", line.strip()))
    except FileNotFoundError:
        exclude_set = set()

    document_number_pattern = config["document_number_pattern"]
    factory_number_pattern = config["factory_number_pattern"]

    document_number_match = re.search(document_number_pattern, text)
    document_number = f"府經工行字第{document_number_match.group(1)}號" if document_number_match else "未匹配"

    matches = re.findall(factory_number_pattern, text)
    factory_numbers = [m1 if m1 else m2 for m1, m2 in matches]

    filtered_factory_numbers = [
        num for num in factory_numbers
        if not (re.fullmatch(r"\d{8}", num) and num in exclude_set)
    ]

    unique_factory_numbers = sorted(set(filtered_factory_numbers))
    factory_numbers_result = ", ".join(unique_factory_numbers) if unique_factory_numbers else "無"

    return {
        "發文字號": document_number,
        "工廠編號": factory_numbers_result,
    }


def process_single_pdf(pdf_file, config):
    data = extract_pdf_data(pdf_file, config)
    data["檔名"] = os.path.basename(pdf_file)
    return data


def process_folder_multiprocessing(config):
    ensure_tesseract_path(config)

    folder_path = config['process_folder']
    output_excel = config['output_excel']
    max_processes = config["max_processes"]

    cpu_count = multiprocessing.cpu_count()
    if max_processes is None:
        max_processes = max(1, cpu_count - 1)

    if not os.path.exists(folder_path):
        raise FileNotFoundError(f"❌ 找不到指定資料夾：{folder_path}")

    pdf_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".pdf")]
    extracted_data = []

    with ProcessPoolExecutor(max_workers = max_processes) as executor:
        job = partial(process_single_pdf, config=config)
        futures = [executor.submit(job, pdf_file) for pdf_file in pdf_files]
        for future in futures:
            extracted_data.append(future.result())
            print(f"Processed: {extracted_data[-1]['檔名']}")

    df = pd.DataFrame(extracted_data)
    df["編號"] = df["檔名"].apply(lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else None)
    columns_order = ["編號", "檔名"] + [col for col in df.columns if col not in ["檔名", "編號"]]
    df = df[columns_order]
    df.to_excel(output_excel, index=False)
    print(f"\n✅ 提取結果已保存至：{output_excel}")

    
    # 單核處理，除錯用
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
    



if __name__ == "__main__":
    config = load_config()
    folder_path = config["process_folder"]
    output_excel = config["output_excel"]

    #process_folder_multiprocessing(folder_path, output_excel, config)
    process_folder_multiprocessing(config)
