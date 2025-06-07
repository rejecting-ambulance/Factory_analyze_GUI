import cv2
import numpy as np
import fitz  # PyMuPDF
import os
import sys
import shutil
import json
import pytesseract
import time
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from concurrent.futures import ProcessPoolExecutor
import multiprocessing

from factory_to_sheet_mc import process_folder_multiprocessing
from factory_query import process_excel_data
"""
這段程式碼會讀取1個PDF
並依指定的特徵分割成不同檔案
目前設定的特徵圖片為市長官印
mayor_stamp
"""

def select_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        print("選擇的 PDF:", file_path)
        return file_path    


def load_config(config_path="config.json", default_config=None):
    if not os.path.exists(config_path):
        print(f"[錯誤] 找不到設定檔：{config_path}")
        if default_config:
            print("[提示] 使用預設設定。")
            config = default_config
        else:
            print("[中止] 無法繼續執行，請確認設定檔存在。")
            sys.exit(1)
    else:
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
        except json.JSONDecodeError as e:
            print(f"[錯誤] 設定檔格式錯誤：{e}")
            if default_config:
                print("[提示] 使用預設設定。")
                config = default_config
            else:
                print("[中止] 請修正 config.json 後再執行。")
                sys.exit(1)

    # 處理 tesseract 路徑
    tesseract_path = config.get("tesseract_path", "").strip()
    if tesseract_path:
        if os.path.exists(tesseract_path):
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
            os.environ['TESSERACT_PATH'] = pytesseract.pytesseract.tesseract_cmd
            print(f"[設定] 使用指定的 Tesseract 路徑：{tesseract_path}")
        else:
            print(f"[錯誤] 指定的 Tesseract 路徑不存在：{tesseract_path}，將使用預設 PATH。")
    else:
        print("[設定] 未指定 Tesseract 路徑，將使用預設 PATH。")

    return config

def is_blank_page(page, threshold=0.95):
    """檢查單頁是否為空白"""
    text = page.get_text("text").strip()
    if text:  # 若頁面包含文本，則不空白
        return False
    
    pix = page.get_pixmap()  # 渲染為圖像
    img = np.frombuffer(pix.samples, dtype=np.uint8)  # 轉換為 NumPy 陣列
    white_pixels = np.sum(img == 255)  # 計算純白像素數
    total_pixels = pix.width * pix.height * pix.n  # 總像素數

    return (white_pixels / total_pixels) >= threshold  # 若白色比例超過閾值，則視為空白


def remove_blank_pages(pdf_path, config):
    """移除空白頁"""
    print(str_line('1.移除空白頁面'))
    doc = fitz.open(pdf_path)
    new_doc = fitz.open()

    removed_pages = []  # 記錄被移除的頁碼
    total_pages = len(doc)  # 原始總頁數

    for page in doc:  
        if is_blank_page(page, threshold = config["blank_page_threshold"]):  
            removed_pages.append(page.number + 1)  # PyMuPDF 頁碼從 0 開始，+1 轉為人類可讀的頁碼
        else:
            new_doc.insert_pdf(doc, from_page=page.number, to_page=page.number)

    temp_path = config["cleaned_pdf"]
    new_doc.save(temp_path)
    new_doc.close()
    doc.close()

    # 顯示移除頁面資訊
    remaining_pages = total_pages - len(removed_pages)
    print("移除空白頁面:")
    for i in range(0, len(removed_pages), 5):
        print("  " + ", ".join(map(str, removed_pages[i:i+5])))
    print(f"剩餘頁數: {remaining_pages} / {total_pages}")




def compare_images_sift(img1, img2, threshold=10):
    """
    使用 SIFT 特徵點比對兩張圖片是否相似。

    Args:
        img1: 第一張圖片的物件。
        img2: 第二張圖片的物件。
        threshold: 相似度閾值，預設為 10。

    Returns:
        如果兩張圖片相似，則返回 True，否則返回 False。
    """

    # 將圖片轉換為灰度圖像 (如果需要)
    if len(img1.shape) == 3:
        img1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
    if len(img2.shape) == 3:
        img2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)

    sift = cv2.SIFT_create()
    kp1, des1 = sift.detectAndCompute(img1, None)
    kp2, des2 = sift.detectAndCompute(img2, None)

    if des1 is None or des2 is None:  # 檢查是否有檢測到特徵點
        return False

    bf = cv2.BFMatcher()
    matches = bf.knnMatch(des1, des2, k=2)

    good_matches = []
    for m, n in matches:
        if m.distance < 0.75 * n.distance:
            good_matches.append(m)

    return len(good_matches) > threshold  # 直接返回比較結果


def compare_image_with_pdf_page(image_paths, pdf_path, page_num, threshold=10):
    """比較多張圖片與單一 PDF 頁面是否相似。"""
    imgs = []
    for image_path in image_paths:
        img = cv2.imread(image_path)
        if img is None:
            print(f"Error: Could not read image at {image_path}")
            return None
        imgs.append(img)

    doc = fitz.open(pdf_path)
    page = doc[page_num]
    pix = page.get_pixmap()
    img2 = cv2.imdecode(np.frombuffer(pix.tobytes(), np.uint8), cv2.IMREAD_COLOR)

    similar_image_indices = []
    for i, img1 in enumerate(imgs):
        if compare_images_sift(img1, img2, threshold):
            similar_image_indices.append(i)

    if similar_image_indices:
        return page_num, similar_image_indices  # 返回頁碼和相似圖片索引
    return None

def compare_image_with_pdf_pages_multiprocessing(image_paths, config):
    """使用多核心比較多張圖片與 PDF 每一頁是否相似。"""
    print(str_line('2.比對檔案分割點'))
    max_processes = config["max_processes"]
    if max_processes is None:
        max_processes = max(1, multiprocessing.cpu_count() - 1)

    threshold = config["sift_threshold"]

    pdf_path = config["cleaned_pdf"]
    results = {}
    doc = fitz.open(pdf_path)

    with ProcessPoolExecutor(max_workers=max_processes) as executor:
        tasks = [executor.submit(compare_image_with_pdf_page, image_paths, pdf_path, page_num, threshold) for page_num in range(doc.page_count)]
        for future in tasks:
            result = future.result()
            if result:
                page_num, image_indices = result
                results[page_num] = image_indices

    doc.close()
    return results

def split_pdf(pdf_path, split_points, output_dir="split_pdf"):
    """
    將 PDF 分割成多個檔案。

    Args:
        pdf_path: PDF 路徑。
        split_points: 分割點列表，包含頁碼。
        output_dir: 輸出目錄，預設為 "split_pdf"。
    """
    print(str_line('3.分割檔案'))
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)  # 建立輸出目錄

    try:
        doc = fitz.open(pdf_path)
        start_page = 0

        for i, split_point in enumerate(split_points):
            new_doc = fitz.open()  # 建立新的 PDF 文件
            for page_num in range(start_page, split_point + 1):
                new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)  # 複製頁面

            output_path = os.path.join(output_dir, f"split_{i + 1}.pdf")
            new_doc.save(output_path)  # 儲存分割後的 PDF
            new_doc.close()

            start_page = split_point + 1

        # 處理最後一個部分
        if start_page < doc.page_count:
            new_doc = fitz.open()
            for page_num in range(start_page, doc.page_count):
                new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)

            output_path = os.path.join(output_dir, f"split_{len(split_points) + 1}.pdf")
            new_doc.save(output_path)
            new_doc.close()

    except Exception as e:
        print(f"分割時遇到錯誤: {e}")

    finally:
        doc.close()  # 確保關閉 PDF 文件

def get_images_from_folder(folder_path, extensions=('.jpg', '.jpeg', '.png', '.bmp')):
    image_paths = []
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(extensions):
            image_paths.append(os.path.join(folder_path, filename))
    return image_paths   

def str_line(show = str):
    max_len = 50
    dash_len = int((max_len - len(show))/2)
    dash = ''
    for i in range(dash_len):
        dash = dash + '='
    show = f'{dash} {show} {dash}\n'
    
    return show

def wait_for_file(path, timeout=10):
    start = time.time()
    while not os.path.exists(path):
        if time.time() - start > timeout:
            raise TimeoutError(f"檔案產生超時: {path}")
        time.sleep(0.5)

def check_and_handle_split_folder(config):
    folder_path = config['process_folder']
    if not os.path.exists(folder_path):
        return 'y'

    contents = os.listdir(folder_path)
    if not contents:
        return 'y'

    print(f"⚠️ 偵測到資料夾 '{folder_path}' 有檔案，共 {len(contents)} 個項目。")

    # 啟動 Tkinter 並隱藏主視窗
    root = tk.Tk()
    root.withdraw()

    result = messagebox.askyesno(
        "清空資料夾",
        f"偵測到資料夾 '{folder_path}' 內已有 {len(contents)} 個項目。\n\n是否要清空？\n\n"
        "是：清空後自選PDF\n否：分析資料夾內容"
    )

    root.destroy()

    if result:
        for item in contents:
            item_path = os.path.join(folder_path, item)
            if os.path.isdir(item_path):
                shutil.rmtree(item_path)
            else:
                os.remove(item_path)
        print(f"✅ 已清空資料夾 '{folder_path}'。")
        return 'y'
    else:
        print("✅ 保留原資料夾內容。")
        return 'n'

def print_intro():
    print("""
    ================ 工廠登記公文自動化處理系統 ================
    - 開啟 - 找印章 - 分割 - 找號碼 - 擷取 - 查詢 - 下班 - 
    ===========================================================
    """)


def show_manual_step(root, config):
    output_excel = config['output_excel']
    def on_yes():
        root.destroy()  # 關閉目前提示視窗
        print(str_line('5.查詢工廠編號'))
        process_excel_data(output_excel, 4)
        show_finish_window()

    def on_no():
        root.destroy()
        print("不查詢。")
        show_finish_window()

    root.title("步驟五：手動檢查工廠編號")
    label = tk.Label(root, text="請到 Excel 設定工廠編號，完成後儲存並關閉。", font=('Arial', 12))
    label.pack(padx=20, pady=20)

    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    btn_yes = tk.Button(button_frame, text="已完成 (y)", command=on_yes, width=15)
    btn_yes.grid(row=0, column=0, padx=10)

    btn_no = tk.Button(button_frame, text="尚未完成 (n)", command=on_no, width=15)
    btn_no.grid(row=0, column=1, padx=10)

    root.mainloop()

def show_finish_window():
    finish_root = tk.Tk()
    finish_root.title("完成")
    label = tk.Label(finish_root, text=str_line("有了耳朵，歡樂多更多！完成囉"), font=('Arial', 12))
    label.pack(padx=20, pady=20)

    btn_close = tk.Button(finish_root, text="結束程式", command=finish_root.destroy, width=20)
    btn_close.pack(pady=10)

    finish_root.mainloop()



def main ():
    print_intro()
    

    default_config = {
        "blank_page_threshold" : 0.95,
        "sift_threshold" : 20,
        "tesseract_path": "",
        "clean_pdf": "remove_blank.pdf",
        "process_folder": "split_pdf",
        "image_folder": "footer_images",
        "output_excel": "factory_extraction.xlsx",
        "max_processes": multiprocessing.cpu_count() - 1,
        "clean_temp_pdf": "True"
    }

    config = load_config(default_config=default_config)

    if_split = check_and_handle_split_folder(folder_path=config['process_folder'])

    if if_split == 'y':
        pdf_path = select_pdf()

        remove_blank_pages(pdf_path, config)
        wait_for_file(config["cleaned_pdf"])
        temp_path = config["cleaned_pdf"]

        image_folder = config['image_folder']  # 假設你的圖片都放在 templates 資料夾內
        image_paths = get_images_from_folder(image_folder)

        results = compare_image_with_pdf_pages_multiprocessing(image_paths, config)  # 使用多核心版本

        if results:
            similar_pages = []
            for page_num, image_indices in results.items():
                similar_pages.append(page_num)

            if similar_pages:
                split_points = []

                for i in range(len(similar_pages) - 1):
                    if similar_pages[i + 1] - similar_pages[i] >= 1:
                        split_points.append(similar_pages[i])

                if similar_pages[-1] not in split_points:
                    split_points.append(similar_pages[-1])

                split_file_count = len(split_points) if split_points else 1

                print("分割結果：")
                print("-" * 50)
                print(f"{'頁面':<10}{'相似圖片數':<15}{'已分割檔案'}")
                print("-" * 50)

                for i, page_num in enumerate(similar_pages):
                    image_indices = results[page_num]
                    print(f"Page {page_num + 1:<6}→ {len(image_indices):<13}張     {i + 1}/{split_file_count}")

                split_pdf(temp_path, split_points)
                print("PDF 分割完成！")
            else:
                print("PDF 中沒有與圖片相似的頁面。")
        else:
            print("Error during PDF processing. No similar pages found.")
        
        # 清理暫存檔案 (可選)
        if config['clean_temp_pdf']:
            os.remove(config["cleaned_pdf"])


    print(str_line('4.擷取文件內工廠編號'))
    process_folder_multiprocessing(config)
    
    show_manual_step(tk.Tk(), config)

# 範例
if __name__ == '__main__':
    multiprocessing.freeze_support()
    main()