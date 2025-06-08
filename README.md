# 工廠登記公文自動化處理系統

> 神說要有光，於是就有了光。  
> 你說想提早下班，於是我就出現了。


---


## 前置作業
### 檔案準備
1. 將廠登公文掃描成一份連續的PDF檔案
* 建議先用鉛筆編號，方便確認結果有沒有漏掉。
* 雙面掃描，空白頁也一起掃描。  

### 背景軟體設定
1.安裝Tesseract.exe
* 下載網站 (https://github.com/UB-Mannheim/tesseract/wiki)
* 照步驟安裝，安裝路徑應該都是"C:/Program Files/Tesseract-OCR"
* 要記錄安裝路徑，轉貼到config.json裡面。

2.解壓縮Poppler.zip
* 下載網站 (https://github.com/oschwartz10612/poppler-windows/releases/)
* 紀錄解壓縮的路徑，轉貼到config.json裡面。
* 有先放一份在資料夾內了。

## 操作步驟


1. 點選【工廠登記公文自動化處理系統.exe】
* 會出現黑色小視窗，等待他說第一句話。

2. 跳出選擇視窗，選要分析的PDF
* 如果是第二次分析，split_pdf資料夾有檔案的話，點【否】可以跳過選擇直接分析資料夾。

3.自動偵測及擷取
* 移除空白頁、偵測大印、分割、偵測公文內容、彙整成Excel。

4.自動查工廠編號前，打開Excel，把多截取、沒截取的廠編整理好
* 存檔關閉！
* 存檔關閉！
* 存檔關閉！

5. 回到【處理系統】的詢問視窗，選擇【是】
* 找不到可能是縮小在工具列

6. 自動查詢
* 等查詢完成就可以打開Excel看結果囉。
* 分割完的PDF會保存在split_pdf資料夾內。


## 注意事項

1. footer_images內可以放入想比對的圖片，這是要比對文件結尾的內容。  

2.excleude_numbers.txt可以加入要排除的數字，因為有時候會把電話或是一些不需要的資訊讀進來。

### 
* Tools: ChatGPT 
* Contact: zhandezhong861131@gmail.com
