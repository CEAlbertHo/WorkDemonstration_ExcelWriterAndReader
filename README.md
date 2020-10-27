# ExcelWriterAndReader ( Excel 讀檔 & 寫黨工具 )

## -簡介-
這是一個能夠逐行讀取支援的 Excel 格式工作表，搭配指定的關鍵字，來做到把 Excel 資料序列化成二進制格式 ( binary format ) 的檔案的工具。

---

## 主要功能介紹
- 批次轉檔 : 一次進行複數檔案的轉換 ( 保留相對路徑 )
- 檔案轉換 : 把 Excel 寫出成為 Binary 格式
- 簡易加密 : 寫出 Steam 時可以用指定的 Key 和 IV 作加密 ( 以及之後解密用 )
- 簡易紀錄 : 跑完後會有簡易的結果資訊，記錄那些檔案成功，那些失敗，例外的檔案會有例外資訊，另外會有檔案轉換的時間記錄。

---

## 使用介紹

主畫面如圖所示。

![主視窗](https://i.ibb.co/2s1rDWX/Excel-Converter.png)

目前有兩個功能

1. 產生資料夾功能 ( Creative Directory 按鈕 )   
   1-1. 如果執行檔所處的路徑中沒有 "來源資料夾" 及 "目標資料夾" 的話，會建立這兩個資料夾。
   
   ![建立資料夾](https://i.ibb.co/BZQfRfT/Excel-Converter-Create-Directory.png)

2. Excel 檔案轉換功能 ( Convert Excel 按鈕 )   
   2-1. 把符合指定格式的 Excel 檔案放到來源資料夾裡面後，直接點轉換按鈕就會開始執行。

   2-2. 點了按鈕進行轉換後，會在程式跑完後看到轉換結果。
   ![批次檔案轉換 1](https://i.ibb.co/7KMxmjh/Excel-Converter-Convert-Excel-Result.png)

   2-3. 可以在 "目標資料夾" 找到轉換完成的檔案以及詳細的轉換紀錄，轉換完成的檔案也會維持原本的資料夾結構。
   ![批次檔案轉換 2](https://i.ibb.co/8MdnhGX/vs.png)

   ![轉換紀錄 Log](https://i.ibb.co/8d7pgdM/Log.png)

   2-4. 附上原始 Excel 檔案以及轉換後的結果。
   ![Excel 原始內容](https://i.ibb.co/FVvzXYR/Excel.png)
   ![Binary 加密](https://i.ibb.co/DQd0nPz/Encrypted-Binary.png)


檔案轉換完之後，可以在其他地方用對應的 Key 和 IV 來解開 Binary 資料，並可以還原成想儲存的資料結構格式。

 - 解開 Binary 檔案範例
 ![Binary 解碼](https://i.ibb.co/GdxMBRR/Decrypted-Binary.png)











// Copyright ©2020 Albert Ho ( rt135792005@gmail.com ). All rights reserved.