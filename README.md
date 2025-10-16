# 工程施工照片管理系統 (Construction Photo Management System)

## 📋 專案簡介

這是一個基於 **Excel VBA** 開發的工程施工照片管理系統，專門用於工程查驗照片的整理、重新命名、報表產出與列印管理。系統能夠自動化處理大量施工照片，並產生符合工程需求的查驗報表。

## 🏗️ 專案架構

### 核心模組 (Modules)

#### 1. **cmdMain.bas** - 主要命令模組
主要功能入口點，包含以下核心功能：
- `PrintOut()` - 列印 PDF 報表
- `getDataFromFolder()` - 從資料夾讀取照片資料
- `SelectFolder()` - 選擇資料夾對話框
- `ChangeAllFileName()` - 批次更改檔名
- `combineWorkbooks()` - 合併多個工作簿
- `showFolder()` - 開啟目標資料夾

#### 2. **test.bas** - 測試與輔助功能
包含資料提取、篩選與命名相關功能

#### 3. **PastePic.bas** - 照片貼上功能
- 自動將資料夾內的照片貼到 Excel 工作表
- 支援照片自動縮放與排版
- 處理 JPG/JPEG 格式照片

#### 4. **FetchURL.bas** & **GIT.bas**
網路資源獲取與版本控制相關功能

---

### 類別模組 (Class Modules)

#### 1. **clsMyfunction.cls** - 核心函數類別 (最大的類別，約 918 行)
包含大量通用函數：
- **陣列排序功能**：
  - `BubbleSort_array()` - 冒泡排序（適用於小量資料）
  - `MergeSort_array()` - 合併排序（適用於大量資料）
  - `SortPTArray()` - 點陣列排序
  
- **資料轉換功能**：
  - `tranColls2Array()` - 集合轉陣列
  - `tranColl2Array()` - 單一集合轉陣列
  - `tranArray2Coll()` - 陣列轉集合
  - `tranColls2OneArray()` - 多集合合併為一維陣列
  
- **資料查詢與篩選**：
  - `getRowsByUser()` - 根據條件取得列號
  - `getUniqueItems()` - 取得唯一值
  - `getUniqueItemsInCollRows()` - 在指定列集合中取得唯一值
  
- **字串處理**：
  - `tranCharcter()` - 字元轉換
  - `ConvertToLetter()` - 數字轉欄位字母
  - `tranDateToStr()` / `tranStrToDate()` - 日期格式轉換
  
- **工程專用功能**：
  - `TranLoc()` - 樁號轉換
  - `SplitAllLocs()` - 區間分割
  - `IsRecLocPass()` - 樁號區間檢查
  - `splitFileName_Check()` - 檔名分割檢查

- **輔助功能**：
  - `showList()` - 顯示集合或陣列內容
  - `getBlankColl()` - 取得空白列集合
  - `ClearData()` - 清除資料
  - `FileOpen()` / `IsFileExists()` - 檔案操作

#### 2. **clsReport.cls** - 報表產出類別
照片查驗報表的核心類別：
- `CollectItem()` - 收集查驗項目
- `GetReportByItem()` - 依項目產生報表
- `PastePhoto()` - 貼上照片到報表
- `PasteDetail()` - 貼上查驗詳細資料
- `PrintXLS()` - 輸出為 Excel 格式
- `printReport()` - 輸出為 PDF 格式
- `AddText()` - 在照片上添加文字（如日期）
- `tranDate()` - 日期格式轉換
- `ClearAllPhoto()` / `ClearAllDetail()` - 清除舊資料

#### 3. **clsPrintOut.cls** - 列印輸出類別
處理各種列印與輸出需求：
- `ToPDF()` - 輸出為 PDF
- `ToXLS()` - 輸出為 Excel
- `ToPDF_Check()` - 查驗表輸出為 PDF
- `ShtToPDF()` - 工作表轉 PDF
- `combineFiles()` - 合併多個檔案
- `clearMark()` - 清除標記文字
- `SpecificShtToXLS()` - 特定工作表輸出為 Excel

#### 4. **clsmyFile.cls** - 檔案處理類別
處理照片檔案的讀取與命名：
- `getAllFolder()` - 遍歷所有資料夾
- `getAllFile()` - 取得所有檔案
- `PastePictures()` - 貼上照片到工作表
- `changeFileName()` - 更改檔名
- `getRenameFile()` - 產生重新命名規則
- `getFileName()` - 取得檔名
- `getParentFolder()` - 取得上層資料夾
- `delRng()` - 清除舊資料與圖片
- `IsPhoto()` - 判斷是否為照片檔案

#### 5. **clsInformation.cls** - 設定資訊類別
管理系統設定與參數：
- `IsShowEditForm()` - 是否顯示編輯表單
- `IsPrintDate()` - 是否列印日期
- `IsPrintDateBack()` - 是否顯示日期背景
- `IsPasteIMG()` - 是否貼上圖片縮圖
- `getIMGwidth()` - 取得圖片寬度
- `getMainPath()` - 取得主要路徑
- `getFontColor()` / `getInteriorColor()` - 取得顏色設定
- `getReNameStruc()` - 取得重新命名結構
- `getCollStructure()` - 解析結構字串為欄位集合
- `VBLongToRGB()` - 顏色值轉換

#### 6. **clsFetchURL.cls** - URL 擷取類別
處理網路資源獲取

#### 7. **clsFSO.cls** - 檔案系統物件類別
檔案系統操作的封裝

---

## 🔄 系統運作流程

### 1️⃣ **照片匯入流程**
```
選擇資料夾 → 掃描所有照片檔案 → 讀取檔案資訊 → 貼上縮圖（可選）→ 顯示於 Result 工作表
```

### 2️⃣ **檔名重新命名流程**
```
設定命名規則 → 解析檔案資訊 → 產生新檔名 → 批次重新命名
```
**命名格式範例**：`渠道名稱_測點_日期_查驗項目_備註.jpg`

### 3️⃣ **報表產出流程**
```
收集查驗項目 → 依項目/日期排序 → 產生報表頁面 → 貼上照片與詳細資料 → 輸出 PDF/Excel
```

### 4️⃣ **工作簿合併流程**
```
選擇多個 Excel 檔案 → 依序複製工作表 → 合併到新活頁簿 → 儲存結果
```

---

## 📂 專案檔案結構

```
constructionPhoto/
├── README.md                    # 專案說明文件（本檔案）
├── .git/                        # Git 版本控制資料夾
│
├── 【主要模組 Modules】
├── cmdMain.bas                  # 主命令模組（核心入口）
├── test.bas                     # 測試與輔助功能
├── PastePic.bas                 # 照片貼上功能
├── FetchURL.bas                 # URL 擷取功能
├── GIT.bas                      # Git 相關功能
├── douliu.bas                   # 特定專案功能
├── UnitTest.bas                 # 單元測試
├── useFunction.bas              # 輔助函數
├── Module1.bas                  # 通用模組
│
├── 【類別模組 Classes】
├── clsMyfunction.cls            # 核心函數類別（最大，918行）
├── clsReport.cls                # 報表產出類別
├── clsPrintOut.cls              # 列印輸出類別
├── clsmyFile.cls                # 檔案處理類別
├── clsInformation.cls           # 設定資訊類別
├── clsFetchURL.cls              # URL 擷取類別
├── clsFSO.cls                   # 檔案系統類別
│
├── 【使用者表單 UserForms】
├── ImageTmp.frm                 # 圖片暫存表單
├── ImageTmp.frx                 # 圖片暫存表單資源
├── Inf.frm                      # 資訊設定表單
├── Inf.frx                      # 資訊設定表單資源
├── UserForm2.frm                # 通用表單
├── UserForm2.frx                # 通用表單資源
│
└── 【工作表模組 Worksheets】
    ├── ThisWorkbook.doccls      # 活頁簿事件
    ├── 工作表1.doccls           # 主要工作表（Main）
    ├── 工作表2.doccls           # 輔助工作表（Result）
    ├── 工作表4.doccls           # 輔助工作表
    └── 工作表5.doccls           # 輔助工作表
```

---

## 🎯 主要功能特色

### ✅ 批次照片管理
- 自動掃描資料夾及子資料夾中的所有照片
- 支援照片縮圖預覽
- 智慧型照片尺寸調整（保持長寬比）

### ✅ 智慧檔名管理
- 自訂命名規則結構
- 批次重新命名檔案
- 檔名規則驗證

### ✅ 專業報表產出
- 支援多種排序方式（查驗項目、查驗日期）
- 自動分頁（每頁 3-4 張照片）
- 可輸出 PDF 或 Excel 格式
- 照片上可標註日期與相關資訊

### ✅ 工程專用功能
- 樁號轉換與驗證
- 區間範圍檢查
- 查驗項目管理

### ✅ 資料處理工具
- 多種排序演算法（冒泡排序、合併排序）
- 集合與陣列轉換
- 唯一值篩選
- 日期格式轉換

---

## 🚀 使用方式

### 基本操作流程

1. **開啟 Excel 檔案**
   - 啟用巨集功能（信任此文件）

2. **設定主資料夾**
   - 執行 `SelectFolder()` 選擇照片所在資料夾
   - 系統會將路徑儲存在 `Main` 工作表的 `B2` 儲存格

3. **匯入照片資料**
   - 執行 `getDataFromFolder()`
   - 照片清單會顯示在 `Result` 工作表
   - 可選擇是否顯示縮圖

4. **設定檔名重新命名規則**（可選）
   - 在 `Main` 工作表 `B6` 設定命名結構
   - 例如：`渠道名稱_測點_日期_查驗項目_備註`
   - 執行 `ChangeAllFileName()` 進行重新命名

5. **產生報表**
   - 執行 `PrintOut()`
   - 選擇排序方式（I=查驗時間, J=查驗項目）
   - 選擇輸出格式（1=XLS, 2=PDF）
   - 報表會儲存在專案資料夾的 `查驗照片Output` 子資料夾

---

## ⚙️ 系統設定

系統設定位於 `Main` 工作表：

| 設定項目 | 儲存格 | 說明 |
|---------|--------|------|
| 主資料夾路徑 | B2 | 照片所在的根目錄 |
| 縮圖寬度 | C3 | 縮圖顯示寬度（像素） |
| 報表範本工作表 | B4 | 使用的報表範本名稱 |
| 日期字型顏色 | C5 | 照片上日期文字的顏色 |
| 重新命名結構 | B6 | 檔名規則（用 `_` 分隔欄位） |

### CheckBox 設定
- `CheckBox1` - 是否顯示編輯表單
- `CheckBox2` - 是否在照片上顯示日期
- `CheckBox3` - 日期是否顯示背景色
- `CheckBox4` - 是否貼上圖片縮圖

---

## 📊 資料結構

### Result 工作表欄位
| 欄位 | 說明 |
|-----|------|
| A | ID（序號） |
| B | 縮圖 |
| C | 完整路徑 |
| D | 資料夾 |
| E | 檔名 |
| F | 新檔名 |
| G | 渠道名稱 |
| H | 測點 |
| I | 日期 |
| J | 查驗項目 |
| K | 備註 |

---

## 🔧 技術細節

### 採用的演算法
- **冒泡排序**：用於少量資料（< 70 筆）
- **合併排序**：用於大量資料排序，效能較佳
- **遞迴式資料夾遍歷**：深度優先搜尋所有子資料夾

### 外部相依性
- **Scripting.FileSystemObject**：檔案系統操作
- **Office.FileDialog**：檔案選擇對話框
- **Excel 物件模型**：工作表、範圍、圖片等操作

### 檔案格式支援
- **照片格式**：JPG, JPEG
- **輸出格式**：PDF（需 Excel 2007 以上）、XLS

---

## 📝 注意事項

1. **Excel 版本需求**
   - 建議使用 Excel 2010 以上版本
   - PDF 輸出功能需要 Excel 2007 SP2 以上

2. **巨集安全性設定**
   - 需要啟用巨集才能使用本系統
   - 建議將檔案另存為 `.xlsm` 格式

3. **效能考量**
   - 大量照片（> 100 張）建議不要顯示縮圖
   - 照片尺寸建議控制在 5MB 以下

4. **資料夾路徑**
   - 避免使用特殊字元
   - 路徑不宜過長（建議 < 200 字元）

5. **檔名規則**
   - 檔名不可包含：`\ / : * ? " < > |`
   - 建議使用底線 `_` 作為欄位分隔符號

---

## 🐛 常見問題

### Q1: 執行時出現「無法另存為 PDF」的錯誤
**A:** Excel 版本過舊，請升級至 Excel 2007 SP2 以上版本，或改用列印功能。

### Q2: 照片無法正確顯示
**A:** 請確認照片格式為 JPG 或 JPEG，且檔案未損毀。

### Q3: 重新命名後檔案找不到
**A:** 檢查新檔名是否包含非法字元，或路徑是否過長。

### Q4: 報表產生速度很慢
**A:** 建議：
   - 減少每次處理的照片數量
   - 不要同時開啟太多 Excel 檔案
   - 關閉不必要的應用程式

---

## 🔄 版本控制

本專案使用 **Git** 進行版本控制，`.git` 資料夾包含完整的版本歷史記錄。

---

## 👨‍💻 開發者資訊

- **開發語言**：VBA (Visual Basic for Applications)
- **開發環境**：Microsoft Excel
- **專案類型**：工程管理工具

---

## 📄 授權說明

本專案為內部使用工具，請勿未經授權擅自散布或商業使用。
---

## 📧 聯絡資訊

如有問題或建議，請聯繫專案維護人員。

---

**最後更新**：2025-10-16  
**文件版本**：1.0
