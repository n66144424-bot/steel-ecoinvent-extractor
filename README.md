# steel-ecoinvent-extractor

ecoinvent SPOLD 多批次鋼鐵製程技術投入萃取工具

## 📋 專案概述

這個工具可以自動掃描多個 ecoinvent SPOLD 資料庫檔案夾，找出所有鋼鐵相關製程，並萃取其技術投入（technosphere inputs）資料，最後整理成 Excel 檔案供分析使用。

### 主要功能
- ✅ 支援多批次資料庫處理
- ✅ 自動篩選鋼鐵相關製程
- ✅ 萃取所有技術投入資料
- ✅ 生成 4 張 Excel 工作表
- ✅ 跨批次去重功能
- ✅ 批次來源追蹤

---

## 🔧 安裝說明

### 前置需求
- **Windows 10/11**
- **Python 3.10 以上**（已安裝 pip）

### Step 1：安裝 Python

1. 去 [python.org/downloads](https://www.python.org/downloads/) 下載最新版本
2. 執行安裝程式
3. **務必勾選** ☑️ "Add python.exe to PATH"
4. 點 "Install Now"

驗證安裝：
```bash
python --version
```

### Step 2：下載本專案

方式 A：使用 Git（推薦）
```bash
git clone https://github.com/n66144424-bot/steel-ecoinvent-extractor.git
cd steel-ecoinvent-extractor
```

方式 B：直接下載
- 在 GitHub 點 **Code** → **Download ZIP**
- 解壓縮到你想要的位置

### Step 3：安裝依賴套件

進入專案資料夾，執行：
```bash
pip install -r requirements.txt
```

等待完成，應該會看到：
```
Successfully installed openpyxl lxml tqdm
```

---

## 📂 資料夾結構設定

### 建議的檔案組織方式

在 **D 槽**（或其他位置）建立以下結構：

```
D:\ecoinvent\
├── extract_steel_inputs.py       ← 從本倉庫複製或下載
├── requirements.txt              ← 從本倉庫複製或下載
│
├── batch_1\
│   └── datasets\
│       ├── file1.spold
│       ├── file2.spold
│       └── ... (數千個 .spold 檔案)
│
├── batch_2\
│   └── datasets\
│       ├── file1.spold
│       ├── file2.spold
│       └── ... (數千個 .spold 檔案)
│
└── steel_technosphere_inputs.xlsx  ← 執行後自動產生
```

### 說明
- **batch_1/datasets/** - 第一批 ecoinvent 資料庫解壓縮後的 datasets 資料夾
- **batch_2/datasets/** - 第二批 ecoinvent 資料庫解壓縮後的 datasets 資料夾
- 若有更多批次，依此類推建立 batch_3, batch_4 等

---

## ⚙️ 配置說明

### 修改 DATASETS_FOLDERS

打開 `extract_steel_inputs.py`，找到這部分（約第 27-30 行）：

```python
DATASETS_FOLDERS = [
    r"C:\ecoinvent\batch_1\datasets",      # ← 改這行
    r"C:\ecoinvent\batch_2\datasets",      # ← 改這行
]
```

改成你的實際路徑，例如：

```python
DATASETS_FOLDERS = [
    r"D:\ecoinvent\batch_1\datasets",
    r"D:\ecoinvent\batch_2\datasets",
    r"D:\ecoinvent\batch_3\datasets",      # 若有第三批
]
```

---

## 🚀 使用步驟

### Step 1：準備資料

1. 從 ecoinvent 下載 SPOLD 資料庫
2. 解壓縮，找到 `datasets/` 資料夾
3. 複製到 `D:\ecoinvent\batch_1\datasets\`、`D:\ecoinvent\batch_2\datasets\` 等位置

### Step 2：進入專案資料夾

在檔案管理員中：
1. 找到 `D:\ecoinvent\` 資料夾
2. 在空白處 **Shift + 右鍵**
3. 選 **"在此處開啟 PowerShell 視窗"** 或 **CMD**

### Step 3：執行腳本

在 PowerShell 或 CMD 輸入：

```bash
python extract_steel_inputs.py
```

### Step 4：等待完成

會看到進度條和完成訊息。

### Step 5：檢查結果

執行完後，在同一資料夾會看到 `steel_technosphere_inputs.xlsx` - 用 Excel 打開查看結果！

---

## 📊 輸出檔案說明

### steel_technosphere_inputs.xlsx

包含 4 張工作表：

#### **Sheet 1：執行記錄**
- 執行時間戳
- 總製程數、技術投入總筆數
- 去重模式和批次來源路徑

#### **Sheet 2：製程總覽**
所有找到的鋼鐵製程概覽

#### **Sheet 3：技術投入明細**
所有技術投入的詳細資料

#### **Sheet 4：投入物質加總**
跨所有製程的技術投入物質聚合統計

---

## ❓ 常見問題

### Q：執行時出現 "找不到資料夾"
**A：** 檢查 DATASETS_FOLDERS 中的路徑是否正確。用檔案管理員複製完整路徑。

### Q：沒有找到鋼鐵製程
**A：** 確認 SPOLD 檔案是否有效，或檢查 STEEL_KEYWORDS 設定。

### Q：執行很慢
**A：** 這是正常的。處理數千個 XML 檔案需要時間（通常 5-30 分鐘）。

---

## 📝 版本歷史

### v1.0.0 (2026-06-04)
- 初始版本，支援多批次處理

---

**祝你使用愉快！** 🎉
