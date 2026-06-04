# steel-ecoinvent-extractor

本專案用於從 **ecoinvent 3.12 SPOLD** 檔案中，自動擷取與鋼鐵製造相關的 technosphere 投入資料，並輸出為 Excel，方便後續分析。

## 1. 專案目的

- 掃描多個批次（多個資料夾）的 `.spold` 檔案
- 篩選鋼鐵相關製程
- 擷取 technosphere inputs
- 匯出 4 個工作表的 Excel 報表（含執行紀錄、製程總覽、投入明細、聚合材料）
- 支援去重模式，避免重複製程重複計算

## 2. 安裝需求

- Python 3.9+
- pip

安裝套件：

```bash
pip install -r requirements.txt
```

## 3. 建議資料夾結構（Windows）

```text
D:\ecoinvent\
├── extract_steel_inputs.py
├── batch_1\
│   └── datasets\
│       ├── *.spold
├── batch_2\
│   └── datasets\
│       ├── *.spold
└── ...
```

## 4. 設定方式（DATASETS_FOLDERS）

開啟 `/tmp/workspace/n66144424-bot/steel-ecoinvent-extractor/extract_steel_inputs.py`，修改：

```python
DATASETS_FOLDERS = [
    r"D:\ecoinvent\batch_1\datasets",
    r"D:\ecoinvent\batch_2\datasets",
]
```

可依批次數量自由增減。

## 5. 使用步驟

1. 將 ecoinvent 各批次資料解壓到不同 `datasets` 資料夾
2. 編輯 `DATASETS_FOLDERS`
3. （可選）設定 `DEDUPLICATE_PROCESSES = True/False`
4. 在專案目錄執行：

```bash
python extract_steel_inputs.py
```

5. 執行完成後查看輸出 Excel：`steel_technosphere_inputs.xlsx`

## 6. 輸出檔案說明

輸出為 `steel_technosphere_inputs.xlsx`，含 4 個工作表：

1. `execution_log`：執行時間、等級與訊息
2. `process_overview`：每個鋼鐵製程摘要
3. `input_details`：每筆 technosphere input 明細
4. `aggregated_materials`：依批次與投入名稱彙總用量

## 7. 疑難排解

- **找不到資料夾**：確認 `DATASETS_FOLDERS` 路徑正確
- **沒有抓到資料**：確認資料夾內是否有 `.spold`，以及製程名稱是否含鋼鐵關鍵字
- **Excel 開不起來**：請關閉同名 Excel 檔後重新執行
- **執行很慢**：大量檔案屬正常，可先測試單一批次

## 8. FAQ

### Q1：一定要放在 D 槽嗎？
不一定，任何路徑都可，重點是 `DATASETS_FOLDERS` 要對應正確。

### Q2：有多個批次怎麼處理？
在 `DATASETS_FOLDERS` 中加入更多資料夾路徑即可。

### Q3：為什麼要開啟 deduplication？
多批次可能有重複製程，開啟後可避免重複統計。

### Q4：可否輸出 PDF？
建議先輸出 Excel，若需要再由 Excel 另存 PDF。
