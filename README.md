# Weekly Battery Report Automation (R / Python)

每週電池循環測試資料的自動化整理與報表輸出。  
流程包含：解壓縮週報 zip、讀取多顆電池的 Excel（step/cycle）、擷取「恆流放電」容量序列、分組為 Plan A / Plan B、計算標準化容量（capacity / capacity[1]）、輸出 4×4 總覽圖，並保存中間狀態供後續分析使用。

本專案提供兩個等價版本：
- `R` 版本：輸出 `.Rdata` 與圖檔
- `Python` 版本：輸出 `.pkl` 與圖檔

## What it does

- 從使用者 `Downloads` 取得 `{date}_cs_newexp.zip` 並解壓至工作目錄
- 讀入 32 顆電池（i=1..4, j=1..8）的資料檔案
- 從 `step` 工作表篩選 `工步類型 == "恆流放電"`，擷取 `容量(Ah)` 作為容量序列
- 以容量序列長度作為 cycle 數
- 分組：
  - Plan A：j 為奇數（1,3,5,7），共 16 顆
  - Plan B：j 為偶數（2,4,6,8），共 16 顆
- 統一各 plan 的 X 軸上限為該 plan 中 cycle 最短的那顆電池
- 輸出報表到 `report/`：
  - `PlanA_capacity.png`
  - `PlanB_capacity.png`
- 保存狀態：
  - R：`{date}_cs_newexp.Rdata`
  - Python：`{date}_cs_newexp.pkl`

## Data expectations

### Zip / folder naming

- Zip 檔：`{date}_cs_newexp.zip`
- 預期解壓後資料夾：`{date}_cs_newexp/`

### Excel file pattern

每顆電池會嘗試下列檔名尾碼（取第一個存在的）：
- `240076-{i}-{j}-2818573959.xlsx`
- `240076-{i}-{j}-2818573960.xlsx`
- `240076-{i}-{j}-2818573961.xlsx`
- `240076-{i}-{j}-2818573962.xlsx`
- `240076-{i}-{j}-2818573963.xlsx`

### Required sheets / columns

- 工作表：
  - `step`
  - `cycle`
- `step` 表需包含欄位：
  - `工步類型`
  - `容量(Ah)`

## Repository structure

```text
.
├─ R/
│  └─ weekly_report.R
├─ python/
│  └─ weekly_report.py
├─ examples/
│  ├─ PlanA_capacity.png
│  └─ PlanB_capacity.png
├─ requirements.txt
├─ LICENSE
└─ README.md
````

## Setup

### Python

* Python 3.9+（建議 3.10+）
* 主要套件：

  * pandas
  * numpy
  * matplotlib
  * openpyxl

安裝：

```bash
pip install -r requirements.txt
```

`requirements.txt` 可參考：

```txt
pandas
numpy
matplotlib
openpyxl
```

### R

* R 4.2+
* 套件：

  * readxl
  * tidyverse
  * lubridate

安裝：

```r
install.packages(c("readxl", "tidyverse", "lubridate"))
```

## Usage

目前兩個版本都以檔案開頭的變數控制日期與工作路徑（可依需求改成參數化）。

### Python

1. 編輯 `python/weekly_report.py`：

   * `date = 20260107`
   * `file_name = r"D:\研究\weekly"`

2. 執行：

```bash
python python/weekly_report.py
```

輸出位置：

* `D:\研究\weekly\{date}_cs_newexp\report\PlanA_capacity.png`
* `D:\研究\weekly\{date}_cs_newexp\report\PlanB_capacity.png`
* `D:\研究\weekly\{date}_cs_newexp\report\{date}_cs_newexp.pkl`

### R

1. 編輯 `R/weekly_report.R`：

   * `date <- 20260107`
   * `file_name <- "D:\\研究\\weekly"`

2. 執行：

```r
source("R/weekly_report.R")
```

輸出位置：

* `D:\研究\weekly\{date}_cs_newexp\report\PlanA_capacity.png`
* `D:\研究\weekly\{date}_cs_newexp\report\PlanB_capacity.png`
* `D:\研究\weekly\{date}_cs_newexp\report\{date}_cs_newexp.Rdata`

## Output

輸出都在 `report/` 目錄下：

* `PlanA_capacity.png`：Plan A（16 顆）標準化容量 vs cycle，4×4 總覽圖
* `PlanB_capacity.png`：Plan B（16 顆）標準化容量 vs cycle，4×4 總覽圖

## Notes / Troubleshooting

* `Cannot find any file...`：表示該電池對應的 59~63 尾碼 Excel 檔都不存在，檢查檔名或資料是否完整解壓。
* `sheet not found`：確認 Excel 內是否有 `step` / `cycle` 工作表。
* 欄位名稱不一致：若不同批次欄位命名略有差異，需在腳本內做對應或清理。

## Why two implementations

* 可依個人習慣直接選用 R 或 Python
* 兩邊輸出一致的圖檔，方便交叉驗證結果
* 保留 `.Rdata` / `.pkl` 作為後續分析的中間層，避免每次重讀 Excel

## License

MIT


