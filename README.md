# XML to CSV 轉換工具

這是一個將多個 XML 檔案轉換為單一 CSV 主檔和可選 TSV 明細檔的 Python 工具。

## 功能特色

- 🔄 **批次處理**：一次處理多個 XML 檔案
- 📊 **雙重輸出**：主檔 CSV + 明細檔 TSV（可選）
- 🗺️ **距離計算**：自動計算路線總距離（使用 Haversine 公式）
- 🎯 **完整 JSON**：主檔中的 ViaPoint 資料以完整 JSON 格式儲存，包含所有欄位
- 🛡️ **錯誤處理**：跳過無法解析的檔案，繼續處理其他檔案
- 🌍 **UTF-8 支援**：完整支援中文和特殊字元

## 輸出檔案格式

### 主檔 CSV 欄位
- `name` - 路線名稱
- `description` - 路線描述
- `CreationTimeUTC` - 建立時間（UTC）
- `IsManuallyCorrected` - 是否手動修正
- `TotalDistanceKm` - 總距離（公里，依 ViaPoint 順序計算）
- `RouteInfo_IgnoringRestrictions` - 忽略限制資訊
- `RouteInfo_MapCorrectionInfo_DatasetInfo_ImageInfo_ImageName` - 地圖影像名稱
- `RouteInfo_MapCorrectionInfo_DatasetInfo_ImageInfo_StartMapId` - 起始地圖 ID
- `RouteInfo_ViaPoints_NumVia` - ViaPoint 數量
- `RouteInfo_ViaPoints_ViaPoint` - ViaPoint 完整 JSON 資料（包含所有欄位）

### 明細檔 TSV 欄位（可選）
- `placemark_index` - Placemark 索引
- `placemark_name` - Placemark 名稱
- `seq` - 順序
- `Position` - 位置座標
- `Lat` - 緯度
- `Lon` - 經度
- `GroupID` - 群組 ID
- `Segment` - 路段
- `Heading` - 方向
- `Type` - 類型
- `LinkToGeom` - 幾何連結
- `Direction` - 方向
- `TTSRemark` - TTS 備註
- `WorkType` - 工作類型
- `MMRule` - MM 規則
- `ManeuverID` - 操作 ID
- `ManeuverNumber` - 操作編號
- `IsDeadEnd` - 是否為死路

## 安裝需求

- Python 3.6 或以上版本
- 無需額外套件（使用標準函式庫）

## 使用方法

### 基本用法

```bash
# 處理當前目錄下所有 XML 檔案
python xml_to_csv.py *.xml -o output.csv

# 處理指定資料夾中的所有 XML 檔案
python xml_to_csv.py --dir "C:\path\to\xml\files" -o output.csv

# 產生主檔和明細檔
python xml_to_csv.py *.xml -o output.csv --detail-tsv detail.tsv
```

### 進階用法

```bash
# 遞迴掃描子資料夾
python xml_to_csv.py --dir "C:\xml\files" --recursive -o output.csv

# 指定檔案模式
python xml_to_csv.py --dir "C:\xml\files" --pattern "route_*.xml" -o output.csv

# 處理多個特定檔案
python xml_to_csv.py file1.xml file2.xml file3.xml -o output.csv
```

## 參數說明

| 參數 | 說明 | 範例 |
|------|------|------|
| `input` | 輸入檔案或萬用字元（可多個） | `*.xml`, `file1.xml file2.xml` |
| `-o, --output-csv` | 輸出主檔 CSV 路徑（必填） | `-o routes.csv` |
| `--detail-tsv` | 可選：輸出明細 TSV 路徑 | `--detail-tsv detail.tsv` |
| `--dir` | 要掃描的資料夾路徑 | `--dir "C:\xml\files"` |
| `--pattern` | 搭配 --dir 的檔案樣式（預設 *.xml） | `--pattern "route_*.xml"` |
| `--recursive` | 搭配 --dir 遞迴掃描 | `--recursive` |

## 使用範例

### 範例 1：基本轉換
```bash
python xml_to_csv.py *.xml -o all_routes.csv
```
輸出：`all_routes.csv` 包含所有 XML 檔案的合併資料

### 範例 2：完整轉換（含明細檔）
```bash
python xml_to_csv.py *.xml -o routes.csv --detail-tsv routes_detail.tsv
```
輸出：
- `routes.csv` - 主檔（每個 Placemark 一列）
- `routes_detail.tsv` - 明細檔（每個 ViaPoint 一列）

### 範例 3：資料夾掃描
```bash
python xml_to_csv.py --dir "D:\Download\DSNY\Routes-Stage.2025-09-05T09_57_29" -o 1002.csv
```
輸出：`1002.csv` 包含指定資料夾中所有 XML 檔案的資料

## 輸出檔案格式

### CSV 主檔
- 編碼：UTF-8 with BOM
- 分隔符：逗號 (,)
- 引號：所有欄位都加引號
- 換行：Windows 格式 (\r\n)

### TSV 明細檔
- 編碼：UTF-8 with BOM
- 分隔符：Tab (\t)
- 引號：最小化引號
- 換行：Windows 格式 (\r\n)

## 錯誤處理

程式會自動處理以下情況：
- 跳過無法解析的 XML 檔案（顯示警告訊息）
- 跳過無法讀取的檔案（顯示警告訊息）
- 如果沒有可處理的檔案，程式會退出並顯示錯誤訊息

## 注意事項

1. **檔案大小限制**：主檔中的 JSON 欄位會自動移除換行符號，避免 Excel 單格字元限制
2. **距離計算**：使用 Haversine 公式計算球面距離，精確度約 6 位小數
3. **編碼支援**：所有輸出檔案都使用 UTF-8 with BOM 編碼，確保中文正確顯示
4. **記憶體使用**：程式會將所有資料載入記憶體，處理大量檔案時請注意記憶體使用量

## 技術細節

- **距離計算**：使用 Haversine 公式計算地球表面兩點間距離
- **JSON 處理**：自動清理控制字元，確保 CSV 格式正確
- **檔案處理**：支援 Windows 路徑和萬用字元展開
- **錯誤恢復**：單一檔案錯誤不會影響其他檔案的處理

## 授權

此工具為開源軟體，可自由使用和修改。

## 版本歷史

- v1.0 - 初始版本，支援基本 XML 到 CSV 轉換功能
