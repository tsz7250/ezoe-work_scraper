# 書報爬蟲

針對 ezoe.work 抓取文章，並輸出為格式化的 DOCX 與 PDF。

## 功能摘要

- **批次爬取**：從 `urls.txt` 或指定 txt 檔案讀取 URL，每篇產生 `書名_篇數.md`
- **合併轉換**：依書名合併多篇 Markdown，轉成一本 DOCX（可設定行距、標題樣式）
- **簡轉繁**：使用新同文堂通用詞典修正已知多義詞，再由 BertCC 依語境選擇其餘繁體字形
- **PDF**：Windows 上可用 Microsoft Word 將 DOCX 轉成 PDF；其他平台會略過 PDF
- **一鍵流程**：執行 `main.py` 可依序完成「爬取 → 轉 DOCX/PDF → 刪除中間 Markdown」

## 環境需求

- Python 3.x
- 依賴套件：`pip install -r requirements.txt`
- 首次簡轉繁：需連線下載約 400 MB 的 `bert-base-chinese` 模型；下載後可離線轉換
- 選用：Pandoc（若未安裝，pypandoc 可代為下載）
- 選用：Windows + Microsoft Word（用於 PDF 輸出）

## 安裝

```bash
pip install -r requirements.txt
```

## 使用方式

1. **編輯 URL 列表**：在 `urls.txt` 中每行填入一個文章 URL。例如：

   ```
   https://ezoe.work/books/3/3061-1.html
   https://ezoe.work/books/3/3061-2.html
   ```

2. **完整流程**（爬取 → 轉 DOCX/PDF ）：
   ```bash
   python main.py
   ```
   需先存在 `urls.txt`。

---

**分開執行**：

- **僅爬取**（只產生 Markdown）：
  ```bash
  python scrape_to_markdown.py
  ```
  或指定 URL 檔案：`python scrape_to_markdown.py 你的網址列表.txt`

- **僅轉換**（已有 `書名_數字.md` 時，合併並轉 DOCX/PDF）：
  ```bash
  python convert_md_to_docx.py
  ```
  會掃描當前目錄下符合 `書名_數字.md` 的檔案，按書名分組後各產生一本 DOCX（與可選的 PDF）。

## 專案檔案說明

| 檔案 | 說明 |
|------|------|
| `scrape_to_markdown.py` | 從網頁抓取內容並轉成 Markdown，可單一 URL 或從 txt 批次處理 |
| `convert_md_to_docx.py` | 將當前目錄的 `書名_數字.md` 依書名合併，以語境模型簡轉繁後轉成 DOCX，並可選轉 PDF |
| `main.py` | 一鍵執行：讀取 `urls.txt` 爬取 → 呼叫轉換 → 刪除產生的 Markdown |
| `urls.txt` | 預設的 URL 列表，每行一個文章網址 |

## 注意事項

- **首次模型下載**：BertCC 首次載入時會從 Hugging Face 下載約 400 MB 的 `bert-base-chinese` 權重；安裝 PyTorch 也會增加約 124 MB 套件，下載後可離線執行簡轉繁。
- **簡轉繁限制**：新同文堂詞典只套用字形對應，不採用地區用詞改寫；BertCC 仍可能判錯，須以實文檢查重要輸出。詞典授權與來源見 `data/TONGWEN_LICENSE.txt`。
- **PDF**：需在 **Windows** 上且已安裝 Microsoft Word；macOS/Linux 會略過 PDF 輸出，簡轉繁仍可執行。
- **pywin32**：僅在 Windows 上且要輸出 PDF 時需要。
