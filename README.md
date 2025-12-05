# PPT Secretary (Gemini Powered)

## 專案目的
- 透過 Google Gemini 文字與視覺模型自動依照指令修改既有簡報，並於修改後進行視覺化驗證以減少誤差。
- 內建 `ppt_tool/ppt_api.py` 的高階 helper，減少直接呼叫 python-pptx API 所帶來的風險。
- 互動式 CLI（`ppt_tool/main.py`）會先檢視簡報內容再按照使用者需求自動產生並執行修改程式。

## 系統需求與環境設定
1. **Python**：建議 3.10 以上（需支援 `python-pptx`、`google-generativeai`）。
2. **作業系統**：
   - Windows：可使用 PowerPoint COM 進行 PDF 轉檔（需已安裝 Office PowerPoint）。
   - macOS / Linux：需安裝 LibreOffice (`soffice`) 才能輸出 PDF 供視覺驗證；若無則僅能使用純文字摘要。
3. **套件**：`pip install -r ppt_tool/requirements.txt`，其中包含 `python-pptx`, `google-generativeai`, `python-dotenv`，以及 Windows 平台的 `pywin32`。
4. **環境變數**：於專案根目錄建立 `.env`，內容範例如下：
   ```
   GOOGLE_API_KEY=你的 Gemini API Key
   GEMINI_TEXT_MODEL=gemini-2.5-flash
   GEMINI_VISION_MODEL=gemini-2.5-flash
   ```
   `main.py` 會自動載入並套用設定；若未設置金鑰將無法使用模型。
5. **預設目標檔**：根目錄下的 `presentation.pptx`。若檔案不存在，系統會在首次指令時建立空白簡報。

## 安裝與操作步驟
1. 取得程式碼並進入專案目錄：
   ```bash
   git clone <repo-url>
   cd P2025_PPTOR
   ```
2. 建立虛擬環境並啟用：
   ```bash
   python -m venv .venv
   # Windows
   .venv\Scripts\activate
   # macOS/Linux
   source .venv/bin/activate
   ```
3. 安裝依賴：
   ```bash
   pip install -r ppt_tool/requirements.txt
   ```
4. 設定 `.env`，填入 Gemini API Key 及模型名稱。
5. 啟動工具：
   ```bash
   python -m ppt_tool.main          # 一般模式
   python -m ppt_tool.main -d       # Debug 模式，會印出 Gemini 產生的程式碼
   ```
6. 依 CLI 提示輸入自然語言指令。成功修改後程式會嘗試開啟 `presentation.pptx` 供檢視，並輸出視覺驗證回饋。

## 專案結構與簡述
```
P2025_PPTOR/
├── ppt_tool/
│   ├── main.py          # CLI 入口，負責載入 .env 與整合 inspector / modifier
│   ├── converter.py     # 偵測 COM/LibreOffice 並輸出 PDF，供視覺參考
│   ├── inspector.py     # 產生目前簡報的文字摘要並觸發 PDF 轉檔
│   ├── modifier.py      # 建構 Gemini Prompt、執行產生程式碼、並呼叫 Vision 驗證
│   ├── ppt_api.py       # 封裝常用的 python-pptx 操作 helper
│   ├── test.py          # 手動測試或示範腳本
│   └── requirements.txt # 套件列表
├── tests/
│   ├── test_ppt_api.py  # 針對 helper API 的 pytest 測試
│   └── artifacts/       # 測試輸出樣本
├── temp_visuals/        # PDF 轉檔輸出，供視覺模型參考與驗證
├── presentation.pptx    # 預設操作目標
└── .env                 # 存放 Gemini 金鑰與模型設定
```

## 開發者注意事項
- **API Key 安全**：`GOOGLE_API_KEY` 僅放在 `.env`，勿提交到版本控制。
- **PowerPoint Lock**：若使用者開啟 `presentation.pptx`，系統會偵測 `~$` lock 檔並提示先關閉。
- **PDF 轉檔引擎**：`converter.py` 依序嘗試 COM → LibreOffice → 無法轉檔；缺少引擎將無法進行視覺驗證。
- **Debug 模式**：啟用 `-d/--debug` 以檢視 Gemini 產生的 Python 程式，方便排查失敗原因。
- **測試**：執行 `pytest` 會在 `tests/artifacts` 產出樣本檔案，必要時可清理；CI 流程請記得忽略該資料夾。
- **暫存檔案**：`temp_visuals/` 會累積 PDF，若想清空可安全刪除資料夾，程式會於下次需要時重新建立。
- **依賴更新**：若新增 helper 或外部套件，請同步更新 `ppt_tool/requirements.txt` 與 README，並為新功能補上測試。

