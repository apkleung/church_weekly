# 主日崇拜程序表（基督教宣聖會馬鞍山堂）

一個網頁工具：填好每週崇拜程序（講員、領會、詩歌、宣召/讀經經文、報告、奉獻、人數…），
一撳「輸出Word檔」就自動生成一份跟足教會原版 Word 版式嘅程序表 `.docx` 落嚟，可直接列印／匯出。

## 用家（填表者）
1. 開啟網頁（Vercel 提供嘅 URL，或直接開 `index.html`）
2. 揀崇拜日期 → 填晒各欄位（人名下拉可以直接揀，或打字）
3. 撳 **「📤 輸出Word檔」** → 自動生成並下載 `程序表-YYYY-MM-DD.docx`
4. 用 Word 打開，列印／另存 PDF 都得

貼士：
- 人名下拉（講員/領會…）可以揀可以打字；揀唔岩想直接用下拉再揀都得。
- 宣召／讀經／每月金句：揀書卷・章・節（留空=全章），會自動拉返經文內文。
- 聖餐：第4個主日先有，撳開「聖餐」掣再揀負責人。
- 揀咗第4主日，日期下會提示「係第 N 個星期日／有聖餐安排」。
- 儲存版本：揀日期 → 💾儲存；之後喺「已儲存版本」下拉揀返 → 🕘顯示此日期版本。

## 技術部署（管理者）
- Vercel + GitHub。
- 前端 `index.html` 撳掣會 POST JSON 去 `/api/generate`。
- `api/generate.py`（Vercel Python serverless function）接收 JSON，
  行 `make_program_tpl.py`，用原版 Word 模板生成 `.docx` 回傳。
- 一定要上傳嘅檔案（MUST）：
  - `index.html`（網頁）
  - `api/generate.py`（Python function，必須放喺 `api/` 子資料夾）
  - `make_program_tpl.py`（生成邏輯）
  - `requirements.txt`（python-docx、requests）
  - `vercel.json`（Vercel 設定）
  - `MOS程序表2026-09-20 (1).docx`（普通主日模板）
  - `MOS程序表2026-08-2.docx`（有聖餐主日模板）
- 模板檔必須同 `make_program_tpl.py` 放埋同一層（script 用 `__file__` 定位模板，唔關 cwd 事）。

## 本地測試（無需 Vercel）
```
pip install -r requirements.txt
python make_program_tpl.py 程序表-2026-09-27.json   # 生成 .docx
```

## 模板更換
想改版式，直接用 Word 改 `MOS程序表2026-09-20 (1).docx`（普通版）或
`MOS程序表2026-08-2.docx`（聖餐版）就得——generator 只會更新文字，版面跟晒你改。
注意：script 會自動移除 Word 內容控制（combo box / 下拉人名）轉做純文字。