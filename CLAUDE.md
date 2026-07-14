# Excel 學習站

## 部署
- **Vercel（主）**: https://excel-learning-theta.vercel.app
- **Netlify（301→Vercel）**: https://splendorous-alpaca-7e65de.netlify.app
- **GitHub**: Astor0504/excel-learning

## 環境變數（Vercel Production）
- `ANTHROPIC_KEY`、`ANTHROPIC_MODEL=claude-haiku-4-5`（教學對話 Haiku 夠用、省 token）
- `AZURE_KEY`、`AZURE_REGION=eastasia`

## 已解決的歷史雷點（不要走回頭路）
- **2026-04-14**：前端 `app.js` 裡的 `getApiKey()` 曾把 Anthropic key base64 寫死，已改走 `/api/chat`。**永遠不要把 key 放回前端。**
- 前端用 `<meta name="api-base">` 拿後端 URL；打 `/api/chat`、`/api/tts`、`/api/voices`。
- TTS epoch guard 保留原樣，不要拿掉。
- **2026-07-05**：WPS 宏的正解是**內建 JS 宏（JSA）**，macOS 國際版免費 WPS 實測可用（見 P5-05 番外課）。舊說法「WPS 裝 VBA 支援元件就能跑巨集」查證不成立、已從 P5-02/P5-03 移除——**不要寫回去**。細節見 auto-memory `reference_wps_js_macro_mac`。

## 教學素材慣例（可復用元件）
Excel 站建立了一套 lesson 結構，新增課程請沿用：
- **公式動畫**：分步拆解（highlight → 代入 → 結果）
- **決策樹**：用 Mermaid（`npx --yes @mermaid-js/mermaid-cli`）產 SVG，light/dark 兩版
- **平台 callout（兩維度都要查）**：凡遇平台差異一律加 callout，且**分開檢查兩維度**——①macOS 版 Excel 限制（本站 macOS 取向）②WPS 差異（台灣 WPS 使用者多）。常見只補了 macOS 卻漏 WPS（如 Power Query/Power Pivot 課）；新增或修課時兩維度都要確認。
- 每課 HTML 行數 `wc -l` 控制在可維護範圍
- **課程資料按課切片(2026-07-14 起)**:每課的互動內容與示範資料在 `lesson-data/{slug}.js`(`export const CONTENT / DEMOS`),首頁用的全站資料+各課 meta 在 `site-data.js`。舊的 xlsx-content.js / lesson-demos-data.js 已移除——**新增課程時直接建立新切片檔,不要重建整包大檔**。

## 建站時的部署流程
1. `git push` → Vercel 自動部署
2. `npx netlify-cli deploy --dir=. --prod --site=<excel-site-id>`（reference_deployment_accounts）
3. 驗證 5 個關鍵 URL 都 200：`/`、`/quiz-data.js`、`/lessons/<id>.html`、`/lessons/img/<svg>`、`/styles.css`
