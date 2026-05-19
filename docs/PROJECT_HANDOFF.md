# Project Handoff

## Workspace
- Local path: `/Users/user/Desktop/句型辞典html创建`
- GitHub repo: `https://github.com/zhangwenluan22-cyber/pdf-excel-proofread`

## Current Layout
- Root:
  Only active runtime/deploy files stay here: the two site entry files, backend, static dictionary data, main Excel, `render.yaml`, `requirements.txt`
- `docs/`:
  Project notes and handoff docs
- `docs/reference-materials/`:
  Reference PDF and sample split workbook
- `archive/`:
  Backup files not used by the current deployment
- `scratch/`:
  Experimental tools, temporary scripts, MVP files

## Two Sites

### 1. 校对站
- Purpose: Excel 校对、标记 `待校对 / 校对无误 / 有错误`、导出有错误行
- Deployment: Render
- Main frontend file: `pdf_excel_proofread_web.html`
- Main backend file: `pdf_excel_proofread_server.py`
- Online URL: `https://pdf-excel-proofread.onrender.com/`

### 2. 检索站
- Purpose: 内部语法检索、搜索条目、查看释义/接续/解释/例句/参考项目
- Deployment: GitHub Pages
- Main frontend file: `grammar_lookup_web.html`
- Static data file: `dictionary.json`
- Data build script: `build_dictionary_json.py`
- Online URL:
  - `https://zhangwenluan22-cyber.github.io/pdf-excel-proofread/grammar_lookup_web.html`

## Source Data
- Main dictionary Excel: `句型辞典_已校对合并_2-851.xlsx`
- Current server reads this file for `/api/dictionary`
- Static lookup site can use `dictionary.json`
- Old backup Excel moved to: `archive/句型辞典_已校对合并_2-851_参考项目修正前备份.xlsx`

## Update Flow

### Update 校对站
1. Edit:
   - `pdf_excel_proofread_web.html`
   - `pdf_excel_proofread_server.py` if backend changes are needed
2. Run:
   ```bash
   git add .
   git commit -m "describe change"
   git push
   ```
3. Render auto-redeploys

### Update 检索站
1. Edit:
   - `grammar_lookup_web.html`
2. If dictionary data changed:
   ```bash
   python3 build_dictionary_json.py
   ```
3. Run:
   ```bash
   git add .
   git commit -m "describe change"
   git push
   ```
4. GitHub Pages auto-updates

## Current UX Notes
- 校对站 and 检索站 are intentionally separate
- 检索站 has no PDF preview
- 检索站 supports search + kana browsing
- 校对站 focuses on review workflow and error export

## Good Context To Provide In A New Chat
- Which site to change: 校对站 or 检索站
- Target file(s)
- Screenshot or example entry
- Whether change should affect Render, GitHub Pages, or both
