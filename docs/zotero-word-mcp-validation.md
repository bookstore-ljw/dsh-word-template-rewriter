# zotero-word-mcp 在 Zotero 10.0.7 上的兼容性验证

日期：2026-11-11
目标：确认 `zotero-word-mcp`（README 仅声明支持 Zotero 7）能否在本机 Zotero 10.0.7 上
读写 Zotero 原生字段。全程只操作临时副本，未改 DSH profile、未改全局 Python、未改 Zotero 库。

## 环境

| 项 | 值 |
|---|---|
| Word | 16.0.20326（`C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE`） |
| Zotero | 10.0.7，安装于 `D:\zetro`（`compatibility.ini` 的 `LastVersion=10.0.7_...`） |
| Zotero Word 集成 | `D:\zetro\integration\word-for-windows\libzoteroWinWordIntegration.dll` + `%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm` |
| Zotero 运行状态 | 验证时**未启动** |
| Python | 3.12.10，独立 venv：`%TEMP%\zotero-word-mcp-check\.venv` |
| 安装方式 | `pip install -e . -i https://pypi.tuna.tsinghua.edu.cn/simple`（PyPI 直连不通，清华镜像可用） |
| 被测文档 | `fixtures/build/out.docx` 的临时副本 `%TEMP%\zwmcp-full.docx` |

## 结果：字段级可用

`NativeWordBridge`（COM 侧）+ `libzoteroWinWordIntegration.dll`（DLL 侧）在 Zotero 10 上
都能工作，写入的原生字段能被 DLL 读回：

1. `probe_before`：文档打开后 `getDocument` / `getDocumentData` / `getFields` 全部成功，
   空文档返回 `docData=""`、`fields=[]`；
2. `insert_native_bibliography(placement="end")` 写入 `BIBL {} CSL_BIBLIOGRAPHY`，
   占位文本 `{Bibliography}`，字段类型 35；
3. `list_native_fields` 读回 1 个 bibliography 字段；
4. `load_native_document_data` 读回 `ZOTERO_PREF`：
   `styleID=gb-t-7714-2015-numeric-bilingual-no-uppercase-no-url-doi`、`locale=zh-CN`、
   `dataVersion=4`、`sessionID=...`；
5. 再次 `probe` 时 DLL 侧同时返回该字段与 docData JSON，说明 Zotero 10 的集成 DLL
   认得这份文档与写入的字段。

## 未验证 / 风险

- **刷新**（`refresh_document`）没测：它通过 Zotero 的 Word 宏触发，需要 Zotero 进程运行；
- 真实引文插入没测：需要 Zotero 库里真实 item key，`citation_map` 转换同理；
- `allResolved=false` 是预期结果——没有 Zotero 刷新，占位符保持未解析；
- 项目 README 写“Zotero 7”，代码里没有硬编码 7.x 版本判断，DLL 路径靠
  `ZOTERO_WORD_DLL` 覆盖（默认只找 `D:\Zotero\...`、`~/Zotero/...`，本机是 `D:\zetro`，
  必须显式设置）。

## 结论

在 Zotero 10.0.7 上，**原生字段的读写链路可用**，README 的 Zotero 7 是保守声明而非硬限制。
要真正接入 DSH，还需要：Zotero 启动后验证一次 refresh，并用真实 item key 插一条引文。

## 复现

```powershell
$v = "$env:TEMP\zotero-word-mcp-check\.venv\Scripts\python.exe"
$env:ZOTERO_WORD_DLL = 'D:\zetro\integration\word-for-windows\libzoteroWinWordIntegration.dll'
$env:ZOTERO_DATA_DIR = "$env:USERPROFILE\Zotero"
& $v "$env:TEMP\zotero-word-mcp-check\probe_full.py"
```

验证脚本只复制 `fixtures/build/out.docx` 到 `%TEMP%`，写入后不覆盖原文件。
