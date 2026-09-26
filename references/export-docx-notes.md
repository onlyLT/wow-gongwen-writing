# 导出 Word 操作手册（export_docx.py）

执行 SKILL.md 第 6 步「导出 Word」时读本文件。JSON 字段契约以 `scripts/export_docx.py` 顶部文档字符串为准——**编 JSON 前先读它**。

## 命令用法

```bash
python scripts/export_docx.py spec.json          # 按 JSON 里的 output 落盘
python scripts/export_docx.py spec.json -o x.docx # 覆盖输出路径
python scripts/export_docx.py --demo demo.json    # 生成示例 JSON 参考
```

- 脚本路径以本 skill 根目录为基准，从其他目录调用时写全路径。
- 输出 `SAVED:` 即成功；stderr 出现 `WARNING: unknown key(s)` 说明 JSON 写了契约外的字段名，改正后重跑。

## 编 JSON 的要点

- **版头**：`red_header`（发文机关标志）+ `wenhao`；上行文加 `signer`（只写姓名，脚本自动排成与发文字号同行、居右空一字）。简报/专报用 `brief_header` + `issue_no` + `issuer_left/right`。纪要用 `red_header: "××县人民政府常务会议纪要"` + `issue_no` + `issuer_left/right`。
- **标题**：超过一行（二号字每行约 20 字）时在 `title` 里用 `\n` 手动断行——按词意断、排成梯形或菱形，不把"经费""工作"之类词拆到两行。
- **正文块**：`h1/h2/h3` 与 `text` 同段时，标题末尾带句号（`"h1": "一、总体要求。"`）；独立成行的标题不带。
- **附件说明**：用 `fujian` 列表，只写名称；1 个不编号、多个自动编 `1.` `2.`，名称后的标点会被去掉。附件正文放 `attachments`，排在版记之前另起一页。
- **署名**：`seal` 缺省为 true（加盖印章的排法：日期右空四字、署名居中于日期之上，正文下留两行给印章）；纪要、以署名代章的文件设 `"seal": false`。
- **版记**：`chaosong`（或纪要/简报的 `fensong`）+ `yinfa_org` + `yinfa_date`，固定在最后一页版心底部；放不下时整块移到下一页。
- **纪要名单**：`chuxi` / `liexi` / `qingjia`，排在正文之后。
- 占位符 `【待补：…】` 原样写进 JSON，生成后仍保留在 docx 中，提示用户填定后方可发文。红头、标题、署名里的机关名称若是占位，随稿提示"补齐后改 JSON 重新导出"——直接在 Word 里改，红头压缩比例和标题断行不会跟着变。机关名称占位写成 `【待补：县名】人民政府`（县名含"县"字），不写成 `【待补：县名】县`。

## 版式与字体缺省（GB/T 9704-2012）

A4，天头 37mm、订口 28mm，版心 156mm×225mm；文档网格每页 22 行、每行 28 字；页码四号宋体"— 1 —"，单页居右、双页居左。

| 要素 | 字体 | 字号 |
|---|---|---|
| 发文机关标志 | 方正小标宋简体（红，加粗，超宽自动横向压缩） | 约 40 磅 |
| 标题 | 方正小标宋简体（加粗） | 二号 |
| 一级标题 (h1) | 黑体 | 三号 |
| 二级标题 (h2) | 楷体_GB2312 | 三号 |
| 三级标题 (h3) | 仿宋_GB2312（加粗） | 三号 |
| 正文、附注、附件说明 | 仿宋_GB2312 | 三号 |
| 签发人姓名 | 楷体_GB2312 | 三号 |
| 版记（抄送、印发） | 仿宋_GB2312 | 四号 |
| 数字 / 西文（全部层级） | Times New Roman | 随所在层级 |

单位另有要求时在 JSON 的 `fonts` 里覆盖（数字/西文键名为 `latin`）；本机未装的字体 Word 会自动替换，不影响生成。
数字、字母走 Times New Roman 是按字符自动分槽的，无需单独拆 run；但**全角数字（１２３）算中文字符**会走中文字体，成稿里数字一律用半角。

## 依赖与 python 不可用时的替代调用

- 依赖 `python-docx`（`pip install -r scripts/requirements.txt`）。
- 命令中的 `python` 若不可用或**静默失败**（退出码非 0 且无输出，典型如 Windows 商店占位 stub），改用 `uv run --with python-docx python`（推荐，无需预装依赖）或 `py -3`。

## 可移植性与降级（跨 agent 平台运行时）

本 skill 的**核心写作能力零依赖**——SKILL.md 第 1–5 步只产出 markdown/纯文本，任何 Agent Skills 兼容的 harness 都能跑。第 6 步导出 Word 是**可选能力**：

- 若运行环境**有** `python-docx` → 调 `export_docx.py` 出 .docx。
- 若**无法运行**（python-docx 未装、python 不可用或调用静默失败、无子进程权限）→ **降级**：照常交付 markdown/纯文本成稿 + 版式说明，并提示用户"本环境未装 python-docx，如需 Word 请 pip/uv 安装 python-docx 后重跑第 6 步，或把成稿贴入公文模板"。不因导出不可用而阻断主体交付。

## 导出后随稿提示（必做）

列出 docx **未包含**、须由办公室补全的要素：**印章**（位置已留好）；如有，还包括联合行文的多机关标志与多印章、信函格式版头、版记页前空白页的页码处理。并提醒：本机缺方正小标宋等字体时，打开效果与印制效果会有差异，以单位模板复核一遍。
