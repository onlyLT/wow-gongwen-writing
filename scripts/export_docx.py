# -*- coding: utf-8 -*-
"""
export_docx.py — 通用公文 → Word(.docx) 导出器（GB/T 9704-2012 近似版式）

数据驱动：吃一个 JSON 规格，产出排版好的 .docx。任意文种通用
（通知/请示/报告/批复/函/意见/纪要 + 专报/简报/总结/方案/讲话稿…）。

用法：
    python export_docx.py spec.json                 # 输出到 spec 里的 output
    python export_docx.py spec.json -o 目标.docx    # 覆盖输出路径
    python export_docx.py --demo demo.json          # 写一份示例 JSON 供参考

页面（按 GB/T 9704-2012）：A4；天头 37mm、订口 28mm，版心 156mm×225mm；
文档网格每页 22 行、每行 28 字；正文仿宋_GB2312 三号，行距固定 28.95 磅。

JSON 规格（所有字段可选，缺省即不出该要素）：
{
  "output": "outputs/xxx.docx",           // 输出路径（-o 可覆盖）

  // —— 版头 ——
  "fenhao": "000123",                     // 份号（涉密件；版心左上第一行）
  "miji": "秘密★1年",                      // 密级和保密期限（黑体，左上）
  "jinji": "特急",                         // 紧急程度（黑体，左上）
  "red_header": "武义县财政局文件",         // 发文机关标志 / 纪要标志（红、居中，距版心上缘约
                                           //   35mm；超宽时自动横向压缩）
  "brief_header": "财政专报",              // 简报/专报报头（红、居中；与 red_header 二选一）
  "issue_no": "第5期",                     // 期号（简报/专报/纪要）
  "issuer_left": "武义县财政局编",          // 报头下左侧编发单位（左空一字）
  "issuer_right": "2026年7月2日",          // 报头下右侧日期（右空一字）
  "wenhao": "武财〔2026〕5号",             // 发文字号（红头下空二行，居中）
  "signer": "张三",                        // 签发人姓名（上行文）。有它时发文字号居左空一字、
                                           //   "签发人：姓名"居右空一字，两者同处一行
  "red_line": true,                        // 红色分隔线（有红头/报头时默认 true）

  // —— 主体 ——
  "title": "××关于××的报告",              // 标题（分隔线下空二行，二号小标宋居中）。
                                           //   超过一行时用 \\n 手动断行：词意完整、排成梯形或菱形
  "zhusong": "金华市财政局：",             // 主送机关（标题下空一行，顶格，末尾全角冒号）
  "body": [ ... ],                         // 正文块，见下
  "fujian": ["附件名称一", "附件名称二"],   // 附件说明（正文下空一行、左空二字；1 个不编号，
                                           //   多个自动编 1. 2.；名称后不加标点）
  "chuxi": "王某某、李某某",                // 纪要出席人员（黑体"出席："+ 仿宋名单，回行对齐）
  "liexi": "…", "qingjia": "…",            // 纪要列席、请假人员（同上）
  "seal": true,                            // 是否加盖印章，决定署名与日期的排法（默认 true）
  "signoff_name": "武义县财政局",           // 发文机关署名
  "signoff_date": "2026年7月2日",          // 成文日期（阿拉伯数字，月日不编虚位）
  "fuzhu": "（联系人：李四　电话：0579-…）", // 附注（成文日期下一行，左空二字，加圆括号）
  "attachments": [                         // 附件正文（另面编排，排在版记之前）
    {"label": "附件1", "title": "附件标题", "body": [ ...同正文块... ]}
  ],

  // —— 版记（浮动表格，固定在最后一页版心底部）——
  "chaosong": "市财政局，县审计局。",       // 抄送机关（末尾句号；缺句号自动补）
  "fensong": "…",                          // 分送（纪要/简报用，与 chaosong 二选一）
  "yinfa_org": "武义县财政局办公室",        // 印发机关（左空一字）
  "yinfa_date": "2026年7月2日",            // 印发日期（右空一字，自动补"印发"）
  "page_number": true,                     // 页码（默认 true："— 1 —"，四号宋体，
                                           //   单页居右空一字、双页居左空一字）
  "fonts": {"body": "仿宋_GB2312", "h1": "黑体", "h2": "楷体_GB2312",
            "h3": "仿宋_GB2312", "title": "方正小标宋简体", "header": "方正小标宋简体",
            "latin": "Times New Roman"}      // 可覆盖字体名；latin=数字/西文
}

正文块类型：
  {"text": "..."}                  普通正文段（仿宋三号，首行缩进2字，两端对齐）
  {"h1": "一、…", "text": "..."}    一级标题（黑体三号）+ 可选同段正文
  {"h2": "（一）…", "text": "..."}  二级标题（楷体三号）+ 可选同段正文
  {"h3": "1.…", "text": "..."}      三级标题（仿宋三号加粗）+ 可选同段正文
  {"plain": "...", ...}            自由行，可带 align/font/size/bold/color/no_indent/line
  {"pagebreak": true}              分页
  标题与正文同段（run-on）时，标题末尾须带句号，如 {"h1": "一、总体要求。", "text": "…"}。
  文本中的 \\n 会变成段内换行。

字体缺省（GB/T 9704 标准层级，经实际发文核校）：标题=方正小标宋简体二号加粗，
一级标题=黑体三号，二级标题=楷体_GB2312 三号，三级标题=仿宋_GB2312 三号加粗，
正文/附注=仿宋_GB2312 三号，发文机关标志=方正小标宋简体加粗（红）。可在 fonts 里
覆盖任一层级（需本机已安装；未安装时 Word 会自动替换，不影响生成）。
所有层级中的半角数字、西文字母（发文字号、日期、电话、编号等）统一用
Times New Roman（fonts.latin 可覆盖），汉字与中文标点仍用该层级的中文字体；
全角数字属中文字符，走中文字体，故正文数字请用半角。页码按国标用宋体。

不产出的要素：印章（署名与日期已按有无印章的国标位置留好，章由办公室加盖）、
版记前空白页的页码处理、联合行文的多机关标志与多印章、信函格式与命令格式版头。
未知字段会在 stderr 打印 "WARNING: unknown key(s)" 并忽略——编 spec 时不要臆造字段名。

依赖：python-docx（pip install python-docx）
"""
import sys, json, argparse, os
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

RED = RGBColor(0xFF, 0x00, 0x00)
ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
         "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}

# 版面（GB/T 9704-2012 第 5 章）：A4，天头 37mm，订口 28mm，版心 156×225mm。
TEXT_W = 156 / 25.4 * 72            # 版心宽 ≈ 442.2 磅
LINE = 28.95                        # 22 行/页：225mm ÷ 22 ≈ 28.95 磅（取 29 磅会只排下 21 行）
BODY = 16                           # 三号
CHAR_PITCH = TEXT_W / 28            # 每行 28 字的字符网格间距 ≈ 15.79 磅

# 层级用字（经实际公文核校）：标题=方正小标宋二号，一级=黑体三号，
# 二级=楷体三号，三级=仿宋三号加粗，正文=仿宋三号。未装字体时 Word 自动替换。
DEFAULT_FONTS = {"body": "仿宋_GB2312", "h1": "黑体", "h2": "楷体_GB2312",
                 "h3": "仿宋_GB2312", "title": "方正小标宋简体",
                 "header": "方正小标宋简体", "latin": "Times New Roman"}
# 数字/西文字母统一用 Times New Roman（公文通行做法）：Word 按字符分槽取字，
# 半角数字、字母走 w:ascii/w:hAnsi，汉字与中文标点走 w:eastAsia。
# build() 会用 fonts["latin"] 覆盖该值。
LATIN_FONT = DEFAULT_FONTS["latin"]

KNOWN_KEYS = {
    "output", "fenhao", "miji", "jinji", "red_header", "brief_header", "issue_no",
    "issuer_left", "issuer_right", "wenhao", "signer", "red_line", "title", "zhusong",
    "body", "fujian", "chuxi", "liexi", "qingjia", "seal", "signoff_name",
    "signoff_date", "fuzhu", "attachments", "chaosong", "fensong", "yinfa_org",
    "yinfa_date", "page_number", "fonts",
}
KNOWN_BLOCK_KEYS = {"text", "h1", "h2", "h3", "plain", "align", "font", "size", "bold",
                    "color", "no_indent", "line", "pagebreak"}


def warn(msg):                                # ASCII：Windows 管道输出走 GBK，中文会乱码
    print("WARNING:", msg, file=sys.stderr)


def em_width(s):
    """估算三号字文本宽度（以网格字为单位）：汉字、全角符号计 1，半角数字字母约 0.5，
    汉字与半角字符交界处 Word 自动加约 0.2 字间距（实测于 Word 渲染）。"""
    w, prev = 0.0, None
    for ch in s:
        cjk = ord(ch) >= 0x2E80
        w += 1 if cjk else 8 / CHAR_PITCH
        if prev is not None and cjk != prev:
            w += 0.2
        prev = cjk
    return w


def set_font(run, cn, size=BODY, bold=False, color=None, latin=None, scale=None):
    latin = latin or LATIN_FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = latin                     # 写入 w:ascii / w:hAnsi（数字、西文）
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    rfonts.set(qn('w:eastAsia'), cn)          # 汉字、中文标点
    rfonts.set(qn('w:cs'), cn)
    rfonts.set(qn('w:hint'), 'eastAsia')      # 中文引号等两可字符按中文字体取字
    if color is not None:
        run.font.color.rgb = color
    if scale and scale < 100:                 # 字符横向缩放（发文机关标志过长时）
        w = OxmlElement('w:w'); w.set(qn('w:val'), str(scale))
        rpr.append(w)


def add_run(p, text, cn, size=BODY, bold=False, color=None, latin=None, scale=None):
    r = p.add_run(text)                       # \n → 段内换行，\t → 制表符
    set_font(r, cn, size, bold, color, latin, scale)
    return r


def set_ind(p, left=None, right=None, first=None, hanging=None, unit=None):
    """按"字"设置缩进。一律写绝对值（twips），不写 *Chars 属性：Word 对 leftChars 与
    hangingChars 并用时的解释与规范不一致（首行落在 leftChars 处），WPS 等也各有差异。
    三号正文 1 字 = 网格字距 CHAR_PITCH；四号行（版记、页码）传 unit=14。"""
    unit = unit or CHAR_PITCH
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind'); pPr.append(ind)
    for attr, chars in (("left", left), ("right", right),
                        ("firstLine", first), ("hanging", hanging)):
        if chars is not None:
            ind.set(qn('w:' + attr), str(int(round(chars * unit * 20))))


def new_para(doc, align=None, line=LINE, space_before=0, space_after=0,
             first_indent=None, snap=True):
    p = doc.add_paragraph()
    if align in ALIGN:
        p.alignment = ALIGN[align]
    pf = p.paragraph_format
    pf.line_spacing = Pt(line)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if first_indent:
        set_ind(p, first=first_indent)
    if not snap:                              # 大字号行不对齐字符网格，避免字距被拉开
        s = OxmlElement('w:snapToGrid'); s.set(qn('w:val'), '0')
        p._p.get_or_add_pPr().append(s)
    return p


def blank(doc, n=1):
    for _ in range(n):
        new_para(doc)


def bottom_border(p, color='FF0000', sz=12, space_pt=11):
    """段落下边框作分隔线；space_pt 为线与文字的距离（发文字号下 4mm ≈ 11 磅）。"""
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), str(sz))
    b.set(qn('w:space'), str(space_pt)); b.set(qn('w:color'), color)
    pbdr.append(b); pPr.append(pbdr)


def setup_page(doc, fonts, page_number):
    s = doc.sections[0]
    s.page_width, s.page_height = Mm(210), Mm(297)
    s.top_margin, s.bottom_margin = Mm(37), Mm(35)
    s.left_margin, s.right_margin = Mm(28), Mm(26)
    s.footer_distance = Mm(21)                # 页码行上缘约在版心下缘之下 7mm（实测）

    grid = s._sectPr.find(qn('w:docGrid'))
    if grid is None:
        grid = OxmlElement('w:docGrid'); s._sectPr.append(grid)
    grid.set(qn('w:type'), 'linesAndChars')
    grid.set(qn('w:linePitch'), str(int(round(LINE * 20))))
    grid.set(qn('w:charSpace'), str(int(round((CHAR_PITCH - BODY) * 4096))))

    normal = doc.styles['Normal']             # 网格字距以 Normal 字号为基准
    normal.font.size = Pt(BODY)
    normal.font.name = LATIN_FONT
    normal.element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fonts["body"])
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = Pt(LINE)

    if page_number:
        doc.settings.odd_and_even_pages_header_footer = True
        for footer, align in ((s.footer, "right"), (s.even_page_footer, "left")):
            p = footer.paragraphs[0]
            p.alignment = ALIGN[align]
            p.paragraph_format.line_spacing = Pt(LINE)
            set_ind(p, right=1 if align == "right" else None,
                    left=1 if align == "left" else None, unit=14)
            add_run(p, "— ", "宋体", 14, latin="宋体")
            add_page_field(p)
            add_run(p, " —", "宋体", 14, latin="宋体")


def add_page_field(p):
    def fld(kind=None, instr=None, text=None):
        r = p.add_run(text or "")
        set_font(r, "宋体", 14, latin="宋体")
        if kind:
            fc = OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'), kind)
            r._r.append(fc)
        if instr:
            it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve')
            it.text = instr; r._r.append(it)
    fld("begin"); fld(instr=" PAGE "); fld("separate"); fld(text="1"); fld("end")


def build_header(doc, spec, fonts):
    last = None                               # 最后一个版头段落，红线挂在它下方
    top_lines = 0
    if spec.get("fenhao"):
        last = new_para(doc, "left"); add_run(last, spec["fenhao"], fonts["body"])
        top_lines += 1
    for key in ("miji", "jinji"):
        if spec.get(key):
            last = new_para(doc, "left"); add_run(last, spec[key], fonts["h1"])
            top_lines += 1

    if spec.get("red_header"):
        text, size = spec["red_header"], 40
        need = em_width(text) * size          # 按实际文字压缩；含占位符的，补齐后须重新导出
        scale = int(TEXT_W * 0.98 / need * 100) if need > TEXT_W * 0.98 else None
        # 发文机关标志上边缘距版心上缘 35mm ≈ 99 磅（扣除已占行与行内留白）
        before = max(0, 99.2 - top_lines * LINE - 8)
        last = new_para(doc, "center", line=52, space_before=before, snap=False)
        add_run(last, text, fonts["header"], size, True, RED, scale=max(scale or 100, 50))
    if spec.get("brief_header"):
        last = new_para(doc, "center", line=52, space_after=6, snap=False)
        add_run(last, spec["brief_header"], fonts["header"], 36, True, RED)
    if spec.get("issue_no"):
        last = new_para(doc, "center"); add_run(last, spec["issue_no"], fonts["body"])
    if spec.get("issuer_left") or spec.get("issuer_right"):
        last = new_para(doc, "left")
        set_ind(last, left=1)
        last.paragraph_format.tab_stops.add_tab_stop(
            Pt(TEXT_W - CHAR_PITCH), WD_TAB_ALIGNMENT.RIGHT)
        add_run(last, spec.get("issuer_left", "") + "\t" + spec.get("issuer_right", ""),
                fonts["body"])

    signer = (spec.get("signer") or "").replace("签发人：", "").replace("签发人:", "").strip()
    if spec.get("wenhao") or signer:
        if spec.get("red_header"):
            blank(doc, 2)                     # 发文字号在发文机关标志下空二行
        if signer:
            last = new_para(doc, "left")
            set_ind(last, left=1)
            last.paragraph_format.tab_stops.add_tab_stop(
                Pt(TEXT_W - CHAR_PITCH), WD_TAB_ALIGNMENT.RIGHT)
            add_run(last, spec.get("wenhao", "") + "\t签发人：", fonts["body"])
            add_run(last, signer, "楷体_GB2312")   # 签发人姓名用楷体
        else:
            last = new_para(doc, "center"); add_run(last, spec["wenhao"], fonts["body"])

    has_header = bool(spec.get("red_header") or spec.get("brief_header"))
    if spec.get("red_line", has_header) and last is not None:
        bottom_border(last)                   # 发文字号之下 4mm 的红色分隔线
    return last is not None


def render_blocks(doc, blocks, fonts):
    for blk in blocks:
        unknown = set(blk) - KNOWN_BLOCK_KEYS
        if unknown:
            warn("unknown body-block key(s) %s ignored" % sorted(unknown))
        if blk.get("pagebreak"):
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
            continue
        if "plain" in blk:
            size = blk.get("size", BODY)
            p = new_para(doc, blk.get("align"), line=blk.get("line", LINE),
                         first_indent=None if blk.get("no_indent") else 2)
            color = RED if blk.get("color") == "red" else None
            add_run(p, blk["plain"], blk.get("font", fonts["body"]),
                    size, blk.get("bold", False), color)
            continue
        head_key = next((k for k in ("h1", "h2", "h3") if k in blk), None)
        if head_key:
            p = new_para(doc, "justify", first_indent=2)
            add_run(p, blk[head_key], fonts[head_key], bold=(head_key == "h3"))
            if blk.get("text"):
                add_run(p, blk["text"], fonts["body"])
        elif "text" in blk:
            p = new_para(doc, "justify", first_indent=2)
            add_run(p, blk["text"], fonts["body"])


def render_title(doc, text, fonts):
    p = new_para(doc, "center", line=36, snap=False)
    add_run(p, text, fonts["title"], 22, True)


def render_fujian(doc, items, fonts):
    """附件说明：正文下空一行、左空二字；回行与附件名称首字对齐。"""
    if isinstance(items, str):
        items = [items]
    items = [i.rstrip("。；;，,") for i in (items or []) if i]
    if not items:
        return
    blank(doc)
    if len(items) == 1:
        p = new_para(doc, "justify"); set_ind(p, left=5, hanging=3)
        add_run(p, "附件：" + items[0], fonts["body"])
        return
    for n, name in enumerate(items, 1):
        p = new_para(doc, "justify")
        if n == 1:
            set_ind(p, left=5.75, hanging=3.75)
            add_run(p, "附件：%d.%s" % (n, name), fonts["body"])
        else:
            set_ind(p, left=5.75, hanging=0.75)
            add_run(p, "%d.%s" % (n, name), fonts["body"])


def render_attendees(doc, spec, fonts):
    rows = [(k, spec[k]) for k in ("chuxi", "liexi", "qingjia") if spec.get(k)]
    label = {"chuxi": "出席", "liexi": "列席", "qingjia": "请假"}
    for i, (k, names) in enumerate(rows):
        if i == 0:
            blank(doc)
        p = new_para(doc, "justify"); set_ind(p, left=5, hanging=3)
        add_run(p, label[k] + "：", fonts["h1"])
        add_run(p, names, fonts["body"])


def render_signoff(doc, spec, fonts):
    """署名与成文日期（GB/T 9704-2012 7.3.5）。
    加盖印章：成文日期右空四字，署名以成文日期为准居中；正文下留两行给印章。
    不加盖印章：署名右空二字，成文日期首字比署名首字右移二字（日期更长时署名右移）。
    用右缩进而非尾部空格定位——右对齐段落的尾部空格 Word 不计宽度。"""
    name, date = spec.get("signoff_name"), spec.get("signoff_date")
    if not (name or date):
        return
    nw, dw = em_width(name or ""), em_width(date or "")
    # 正文末段与署名、日期同页（印章不得孤悬无正文的页面）
    if doc.paragraphs:
        doc.paragraphs[-1].paragraph_format.keep_with_next = True
    seal = spec.get("seal", True)
    for _ in range(2 if seal else 1):
        new_para(doc).paragraph_format.keep_with_next = True
    if seal:
        pd = 4
        pn = pd + (dw - nw) / 2 if date else 4
    else:
        pn, pd = 2, nw - dw
        if date and pd < 2:
            pd, pn = 2, dw + 4 - nw
    if name:
        p = new_para(doc, "right"); set_ind(p, right=max(pn, 0))
        p.paragraph_format.keep_with_next = bool(date)
        add_run(p, name, fonts["body"])
    if date:
        p = new_para(doc, "right"); set_ind(p, right=max(pd, 0))
        add_run(p, date, fonts["body"])


def render_banji(doc, spec, fonts):
    """版记：单格浮动表格锚定在版心底部，首末条粗线（表格上下框）、中间细线（印发行
    段落上框）（GB/T 9704-2012 7.4）。只用一行且禁止跨页拆分——放不下时整块移到
    下一页，不会出现抄送与印发分居两页。"""
    cc_key = "chaosong" if spec.get("chaosong") else ("fensong" if spec.get("fensong") else None)
    has_yinfa = bool(spec.get("yinfa_org") or spec.get("yinfa_date"))
    rows = ([cc_key] if cc_key else []) + (["yinfa"] if has_yinfa else [])
    if not rows:
        return
    w_twips = str(int(round(TEXT_W * 20)))
    tbl = doc.add_table(rows=1, cols=1)
    t = tbl._tbl
    tblPr = t.tblPr
    for child in list(tblPr):
        tblPr.remove(child)

    def el(tag, **attrs):
        e = OxmlElement(tag)
        for k, v in attrs.items():
            e.set(qn('w:' + k), v)
        return e
    tblPr.append(el('w:tblpPr', vertAnchor='margin', tblpYSpec='bottom',
                    horzAnchor='margin', tblpXSpec='center'))
    tblPr.append(el('w:tblW', w=w_twips, type='dxa'))
    borders = el('w:tblBorders')
    for side, sz in (("top", "8"), ("left", None), ("bottom", "8"), ("right", None),
                     ("insideH", None), ("insideV", None)):
        borders.append(el('w:' + side, val='single', sz=sz, space='0', color='000000')
                       if sz else el('w:' + side, val='nil'))
    tblPr.append(borders)
    tblPr.append(el('w:tblLayout', type='fixed'))
    mar = el('w:tblCellMar')
    for side in ("left", "right"):
        mar.append(el('w:' + side, w='0', type='dxa'))
    tblPr.append(mar)
    for gc in t.tblGrid.findall(qn('w:gridCol')):
        gc.set(qn('w:w'), w_twips)
    tbl.rows[0]._tr.get_or_add_trPr().append(el('w:cantSplit'))

    cell = tbl.rows[0].cells[0]
    for i, kind in enumerate(rows):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.line_spacing = Pt(LINE)
        if kind in ("chaosong", "fensong"):
            label = "抄送" if kind == "chaosong" else "分送"
            names = spec[kind].rstrip("，,；;")
            if not names.endswith("。"):
                names += "。"
            set_ind(p, left=4, right=1, hanging=3, unit=14)   # 左右各空一字，回行与冒号后首字对齐
            add_run(p, label + "：" + names, fonts["body"], 14)
        else:
            date = spec.get("yinfa_date", "")
            if date and not date.endswith("印发"):
                date += "印发"
            if i > 0:                         # 中间细线：段落上框，段落不设缩进才能通栏
                pPr = p._p.get_or_add_pPr()
                pbdr = el('w:pBdr')
                pbdr.append(el('w:top', val='single', sz='6', space='0', color='000000'))
                pPr.append(pbdr)
            stops = p.paragraph_format.tab_stops   # 用制表位代替缩进实现左右各空一字
            stops.add_tab_stop(Pt(14), WD_TAB_ALIGNMENT.LEFT)
            stops.add_tab_stop(Pt(TEXT_W - 14), WD_TAB_ALIGNMENT.RIGHT)
            add_run(p, "\t" + spec.get("yinfa_org", "") + "\t" + date, fonts["body"], 14)
    tail = new_para(doc, line=1)              # 表格后须有段落；压到 1 磅避免多出一页
    add_run(tail, "", fonts["body"], 1)


def build(spec):
    fonts = dict(DEFAULT_FONTS); fonts.update(spec.get("fonts", {}))
    global LATIN_FONT
    LATIN_FONT = fonts["latin"]
    unknown = set(spec) - KNOWN_KEYS
    if unknown:
        warn("unknown key(s) %s ignored -- see the JSON contract in this script's docstring" % sorted(unknown))

    doc = Document()
    setup_page(doc, fonts, spec.get("page_number", True))
    has_header = build_header(doc, spec, fonts)

    if spec.get("title"):
        if has_header:
            blank(doc, 2)                     # 标题在红色分隔线下空二行
        render_title(doc, spec["title"], fonts)
    if spec.get("zhusong"):
        blank(doc)                            # 主送机关在标题下空一行
        p = new_para(doc, "justify")
        add_run(p, spec["zhusong"], fonts["body"])
    elif spec.get("title") or has_header:
        blank(doc)                            # 无主送（纪要、简报等）时正文前空一行

    render_blocks(doc, spec.get("body", []), fonts)
    render_fujian(doc, spec.get("fujian"), fonts)
    render_attendees(doc, spec, fonts)
    render_signoff(doc, spec, fonts)

    if spec.get("fuzhu"):
        p = new_para(doc, "justify", first_indent=2)
        add_run(p, spec["fuzhu"], fonts["body"])

    for att in spec.get("attachments", []):
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        p = new_para(doc, "left")             # "附件"及序号：黑体，版心左上第一行顶格
        add_run(p, att.get("label", "附件"), fonts["h1"])
        blank(doc)
        if att.get("title"):                  # 附件标题居中于第三行
            render_title(doc, att["title"], fonts)
            blank(doc)
        render_blocks(doc, att.get("body", []), fonts)

    render_banji(doc, spec, fonts)
    return doc


DEMO = {
    "output": "outputs/示例-请示.docx",
    "red_header": "××县教育局文件",
    "wenhao": "×教〔2026〕1号",
    "signer": "×××",
    "title": "××县教育局关于追加2026年\n学校食堂改造经费的请示",
    "zhusong": "县人民政府：",
    "body": [
        {"text": "为改善全县中小学食堂就餐条件，根据《××》，我局……。现将有关情况请示如下。"},
        {"h1": "一、项目进展情况。", "text": "……"},
        {"h1": "二、资金缺口情况。", "text": "……"},
        {"text": "为此，特请示追加经费……万元。"},
        {"text": "妥否，请批复。"}
    ],
    "fujian": ["2026年学校食堂改造项目经费测算表"],
    "signoff_name": "××县教育局",
    "signoff_date": "2026年×月×日",
    "fuzhu": "（联系人：×××　电话：×××）",
    "chaosong": "县财政局",
    "yinfa_org": "××县教育局办公室",
    "yinfa_date": "2026年×月×日"
}


def main():
    ap = argparse.ArgumentParser(description="通用公文 → Word 导出器")
    ap.add_argument("spec", nargs="?", help="JSON 规格文件路径")
    ap.add_argument("-o", "--out", help="覆盖输出路径")
    ap.add_argument("--demo", metavar="PATH", help="写一份示例 JSON 到 PATH")
    args = ap.parse_args()

    if args.demo:
        with open(args.demo, "w", encoding="utf-8") as f:
            json.dump(DEMO, f, ensure_ascii=False, indent=2)
        print("DEMO written:", args.demo); return

    if not args.spec:
        ap.error("需要提供 JSON 规格文件（或用 --demo 生成示例）")

    with open(args.spec, "r", encoding="utf-8-sig") as f:   # 兼容 PowerShell 写出的带 BOM 文件
        spec = json.load(f)
    out = args.out or spec.get("output")
    if not out:
        ap.error("未指定输出路径（JSON 的 output 或 -o）")

    d = os.path.dirname(os.path.abspath(out))
    if d and not os.path.isdir(d):
        os.makedirs(d, exist_ok=True)
    build(spec).save(out)
    print("SAVED:", out)


if __name__ == "__main__":
    main()
