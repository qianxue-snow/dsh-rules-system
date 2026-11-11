# -*- coding: utf-8 -*-
"""
md2docx.py —— 把本任务的 Markdown 文书转成 Word (.docx)

纯 Python 标准库实现：手写 OOXML（WordprocessingML）后用 zipfile 打包。
不依赖 Word COM、pandoc、python-docx。Python 3.8+ 可跑。

排版规则（中文公文/作业习惯）：
  #        文档大标题：二号 宋体 加粗 居中
  ## ——xxx 副标题  ：三号 加粗 居中
  ##/###   小节标题：三号 黑体 加粗 左对齐
  正文               ：三号 仿宋，固定行距 28 磅
  行首 1~2 个全角空格 → 首行缩进 2 字符；0 个 → 顶格
  行首 3~5 个全角空格 → 左缩进对应字数（附件条目）
  行首 6+ 个全角空格 → 右对齐（落款、署名）
  > 引用             ：五号 楷体 灰色 + 左侧竖线（“交付说明”框）
  | 表格 |           ：五号 宋体，全边框，表头加粗浅灰底
"""

import os
import re
import sys
import zipfile
from xml.sax.saxutils import escape

# ---------------------------------------------------------------- 页面与字号

PAGE_W = 11906          # A4 宽（twips）
MARGIN_L = 1701         # 左 3.0 cm
MARGIN_R = 1474         # 右 2.6 cm
CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R

SZ_H1 = 44              # 二号 22pt
SZ_H2 = 36              # 三号 16pt（副标题）
SZ_BODY = 32            # 三号 16pt
SZ_TBL = 21             # 五号 10.5pt
SZ_NOTE = 21            # 五号 10.5pt
LINE_EXACT = 560        # 固定行距 28 磅

FONT_BODY = '仿宋'
FONT_HEAD = '黑体'
FONT_TITLE = '宋体'
FONT_TBL = '宋体'
FONT_NOTE = '楷体'
FONT_ASCII = 'Times New Roman'

FW_SPACE = '\u3000'     # 全角空格


# ---------------------------------------------------------------- 基础工具

def esc(t):
    return escape(t, {'"': '&quot;'})


def rpr(font_east, bold=False, color=None, sz=SZ_BODY, ascii_font=FONT_ASCII):
    """按 OOXML 规定的元素顺序生成 w:rPr。"""
    s = ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
         % (ascii_font, ascii_font, font_east, ascii_font))
    if bold:
        s += '<w:b/><w:bCs/>'
    if color:
        s += '<w:color w:val="%s"/>' % color
    s += '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz)
    return s


def runs(text, base_rpr):
    """把 **粗体** 拆成多个 run。"""
    bold_rpr = re.sub(r'(<w:rFonts[^>]*/>)', r'\1<w:b/><w:bCs/>', base_rpr, count=1)
    out = []
    for i, part in enumerate(text.split('**')):
        if part == '':
            continue
        r = bold_rpr if i % 2 == 1 else base_rpr
        out.append('<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
                   % (r, esc(part)))
    if not out:
        out.append('<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve"></w:t></w:r>' % base_rpr)
    return ''.join(out)


def para(text, font=FONT_BODY, sz=SZ_BODY, bold=False, color=None, jc=None,
         first_line=None, left=None, right=None, line_exact=LINE_EXACT,
         before=0, after=0, note=False, space_auto=False):
    """生成一个 w:p。pPr 内部顺序：pBdr → shd → spacing → ind → jc → rPr。"""
    ppr = ''
    if note:
        ppr += ('<w:pBdr><w:left w:val="single" w:sz="18" w:space="6" '
                'w:color="BFBFBF"/></w:pBdr>'
                '<w:shd w:val="clear" w:color="auto" w:fill="FAFAFA"/>')
    if space_auto:
        ppr += '<w:spacing w:before="%d" w:after="%d" w:line="240" w:lineRule="auto"/>' % (before, after)
    else:
        ppr += '<w:spacing w:before="%d" w:after="%d" w:line="%d" w:lineRule="exact"/>' % (
            before, after, line_exact)
    ind = ''
    if first_line:
        ind += ' w:firstLine="%d"' % first_line
    if left:
        ind += ' w:left="%d"' % left
    if right:
        ind += ' w:right="%d"' % right
    if ind:
        ppr += '<w:ind%s/>' % ind
    if jc:
        ppr += '<w:jc w:val="%s"/>' % jc
    base = rpr(font, bold, color, sz)
    ppr += '<w:rPr>%s</w:rPr>' % base
    return '<w:p><w:pPr>%s</w:pPr>%s</w:p>' % (ppr, runs(text, base))


def cell(text, width, header=False):
    shd = ('<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' if header else '')
    tcpr = ('<w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s<w:vAlign w:val="center"/></w:tcPr>'
            % (width, shd))
    p = ('<w:p><w:pPr><w:spacing w:before="20" w:after="20" w:line="240" '
         'w:lineRule="auto"/><w:jc w:val="%s"/><w:rPr>%s</w:rPr></w:pPr>%s</w:p>'
         % ('center' if header else 'left',
            rpr(FONT_TBL, header, None, SZ_TBL),
            runs(text, rpr(FONT_TBL, header, None, SZ_TBL))))
    return '<w:tc>%s%s</w:tc>' % (tcpr, p)


def table(rows, has_header=True):
    ncol = max(len(r) for r in rows)
    rows = [r + [''] * (ncol - len(r)) for r in rows]

    weights = []
    for c in range(ncol):
        mx = max(len(r[c]) for r in rows)
        weights.append(max(mx, 4) ** 0.85)
    tot = sum(weights)
    widths = [max(int(CONTENT_W * w / tot), 700) for w in weights]
    widths[-1] += CONTENT_W - sum(widths)

    borders = ('<w:tblBorders>'
               + ''.join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="808080"/>' % side
                         for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
               + '</w:tblBorders>')
    tblpr = ('<w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:jc w:val="center"/>%s'
             '<w:tblLayout w:type="fixed"/>'
             '<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>'
             '<w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar>'
             '</w:tblPr>' % (CONTENT_W, borders))
    grid = '<w:tblGrid>' + ''.join('<w:gridCol w:w="%d"/>' % w for w in widths) + '</w:tblGrid>'

    trs = []
    for ri, row in enumerate(rows):
        head = has_header and ri == 0
        trpr = ('<w:trPr><w:trHeight w:val="380" w:hRule="atLeast"/>%s</w:trPr>'
                % ('<w:tblHeader/>' if head else ''))
        trs.append('<w:tr>%s%s</w:tr>'
                   % (trpr, ''.join(cell(row[c], widths[c], head) for c in range(ncol))))
    return '<w:tbl>%s%s%s</w:tbl>' % (tblpr, grid, ''.join(trs))


# ---------------------------------------------------------------- Markdown 解析

SEP_RE = re.compile(r'^:?-{2,}:?$')


def parse(md):
    lines = md.split('\n')
    blocks = []
    i = 0
    while i < len(lines):
        raw = lines[i].rstrip('\n').rstrip()
        s = raw.strip()

        if s == '' or s == '---' or s == '***':
            i += 1
            continue

        if s.startswith('>'):
            buf = []
            while i < len(lines) and lines[i].strip().startswith('>'):
                buf.append(lines[i].strip().lstrip('>').strip())
                i += 1
            blocks.append(('note', buf))
            continue

        if s.startswith('|'):
            buf = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                buf.append(lines[i].strip())
                i += 1
            has_header = False
            rows = []
            for k, b in enumerate(buf):
                cells = [c.strip() for c in b.strip('|').split('|')]
                if (cells and any(c != '' for c in cells)
                        and all(SEP_RE.match(c) for c in cells if c != '')):
                    has_header = (k == 1 and len(rows) == 1)
                    continue
                rows.append(cells)
            blocks.append(('table', rows, has_header))
            continue

        m = re.match(r'^(#{1,6})\s+(.*)$', raw)
        if m:
            blocks.append(('h', len(m.group(1)), m.group(2).strip()))
            i += 1
            continue

        blocks.append(('p', raw))
        i += 1
    return blocks


def count_fw_lead(s):
    n = 0
    while n < len(s) and s[n] == FW_SPACE:
        n += 1
    return n


def render(blocks):
    body = []
    prev_was_table = False
    for b in blocks:
        kind = b[0]

        if kind == 'table':
            rows, has_header = b[1], b[2]
            if prev_was_table:
                body.append(para('', sz=SZ_BODY))
            body.append(table(rows, has_header))
            body.append(para('', sz=SZ_BODY))     # 表后空段，避免两表相连
            prev_was_table = True
            continue

        prev_was_table = False

        if kind == 'note':
            for line in b[1]:
                if line:
                    body.append(para(line, font=FONT_NOTE, sz=SZ_NOTE, color='808080',
                                     note=True, before=20, after=20))
            continue

        if kind == 'h':
            level, text = b[1], b[2]
            if level == 1:
                body.append(para(text, font=FONT_TITLE, sz=SZ_H1, bold=True, jc='center',
                                 before=120, after=240))
            elif text.startswith('——') or text.startswith('--'):
                body.append(para(text, font=FONT_TITLE, sz=SZ_H2, bold=True, jc='center',
                                 before=0, after=240))
            else:
                body.append(para(text, font=FONT_HEAD, sz=SZ_H2, bold=True,
                                 first_line=640, before=160, after=80))
            continue

        # 普通段落：按行首全角空格数决定缩进/对齐
        n = count_fw_lead(b[1])
        text = b[1].lstrip(FW_SPACE) if n else b[1]
        if n == 0:
            body.append(para(text))
        elif n <= 2:
            body.append(para(text, first_line=640))
        elif n <= 5:
            body.append(para(text, left=n * 320))
        else:
            body.append(para(text, jc='right', right=640))
    return ''.join(body)


# ---------------------------------------------------------------- docx 打包

DOC_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s'
           '<w:sectPr><w:pgSz w:w="%d" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="%d" w:bottom="1440" w:left="%d" '
           'w:header="851" w:footer="992" w:gutter="0"/>'
           '<w:cols w:space="425"/></w:sectPr>'
           '</w:body></w:document>')

CONTENT_TYPES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                 '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                 '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                 '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
                 '</Types>')

ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
             '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
             '</Relationships>')

DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>')

STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="仿宋" w:cs="Times New Roman"/>'
          '<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:pPrDefault>'
          '</w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/><w:qFormat/>'
          '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>'
          '</w:styles>')


def core_props(title, author):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:creator>%s</dc:creator><cp:lastModifiedBy>%s</cp:lastModifiedBy>'
            '<dc:title>%s</dc:title>'
            '</cp:coreProperties>' % (esc(author), esc(author), esc(title)))


APP_PROPS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
             'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
             '<Application>md2docx.py</Application></Properties>')


def write_docx(path, body_xml, title, author):
    doc = DOC_XML % (body_xml, PAGE_W, MARGIN_R, MARGIN_L)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('docProps/core.xml', core_props(title, author))
        z.writestr('docProps/app.xml', APP_PROPS)
    return path


# ---------------------------------------------------------------- 主流程

# 署名：写进 docx 属性。用 --author 覆盖；留空则写文件标题。
AUTHOR = ''


def convert(src_path, dst_path, author=None):
    with open(src_path, 'r', encoding='utf-8') as f:
        md = f.read()
    blocks = parse(md)
    body_xml = render(blocks)
    title = ''
    for b in blocks:
        if b[0] == 'h' and b[1] == 1:
            title = b[2]
            break
    if not title:
        title = os.path.splitext(os.path.basename(src_path))[0]
    write_docx(dst_path, body_xml, title, author if author is not None else AUTHOR)
    return len(body_xml)


def main():
    args = [a for a in sys.argv[1:]]
    out_dir = '.'
    author = AUTHOR
    i = 0
    while i < len(args):
        a = args[i]
        if a in ('--author', '-a') and i + 1 < len(args):
            author = args[i + 1]
            i += 2
            continue
        if a in ('-h', '--help'):
            print('用法: python md2docx.py [输出目录] [--author "署名"]')
            print('  把目录下所有 .md 转成同名 .docx。不指定目录则用当前目录。')
            return 0
        out_dir = a
        i += 1

    targets = [f for f in sorted(os.listdir(out_dir)) if f.lower().endswith('.md')]
    if not targets:
        print('没有找到 .md 文件：%s' % out_dir)
        return 1
    for name in targets:
        src = os.path.join(out_dir, name)
        dst = os.path.join(out_dir, os.path.splitext(name)[0] + '.docx')
        size = convert(src, dst, author)
        print('OK  %-46s -> %s  (%.1f KB)' % (name, os.path.basename(dst),
                                              os.path.getsize(dst) / 1024.0))
    return 0


if __name__ == '__main__':
    sys.exit(main())


if __name__ == '__main__':
    sys.exit(main())
