#!/usr/bin/env python3
"""只读体检一份 xlsx：结构、工作表映射、字数分布、长文案清单、残留统计。

用法：
    python inspect_workbook.py "原文件.xlsx" --out "体检报告.txt" \
        [--long 55] [--brands "rabbitgoo,Aevyri"] [--blacklist "甜点带,红海"]

只读，不修改任何文件。单元格一律用 ElementTree 解析，避免自闭合单元格
把后一格吞进同一个正则匹配（见 references/pitfalls.md P1/P2）。
"""
import argparse
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

DEFAULT_BLACKLIST = ("甜点带,甜点词层,红海,内卷,蓝海,天花板,腰部,卡位,截流,爬坡,护城河,"
                     "绝对入口,红利期,洗牌市场,一超多强,三峰一谷,倒金字塔,中高集中,"
                     "没戏,硬打,死路,打不动,白捡,吃长尾")
SYMBOLS = ['「', '」', '——', '→', '①', '②', '③']


def decode(t):
    t = re.sub(r'&#(\d+);', lambda m: chr(int(m.group(1))), t)
    return (t.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
             .replace('&quot;', '"').replace('&#10;', '\n'))


def sheet_map(z):
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rm = {e.get('Id'): e.get('Target') for e in rels}
    out = []
    for s in wb.find(NS + 'sheets'):
        target = rm[s.get(RID)].lstrip('/')
        if not target.startswith('xl/'):
            target = 'xl/' + target
        out.append((s.get('name'), target))
    return out


def read_sst(z):
    if 'xl/sharedStrings.xml' not in z.namelist():
        return [], 0, 0
    raw = z.read('xl/sharedStrings.xml').decode('utf-8')
    items = re.findall(r'<si>(.*?)</si>', raw, re.S)
    texts = [decode(''.join(re.findall(r'<t[^>]*>(.*?)</t>', x, re.S))) for x in items]
    c = re.search(r'\bcount="(\d+)"', raw)
    u = re.search(r'\buniqueCount="(\d+)"', raw)
    return texts, (int(c.group(1)) if c else 0), (int(u.group(1)) if u else 0)


def read_cells(z, path, sst):
    root = ET.fromstring(z.read(path))
    out = {}
    for c in root.iter(NS + 'c'):
        ref, typ = c.get('r'), c.get('t')
        if not ref:
            continue
        if typ == 's':
            v = c.find(NS + 'v')
            if v is None or not v.text or not v.text.isdigit():
                continue
            idx = int(v.text)
            if idx < len(sst):
                out[ref] = sst[idx]
        elif typ == 'inlineStr':
            node = c.find(NS + 'is')
            if node is not None:
                out[ref] = decode(''.join(x.text or '' for x in node.iter(NS + 't')))
    return out


def merge_count(z, path):
    root = ET.fromstring(z.read(path))
    mg = root.find(NS + 'mergeCells')
    return len(list(mg)) if mg is not None else 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsx', help='待体检的 xlsx 路径')
    ap.add_argument('--out', default=None, help='详细清单输出文件，不传则只打印摘要')
    ap.add_argument('--long', type=int, default=55, help='长文案阈值，默认 55 字')
    ap.add_argument('--brands', default='', help='逗号分隔的品牌词')
    ap.add_argument('--blacklist', default=DEFAULT_BLACKLIST, help='逗号分隔的黑话词')
    args = ap.parse_args()

    if not os.path.isfile(args.xlsx):
        print('文件不存在:', args.xlsx)
        return 2

    z = zipfile.ZipFile(args.xlsx)
    names = z.namelist()
    media = [n for n in names if n.startswith('xl/media/') and not n.endswith('/')]
    drawings = [n for n in names if re.match(r'xl/drawings/drawing\d+\.xml$', n)]
    cellimages = [n for n in names if 'cellimages' in n]
    dispimg = sum(z.read(n).decode('utf-8', 'ignore').count('DISPIMG')
                  for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n))
    sst, sst_count, sst_uniq = read_sst(z)
    sheets = sheet_map(z)

    lines = []
    lines.append('文件: %s' % args.xlsx)
    lines.append('大小: %d bytes, zip 条目: %d' % (os.path.getsize(args.xlsx), len(names)))
    lines.append('media: %d, drawings: %d, cellimages: %d, DISPIMG: %d'
                 % (len(media), len(drawings), len(cellimages), dispimg))
    lines.append('sharedStrings: count=%d uniqueCount=%d si=%d'
                 % (sst_count, sst_uniq, len(sst)))
    lines.append('嵌图判定: %s' % ('必须用 XML 直改，禁止 openpyxl.save'
                                if (media or drawings or cellimages) else
                                '无嵌图，可用 openpyxl 保存'))
    lines.append('工作表映射:')
    for i, (nm, path) in enumerate(sheets, 1):
        lines.append('  %d. %s -> %s' % (i, nm, path))
    lines.append('')

    alltext = {}
    summary = []
    long_items = []
    for i, (nm, path) in enumerate(sheets, 1):
        cells = read_cells(z, path, sst)
        alltext[i] = cells
        chars = sum(len(v) for v in cells.values())
        longs = [(r, v) for r, v in cells.items() if len(v) >= args.long]
        badsym = sum(v.count(s) for v in cells.values() for s in SYMBOLS)
        summary.append((nm, len(cells), chars, len(longs), badsym, merge_count(z, path)))
        for r, v in longs:
            if not v.startswith('http'):
                long_items.append((i, nm, r, len(v), v))

    lines.append('%-30s %7s %9s %8s %8s %7s' %
                 ('SHEET', 'cells', 'chars', 'long', 'badsym', 'merges'))
    for s in summary:
        lines.append('%-30s %7d %9d %8d %8d %7d' % s)
    lines.append('TOTAL chars: %d' % sum(s[2] for s in summary))
    lines.append('')

    blob = '\n'.join(v for cells in alltext.values() for v in cells.values())
    brands = [b.strip() for b in args.brands.split(',') if b.strip()]
    blacks = [b.strip() for b in args.blacklist.split(',') if b.strip()]
    lines.append('品牌残留: %s' % ({b: blob.count(b) for b in brands if b in blob} or '无'))
    lines.append('黑话残留: %s' % ({b: blob.count(b) for b in blacks if b in blob} or '无'))
    lines.append('禁用符号: %s' % ({s: blob.count(s) for s in SYMBOLS if s in blob} or '无'))
    lines.append('')

    lines.append('=== 长文案清单（>= %d 字，已排除 URL）===' % args.long)
    for i, nm, r, ln, v in long_items:
        lines.append('sheet%d %s!%-5s len=%-4d | %s'
                     % (i, nm, r, ln, v.replace('\n', ' / ')[:120]))

    # 提示可能是原始数据表的 Sheet
    big = [(nm, chars) for nm, _, chars, _, _, _ in summary if chars > 20000]
    if big:
        lines.append('')
        lines.append('疑似原始数据表（字数大，建议整表不动）: %s'
                     % ', '.join('%s(%d字)' % x for x in big))

    text = '\n'.join(lines)
    print('\n'.join(lines[:len(sheets) + 12]))
    if args.out:
        with open(args.out, 'w', encoding='utf-8') as f:
            f.write(text)
        print('\n详细清单已写入: %s' % args.out)
    z.close()
    return 0


if __name__ == '__main__':
    sys.exit(main())
