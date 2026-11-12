#!/usr/bin/env python3
"""校验改写结果：结构完整性、图片公式、sharedStrings、目标值与残留统计。

用法：
    python verify_rewrite.py --orig "原文件.xlsx" --new "半精简版.xlsx" \
        --map rewrite_map.json [--brands "rabbitgoo,Aevyri"] [--blacklist "甜点带,红海"]

任一硬性项为 FAIL 时退出码为 1，不要带病交付。
"""
import argparse
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from inspect_workbook import (sheet_map, read_sst, read_cells,  # noqa: E402
                              DEFAULT_BLACKLIST, SYMBOLS)
from rewrite_cells import load_map  # noqa: E402

results = []


def check(label, ok, detail=''):
    results.append((label, bool(ok), detail))
    print('[%s] %s%s' % ('PASS' if ok else 'FAIL', label,
                         (' | ' + str(detail)) if detail else ''))
    return bool(ok)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--orig', required=True)
    ap.add_argument('--new', required=True)
    ap.add_argument('--map', required=True)
    ap.add_argument('--brands', default='')
    ap.add_argument('--blacklist', default=DEFAULT_BLACKLIST)
    ap.add_argument('--skip-sheets', default='',
                    help='统计残留时排除的工作表名，逗号分隔，通常填原始数据表')
    args = ap.parse_args()

    if not all(os.path.isfile(p) for p in (args.orig, args.new, args.map)):
        print('文件缺失')
        return 2

    za, zb = zipfile.ZipFile(args.orig), zipfile.ZipFile(args.new)
    na, nb = za.namelist(), zb.namelist()
    sheets = sheet_map(za)
    mapping = load_map(args.map, sheets)

    check('zip 完整性', zb.testzip() is None)
    check('zip 条目名一致', na == nb, '%d -> %d' % (len(na), len(nb)))

    protect = [n for n in na if n.startswith('xl/media/') and not n.endswith('/')
               or re.match(r'xl/drawings/drawing\d+\.xml$', n) or 'cellimages' in n]
    changed = [n for n in protect if za.read(n) != zb.read(n)]
    check('图片与嵌图部件字节一致', not changed,
          '%d 个受保护条目, %d 个变化' % (len(protect), len(changed)))

    def dispimg(z, names):
        return sum(z.read(n).decode('utf-8', 'ignore').count('DISPIMG')
                   for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n))
    da, db = dispimg(za, na), dispimg(zb, nb)
    check('DISPIMG 公式数一致', da == db, '%d -> %d' % (da, db))

    sa, sb = read_sst(za), read_sst(zb)
    check('sharedStrings count 不变', sa[1] == sb[1], '%d -> %d' % (sa[1], sb[1]))
    si_a = za.read('xl/sharedStrings.xml').decode('utf-8').count('<si>') if 'xl/sharedStrings.xml' in na else 0
    si_b = zb.read('xl/sharedStrings.xml').decode('utf-8').count('<si>') if 'xl/sharedStrings.xml' in nb else 0
    # 统计映射里声明为共享字符串的格数
    need = 0
    for path, cells in mapping.items():
        rawa = za.read(path).decode('utf-8')
        for ref in cells:
            mo = re.search(r'<c r="' + re.escape(ref) + r'"([^>]*?)>', rawa)
            if mo and 't="s"' in mo.group(1):
                need += 1
    check('sharedStrings uniqueCount = 原值 + 改写格数',
          sb[2] == sa[2] + need, '%d -> %d (预期 +%d)' % (sa[2], sb[2], need))
    check('si 条目数与 uniqueCount 一致', si_b == sb[2], 'si=%d uniq=%d' % (si_b, sb[2]))

    untouched = [p for _, p in sheets if p not in mapping]
    same = [p for p in untouched if za.read(p) == zb.read(p)]
    check('未改写工作表字节一致', len(same) == len(untouched),
          '%d/%d' % (len(same), len(untouched)))

    ca = {p: read_cells(za, p, sa[0]) for _, p in sheets}
    cb = {p: read_cells(zb, p, sb[0]) for _, p in sheets}

    mismatch = []
    for path, cells in mapping.items():
        for ref, want in cells.items():
            got = cb.get(path, {}).get(ref)
            if got != want:
                mismatch.append((path, ref, str(got)[:40]))
    check('目标单元格值精确匹配映射', not mismatch,
          '%d 格, %d 处不符' % (sum(len(c) for c in mapping.values()), len(mismatch)))

    drift = []
    for path in cb:
        targets = mapping.get(path, {})
        for ref, val in ca.get(path, {}).items():
            if ref in targets:
                continue
            if cb[path].get(ref) != val:
                drift.append((path, ref))
    check('非目标文本单元格值不变', not drift,
          '%d 处意外变化' % len(drift))

    old_chars = sum(len(v) for cells in ca.values() for v in cells.values())
    new_chars = sum(len(v) for cells in cb.values() for v in cells.values())
    target_old = sum(len(ca.get(p, {}).get(r, '')) for p, cs in mapping.items() for r in cs)
    target_new = sum(len(v) for cs in mapping.values() for v in cs.values())
    print('\n--- 字数 ---')
    print('全文件: %d -> %d (%.2f%%)' % (old_chars, new_chars,
          100.0 * new_chars / old_chars if old_chars else 0))
    print('被改写文案: %d -> %d (%.1f%% 保留)' % (target_old, target_new,
          100.0 * target_new / target_old if target_old else 0))

    skip_names = {s.strip() for s in args.skip_sheets.split(',') if s.strip()}
    skip_paths = {p for nm, p in sheets if nm in skip_names}
    blob_all = '\n'.join(v for cells in cb.values() for v in cells.values())
    blob_analysis = '\n'.join(v for p, cells in cb.items() if p not in skip_paths
                              for v in cells.values() if not v.startswith('http'))
    brands = [b.strip() for b in args.brands.split(',') if b.strip()]
    blacks = [b.strip() for b in args.blacklist.split(',') if b.strip()]

    def tally(blob, words):
        return {w: blob.count(w) for w in words if w in blob} or '无'

    print('\n--- 残留 ---')
    print('口径: 全部文本 / 分析类（排除 %s 与商品 URL）'
          % (', '.join(skip_names) if skip_names else '无'))
    print('全部文本 | 品牌:', tally(blob_all, brands), '| 黑话:', tally(blob_all, blacks))
    print('分析类   | 品牌:', tally(blob_analysis, brands), '| 黑话:', tally(blob_analysis, blacks))
    print('禁用符号:', tally(blob_all, SYMBOLS))

    try:
        import openpyxl
        wb = openpyxl.load_workbook(args.new, read_only=True)
        n = len(wb.sheetnames)
        wb.close()
        check('openpyxl 只读打开', n == len(sheets), '%d 个 Sheet' % n)
    except Exception as e:
        check('openpyxl 只读打开', False, str(e)[:80])

    za.close()
    zb.close()
    failed = [r for r in results if not r[1]]
    print('\nSUMMARY: %d 项通过, %d 项失败' % (len(results) - len(failed), len(failed)))
    return 1 if failed else 0


if __name__ == '__main__':
    sys.exit(main())
