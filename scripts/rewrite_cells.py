#!/usr/bin/env python3
"""通用 XML 直改引擎：按映射改写单元格文案，保留其余全部 zip 条目字节。

用法：
    python rewrite_cells.py --src "原文件.xlsx" --out "原文件_半精简版.xlsx" \
        --map rewrite_map.json [--backup "目录/_backup/改前备份.xlsx"]

映射 JSON 的键为工作表序号(字符串)或工作表名，值为坐标到新文案。
内置：目标格预检、按坐标精确定位（规避自闭合单元格吞并）、每格独立追加
sharedStrings 条目、count/uniqueCount 重算。预检不通过不会写出半成品。
"""
import argparse
import json
import os
import re
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from inspect_workbook import sheet_map, read_sst, decode  # noqa: E402


def esc(s):
    s = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return s.replace('\n', '&#10;')


def load_map(path, sheets):
    with open(path, 'r', encoding='utf-8-sig') as f:
        raw = json.load(f)
    by_index = {str(i): p for i, (_, p) in enumerate(sheets, 1)}
    by_name = {nm: p for nm, p in sheets}
    out = {}
    for key, cells in raw.items():
        if key.startswith('_'):
            continue
        target = by_index.get(key) or by_name.get(key)
        if not target:
            raise KeyError('映射里的表不存在: %s' % key)
        bucket = out.setdefault(target, {})
        for ref, text in cells.items():
            if not isinstance(text, str):
                raise TypeError('%s!%s 的新文案必须是字符串' % (key, ref))
            bucket[ref] = text
    return out


def cell_pattern(ref):
    return re.compile(r'(<c r="' + re.escape(ref) + r'"([^>]*?)>)(.*?)(</c>)', re.S)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--src', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--map', required=True)
    ap.add_argument('--backup', default=None)
    args = ap.parse_args()

    if not os.path.isfile(args.src):
        print('原文件不存在:', args.src)
        return 2
    if os.path.abspath(args.src) == os.path.abspath(args.out):
        print('输出路径不能与原文件相同；本技能要求另存新文件。')
        return 2

    zin = zipfile.ZipFile(args.src, 'r')
    names = set(zin.namelist())
    sheets = sheet_map(zin)
    mapping = load_map(args.map, sheets)

    has_sst = 'xl/sharedStrings.xml' in names
    sst_raw = zin.read('xl/sharedStrings.xml').decode('utf-8') if has_sst else ''
    sst_texts, sst_count, sst_uniq = read_sst(zin)
    new_si = []
    next_idx = sst_uniq
    xmls = {p: zin.read(p).decode('utf-8') for p in mapping}

    errors, hit, old_len, new_len = [], {}, 0, 0
    for path, cells in mapping.items():
        found = set()
        raw = xmls[path]
        for ref, text in cells.items():
            mo = cell_pattern(ref).search(raw)
            if not mo:
                errors.append((path, ref, '单元格不存在'))
                continue
            open_tag, attrs, body, close_tag = mo.groups()
            if attrs.rstrip().endswith('/'):
                errors.append((path, ref, '自闭合单元格'))
                continue
            if 't="s"' in attrs:
                if not has_sst:
                    errors.append((path, ref, '声明为共享字符串但文件无 sst'))
                    continue
                vm = re.search(r'<v>(\d+)</v>', body)
                if not vm:
                    errors.append((path, ref, '缺少 <v>'))
                    continue
                idx = int(vm.group(1))
                if idx < len(sst_texts):
                    old_len += len(sst_texts[idx])
                new_si.append('<si><t>' + esc(text) + '</t></si>')
                nb = body[:vm.start()] + '<v>%d</v>' % next_idx + body[vm.end():]
                next_idx += 1
            elif 't="inlineStr"' in attrs:
                im = re.search(r'<is>(.*?)</is>', body, re.S)
                if not im:
                    errors.append((path, ref, '缺少 <is>'))
                    continue
                old_txt = decode(''.join(re.findall(r'<t[^>]*>(.*?)</t>', im.group(1), re.S)))
                old_len += len(old_txt)
                nb = body[:im.start()] + '<is><t>' + esc(text) + '</t></is>' + body[im.end():]
            else:
                errors.append((path, ref, '不是文本单元格(t="s" 或 inlineStr)'))
                continue
            found.add(ref)
            new_len += len(text)
            raw = raw[:mo.start()] + open_tag + nb + close_tag + raw[mo.end():]
        xmls[path] = raw
        hit[path] = len(found)

    if errors:
        print('预检失败，未写出任何文件：')
        for e in errors:
            print('  ', e)
        zin.close()
        return 1

    if has_sst and new_si:
        updated = sst_raw.replace('uniqueCount="%d"' % sst_uniq,
                                  'uniqueCount="%d"' % next_idx, 1)
        updated = updated.rstrip()
        if not updated.endswith('</sst>'):
            print('sharedStrings 结尾异常')
            zin.close()
            return 1
        xmls['xl/sharedStrings.xml'] = (updated[:-len('</sst>')]
                                        + ''.join(new_si) + '</sst>')

    if args.backup:
        os.makedirs(os.path.dirname(os.path.abspath(args.backup)), exist_ok=True)
        if not os.path.exists(args.backup):
            shutil.copy2(args.src, args.backup)

    tmp = args.out + '.tmp'
    with zipfile.ZipFile(tmp, 'w') as zout:
        for item in zin.infolist():
            data = (xmls[item.filename].encode('utf-8')
                    if item.filename in xmls else zin.read(item.filename))
            zi = zipfile.ZipInfo(item.filename, item.date_time)
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zi.internal_attr = item.internal_attr
            zi.create_system = item.create_system
            zi.comment = item.comment
            zout.writestr(zi, data)
    zin.close()
    if os.path.exists(args.out):
        os.remove(args.out)
    os.replace(tmp, args.out)

    total = sum(hit.values())
    print('OK ->', args.out)
    print('改写格数: %d, 覆盖工作表: %d' % (total, len(hit)))
    for p, n in hit.items():
        print('  %s: %d 格' % (p, n))
    if has_sst:
        print('sharedStrings: count %d -> %d, uniqueCount %d -> %d'
              % (sst_count, sst_count, sst_uniq, next_idx))
    print('目标文案字数: %d -> %d (保留 %.1f%%)'
          % (old_len, new_len, 100.0 * new_len / old_len if old_len else 0))
    if args.backup:
        print('备份:', args.backup)
    return 0


if __name__ == '__main__':
    sys.exit(main())
