# -*- coding: utf-8 -*-
"""只读文档文本提取工具。

用法:
    python 读取文档.py <文件路径> [--max-chars N]

支持 .docx / .pptx / .xlsx / .txt / .md
仅以只读方式打开文件，绝不写入、修改或删除源文件。
"""
import sys
import os
import zipfile
import re
from xml.etree import ElementTree as ET

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
}


def _text_of(elem, tag):
    return ''.join(t.text or '' for t in elem.iter(tag))


def read_docx(path, max_chars):
    out = []
    with zipfile.ZipFile(path) as z:
        with z.open('word/document.xml') as f:
            root = ET.parse(f).getroot()
        for p in root.iter('{%s}p' % NS['w']):
            line = _text_of(p, '{%s}t' % NS['w']).strip()
            if line:
                out.append(line)
            if sum(len(x) for x in out) > max_chars:
                out.append('...[截断]')
                break
    return '\n'.join(out)


def read_pptx(path, max_chars):
    out = []
    with zipfile.ZipFile(path) as z:
        slides = sorted(
            (n for n in z.namelist()
             if re.match(r'ppt/slides/slide\d+\.xml$', n)),
            key=lambda n: int(re.search(r'(\d+)', n.split('/')[-1]).group(1)),
        )
        total = 0
        for name in slides:
            with z.open(name) as f:
                root = ET.parse(f).getroot()
            texts = [_text_of(p, '{%s}t' % NS['a']).strip()
                     for p in root.iter('{%s}p' % NS['a'])]
            texts = [t for t in texts if t]
            if texts:
                num = re.search(r'(\d+)', name.split('/')[-1]).group(1)
                out.append('--- 第%s页 ---' % num)
                out.extend(texts)
                total += sum(len(t) for t in texts)
            if total > max_chars:
                out.append('...[截断]')
                break
    return '\n'.join(out)


def read_xlsx(path, max_chars, max_rows=30, max_cols=12):
    """读取 xlsx：列出工作表名，并预览每表左上角若干行。"""
    out = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        shared = []
        if 'xl/sharedStrings.xml' in names:
            with z.open('xl/sharedStrings.xml') as f:
                root = ET.parse(f).getroot()
            for si in root:
                shared.append(_text_of(si, '{%s}t' % NS['ss']))
        # 工作表名
        wb = 'xl/workbook.xml'
        sheet_names = []
        if wb in names:
            with z.open(wb) as f:
                root = ET.parse(f).getroot()
            sheet_names = [s.get('name') for s in root.iter(
                '{%s}sheet' % NS['ss'])]
        out.append('工作表: ' + ' | '.join(sheet_names))
        sheets = sorted(n for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n))
        for name in sheets[:4]:
            with z.open(name) as f:
                root = ET.parse(f).getroot()
            out.append('--- %s ---' % name)
            rows_done = 0
            for row in root.iter('{%s}row' % NS['ss']):
                cells = []
                for c in row.iter('{%s}c' % NS['ss']):
                    v = c.find('{%s}v' % NS['ss'])
                    t = c.get('t')
                    if v is None or v.text is None:
                        cells.append('')
                    elif t == 's':
                        idx = int(v.text)
                        cells.append(shared[idx] if idx < len(shared) else '')
                    else:
                        cells.append(v.text)
                cells = cells[:max_cols]
                while cells and cells[-1] == '':
                    cells.pop()
                if cells:
                    out.append(' | '.join(cells))
                rows_done += 1
                if rows_done >= max_rows:
                    out.append('...[仅预览前%d行]' % max_rows)
                    break
            if sum(len(x) for x in out) > max_chars:
                break
    return '\n'.join(out)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    path = sys.argv[1]
    max_chars = 4000
    if '--max-chars' in sys.argv:
        max_chars = int(sys.argv[sys.argv.index('--max-chars') + 1])
    if not os.path.isfile(path):
        print('文件不存在: %s' % path)
        return
    ext = os.path.splitext(path)[1].lower()
    print('=' * 60)
    print('文件: %s' % path)
    print('=' * 60)
    try:
        if ext == '.docx':
            print(read_docx(path, max_chars))
        elif ext == '.pptx':
            print(read_pptx(path, max_chars))
        elif ext == '.xlsx':
            print(read_xlsx(path, max_chars))
        elif ext in ('.txt', '.md', '.py', '.csv'):
            with open(path, 'r', encoding='utf-8', errors='replace') as f:
                print(f.read(max_chars))
        else:
            print('暂不支持的类型: %s' % ext)
    except Exception as e:
        print('读取失败: %r' % (e,))


if __name__ == '__main__':
    main()
