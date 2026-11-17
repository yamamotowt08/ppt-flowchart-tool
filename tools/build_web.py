"""HTML 版（フローチャート作成ツール.html）と、EXE 版の中に入れる本体（app/tool.html）を作る。

web/app.html の __PPTX_TEMPLATE__ に、python-pptx の標準テンプレート
（スライドマスター・レイアウト・テーマ）を埋め込む。

使い方: python3 tools/build_web.py [アーティファクト用の出力先]
"""
import io, json, os, sys, zipfile
from pptx import Presentation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, 'web', 'app.html')
OUT = os.path.join(ROOT, 'フローチャート作成ツール.html')   # ダブルクリックで開く HTML 版
APP_OUT = os.path.join(ROOT, 'app', 'tool.html')          # EXE 版の中に入れる本体（同じ中身）

SKIP = ('ppt/printerSettings/', 'docProps/thumbnail', 'ppt/slides/')


def template_parts():
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    buf = io.BytesIO()
    prs.save(buf)
    z = zipfile.ZipFile(buf)
    parts = {}
    for name in z.namelist():
        if name.startswith(SKIP):
            continue
        t = z.read(name).decode('utf-8')
        if name == '[Content_Types].xml':
            t = t.replace('<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.presentationml.printerSettings"/>', '')
            t = t.replace('<Default Extension="jpeg" ContentType="image/jpeg"/>', '')
        elif name == '_rels/.rels':
            t = t.replace('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/>', '')
        elif name == 'ppt/_rels/presentation.xml.rels':
            t = t.replace('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="printerSettings/printerSettings1.bin"/>', '')
        parts[name] = t
    assert 'printerSettings' not in parts['ppt/_rels/presentation.xml.rels']
    assert 'thumbnail' not in parts['_rels/.rels']
    assert 'slides/slide1.xml' in parts['ppt/_rels/presentation.xml.rels']
    return parts


def main():
    html = open(SRC, encoding='utf-8').read()
    data = json.dumps(template_parts(), ensure_ascii=False, separators=(',', ':'))
    data = data.replace('</', '<\\/')
    assert html.count('__PPTX_TEMPLATE__') == 1
    page = html.replace('__PPTX_TEMPLATE__', data)

    doc = ('<!doctype html>\n<html lang="ja">\n<head>\n<meta charset="utf-8">\n'
           '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">\n'
           '</head>\n<body>\n' + page + '\n</body>\n</html>\n')
    for out in (OUT, APP_OUT):
        with open(out, 'w', encoding='utf-8') as f:
            f.write(doc)
    # Artifact 用（骨組みは公開時に付く）
    if len(sys.argv) > 1:
        with open(sys.argv[1], 'w', encoding='utf-8') as f:
            f.write(page)
    print('OK', len(page))


if __name__ == '__main__':
    main()
