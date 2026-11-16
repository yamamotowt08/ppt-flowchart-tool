"""フローチャート作成ツール（.pptm / .ppam）と取り込み用 .bas を作る。

使い方: python3 tools/build.py
"""
import io, os, re, sys, uuid, zipfile
from pptx import Presentation
from pptx.util import Emu, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vbabin import compress, cfb, _dir_stream, _encrypt, CP

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, 'FlowchartMaker_コード.txt')
MODULE = 'FlowchartMaker'

RIBBON = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="fcTab" label="フローチャート">
        <group id="fcGroup" label="フローチャート自動作成">
          <button id="fcMake" label="作成" size="large" imageMso="SmartArtInsert"
                  onAction="FC_RibbonMake"
                  screentip="フローチャートを作成"
                  supertip="1回目：スライドに入力欄を出します。手順を書いてもう一度押すと、フローチャートを描きます。何も選ばずに押すと、前回の入力欄が戻ります。"/>
          <button id="fcQuick" label="一行入力" size="large" imageMso="TextBoxInsert"
                  onAction="FC_RibbonQuick"
                  screentip="一行入力で作成"
                  supertip="「開始/申請する/終了」のように、手順を「/」で区切って入力して作ります。"/>
          <button id="fcDelete" label="削除" size="large" imageMso="Delete"
                  onAction="FC_RibbonDelete"
                  screentip="フローチャートを削除"
                  supertip="このスライドで自動作成したフローチャートだけを消します。"/>
          <button id="fcHelp" label="書き方" size="large" imageMso="Help"
                  onAction="FC_RibbonHelp"
                  screentip="書き方を表示"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
'''

CT_MAIN = {
    'pptm': 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
    'ppam': 'application/vnd.ms-powerpoint.addin.macroEnabled.main+xml',
}
REL_VBA = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'
REL_UI = 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'


def vba_project(src):
    """標準モジュール 1 つだけの vbaProject.bin（PowerPoint 用）。"""
    src = src.replace('\r\n', '\n').replace('\n', '\r\n')
    mods = [(MODULE, src, False)]
    pid = '{%s}' % str(uuid.uuid4()).upper()
    lines = ['ID="%s"' % pid, 'Module=%s' % MODULE,
             'Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
             'CMG="%s"' % _encrypt(b'\0\0\0\0', pid), 'DPB="%s"' % _encrypt(b'\0', pid),
             'GC="%s"' % _encrypt(b'\xff', pid), '',
             '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '',
             '[Workspace]', '%s=0, 0, 0, 0, C' % MODULE]
    project = ('\r\n'.join(lines) + '\r\n').encode(CP)
    wm = MODULE.encode(CP) + b'\0' + MODULE.encode('utf-16-le') + b'\0\0' + b'\0\0'
    vba = {'_VBA_PROJECT': bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]),
           'dir': compress(_dir_stream('VBAProject', mods)),
           MODULE: compress(src.encode(CP))}
    return cfb({'VBA': vba, 'PROJECT': project, 'PROJECTwm': wm})


def widen_to_16x9(prs):
    """標準テンプレート（4:3）を 16:9 にして、マスターとレイアウトの枠も横に広げる。"""
    k = 12192000 / prs.slide_width
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    for part in [prs.slide_master] + list(prs.slide_layouts):
        for shp in part.shapes:
            # 自分で位置を持つ枠だけ広げる（マスターから受け継ぐ枠を二重に広げない）
            off = shp._element.xpath('./p:spPr/a:xfrm/a:off')
            ext = shp._element.xpath('./p:spPr/a:xfrm/a:ext')
            if off and ext:
                off[0].set('x', str(int(int(off[0].get('x')) * k)))
                ext[0].set('cx', str(int(int(ext[0].get('cx')) * k)))


def add_text_slide(prs, title, lines, size=24):
    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = title
    tf = s.placeholders[1].text_frame
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = Pt(size)


def make_pptx(with_slides):
    prs = Presentation()
    widen_to_16x9(prs)
    if with_slides:
        s = prs.slides.add_slide(prs.slide_layouts[0])
        s.shapes.title.text = 'フローチャート自動作成ツール'
        s.placeholders[1].text = '上のリボンの［フローチャート］タブ →［作成］で使えます'
        add_text_slide(prs, '使い方', [
            'フローチャートを入れたいスライドを開く',
            '［フローチャート］タブ →［作成］→ 黄色い入力欄が出る',
            '入力欄に手順を 1 行に 1 つずつ書く',
            'もう一度［作成］→ 入力欄が消えてフローチャートが描かれる',
            '直すときは、何も選ばずに［作成］→ 前回の入力欄が戻る',
        ])
        add_text_slide(prs, '書き方', [
            '開始 / 終了 → 角丸　　「？」で終わる行 → ひし形　　その他 → 四角',
            '判断？ > 処理 →「いいえ」のとき右に処理（そのあと次の手順に合流）',
            '判断？ > 処理 > #2 → 処理のあと 2 番目の手順へ戻る',
            '判断？ > 終了 →「いいえ」のとき右に終了',
            '※ で始まる行 → メモ（図に出ません）',
        ], size=20)
        s = prs.slides.add_slide(prs.slide_layouts[5])
        s.shapes.title.text = 'ここで試せます'
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def package(pptx_bytes, kind, vba_bin):
    zin = zipfile.ZipFile(io.BytesIO(pptx_bytes))
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == '[Content_Types].xml':
                t = data.decode('utf-8')
                t = re.sub(r'(PartName="/ppt/presentation.xml" ContentType=")[^"]+"', r'\g<1>%s"' % CT_MAIN[kind], t)
                t = t.replace('<Default Extension="xml"',
                              '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
                              '<Default Extension="xml"', 1)
                assert CT_MAIN[kind] in t and 'vbaProject' in t
                data = t.encode('utf-8')
            elif item.filename == 'ppt/_rels/presentation.xml.rels':
                t = data.decode('utf-8').replace(
                    '</Relationships>',
                    '<Relationship Id="rIdVba1" Type="%s" Target="vbaProject.bin"/></Relationships>' % REL_VBA)
                data = t.encode('utf-8')
            elif item.filename == '_rels/.rels':
                t = data.decode('utf-8').replace(
                    '</Relationships>',
                    '<Relationship Id="rIdUi1" Type="%s" Target="customUI/customUI.xml"/></Relationships>' % REL_UI)
                data = t.encode('utf-8')
            z.writestr(item, data)
        z.writestr('ppt/vbaProject.bin', vba_bin)
        z.writestr('customUI/customUI.xml', RIBBON.encode('utf-8'))
    return out.getvalue()


def main():
    src = open(SRC, encoding='utf-8').read()
    # 取り込み用 .bas（Shift_JIS / CRLF）
    with open(os.path.join(ROOT, 'FlowchartMaker.bas'), 'wb') as f:
        f.write(src.replace('\r\n', '\n').replace('\n', '\r\n').encode(CP))
    vba_bin = vba_project(src)
    with open(os.path.join(ROOT, 'フローチャート作成ツール.pptm'), 'wb') as f:
        f.write(package(make_pptx(True), 'pptm', vba_bin))
    with open(os.path.join(ROOT, 'フローチャート作成アドイン.ppam'), 'wb') as f:
        f.write(package(make_pptx(False), 'ppam', vba_bin))
    print('OK')


if __name__ == '__main__':
    main()
