# -*- coding: utf-8 -*-
"""
任意の pptx に「開発者アドイン(レジストリ登録済み)のタスクペインを自動で開く」参照を埋め込む。
office-addin-debugging が生成するランチャー pptx と同じ仕組み(ppt/webextensions/*)を移植する。

  python scripts/inject-addin.py <src.pptx> <out.pptx> [launcher.pptx]

launcher.pptx 省略時は %TEMP%\\PowerPoint add-in <manifest id>.pptx を使う(npm run sideload 実行後に存在)。
用途: 自分のデッキで Flash Slide をテストする / 開発中に既存デッキへ挿入検証する。
"""
import os, re, sys, zipfile, shutil, tempfile

sys.stdout.reconfigure(encoding="utf-8")
MANIFEST_ID = "f8e5e78d-2ff1-4cf2-b9d3-b03e2363f9c0"
WE_PARTS = ["ppt/webextensions/webextension.xml", "ppt/webextensions/taskpanes.xml", "ppt/webextensions/_rels/taskpanes.xml.rels"]
REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    src, out = sys.argv[1], sys.argv[2]
    launcher = sys.argv[3] if len(sys.argv) > 3 else os.path.join(os.environ.get("TEMP", ""), f"PowerPoint add-in {MANIFEST_ID}.pptx")
    if not os.path.exists(launcher):
        sys.exit("launcher not found: " + launcher + "  (先に npm run sideload を実行)")

    lz = zipfile.ZipFile(launcher)
    we = {n: lz.read(n) for n in WE_PARTS}
    lroot = lz.read("_rels/.rels").decode("utf-8")
    m = re.search(r'<Relationship[^>]*Type="' + re.escape(REL_TYPE) + r'"[^>]*/>', lroot)
    if not m:
        sys.exit("launcher has no webextensiontaskpanes relationship")
    rel_xml = m.group(0)
    lct = lz.read("[Content_Types].xml").decode("utf-8")
    overrides = re.findall(r'<Override PartName="/ppt/webextensions/[^"]+"[^>]*/>', lct)

    sz = zipfile.ZipFile(src)
    tmp = tempfile.mktemp(suffix=".pptx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as oz:
        for item in sz.infolist():
            data = sz.read(item.filename)
            if item.filename == "_rels/.rels":
                s = data.decode("utf-8")
                if REL_TYPE not in s:
                    s = s.replace("</Relationships>", rel_xml + "</Relationships>")
                data = s.encode("utf-8")
            elif item.filename == "[Content_Types].xml":
                s = data.decode("utf-8")
                for ov in overrides:
                    if ov not in s:
                        s = s.replace("</Types>", ov + "</Types>")
                data = s.encode("utf-8")
            elif item.filename in we:
                continue
            oz.writestr(item, data)
        for n, d in we.items():
            oz.writestr(n, d)
    shutil.move(tmp, out)
    print("injected ->", out)


if __name__ == "__main__":
    main()
