"""
manifest.template.xml から manifest.xml を作る。

  python make_manifest.py https://<組織名>.github.io/<リポジトリ名> [バージョン]

- アドインの ID は初回に addin_id.txt に作って保存し、以後は同じ ID を使う（ID が変わると別のアドイン扱いになる）
- 画面（site/）を直したら、バージョンを上げて manifest.xml を作り直し、管理センターで更新する
"""
import sys, uuid, pathlib
from urllib.parse import urlparse

here = pathlib.Path(__file__).parent
if len(sys.argv) < 2:
    sys.exit("使い方: python make_manifest.py https://<組織名>.github.io/<リポジトリ名> [バージョン]")
base = sys.argv[1].rstrip("/")
version = sys.argv[2] if len(sys.argv) > 2 else "1.0.0.0"
u = urlparse(base)
if u.scheme != "https":
    sys.exit("URL は https で始めてください")
idf = here / "addin_id.txt"
if not idf.exists():
    idf.write_text(str(uuid.uuid4()) + "\n")
addin_id = idf.read_text().strip()
xml = (here / "manifest.template.xml").read_text()
xml = (xml.replace("{{BASE}}", base).replace("{{ORIGIN}}", f"{u.scheme}://{u.netloc}")
          .replace("{{ID}}", addin_id).replace("{{VERSION}}", version))
xml = xml.replace("<!-- make_manifest.py で {{...}} を置き換えて manifest.xml を作る。直接編集しない -->\n", "")
(here / "manifest.xml").write_text(xml)
print(f"manifest.xml を作りました（ID {addin_id}、バージョン {version}、{base}）")
