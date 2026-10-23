#!/usr/bin/env python3
"""~/dev/grader/ の grader.py と checkers/ を、Webアプリ(docs/py/)にコピーし、manifest.json を更新する。
使い方: webapp フォルダで  python3 publish.py   → そのあと、GitHubに反映(git add / commit / push、またはGitHub Desktop)。
内容が変わっていなければ version も変わらないので、余計な更新は起きない。"""
import hashlib, json, py_compile, shutil, sys, tempfile
from datetime import date
from pathlib import Path

here = Path(__file__).resolve().parent
src = here.parent                               # ローカル版が入っているフォルダ(~/dev/grader/)
py = here / "docs" / "py"

def die(msg):
    print("中止:", msg); sys.exit(1)

if not (src / "grader.py").exists(): die(f"{src}/grader.py が見つかりません。webapp フォルダを、grader.py と同じフォルダの中に置いてください。")
if not (py / "manifest.json").exists(): die(f"{py}/manifest.json が見つかりません。")

files = {"grader.py": src / "grader.py"}
plugins = {p.name: p for p in sorted((src / "checkers").glob("*.py")) if not p.name.startswith("_")} if (src / "checkers").is_dir() else {}

for name, p in {**files, **plugins}.items():    # 構文エラーのファイルを公開しないよう、先に確認する
    try: py_compile.compile(str(p), cfile=str(Path(tempfile.gettempdir()) / "chk.pyc"), doraise=True)
    except py_compile.PyCompileError as e: die(f"{name} に構文エラーがあります:\n{e}")

(py / "checkers").mkdir(exist_ok=True)
shutil.copy2(files["grader.py"], py / "grader.py")
for name, p in plugins.items(): shutil.copy2(p, py / "checkers" / name)
for old in (py / "checkers").glob("*.py"):      # ローカルから削除した章のファイルは、Web側からも消す
    if old.name not in plugins: old.unlink(); print("削除:", old.name)
manual = src / "manual.pdf"              # マニュアルのPDFがあれば、Webアプリにもコピーする
if manual.exists():
    shutil.copy2(manual, here / "docs" / "manual.pdf"); print("コピーしました: manual.pdf → docs/manual.pdf")
h = hashlib.sha1()
for name in ["grader.py", "webcore.py"]: h.update((py / name).read_bytes())
for name in plugins: h.update((py / "checkers" / name).read_bytes())
manifest = json.loads((py / "manifest.json").read_text(encoding="utf-8"))
digest = h.hexdigest()[:8]
if not str(manifest.get("version", "")).endswith(digest) or manifest.get("checkers") != list(plugins):
    manifest["version"] = f"{date.today():%Y-%m-%d}-{digest}"
manifest["checkers"] = list(plugins)
(py / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
print(f"コピーしました: grader.py、章のチェッカー {len(plugins)} 個 {list(plugins)}")
print(f"version: {manifest['version']}")
print("次に、GitHubに反映してください(git add -A && git commit -m \"更新\" && git push)。")
