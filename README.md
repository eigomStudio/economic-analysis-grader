# Eco-Analysis Grader(経済分析入門 章末課題 採点ツール・ブラウザ版)

学生の提出ファイル(Excel)を、**ブラウザの中だけで**採点します。ファイルはサーバーにアップロードされず、このパソコンの外には出ません。

## 構成
- `docs/` … 公開するサイトの中身(GitHub Pagesの公開元)
  - `index.html` / `app.js` … 画面
  - `worker.js` / `core.js` … 採点を動かす部分(Pyodide = ブラウザ上のPython)
  - `pyodide/` … Pyodide本体(同じサイトに置いてあるので、外部のサーバーを使いません)
  - `py/grader.py` … **採点の判定ルール(ローカル版の `grader.py` と同じファイル)**
  - `py/webcore.py` … ブラウザ版の入口と、結果のExcelの作成
  - `py/checkers/` … 第2章以降の判定ルール(章ごとに1ファイル)
  - `py/wheels/` … openpyxl など、同じサイトに置いたライブラリ
  - `py/manifest.json` … 読み込むファイルの一覧とバージョン

## 公開(GitHub Pages)
1. このフォルダの中身を、GitHubのリポジトリに置く。
2. リポジトリの Settings → Pages で、Source を「Deploy from a branch」、Branch を `main`、フォルダを `/docs` にして保存する。
3. 表示されたURL(`https://アカウント名.github.io/リポジトリ名/`)を、先生方に伝える。リポジトリ名は、たとえば `eco-analysis-grader`(英小文字とハイフン)がおすすめ。

※リポジトリが公開の場合、採点ルール(`grader.py`)も、誰でも読めます。学生のデータは、サーバーに存在しません。

## 更新(判定ルールを直したとき)
1. `docs/py/grader.py` を新しいものに差し替える(第2章以降は `docs/py/checkers/` にファイルを足す)。
2. `docs/py/manifest.json` の `version` を書き換える(新しい版を読み込ませるため)。第2章以降を足したときは、`checkers` の一覧に、ファイル名も書く。
3. GitHubにアップロードする。先生方は、ページを開き直すだけでよい。

## 自分のパソコンで試す
```bash
cd docs
python3 -m http.server 8000
```
ブラウザで `http://localhost:8000/` を開く。(`index.html` をダブルクリックで開くと、アドレスが `file://` で始まり、ブラウザの制限で動きません。その場合は、ページに理由が表示されます。)

画面の上のアドレスが `http://localhost:8000/` または `https://` で始まっていることを確認してください。

## 「アップロード」と表示されることについて
フォルダやファイルを選ぶとき、ブラウザが「アップロード」という文言の確認画面を出すことがあります。ブラウザがファイル選択で必ず使う表現で、このページでは、ファイルはどこにも送信されません(ページの上にも説明を載せています)。

## 動作環境と注意
- 最新のChrome、Edge、Safari、Firefox。初回は約14MBを読み込む(2回目以降はブラウザが覚えている)。
- 「ファイルが外に出ていない」ことは、ブラウザの開発者ツール(F12)の「ネットワーク」で、採点中に通信が増えないことで確認できる。
- Pyodide(MPL-2.0)、openpyxl(MIT)を同梱している。

## Macでの運用(~/dev/grader/webapp にこのフォルダを置く場合)
```
~/dev/grader/
  grader.py            ← ローカル版(判定ルールの本体)
  checkers/            ← 第2章以降の判定ルール
  webapp/              ← このフォルダ。ここだけをGitHubと同期する
    docs/  README.md  publish.py  .gitignore
```
- **GitHubと同期するのは `webapp` だけです。** `~/dev/grader/` 全体はGitHubに上げません(学生のファイルや採点結果が入る可能性があるため)。`webapp/.gitignore` が、xlsx / csv / out などの、誤った追加を防ぎます。
- 最初の1回:
  ```bash
  cd ~/dev/grader/webapp
  git init -b main
  git status            # docs/ README.md publish.py .gitignore だけが出ることを確認
  git add -A && git commit -m "first version"
  git remote add origin https://github.com/アカウント名/リポジトリ名.git
  git push -u origin main
  ```
  (先に、GitHubで、空のリポジトリを作っておく。README等は追加しない。そのあと Settings → Pages で `main` と `/docs` を指定する。)
- 判定ルールを直したあと:
  ```bash
  cd ~/dev/grader/webapp
  python3 publish.py                     # grader.py と checkers/ をdocs/py/にコピーし、versionを更新
  git add -A && git commit -m "更新" && git push
  ```
  1〜2分後に、公開ページに反映される。反映されないときは、ブラウザで強制再読み込み(Command+Shift+R)。
- `publish.py` は、コピー前に構文エラーを確認し、エラーがあれば中止する(壊れた版を公開しないため)。
