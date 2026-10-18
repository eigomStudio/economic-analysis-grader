"""章末課題の診断・採点補助(Eco-Analysis Grader)  バージョン: 2026-10-18b(ローカル専用・外部通信なし)。
使い方: python3 grader.py          → フォルダ選択ダイアログが開く(採点 → 採点結果.xlsx)
        python3 grader.py --export  → 確認・修正済みの採点結果.xlsx から、Canvas取り込み用CSVを作る
        python3 grader.py 提出フォルダ [--ref 模範解答フォルダ] [--out 出力先]   → 引数で直接指定
ファイル名「学生番号－氏名－課題N.xlsx」から章Nを判定し、対応するチェッカーで処理する。"""
import argparse, csv, importlib.util, io, json, os, re, subprocess, time, types, unicodedata, urllib.error, urllib.parse, urllib.request, zipfile
from collections import Counter
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl

# ───────── 共通部品 ─────────
GRADER_VERSION = "2026-10-18b"                  # 判定ルールを直したら、ここを新しい日付にする(結果のExcelの「集計」シートにも出る)
FINAL = "最終評価(修正可)"
COMMENT_COL = "学生向けコメント(修正可)"
DEFAULT_POINTS = {"S": 4, "A": 3, "B": 2, "C": 1, "D": 0}   # Canvas点数。「Canvas対応表」シートで変更できる
CHECKERS = {}                                   # 章番号 -> 関数。新しい章は @checker(N) で追加
LABELS = {}                                     # できていない箇所のコード -> 説明
COMMENTS = {}                                   # できていない箇所のコード -> 学生向けのコメント(1〜2文)。章ごとのファイルから足せる

def make_comment(codes, grade):
    """できていない箇所のコードから、学生向けコメントの下書きを作る(先生が確認・修正する前提)。"""
    if grade in ("", "評価外", "未対応"): return ""
    lines = []
    for c in codes:                              # 「コード@詳細」の形なら、コメント中の {items} に詳細を入れる(例: どの項目の書式か)
        base, _, items = c.partition("@")
        if base in COMMENTS:
            t = COMMENTS[base]
            lines.append(t.replace("{items}", items.replace(",", "・")) if items else re.sub(r"\{items\}[のに]?", "", t))
    notes = [x for x in lines if x.startswith("※")]; lines = [x for x in lines if not x.startswith("※")]   # 「※」で始まるもの(遅延など)は最後に添える
    head = "次の点を確認してください。\n" + "\n".join("・" + x for x in lines) if lines else "よくできています。"
    return "\n".join([head] + notes)

def label_of(c):
    base, _, items = c.partition("@")
    return LABELS.get(base, base) + (f"({items.replace(',', '・')})" if items else "")

def chapter_from_path(p):
    """ファイル名に章がないとき、フォルダ名(「課題2」「第2章」など)から章を読む。"""
    for part in reversed(Path(p).parts[:-1]):
        ch = parse_name(part)[2]
        if ch: return ch
    return None

def is_late(p):
    """遅延の提出か: Canvasの一括ダウンロードのファイル名の「late」、または先生が仕分けたフォルダ名の「遅延」。"""
    p = Path(p)
    return bool(re.search(r"(?i)(?:^|_)late(?:_|$)", p.stem)) or any("遅延" in part for part in p.parts[:-1])

GRADES_SD = ["S", "A", "B", "C", "D"]
LABELS_LATE = "提出期限後の提出(評価を1段階下げた)"

def apply_file_late(p, g, iss):
    """遅延の提出なら、評価を1段階下げ(Dはそのまま)、コメントに注記を足す。(評価, コード, 遅延か)"""
    if not is_late(p) or g not in GRADES_SD: return g, iss, False
    g2 = GRADES_SD[min(GRADES_SD.index(g) + 1, 4)]
    return g2, iss + ([f"LATE@{g}→{g2}"] if g2 != g else []), True

def checker(n):
    def deco(f): CHECKERS[n] = f; return f
    return deco

KANJI_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}
ID_PAT = re.compile(r"(?<![A-Za-z0-9])(\d{2}[A-Za-z]{2}\d{3}|\d{7,8})(?![0-9])")
LATIN_NAME_PAT = re.compile(r"[\s_\-－ー−]*([A-Za-z][A-Za-z .']*)")
NAME_PAT = re.compile(r"[\s_\-－ー−]*([\u3040-\u30ff\u4e00-\u9fff々]+)")

def parse_name(stem, default_chapter=None):
    """学生番号・氏名・章をファイル名から読む。「25xx000山田太郎_…_課題１」「2024001－山田太郎－課題1」等に対応。"""
    s = unicodedata.normalize("NFKC", stem)
    m = re.search(r"(?:課題|课题)\s*(\d+)", s) or re.search(r"第\s*(\d+)\s*章", s)
    chapter = int(m.group(1)) if m else None
    if chapter is None:
        k = re.search(r"第\s*([一二三四五六七八九十])\s*章", s)
        chapter = KANJI_NUM[k.group(1)] if k else default_chapter
    sid = name = ""
    m = ID_PAT.search(s)
    if m:
        sid = m.group(1).upper()
        n = NAME_PAT.match(s, m.end())
        if n:
            name = re.split(r"第|課題|课题|章末", n.group(1))[0].rstrip("ー")
        if not name:                                  # 留学生など、アルファベットの氏名(例: 26xx000Taro Yamada_…)
            n = LATIN_NAME_PAT.match(s, m.end())
            if n: name = re.sub(r"\s+", " ", n.group(1)).strip()
    return sid, name, chapter, bool(sid and name)

def textboxes(path):
    out = []
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.startswith("xl/drawings/drawing") and n.endswith(".xml"):
                xml = z.read(n).decode("utf8")
                for sp in re.split(r"<xdr:(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)", xml)[1:]:
                    paras = re.findall(r"<a:p>.*?</a:p>", sp, flags=re.S)
                    t = "\n".join("".join(re.findall(r"<a:t>(.*?)</a:t>", p, flags=re.S)) for p in paras)
                    if t.strip(): out.append(t.strip())
    return "\n".join(out)

def cell_essay(ws, h=None):
    """セルに書かれた考察の文章。元データ(1〜12行目のA列・タイトル)、表の項目名・表頭、単位・出所の記載は除く。"""
    skip = {norm_label(ws.cell(r, 1).value) for r in range(1, 13)} - {""}
    side = {r for r, _ in table_rows(ws, h)} if h else set()      # 表の行(表側の項目名がある行)
    out = []
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if not isinstance(v, str) or v.startswith("="): continue
            t = norm_label(v)
            if len(t) < 2 or t in skip or re.match(r"(単位|出所|出典|資料|注)", t) or re.fullmatch(r"[\d.,%年()()~〜-]+", t): continue
            if c.row <= 12 and c.column == 1: continue
            if h and c.column == 1 and (h - 3 <= c.row <= h or c.row in side) and "。" not in t: continue     # 表のタイトル・表頭・表側(表の下の行は考察として読む。「。」があれば考察とみなす)
            out.append(v.strip())
    return "\n".join(out)

def judge_essay(text, chapter):                 # ローカルLLM(MLX)に差し替える場所。None=未判定
    return None

# ───────── 第1章 ─────────
LABELS.update({
 "C1_NO_TABLE": "構成比の表が見つからない/未完成", "C1_HARDCODED": "数式を使わず数値を入力している",
 "C1_NO_ABS_REF": "絶対セル参照を使っていない", "C1_PARTIAL_ABS_REF": "一部の行で絶対セル参照が使われていない(注意のみ・減点なし)", "C1_WRONG_ABS_REF": "絶対セル参照の付け方が誤り($B$4型など)",
 "C1_WRONG_VALUE": "構成比の値/式が誤り(絶対参照以外)", "C1_WRONG_NUM": "分子のセル参照が誤り(行のずれなど)", "C1_SUM_TRAP": "雇用者と内訳を合計している(ひっかけ)",
 "C1_INCOMPLETE": "表に不足している行・空欄がある", "C1_FMT": "表示形式が適切でない(パーセントスタイル、または小数点第1位までの数値)", "C1_NO_UNIT": "「単位」の記載がない(注意のみ)", "C1_NO_TITLE": "表のタイトルがない(注意のみ)", "C1_NO_HEADER": "表頭(年)が不足している(注意のみ)", "C1_NO_SIDEHEAD": "表側(項目名)が入力されていない行がある(注意のみ)",
 "C1_SRC_MODIFIED": "元データを書き換えている", "C1_NO_ESSAY": "考察がない", "C1_ESSAY_IN_CELL": "考察をテキストボックスでなくセルに書いている(注意のみ)", "C1_SHORT_ESSAY": "考察が短い(注意のみ・減点なし)", "C1_TERM": "用語の使い分けを要確認(割合に増加/減少)",
 "C1_MULTI_SHEET": "シートが複数ある", "C1_EMPTY_REF": "元データのない行(空のセル)を参照して、0%などになっている行がある(注意のみ)", "C1_PARTIAL_HARD": "100%の行(雇用者)を数式でなく手入力している(注意のみ・採点には影響しない)"})
COMMENTS.update({
 "C1_NO_TABLE": "構成比の表が見つかりませんでした。「被雇用者の雇用形態別構成比」の表を作成してください。",
 "C1_HARDCODED": "構成比が数式ではなく数値で入力されています。セル参照を使った数式で計算しましょう。",
 "C1_PARTIAL_HARD": "100%の行が数式ではなく数値で入力されています。=B4/B$4 のように数式で入力すると、他の行と同じ式をコピーできます。",
 "C1_NO_ABS_REF": "分母の「雇用者(役員を除く)」に絶対参照(行を固定する $ 、例: B$4)を使うと、1つの式をコピーするだけで全体を計算できます。",
 "C1_PARTIAL_ABS_REF": "一部の行で絶対参照が使われていません。分母を B$4 のように行固定にして、全ての行で同じ式をコピーできるようにしましょう。",
 "C1_WRONG_ABS_REF": "絶対参照の付け方が違うため、一部のセルで構成比が正しく計算されていません。分母は「同じ年(列)の雇用者数」なので、行だけを固定します(B$4)。",
 "C1_WRONG_NUM": "分子のセルがずれているため、構成比が正しく計算されていません。分子は同じ行の項目の人数です(例: 正規の職員・従業員の行なら B5)。100%の行も確認しましょう。",
 "C1_WRONG_VALUE": "構成比の式を確認してください。分子は各雇用形態の人数、分母は同じ年の「雇用者(役員を除く)」です。",
 "C1_SUM_TRAP": "「雇用者(役員を除く)」は全体で、他の項目はその内訳です。内訳と足し合わせないよう注意しましょう。",
 "C1_INCOMPLETE": "表に不足している行や空欄があります。元の表と同じ項目を、すべて入力しましょう。",
 "C1_NO_UNIT": "表に「単位:%」(または単位の記載)がありません。",
 "C1_NO_TITLE": "表にタイトル(「被雇用者の雇用形態別構成比」など)を入力しましょう。",
 "C1_EMPTY_REF": "元データにない行を参照していて、構成比が0%などになっている行があります。不要な行は削除しましょう。",
 "C1_NO_HEADER": "表頭(年)が、元の表と同じように、すべて入力されていません。",
 "C1_NO_SIDEHEAD": "表側(項目名)が入力されていない行があります。元の表と同じように入力しましょう。",
 "C1_FMT": "構成比の表示形式を、パーセントスタイル(小数点第1位まで)、または小数点第1位までの数値にそろえましょう。",
 "C1_SRC_MODIFIED": "元のデータが書き換えられているようです。元データは変更しないでください。",
 "C1_NO_ESSAY": "考察が書かれていません。計算結果から読み取れる近年の雇用形態の変化を、テキストボックスにまとめましょう。",
 "C1_ESSAY_IN_CELL": "考察はセルではなく、テキストボックスに書きましょう(今回はセルの文章を考察として評価しています)。",
 "C1_SHORT_ESSAY": "考察をもう少し詳しく書いてみましょう。正規・非正規の全体の動きに加えて、内訳(パート、アルバイトなど)の変化にも触れると、さらに良い考察になります。",
 "C1_TERM": "割合の変化には「上昇・低下」、人数の変化には「増加・減少」を使い分けましょう。",
 "C1_MULTI_SHEET": "提出するシートは、課題のシート1枚だけにしましょう。"})
PAT1 = re.compile(r'^=\s*(\$?)([A-Z]{1,2})(\$?)(\d+)\s*/\s*(\$?)([A-Z]{1,2})(\$?)(\d+)\s*$')
SUMRNG = re.compile(r'SUM\(\s*\$?([A-Z]{1,2})\$?(\d+)\s*:\s*\$?([A-Z]{1,2})\$?(\d+)\s*\)')
COLS1, SRC_ROWS1, MIN_ESSAY1 = range(2, 25), range(4, 11), 100      # 元データは4〜10行目(雇用者〜その他)
MIN_ESSAY_NONE = 10          # これ未満は「考察なし」、MIN_ESSAY1未満は「考察が短い」(改行・空白は数えない)
SHORT_ESSAY_PENALTY = False   # 課題1では、短い考察は減点しない(Sのまま、詳しく書くよう促すコメントだけ付ける)
HAND_TYPED_MAX = 0            # 表の中の手入力の数値(数式でないセル)がこの個数を超えると「式なし」=D。授業の方針は「手入力はD」なので0。ただし100%の行(雇用者)だけの手入力は許容し、注意のみ
VALUE_NG_RATIO = 0.05         # 数式セルのうち、計算結果が違うセルの割合がこれを超えると「値誤り」(0.05=約1行でも誤りがあれば。1行までなら許容する場合は0.13程度)
ABS_OK_RATIO = 0.75          # 数式セルのうち、これ以上で絶対参照を使っていれば「使用」とみなす(全部でなければ注意のみ・減点なし)
CHECK_FMT = True             # 表示形式: パーセントスタイル、または小数点第1位までの数値ならOK(違っても減点せず注意のみ)
ROWNAMES1 = ["雇用者(100%)", "正規", "パート", "アルバイト", "派遣", "契約・嘱託", "その他"]

def val_ok(vv, exp):
    """セルの値が期待値(構成比)と合っているか。Trueで一致、Falseで不一致、Noneで判定不能(未計算など)。"""
    if not isinstance(vv, (int, float)) or isinstance(vv, bool) or exp is None: return None
    return abs(vv - exp) <= 0.0006 or abs(vv - exp * 100) <= 0.06       # 割合のまま/百分率(*100)のどちらもよい。四捨五入の差は許容

def classify1(cell, c, src_row, vv=None, exp=None):
    v = cell.value
    if isinstance(v, (int, float)) and not isinstance(v, bool):          # 手入力の数値
        return "hard_wrong" if val_ok(v, exp) is False else "hard"
    if not (isinstance(v, str) and v.startswith("=")):
        return "empty"
    f = re.sub(r"\s+", "", v).upper()
    f = re.sub(r"\*100(?:\.0+)?(?=\)|$)", "", f)     # 「…*100」(百分率に直す書き方)も許容
    f = re.sub(r"\*100(?:\.0+)?/", "/", f)            # 「=B5*100/B$4」
    f = re.sub(r"^=100\*", "=", f)                     # 「=100*B5/B$4」
    f = re.sub(r"^=\(([^()]*)\)$", r"=\1", f)          # 全体を囲む括弧
    f = re.sub(r"\(([^()/]*)\)/", r"\1/", f)           # 「=(B5*100)/B$4」型の分子の括弧
    m = PAT1.match(f)
    k = "other"
    if m:
        _, nc, _, nr, d1, dc, d2, dr = m.groups()
        n1 = m.group(1)
        num_ok, den_ok = ci(nc) == c and int(nr) == src_row, ci(dc) == c and int(dr) == 4
        if num_ok and den_ok:
            k = "ok_abs" if (d2 and not d1) else "ok_rel"
        elif not num_ok and not (int(nr) == src_row and n1):     # 分子が別の行・別の列(行のずれなど)。$B5 のような列固定は絶対参照の誤りとして扱う
            k = "wrong_num_abs" if (not den_ok and d1) else "wrong_num"
        else:
            k = "wrong_abs" if (d1 or d2) else "wrong"
    ok = val_ok(vv, exp)                          # 計算済みの値があれば、実際の計算結果も照合する
    if k in ("ok_abs", "ok_rel") and ok is False: return "badval_abs" if k == "ok_abs" else "badval_rel"   # 式の形は正しいが、計算結果が違う
    if k == "other" and ok is True: return "ok_valabs" if re.search(r"\$4(?!\d)", f) else "ok_val"   # 見慣れない書き方でも、結果が合っていれば正しい
    return k

def src_value(src, c, sheets):
    """元データの値。"""
    for w in sheets:
        v = w.cell(src, c).value
        if isinstance(v, (int, float)) and not isinstance(v, bool): return v
    return None

def fmt_good(nf):
    """パーセントスタイル(0.0%)、または小数点第1位までの数値(0.0、#,##0.0 など)か。"""
    t = re.sub(r'_.|"[^"]*"|\\.|\*.|\[[^\]]*\]|\s', "", str(nf or "General").split(";")[0])
    return re.fullmatch(r"[#,0]*0\.0%?", t) is not None

def norm_label(v):
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(v or "")))

def resolve_label(ws, cell):                    # 表側の項目名(文字列、または =A4 のような参照)
    v = cell.value
    if isinstance(v, str) and v.startswith("="):
        m = re.fullmatch(r"=\$?([A-Z]{1,2})\$?(\d+)", v.strip())
        v = ws.cell(int(m.group(2)), ci(m.group(1))).value if m else ""
    return norm_label(v)

def yearish(x):
    return (isinstance(x, (int, float)) and not isinstance(x, bool) and 1990 <= x <= 2100) or \
        (isinstance(x, str) and (re.fullmatch(r"\s*\d{4}\s*", x) is not None or
                                 re.fullmatch(r"=\$?[A-Z]{1,2}\$?3|=\$?[A-Z]{1,2}\$?\d+\+1", x.replace(" ", "")) is not None))   # 年の文字・=B3 や =B16+1 のような式

def find_header(ws):                            # 年(2002, 2003…)が並ぶ表頭の行。元データ(1〜12行目)より下で探す
    for r in range(13, 61):
        if yearish(ws.cell(r, 2).value) and yearish(ws.cell(r, 3).value): return r
    labels = {norm_label(ws.cell(r, 1).value) for r in SRC_ROWS1} - {""}     # 表頭がない答案: 表側の項目名から表を探す
    for r in range(13, 61):
        if resolve_label(ws, ws.cell(r, 1)) in labels: return r - 1
    return None

def table_rows(ws, h):
    """表頭の下の各行を、元データの行(4〜10)に対応づける。それ以外の行(合計など)は判定しない。項目名で照合し、できなければ並び順で補う。"""
    src_labels = {norm_label(ws.cell(r, 1).value): r for r in SRC_ROWS1}
    out, used = [], set()
    for k in range(10):
        r = h + 1 + k
        if all(ws.cell(r, c).value is None for c in range(1, 25)): break
        src = src_labels.get(resolve_label(ws, ws.cell(r, 1)))
        if src is None or src in used: src = 4 + k
        if src not in SRC_ROWS1 or src in used: continue
        used.add(src); out.append((r, src))
    return out

def grade1(calc, abs_, essay):                  # 「採点区分(案)」シートの対応表
    if calc == "表なし": return "評価外"
    if calc == "式なし": return "D"
    if calc == "値誤り": return "D" if essay == "なし" else "C"
    if abs_ == "使用": return {"妥当": "S", "不十分": "A", "なし": "B"}[essay]
    return {"妥当": "A", "不十分": "B", "なし": "C"}[essay]

@checker(1)
def check1(path, ref):
    wb = load_workbook(path); ws = wb.worksheets[0]
    wsv = load_workbook(path, data_only=True).worksheets[0]          # 計算済みの値(Excelで保存したファイルにある)
    iss = []
    if len(wb.worksheets) > 1: iss.append("C1_MULTI_SHEET")
    h = find_header(ws)                             # タイトルの表記ゆれに左右されないよう、表頭の行から表を見つける
    essay_text = textboxes(path)
    in_cell = cell_essay(ws, h)
    if len(norm_label(in_cell)) >= MIN_ESSAY_NONE and len(norm_label(essay_text)) < MIN_ESSAY1:   # テキストボックスに十分な考察がなく、セルに文章がある → セルの文章も考察として扱う
        essay_text = (essay_text + "\n" + in_cell).strip(); iss.append("C1_ESSAY_IN_CELL")
    n_es = len(norm_label(essay_text))
    if n_es < MIN_ESSAY_NONE: essay = "なし"; iss.append("C1_NO_ESSAY")
    elif n_es < MIN_ESSAY1: essay = "不十分" if SHORT_ESSAY_PENALTY else "妥当"; iss.append("C1_SHORT_ESSAY")
    else: essay = judge_essay(essay_text, 1) or "妥当"
    if h is None:
        return "評価外", iss + ["C1_NO_TABLE"], {}
    cnt, fmt_ok, hard_rows, rel_rows = Counter(), 0, [], []
    trs = table_rows(ws, h)
    row_of = {src: r for r, src in trs}
    sheets = ([load_workbook(ref, data_only=True).worksheets[0]] if ref else []) + [ws, wsv]   # 期待値の元データ(模範解答があればそれを優先)
    for r, src in trs:
        row_hard = row_rel = 0
        for c in COLS1:
            cell = ws.cell(r, c)
            num, den = src_value(src, c, sheets), src_value(4, c, sheets)
            exp = num / den if num is not None and den else None
            k = classify1(cell, c, src, wsv.cell(r, c).value, exp)
            if src == 4 and k == "hard": k = "hard100"              # 100%の行の手入力は許容(値が合っていれば注意のみ)
            elif src == 4 and k == "hard_wrong": k = "wrong100"     # 値が違えば、手入力ではなく「計算結果の誤り」
            if k in ("badval_abs", "badval_rel") and ws.cell(src, c).value is None:     # 元データのない(空の)セルを参照した式 → 計算の誤りではなく注意のみ
                k = "emptyref_abs" if k == "badval_abs" else "emptyref_rel"
            cnt[k] += 1
            row_hard += k in ("hard", "hard_wrong", "hard100", "wrong100"); row_rel += k == "ok_rel"
            fmt_ok += cell.value is not None and fmt_good(cell.number_format)
        if row_hard >= len(COLS1) / 2: hard_rows.append(ROWNAMES1[src - 4])
        if row_rel >= len(COLS1) / 2: rel_rows.append(ROWNAMES1[src - 4])
    missing = [r for r in SRC_ROWS1 if r not in row_of]
    cnt["empty"] += len(missing) * len(COLS1)               # 表にない行は空欄として数える
    total = len(SRC_ROWS1) * len(COLS1)
    if (total - cnt["empty"]) / total < 0.2:        # 表の2割未満しか埋まっていなければ「表なし」扱い
        return "評価外", iss + ["C1_NO_TABLE"], dict(cnt)
    if cnt["empty"]: iss.append("C1_INCOMPLETE")
    filled = total - cnt["empty"]
    handtyped = cnt["hard"] + cnt["hard_wrong"]
    fcells = sum(cnt[k] for k in ("ok_abs", "ok_rel", "ok_val", "ok_valabs", "emptyref_abs", "emptyref_rel", "badval_abs", "badval_rel", "wrong", "wrong_abs", "wrong_num", "wrong_num_abs", "other"))   # 数式が入っているセル
    bad = cnt["wrong"] + cnt["wrong_abs"] + cnt["wrong_num"] + cnt["wrong_num_abs"] + cnt["other"] + cnt["hard_wrong"] + cnt["wrong100"] + cnt["badval_abs"] + cnt["badval_rel"]
    if handtyped > HAND_TYPED_MAX or handtyped / filled > 0.5:
        calc, abs_ = "式なし", "-"; iss.append("C1_HARDCODED")
    else:
        calc = "値誤り" if bad / max(fcells + cnt["hard_wrong"] + cnt["wrong100"], 1) > VALUE_NG_RATIO else "正"
        abs_ratio = (cnt["ok_abs"] + cnt["ok_valabs"] + cnt["badval_abs"] + cnt["emptyref_abs"]) / max(fcells, 1)
        abs_ = "使用" if abs_ratio >= ABS_OK_RATIO else "不使用"
        if cnt["wrong_abs"] + cnt["wrong_num_abs"]: iss.append("C1_WRONG_ABS_REF")
        elif abs_ == "不使用": iss.append("C1_NO_ABS_REF")
        elif abs_ratio < 0.999: iss.append("C1_PARTIAL_ABS_REF")     # ほとんどで使っている → 注意のみ
        if cnt["wrong_num"] + cnt["wrong_num_abs"]: iss.append("C1_WRONG_NUM")
        if cnt["wrong"] + cnt["other"] + cnt["hard_wrong"] + cnt["wrong100"] + cnt["badval_abs"] + cnt["badval_rel"]: iss.append("C1_WRONG_VALUE")
        if cnt["emptyref_abs"] + cnt["emptyref_rel"]: iss.append("C1_EMPTY_REF")        # 注意のみ
        if cnt["hard100"] or cnt["wrong100"]: iss.append("C1_PARTIAL_HARD")     # 100%の行の手入力 → 注意のみ
    if CHECK_FMT and fmt_ok / max(filled, 1) < 0.9: iss.append("C1_FMT")
    unit = any("単位" in str(c.value or "") for row in ws.iter_rows(min_row=max(1, h - 3), max_row=h + 1) for c in row)
    if not unit: iss.append("C1_NO_UNIT")                                   # 以下4つは、ない場合も減点せず注意のみ
    if not any(isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip() for r in range(max(13, h - 3), h)):
        iss.append("C1_NO_TITLE")
    if sum(yearish(ws.cell(h, c).value) for c in COLS1) < len(COLS1): iss.append("C1_NO_HEADER")
    if any(not resolve_label(ws, ws.cell(r, 1)) for r, _ in trs): iss.append("C1_NO_SIDEHEAD")
    trap = any(isinstance(c.value, str) and any(int(m.group(2)) <= 4 <= int(m.group(4))
               for m in SUMRNG.finditer(c.value.upper())) for row in ws.iter_rows() for c in row)
    if trap: iss.append("C1_SUM_TRAP")
    if ref:
        rs = load_workbook(ref, data_only=True).worksheets[0]
        if any(ws.cell(r, c).value != rs.cell(r, c).value for r in range(4, 11) for c in COLS1):
            iss.append("C1_SRC_MODIFIED")
    if re.search(r"(構成比|割合|比率|シェア)[^。、，,]{0,20}(増加|減少|増え|減っ)", essay_text): iss.append("C1_TERM")
    detail = {f"n_{k}": v for k, v in cnt.items()} | {"考察文字数": n_es, "手入力の行": "、".join(hard_rows), "絶対参照なしの行": "、".join(rel_rows),
                                                "不足している行": "、".join(ROWNAMES1[k - 4] for k in missing)}
    return grade1(calc, abs_, essay), iss, detail

# ───────── 章ごとのチェッカー(プラグイン) ─────────
def load_plugins():
    """checkers/ フォルダ(grader.pyと同じ場所)にある *.py を読み込む。ファイル名が _ で始まるものは無視。
    各ファイルは register(api) を持ち、api.checker(章番号) でチェッカーを登録し、api.LABELS に「できていない箇所」の説明を足す。"""
    d = Path(__file__).resolve().parent / "checkers"
    if not d.is_dir(): return
    api = types.SimpleNamespace(checker=checker, label_of=label_of, LABELS=LABELS, COMMENTS=COMMENTS, textboxes=textboxes, judge_essay=judge_essay,
                                norm_label=norm_label, resolve_label=resolve_label, parse_name=parse_name)
    for p in sorted(d.glob("*.py")):
        if p.name.startswith("_"): continue
        try:
            spec = importlib.util.spec_from_file_location(f"checker_{p.stem}", p)
            m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m); m.register(api)
            print(f"チェッカーを読み込みました: {p.name}")
        except Exception as e:
            print(f"チェッカー {p.name} を読み込めませんでした(スキップします): {e}")

# ───────── 実行部(出力はExcel) ─────────
LABELS["FILENAME"] = "ファイル名から学生番号または氏名を読み取れない"
LABELS["LATE"] = LABELS_LATE
COMMENTS["LATE"] = "※提出期限を過ぎてから提出されたため、評価を1段階下げています({items})。"
GRADES = ["S", "A", "B", "C", "D", "評価外", "未対応"]
FONT, HFILL = "Meiryo", PatternFill("solid", fgColor="1F3864")

def sheet(ws, header, data, widths):
    ws.append(header)
    for c in ws[1]:
        c.font = Font(name=FONT, bold=True, color="FFFFFF", size=10); c.fill = HFILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for r in data: ws.append(r)
    for row in ws.iter_rows(min_row=2):
        for c in row: c.font = Font(name=FONT, size=10); c.alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate(widths, 1): ws.column_dimensions[gl(i)].width = w
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions

def write_xlsx(path, rows, longrows):
    wb = Workbook()
    first = ["学生番号", "氏名", "課題", "採点結果", FINAL, "Canvas点数", COMMENT_COL, "できていない箇所", "できていない箇所コード", "ファイル"]
    extra = sorted({k for r in rows for k in r} - set(first))
    data = []
    for i, r in enumerate(rows, 2):                       # 最終評価は自動評価をそのまま初期値にし、先生が必要に応じて直す
        r = dict(r); r[FINAL] = r["採点結果"]
        r[COMMENT_COL] = make_comment([c for c in str(r.get("できていない箇所コード", "")).split(";") if c], r["採点結果"])
        r["Canvas点数"] = f"=IFERROR(VLOOKUP(E{i},'Canvas対応表'!$A$2:$B$6,2,FALSE),\"\")"
        data.append([r.get(k, "") for k in first + extra])
    ws = wb.active; ws.title = "採点結果"
    sheet(ws, first + extra, data, [12, 14, 6, 10, 14, 10, 55, 55, 30, 34] + [10] * len(extra))
    for g, col in (("D", "FFC7CE"), ("評価外", "FFC7CE"), ("未対応", "D9D9D9")):
        for L in "DE":
            ws.conditional_formatting.add(f"{L}2:{L}{ws.max_row}", CellIsRule(operator="equal", formula=[f'"{g}"'], fill=PatternFill("solid", bgColor=col)))
    ws["E1"].fill = PatternFill("solid", fgColor="BF8F00")  # 修正する列だと分かる色
    ws["G1"].fill = PatternFill("solid", fgColor="BF8F00")
    dv = DataValidation(type="list", formula1='"S,A,B,C,D,評価外,未対応"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"E2:E{max(ws.max_row, 2)}")
    wm = wb.create_sheet("Canvas対応表")                  # 評価→Canvas点数。ここを直すとCSV出力にも反映される
    sheet(wm, ["評価", "Canvas点数"], [[g, p] for g, p in DEFAULT_POINTS.items()], [10, 12])
    wm["D1"] = "Canvasの課題の配点を最大点(S)に合わせ、評価スキームの範囲に各点数が入るよう設定してください。"
    wm["D1"].font = Font(name=FONT, size=10)
    wl = wb.create_sheet("減点明細")     # 1行=学生×できていない箇所。後で集計・並べ替えに使う
    sheet(wl, ["学生番号", "氏名", "課題", "コード", "内容"],
          [[r["学生番号"], r["氏名"], r["課題"], r["コード"], r["内容"]] for r in longrows], [12, 14, 6, 22, 50])
    wsum = wb.create_sheet("集計")
    chapters = sorted({r["課題"] for r in rows if r["課題"] != ""})
    wsum["A1"] = "採点結果の人数(最終評価)"; wsum["A1"].font = Font(name=FONT, bold=True, size=12)
    wsum["A2"] = f"採点プログラムのバージョン: {GRADER_VERSION}"; wsum["A2"].font = Font(name=FONT, size=9, color="888888")
    hdr = ["課題", "提出数(未対応を除く)"] + GRADES
    for j, h in enumerate(hdr, 1):
        c = wsum.cell(3, j, h); c.font = Font(name=FONT, bold=True, color="FFFFFF", size=10); c.fill = HFILL
        c.alignment = Alignment(wrap_text=True)
    for i, ch in enumerate(chapters, 4):
        wsum.cell(i, 1, ch)
        wsum.cell(i, 2, f"=COUNTIFS('採点結果'!$C:$C,$A{i},'採点結果'!$E:$E,\"<>未対応\")")
        for j in range(len(GRADES)):
            wsum.cell(i, 3 + j, f"=COUNTIFS('採点結果'!$C:$C,$A{i},'採点結果'!$E:$E,{gl(3 + j)}$3)")
    end1 = 3 + len(chapters)
    r0 = end1 + 3
    wsum.cell(r0 - 1, 1, "できていない箇所の人数(減点明細から集計)").font = Font(name=FONT, bold=True, size=12)
    for j, h in enumerate(["課題", "コード", "内容", "人数", "割合"], 1):
        c = wsum.cell(r0, j, h); c.font = Font(name=FONT, bold=True, color="FFFFFF", size=10); c.fill = HFILL
    for i, (code, label) in enumerate(LABELS.items(), r0 + 1):
        m = re.match(r"C(\d+)_", code)
        ch = int(m.group(1)) if m else "共通"
        wsum.cell(i, 1, ch); wsum.cell(i, 2, code); wsum.cell(i, 3, label)
        if m:
            wsum.cell(i, 4, f"=COUNTIFS('減点明細'!$C:$C,$A{i},'減点明細'!$D:$D,$B{i})")
            wsum.cell(i, 5, f"=IFERROR(D{i}/INDEX($B$4:$B${end1},MATCH($A{i},$A$4:$A${end1},0)),0)")
        else:
            wsum.cell(i, 4, f"=COUNTIF('減点明細'!$D:$D,$B{i})")
            wsum.cell(i, 5, f"=IFERROR(D{i}/SUM($B$4:$B${end1}),0)")
        wsum.cell(i, 5).number_format = "0.0%"
    for row in wsum.iter_rows(min_row=4):
        for c in row:
            if c.font.name != FONT or not c.font.bold: c.font = Font(name=FONT, size=10)
    for i, w in enumerate([8, 24, 52, 8, 8, 8, 8, 8, 8, 8], 1): wsum.column_dimensions[gl(i)].width = w
    wb.move_sheet("集計", offset=-2)
    wb.save(path)

def pick_folder(prompt, start=None):
    """macOSのフォルダ選択ダイアログ(osascript)。使えなければtkinter。キャンセルならNone。"""
    loc = f' default location (POSIX file "{start}")' if start and Path(start).exists() else ""
    try:
        r = subprocess.run(["osascript", "-e", f'POSIX path of (choose folder with prompt "{prompt}"{loc})'],
                           capture_output=True, text=True)
        if r.returncode == 0 and r.stdout.strip(): return Path(r.stdout.strip())
        if "-128" in r.stderr: return None          # ユーザーがキャンセル
    except FileNotFoundError:
        pass
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        p = filedialog.askdirectory(title=prompt, initialdir=start or str(Path.home())); root.destroy()
        return Path(p) if p else None
    except Exception:
        return None

def ask_chapter():
    """ファイル名から章を判別できないファイルに使う章番号を、ダイアログで尋ねる。"""
    try:
        r = subprocess.run(["osascript", "-e", 'text returned of (display dialog "この提出フォルダは第何章の課題ですか?(ファイル名から章を判別できないファイルに使います)" default answer "1")'],
                           capture_output=True, text=True)
        if r.returncode == 0 and r.stdout.strip().isdigit(): return int(r.stdout.strip())
    except FileNotFoundError:
        pass
    return None

# ───────── Canvas取り込み用CSVの作成 ─────────
def osa(script):
    try:
        r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    except FileNotFoundError:
        return None
    return r.stdout.strip() if r.returncode == 0 else None

def pick_file(prompt, start=None):
    loc = f' default location (POSIX file "{start}")' if start and Path(start).exists() else ""
    out = osa(f'POSIX path of (choose file with prompt "{prompt}"{loc})')
    return Path(out) if out else None

def pick_save(prompt, name, start=None):
    loc = f' default location (POSIX file "{start}")' if start and Path(start).exists() else ""
    out = osa(f'POSIX path of (choose file name with prompt "{prompt}" default name "{name}"{loc})')
    return Path(out) if out else None

def choose_from(prompt, items, default=None):
    q = lambda x: x.replace("\\", "\\\\").replace('"', '\\"')
    lst = ", ".join(f'"{q(i)}"' for i in items)
    d = f' default items {{"{q(default)}"}}' if default else ""
    out = osa(f'choose from list {{{lst}}} with prompt "{q(prompt)}"{d}')
    return out if out and out != "false" else None

def read_csv_any(path):
    b = Path(path).read_bytes()
    bom = b.startswith(b"\xef\xbb\xbf")
    for enc in ("utf-8-sig", "cp932"):
        try: return list(csv.reader(io.StringIO(b.decode(enc)))), bom
        except UnicodeDecodeError: pass
    raise ValueError("CSVの文字コードを読み取れません(UTF-8で保存し直してください)")

def nrm(x): return unicodedata.normalize("NFKC", str(x or "")).strip().upper()
def nname(x): return re.sub(r"[\s,、]+", "", unicodedata.normalize("NFKC", str(x or ""))).lower()

def prepare_export(a):
    res = Path(a.results) if a.results else pick_file("確認・修正済みの採点結果.xlsxを選んでください", "/Volumes")
    if not res: print("採点結果が選ばれなかったため中止します"); return
    cv = Path(a.canvas) if a.canvas else pick_file("Canvasからエクスポートした成績表CSVを選んでください", str(res.parent))
    if not cv: print("Canvasの成績表CSVが選ばれなかったため中止します"); return
    wb = load_workbook(res, data_only=True)
    ws = wb["採点結果"]
    hd = [c.value for c in ws[1]]
    if FINAL not in hd: print(f"「{FINAL}」列が見つかりません。grader.pyで作った採点結果.xlsxを使ってください"); return
    pts = dict(DEFAULT_POINTS)
    if "Canvas対応表" in wb.sheetnames:
        pts = {str(r[0]).strip().upper(): r[1] for r in wb["Canvas対応表"].iter_rows(min_row=2, values_only=True) if r[0] is not None and r[1] is not None}
    items, skipped = [], []
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(hd, r))
        if not d.get("学生番号") and not d.get("氏名"): continue
        g = str(d.get(FINAL) or "").strip().upper()
        if g in pts: items.append(d | {"_g": g})
        elif str(d.get(COMMENT_COL) or "").strip() and g in ("評価外", ""): items.append(d | {"_g": None})   # 評価外でもコメントがあれば、コメントだけ送る(点数は付けない)
        else: skipped.append(d | {"_g": g})
    chs = sorted({d["課題"] for d in items if d.get("課題") not in (None, "")})
    if len(chs) > 1 and not a.chapter: print(f"複数の章({chs})が含まれています。--chapter で章を指定してください"); return
    if a.chapter: items = [d for d in items if d.get("課題") == a.chapter]
    rows, bom = read_csv_any(cv)
    head = rows[0]
    asg = [i for i, h in enumerate(head) if re.search(r"\(\d+\)\s*$", h)]
    if not asg: print("Canvasの成績表CSVで、課題の列(「課題名 (数字)」の形)が見つかりません"); return
    ident = list(range(asg[0]))
    if a.column:
        nf = lambda x: unicodedata.normalize("NFKC", x)           # 全角の数字・空白の違いを吸収する
        hit = [i for i in asg if nf(a.column) in nf(head[i])]
        if len(hit) != 1: print(f"「{a.column}」に一致する課題列が{len(hit)}個です。候補:", [head[i] for i in asg]); return
        col = hit[0]
    else:
        ch = a.chapter or (chs[0] if chs else None)
        nfh = lambda x: unicodedata.normalize("NFKC", x)
        cand = [head[i] for i in asg if ch and f"{ch}章" in nfh(head[i]) and "章末" in nfh(head[i])] or \
               [head[i] for i in asg if ch and f"{ch}章" in nfh(head[i])]
        pick = choose_from("取り込む課題の列を選んでください", [head[i] for i in asg], cand[0] if len(cand) == 1 else None)
        if not pick: print("課題の列が選ばれなかったため中止します"); return
        col = head.index(pick)
    meta, students = [], []
    for r in rows[1:]:
        r = r + [""] * (len(head) - len(r))
        (students if any(r[i].strip() for i in ident[1:]) else meta).append(r)
    out_rows, used, unmatched = [], set(), []
    idc = head.index("ID") if "ID" in head else None
    for d in items:
        sid, nm = nrm(d.get("学生番号")), nname(d.get("氏名"))
        m = None
        cu = re.search(r"_(?:LATE_)?(\d{3,})_\d{3,}_", str(d.get("ファイル") or ""))     # Canvasの一括ダウンロードのファイル名「氏名_ユーザーID_提出ID_元の名前」
        if cu and idc is not None: m = next((k for k, r in enumerate(students) if k not in used and r[idc].strip() == cu.group(1)), None)
        sp = re.compile(rf"(?<![0-9A-Z]){re.escape(sid)}(?![0-9])") if sid else None     # 学生番号(「B03経済01経済-24AK000 山田 太郎」のように氏名の列に入っていてもよい)
        if m is None and sp: m = next((k for k, r in enumerate(students) if k not in used and any(sp.search(nrm(r[i])) for i in ident)), None)
        if m is None and nm: m = next((k for k, r in enumerate(students) if k not in used and nname(r[0]).endswith(nm)), None)
        if m is None: unmatched.append(d); continue
        used.add(m); out_rows.append((m, d))
    return dict(res=res, hd=hd, pts=pts, items=items, skipped=skipped, chs=chs, head=head, bom=bom, col=col,
                ident=ident, meta=meta, students=students, out_rows=out_rows, unmatched=unmatched)

def is_test_student(r):
    """Canvasのテスト用受講生(「受講者として表示」用)の行か。採点対象ではないので、「採点結果にない学生」の数には入れない。"""
    return nname(r[0]).lower() in ("テスト用受講生", "テスト受講生", "studenttest", "teststudent")

def export_canvas(a):
    c = prepare_export(a)
    if not c: return
    res, pts, chs, head, bom, col, ident, meta, students, out_rows, unmatched, skipped = (
        c[k] for k in ("res", "pts", "chs", "head", "bom", "col", "ident", "meta", "students", "out_rows", "unmatched", "skipped"))
    out_cols = ident + [col]
    ch0 = a.chapter or (chs[0] if chs else None)
    outp = Path(a.csv_out) if a.csv_out else pick_save("Canvas取り込み用CSVの保存先", f"canvas取り込み用_第{ch0}章.csv" if ch0 else "canvas取り込み用.csv", str(res.parent))
    if not outp: print("保存先が選ばれなかったため中止します"); return
    with open(outp, "w", newline="", encoding="utf-8-sig" if bom else "utf-8") as f:
        w = csv.writer(f)
        w.writerow([head[i] for i in out_cols])
        for r in meta: w.writerow([r[i] for i in out_cols])
        for m, d in sorted(out_rows):
            if d["_g"] is None: continue                  # コメントのみの学生はCSVに入れない
            row = [students[m][i] for i in out_cols]; row[-1] = d["_g"] if a.letters else pts[d["_g"]]; w.writerow(row)
    print(f"出力: {outp}\n課題列: {head[col]}")
    print(f"出力 {len(out_rows)} 人 / 評価の内訳:", dict(Counter(d["_g"] for _, d in out_rows)))
    if skipped: print(f"評価がS〜Dでないため出力しなかった学生 {len(skipped)} 人:", ", ".join(f"{d.get('学生番号') or ''} {d.get('氏名') or ''}({d['_g'] or '空欄'})" for d in skipped))
    if unmatched: print(f"Canvasの成績表と照合できなかった学生 {len(unmatched)} 人:", ", ".join(f"{d.get('学生番号')} {d.get('氏名')}" for d in unmatched))
    used_idx = {m for m, _ in out_rows}
    others = [r for k, r in enumerate(students) if k not in used_idx and not is_test_student(r)]
    print(f"Canvasの成績表のうち採点結果にない学生: {len(others)} 人(出力に含めていません。テスト用受講生は数えていません)")


# ───────── SpeedGrader相当: Canvas APIで点数とコメントを送る ─────────
def parse_course_url(u):
    m = re.match(r"\s*(https?)://([^/\s]+)/courses/(\d+)", u or "")
    return (f"{m.group(1)}://{m.group(2)}", m.group(3)) if m else None

def confirm_dialog(msg):
    out = osa(f'display dialog "{msg}" buttons {{"キャンセル", "送信する"}} default button "キャンセル" with icon caution')
    return bool(out and "送信する" in out)

GRADE_ORDER = ["S", "A", "B", "C", "D"]

def fetch_submissions(base, course, aid, token):
    """課題の全提出物(遅延かどうか、期限、コメントを含む)をCanvasから取得する。学生ID -> 提出物"""
    nxt = f"{base}/api/v1/courses/{course}/assignments/{aid}/submissions?include[]=submission_comments&per_page=100"
    subs = {}
    while nxt:
        req = urllib.request.Request(nxt, headers={"Authorization": f"Bearer {token.strip()}"})
        with urllib.request.urlopen(req, timeout=30) as r:
            for x in json.loads(r.read().decode("utf-8")): subs[str(x.get("user_id"))] = x
            link = r.headers.get("Link", "")
        m = re.search(r'<([^>]+)>;\s*rel="next"', link)
        nxt = m.group(1) if m else None
    return subs

def late_note(sub):
    """期限後の提出なら、期限の日時(この端末の時刻帯)の文字列。期限内・未提出なら None。"""
    if not sub or not sub.get("late"): return None
    from datetime import datetime
    due = sub.get("cached_due_date")
    try:
        t = datetime.fromisoformat(due.replace("Z", "+00:00")).astimezone()
        return f"{t.month}月{t.day}日 {t:%H:%M}"
    except Exception: return ""

def apply_late(g, comment, sub, use_late=True):
    """遅延の提出は、評価を1段階下げ、その旨をコメントの最後に足す(Dはそのまま)。(評価, コメント, 下げたか)"""
    due = late_note(sub) if use_late else None
    if due is None or g not in GRADE_ORDER[:-1]: return g, comment, False
    g2 = GRADE_ORDER[GRADE_ORDER.index(g) + 1]
    note = f"※提出期限{f'({due})' if due else ''}を過ぎてから提出されたため、評価を1段階下げています({g}→{g2})。"
    return g2, (comment + "\n" + note).strip(), True

def post_feedback(a):
    c = prepare_export(a)
    if not c: return
    head, col, students, out_rows, pts = c["head"], c["col"], c["students"], c["out_rows"], c["pts"]
    if "ID" not in head: print("Canvasの成績表CSVに「ID」列がありません(学生をAPIで指定するために必要です)"); return
    idcol = head.index("ID")
    m = re.search(r"\((\d+)\)\s*$", head[col]); aid = m.group(1)
    url = a.url or osa('text returned of (display dialog "CanvasのコースのURL(例: https://〜/courses/101839)を貼り付けてください" default answer "")')
    pc = parse_course_url(url)
    if not pc: print("コースのURLの形式が正しくありません(例: https://〜/courses/101839)"); return
    base, course = pc
    sent_path = Path(c["res"]).parent / f"canvas_sent_{aid}.json"
    sent = json.loads(sent_path.read_text(encoding="utf-8")) if sent_path.exists() else {}
    token = get_token()                               # 遅延の提出を調べるため、確認の段階でも使う(保存はしない)
    if not token: print("トークンが入力されなかったため中止します"); return
    try: subs = fetch_submissions(base, course, aid, token)
    except urllib.error.HTTPError as e: print(f"Canvasから提出状況を取得できませんでした: HTTP {e.code}(トークンやURLを確認してください)"); return
    todo, already, lates = [], 0, []
    for mi, d in sorted(out_rows):
        uid = students[mi][idcol].strip()
        g, comment, down = apply_late(d["_g"], str(d.get(COMMENT_COL) or "").strip(), subs.get(uid), not a.no_late and d.get("遅延") != "あり")   # 採点の段階で遅延として下げた学生は、二重に下げない
        if down: lates.append((uid, d["_g"], g))
        val = "" if g is None else g if a.letters else pts[g]
        sig = {"grade": str(val), "comment": comment}
        if not a.resend and sent.get(uid) == sig: already += 1; continue
        todo.append((uid, d, sig))
    if a.only:                                        # 氏名または学生番号の一部で絞る(例: --only テスト用 でテスト用受講生だけに送る)
        todo = [t for t in todo if nname(a.only) in nname(t[1].get("氏名")) or nrm(a.only) in nrm(t[1].get("学生番号"))]
    if a.limit: todo = todo[: a.limit]
    print(f"課題: {head[col]} / コース {course} / 送信対象 {len(todo)} 人(送信済みでスキップ {already} 人)")
    nco = sum(d["_g"] is None for d in c["items"])
    print(f"  採点結果の学生 {len(c['items']) + len(c['skipped'])} 人 → 送る対象 {len(c['items'])} 人(S〜D {len(c['items']) - nco} 人、評価外でコメントのみ {nco} 人) → Canvasの成績表と照合できた {len(out_rows)} 人(成績表CSVの学生 {len(students)} 人)")
    if c["unmatched"] and not out_rows: print(f"  成績表CSVの見出し(学生を見分ける列): {[head[i] for i in c['ident']]}")
    if c["skipped"]: print(f"  評価がS〜Dでないため対象外: {len(c['skipped'])} 人(評価の内訳: {dict(Counter(d['_g'] or '空欄' for d in c['skipped']))})")
    if c["unmatched"]: print(f"  Canvasの成績表と照合できなかった学生: {len(c['unmatched'])} 人(学生番号の例: {', '.join(str(d.get('学生番号') or '空欄') for d in c['unmatched'][:3])})")
    if a.only and not todo: print(f"  --only {a.only} に一致する学生がいません")
    for uid, d, sig in todo[:3]:
        print(f"  例) 学生ID {uid}: 点数 {sig['grade'] or 'なし(コメントのみ)'} / コメント: {sig['comment'][:60].replace(chr(10), ' ')}")
    for uid, d, sig in todo:
        if not sig["grade"]: print(f"  コメントのみ) 学生ID {uid}: {sig['comment'][:60].replace(chr(10), ' ')}")
    if lates: print(f"  期限後の提出(遅延): {len(lates)} 人 → 評価を1段階下げ、その旨をコメントに書き足して送ります: " + ", ".join(f"学生ID {u}({g0}→{g1})" for u, g0, g1 in lates))
    if any((x.get("points_deducted") or 0) > 0 for x in subs.values()):
        print("  ※この課題では、Canvasの遅延ポリシーによる自動減点がかかっています。二重に下がらないよう、成績表の設定で遅延ポリシーの自動減点を切ってください(または --no-late で、このツールの減点を止めてください)。")
    if not a.send:
        print("(確認のみです。実際に送るには --send を付けてください。成績表の公開設定が手動なら、送信後も学生には見えません)"); return
    if not todo: print("送信する学生がいません"); return
    if not a.yes and not confirm_dialog(f"{len(todo)}人の点数とコメントを、Canvasに送信します。よろしいですか?"): print("キャンセルしました"); return
    ok = ng = 0
    for uid, d, sig in todo:
        body = {"submission[posted_grade]": sig["grade"]} if sig["grade"] else {}      # 評価外はコメントだけ
        prev = (sent.get(uid) or {}).get("comment", "")
        add = sig["comment"][len(prev):].strip() if prev and sig["comment"].startswith(prev) else sig["comment"]   # 送信済みのコメントに書き足しただけなら、足した部分だけを送る(同じコメントを二重に付けない)
        if add: body["comment[text_comment]"] = add
        req = urllib.request.Request(f"{base}/api/v1/courses/{course}/assignments/{aid}/submissions/{uid}",
                                     data=urllib.parse.urlencode(body).encode(), method="PUT",
                                     headers={"Authorization": f"Bearer {token.strip()}"})
        try:
            with urllib.request.urlopen(req, timeout=30): pass
            ok += 1; sent[uid] = sig
            sent_path.write_text(json.dumps(sent, ensure_ascii=False, indent=1), encoding="utf-8")
        except urllib.error.HTTPError as e:
            ng += 1; print(f"  失敗 学生ID {uid}: HTTP {e.code} {e.read()[:150]!r}")
            if e.code in (401, 403): print("認証または権限のエラーのため、中断します(トークンや権限を確認してください)"); break
        except Exception as e:
            ng += 1; print(f"  失敗 学生ID {uid}: {e}")
        time.sleep(0.3)
    print(f"送信完了: 成功 {ok} 人 / 失敗 {ng} 人。送信済みの記録: {sent_path.name}(同じ内容を二重に送らないために使います)")

def get_token():
    return os.environ.get("CANVAS_TOKEN") or osa('text returned of (display dialog "Canvasのアクセストークンを入力してください(保存されません)" default answer "" with hidden answer)')

def verify_feedback(a):
    """送ったはずの点数とコメントが、Canvas上に実際にあるかを学生ごとに照合する(送信漏れ・取り違えの検出)。"""
    c = prepare_export(a)
    if not c: return
    head, col, students, out_rows, pts = c["head"], c["col"], c["students"], c["out_rows"], c["pts"]
    if "ID" not in head: print("Canvasの成績表CSVに「ID」列がありません"); return
    idcol = head.index("ID"); aid = re.search(r"\((\d+)\)\s*$", head[col]).group(1)
    url = a.url or osa('text returned of (display dialog "CanvasのコースのURL(例: https://〜/courses/101839)を貼り付けてください" default answer "")')
    pc = parse_course_url(url)
    if not pc: print("コースのURLの形式が正しくありません"); return
    base, course = pc
    token = get_token()
    if not token: print("トークンが入力されなかったため中止します"); return
    try: subs = fetch_submissions(base, course, aid, token)
    except urllib.error.HTTPError as e: print(f"取得に失敗しました: HTTP {e.code} {e.read()[:150]!r}"); return
    flat = lambda t: re.sub(r"\s+", "", str(t or "").replace("\r\n", "\n"))
    res, ng = Counter(), []
    for mi, d in sorted(out_rows):
        uid = students[mi][idcol].strip()
        sub = subs.get(uid)
        g, comment, _ = apply_late(d["_g"], str(d.get(COMMENT_COL) or "").strip(), sub, not a.no_late and d.get("遅延") != "あり")
        val = "" if g is None else g if a.letters else pts[g]
        want = flat(comment)
        who = f"{d.get('学生番号') or ''} {d.get('氏名') or ''}".strip()
        if sub is None: res["提出情報なし"] += 1; ng.append((who, "Canvasに該当する提出情報がありません")); continue
        why = []
        got = sub.get("grade") if a.letters else (sub.get("entered_score") if sub.get("entered_score") is not None else sub.get("score"))   # Canvasの自動減点前の点数
        try: same = (str(got) == str(val)) if a.letters else (got is not None and float(got) == float(val))
        except (TypeError, ValueError): same = False
        if val == "": same = True                          # コメントのみの学生は点数を照合しない
        if not same: why.append(f"点数が違う(期待 {val} / Canvas {got})")
        got_c = "".join(flat(x.get("comment")) for x in sub.get("submission_comments", []))        # 書き足しで2つに分かれたコメントもまとめて照合
        if want and want not in got_c and not all(flat(q) in got_c for q in comment.split("\n※")): why.append("コメントが見つからない")
        if why: res["要確認"] += 1; ng.append((who, "、".join(why)))
        else: res["一致"] += 1
    print(f"照合結果 {len(out_rows)} 人: 一致 {res['一致']} / 要確認 {res['要確認']} / 提出情報なし {res['提出情報なし']}")
    for who, why in ng: print(f"  要確認 {who}: {why}")
    if c["unmatched"]: print(f"採点結果にあるが、Canvasの成績表と照合できなかった学生が {len(c['unmatched'])} 人います(送信されていません)")
    if c["skipped"]: print(f"評価がS〜Dでないため対象外の学生が {len(c['skipped'])} 人います")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder", nargs="?"); ap.add_argument("--ref"); ap.add_argument("--out", default="out")
    ap.add_argument("--chapter", type=int, help="ファイル名から章を判別できないときに使う章番号(--export では対象の章)")
    ap.add_argument("--export", action="store_true", help="採点結果.xlsxからCanvas取り込み用CSVを作る")
    ap.add_argument("--letters", action="store_true", help="--export で、点数ではなくS〜Dの文字をそのまま出力する(試験用)")
    ap.add_argument("--post", action="store_true", help="点数とコメントをCanvas APIで送る(--send を付けない限り、確認のみ)")
    ap.add_argument("--verify", action="store_true", help="送った点数・コメントがCanvas上にあるかを照合する")
    ap.add_argument("--send", action="store_true"); ap.add_argument("--yes", action="store_true"); ap.add_argument("--resend", action="store_true")
    ap.add_argument("--no-late", action="store_true", help="遅延の提出でも評価を下げない(Canvasの遅延ポリシーを使う場合など)")
    ap.add_argument("--url"); ap.add_argument("--limit", type=int); ap.add_argument("--only", help="氏名または学生番号の一部で、送る学生を絞る")
    ap.add_argument("--results"); ap.add_argument("--canvas"); ap.add_argument("--column"); ap.add_argument("--csv-out")
    ap.add_argument("--version", action="version", version=f"Eco-Analysis Grader {GRADER_VERSION}")
    a = ap.parse_args()
    print(f"Eco-Analysis Grader  バージョン {GRADER_VERSION}")
    if a.export: return export_canvas(a)
    if a.post: return post_feedback(a)
    if a.verify: return verify_feedback(a)
    load_plugins()
    if a.folder:
        sub, ref, outdir = Path(a.folder), (Path(a.ref) if a.ref else None), Path(a.out)
    else:                                            # 引数なし → ダイアログで指定
        sub = pick_folder("提出フォルダ(学生のxlsxが入っているフォルダ)を選んでください", "/Volumes")
        if not sub: print("提出フォルダが選ばれなかったため中止します"); return
        ref = pick_folder("模範解答フォルダを選んでください(使わない場合はキャンセル)", str(sub.parent))
        if not ref: print("模範解答フォルダなしで実行します")
        outdir = pick_folder("採点結果の出力先フォルダを選んでください", str(sub.parent))
        if not outdir: print("出力先が選ばれなかったため中止します"); return
    refs = {}
    if ref:
        for n in CHECKERS:
            hit = sorted(ref.glob(f"*ch{n:02d}*.xlsx"))
            if hit: refs[n] = hit[0]
    files = sorted(p for p in sub.rglob("*.xlsx") if not p.name.startswith(("~$", ".")))
    outdir.mkdir(parents=True, exist_ok=True)
    if a.chapter is None:       # フォルダの章は、ファイル名に章が書かれているものの最頻値から自動で決める
        found = Counter(parse_name(p.stem)[2] or chapter_from_path(p) for p in files) - Counter({None: 10**9})
        found.pop(None, None)
        if found:
            a.chapter = found.most_common(1)[0][0]
            print(f"フォルダの章を第{a.chapter}章と判断しました(ファイル名に章がないファイルも第{a.chapter}章として処理します)")
        elif not a.folder:
            a.chapter = ask_chapter()
    rows, longrows = [], []
    for p in files:
        sid, name, ch, fn_ok = parse_name(p.stem, chapter_from_path(p) or a.chapter)
        base = {"学生番号": sid, "氏名": name, "課題": ch or "", "ファイル": p.name}
        if ch not in CHECKERS:
            rows.append(base | {"採点結果": "未対応", "できていない箇所": "", "できていない箇所コード": ""}); continue
        try:
            g, iss, det = CHECKERS[ch](p, refs.get(ch))
        except Exception as e:
            g, iss, det = "評価外", [], {"error": str(e)[:80]}
        if not fn_ok: iss = iss + ["FILENAME"]
        g, iss, late = apply_file_late(p, g, iss)
        rows.append(base | {"採点結果": g, "遅延": "あり" if late else "", "できていない箇所": "、".join(label_of(c) for c in iss),
                            "できていない箇所コード": ";".join(iss)} | det)
        longrows += [{"学生番号": sid, "氏名": name, "課題": ch, "コード": c.partition("@")[0], "内容": label_of(c)} for c in iss]
    chs_done = Counter(r["課題"] for r in rows if r["課題"] != "")
    base = f"採点結果_第{chs_done.most_common(1)[0][0]}章" if chs_done else "採点結果"
    out = outdir / f"{base}.xlsx"
    if out.exists():                                  # 確認・修正済みのファイルを上書きしないよう、日時を付けて別名にする
        from datetime import datetime
        out = outdir / f"{base}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        print(f"同名のファイルがあるため、別名で保存します: {out.name}")
    write_xlsx(out, rows, longrows)
    print(f"処理 {len(rows)} 件 / 未対応 {sum(r['採点結果']=='未対応' for r in rows)} 件 → {out}")
    if not a.folder:
        try: subprocess.run(["open", "-R", str(out)])      # Finderで結果ファイルを表示
        except Exception: pass

if __name__ == "__main__":
    main()
