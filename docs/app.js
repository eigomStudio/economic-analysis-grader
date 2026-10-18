const $ = (id) => document.getElementById(id);
let worker = null, failed = false;
let ready = false, pending = false, running = false, t0 = performance.now();

const setStatus = (t) => { $("status").textContent = t; };
const GRADES = ["S", "A", "B", "C", "D", "評価外", "未対応"];
window.__ecoApp = true;                       // index.html の見張りに「このプログラムは動き出した」と知らせる
try {
  worker = new Worker(new URL("./worker.js", import.meta.url), { type: "module" });
} catch (err) {                                // 古いブラウザなどで、Workerを起動できないとき(以降の画面の動きは止めない)
  failed = true;
  worker = { postMessage() {}, set onerror(v) {}, set onmessage(v) {} };
  setStatus("エラー: 採点の部品を起動できませんでした。最新のChrome、Edge、Safariで、公開されたURL(https://…)から開いてください。(" + (err && err.message) + ")");
}

if (location.protocol === "file:") {
  setStatus("エラー: このページは、ファイルを直接開くと動きません。公開されたURL(https://…)から開くか、README.mdの「自分のパソコンで試す」の手順で開いてください。");
  $("run").disabled = true;
}
worker.onerror = (e) => {                       // 部品(Pyodideなど)を読み込めなかったとき
  setStatus("エラー: 採点の部品を読み込めませんでした。ネットワークの制限や、古いブラウザの可能性があります。(" + (e.message || "詳細なし") + ")");
  pending = false; running = false;
};
setTimeout(() => {
  if (!ready && !$("status").textContent.startsWith("エラー"))
    setStatus($("status").textContent + "(時間がかかっています。回線が遅いか、読み込みが止まっている可能性があります。ページを再読み込みしても直らない場合は、F12キーのコンソールの表示を教えてください。)");
}, 60000);

worker.onmessage = (e) => {
  const m = e.data;
  if (m.type === "status") setStatus(m.text);
  else if (m.type === "ready") {
    ready = true;
    setStatus(`準備ができました(${((performance.now() - t0) / 1000).toFixed(1)}秒)。ファイルを選んで「採点する」を押してください。`);
    $("version").textContent = "バージョン " + m.version;
    if (pending) { pending = false; start(); }
  } else if (m.type === "done") showResult(m);
  else if (m.type === "error") { running = false; $("run").disabled = false; setStatus("エラー: " + m.message); }
};

function selectedFiles() {
  const all = [...$("folder").files, ...$("files").files];
  return all.filter((f) => /\.(xlsx|zip)$/i.test(f.name) && !f.name.startsWith("~$") && !f.name.startsWith("."));
}

function guessChapter(files) {               // ファイル名・フォルダ名の「課題2」「第2章」から章を推定する(いちばん多いもの)
  const cnt = {};
  for (const f of files) {
    const t = (f.webkitRelativePath || f.name).normalize("NFKC");
    const m = t.match(/課題\s*(\d+)/) || t.match(/第\s*(\d+)\s*章/);
    if (m) cnt[m[1]] = (cnt[m[1]] || 0) + 1;
  }
  const best = Object.entries(cnt).sort((a, b) => b[1] - a[1])[0];
  return best ? best[0] : "";
}

function updateSelection() {
  const files = selectedFiles(), n = files.length;
  $("selected").textContent = n ? `${n} 個のファイル(.xlsx / .zip)を選択中` : "ファイルが選ばれていません";
  const ch = guessChapter(files);
  if (ch) { $("chapter").value = ch; $("chapter-box").hidden = false; $("chapter-note").textContent = "(ファイル名・フォルダ名から読み取りました)"; }
  else { $("chapter-note").textContent = ""; }
}
$("folder").addEventListener("change", updateSelection);
$("files").addEventListener("change", updateSelection);

$("run").addEventListener("click", () => {
  if (failed) return;
  if (!selectedFiles().length) { setStatus("提出ファイルを選んでください。"); return; }
  if (!ready) { pending = true; setStatus("準備が終わったら、自動で採点を始めます。" ); return; }
  start();
});

async function start() {
  if (running) return;
  const files = selectedFiles();
  if (!files.length) { setStatus("提出ファイルを選んでください。"); return; }
  running = true; $("run").disabled = true; $("result").hidden = true;
  setStatus(`${files.length} 個のファイルを読み込んでいます…`);
  const items = [], transfer = [];
  // フォルダで選んだときは、フォルダ名(課題名・「遅延」)も渡す
  for (const f of files) { const data = await f.arrayBuffer(); items.push({ name: f.name, path: f.webkitRelativePath || f.name, data }); transfer.push(data); }
  let ref = null;
  if ($("ref") && $("ref").files[0]) { const f = $("ref").files[0]; const data = await f.arrayBuffer(); ref = { name: f.name, data }; transfer.push(data); }
  const chapter = parseInt($("chapter").value, 10) || 0;
  t0 = performance.now();
  setStatus("採点しています…");
  worker.postMessage({ type: "run", files: items, ref, chapter }, transfer);
}

function showResult(m) {
  const s = m.summary;
  if (s.need_chapter) {                          // ファイル名から章が分からない → 章番号の欄を出して、入力してもらう
    running = false; $("run").disabled = false;
    $("chapter-box").hidden = false; $("chapter-note").textContent = ""; $("chapter").focus();
    setStatus(`${s.count} 個のファイルの名前から、課題の章番号が分かりませんでした。「課題の章番号」に番号を入力して、もう一度「採点する」を押してください。`);
    return;
  }
  const blob = new Blob([m.xlsx], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = $("download");
  if (a.dataset.url) URL.revokeObjectURL(a.dataset.url);
  a.dataset.url = URL.createObjectURL(blob);
  a.href = a.dataset.url;
  a.download = s.chapter ? `採点結果_第${s.chapter}章.xlsx` : "採点結果.xlsx";
  const rows = GRADES.filter((g) => s.grades[g]).map((g) => `<tr><th>${g}</th><td>${s.grades[g]}</td></tr>`).join("");
  $("summary").innerHTML = `<p>${s.count} 件を処理しました(第${s.chapter || "?"}章、未対応 ${s.unsupported} 件、${s.seconds} 秒)。<br><span class="small">採点プログラムのバージョン: ${s.version || "不明"}</span></p><table>${rows}</table>`;
  if (s.chapter) { $("chapter").value = s.chapter; $("chapter-box").hidden = false; }   // 採点に使った章を表示する
  $("log").textContent = (s.log || []).join("\n");
  $("result").hidden = false; $("run").disabled = false; running = false;
  setStatus("採点が終わりました。下のボタンから結果のExcelをダウンロードしてください。");
}
