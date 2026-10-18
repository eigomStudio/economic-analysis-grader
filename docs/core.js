/* Pyodide上でPython(grader.py)を動かすための共通処理。画面(app.js)とテスト(Node)の両方から使う。
   学生のファイルは、このブラウザ(PC)の中のメモリだけで処理され、ネットワークには送られない。 */
(function (root) {
  async function setup(py, assets) {
    for (const w of assets.wheels) py.unpackArchive(new Uint8Array(w), "wheel");          // openpyxl など(同じサイトに置いたファイル)
    py.FS.mkdirTree("/home/pyodide/checkers");
    py.FS.writeFile("/home/pyodide/grader.py", assets.grader);
    py.FS.writeFile("/home/pyodide/webcore.py", assets.webcore);
    for (const [name, text] of Object.entries(assets.checkers || {})) py.FS.writeFile("/home/pyodide/checkers/" + name, text);
    py.runPython("import sys; sys.path.insert(0, '/home/pyodide')");
    py.runPython("import webcore");
  }

  async function run(py, job) {
    py.runPython("import shutil\nfor d in ('/tmp/in', '/tmp/ref', '/tmp/out'): shutil.rmtree(d, ignore_errors=True)");
    py.FS.mkdirTree("/tmp/in"); py.FS.mkdirTree("/tmp/ref"); py.FS.mkdirTree("/tmp/out");
    job.files.forEach((f, i) => {                                           // 同じ名前のファイルがあっても混ざらないよう、番号のフォルダに入れる
      const rel = String(f.path || f.name).split("/").filter((x) => x && x !== "." && x !== "..");   // フォルダ名も残す(課題名・「遅延」の判定に使う)
      const dir = ["/tmp/in", String(i)].concat(rel.slice(0, -1)).join("/");
      py.FS.mkdirTree(dir);
      py.FS.writeFile(dir + "/" + rel[rel.length - 1], new Uint8Array(f.data));
    });
    let refPath = "";
    if (job.ref) { refPath = "/tmp/ref/" + job.ref.name; py.FS.writeFile(refPath, new Uint8Array(job.ref.data)); }
    py.globals.set("job_ref", refPath);
    py.globals.set("job_chapter", job.chapter || 0);
    const json = py.runPython(
      "import json, webcore\n" +
      "json.dumps(webcore.run('/tmp/in', '/tmp/out/result.xlsx', job_ref or None, int(job_chapter) or None), ensure_ascii=False)");
    const summary = JSON.parse(json);
    if (summary.need_chapter) return { xlsx: null, summary };            // 章番号が分からないとき(結果のExcelは作らない)
    return { xlsx: py.FS.readFile("/tmp/out/result.xlsx"), summary };
  }

  const api = { setup, run };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.GraderCore = api;
})(typeof self !== "undefined" ? self : this);
