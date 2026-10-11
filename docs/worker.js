// 採点の本体。画面(app.js)とは別のスレッドで動き、学生のファイルはこのパソコンの中だけで処理される。
import { loadPyodide } from "./pyodide/pyodide.mjs";
import "./core.js";

const here = (p) => new URL(p, import.meta.url);
let py = null;

async function init() {
  postMessage({ type: "status", text: "採点の準備をしています(1/3)Pyodide(ブラウザ上のPython)を読み込み中…" });
  const idx = here("./pyodide/");
  py = await loadPyodide({ indexURL: idx.protocol === "file:" ? decodeURIComponent(idx.pathname) : idx.href });   // file: はNodeでのテスト用
  postMessage({ type: "status", text: "採点の準備をしています(2/3)採点のプログラムを読み込み中…" });
  const manifest = await (await fetch(here("./py/manifest.json"), { cache: "no-cache" })).json();
  const v = "?v=" + encodeURIComponent(manifest.version);
  const buf = async (u) => (await fetch(here(u + v))).arrayBuffer();
  const text = async (u) => (await fetch(here(u + v))).text();
  const assets = {
    wheels: await Promise.all(manifest.wheels.map((w) => buf("./py/wheels/" + w))),
    grader: await text("./py/grader.py"),
    webcore: await text("./py/webcore.py"),
    checkers: {},
  };
  for (const n of manifest.checkers) assets.checkers[n] = await text("./py/checkers/" + n);
  postMessage({ type: "status", text: "採点の準備をしています(3/3)Excelを読むライブラリを準備中…" });
  await self.GraderCore.setup(py, assets);
  postMessage({ type: "ready", version: manifest.version });
}

self.onmessage = async (e) => {
  const m = e.data;
  try {
    if (m.type === "run") {
      const job = { files: m.files, ref: m.ref, chapter: m.chapter };
      const res = await self.GraderCore.run(py, job);
      postMessage({ type: "done", xlsx: res.xlsx, summary: res.summary }, [res.xlsx.buffer]);
    }
  } catch (err) {
    postMessage({ type: "error", message: String(err && err.message ? err.message : err) });
  }
};

init().catch((err) => postMessage({ type: "error", message: "準備に失敗しました: " + String(err && err.message ? err.message : err) }));
