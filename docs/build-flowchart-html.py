# -*- coding: utf-8 -*-
"""把 FLOWCHART.md 渲染成可离线查看的 flowchart.html。

单一真源是 FLOWCHART.md：本脚本读取它，把 Markdown 内嵌进 HTML 模板，
页面加载时用本地 vendor/marked.min.js 解析、vendor/mermaid.min.js 画图。
改了流程图只需重新运行：  py -3 docs/build-flowchart-html.py
"""
from __future__ import annotations

import io
import json
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))

TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Email Manager v3.0.0 · 程序完整流程图</title>
<style>
  :root { color-scheme: light; }
  body { margin: 0; font-family: "Segoe UI", "Microsoft YaHei", system-ui, sans-serif;
         background: #f6f7f9; color: #1f2328; }
  header { position: sticky; top: 0; z-index: 10; background: #0f172a; color: #fff;
           padding: 12px 28px; box-shadow: 0 2px 10px rgba(15, 23, 42, .18); }
  header h1 { font-size: 16px; margin: 0; font-weight: 600; letter-spacing: .2px; }
  header span { opacity: .72; font-size: 12px; margin-left: 12px; }
  main { max-width: 1180px; margin: 26px auto 90px; padding: 0 20px; }
  h1, h2 { border-bottom: 1px solid #e5e7eb; padding-bottom: 6px; }
  h2 { margin-top: 52px; font-size: 21px; }
  p, li { line-height: 1.65; }
  .mermaid { background: #fff; border: 1px solid #e5e7eb; border-radius: 10px;
             padding: 18px; margin: 18px 0 30px; overflow: auto; text-align: center; }
  .mermaid svg { max-width: 100%; height: auto; }
  table { border-collapse: collapse; width: 100%; background: #fff; font-size: 14px; margin: 14px 0 26px; }
  th, td { border: 1px solid #e5e7eb; padding: 8px 10px; text-align: left; vertical-align: top; }
  th { background: #f1f5f9; }
  code { background: #f1f5f9; padding: 1px 5px; border-radius: 4px; font-size: 13px; }
  blockquote { border-left: 4px solid #cbd5e1; margin: 14px 0; padding: 6px 14px;
               color: #475569; background: #fff; }
  #err { display: none; background: #fee2e2; border: 1px solid #ef4444; color: #7f1d1d;
         padding: 12px 14px; margin: 16px 0; border-radius: 8px; white-space: pre-wrap; }
</style>
</head>
<body>
<header>
  <h1>Email Manager v3.0.0 · 程序完整流程图
    <span>依据 commit 45ed398 · 离线 Mermaid 渲染</span></h1>
</header>
<main>
  <div id="err"></div>
  <div id="content"></div>
</main>
<script src="vendor/marked.min.js"></script>
<script src="vendor/mermaid.min.js"></script>
<script>
const MD = __MD_JSON__;
marked.use({
  renderer: {
    code(code, infostring) {
      const lang = (infostring || "").trim().split(/\\s+/)[0];
      if (lang === "mermaid") return '<div class="mermaid">' + code + "</div>";
      const esc = code.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
      return "<pre><code>" + esc + "</code></pre>";
    }
  }
});
document.getElementById("content").innerHTML = marked.parse(MD);

// marked v4 不给标题生成 id —— 按 GitHub 的 slug 规则补上，让目录锚点可用。
function slugify(text) {
  return text.trim().toLowerCase()
    .replace(/[^\\p{L}\\p{N} \\-_]/gu, "")   // 去掉标点（. ： （ ） / 等）
    .replace(/ /g, "-");                      // 每个空格变一个连字符（与 GitHub 一致）
}
const usedIds = new Set();
document.querySelectorAll("#content h1, #content h2, #content h3").forEach((h) => {
  let id = slugify(h.textContent);
  let unique = id, n = 2;
  while (usedIds.has(unique)) { unique = id + "-" + n++; }
  usedIds.add(unique);
  h.id = unique;
});

mermaid.initialize({
  startOnLoad: false,
  theme: "default",
  securityLevel: "loose",
  flowchart: { htmlLabels: true, curve: "basis", nodeSpacing: 34, rankSpacing: 46 },
  sequence: { useMaxWidth: true },
  themeVariables: { fontFamily: '"Segoe UI", "Microsoft YaHei", sans-serif' }
});
mermaid.run({ querySelector: ".mermaid" }).then(() => {
  window.__renderOk = true;
  document.title = "Email Manager 流程图 · 渲染完成";
}).catch((e) => {
  const box = document.getElementById("err");
  box.style.display = "block";
  box.textContent = "Mermaid 渲染错误：" + (e && e.message ? e.message : String(e));
  document.title = "RENDER_ERROR";
});
</script>
</body>
</html>
"""


def main() -> int:
    md_path = os.path.join(HERE, "FLOWCHART.md")
    out_path = os.path.join(HERE, "flowchart.html")
    with io.open(md_path, "r", encoding="utf-8") as handle:
        markdown = handle.read()
    mermaid_blocks = len(re.findall(r"```mermaid", markdown))
    html = TEMPLATE.replace("__MD_JSON__", json.dumps(markdown, ensure_ascii=False))
    with io.open(out_path, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(html)
    print("wrote %s (%d mermaid diagrams, %.1f KB markdown)"
          % (out_path, mermaid_blocks, len(markdown.encode("utf-8")) / 1024))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
