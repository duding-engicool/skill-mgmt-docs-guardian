# -*- coding: utf-8 -*-
"""生成自包含的单文件 HTML 关系图（内嵌力导向布局，离线可用，双击即开）。

不依赖任何 CDN / 外部 JS 库，不发起网络请求。
"""
from __future__ import annotations

import json
import os
from typing import Any, Dict, Set

from lib.core import INDEX_DIR

_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>体系文件关系图</title>
<style>
  * { box-sizing: border-box; }
  body { margin:0; font-family: "Microsoft YaHei","Segoe UI",sans-serif; background:#f4f6f9; color:#1f2933; }
  #toolbar { display:flex; align-items:center; gap:10px; padding:10px 14px; background:#fff;
             border-bottom:1px solid #dde3ea; flex-wrap:wrap; }
  #toolbar h1 { font-size:15px; margin:0 12px 0 0; font-weight:600; color:#1a3a5c; }
  button, select, input { font-family:inherit; font-size:13px; padding:5px 10px;
             border:1px solid #c7d0da; border-radius:4px; background:#fff; color:#1f2933; }
  button { cursor:pointer; } button:hover { background:#eef3f8; }
  button.active { background:#1a3a5c; color:#fff; border-color:#1a3a5c; }
  .sep { width:1px; height:20px; background:#dde3ea; }
  #main { display:flex; height:calc(100vh - 51px); }
  #cvwrap { flex:1; position:relative; overflow:hidden; }
  canvas { display:block; width:100%; height:100%; cursor:grab; }
  canvas.dragging { cursor:grabbing; }
  #panel { width:300px; background:#fff; border-left:1px solid #dde3ea; padding:14px;
           overflow-y:auto; font-size:13px; line-height:1.7; }
  #panel h2 { font-size:14px; margin:0 0 8px; color:#1a3a5c; }
  #panel .row { margin-bottom:6px; word-break:break-all; }
  #panel .k { color:#6b7684; font-size:12px; }
  #panel ul { padding-left:18px; margin:4px 0; }
  #panel li { margin-bottom:3px; cursor:pointer; color:#2b5f8e; }
  #panel li:hover { text-decoration:underline; }
  #legend { position:absolute; left:12px; bottom:12px; background:rgba(255,255,255,.92);
            border:1px solid #dde3ea; border-radius:4px; padding:8px 10px; font-size:12px; line-height:1.8; }
  .dot { display:inline-block; width:10px; height:10px; border-radius:50%; margin-right:5px; vertical-align:-1px; }
  #status { position:absolute; right:12px; bottom:12px; background:rgba(255,255,255,.92);
            border:1px solid #dde3ea; border-radius:4px; padding:6px 10px; font-size:12px; color:#52606d; }
  .muted { color:#8895a4; }
</style>
</head>
<body>
<div id="toolbar">
  <h1>体系文件关系图</h1>
  <button id="btn-all" class="active">全量视图</button>
  <button id="btn-chg">变更影响视图</button>
  <div class="sep"></div>
  <input id="search" type="text" placeholder="搜索编号/标题/路径…" style="width:190px">
  <button id="btn-find">定位</button>
  <div class="sep"></div>
  <button id="btn-relayout">重新布局</button>
  <button id="btn-fit">适应窗口</button>
  <span class="muted" id="meta"></span>
</div>
<div id="main">
  <div id="cvwrap">
    <canvas id="cv"></canvas>
    <div id="legend">
      <div><span class="dot" style="background:#1a3a5c"></span>普通文件</div>
      <div><span class="dot" style="background:#d9534f"></span>本次变更（高亮描边）</div>
      <div><span class="dot" style="background:#9aa5b1"></span>未解析（盲区）</div>
      <div><span class="dot" style="background:#e08a2e"></span>变更相关连线</div>
    </div>
    <div id="status"></div>
  </div>
  <div id="panel"><h2>说明</h2><div class="row">点击节点查看引用关系；滚轮缩放，拖拽平移；拖动节点可固定位置。</div></div>
</div>
<script>
const DATA = __DATA__;
const canvas = document.getElementById('cv');
const ctx = canvas.getContext('2d');
const wrap = document.getElementById('cvwrap');

// ---------- 数据准备 ----------
const nodes = DATA.nodes.map((n, i) => ({
  ...n, idx: i, x: 0, y: 0, vx: 0, vy: 0, fixed: false
}));
const byId = {};
nodes.forEach(n => byId[n.id] = n);
const edges = DATA.edges.filter(e => byId[e.source] && byId[e.target]).map(e => ({
  ...e, s: byId[e.source], t: byId[e.target]
}));
const deg = {};
edges.forEach(e => { deg[e.source] = (deg[e.source]||0)+1; deg[e.target] = (deg[e.target]||0)+1; });

// group 配色
const groups = [...new Set(nodes.map(n => n.group))].sort();
const gcolor = {};
groups.forEach((g, i) => {
  const hue = (200 + i * 47) % 360;
  gcolor[g] = 'hsl(' + hue + ',42%,52%)';
});

let view = 'all';        // all | changed
let visible = nodes.slice();
let scale = 1, offsetX = 0, offsetY = 0;
let hover = null, selected = null, drag = null;
let alpha = 1;

function changedSet() {
  return new Set(nodes.filter(n => n.changed).map(n => n.id));
}
function neighborsOf(set) {
  const out = new Set();
  edges.forEach(e => {
    if (set.has(e.source)) out.add(e.target);
    if (set.has(e.target)) out.add(e.source);
  });
  return out;
}
function applyView() {
  if (view === 'all') { visible = nodes.slice(); }
  else {
    const c = changedSet();
    const nb = neighborsOf(c);
    visible = nodes.filter(n => c.has(n.id) || nb.has(n.id));
  }
  visible.forEach((n, i) => {
    if (n.x === 0 && n.y === 0) {
      const a = i / Math.max(1, visible.length) * Math.PI * 2;
      const r = 120 + Math.sqrt(visible.length) * 26;
      n.x = Math.cos(a) * r; n.y = Math.sin(a) * r;
    }
  });
  alpha = 1;
  document.getElementById('meta').textContent =
    '显示 ' + visible.length + ' / ' + nodes.length + ' 个文件，' +
    edges.filter(e => isVisible(e.s) && isVisible(e.t)).length + ' 条引用关系';
  layout(260);
}
function isVisible(n) { return visible.indexOf(n) >= 0; }

// ---------- 力导向 ----------
function layout(steps) {
  const K = 2600, L = 95, CENTER = 0.004, DAMP = 0.86;
  for (let s = 0; s < steps; s++) {
    const a = alpha * (1 - s / steps);
    if (a <= 0.01) break;
    // 斥力
    for (let i = 0; i < visible.length; i++) {
      for (let j = i + 1; j < visible.length; j++) {
        const p = visible[i], q = visible[j];
        let dx = q.x - p.x, dy = q.y - p.y;
        let d2 = dx*dx + dy*dy;
        if (d2 < 1) { dx = (Math.random()-0.5); dy = (Math.random()-0.5); d2 = 1; }
        const d = Math.sqrt(d2);
        const f = Math.min(K / d2, 40) * a;
        const fx = dx / d * f, fy = dy / d * f;
        p.vx -= fx; p.vy -= fy; q.vx += fx; q.vy += fy;
      }
    }
    // 弹簧
    for (const e of edges) {
      if (!isVisible(e.s) || !isVisible(e.t)) continue;
      let dx = e.t.x - e.s.x, dy = e.t.y - e.s.y;
      const d = Math.max(1, Math.hypot(dx, dy));
      const f = (d - L) * 0.012 * a;
      const fx = dx / d * f, fy = dy / d * f;
      e.s.vx += fx; e.s.vy += fy; e.t.vx -= fx; e.t.vy -= fy;
    }
    // 中心引力 + 积分
    for (const n of visible) {
      n.vx += -n.x * CENTER * a; n.vy += -n.y * CENTER * a;
      if (n.fixed) { n.vx = n.vy = 0; continue; }
      n.vx *= DAMP; n.vy *= DAMP;
      const sp = Math.hypot(n.vx, n.vy);
      if (sp > 12) { n.vx = n.vx / sp * 12; n.vy = n.vy / sp * 12; }
      n.x += n.vx; n.y += n.vy;
    }
  }
  draw();
}

function resize() {
  const r = wrap.getBoundingClientRect();
  const dpr = window.devicePixelRatio || 1;
  canvas.width = r.width * dpr; canvas.height = r.height * dpr;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  draw();
}

function radius(n) { return 4.5 + Math.min(9, Math.sqrt(deg[n.id] || 0) * 2.6); }

function draw() {
  const w = canvas.clientWidth, h = canvas.clientHeight;
  ctx.save();
  ctx.clearRect(0, 0, w, h);
  ctx.translate(w / 2 + offsetX, h / 2 + offsetY);
  ctx.scale(scale, scale);

  // 边
  for (const e of edges) {
    if (!isVisible(e.s) || !isVisible(e.t)) continue;
    const hot = e.hot || (selected && (selected.id === e.source || selected.id === e.target));
    ctx.strokeStyle = hot ? '#e08a2e' : (e.kind === 'hierarchy' ? '#9ec5e8' : '#c3ccd6');
    ctx.lineWidth = hot ? 1.9 : 1;
    if (e.kind === 'hierarchy') ctx.setLineDash([4, 3]); else ctx.setLineDash([]);
    ctx.beginPath(); ctx.moveTo(e.s.x, e.s.y); ctx.lineTo(e.t.x, e.t.y); ctx.stroke();
  }
  ctx.setLineDash([]);

  // 节点
  for (const n of visible) {
    const r = radius(n);
    const isSel = selected && selected.id === n.id;
    const isHov = hover && hover.id === n.id;
    ctx.beginPath(); ctx.arc(n.x, n.y, r, 0, Math.PI * 2);
    ctx.fillStyle = n.status !== 'ok' ? '#9aa5b1' : gcolor[n.group];
    ctx.fill();
    ctx.lineWidth = n.changed ? 3 : (isSel || isHov ? 2.5 : 1);
    ctx.strokeStyle = n.changed ? '#d9534f' : (isSel ? '#1a3a5c' : (isHov ? '#52606d' : '#ffffff'));
    if (n.status !== 'ok') ctx.setLineDash([3, 2]);
    ctx.stroke();
    ctx.setLineDash([]);

    if (scale > 0.55 || n.changed || isSel || isHov) {
      ctx.font = (n.changed ? '600 ' : '') + Math.max(9, Math.min(13, 11 / scale)) + 'px "Microsoft YaHei",sans-serif';
      ctx.fillStyle = n.changed ? '#b8322e' : '#39414d';
      ctx.textAlign = 'center';
      ctx.fillText(n.label.length > 18 ? n.label.slice(0, 17) + '…' : n.label, n.x, n.y + r + 11);
    }
  }
  ctx.restore();
  document.getElementById('status').textContent =
    '缩放 ' + scale.toFixed(2) + 'x　节点 ' + visible.length;
}

// ---------- 交互 ----------
function toWorld(mx, my) {
  const w = canvas.clientWidth, h = canvas.clientHeight;
  return { x: (mx - w / 2 - offsetX) / scale, y: (my - h / 2 - offsetY) / scale };
}
function hit(mx, my) {
  const p = toWorld(mx, my);
  let best = null, bd = 1e9;
  for (const n of visible) {
    const d = Math.hypot(n.x - p.x, n.y - p.y);
    if (d < radius(n) + 5 && d < bd) { bd = d; best = n; }
  }
  return best;
}

canvas.addEventListener('mousedown', e => {
  const r = canvas.getBoundingClientRect();
  const mx = e.clientX - r.left, my = e.clientY - r.top;
  const n = hit(mx, my);
  if (n) { drag = { node: n, moved: false, mx, my }; select(n); }
  else { drag = { pan: true, mx, my, ox: offsetX, oy: offsetY }; }
});
window.addEventListener('mousemove', e => {
  const r = canvas.getBoundingClientRect();
  const mx = e.clientX - r.left, my = e.clientY - r.top;
  if (drag && drag.node) {
    const p0 = toWorld(drag.mx, drag.my), p1 = toWorld(mx, my);
    drag.node.x += p1.x - p0.x; drag.node.y += p1.y - p0.y;
    drag.node.vx = drag.node.vy = 0; drag.node.fixed = true;
    drag.mx = mx; drag.my = my; drag.moved = true;
    canvas.classList.add('dragging');
    draw();
  } else if (drag && drag.pan) {
    offsetX = drag.ox + (mx - drag.mx); offsetY = drag.oy + (my - drag.my);
    draw();
  } else {
    const n = hit(mx, my);
    if (n !== hover) { hover = n; draw(); }
  }
});
window.addEventListener('mouseup', () => { drag = null; canvas.classList.remove('dragging'); });
canvas.addEventListener('wheel', e => {
  e.preventDefault();
  const r = canvas.getBoundingClientRect();
  const mx = e.clientX - r.left, my = e.clientY - r.top;
  const before = toWorld(mx, my);
  scale = Math.min(3, Math.max(0.15, scale * (e.deltaY < 0 ? 1.12 : 0.89)));
  const after = toWorld(mx, my);
  offsetX += (after.x - before.x) * scale; offsetY += (after.y - before.y) * scale;
  draw();
}, { passive: false });

function select(n) {
  selected = n;
  const out = edges.filter(e => e.source === n.id);
  const inn = edges.filter(e => e.target === n.id);
  const li = arr => arr.map(e => {
    const other = byId[e.source === n.id ? e.target : e.source];
    return '<li data-id="' + other.id + '">' + (other.label) + '</li>';
  }).join('') || '<li class="muted">无</li>';
  document.getElementById('panel').innerHTML =
    '<h2>' + n.label + '</h2>' +
    '<div class="row"><span class="k">编号</span> ' + (n.label) + '</div>' +
    '<div class="row"><span class="k">标题</span> ' + n.title + '</div>' +
    '<div class="row"><span class="k">路径</span> ' + n.path + '</div>' +
    '<div class="row"><span class="k">责任人</span> ' + (n.owner || '—') + '</div>' +
    '<div class="row"><span class="k">状态</span> ' + (n.status === 'ok' ? '已解析' : '未解析（盲区）') +
      (n.changed ? '　<b style="color:#b8322e">本次变更</b>' : '') + '</div>' +
    '<div class="row"><span class="k">引用了（' + out.length + '）</span><ul>' + li(out) + '</ul></div>' +
    '<div class="row"><span class="k">被引用（' + inn.length + '）</span><ul>' + li(inn) + '</ul></div>';
  document.getElementById('panel').querySelectorAll('li[data-id]').forEach(el => {
    el.onclick = () => select(byId[el.dataset.id]);
  });
  draw();
}

document.getElementById('btn-all').onclick = () => {
  view = 'all';
  document.getElementById('btn-all').classList.add('active');
  document.getElementById('btn-chg').classList.remove('active');
  applyView();
};
document.getElementById('btn-chg').onclick = () => {
  view = 'changed';
  document.getElementById('btn-chg').classList.add('active');
  document.getElementById('btn-all').classList.remove('active');
  applyView();
};
document.getElementById('btn-relayout').onclick = () => {
  visible.forEach(n => { n.fixed = false; n.vx = n.vy = 0; });
  alpha = 1; layout(400); fit();
};
document.getElementById('btn-fit').onclick = () => fit();
document.getElementById('btn-find').onclick = () => doSearch();
document.getElementById('search').addEventListener('keydown', e => { if (e.key === 'Enter') doSearch(); });

function doSearch() {
  const q = document.getElementById('search').value.trim().toLowerCase();
  if (!q) return;
  const n = nodes.find(x => (x.label + ' ' + x.title + ' ' + x.path).toLowerCase().includes(q));
  if (!n) { document.getElementById('meta').textContent = '未找到匹配文件'; return; }
  if (view === 'changed' && !isVisible(n)) {
    document.getElementById('btn-all').click();
  }
  offsetX = -n.x * scale; offsetY = -n.y * scale;
  select(n);
}
function fit() {
  if (!visible.length) return;
  let x0 = 1e9, x1 = -1e9, y0 = 1e9, y1 = -1e9;
  visible.forEach(n => { x0 = Math.min(x0, n.x); x1 = Math.max(x1, n.x);
                         y0 = Math.min(y0, n.y); y1 = Math.max(y1, n.y); });
  const w = canvas.clientWidth, h = canvas.clientHeight;
  const sx = (w - 120) / Math.max(80, x1 - x0), sy = (h - 120) / Math.max(80, y1 - y0);
  scale = Math.min(2, Math.max(0.15, Math.min(sx, sy)));
  offsetX = -(x0 + x1) / 2 * scale; offsetY = -(y0 + y1) / 2 * scale;
  draw();
}

window.addEventListener('resize', resize);
resize();
applyView();
fit();
</script>
</body>
</html>
"""


def render_graph(root: str, visual: Dict[str, Any], day: str) -> str:
    d = os.path.join(root, INDEX_DIR, "reports")
    os.makedirs(d, exist_ok=True)
    path = os.path.join(d, f"{day}_关系图.html")
    payload = json.dumps(visual, ensure_ascii=False)
    html = _TEMPLATE.replace("__DATA__", payload)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(html)
    os.replace(tmp, path)
    return path
