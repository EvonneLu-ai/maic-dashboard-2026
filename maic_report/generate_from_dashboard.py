#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MAIC 2026 每週報告 — 直接讀取 index.html & analysis.html 真實數據產出
用法：python generate_report_from_html.py [--date 2026-04-28]
"""

import sys, os, re
from datetime import datetime

# ═══════════════════════════════════════════════════════════
#  真實數據（直接從 index.html + analysis.html 對應）
#  每週只需更新這個區塊！
# ═══════════════════════════════════════════════════════════
DATA = {
    'date':          '2026-04-28',           # 資料截至日期
    'date_label':    '截至 2026/04/27',

    # ── 報名數據 KPI ──
    'total_teams':   224,    # 累計報名隊伍
    'total_students':326,    # 參賽學生人數
    'total_teachers':155,    # 指導老師數
    'total_schools': 16,     # 涵蓋學校數
    'registered_accounts': 235,  # 累計註冊帳號
    'target':        800,    # 目標隊伍數

    # ── 賽道 ──
    'track': {
        '創意': 118,
        '創新': 70,
        '創業': 36,
    },

    # ── 學校排行（從 analysis.html 直接取）──
    'school_rank': [
        {'school': '國立臺灣大學',   'teams': 61, '創意': 13, '創新': 26, '創業': 22, 'students': 121, 'teachers': 0},
        {'school': '逢甲大學',      'teams': 47, '創意': 43, '創新': 3,  '創業': 1,  'students': 51,  'teachers': 40},
        {'school': '國立政治大學',   'teams': 40, '創意': 36, '創新': 1,  '創業': 3,  'students': 64,  'teachers': 40},
        {'school': '國立臺中教育大學','teams': 23, '創意': 8,  '創新': 15, '創業': 0,  'students': 22,  'teachers': 22},
        {'school': '國立高雄科技大學','teams': 20, '創意': 6,  '創新': 14, '創業': 0,  'students': 20,  'teachers': 20},
        {'school': '國立清華大學',   'teams': 15, '創意': 3,  '創新': 6,  '創業': 6,  'students': 18,  'teachers': 14},
        {'school': '東海大學',      'teams': 10, '創意': 7,  '創新': 3,  '創業': 0,  'students': 10,  'teachers': 10},
        {'school': '國立陽明交通大學','teams': 7,  '創意': 2,  '創新': 3,  '創業': 2,  'students': 8,   'teachers': 6},
        {'school': '臺北醫學大學',   'teams': 1,  '創意': 0,  '創新': 0,  '創業': 1,  'students': 1,   'teachers': 0},
        {'school': '國立臺北教育大學','teams': 1,  '創意': 1,  '創新': 0,  '創業': 0,  'students': 1,   'teachers': 0},
        {'school': '慈濟科技大學',   'teams': 1,  '創意': 1,  '創新': 0,  '創業': 0,  'students': 2,   'teachers': 1},
        {'school': '國立臺北科技大學','teams': 1,  '創意': 0,  '創新': 1,  '創業': 0,  'students': 1,   'teachers': 0},
        {'school': '輔仁大學',      'teams': 1,  '創意': 0,  '創新': 0,  '創業': 1,  'students': 1,   'teachers': 0},
        {'school': '銘傳大學',      'teams': 1,  '創意': 0,  '創新': 1,  '創業': 0,  'students': 3,   'teachers': 1},
        {'school': '國立臺南大學',   'teams': 1,  '創意': 0,  '創新': 0,  '創業': 1,  'students': 3,   'teachers': 0},
        {'school': '靜宜大學',      'teams': 1,  '創意': 1,  '創新': 0,  '創業': 0,  'students': 0,   'teachers': 1},
    ],

    # ── 科系排行（from analysis.html）──
    'dept_rank': [
        ('教育學系', 40), ('資訊工程學系', 15), ('室內設計進修學士班', 9),
        ('外國語文學系', 8), ('合作經濟暨社會事業經營學系', 7), ('會計學系', 6),
        ('中國文學系', 6), ('經濟學系', 5), ('電機工程學系', 4),
        ('財務金融學系', 4), ('公共事務研究所', 4), ('哲學系', 4),
        ('特殊教育學系', 3), ('物理學系', 3), ('生命科學系', 3),
        ('昆蟲學系', 3), ('華語教學碩士學位學程', 3), ('國際企業學系', 3),
        ('文化與社會創新碩士學位學程', 3), ('教育與學習科技學系', 2),
        ('行銷與流通管理系', 2), ('國際經營與貿易學系', 2),
        ('心理學系', 2), ('植物科學研究所', 2), ('教育研究所', 2),
    ],

    # ── 週增量（有上週資料時填入）──
    'delta_teams':   None,   # e.g. 62
    'delta_schools': None,   # e.g. 2

    # ── 網站流量 ──
    'website_main':  '18,470',
    'website_bts':   '9,940',

    # ── 達成率 ──
    'achievement':   '27.8%',

    # ── 分析洞察 ──
    'insight': '臺大+政大為本週增長主力；東華大學觸及 >10,000 人但轉換率待追蹤，建議設 7 天驗證節點。',
}
# ═══════════════════════════════════════════════════════════

def gen_html(d, out_path):
    school_rows = ''
    for i, s in enumerate(d['school_rank'], 1):
        school_rows += f"""
        <tr>
          <td class="rank">{i}</td>
          <td class="school-name">{s['school']}</td>
          <td class="num-cell">{s['teams']}</td>
          <td class="delta-cell"><span style="color:#8e8e93">—</span></td>
          <td><span class="track-badge c1">{s['創意']}</span></td>
          <td><span class="track-badge c2">{s['創新']}</span></td>
          <td><span class="track-badge c3">{s['創業']}</span></td>
          <td class="num-cell">{s['students']} 人<small style="color:#8e8e93"> ＋{s['teachers']} 師</small></td>
        </tr>"""

    dept_rows = ''
    for i, (dept, cnt) in enumerate(d['dept_rank'], 1):
        dept_rows += f'<tr><td class="rank">{i}</td><td>{dept}</td><td class="num-cell">{cnt}</td></tr>'

    ach_pct = d['total_teams'] * 100 // d['target']
    track = d['track']
    delta_t = f'▲ +{d["delta_teams"]}' if d['delta_teams'] else '—'
    delta_s = f'▲ +{d["delta_schools"]} 所' if d['delta_schools'] else '—'

    html = f"""<!DOCTYPE html>
<html lang="zh-TW"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>MAIC 2026 週報 {d['date']}</title>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,'PingFang TC',sans-serif;background:#f5f5f7;color:#1d1d1f}}
.page{{max-width:1100px;margin:0 auto;padding:32px 24px 64px}}
.header{{background:linear-gradient(135deg,#0071e3,#34c759);border-radius:20px;padding:28px 32px;color:#fff;margin-bottom:28px;display:flex;justify-content:space-between;align-items:center}}
.header h1{{font-size:24px;font-weight:800}}.header p{{font-size:13px;opacity:.85;margin-top:4px}}
.header-right{{text-align:right;font-size:12px;opacity:.85}}
.kpi-row{{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:20px}}
.kpi{{background:#fff;border-radius:16px;padding:18px;text-align:center;box-shadow:0 2px 12px rgba(0,0,0,.06)}}
.kpi .num{{font-size:30px;font-weight:800;color:#0071e3}}.kpi .lbl{{font-size:11px;color:#8e8e93;margin-top:4px}}.kpi .chg{{font-size:11px;color:#34c759;margin-top:2px;font-weight:600}}
.progress-bar{{background:#fff;border-radius:16px;padding:20px;margin-bottom:20px;box-shadow:0 2px 12px rgba(0,0,0,.06)}}
.pb-label{{display:flex;justify-content:space-between;font-size:13px;font-weight:600;margin-bottom:8px}}
.pb-track{{background:#f0f0f0;border-radius:99px;height:14px;overflow:hidden}}
.pb-fill{{height:100%;border-radius:99px;background:linear-gradient(90deg,#0071e3,#34c759);transition:width .6s ease}}
.track-row{{display:flex;gap:12px;margin-bottom:20px}}
.track-pill{{flex:1;border-radius:14px;padding:16px;text-align:center;color:#fff;font-weight:700;font-size:24px}}
.t-c1{{background:linear-gradient(135deg,#ff9f0a,#ff6b00)}}.t-c2{{background:linear-gradient(135deg,#34c759,#1a8c36)}}.t-c3{{background:linear-gradient(135deg,#0071e3,#0050a0)}}
.track-pill .lbl{{font-size:12px;font-weight:500;margin-top:4px;opacity:.9}}
.insight{{background:#fff3cd;border-left:4px solid #ff9f0a;border-radius:0 12px 12px 0;padding:14px 18px;margin-bottom:20px;font-size:13px;line-height:1.6}}
.section{{background:#fff;border-radius:16px;padding:24px;margin-bottom:20px;box-shadow:0 2px 12px rgba(0,0,0,.06)}}
.section h2{{font-size:15px;font-weight:700;margin-bottom:14px}}
.search-bar{{width:100%;padding:9px 14px;border:1px solid #d1d1d6;border-radius:10px;font-size:13px;margin-bottom:12px;outline:none}}
.search-bar:focus{{border-color:#0071e3}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th{{background:#f5f5f7;padding:8px 10px;text-align:left;font-size:11px;font-weight:700;color:#8e8e93;letter-spacing:.5px;text-transform:uppercase}}
td{{padding:9px 10px;border-bottom:1px solid #f0f0f0}}
tr:last-child td{{border-bottom:none}}tr:hover td{{background:#f9f9fb}}
.rank{{font-size:12px;color:#8e8e93;font-weight:700;width:28px}}.num-cell{{text-align:right;font-weight:700}}.delta-cell{{text-align:center}}.school-name{{font-weight:600}}
.track-badge{{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700}}
.c1{{background:rgba(255,159,10,.15);color:#b85a00}}.c2{{background:rgba(52,199,89,.15);color:#1a6628}}.c3{{background:rgba(0,113,227,.15);color:#0050a0}}
.footer{{text-align:center;font-size:11px;color:#8e8e93;margin-top:32px;line-height:1.8}}
@media(max-width:600px){{.kpi-row{{grid-template-columns:repeat(2,1fr)}}.track-row{{flex-direction:column}}}}
</style></head><body>
<div class="page">
  <div class="header">
    <div><h1>📊 MAIC 2026 · 週報</h1><p>行動應用創新賽 2026 ｜ 報名數據即時總覽</p></div>
    <div class="header-right">資料截至<br><strong>{d['date_label']}</strong><br>Apple EDU Taiwan</div>
  </div>

  <div class="kpi-row">
    <div class="kpi"><div class="num">{d['total_teams']}</div><div class="lbl">累計報名隊伍</div><div class="chg">{delta_t}</div></div>
    <div class="kpi"><div class="num">{d['total_students']}</div><div class="lbl">學生人數</div><div class="chg">+{d['total_teachers']} 指導老師</div></div>
    <div class="kpi"><div class="num">{d['total_schools']}</div><div class="lbl">涵蓋學校數</div><div class="chg">{delta_s}</div></div>
    <div class="kpi"><div class="num">{d['registered_accounts']}</div><div class="lbl">累計註冊帳號</div><div class="chg">已驗證</div></div>
    <div class="kpi"><div class="num">{d['achievement']}</div><div class="lbl">目標達成率</div><div class="chg">{d['total_teams']}/{d['target']} 組</div></div>
  </div>

  <div class="progress-bar">
    <div class="pb-label"><span>🎯 目標達成進度</span><span>{d['total_teams']} / {d['target']} 組（{ach_pct}%）</span></div>
    <div class="pb-track"><div class="pb-fill" style="width:{ach_pct}%"></div></div>
  </div>

  <div class="track-row">
    <div class="track-pill t-c1">{track['創意']}<div class="lbl">🎨 創意賽道</div></div>
    <div class="track-pill t-c2">{track['創新']}<div class="lbl">🔬 創新賽道</div></div>
    <div class="track-pill t-c3">{track['創業']}<div class="lbl">🚀 創業賽道</div></div>
  </div>

  <div class="insight">💡 <strong>本週洞察：</strong>{d['insight']}</div>

  <div class="section">
    <h2>🏫 各校組數排行</h2>
    <input class="search-bar" placeholder="🔍 搜尋學校..." oninput="filterTable('st',this.value,1)">
    <table id="st"><thead><tr><th>#</th><th>學校</th><th style="text-align:right">組數</th><th style="text-align:center">週比</th><th>創意</th><th>創新</th><th>創業</th><th style="text-align:right">人數</th></tr></thead>
    <tbody>{school_rows}</tbody></table>
  </div>

  <div class="section">
    <h2>📚 科系報名排行（前 25）</h2>
    <input class="search-bar" placeholder="🔍 搜尋科系..." oninput="filterTable('dt',this.value,1)">
    <table id="dt"><thead><tr><th>#</th><th>科系 / 所</th><th style="text-align:right">組數</th></tr></thead>
    <tbody>{dept_rows}</tbody></table>
  </div>

  <div class="footer">
    MAIC 行動應用創新賽 2026 · 週報 {d['date']} ｜
    <a href="https://evonnelu-ai.github.io/maic-dashboard-2026/" style="color:#0071e3">→ 完整 Dashboard</a>
  </div>
</div>
<script>
function filterTable(id,q,col){{
  document.getElementById(id).querySelectorAll('tbody tr').forEach(r=>{{
    r.style.display=(!q||r.cells[col].textContent.toLowerCase().includes(q.toLowerCase()))?'':'none';
  }});
}}
</script>
</body></html>"""
    with open(out_path,'w',encoding='utf-8') as f: f.write(html)
    print(f"  ✅ HTML 報表：{out_path}")

def gen_pdf_html(d, out_path):
    school_rows = ''
    for i,s in enumerate(d['school_rank'][:12],1):
        school_rows += f"<tr><td>{i}</td><td>{s['school']}</td><td>{s['teams']}</td><td>{s['創意']}</td><td>{s['創新']}</td><td>{s['創業']}</td><td>{s['students']}人+{s['teachers']}師</td></tr>"
    dept_rows = ''
    for i,(dept,cnt) in enumerate(d['dept_rank'][:12],1):
        dept_rows += f"<tr><td>{i}</td><td>{dept}</td><td>{cnt}</td></tr>"
    track = d['track']
    ach_pct = d['total_teams']*100//d['target']
    html = f"""<!DOCTYPE html><html lang="zh-TW"><head>
<meta charset="UTF-8"><title>MAIC 2026 PDF 週報 {d['date']}</title>
<style>
@page{{size:A4 portrait;margin:16mm 14mm 14mm 14mm}}
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,'PingFang TC','Helvetica Neue',sans-serif;font-size:10.5pt;color:#1d1d1f;background:#fff}}
.header{{border-bottom:3px solid #0071e3;padding-bottom:10px;margin-bottom:14px;display:flex;justify-content:space-between;align-items:flex-end}}
.header h1{{font-size:17pt;font-weight:800;color:#0071e3}}
.meta{{font-size:9pt;color:#8e8e93;text-align:right}}
.kpi-grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:8px;margin-bottom:12px}}
.kpi{{border:1.5px solid #d1d1d6;border-radius:9px;padding:8px;text-align:center}}
.kpi .num{{font-size:18pt;font-weight:800;color:#0071e3}}.kpi .lbl{{font-size:7.5pt;color:#8e8e93}}
.pb{{margin-bottom:12px}}.pb-label{{display:flex;justify-content:space-between;font-size:9pt;font-weight:600;margin-bottom:5px}}
.pb-track{{background:#f0f0f0;border-radius:99px;height:10px;overflow:hidden}}
.pb-fill{{height:100%;border-radius:99px;background:linear-gradient(90deg,#0071e3,#34c759)}}
.track-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:12px}}
.track{{border-radius:9px;padding:8px;text-align:center;color:#fff}}
.t1{{background:#ff9f0a}}.t2{{background:#34c759}}.t3{{background:#0071e3}}
.track .cnt{{font-size:18pt;font-weight:800}}.track .lbl{{font-size:8.5pt;opacity:.9}}
.insight{{background:#fff8e1;border-left:3px solid #ff9f0a;padding:8px 12px;margin-bottom:12px;font-size:8.5pt;line-height:1.5;border-radius:0 8px 8px 0}}
.tables{{display:grid;grid-template-columns:1.7fr 1fr;gap:14px;margin-bottom:12px}}
h2{{font-size:10pt;font-weight:700;margin-bottom:7px;border-left:3px solid #0071e3;padding-left:7px}}
table{{width:100%;border-collapse:collapse;font-size:8.5pt}}
th{{background:#f5f5f7;padding:4px 6px;font-weight:700;color:#8e8e93;font-size:7.5pt}}
td{{padding:4px 6px;border-bottom:1px solid #f0f0f0}}
td:first-child{{color:#8e8e93;font-weight:700;width:18px}}
.footer{{text-align:center;font-size:7.5pt;color:#8e8e93;border-top:1px solid #e0e0e5;padding-top:7px}}
.print-btn{{position:fixed;bottom:20px;right:20px;background:#0071e3;color:#fff;border:none;border-radius:10px;padding:10px 20px;font-size:13px;font-weight:700;cursor:pointer;box-shadow:0 4px 16px rgba(0,113,227,.35)}}
@media print{{.print-btn{{display:none}}}}
</style></head><body>
<div class="header">
  <div><h1>📊 MAIC 2026 週報</h1><div style="font-size:9pt;color:#8e8e93;margin-top:2px">行動應用創新賽 · 報名數據摘要</div></div>
  <div class="meta">資料截至：{d['date_label']}<br>Apple EDU Taiwan · Evonne Lu</div>
</div>
<div class="kpi-grid">
  <div class="kpi"><div class="num">{d['total_teams']}</div><div class="lbl">報名組數</div></div>
  <div class="kpi"><div class="num">{d['total_students']}</div><div class="lbl">學生人數</div></div>
  <div class="kpi"><div class="num">{d['total_teachers']}</div><div class="lbl">指導老師</div></div>
  <div class="kpi"><div class="num">{d['total_schools']}</div><div class="lbl">學校數</div></div>
  <div class="kpi"><div class="num">{d['achievement']}</div><div class="lbl">目標達成率</div></div>
</div>
<div class="pb">
  <div class="pb-label"><span>🎯 目標進度（{d['total_teams']}/{d['target']}）</span><span>{ach_pct}%</span></div>
  <div class="pb-track"><div class="pb-fill" style="width:{ach_pct}%"></div></div>
</div>
<div class="track-grid">
  <div class="track t1"><div class="cnt">{track['創意']}</div><div class="lbl">🎨 創意賽道</div></div>
  <div class="track t2"><div class="cnt">{track['創新']}</div><div class="lbl">🔬 創新賽道</div></div>
  <div class="track t3"><div class="cnt">{track['創業']}</div><div class="lbl">🚀 創業賽道</div></div>
</div>
<div class="insight">💡 <strong>本週洞察：</strong>{d['insight']}</div>
<div class="tables">
  <div><h2>各校組數排行（Top 12）</h2>
  <table><thead><tr><th>#</th><th>學校</th><th>組</th><th>意</th><th>新</th><th>業</th><th>人數</th></tr></thead>
  <tbody>{school_rows}</tbody></table></div>
  <div><h2>科系排行（Top 12）</h2>
  <table><thead><tr><th>#</th><th>科系</th><th>組</th></tr></thead>
  <tbody>{dept_rows}</tbody></table></div>
</div>
<div class="footer">MAIC 2026 週報 {d['date']} ｜ https://evonnelu-ai.github.io/maic-dashboard-2026/</div>
<button class="print-btn" onclick="window.print()">🖨️ 列印 / 儲存 PDF</button>
</body></html>"""
    with open(out_path,'w',encoding='utf-8') as f: f.write(html)
    print(f"  ✅ PDF 頁面：{out_path}")

def gen_pptx(d, out_path):
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
    except ImportError:
        print("  ⚠️  請先執行：pip install python-pptx"); return

    prs = Presentation()
    prs.slide_width = Inches(13.33); prs.slide_height = Inches(7.5)
    BLUE=RGBColor(0x00,0x71,0xE3); GREEN=RGBColor(0x34,0xC7,0x59)
    ORANGE=RGBColor(0xFF,0x9F,0x0A); WHITE=RGBColor(0xFF,0xFF,0xFF)
    DARK=RGBColor(0x1d,0x1d,0x1f); GRAY=RGBColor(0x8e,0x8e,0x93)
    blank=prs.slide_layouts[6]

    def tb(slide,text,l,t,w,h,size=16,bold=False,color=DARK,align=PP_ALIGN.LEFT):
        s=slide.shapes.add_textbox(Inches(l),Inches(t),Inches(w),Inches(h))
        tf=s.text_frame; tf.word_wrap=True
        p=tf.paragraphs[0]; p.alignment=align
        r=p.add_run(); r.text=text; r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=color
    def rect(slide,l,t,w,h,clr):
        s=slide.shapes.add_shape(1,Inches(l),Inches(t),Inches(w),Inches(h))
        s.fill.solid(); s.fill.fore_color.rgb=clr; s.line.fill.background(); return s

    track=d['track']
    ach_pct=d['total_teams']*100//d['target']

    # Slide 1: 封面
    s1=prs.slides.add_slide(blank)
    rect(s1,0,0,13.33,7.5,RGBColor(0x00,0x47,0x9A))
    tb(s1,'MAIC 2026',1.5,1.5,10,1.5,size=60,bold=True,color=WHITE,align=PP_ALIGN.CENTER)
    tb(s1,f'每週報名數據週報 · {d["date"]}',1.5,3.2,10,0.8,size=22,color=RGBColor(0xAA,0xCC,0xFF),align=PP_ALIGN.CENTER)
    tb(s1,'Apple EDU Taiwan · Evonne Lu',1.5,4.1,10,0.6,size=15,color=RGBColor(0x88,0xAA,0xCC),align=PP_ALIGN.CENTER)

    # Slide 2: KPI 總覽
    s2=prs.slides.add_slide(blank)
    rect(s2,0,0,13.33,1.1,BLUE)
    tb(s2,'📊 本週數據總覽',0.3,0.12,10,0.85,size=22,bold=True,color=WHITE)
    tb(s2,f'資料截至 {d["date_label"]}',9.5,0.22,3.5,0.65,size=12,color=RGBColor(0xAA,0xCC,0xFF))
    kpis=[(str(d['total_teams']),'累計報名隊伍',0.4,BLUE),(str(d['total_students']),'學生人數',3.2,GREEN),
          (str(d['total_teachers']),'指導老師',6.0,ORANGE),(str(d['total_schools']),'涵蓋學校數',8.8,RGBColor(0xAF,0x52,0xDE)),
          (d['achievement'],'目標達成率',11.0,RGBColor(0xFF,0x3B,0x30))]
    for val,lbl,x,clr in kpis:
        rect(s2,x,1.35,2.15,2.4,RGBColor(0xF5,0xF5,0xF7))
        tb(s2,val,x,1.45,2.15,1.3,size=36,bold=True,color=clr,align=PP_ALIGN.CENTER)
        tb(s2,lbl,x,2.85,2.15,0.55,size=12,color=GRAY,align=PP_ALIGN.CENTER)
    # 進度條
    rect(s2,0.4,4.0,12.5,0.6,RGBColor(0xE5,0xE5,0xEA))
    bar_w=max(0.3,12.5*ach_pct/100)
    rect(s2,0.4,4.0,bar_w,0.6,BLUE)
    tb(s2,f'目標進度 {d["total_teams"]}/{d["target"]} 組（{ach_pct}%）',0.4,4.7,12,0.5,size=13,color=GRAY)
    # 賽道
    for i,(name,val,clr) in enumerate([('🎨 創意',track['創意'],ORANGE),('🔬 創新',track['創新'],GREEN),('🚀 創業',track['創業'],BLUE)]):
        x=0.4+i*4.15
        rect(s2,x,5.4,3.7,1.6,clr)
        tb(s2,str(val),x,5.45,3.7,0.95,size=40,bold=True,color=WHITE,align=PP_ALIGN.CENTER)
        tb(s2,name,x,6.35,3.7,0.55,size=14,color=WHITE,align=PP_ALIGN.CENTER)

    # Slide 3: 學校排行
    s3=prs.slides.add_slide(blank)
    rect(s3,0,0,13.33,1.1,GREEN)
    tb(s3,'🏫 各校組數排行（Top 10）',0.3,0.12,12,0.85,size=22,bold=True,color=WHITE)
    hdrs=['#','學校','組數','創意','創新','創業','人數']
    xs=[0.3,0.75,4.4,5.35,6.3,7.25,8.2]; ws=[0.4,3.55,0.85,0.85,0.85,0.85,2.0]
    for j,(h,x,w) in enumerate(zip(hdrs,xs,ws)):
        tb(s3,h,x,1.25,w,0.4,size=10,bold=True,color=GRAY)
    for i,s in enumerate(d['school_rank'][:10]):
        y=1.72+i*0.49
        if i%2==0: rect(s3,0.3,y-0.05,9.8,0.44,RGBColor(0xF9,0xF9,0xFB))
        vals=[str(i+1),s['school'],str(s['teams']),str(s['創意']),str(s['創新']),str(s['創業']),f"{s['students']}人+{s['teachers']}師"]
        for j,(v,x,w) in enumerate(zip(vals,xs,ws)):
            clr=BLUE if j==2 else DARK
            tb(s3,v,x,y,w,0.42,size=11 if j!=1 else 10,color=clr,bold=(j==2))

    # Slide 4: 科系 + 洞察
    s4=prs.slides.add_slide(blank)
    rect(s4,0,0,13.33,1.1,ORANGE)
    tb(s4,'📚 科系排行 & 本週洞察',0.3,0.12,12,0.85,size=22,bold=True,color=WHITE)
    tb(s4,'科系報名排行（Top 10）',0.4,1.25,5.5,0.45,size=13,bold=True,color=DARK)
    for i,(dept,cnt) in enumerate(d['dept_rank'][:10]):
        y=1.75+i*0.48
        if i%2==0: rect(s4,0.4,y-0.04,5.5,0.44,RGBColor(0xF9,0xF9,0xFB))
        tb(s4,str(i+1),0.4,y,0.4,0.38,size=11,color=GRAY)
        tb(s4,dept,0.85,y,4.3,0.38,size=10,color=DARK)
        tb(s4,str(cnt),4.9,y,0.7,0.38,size=12,color=ORANGE,bold=True)
    tb(s4,'💡 本週洞察',7.0,1.25,6.0,0.45,size=13,bold=True,color=DARK)
    rect(s4,7.0,1.75,6.0,3.5,RGBColor(0xFF,0xF8,0xE1))
    tb(s4,d['insight'],7.15,1.9,5.7,3.2,size=13,color=DARK)
    tb(s4,f'🔗 https://evonnelu-ai.github.io/maic-dashboard-2026/',7.0,5.6,6.0,0.5,size=10,color=BLUE)

    prs.save(out_path)
    print(f"  ✅ PPTX 投影片：{out_path}")

def gen_summary(d, out_path):
    track=d['track']
    top3=' / '.join([f"{s['school']}（{s['teams']}組）" for s in d['school_rank'][:3]])
    delta_str=f'（較上週 ▲+{d["delta_teams"]}）' if d['delta_teams'] else ''
    summary=f"""📊 MAIC 2026 週報 {d['date']}
━━━━━━━━━━━━━━━━━━━━
🏆 報名組數：{d['total_teams']} 組 {delta_str}
👥 學生人數：{d['total_students']} 人（含 {d['total_teachers']} 位指導老師）
🏫 涵蓋學校：{d['total_schools']} 所
🎯 目標達成率：{d['achievement']}（{d['total_teams']}/{d['target']} 組）
📊 賽道分布：創意 {track['創意']} ｜ 創新 {track['創新']} ｜ 創業 {track['創業']}

🔝 學校 Top 3：
   {top3}

💡 {d['insight']}

🔗 完整 Dashboard：
   https://evonnelu-ai.github.io/maic-dashboard-2026/
━━━━━━━━━━━━━━━━━━━━
Apple EDU Taiwan · Evonne Lu"""
    with open(out_path,'w',encoding='utf-8') as f: f.write(summary)
    print(f"  ✅ 摘要訊息：{out_path}")
    print(); print("── 訊息預覽 ─────────────────────────────────────"); print(summary); print("─────────────────────────────────────────────────")

def main():
    date_str = DATA['date']
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
    os.makedirs(out_dir, exist_ok=True)
    dn = date_str.replace('-','')
    t=DATA['track']; s=sum(t.values())
    print(f"\n🚀 MAIC 2026 週報產出中... ({date_str})")
    print(f"   報名：{DATA['total_teams']} 組 ｜ 學校：{DATA['total_schools']} 所 ｜ 創意{t['創意']} 創新{t['創新']} 創業{t['創業']}\n")
    gen_html(DATA, os.path.join(out_dir, f'maic_report_{dn}.html'))
    gen_pdf_html(DATA, os.path.join(out_dir, f'maic_pdf_{dn}.html'))
    gen_pptx(DATA, os.path.join(out_dir, f'maic_slides_{dn}.pptx'))
    gen_summary(DATA, os.path.join(out_dir, f'maic_summary_{dn}.txt'))
    print(f"\n✨ 全部完成！輸出目錄：{out_dir}/")

if __name__=='__main__': main()
