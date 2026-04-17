import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib import rcParams
import matplotlib.font_manager as fm
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import csv, io, urllib.request, os

import os; FONT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "../NotoSansCJKtc.otf")
fe = fm.FontEntry(fname=FONT_PATH, name='NotoTC')
fm.fontManager.ttflist.append(fe)
rcParams['font.family'] = 'NotoTC'
rcParams['axes.unicode_minus'] = False

# ── 色彩 ──
CARD=(255,255,255); DARK=(29,29,31); GRAY=(110,110,115)
LIGHT=(245,245,247); BORDER=(210,210,215)
BLUE=(58,122,200); GREEN=(60,179,100); ORANGE=(222,140,60); RED=(210,72,68)
PURPLE=(148,100,200); TEAL=(72,185,210)
def rgb(c): return tuple(v/255 for v in c)

TYPE_COL={'op':BLUE,'mk':ORANGE,'tr':GREEN,'gc':PURPLE}
TYPE_LBL={'op':'OP','mk':'MK','tr':'TR','gc':'GC'}
PRI_COL={'high':RED,'mid':ORANGE,'low':TEAL}
STAT_COL={'open':BLUE,'done':GREEN,'delay':RED}
STAT_LBL={'open':'進行中','done':'已完成','delay':'延遲'}

def csv_get(url):
    with urllib.request.urlopen(url) as r:
        return list(csv.DictReader(io.StringIO(r.read().decode('utf-8'))))

BASE="https://raw.githubusercontent.com/EvonneLu-ai/maic-dashboard-2026/main/data"
kpi   = {r['field']:r['value'] for r in csv_get(f"{BASE}/operation_kpi.csv")}
tl    = csv_get(f"{BASE}/timeline.csv")
sess  = csv_get(f"{BASE}/sessions.csv")
rtc   = csv_get(f"{BASE}/rtc.csv")
# ── 決議資料（以截圖為準，僅 open 項目）──
dec = [
    # W17 | 2026/04/17 | 10 筆
    {'meeting':'週會 W17｜2026/04/17','id':'w01','owner':'正奎',          'item':'發送資料給中教大郭校長與中心，並通知 Edick 追蹤',                  'deadline':'4/22','priority':'high','status':'open'},
    {'meeting':'週會 W17｜2026/04/17','id':'w02','owner':'正奎',          'item':'上傳 GC K-12 賽事規則資料至共享資料夾',                            'deadline':'4/17','priority':'high','status':'open'},
    {'meeting':'週會 W17｜2026/04/17','id':'w03','owner':'正奎 / SA',     'item':'向老闆報備 K-12 賽道狀況',                                        'deadline':'4/21','priority':'high','status':'open'},
    {'meeting':'週會 W17｜2026/04/17','id':'w04','owner':'正奎',          'item':'確認夏季訓練營 SA 具體分工，並回報是否參與台大協調會',              'deadline':'4/24','priority':'mid', 'status':'open'},
    {'meeting':'週會 W17｜2026/04/17','id':'w05','owner':'Bella',         'item':'製作 4/30 台大宣講會 IG / FB 圖卡（需含學校官稱）',               'deadline':'儘速','priority':'high','status':'open'},
    {'meeting':'週會 W17｜2026/04/17','id':'w06','owner':'Edick',         'item':'與清大邱教授初步聯繫，評估小型宣講會可能性',                       'deadline':'下週','priority':'high','status':'open'},
    {'meeting':'週會 W17｜2026/04/17','id':'w07','owner':'Kat / SA',      'item':'將 PDF 行銷計劃書內容填入 Keynote 簡報模組（含歷年數據）',         'deadline':'4/24','priority':'high','status':'open'},
    {'meeting':'週會 W17｜2026/04/17','id':'w08','owner':'Willy / Brendi','item':'提供主視覺 AI 原始檔給正奎及傳揚（David）',                       'deadline':'儘速','priority':'high','status':'open'},
    {'meeting':'週會 W17｜2026/04/17','id':'w09','owner':'Brendi',        'item':'更新 Numbers 欄位（標註線上/實體課程），並追蹤創業組名單',          'deadline':'每週一','priority':'mid','status':'open'},
    {'meeting':'週會 W17｜2026/04/17','id':'w10','owner':'Evonne',        'item':'串接 Dashboard，產出視覺化週報圖表',                              'deadline':'每週會後','priority':'mid','status':'open'},
    # W16 | 2026/04/10 | 2 筆
    {'meeting':'週會 W16｜2026/04/10','id':'d13','owner':'Edick',         'item':'回覆成功大學場地費洽談進展',                                       'deadline':'4/17','priority':'mid','status':'open'},
    {'meeting':'週會 W16｜2026/04/10','id':'d20','owner':'正奎',          'item':'確認清華大學宣講會執行可能性（對口：邱富源教授）',                  'deadline':'4/17','priority':'mid','status':'open'},
    # W15 | 2026/04/09 | 2 筆
    {'meeting':'週會 W15｜2026/04/09','id':'d01','owner':'Brandi',        'item':'清華大學宣講時間確認，回覆給 Max',                                 'deadline':'4/15','priority':'high','status':'open'},
    {'meeting':'週會 W15｜2026/04/09','id':'d03','owner':'13',            'item':'成功大學宣講等 Edick 確認後排定時間',                              'deadline':'4/20','priority':'mid', 'status':'open'},
]
schl  = csv_get(f"{REPO}/data/schools.csv")
dept  = csv_get(f"{REPO}/data/departments.csv")
email = csv_get(f"{REPO}/data/emails.csv")
ws    = csv_get(f"{REPO}/data/workshops.csv")

# ═══════════════════════════════════════
# 生成圖表 PNG
# ═══════════════════════════════════════

def pie_chart(labels, sizes, colors, title, path, legend=True):
    fig, ax = plt.subplots(figsize=(5.5, 3.8))
    fig.patch.set_facecolor('white')
    wedges, texts, autotexts = ax.pie(
        sizes, labels=None, colors=[rgb(c) for c in colors],
        autopct='%1.0f%%', startangle=140,
        wedgeprops=dict(width=0.55, edgecolor='white', linewidth=1.5),
        pctdistance=0.78
    )
    for at in autotexts:
        at.set_fontsize(7); at.set_color('white'); at.set_fontweight('bold')
    ax.set_title(title, fontsize=10, fontweight='bold', pad=10, color='#1d1d1f')
    if legend:
        patches = [mpatches.Patch(color=rgb(c), label=f'{l} ({s}%)') for l,s,c in zip(labels,sizes,colors)]
        ax.legend(handles=patches, loc='center left', bbox_to_anchor=(1,0.5),
                  fontsize=6.5, frameon=False)
    plt.tight_layout()
    plt.savefig(path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()

def bar_h_chart(labels, values, colors, title, path):
    fig, ax = plt.subplots(figsize=(5.5, 3.2))
    fig.patch.set_facecolor('white')
    bars = ax.barh(labels[::-1], values[::-1], color=[rgb(c) for c in colors[::-1]],
                   edgecolor='white', height=0.55)
    for bar, val in zip(bars, values[::-1]):
        ax.text(bar.get_width()+0.3, bar.get_y()+bar.get_height()/2,
                f'{val}%', va='center', fontsize=7, color='#6e6e73')
    ax.set_xlim(0, max(values)*1.2)
    ax.set_title(title, fontsize=10, fontweight='bold', pad=8, color='#1d1d1f')
    ax.spines[['top','right','bottom']].set_visible(False)
    ax.tick_params(axis='x', which='both', bottom=False, labelbottom=False)
    ax.tick_params(axis='y', labelsize=8)
    plt.tight_layout()
    plt.savefig(path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()

def donut_act(labels, values, colors, title, path):
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(5.8, 3.0),
                                    gridspec_kw={'width_ratios':[1,1.2]})
    fig.patch.set_facecolor('white')
    ax1.pie(values, colors=[rgb(c) for c in colors],
            wedgeprops=dict(width=0.55, edgecolor='white', linewidth=1.5),
            startangle=90)
    ax1.set_title(title, fontsize=9, fontweight='bold', color='#1d1d1f')
    ax2.axis('off')
    for i,(l,v,c) in enumerate(zip(labels,values,colors)):
        y = 1 - i*0.18
        ax2.add_patch(plt.Rectangle((0,y-0.05),0.12,0.1,color=rgb(c),transform=ax2.transAxes))
        ax2.text(0.16, y, f'{l}  {v}場', transform=ax2.transAxes,
                 fontsize=7.5, va='center', color='#3a3a3c')
    plt.tight_layout(pad=0.5)
    plt.savefig(path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()

def email_bar(data, path):
    fig, ax = plt.subplots(figsize=(5.8, 2.8))
    fig.patch.set_facecolor('white')
    names = [d['name'].replace('2025年參賽隊伍','').replace('「','').replace('」','').strip() for d in data]
    open_rates = [float(d['open_rate']) for d in data]
    ctrs       = [float(d['ctr']) for d in data]
    x = range(len(names))
    w = 0.35
    ax.bar([i-w/2 for i in x], open_rates, w, label='開信率 %', color=rgb(BLUE), alpha=0.85)
    ax.bar([i+w/2 for i in x], ctrs,       w, label='CTR %',   color=rgb(ORANGE), alpha=0.85)
    for i,v in enumerate(open_rates):
        ax.text(i-w/2, v+0.5, f'{v}%', ha='center', fontsize=7, color='#3a3a3c')
    for i,v in enumerate(ctrs):
        ax.text(i+w/2, v+0.5, f'{v}%', ha='center', fontsize=7, color='#3a3a3c')
    ax.set_xticks(list(x)); ax.set_xticklabels(names, fontsize=7)
    ax.set_title('EDM 信件成效', fontsize=10, fontweight='bold', color='#1d1d1f')
    ax.legend(fontsize=8, frameon=False)
    ax.spines[['top','right']].set_visible(False)
    ax.set_ylim(0, max(open_rates)*1.3)
    plt.tight_layout()
    plt.savefig(path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()

# 科系分佈
dept_labels = [r['label'] for r in dept]
dept_vals   = [int(r['percentage']) for r in dept]
dept_colors = [BLUE, TEAL, PURPLE, ORANGE, GREEN, RED, (142,142,147)]
pie_chart(dept_labels, dept_vals, dept_colors, '科系分佈', f"/tmp/chart_dept.png")

# 學校分佈 (horizontal bar)
schl_labels = [r['label'] for r in schl]
schl_vals   = [int(r['percentage']) for r in schl]
schl_colors = [BLUE, ORANGE, GREEN, TEAL, PURPLE, RED,
               (255,200,0),(200,100,200),(100,200,150),(200,150,100),(142,142,147)][:len(schl_labels)]
bar_h_chart(schl_labels, schl_vals, schl_colors, '學校分佈', f"/tmp/chart_school.png")

# 活動類型分佈
act_type_labels = ['多元能力講座','宣講會/宣講+工作坊','RTC課程','TW初賽/決賽','GC夏令營/決賽']
act_type_vals   = [12, 10, 10, 7, 2]
act_type_colors = [TEAL, ORANGE, GREEN, BLUE, PURPLE]
donut_act(act_type_labels, act_type_vals, act_type_colors, '活動類型分佈', f"/tmp/chart_act.png")

# EDM 開信率
email_bar(email, f"/tmp/chart_email.png")


print("圖表生成完成")

# ═══════════════════════════════════════
# 建構 PDF
# ═══════════════════════════════════════
class PDF(FPDF):
    def header(self): pass
    def footer(self):
        self.set_y(-11); self.set_font('f','',7.5)
        self.set_text_color(*GRAY)
        self.cell(0,6,f'MAIC 2026 Dashboard  ·  資料截至 2026/04/17  ·  {self.page_no()} / {self.pages_count}',align='C')

    def sec(self, txt, col=BLUE):
        self.ln(5)
        x,y = self.l_margin, self.get_y()
        self.set_fill_color(*col); self.rect(x, y+1, 3, 8, 'F')
        bg = (236,244,255) if col==BLUE else \
             (255,244,228) if col==ORANGE else \
             (234,250,240) if col==GREEN else \
             (255,238,238) if col==RED else (246,236,255)
        self.set_fill_color(*bg); self.rect(x+4, y, self.epw-4, 10, 'F')
        self.set_font('f','',10); self.set_text_color(*DARK)
        self.set_xy(x+8, y+0.5)
        self.cell(0, 10, txt, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(2)

    def kcard(self, lbl, val, sub='', col=BLUE, bw=44):
        x,y = self.get_x(), self.get_y()
        self.set_fill_color(*CARD); self.set_draw_color(*BORDER)
        self.rect(x, y, bw, 24, 'FD')
        self.set_fill_color(*col); self.rect(x, y, bw, 2.5, 'F')
        self.set_font('f','',7); self.set_text_color(*GRAY)
        self.set_xy(x+3,y+4); self.cell(bw-6,4,lbl)
        self.set_font('f','',17); self.set_text_color(*DARK)
        self.set_xy(x+3,y+8.5); self.cell(bw-6,8,val)
        if sub:
            self.set_font('f','',6.5); self.set_text_color(*GRAY)
            self.set_xy(x+3,y+18); self.cell(bw-6,4,sub)
        self.set_xy(x+bw+2, y)

    def thead(self, cols, widths):
        self.set_fill_color(*DARK); self.set_text_color(*CARD); self.set_font('f','',8.5)
        for c,w in zip(cols,widths): self.cell(w,7,f'  {c}',fill=True)
        self.ln()

pdf = PDF()
pdf.add_font('f', fname=FONT_PATH)
pdf.set_auto_page_break(True, 14)
pdf.set_margins(13,13,13)

# ─────────────── PAGE 1: KPI + 圖表 ───────────────
pdf.add_page()
pdf.set_fill_color(*DARK); pdf.rect(0,0,210,30,'F')
pdf.set_font('f','',20); pdf.set_text_color(*CARD)
pdf.set_xy(13,6); pdf.cell(0,10,'MAIC Dashboard 2026')
pdf.set_font('f','',8.5); pdf.set_text_color(160,160,165)
pdf.set_xy(13,19); pdf.cell(0,7,'行動應用創新賽  2026 Season  |  Straight A × Apple  |  資料截至 2026/04/17')
pdf.ln(20)

pdf.sec('報名數據總覽', BLUE)
teams=kpi.get('teams','—'); acct=kpi.get('accounts','—')
trg=int(kpi.get('registrationTarget',800))
pct=round(int(teams)/trg*100,1) if teams.isdigit() else 0
pdf.kcard('累計報名隊伍', teams, f'目標 {trg} 隊  ({pct}%)', BLUE)
pdf.kcard('累計註冊帳號', acct, '含未完成報名', TEAL)
pdf.kcard('涵蓋學校數', kpi.get('schools','—'), '所學校', ORANGE)
pdf.kcard('網站瀏覽數', kpi.get('website1_views','—'), kpi.get('website1_label',''), GREEN)
pdf.ln(27)
pdf.kcard('創意組', kpi.get('creative','—'), '隊', PURPLE)
pdf.kcard('創新組', kpi.get('innovative','—'), '隊', ORANGE)
pdf.kcard('創業組', kpi.get('startup','—'), '隊', RED)
pdf.kcard('企劃書已繳', kpi.get('submitted','—'), f"未繳 {kpi.get('not_submitted','—')} 隊", GREEN)
pdf.ln(28)

# 兩個圖表並排
pdf.sec('科系 & 學校分佈', BLUE)
iw = (pdf.epw - 4) / 2
pdf.image(f"/tmp/chart_dept.png", x=pdf.l_margin, y=pdf.get_y(), w=iw)
pdf.image(f"/tmp/chart_school.png", x=pdf.l_margin+iw+4, y=pdf.get_y(), w=iw)
pdf.ln(58)

# 活動類型
pdf.sec('活動類型分佈', BLUE)
pdf.image(f"/tmp/chart_act.png", x=pdf.l_margin, y=pdf.get_y(), w=pdf.epw*0.6)
pdf.ln(40)

# ─────────────── PAGE 2: 宣講 + EDM ───────────────
pdf.add_page()
pdf.sec('已完成宣講場次', ORANGE)
pdf.thead(['學校','日期','講師','形式','人數'],[52,26,36,16,22])
for i,s in enumerate(sess):
    pdf.set_fill_color(*(LIGHT if i%2==0 else CARD))
    pdf.set_text_color(*DARK); pdf.set_font('f','',8.5)
    pdf.cell(52,7,f"  {s['school']}",fill=True)
    pdf.cell(26,7,s['date'],fill=True)
    pdf.cell(36,7,s['lecturer'],fill=True)
    pdf.set_text_color(*(BLUE if s['type']=='線下' else TEAL))
    pdf.cell(16,7,s['type'],fill=True)
    pdf.set_text_color(*DARK)
    pdf.cell(22,7,f"{s['people']} 人",fill=True,align='R')
    pdf.ln()
pdf.ln(4)

pdf.sec('EDM 信件成效', ORANGE)
pdf.image(f"/tmp/chart_email.png", x=pdf.l_margin, y=pdf.get_y(), w=pdf.epw*0.75)
pdf.ln(50)

# EDM 細項表格
pdf.thead(['信件名稱','寄出','開信','開信率','點擊','CTR'],[88,14,14,18,14,14])
for i,e in enumerate(email):
    bg = LIGHT if i%2==0 else CARD
    pdf.set_fill_color(*bg); pdf.set_text_color(*DARK); pdf.set_font('f','',8.5)
    nm = e['name'][:38]+'…' if len(e['name'])>39 else e['name']
    pdf.cell(88,7,f'  {nm}',fill=True)
    for v,w,al in [(e['sent'],14,'R'),(e['opened'],14,'R'),(f"{e['open_rate']}%",18,'R'),
                    (e['clicks'],14,'R'),(f"{e['ctr']}%",14,'R')]:
        pdf.cell(w,7,str(v),fill=True,align=al)
    pdf.ln()

# ─────────────── PAGE 3: 時間軸 ───────────────
pdf.add_page()
pdf.sec('全年活動時間軸', BLUE)
lx = pdf.l_margin
for t,lbl in [('op','Operation 賽事'),('mk','Marketing 宣傳'),('tr','Training 培訓'),('gc','GC 國際')]:
    c = TYPE_COL[t]
    pdf.set_fill_color(*c); pdf.rect(lx, pdf.get_y()+1.5, 8, 4,'F')
    pdf.set_font('f','',7.5); pdf.set_text_color(*GRAY)
    pdf.set_xy(lx+10, pdf.get_y()); pdf.cell(32,7,lbl)
    lx += 46
pdf.ln(9)
last_m=''
for ev in tl:
    m = ev['month']
    if m != last_m:
        last_m = m
        pdf.set_font('f','',8); pdf.set_text_color(*GRAY)
        pdf.set_fill_color(*LIGHT)
        pdf.cell(pdf.epw,5,f'  {m}',fill=True,new_x=XPos.LMARGIN,new_y=YPos.NEXT)
        pdf.ln(0.5)
    t = ev['type']; col = TYPE_COL.get(t,GRAY)
    pdf.set_fill_color(*col); pdf.rect(pdf.l_margin, pdf.get_y()+0.8, 2.5, 5.5,'F')
    pdf.set_x(pdf.l_margin+5)
    pdf.set_fill_color(*col); pdf.set_text_color(*CARD); pdf.set_font('f','',6.5)
    pdf.cell(10,6,TYPE_LBL.get(t,''),fill=True)
    pdf.set_x(pdf.get_x()+2); pdf.set_text_color(*GRAY); pdf.set_font('f','',7.5)
    pdf.cell(27,6,ev['date'])
    pdf.set_text_color(*DARK); pdf.set_font('f','',8)
    nm = ev['name'][:46]+'…' if len(ev['name'])>47 else ev['name']
    pdf.cell(112,6,nm)
    pdf.set_font('f','',6.5); pdf.set_text_color(*GRAY)
    mt = ev['meta'][:36]+'…' if len(ev['meta'])>37 else ev['meta']
    pdf.cell(0,6,mt,new_x=XPos.LMARGIN,new_y=YPos.NEXT)

# ─────────────── PAGE 4: Training ───────────────
pdf.add_page()
pdf.sec('RTC 課程進度', GREEN)
pdf.thead(['日期','課程名稱','講師','狀態'],[24,102,36,20])
for i,r in enumerate(rtc):
    done = r['done'].strip().lower()=='true'
    pdf.set_fill_color(*((238,252,242) if done else LIGHT if i%2==0 else CARD))
    pdf.set_text_color(*GRAY if done else DARK); pdf.set_font('f','',8.5)
    pdf.cell(24,7,f"  {r['date']}",fill=True)
    nm = r['name'][:43]+'…' if len(r['name'])>44 else r['name']
    pdf.cell(102,7,nm,fill=True)
    pdf.cell(36,7,r['lecturer'][:17],fill=True)
    pdf.set_text_color(*(GREEN if done else ORANGE))
    pdf.cell(20,7,'✓ 已完成' if done else '待執行',fill=True)
    pdf.ln()
pdf.ln(5)

pdf.sec('多元能力強化講座', GREEN)
pdf.thead(['月份','日期','講座名稱','形式','狀態'],[14,22,108,18,20])
for i,w in enumerate(ws):
    done = w['done'].strip().lower()=='true'
    pdf.set_fill_color(*((238,252,242) if done else LIGHT if i%2==0 else CARD))
    pdf.set_text_color(*GRAY if done else DARK); pdf.set_font('f','',8.5)
    pdf.cell(14,7,f"  {w['month']}",fill=True)
    pdf.cell(22,7,w['date'],fill=True)
    nm = w['name'][:46]+'…' if len(w['name'])>47 else w['name']
    pdf.cell(108,7,nm,fill=True)
    pdf.set_text_color(*(BLUE if w['format']=='線上' else ORANGE))
    pdf.cell(18,7,w['format'],fill=True)
    pdf.set_text_color(*(GREEN if done else ORANGE))
    pdf.cell(20,7,'✓ 完成' if done else '待執行',fill=True)
    pdf.ln()

# ─────────────── PAGE 5: 決議（只顯示 open）───────────────
pdf.add_page()
pdf.sec('待辦決議追蹤', RED)

open_dec = dec  # 已預先過濾為 open 項目
delay_dec = []  # 無 delay 項目
show_dec = dec

total_all=len(dec)
n_open=len(open_dec)
n_delay=len(delay_dec)
n_done=sum(1 for d in dec if d['status']=='done')
pdf.set_font('f','',8.5); pdf.set_text_color(*GRAY)
pdf.cell(0,7,f'  待處理 {n_open+n_delay} 項  |  進行中 {n_open}  |  延遲 {n_delay}  |  已完成 {n_done}（不顯示）',
         new_x=XPos.LMARGIN, new_y=YPos.NEXT); pdf.ln(2)

def meeting_key(d):
    for k,v in {'W17':0,'W16':1,'W15':2,'W14':3,'W13':4}.items():
        if k in d['meeting']: return v
    return 99

show_dec.sort(key=meeting_key)
last_m=''
for d in show_dec:
    mtg = d['meeting']
    if mtg != last_m:
        last_m = mtg
        pdf.set_font('f','',8); pdf.set_text_color(*GRAY)
        pdf.set_fill_color(*LIGHT)
        pdf.cell(pdf.epw,5.5,f'  {mtg}',fill=True,new_x=XPos.LMARGIN,new_y=YPos.NEXT)
        pdf.ln(1)

    status = d['status']
    pri    = d['priority']
    pc     = PRI_COL.get(pri, GRAY)
    sc     = RED if status=='delay' else BLUE
    row_h  = 8

    # 左色條
    pdf.set_fill_color(*pc)
    pdf.rect(pdf.l_margin, pdf.get_y()+0.5, 3, row_h-1, 'F')
    pdf.set_x(pdf.l_margin+5)

    # 空 checkbox
    pdf.set_fill_color(*CARD); pdf.set_draw_color(*BORDER)
    pdf.rect(pdf.get_x(), pdf.get_y()+1.5, 5, 5, 'FD')
    pdf.set_x(pdf.get_x()+7)

    # owner badge
    pdf.set_fill_color(*LIGHT); pdf.set_draw_color(*BORDER)
    pdf.set_font('f','',7.5); pdf.set_text_color(*DARK)
    ow = pdf.get_string_width(d['owner'])+6
    pdf.cell(ow, row_h, d['owner'], fill=True)
    pdf.set_x(pdf.get_x()+3)

    # item text
    pdf.set_font('f','',8); pdf.set_text_color(*DARK)
    item_w = pdf.epw - ow - 52
    it = d['item'][:74]+'…' if len(d['item'])>75 else d['item']
    pdf.cell(item_w, row_h, it)

    # deadline
    pdf.set_font('f','',7.5); pdf.set_text_color(*GRAY)
    pdf.cell(16, row_h, d['deadline'], align='C')

    # status badge
    pdf.set_fill_color(*sc); pdf.set_text_color(*CARD); pdf.set_font('f','',7)
    sl = '⚠ 延遲' if status=='delay' else '進行中'
    pdf.cell(pdf.get_string_width(sl)+6, row_h, sl, fill=True)

    # priority badge
    pdf.set_x(pdf.get_x()+2)
    pl = {'high':'高','mid':'中','low':'低'}.get(pri,'')
    pdf.set_fill_color(*pc)
    pdf.cell(pdf.get_string_width(pl)+5, row_h, pl, fill=True)

    pdf.ln(row_h+1)

OUT = os.path.join(REPO, "MAIC2026_Dashboard.pdf")
pdf.output(OUT)
sz = os.path.getsize(OUT)
print(f"完成 {OUT}  {sz//1024}KB  {pdf.page}頁")
