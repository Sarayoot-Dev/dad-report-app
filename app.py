import streamlit as st
import struct
import io
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import SeriesLabel

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title='DAD Report — Valeo NB2',
    page_icon='🏭',
    layout='wide',
    initial_sidebar_state='expanded',
)

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F4E79, #2E75B6);
        color: white; padding: 20px 28px; border-radius: 12px;
        margin-bottom: 24px;
    }
    .main-header h1 { margin: 0; font-size: 1.6rem; }
    .main-header p  { margin: 4px 0 0; opacity: 0.85; font-size: 0.9rem; }
    .stat-box {
        background: white; border: 1px solid #E0E7EF;
        border-radius: 10px; padding: 14px 18px; text-align: center;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    .stat-box .val { font-size: 1.5rem; font-weight: 700; color: #1F4E79; }
    .stat-box .lbl { font-size: 0.78rem; color: #666; margin-top: 2px; }
    .ok-box   { background:#E2EFDA; border-radius:8px; padding:10px 14px; color:#375623; }
    .warn-box { background:#FFF2CC; border-radius:8px; padding:10px 14px; color:#7D6608; }
    .err-box  { background:#FCE4D6; border-radius:8px; padding:10px 14px; color:#843C0C; }
    div[data-testid="stSidebar"] { background: #F5F9FD; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
ALL_CH = [
    ('CH001','Top1 (Z1)',      '°C'),
    ('CH002','Top2 (Z2)',      '°C'),
    ('CH003','Top3 (Z3)',      '°C'),
    ('CH004','Top4 (Z4)',      '°C'),
    ('CH005','Top5 (Z5)',      '°C'),
    ('CH006','Top6 (Z6)',      '°C'),
    ('CH007','Top7 (Z7)',      '°C'),
    ('CH008','Bot1 (Z8)', '°C'),
    ('CH009','Bot2 (Z9)', '°C'),
    ('CH010','Bot3 (Z10)','°C'),
    ('CH011','Bot4 (Z11)','°C'),
    ('CH012','Bot5 (Z12)','°C'),
    ('CH013','Bot6 (Z13)','°C'),
    ('CH014','Bot7 (Z14)','°C'),
    ('CH015','O2 Exit',           'ppm'),
    ('CH016','Dryer zone1',       '°C'),
    ('CH017','Dryer zone2',       '°C'),
    ('CH018','O2 Entrance',       'ppm'),
]
DATA_OFFSET, RECORD_SIZE = 15008, 84
TOP_IDX = list(range(7))
BOT_IDX = list(range(7, 14))
O2_IDX  = [14, 17, 15, 16]

COLORS_TOP = ['#378ADD','#D85A30','#1D9E75','#7F77DD','#BA7517','#D4537E','#639922']
COLORS_BOT = ['#53B8E0','#E2864A','#2CB8A0','#9F77CC','#D0913A','#A85090','#5E9C20']
COLORS_O2  = {'CH015':'#2E75B6','CH018':'#C55A11','CH016':'#E36C09','CH017':'#974706'}

# ── Parser ────────────────────────────────────────────────────────────────────
def parse_dad(raw: bytes) -> list:
    records = []
    total = (len(raw) - DATA_OFFSET) // RECORD_SIZE
    for i in range(total):
        base = DATA_OFFSET + i * RECORD_SIZE
        hdr  = raw[base-8:base]
        yr,mo,dy,hr,mn,sc = hdr[0],hdr[1],hdr[2],hdr[3],hdr[4],hdr[5]
        try:
            ts = datetime(2000+yr, mo, dy, hr, mn, sc)
        except ValueError:
            continue
        rec = {'ts': ts}
        for ci,(ch,_,_) in enumerate(ALL_CH):
            v = struct.unpack_from('>h', raw, base + ci*4)[0]
            rec[ch] = round(v/10, 1) if v != -32767 else None
        records.append(rec)
    return records

# ── Excel Builder ─────────────────────────────────────────────────────────────
def build_excel(data, dt_start, dt_end, zones_sel):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    hdr_fill  = PatternFill('solid', fgColor='1F4E79')
    hdr_font  = Font(bold=True, color='FFFFFF', name='Arial', size=10)
    norm_font = Font(name='Arial', size=9)
    thin      = Side(border_style='thin', color='C0C8D0')
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)
    green_fill= PatternFill('solid', fgColor='E2EFDA')
    blue_fill = PatternFill('solid', fgColor='DEEAF1')
    red_fill  = PatternFill('solid', fgColor='FCE4D6')
    o2_fill   = PatternFill('solid', fgColor='FFF2CC')
    dryer_fill= PatternFill('solid', fgColor='F0FBF4')
    DAY_FILLS = ['EBF3FB','FFF9EC','F0FBF4','FDF2F8','F5F5F5']

    DATES = sorted(set(r['ts'].strftime('%Y/%m/%d') for r in data))
    date_fill_map = {d: PatternFill('solid', fgColor=DAY_FILLS[i%len(DAY_FILLS)])
                     for i,d in enumerate(DATES)}
    daily = defaultdict(lambda: {ch:[] for ch,_,_ in ALL_CH})
    for r in data:
        d = r['ts'].strftime('%Y/%m/%d')
        for ch,_,_ in ALL_CH:
            if r[ch] is not None: daily[d][ch].append(r[ch])

    def sfill(ch):
        if ch in ('CH015','CH018'): return o2_fill
        if ch in ('CH016','CH017'): return dryer_fill
        idx = int(ch[2:])-1
        return PatternFill('solid', fgColor='EBF3FB') if idx%2==0 else PatternFill('solid', fgColor='FFFFFF')

    def hcell(ws, col, row, val):
        c = ws.cell(row, col, val)
        c.fill=hdr_fill; c.font=hdr_font; c.border=border
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        return c

    def dcell(ws, col, row, val, fill=None, fmt=None, bold=False):
        c = ws.cell(row, col, val)
        c.fill=fill or PatternFill('solid',fgColor='FFFFFF')
        c.font=Font(name='Arial',size=9,bold=bold); c.border=border
        c.alignment=Alignment(horizontal='right' if isinstance(val,(int,float)) else 'left')
        if fmt: c.number_format=fmt
        return c

    sel_ch = [(ch,name,unit) for ch,name,unit in ALL_CH if name in zones_sel]
    wb = Workbook()
    n  = len(data)
    ds, de = 3, n+2
    report_name = f'{dt_start.strftime("%d-%m-%Y %H.%M")} to {dt_end.strftime("%d-%m-%Y %H.%M")}'

    # Raw Data
    ws = wb.active; ws.title='Raw Data'
    ws['A1'] = f'DAD Report — {report_name} | {n:,} records'
    ws.merge_cells(f'A1:{get_column_letter(3+len(sel_ch))}1')
    ws['A1'].font=Font(bold=True,name='Arial',size=12,color='1F4E79')
    ws.row_dimensions[1].height=22
    hcell(ws,1,2,'Timestamp'); hcell(ws,2,2,'Date'); hcell(ws,3,2,'Time')
    for i,(ch,name,unit) in enumerate(sel_ch):
        hcell(ws,4+i,2,f'{name}\n({ch})\n[{unit}]')
    ws.row_dimensions[2].height=48
    for ri,rec in enumerate(data,start=3):
        ts=rec['ts']; dp=ts.strftime('%Y/%m/%d'); tp=ts.strftime('%H:%M:%S')
        bg=date_fill_map.get(dp,PatternFill('solid',fgColor='FFFFFF'))
        dcell(ws,1,ri,ts.strftime('%Y/%m/%d %H:%M:%S'),bg)
        dcell(ws,2,ri,dp,bg); dcell(ws,3,ri,tp,bg)
        for i,(ch,_,_) in enumerate(sel_ch):
            c=ws.cell(ri,4+i,rec[ch]); c.fill=bg; c.font=norm_font
            c.border=border; c.alignment=Alignment(horizontal='right'); c.number_format='0.0'
    ws.column_dimensions['A'].width=22; ws.column_dimensions['B'].width=13; ws.column_dimensions['C'].width=10
    for i in range(len(sel_ch)): ws.column_dimensions[get_column_letter(4+i)].width=14
    ws.freeze_panes='A3'; ws.auto_filter.ref=f'A2:{get_column_letter(3+len(sel_ch))}{n+2}'

    # Summary
    ws2=wb.create_sheet('Summary Stats')
    ws2['A1']=f'Summary — {report_name}'
    ws2.merge_cells('A1:H1'); ws2['A1'].font=Font(bold=True,name='Arial',size=13,color='1F4E79')
    ws2['A2']=f'{n:,} records | {dt_start.strftime("%d/%m/%Y %H:%M")} – {dt_end.strftime("%d/%m/%Y %H:%M")}'
    ws2['A2'].font=Font(italic=True,name='Arial',size=9,color='595959')
    for ci,h in enumerate(['Zone / Channel','Channel','Unit','Avg','Min','Max','Range','Std Dev'],1):
        hcell(ws2,ci,4,h)
    for zi,(ch,name,unit) in enumerate(sel_ch):
        r=5+zi; col_l=get_column_letter(4+list(ch2 for ch2,_,_ in ALL_CH).index(ch))
        fill=sfill(ch)
        dcell(ws2,1,r,name,fill,bold=True); dcell(ws2,2,r,ch,fill); dcell(ws2,3,r,unit,fill)
        raw_col = get_column_letter(4 + next(i for i,(c,_,_) in enumerate(sel_ch) if c==ch))
        for ci2,formula,f2,fmt in [
            (4,f"=ROUND(AVERAGE('Raw Data'!{raw_col}{ds}:{raw_col}{de}),1)",green_fill,'0.0'),
            (5,f"=ROUND(MIN('Raw Data'!{raw_col}{ds}:{raw_col}{de}),1)",blue_fill,'0.0'),
            (6,f"=ROUND(MAX('Raw Data'!{raw_col}{ds}:{raw_col}{de}),1)",red_fill,'0.0'),
            (7,f"=F{r}-E{r}",PatternFill('solid',fgColor='FFC7CE'),'0.0'),
            (8,f"=ROUND(STDEV('Raw Data'!{raw_col}{ds}:{raw_col}{de}),2)",PatternFill('solid',fgColor='FFFFFF'),'0.00'),
        ]:
            c=ws2.cell(r,ci2,formula); c.fill=f2; c.font=norm_font; c.border=border
            c.alignment=Alignment(horizontal='right'); c.number_format=fmt
    for col,w in zip('ABCDEFGH',[24,10,8,12,12,12,12,10]): ws2.column_dimensions[col].width=w

    # Daily Summary
    ws3=wb.create_sheet('Daily Summary')
    ws3['A1']=f'Daily Average — {report_name}'
    ws3.merge_cells(f'A1:{get_column_letter(3+len(DATES))}1')
    ws3['A1'].font=Font(bold=True,name='Arial',size=12,color='1F4E79')
    hcell(ws3,1,3,'Name'); hcell(ws3,2,3,'Channel'); hcell(ws3,3,3,'Unit')
    for di,d in enumerate(DATES):
        y,m,day=d.split('/'); hcell(ws3,4+di,3,f'{day}/{m}/{y}')
    for zi,(ch,name,unit) in enumerate(sel_ch):
        r=4+zi; fill=sfill(ch)
        dcell(ws3,1,r,name,fill,bold=True); dcell(ws3,2,r,ch,fill); dcell(ws3,3,r,unit,fill)
        for di,date in enumerate(DATES):
            vals=daily[date][ch]
            avg=round(sum(vals)/len(vals),1) if vals else ''
            c=ws3.cell(r,4+di,avg); c.fill=green_fill; c.font=norm_font; c.border=border
            c.alignment=Alignment(horizontal='right')
            if avg!='': c.number_format='0.0'
    for col,w in zip('ABCDE',[24,10,8,14,14]): ws3.column_dimensions[col].width=w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ── Chart Builder ─────────────────────────────────────────────────────────────
def build_chart(data, zones_top, zones_bot, zones_o2, compare_ranges=None):
    rows = 3 if any([zones_top, zones_bot, zones_o2]) else 1
    active_rows = [bool(zones_top), bool(zones_bot), bool(zones_o2)]
    n_rows = sum(active_rows)
    if n_rows == 0:
        return None

    row_heights = []
    if zones_top: row_heights.append(0.4 if n_rows==1 else 0.35)
    if zones_bot: row_heights.append(0.4 if n_rows==1 else 0.35)
    if zones_o2:  row_heights.append(0.3 if n_rows==1 else 0.25)

    titles = []
    if zones_top: titles.append('🌡️ Top Zones')
    if zones_bot: titles.append('🌡️ Bottom Zones')
    if zones_o2:  titles.append('🫧 O2 & Dryer')

    fig = make_subplots(
        rows=n_rows, cols=1,
        shared_xaxes=True,
        subplot_titles=titles,
        vertical_spacing=0.06,
        row_heights=row_heights,
    )

    def add_traces(ch_list, colors, group, row):
        ts = [r['ts'] for r in data]
        for i,(ch,name,unit) in enumerate(ch_list):
            color = colors[i] if i < len(colors) else COLORS_O2.get(ch,'#888')
            fig.add_trace(go.Scatter(
                x=ts, y=[r[ch] for r in data],
                name=name, mode='lines',
                line=dict(color=color, width=1.5),
                legendgroup=group,
                legendgrouptitle_text=group.title() if i==0 else None,
                hovertemplate=f'<b>{name}</b>: %{{y:.1f}} {unit}<extra></extra>',
            ), row=row, col=1)

    row_num = 1
    if zones_top:
        sel = [(ch,name,unit) for ch,name,unit in ALL_CH[:7] if name in zones_top]
        colors = [COLORS_TOP[i] for i,(_,name,_) in enumerate(ALL_CH[:7]) if name in zones_top]
        add_traces(sel, colors, 'Top Zones', row_num); row_num += 1
    if zones_bot:
        sel = [(ch,name,unit) for ch,name,unit in ALL_CH[7:14] if name in zones_bot]
        colors = [COLORS_BOT[i] for i,(_,name,_) in enumerate(ALL_CH[7:14]) if name in zones_bot]
        add_traces(sel, colors, 'Bottom Zones', row_num); row_num += 1
    if zones_o2:
        sel = [(ch,name,unit) for ch,name,unit in ALL_CH[14:] if name in zones_o2]
        colors = [COLORS_O2.get(ch,'#888') for ch,_,_ in sel]
        add_traces(sel, colors, 'O2 & Dryer', row_num)

    # Compare ranges (vertical lines)
    if compare_ranges:
        for label, color, dt_s, dt_e in compare_ranges:
            for r in range(1, row_num+1):
                fig.add_vrect(x0=dt_s, x1=dt_e, fillcolor=color,
                              opacity=0.08, line_width=0, row=r, col=1)
                fig.add_vline(x=dt_s, line=dict(color=color, width=1, dash='dot'), row=r, col=1)
                fig.add_vline(x=dt_e, line=dict(color=color, width=1, dash='dot'), row=r, col=1)

    fig.update_layout(
        height=max(400, n_rows * 300),
        template='plotly_white',
        hovermode='x unified',
        hoverlabel=dict(bgcolor='white', font_size=11, namelength=30),
        legend=dict(
            groupclick='toggleitem',
            bgcolor='rgba(255,255,255,0.95)',
            bordercolor='#C0C8D0', borderwidth=1,
            font=dict(size=11), tracegroupgap=10,
        ),
        margin=dict(l=60, r=20, t=50, b=20),
        plot_bgcolor='white',
    )
    fig.update_xaxes(
        tickformat='%d/%m\n%H:%M',
        dtick=2*3600*1000,
        gridcolor='rgba(200,200,200,0.25)',
        tickfont=dict(size=10),
    )
    for r in range(1, n_rows+1):
        fig.update_yaxes(gridcolor='rgba(200,200,200,0.25)', row=r, col=1)
        if r < n_rows:
            fig.update_xaxes(showticklabels=False, row=r, col=1)

    return fig

# ══════════════════════════════════════════════════════════════════════════════
# MAIN UI
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="main-header">
    <h1>🏭 DAD Report Generator</h1>
    <p>Valeo APU3 — NB2 Furnace Data Analysis</p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('### 📂 Upload ไฟล์ .DAD')
    uploaded = st.file_uploader(
        'เลือกไฟล์ .DAD (หลายไฟล์ได้)',
        type=['DAD','dad'],
        accept_multiple_files=True,
    )

    st.markdown('---')
    st.markdown('### ⏰ ช่วงเวลา')

    from datetime import date, time
    col1, col2 = st.columns(2)
    with col1:
        d_start = st.date_input('วันเริ่ม', value=date.today())
        t_start = st.time_input('เวลาเริ่ม', value=time(7, 0))
    with col2:
        d_end = st.date_input('วันสิ้นสุด', value=date.today())
        t_end = st.time_input('เวลาสิ้นสุด', value=time(19, 0))

    dt_start = datetime.combine(d_start, t_start)
    dt_end   = datetime.combine(d_end,   t_end)

    st.markdown('---')
    st.markdown('### 📊 เลือก Zone ที่อยากดู')

    all_names_top = [name for _,name,_ in ALL_CH[:7]]
    all_names_bot = [name for _,name,_ in ALL_CH[7:14]]
    all_names_o2  = [name for _,name,_ in ALL_CH[14:]]

    with st.expander('🌡️ Top Zones', expanded=True):
        col_a, col_b = st.columns(2)
        sel_top = []
        for i, name in enumerate(all_names_top):
            col = col_a if i % 2 == 0 else col_b
            if col.checkbox(name, value=True, key=f'top_{i}'):
                sel_top.append(name)
    with st.expander('🌡️ Bottom Zones', expanded=True):
        col_a, col_b = st.columns(2)
        sel_bot = []
        for i, name in enumerate(all_names_bot):
            col = col_a if i % 2 == 0 else col_b
            if col.checkbox(name, value=True, key=f'bot_{i}'):
                sel_bot.append(name)
    with st.expander('🫧 O2 & Dryer', expanded=True):
        col_a, col_b = st.columns(2)
        sel_o2 = []
        for i, name in enumerate(all_names_o2):
            col = col_a if i % 2 == 0 else col_b
            if col.checkbox(name, value=True, key=f'o2_{i}'):
                sel_o2.append(name)

    st.markdown('---')
    st.markdown('### 🔀 เปรียบเทียบช่วงเวลา')
    compare_mode = st.toggle('เปิดโหมดเปรียบเทียบ')
    compare_ranges = []
    if compare_mode:
        n_compare = st.number_input('จำนวนช่วงที่เปรียบเทียบ', 1, 3, 2)
        COMP_COLORS = ['#FF6B6B','#4ECDC4','#45B7D1']
        for ci in range(int(n_compare)):
            with st.expander(f'ช่วงที่ {ci+1}', expanded=True):
                cc1, cc2 = st.columns(2)
                with cc1:
                    cd_s = st.date_input(f'วันเริ่ม', key=f'cd_s{ci}')
                    ct_s = st.time_input(f'เวลาเริ่ม', key=f'ct_s{ci}')
                with cc2:
                    cd_e = st.date_input(f'วันสิ้นสุด', key=f'cd_e{ci}')
                    ct_e = st.time_input(f'เวลาสิ้นสุด', key=f'ct_e{ci}')
                compare_ranges.append((
                    f'ช่วง {ci+1}', COMP_COLORS[ci],
                    datetime.combine(cd_s, ct_s),
                    datetime.combine(cd_e, ct_e),
                ))

    run_btn = st.button('🚀 Generate Report', type='primary', use_container_width=True)

# ── Main Area ─────────────────────────────────────────────────────────────────
if not uploaded:
    st.info('📂 กรุณา Upload ไฟล์ .DAD ในแถบซ้ายก่อนครับ')
    st.stop()

if dt_end <= dt_start:
    st.error('❌ เวลาสิ้นสุดต้องมากกว่าเวลาเริ่มต้นครับ')
    st.stop()

# Parse & Filter
@st.cache_data(show_spinner='⏳ กำลัง parse ไฟล์...')
def load_data(files_data, dt_s, dt_e):
    all_records = []
    for fname, raw in files_data:
        recs = parse_dad(raw)
        filtered = [r for r in recs if dt_s <= r['ts'] <= dt_e]
        all_records.extend(filtered)
    all_records.sort(key=lambda x: x['ts'])
    seen = set()
    unique = []
    for r in all_records:
        k = r['ts']
        if k not in seen:
            seen.add(k); unique.append(r)
    return unique

files_data = [(f.name, f.read()) for f in uploaded]
data = load_data(tuple((n,d) for n,d in files_data), dt_start, dt_end)

# ── Summary Cards ──────────────────────────────────────────────────────────────
duration = dt_end - dt_start
hrs = int(duration.total_seconds() // 3600)
mins = int((duration.total_seconds() % 3600) // 60)

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown(f'<div class="stat-box"><div class="val">{len(data):,}</div><div class="lbl">Records</div></div>', unsafe_allow_html=True)
with col2:
    st.markdown(f'<div class="stat-box"><div class="val">{len(uploaded)}</div><div class="lbl">ไฟล์ .DAD</div></div>', unsafe_allow_html=True)
with col3:
    st.markdown(f'<div class="stat-box"><div class="val">{hrs}h {mins}m</div><div class="lbl">ระยะเวลา</div></div>', unsafe_allow_html=True)
with col4:
    days_with_data = len(set(r['ts'].strftime('%Y/%m/%d') for r in data))
    st.markdown(f'<div class="stat-box"><div class="val">{days_with_data}</div><div class="lbl">วันที่มีข้อมูล</div></div>', unsafe_allow_html=True)

st.markdown('<div style="margin-top:16px"></div>', unsafe_allow_html=True)

# ── Coverage Check ────────────────────────────────────────────────────────────
with st.expander('📋 ตรวจสอบความครบถ้วนของข้อมูล', expanded=(len(data)==0)):
    day_counts = defaultdict(int)
    for r in data:
        day_counts[r['ts'].strftime('%Y/%m/%d')] += 1

    check_day = dt_start.date()
    all_ok = True
    while check_day <= dt_end.date():
        day_str = check_day.strftime('%Y/%m/%d')
        count = day_counts.get(day_str, 0)
        if count > 0:
            st.markdown(f'<div class="ok-box">✅ {check_day.strftime("%d/%m/%Y")} — {count:,} records</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="err-box">❌ {check_day.strftime("%d/%m/%Y")} — ไม่มีข้อมูล (ยังไม่ได้ upload)</div>', unsafe_allow_html=True)
            all_ok = False
        st.markdown('<div style="margin-top:4px"></div>', unsafe_allow_html=True)
        check_day += timedelta(days=1)

if len(data) == 0:
    st.error('🚨 ไม่มีข้อมูลในช่วงเวลาที่ระบุ กรุณาตรวจสอบช่วงวัน/เวลา หรือ Upload ไฟล์เพิ่มครับ')
    st.stop()

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab_chart, tab_summary, tab_compare, tab_download = st.tabs([
    '📈 Chart', '📊 Summary', '🔀 เปรียบเทียบ', '⬇️ Download Excel'
])

with tab_chart:
    fig = build_chart(data, sel_top, sel_bot, sel_o2,
                      compare_ranges if compare_mode else None)
    if fig:
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning('กรุณาเลือก Zone อย่างน้อย 1 Zone ในแถบซ้ายครับ')

with tab_summary:
    st.markdown(f'**ช่วงเวลา:** {dt_start.strftime("%d/%m/%Y %H:%M")} → {dt_end.strftime("%d/%m/%Y %H:%M")}')
    st.markdown(f'**Records:** {len(data):,}')
    st.markdown('---')

    all_sel_names = sel_top + sel_bot + sel_o2
    sel_channels = [(ch,name,unit) for ch,name,unit in ALL_CH if name in all_sel_names]

    import pandas as pd
    rows_summary = []
    for ch,name,unit in sel_channels:
        vals = [r[ch] for r in data if r[ch] is not None]
        if vals:
            rows_summary.append({
                'Zone / Channel': name,
                'Channel': ch,
                'Unit': unit,
                'Avg': round(sum(vals)/len(vals),1),
                'Min': round(min(vals),1),
                'Max': round(max(vals),1),
                'Range': round(max(vals)-min(vals),1),
            })
    if rows_summary:
        df = pd.DataFrame(rows_summary)
        st.dataframe(df, use_container_width=True, hide_index=True)

with tab_compare:
    if not compare_mode:
        st.info('เปิด "โหมดเปรียบเทียบ" ในแถบซ้ายก่อนครับ')
    else:
        for label, color, dt_s, dt_e in compare_ranges:
            seg = [r for r in data if dt_s <= r['ts'] <= dt_e]
            st.markdown(f'**{label}:** {dt_s.strftime("%d/%m/%Y %H:%M")} → {dt_e.strftime("%d/%m/%Y %H:%M")} ({len(seg):,} records)')
        if compare_ranges:
            st.markdown('---')
            import pandas as pd
            all_sel_names = sel_top + sel_bot + sel_o2
            sel_channels = [(ch,name,unit) for ch,name,unit in ALL_CH if name in all_sel_names]
            comp_rows = []
            for ch,name,unit in sel_channels:
                row = {'Zone': name, 'Unit': unit}
                for label,_,dt_s,dt_e in compare_ranges:
                    vals = [r[ch] for r in data if dt_s <= r['ts'] <= dt_e and r[ch] is not None]
                    row[f'{label} Avg'] = round(sum(vals)/len(vals),1) if vals else '-'
                    row[f'{label} Min'] = round(min(vals),1) if vals else '-'
                    row[f'{label} Max'] = round(max(vals),1) if vals else '-'
                comp_rows.append(row)
            st.dataframe(pd.DataFrame(comp_rows), use_container_width=True, hide_index=True)

with tab_download:
    st.markdown('### ⬇️ Download Excel Report')
    all_sel_names = sel_top + sel_bot + sel_o2
    if not all_sel_names:
        st.warning('กรุณาเลือก Zone ในแถบซ้ายก่อนครับ')
    else:
        if st.button('📊 สร้าง Excel Report', type='primary'):
            with st.spinner('⏳ กำลังสร้าง Excel...'):
                excel_data = build_excel(data, dt_start, dt_end, all_sel_names)
            fname = f'DAD Report {dt_start.strftime("%d-%m-%Y %H.%M")} to {dt_end.strftime("%d-%m-%Y %H.%M")}.xlsx'
            st.download_button(
                label='⬇️ Download Excel',
                data=excel_data,
                file_name=fname,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary',
                use_container_width=True,
            )
            st.success(f'✅ พร้อม Download: {fname}')
