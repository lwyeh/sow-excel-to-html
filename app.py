import streamlit as st
import pandas as pd
import openpyxl
import base64
import re
from datetime import datetime

# 設定網頁標題與外觀
st.set_page_config(page_title="竹三團 Excel 轉網頁產生器", page_icon="🏕️")

# ==========================================
# 輔助函式區塊
# ==========================================
def get_cell(df, row, col):
    try:
        val = df.iloc[row, col]
        return val if pd.notna(val) else ""
    except:
        return ""

def format_date(dt):
    if pd.isna(dt) or not isinstance(dt, datetime): return str(dt) if pd.notna(dt) else ""
    weekdays = ['一', '二', '三', '四', '五', '六', '日']
    return f"{dt.month}月{dt.day}日（週{weekdays[dt.weekday()]}）"

def linkify(text):
    if not isinstance(text, str): return text
    url_pattern = re.compile(r'(https?://[^\s]+)')
    text = url_pattern.sub(r'<a href="\1" target="_blank">\1</a>', text)
    return text

def extract_sheet_images(wb, sheet_name):
    if sheet_name not in wb.sheetnames: return []
    ws = wb[sheet_name]
    large_imgs = []
    for img in ws._images:
        data = img._data()
        if len(data) < 50000: continue
        magic = data[:4]
        mime = 'image/png' if magic == b'\x89PNG' else 'image/jpeg'
        b64 = base64.b64encode(data).decode()
        large_imgs.append({'mime': mime, 'b64': b64})
    return large_imgs

def generate_image_html(images, title_prefix):
    if not images: return ""
    html = f'<div class="sec-imgs">\n  <h3>{title_prefix}附圖</h3>\n  <div class="img-grid">\n'
    for i, img in enumerate(images):
        src = f"data:{img['mime']};base64,{img['b64']}"
        html += f'    <div class="img-card">\n      <img src="{src}" alt="附圖 {i+1}" onclick="openLightbox(this.src)">\n    </div>\n'
    html += '  </div>\n</div>\n'
    return html


# ==========================================
# 核心轉換邏輯
# ==========================================
def convert_excel_to_html(uploaded_file):
    # 讀取上傳的 Excel 檔案
    xls = pd.ExcelFile(uploaded_file)
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet_names = xls.sheet_names

    if "年度主題" not in sheet_names:
        return None, "錯誤：找不到「年度主題」工作表，請確認上傳了正確的 Excel 檔案。"

    # 解析 Header 資訊 (從年度主題)
    df_theme = pd.read_excel(uploaded_file, sheet_name="年度主題", header=None)
    gathering_info = str(get_cell(df_theme, 5, 1)).strip()
    year_theme = get_cell(df_theme, 6, 2)
    event_theme = get_cell(df_theme, 7, 2)
    event_date_raw = get_cell(df_theme, 7, 4)
    event_date = format_date(event_date_raw)
    event_time = str(get_cell(df_theme, 8, 4))
    event_loc = str(get_cell(df_theme, 12, 4)).strip()

    year_match = re.search(r'(\d+)年度', gathering_info)
    year = year_match.group(1) if year_match else "114"
    month = event_date_raw.month if isinstance(event_date_raw, datetime) else "08"
    
    file_title = f"竹三團{year}年{month:02d}月團集會總表"
    header_h1 = file_title
    subtitle = f"{gathering_info} {year_theme}".strip()

    # 人員清單解析
    people = []
    for i in range(9, 16):
        col1, col2 = get_cell(df_theme, i, 1), get_cell(df_theme, i, 2)
        col3, col4 = get_cell(df_theme, i, 3), get_cell(df_theme, i, 4)
        for c_title, c_name in [(col1, col2), (col3, col4)]:
            if c_title and ":" in str(c_title):
                parts = str(c_title).split(":", 1)
                if len(parts) == 2:
                    role, group = parts[0].strip(), parts[1].strip()
                    names = str(c_name).replace("\n", "、").split("、")
                    for name in names:
                        name = name.strip()
                        if name and name != "nan":
                            emoji = "🐜" if "蟻" in group else "🐝" if "蜂" in group else "🦌" if "鹿" in group else "🌱" if "育" in group else "💚" if "安心營" in group else "👤"
                            bg = "#fff3cd" if "蟻" in group else "#d8f3dc" if "蜂" in group else "#dbeafe" if ("鹿" in group or "育" in group) else "#f3e8ff"
                            people.append({"role": role, "group": group, "name": name, "emoji": emoji, "bg": bg})

    # CSS 樣式
    css = """
    * { box-sizing: border-box; font-family: 'Noto Sans TC', sans-serif; margin: 0; padding: 0; }
    body { background: #f4f9f4; color: #1b2e22; padding-bottom: 80px; }
    a { color: #1a6b9e; text-decoration: none; }
    a:hover { text-decoration: underline; }
    header { background: linear-gradient(135deg, #2d6a4f, #1b4332, #0a2918); color: #fff; padding: 40px 20px 30px; position: relative; overflow: hidden; }
    header::before { content:''; position:absolute; inset:0; background:url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" viewBox="0 0 100 100"><path d="M0,0 L100,100 M100,0 L0,100" stroke="%23ffffff" stroke-width="2" stroke-opacity="0.07"/></svg>') repeat; pointer-events:none; }
    .badge { display: inline-block; background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 20px; font-size: 13px; font-weight: 600; margin-bottom: 12px; color: #b7e4c7; }
    header h1 { font-size: 28px; font-weight: 900; margin-bottom: 8px; }
    header p { font-size: 16px; color: #b7e4c7; font-weight: 600; margin-bottom: 20px; }
    .chips { display: flex; flex-wrap: wrap; gap: 8px; }
    .chip { background: rgba(255,255,255,0.12); padding: 6px 14px; border-radius: 20px; font-size: 14px; display: flex; align-items: center; gap: 6px; }

    .tab-nav { position: sticky; top: 0; z-index: 100; background: #fff; border-bottom: 2px solid #c8e6c9; display: flex; overflow-x: auto; scrollbar-width: none; box-shadow: 0 2px 12px rgba(0,0,0,0.08); }
    .tab-nav::-webkit-scrollbar { display: none; }
    .tab { white-space: nowrap; padding: 16px 20px; font-size: 15px; font-weight: 600; color: #5a7060; border-bottom: 3px solid transparent; cursor: pointer; transition: 0.2s; }
    .tab.active { color: #2d6a4f; border-bottom-color: #2d6a4f; }

    main { padding: 20px; max-width: 800px; margin: 0 auto; }
    section { display: none; animation: fadeIn 0.3s; }
    section.active { display: block; }
    @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }

    .sec-title { font-size: 20px; font-weight: 700; color: #2d6a4f; margin-bottom: 16px; display: flex; align-items: center; gap: 8px; }
    .info-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 14px; margin-bottom: 24px; }
    .info-card { background: #fff; border-left: 4px solid #52b788; border-radius: 16px; padding: 16px; box-shadow: 0 4px 20px rgba(45,106,79,0.10); }
    .info-card .lbl { font-size: 12px; text-transform: uppercase; letter-spacing: 1px; color: #5a7060; margin-bottom: 6px; }
    .info-card .val { font-size: 15px; font-weight: 700; color: #1b2e22; }

    .people-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(195px, 1fr)); gap: 14px; }
    .person-card { background: #fff; border-radius: 16px; padding: 12px; display: flex; align-items: center; gap: 12px; box-shadow: 0 4px 20px rgba(45,106,79,0.10); }
    .avatar { width: 38px; height: 38px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 20px; }
    .p-info .role { font-size: 11px; color: #5a7060; margin-bottom: 2px; }
    .p-info .name { font-size: 14px; font-weight: 700; color: #1b2e22; }

    .dlbl { margin: 24px 0 16px; position: relative; }
    .dbadge { display: inline-block; background: #2d6a4f; color: #fff; border-radius: 20px; padding: 6px 14px; font-size: 14px; font-weight: 700; z-index: 1; position: relative; }
    .dlbl hr { position: absolute; top: 50%; width: 100%; border: none; border-top: 2px solid #d8f3dc; z-index: 0; }
    .timeline { position: relative; padding-left: 24px; margin-bottom: 30px; }
    .timeline::before { content:''; position: absolute; left: 7px; top: 0; bottom: 0; width: 2px; background: linear-gradient(to bottom, #52b788, #d8f3dc); }
    .tl-item { position: relative; background: #fff; border-radius: 16px; padding: 16px; margin-bottom: 16px; box-shadow: 0 4px 20px rgba(45,106,79,0.10); transition: transform 0.2s; }
    .tl-item:hover { transform: translateX(3px); }
    .tl-item::before { content:''; position: absolute; left: -22px; top: 20px; width: 10px; height: 10px; border-radius: 50%; background: #52b788; border: 2px solid #fff; box-shadow: 0 0 0 2px #52b788; }
    .tl-header { display: flex; align-items: center; gap: 12px; margin-bottom: 10px; }
    .time-tag { background: #d8f3dc; color: #2d6a4f; border-radius: 8px; padding: 4px 10px; font-size: 13px; font-weight: 700; }
    .tl-title { font-size: 15px; font-weight: 700; color: #1b2e22; }
    .tl-meta { font-size: 12px; color: #5a7060; margin-bottom: 10px; display: flex; flex-wrap: wrap; gap: 10px; }
    .tl-meta span { display: flex; align-items: center; gap: 4px; }
    .tl-content { font-size: 13px; color: #1b2e22; white-space: pre-wrap; line-height: 1.5; }
    .tl-note { margin-top: 8px; font-size: 12px; color: #5a7060; font-style: italic; }

    .loc-card { background: #fff; border-radius: 16px; padding: 16px; box-shadow: 0 4px 20px rgba(45,106,79,0.10); margin-bottom: 16px; }
    .loc-title { font-size: 16px; font-weight: 700; color: #1b2e22; margin-bottom: 12px; display: flex; align-items: center; gap: 8px; }
    .loc-info { font-size: 14px; margin-bottom: 8px; color: #5a7060; }
    .notebox { background: #d8f3dc; border-radius: 10px; padding: 12px 16px; font-size: 13px; color: #2d6a4f; line-height: 1.5; }
    .mapbtn { display: inline-block; margin-top: 10px; background: #1a73e8; color: #fff; padding: 8px 18px; border-radius: 20px; font-size: 13px; font-weight: 600; text-decoration: none; box-shadow: 0 2px 8px rgba(26,115,232,0.35); }

    .em-card { background: #fff5f5; border-left: 4px solid #ef4444; border-radius: 16px; padding: 16px; box-shadow: 0 4px 20px rgba(239,68,68,0.10); margin-bottom: 14px; }
    .em-title { font-size: 15px; font-weight: 700; color: #b91c1c; margin-bottom: 6px; }
    .em-phone { font-size: 16px; font-weight: 700; }
    .em-phone a { color: #dc2626; text-decoration: none; }

    .eq-group { margin-bottom: 24px; }
    .eq-title { font-size: 16px; font-weight: 700; color: #2d6a4f; padding-bottom: 8px; border-bottom: 2px solid #d8f3dc; margin-bottom: 16px; display: flex; align-items: center; gap: 8px; }
    .eq-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(225px, 1fr)); gap: 12px; }
    .eq-item { background: #fff; border-radius: 12px; padding: 10px 14px; display: flex; align-items: center; gap: 10px; box-shadow: 0 2px 10px rgba(45,106,79,0.05); }
    .eq-num { width: 22px; height: 22px; border-radius: 50%; background: #d8f3dc; color: #2d6a4f; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 700; flex-shrink: 0; }
    .eq-name { font-size: 14px; font-weight: 600; color: #1b2e22; flex-grow: 1; }
    .eq-note { font-size: 11px; color: #5a7060; }

    .sec-imgs { margin-top: 24px; }
    .sec-imgs h3 { font-size: 15px; font-weight: 700; color: #2d6a4f; margin-bottom: 14px; padding-bottom: 6px; border-bottom: 2px solid #d8f3dc; display: flex; align-items: center; gap: 6px; }
    .img-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 14px; }
    .img-card { background: #fff; border-radius: 16px; padding: 8px; box-shadow: 0 4px 20px rgba(45,106,79,0.10); overflow: hidden; }
    .img-card img { width: 100%; border-radius: 10px; display: block; cursor: zoom-in; }

    #btt { position: fixed; bottom: 20px; right: 20px; background: #2d6a4f; color: #fff; width: 42px; height: 42px; border-radius: 50%; opacity: 0; pointer-events: none; border: none; font-size: 20px; cursor: pointer; transition: 0.3s; box-shadow: 0 4px 12px rgba(45,106,79,0.3); display: flex; align-items: center; justify-content: center; z-index: 101;}
    #btt.show { opacity: 1; pointer-events: auto; }
    """

    # HTML 標頭
    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{header_h1}</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;600;700;900&display=swap" rel="stylesheet">
<style>{css}</style>
</head>
<body>

<header>
  <div class="badge">荒野親子團 竹三團</div>
  <h1>{header_h1}</h1>
  <p>{subtitle}</p>
  <div class="chips">
    <div class="chip">📅 {event_date}</div>
    <div class="chip">⏰ {event_time}</div>
    <div class="chip">📍 {event_loc}</div>
    <div class="chip">🎯 {event_theme}</div>
  </div>
</header>
"""

    # 定義 Tabs 規則
    TABS_SPEC = [
        ("概覽", "📋", "活動概覽", None),
        ("總表", "📅", "總流程", ["活動總表"]),
        ("蟻", "🐜", "小蟻", ["小蟻流程"]),
        ("蜂", "🐝", "小蜂", ["小蜂流程"]),
        ("鹿", "🦌", "小鹿", ["小鹿流程", "鹿團流程表"]),
        ("鹿育", "🦌🌱", "鹿育", ["鹿育流程"]),
        ("育成會", "🌱", "育成會", ["育成會流程"]),
        ("安心營", "💚", "安心營", ["安心營流程", "安心營"]),
        ("裝備", "🎒", "裝備清單", ["活動裝備檢查表"]),
        ("地點", "🗺️", "集合地點", ["活動場域圖面"]),
        ("緊急", "🚨", "緊急聯絡", ["緊急事件發生處理"])
    ]

    def find_sheet(possible_names):
        if not possible_names: return None
        for n in possible_names:
            if n in sheet_names: return n
        return None

    # 產生 Tab Navigation
    tabs_html = '<div class="tab-nav" id="tabNav">\n'
    for t_id, t_emoji, t_name, possible_names in TABS_SPEC:
        actual_sheet = find_sheet(possible_names) if possible_names else True
        if actual_sheet:
            tabs_html += f'  <div class="tab" onclick="switchTab(\'{t_id}\', this)">{t_emoji} {t_name}</div>\n'
    tabs_html += '</div>\n\n<main>\n'
    html += tabs_html

    # --- Section: 概覽 ---
    html += f'<section id="概覽" class="active">\n  <h2 class="sec-title">📋 基本資訊</h2>\n  <div class="info-grid">\n'
    html += f'    <div class="info-card"><div class="lbl">活動日期</div><div class="val">{event_date}</div></div>\n'
    html += f'    <div class="info-card"><div class="lbl">活動時間</div><div class="val">{event_time}</div></div>\n'
    html += f'    <div class="info-card"><div class="lbl">活動地點</div><div class="val">{event_loc}</div></div>\n'
    html += f'    <div class="info-card"><div class="lbl">年度主題</div><div class="val">{year_theme}</div></div>\n'
    html += f'    <div class="info-card"><div class="lbl">活動主題</div><div class="val">{event_theme}</div></div>\n'
    html += f'  </div>\n'

    html += f'  <h2 class="sec-title">👥 人員編組</h2>\n  <div class="people-grid">\n'
    for p in people:
        html += f'    <div class="person-card">\n      <div class="avatar" style="background:{p["bg"]}">{p["emoji"]}</div>\n'
        html += f'      <div class="p-info"><div class="role">{p["role"]} ({p["group"]})</div><div class="name">{p["name"]}</div></div>\n'
        html += f'    </div>\n'
    html += f'  </div>\n</section>\n\n'

    # --- Sections: 各分團流程 ---
    timeline_ids = ["總表", "蟻", "蜂", "鹿", "鹿育", "育成會", "安心營"]
    for t_id, t_emoji, t_name, possible_names in TABS_SPEC:
        if t_id not in timeline_ids: continue
        sheet_name = find_sheet(possible_names)
        if not sheet_name: continue

        html += f'<section id="{t_id}">\n  <h2 class="sec-title">{t_emoji} {t_name}</h2>\n'
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
        df = df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)
        
        header_idx = -1
        for i, r in df.iterrows():
            if any(pd.notna(x) and "時間" in str(x) for x in r):
                header_idx = i
                break
        
        if header_idx != -1:
            df.columns = [str(c).strip().replace(" ", "") for c in df.iloc[header_idx]]
            df = df.iloc[header_idx+1:].reset_index(drop=True)
            df = df.ffill(axis=0) 
            
            def get_col(row, *names):
                for name in names:
                    for c in df.columns:
                        if name in c:
                            v = row[c]
                            return str(v).strip() if pd.notna(v) else ""
                return ""
                
            html += f'  <div class="dlbl"><div class="dbadge">{event_date}</div><hr></div>\n  <div class="timeline">\n'
            for _, row in df.iterrows():
                time_str = get_col(row, "時間", "活動時間")
                if not time_str or time_str == "nan": continue
                
                title = get_col(row, "項目", "活動項目")
                content = get_col(row, "內容", "活動內容")
                owner = get_col(row, "負責", "負責人", "主領")
                loc = get_col(row, "場地", "地點")
                equip = get_col(row, "器材", "準備器材")
                note = get_col(row, "備註")
                
                html += f'    <div class="tl-item">\n      <div class="tl-header">\n        <div class="time-tag">{time_str}</div>\n        <div class="tl-title">{title}</div>\n      </div>\n'
                html += f'      <div class="tl-meta">\n'
                if owner and owner != "nan": html += f'        <span>👤 {owner}</span>\n'
                if loc and loc != "nan": html += f'        <span>📍 {loc}</span>\n'
                if equip and equip != "nan": html += f'        <span>🎒 {equip}</span>\n'
                html += f'      </div>\n'
                if content and content != "nan": html += f'      <div class="tl-content">{linkify(content)}</div>\n'
                if note and note != "nan": html += f'      <div class="tl-note">* {linkify(note)}</div>\n'
                html += f'    </div>\n'
            html += f'  </div>\n'
            
        images = extract_sheet_images(wb, sheet_name)
        html += generate_image_html(images, f"{t_emoji} {t_name}")
        html += f'</section>\n\n'

    # --- Section: 裝備清單 ---
    eq_sheet = find_sheet(["活動裝備檢查表"])
    if eq_sheet:
        html += f'<section id="裝備">\n  <h2 class="sec-title">🎒 裝備清單</h2>\n'
        df = pd.read_excel(uploaded_file, sheet_name=eq_sheet, header=None).dropna(how='all')
        
        current_group = ""
        for i, row in df.iterrows():
            valid_cells = [x for x in row if pd.notna(x) and str(x).strip() != ""]
            if len(valid_cells) == 1 and not any("項次" in str(x) for x in valid_cells):
                g_name = str(valid_cells[0]).strip()
                emoji = "🐜" if "蟻" in g_name else "🐝" if "蜂" in g_name else "🦌" if "鹿" in g_name else "🌱" if "育" in g_name else "🎒"
                if current_group: html += '  </div></div>\n'
                html += f'  <div class="eq-group">\n    <div class="eq-title">{emoji} {g_name}</div>\n    <div class="eq-grid">\n'
                current_group = g_name
            else:
                if current_group == "":
                    current_group = "通用裝備"
                    html += f'  <div class="eq-group">\n    <div class="eq-title">🎒 {current_group}</div>\n    <div class="eq-grid">\n'
                
                row_list = row.tolist()
                for j in range(len(row_list)-1):
                    c_val = str(row_list[j]).strip()
                    if c_val.isdigit() or (c_val and c_val != "nan" and j==0 and "項次" not in str(row_list)):
                        name = str(row_list[j+1]).strip() if pd.notna(row_list[j+1]) else ""
                        if name and name not in ["nan", "品名", "品        名"]:
                            note = ""
                            for k in range(j+2, min(j+4, len(row_list))):
                                n_val = str(row_list[k]).strip()
                                if n_val and n_val not in ["nan", "□", "備註"]:
                                    note = n_val; break
                            num = c_val if c_val.isdigit() else "-"
                            html += f'      <div class="eq-item">\n        <div class="eq-num">{num}</div>\n        <div class="eq-name">{name}</div>\n'
                            if note: html += f'        <div class="eq-note">{note}</div>\n'
                            html += f'      </div>\n'
        if current_group: html += '  </div></div>\n'
        
        images = extract_sheet_images(wb, eq_sheet)
        html += generate_image_html(images, "🎒 裝備清單")
        html += '</section>\n\n'

    # --- Section: 集合地點 ---
    loc_sheet = find_sheet(["活動場域圖面"])
    if loc_sheet:
        html += f'<section id="地點">\n  <h2 class="sec-title">🗺️ 集合地點</h2>\n'
        df = pd.read_excel(uploaded_file, sheet_name=loc_sheet, header=None).dropna(how='all')
        loc_text = []
        map_link = ""
        for _, row in df.iterrows():
            for val in row:
                if pd.notna(val):
                    s_val = str(val).strip()
                    if s_val.startswith("http"): map_link = s_val
                    elif s_val: loc_text.append(s_val)
        
        html += f'  <div class="loc-card">\n    <div class="loc-title">📍 集合資訊</div>\n'
        for t in loc_text: html += f'    <div class="loc-info">{t}</div>\n'
        if map_link: html += f'    <a href="{map_link}" target="_blank" class="mapbtn sky">🗺️ 導航至集合地點</a>\n'
        html += f'  </div>\n'
        
        images = extract_sheet_images(wb, loc_sheet)
        html += generate_image_html(images, "🗺️ 活動場域圖面")
        html += '</section>\n\n'

    # --- Section: 緊急聯絡 ---
    em_sheet = find_sheet(["緊急事件發生處理"])
    if em_sheet:
        html += f'<section id="緊急">\n  <h2 class="sec-title">🚨 緊急聯絡</h2>\n'
        df = pd.read_excel(uploaded_file, sheet_name=em_sheet, header=None).dropna(how='all')
        
        for _, row in df.iterrows():
            text = " ".join([str(x).strip() for x in row if pd.notna(x) and str(x).strip() != "nan"])
            if "緊急事件" in text or not text.strip(): continue
            
            html += f'  <div class="em-card">\n'
            phone_match = re.search(r'09\d{2}-\d{3}-\d{3}|0\d-\d+', text)
            if "電話" in text or phone_match:
                parts = re.split(r'電話[:：]?', text)
                name = parts[0].strip()
                phones_text = parts[1].strip() if len(parts)>1 else text
                html += f'    <div class="em-title">🏥 {name}</div>\n'
                nums = re.findall(r'[0-9\-]+', phones_text)
                for n in nums:
                    if len(n) > 5:
                        clean_n = re.sub(r'\D', '', n)
                        html += f'    <div class="em-phone">📞 <a href="tel:{clean_n}">{n}</a></div>\n'
            else:
                html += f'    <div class="em-title">{text}</div>\n'
            html += f'  </div>\n'
            
        images = extract_sheet_images(wb, em_sheet)
        html += generate_image_html(images, "🚨 緊急聯絡")
        html += '</section>\n\n'

    # --- Footer, Lightbox UI & JS ---
    js = """
    function switchTab(tabId, el) {
      document.querySelectorAll('section').forEach(sec => sec.classList.remove('active'));
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      document.getElementById(tabId).classList.add('active');
      el.classList.add('active');
      window.scrollTo({top: 0, behavior: 'smooth'});
    }
    document.addEventListener('DOMContentLoaded', () => {
      const firstTab = document.querySelector('.tab');
      if(firstTab) firstTab.classList.add('active');
    });

    const btt = document.getElementById('btt');
    window.addEventListener('scroll', () => {
      if (window.scrollY > 300) { btt.classList.add('show'); } 
      else { btt.classList.remove('show'); }
    });
    btt.addEventListener('click', () => {
      window.scrollTo({top: 0, behavior: 'smooth'});
    });
    
    function openLightbox(src) {
      document.getElementById('lb-img').src = src;
      document.getElementById('lightbox').style.display = 'flex';
    }
    """

    html += f"""
</main>
<button id="btt">↑</button>

<div id="lightbox" onclick="this.style.display='none'"
     style="display:none;position:fixed;inset:0;background:rgba(0,0,0,0.85);
            z-index:999;align-items:center;justify-content:center;cursor:zoom-out">
  <img id="lb-img" style="max-width:95vw;max-height:95vh;border-radius:8px;object-fit:contain">
</div>

<script>{js}</script>
</body>
</html>
"""

    return file_title, html


# ==========================================
# Streamlit 網頁介面區塊
# ==========================================
st.title("🏕️ 竹三團 Excel 轉網頁產生器")
st.markdown("請上傳當月的「團集會總表 Excel」，系統會自動將其轉換為互動式網頁檔 (HTML)。")

# 提供上傳元件
uploaded_file = st.file_uploader("點擊或拖曳上傳 Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.info("檔案上傳成功，處理中...")
    try:
        # 執行轉換
        file_title, result_html = convert_excel_to_html(uploaded_file)
        
        if result_html.startswith("錯誤"):
            st.error(result_html)
        else:
            st.success("🎉 轉換完成！請點擊下方按鈕下載 HTML 檔案。")
            
            # 提供下載按鈕
            st.download_button(
                label="⬇️ 下載轉換後的 HTML 檔",
                data=result_html,
                file_name=f"{file_title}.html",
                mime="text/html"
            )
    except Exception as e:
        st.error(f"轉換過程中發生錯誤：{e}")