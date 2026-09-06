import streamlit as st
import pandas as pd
import datetime
import os
import requests
import io
import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Set page config for web app
st.set_page_config(
    page_title="情緒觀察營 — 週總結線上申報與組長關懷系統",
    page_icon="🌟",
    layout="wide",  # Changed to wide for better table and matrix display
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .report-title {
        font-size: 2.2rem;
        font-weight: bold;
        color: #1B5E20;
        text-align: center;
        margin-bottom: 5px;
    }
    .report-subtitle {
        font-size: 1.15rem;
        color: #4E342E;
        text-align: center;
        margin-bottom: 25px;
    }
    .group-box {
        background-color: #E8F5E9;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #2E7D32;
        margin-bottom: 20px;
    }
    .success-box {
        background-color: #D1E7DD;
        color: #0F5132;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #BADBCC;
    }
    .reminder-template {
        background-color: #FFF3E0;
        padding: 18px;
        border-radius: 8px;
        border-left: 5px solid #FF9800;
        font-family: monospace;
        white-space: pre-wrap;
        margin-top: 10px;
    }
</style>
""", unsafe_allow_html=True)

# 1~18組完整分組名單 (來源自保溫分組名單)
groups_data = {
    "第 1 組": {
        "leader": "曾旭正、李麗斐（副）",
        "members": [
            "鍾昆原", "曾旭正", "周文祥", "王坤德", "薛榕婷", "馮豐隆",
            "曾美芳", "張筱雯", "關婉玲", "曾芳美", "吳美毅", "蔡沛妤", "李麗斐",
            "黃淑清", "陳萬和", "梁庭", "嚴惠英", "蔡淵輝", "陳淑芬", "郭瓈灧"
        ]
    },
    "第 2 組": {
        "leader": "楊韻蓉",
        "members": [
            "林翔均", "林文華", "吳煥崇", "黃曉鈺", "劉德政", "陳秀花",
            "楊韻蓉", "蔡宜均", "吳振仁", "王均晨", "林佳慧", "廖袖婷", "葉桂香", "溫婷伊", "陳田恬"
        ]
    },
    "第 3 組": {
        "leader": "陳威廷",
        "members": [
            "何進德", "紀辰諭", "劉俊麟", "張又方", "謝耀德", "葉迷妮", "王韶怡", "郭芳佑", "王碧宏", "柯春黛", "蕭名權", "陳淑媛", "陳薇莉", "張開明", "吳翊菱",
            "簡麗分", "黃玉宙", "戈德", "蘇献珍", "張顥嚴", "林秀婷", "柯春僖", "賴彩鈴", "林耘庄", "黃如榛", "許媛婷", "鍾亞昭", "林金泉"
        ]
    },
    "第 4 組": {
        "leader": "王國美",
        "members": [
            "王國美", "鄭翊汝", "李濱如", "洪筱婷", "陳聖昀", "賀正楨", "侯得正",
            "吳念真", "黃柏豪", "謝佳容", "劉以婕", "陳芸詞", "陳揚旻", "張麗卿", "潘秋華",
            "羅怡和", "留千惠", "張玉芳", "曾麗珍", "鄭美華", "周秀芬"
        ]
    },
    "第 5 組": {
        "leader": "郭宗川",
        "members": [
            "邱淑春", "林麗雪", "戴世昌", "黃壬癸", "蔡宗佑", "郭富貴", "李麗卿", "蕭玉卿", "陳秀美", "劉玉如", "賴幸絹",
            "陳毓旻", "郭宗川", "劉馨檀", "黃瓊慧", "石俊傑", "張志豪", "林淑寧", "楊舒涵", "李平裕", "李俊男", "尤躍勳"
        ]
    },
    "第 6 組": {
        "leader": "林家敬",
        "members": [
            "許婉瑜", "劉洹岑", "粘智淳", "張淑真", "劉淑美", "林俞彣",
            "彭聖森", "林德慶", "羅賢芳", "顏宗宏", "廖鶴翔", "林獻崇", "許婷婷",
            "徐立昇", "林家敬", "粘振清",
            "潘玟玲", "韓倩如", "劉潓雪", "黃乙娟"
        ]
    },
    "第 7 組": {
        "leader": "陳定群",
        "members": [
            "吳幸蓉", "陳定群", "林美芳", "魏媛真", "吳庭妤", "陳淑芳", "楊美蘭", "林伶雯", "張靜娟",
            "黃金輝", "謝萬蒲", "蕭淑勻", "鄭宜涵", "鄭秀娟", "鄭斐文", "徐續仁",
            "黃惠卿", "林靖芳", "吳璧合", "余雅玲"
        ]
    },
    "第 8 組": {
        "leader": "楊婉君",
        "members": [
            "楊婉君", "吳欣緯", "陳秋芬", "王怡誠", "王麗景", "章玲珠", "盧菀萱", "盧彥妤", "劉燕冰", "李瓊華",
            "楊素粉", "王宗平", "李淑瑛", "劉筱愛", "明子又", "蘇郁合", "黃翠月", "蘇國樑", "賴香雪"
        ]
    },
    "第 9 組": {
        "leader": "林如玉",
        "members": [
            "辛天送", "游建邦", "洪雅玲", "徐惠美", "何苑色", "喻榮華",
            "林如玉", "林淑娟", "莊淑娟", "周珈漩", "許湘惠", "柳書玉", "蘇鴻濱",
            "李俊宏", "李慧瑛", "白秀清", "蔡純慧", "周林美麗"
        ]
    },
    "第 10 組": {
        "leader": "黃振育",
        "members": [
            "陳世平", "黃振育", "楊崇誠", "張金㷸", "賴志雄", "吳東昱", "黃品齊", "朱重信", "楊昌學", "葉奕新",
            "林明珠", "胡詠茹", "李宜晏", "陳麗華", "廖曉萍", "趙如薫", "叢思瑋", "李學璁", "蔡春美", "謝佩穎"
        ]
    },
    "第 11 組": {
        "leader": "林美華",
        "members": [
            "賴素貞", "陳品樺", "蘇相瑜", "江盈潔", "阮芬", "林美華", "楊金珠", "蕭淑慧",
            "李素茹", "賴彥帛", "洪于晴", "何如甘", "林珍安", "吳核豫", "葉如玲", "許東華", "蔡菁雯", "林保秀",
            "徐丞億", "韋淑媛"
        ]
    },
    "第 12 組": {
        "leader": "邱品瑄",
        "members": [
            "蔡宗涵", "李淑芬", "江佳怡", "陳雅惠", "童小鈴", "洪唯",
            "何淑琴", "孫詩旻", "洪金城", "謝秀枝", "周素仰", "顏妙如", "陳昭華",
            "蔡麗滿", "邱品瑄", "楊紫璿", "王冠蕙", "簡達謙", "朱穆鳳", "賴盈達"
        ]
    },
    "第 13 組": {
        "leader": "高韻筑",
        "members": [
            "黃忠權", "丁秌全", "林于煒", "朱庭葦", "林珮圻", "鄭鈞懿", "陳彥伶", "陳珈暐",
            "高韻筑", "王上苹", "卓卉綺", "張婉玲", "吳孟霖", "張慈方", "彭美珠", "施瀞雅", "王慧玟", "黃蔚菁"
        ]
    },
    "第 14 組": {
        "leader": "蔡秉諺",
        "members": [
            "李岳璋", "蔡秉諺", "林奕瑄", "葉美鴻", "黃麗燕", "莊媛婷", "廖育君", "吳靜宜", "謝孟君",
            "高翊綺", "張在增", "蕭米佋", "陳桂宸", "顏少萱", "蔡明真", "郭芷萱", "吳承璇", "王馨儀", "葉姿佑", "秦育曐"
        ]
    },
    "第 15 組": {
        "leader": "高忠偉",
        "members": [
            "侯惠倫", "廖以萱", "邵杰（Jordan ）", "白豐銘", "曾品綸", "黃姿蓉", "高忠偉", "張文馨", "戴毓書", "陳星妤", "賴銘璇",
            "張家瑄", "侯惠芸", "黃沛瑋", "張雁菁", "蔡佩君", "鄭伃倢", "劉思嫺", "洪瑩蓁", "莊郁琳"
        ]
    },
    "第 16 組": {
        "leader": "偕魁元",
        "members": [
            "偕魁元", "王芳珠", "黃政和", "徐友信", "吳莉惠", "陳冠同", "劉妙娟", "沈承宗", "蔡雅雅", "凃文鳯",
            "李庭安", "蕭茲方", "林秀菁", "鄭雅文", "莊雅惠", "劉巧領", "劉姿吟", "林宜筠", "湯玉琦"
        ]
    },
    "第 17 組": {
        "leader": "蘇誌盈",
        "members": [
            "王素敏", "李世林", "黃淑芳", "胡瑜玲", "柯孟宜", "葉瓊雅", "劉宇容", "林淑樺",
            "蔡石福地", "蘇誌盈", "黃克經", "黃榮茂", "姚翠華", "林靜香", "鄭人豪", "張秋美", "邵士誠"
        ]
    },
    "第 18 組": {
        "leader": "陳麗華",
        "members": [
            "黃淑梅", "陳麗華", "楊月娥", "陳慶旺", "黃佩琪", "郭宇家",
            "林玲芬", "廖述強", "柯錦鑫", "陳素葉", "謝桂紅", "張湘瑜", "吳美蓉",
            "阮善如", "陳月美", "曹曼蓉", "謝秀錦", "陳錦鳳", "朱龍昌", "陳丁清"
        ]
    }
}

# Deduplicate members within each group
for g, data in groups_data.items():
    seen = set()
    deduped_members = []
    for m in data["members"]:
        if m not in seen:
            seen.add(m)
            deduped_members.append(m)
    groups_data[g]["members"] = deduped_members

# 週回報期間選項 (1 ~ 40 週)
weeks_list = [
    "W1 ( 2026/03/29 ~ 2026/04/04 )", "W2 ( 2026/04/05 ~ 2026/04/11 )", "W3 ( 2026/04/12 ~ 2026/04/18 )",
    "W4 ( 2026/04/19 ~ 2026/04/25 )", "W5 ( 2026/04/26 ~ 2026/05/02 )", "W6 ( 2026/05/03 ~ 2026/05/09 )",
    "W7 ( 2026/05/10 ~ 2026/05/16 )", "W8 ( 2026/05/17 ~ 2026/05/23 )", "W9 ( 2026/05/24 ~ 2026/05/30 )",
    "W10 ( 2026/05/31 ~ 2026/06/06 )", "W11 ( 2026/06/07 ~ 2026/06/13 )", "W12 ( 2026/06/14 ~ 2026/06/20 )",
    "W13 ( 2026/06/21 ~ 2026/06/27 )", "W14 ( 2026/06/28 ~ 2026/07/04 )", "W15 ( 2026/07/05 ~ 2026/07/11 )",
    "W16 ( 2026/07/12 ~ 2026/07/18 )", "W17 ( 2026/07/19 ~ 2026/07/25 )", "W18 ( 2026/07/26 ~ 2026/08/01 )",
    "W19 ( 2026/08/02 ~ 2026/08/08 )", "W20 ( 2026/08/09 ~ 2026/08/15 )", "W21 ( 2026/08/16 ~ 2026/08/22 )",
    "W22 ( 2026/08/23 ~ 2026/08/29 )", "W23 ( 2026/08/30 ~ 2026/09/05 )", "W24 ( 2026/09/06 ~ 2026/09/12 )",
    "W25 ( 2026/09/13 ~ 2026/09/19 )", "W26 ( 2026/09/20 ~ 2026/09/26 )", "W27 ( 2026/09/27 ~ 2026/10/03 )",
    "W28 ( 2026/10/04 ~ 2026/10/10 )", "W29 ( 2026/10/11 ~ 2026/10/17 )", "W30 ( 2026/10/18 ~ 2026/10/24 )",
    "W31 ( 2026/10/25 ~ 2026/10/31 )", "W32 ( 2026/11/01 ~ 2026/11/07 )", "W33 ( 2026/11/08 ~ 2026/11/14 )",
    "W34 ( 2026/11/15 ~ 2026/11/21 )", "W35 ( 2026/11/22 ~ 2026/11/28 )", "W36 ( 2026/11/29 ~ 2026/12/05 )",
    "W37 ( 2026/12/06 ~ 2026/12/12 )", "W38 ( 2026/12/13 ~ 2026/12/19 )", "W39 ( 2026/12/20 ~ 2026/12/26 )",
    "W40 ( 2026/12/27 ~ 2027/01/02 )"
]

# Path to save data
DB_FILE = "emotional_reports.csv"

# Function to load database
def load_data():
    if os.path.exists(DB_FILE):
        try:
            return pd.read_csv(DB_FILE)
        except Exception:
            pass
    # Return empty template dataframe
    return pd.DataFrame(columns=[
        "時間戳記", "組別", "保溫組長", "個人姓名", "電子郵件", 
        "個人編號", "性別", "當週回報期間", "當週填表張數", 
        "累積填表張數", "心得或對法師提問"
    ])

# Function to save record (local + cloud try)
def save_record(new_row, gas_url=""):
    # 1. Save locally
    df = load_data()
    new_df = pd.DataFrame([new_row])
    df = pd.concat([df, new_df], ignore_index=True)
    df.to_csv(DB_FILE, index=False, encoding='utf-8-sig')
    
    # 2. Try saving to Google Sheets via GAS if configured
    if gas_url and gas_url.strip():
        try:
            # Send POST request to Google Apps Script
            response = requests.post(gas_url.strip(), json=new_row, timeout=8)
            if response.status_code == 200:
                return True, "本機備份成功，且順利同步寫入 Google 雲端！"
            else:
                return True, f"本機備份成功，但雲端同步失敗 (狀態碼: {response.status_code})"
        except Exception as e:
            return True, f"本機備份成功，但雲端同步發生錯誤: {str(e)}"
    return True, "填報紀錄已成功儲存於本地伺服器。"

# Sidebar configuration
with st.sidebar:
    st.image("https://img.icons8.com/clouds/200/database.png", width=80)
    st.markdown("### ☁️ 雲端同步設定")
    st.write("設定與 Google 雲端試算表的同步連結。")
    
    save_type = st.radio("選擇儲存方式：", ["本機儲存 + Google Apps Script 同步", "僅本機 CSV 儲存"])
    
    gas_api_url = ""
    if save_type == "本機儲存 + Google Apps Script 同步":
        gas_api_url = st.text_input(
            "請貼上 Google Apps Script 部署網址：", 
            placeholder="https://script.google.com/macros/s/.../exec",
            value=st.session_state.get("gas_url", "")
        )
        if gas_api_url:
            st.session_state["gas_url"] = gas_api_url
            st.success("✅ 已暫存 Google 雲端連線 URL！")
        else:
            st.warning("⚠️ 尚未設定 Google 雲端 API，目前資料僅會儲存在本機。")


# Helper to generate beautiful Word Document (.docx) Weekly Thoughts Report
def generate_docx_report(group_name, leader, week, submissions, roster_list):
    doc = Document()
    
    # Page setup - Margins (1 inch)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    # Set default font style
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Microsoft JhengHei'
    font.size = Pt(11)
    
    # Document Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(f"📋 情緒觀察營 — {group_name} 每週心得報告")
    run_title.font.name = 'Microsoft JhengHei'
    run_title.font.size = Pt(18)
    run_title.bold = True
    run_title.font.color.rgb = RGBColor(27, 94, 32) # Dark Green
    
    # Subtitle Info Table
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_meta = p_meta.add_run(f"回報週別：{week}  |  保溫組長：{leader}  |  報表產出日期：{datetime.date.today().strftime('%Y/%m/%d')}")
    run_meta.font.size = Pt(10)
    run_meta.font.color.rgb = RGBColor(100, 110, 100)
    
    # Add an empty paragraph as space
    doc.add_paragraph()
    
    # Add a heading for Summary Table
    h_sum = doc.add_paragraph()
    run_h_sum = h_sum.add_run("📊 當週填表數據摘要表")
    run_h_sum.font.size = Pt(14)
    run_h_sum.bold = True
    run_h_sum.font.color.rgb = RGBColor(27, 94, 32)
    
    # Create Table
    # Col widths in Inches: 1.2, 1.2, 1.2, 1.2, 1.2 -> Total 6 inches
    cols_data = [("個人編號", Inches(1.2)), ("組員姓名", Inches(1.2)), ("填報狀態", Inches(1.2)), ("當週張數", Inches(1.2)), ("累計張數", Inches(1.2))]
    table = doc.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Format Table Header
    hdr_cells = table.rows[0].cells
    for i, (name, width) in enumerate(cols_data):
        hdr_cells[i].text = name
        hdr_cells[i].width = width
        # Center align text in cell
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Make font white and bold
        for r in p.runs:
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.name = 'Microsoft JhengHei'
            r.font.size = Pt(10.5)
            
        # Shading cell
        tcPr = hdr_cells[i]._tc.get_or_add_tcPr()
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), '1B5E20')  # Dark green hex
        shading.set(qn('w:val'), 'clear')
        tcPr.append(shading)
        
    # Map submissions to name for lookup
    sub_map = {row["個人姓名"]: row for _, row in submissions.iterrows()}
    
    # Fill Table Rows
    total_sheets = 0
    submitted_count = 0
    
    for row in roster_list:
        m_name = row["個人姓名"]
        m_id = row["組員號碼"]
        
        row_cells = table.add_row().cells
        
        # Determine status
        status = "❌ 未填"
        w_sheets = "-"
        c_sheets = "-"
        
        if m_name in sub_map:
            status = "✅ 已填"
            submitted_count += 1
            w_sheets = str(int(sub_map[m_name]["當週填表張數"]))
            c_sheets = str(int(sub_map[m_name]["累積填表張數"]))
            total_sheets += int(sub_map[m_name]["當週填表張數"])
            
        row_data = [m_id, m_name, status, w_sheets, c_sheets]
        
        for i, val in enumerate(row_data):
            row_cells[i].text = val
            row_cells[i].width = cols_data[i][1]
            p = row_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.name = 'Microsoft JhengHei'
                r.font.size = Pt(10)
                if status == "❌ 未填" and i == 2:
                    r.font.color.rgb = RGBColor(183, 28, 28) # Dark Red for Unsubmitted
                elif status == "✅ 已填" and i == 2:
                    r.font.color.rgb = RGBColor(46, 125, 50) # Dark Green for Submitted
                    
    # Metrics paragraph
    p_metrics = doc.add_paragraph()
    p_metrics.paragraph_format.space_before = Pt(12)
    p_metrics.paragraph_format.space_after = Pt(24)
    r_met = p_metrics.add_run(f"💡 本週小組統計：應填報 {len(roster_list)} 人，已填報 {submitted_count} 人，未填報 {len(roster_list) - submitted_count} 人。當週填表總張數：{total_sheets} 張。")
    r_met.font.size = Pt(10)
    r_met.font.bold = True
    r_met.font.color.rgb = RGBColor(100, 100, 100)
    
    # Add Heading for Detailed Thoughts
    h_detail = doc.add_paragraph()
    run_h_detail = h_detail.add_run("💬 組員心得與對法師提問明細")
    run_h_detail.font.size = Pt(14)
    run_h_detail.bold = True
    run_h_detail.font.color.rgb = RGBColor(27, 94, 32)
    h_detail.paragraph_format.space_after = Pt(12)
    
    # Fill detailed thoughts
    for row in roster_list:
        m_name = row["個人姓名"]
        m_id = row["組員號碼"]
        
        p_mem = doc.add_paragraph()
        p_mem.paragraph_format.space_before = Pt(8)
        p_mem.paragraph_format.space_after = Pt(2)
        r_mem = p_mem.add_run(f"【{m_id}】 {m_name}")
        r_mem.bold = True
        r_mem.font.size = Pt(11.5)
        r_mem.font.color.rgb = RGBColor(78, 52, 46) # Warm brown
        
        if m_name in sub_map:
            sub = sub_map[m_name]
            sheets = sub["當週填表張數"]
            cum_sheets = sub["累積填表張數"]
            feedback = sub["心得或對法師提問"]
            
            p_sheets = doc.add_paragraph()
            p_sheets.paragraph_format.left_indent = Inches(0.25)
            p_sheets.paragraph_format.space_after = Pt(2)
            r_sh = p_sheets.add_run(f"📊 本週填表：{int(sheets)} 張 | 累計填表：{int(cum_sheets)} 張")
            r_sh.font.size = Pt(9.5)
            r_sh.italic = True
            r_sh.font.color.rgb = RGBColor(120, 120, 120)
            
            p_feed = doc.add_paragraph()
            p_feed.paragraph_format.left_indent = Inches(0.25)
            p_feed.paragraph_format.space_after = Pt(12)
            
            if str(feedback).strip() in ["", "無", "nan", "None"]:
                r_feed = p_feed.add_run("（組員本週無填寫心得或提問）")
                r_feed.font.color.rgb = RGBColor(150, 150, 150)
                r_feed.italic = True
            else:
                r_feed = p_feed.add_run(str(feedback).strip())
                r_feed.font.color.rgb = RGBColor(33, 33, 33)
                
        else:
            p_feed = doc.add_paragraph()
            p_feed.paragraph_format.left_indent = Inches(0.25)
            p_feed.paragraph_format.space_after = Pt(12)
            r_feed = p_feed.add_run("❌ 本週尚未提交週總結問卷。")
            r_feed.font.color.rgb = RGBColor(183, 28, 28)
            r_feed.italic = True
            
        # Draw a subtle divider paragraph
        p_div = doc.add_paragraph()
        p_div.paragraph_format.space_before = Pt(4)
        p_div.paragraph_format.space_after = Pt(4)
        r_div = p_div.add_run("―" * 50)
        r_div.font.color.rgb = RGBColor(230, 230, 230)
        r_div.font.size = Pt(8)
        
    # Save docx to BytesIO object
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# Render header
st.markdown('<div class="report-title">🌟 情緒觀察營 — 申報與關懷管理系統</div>', unsafe_allow_html=True)
st.markdown('<div class="report-subtitle">第10組及全體保溫小組每週填報與組長一鍵關懷平台</div>', unsafe_allow_html=True)

# Tabs
tab1, tab2, tab3 = st.tabs(["📝 每週填報問卷", "👥 組長關懷專區", "📊 數據後台管理"])

# ----------------- Tab 1: Form Fill -----------------
with tab1:
    st.info("💡 敬請於亞洲時間每週六 23:00 前完成填寫並回傳，感謝您的配合。統計區間為「前週日至本週六」之合計填表張數。")
    
    with st.form("reporting_form", clear_on_submit=True):
        st.subheader("第一步：選擇您的組別與名字")
        
        col1, col2 = st.columns(2)
        with col1:
            group_selected = st.selectbox("請選擇您的組別 *", list(groups_data.keys()), index=9) # Default to Group 10
        
        with col2:
            names_list = groups_data[group_selected]["members"]
            name_selected = st.selectbox("請選擇您的名字 *", names_list)
            
        leader = groups_data[group_selected]["leader"]
        
        # Calculate auto-generated Member ID
        group_num = group_selected.replace("第 ", "").replace(" 組", "")
        try:
            member_idx = names_list.index(name_selected) + 1
            auto_member_id = f"{group_num}-{member_idx}"
        except ValueError:
            auto_member_id = f"{group_num}-X"
            
        st.markdown(f'<div class="group-box">ℹ️ 您選取的是 <b>{group_selected}</b>（保溫組長：{leader}）| 系統配發個人編號：<b>{auto_member_id}</b></div>', unsafe_allow_html=True)
        
        st.write("---")
        st.subheader("第二步：填寫個人與填表資訊")
        
        col3, col4 = st.columns(2)
        with col3:
            email = st.text_input("電子郵件 * (Email)", placeholder="example@gmail.com")
            
            # Female detection list to auto check female gender
            female_names = [
                "謝佩穎", "謝佳容", "謝孟君", "蔡佩君", "許婷婷", "張靜娟", "李慧瑛", 
                "張慈方", "劉姿吟", "胡詠茹", "李宜晏", "陳麗華", "廖曉萍", "趙如薫", 
                "叢思瑋", "蔡春美", "曾美芳", "張筱雯", "關婉玲", "曾芳美", "吳美毅", 
                "蔡沛妤", "李麗斐", "黃淑清", "梁庭", "嚴惠英", "陳淑芬", "郭瓈灧", 
                "黃曉鈺", "陳秀花", "蔡宜均", "林佳慧", "廖袖婷", "葉桂香", "溫婷伊", 
                "陳田恬", "張又方", "葉迷妮", "王韶怡", "柯春黛", "陳淑媛", "陳薇莉", 
                "吳翊菱", "簡麗分", "蘇献珍", "林秀婷", "柯春僖", "賴彩鈴", "林耘庄", 
                "黃如榛", "許媛婷", "鄭翊汝", "李濱如", "洪筱婷", "劉以婕", "陳芸詞", 
                "張麗卿", "潘秋華", "留千惠", "張玉芳", "曾麗珍", "鄭美華", "周秀芬"
            ]
            default_gender_idx = 1 if name_selected in female_names else 0
            gender = st.radio("性別 *", ["男", "女"], horizontal=True, index=default_gender_idx)
        
        with col4:
            personal_id = st.text_input("個人編號 (免填，系統會依組員排序自動對應)", value=auto_member_id, disabled=True)
            period = st.selectbox("當週回報期間 *", weeks_list, index=23) # Default to W24 or typical current week
            
        col5, col6 = st.columns(2)
        with col5:
            weekly_sheets = st.number_input("當週填表張數 *", min_value=0, max_value=100, value=0, step=1)
        with col6:
            cum_sheets = st.number_input("自2026/3/6起累積張數(含當週) *", min_value=0, max_value=1000, value=0, step=1)
            
        st.write("---")
        st.subheader("第三步：心得或對法師提問（選填）")
        feedback = st.text_area("心得或對法師提問", placeholder="請輸入您當週的觀察心得，或是想對法師提出的疑問...")
        
        # Submit button
        submit_btn = st.form_submit_button("提交問卷 📨")
        
        if submit_btn:
            if not email:
                st.error("❌ 請輸入您的電子郵件！")
            elif "@" not in email:
                st.error("❌ 請輸入正確格式的電子郵件！")
            else:
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                new_row = {
                    "時間戳記": timestamp,
                    "組別": group_selected,
                    "保溫組長": leader,
                    "個人姓名": name_selected,
                    "電子郵件": email,
                    "個人編號": auto_member_id,
                    "性別": gender,
                    "當週回報期間": period,
                    "當週填表張數": int(weekly_sheets),
                    "累積填表張數": int(cum_sheets),
                    "心得或對法師提問": feedback if feedback else "無"
                }
                
                success, msg = save_record(new_row, gas_url=gas_api_url)
                if success:
                    st.markdown(f"""
                    <div class="success-box">
                        <h3>🎉 提交成功！</h3>
                        <p>感謝 <b>{name_selected}</b> 的回報！您的填表紀錄已妥善儲存。</p>
                        <p>💡 {msg}</p>
                        <ul>
                            <li><b>個人編號：</b>{auto_member_id}</li>
                            <li><b>回報期間：</b>{period}</li>
                            <li><b>當週張數：</b>{weekly_sheets} 張</li>
                            <li><b>累計張數：</b>{cum_sheets} 張</li>
                        </ul>
                    </div>
                    """, unsafe_allow_html=True)


# ----------------- Tab 2: Group Leader Dashboard (NEW!) -----------------
with tab2:
    st.subheader("👥 保溫組長關懷與數據查閱專區")
    st.write("組長可以透過此平台，快速查看自己組別組員在特定日期範圍內的填表狀況，並一鍵產出未填寫組員的溫馨提醒文字！")
    
    col_l1, col_l2 = st.columns([1, 2])
    
    with col_l1:
        st.markdown("### 🔍 查詢設定")
        leader_group = st.selectbox("請選擇您要讀取的組別：", list(groups_data.keys()), index=9)
        
        # Week Range logic
        st.markdown("**📅 選擇回報週別範圍：**")
        start_week = st.selectbox("開始週別：", weeks_list, index=20) # Default from W21
        end_week = st.selectbox("結束週別：", weeks_list, index=23) # Default to W24
        
        # Get selected weeks list
        try:
            start_idx = weeks_list.index(start_week)
            end_idx = weeks_list.index(end_week)
            if start_idx <= end_idx:
                selected_weeks_range = weeks_list[start_idx:end_idx+1]
            else:
                selected_weeks_range = weeks_list[end_idx:start_idx+1]
        except Exception:
            selected_weeks_range = weeks_list[20:24]
            
        st.success(f"已選取共 **{len(selected_weeks_range)} 週** 的查詢區間")
        
        # Reminder Target Week
        st.write("---")
        st.markdown("### 📣 LINE 關懷提醒設定")
        st.write("選擇欲針對哪一週產生「未填報提醒」：")
        reminder_target_week = st.selectbox("目標提醒週別：", selected_weeks_range, index=len(selected_weeks_range)-1)

    with col_l2:
        st.markdown(f"### 📊 【{leader_group}】填表狀況對照表")
        st.caption(f"保溫組長：{groups_data[leader_group]['leader']} | 查詢區間：{start_week.split(' (')[0]} ～ {end_week.split(' (')[0]}")
        
        # Load data and filter
        all_data = load_data()
        group_roster = groups_data[leader_group]["members"]
        g_num = leader_group.replace("第 ", "").replace(" 組", "")
        
        # Create base roster dataframe with auto-generated ID
        roster_list = []
        for idx, m_name in enumerate(group_roster):
            roster_list.append({
                "組員號碼": f"{g_num}-{idx+1}",
                "個人姓名": m_name,
                "編號排序": idx+1
            })
        roster_base_df = pd.DataFrame(roster_list)
        
        # Build matrix
        matrix_data = roster_base_df.copy()
        
        # Calculate stats for the selected group
        for wk in selected_weeks_range:
            wk_label = wk.split(" (")[0] # e.g. "W1"
            
            # Submissions in this week for this group
            wk_submissions = all_data[(all_data["組別"] == leader_group) & (all_data["當週回報期間"] == wk)]
            
            status_col = []
            for m_name in group_roster:
                user_record = wk_submissions[wk_submissions["個人姓名"] == m_name]
                if not user_record.empty:
                    sheets = user_record["當週填表張數"].values[0]
                    status_col.append(f"✅ 已填 ({int(sheets)}張)")
                else:
                    status_col.append("❌ 未填")
            matrix_data[wk_label] = status_col
        
        # Clean and sort matrix for display
        display_matrix = matrix_data.drop(columns=["編號排序"]).set_index("組員號碼")
        st.dataframe(display_matrix, use_container_width=True)
        
        # ----------------- Unsubmitted & LINE Reminder -----------------
        st.write("---")
        st.markdown(f"### 🚨 {reminder_target_week.split(' (')[0]} 未填寫關懷名單")
        
        # Get submissions for the targeted week
        target_wk_submissions = all_data[(all_data["組別"] == leader_group) & (all_data["當週回報期間"] == reminder_target_week)]
        submitted_names = target_wk_submissions["個人姓名"].tolist()
        
        # Find unsubmitted
        unsubmitted_members = []
        for row in roster_list:
            if row["個人姓名"] not in submitted_names:
                unsubmitted_members.append(row)
                
        if len(unsubmitted_members) == 0:
            st.balloons()
            st.success(f"🎉 太棒了！【{leader_group}】在 {reminder_target_week.split(' (')[0]} 的全體組員皆已完成填報問卷！")
        else:
            col_u1, col_u2 = st.columns([1, 1])
            with col_u1:
                st.warning(f"⚠️ 尚有 **{len(unsubmitted_members)}** 位同修尚未完成本週填寫：")
                
                # Format a table of unsubmitted members
                unsub_df = pd.DataFrame(unsubmitted_members).drop(columns=["編號排序"])
                st.dataframe(unsub_df, use_container_width=True, hide_index=True)
                
            with col_u2:
                st.markdown("💬 **LINE 溫馨關懷提醒範本：**")
                # Create copyable text template
                unsub_mentions = " ".join([f"@{m['個人姓名']}" for m in unsubmitted_members])
                
                template_text = f"""📢【情緒觀察表 — 每週填報溫馨提醒】

各位【{leader_group}】的同修師兄姐吉祥：

感恩大家一週以來精進修行與自我觀察。🙏
目前本週【{reminder_target_week}】之情緒觀察填表統計正在進行中。

若您已完成填表登記，非常感恩您的配合！
若本週尚未完成，請撥冗 1 分鐘點擊下方平台連結回報，感恩合十。❤️

📝 本週尚未完成登記的同修（請抽空上網補登喔）：
👉 {unsub_mentions}

（若您剛剛已提交，系統同步可能有些許延遲，請忽略此訊息。再次感恩大家一同增長智慧與慈悲！🌱）"""
                st.text_area("您可以直接複製以下文字發到 LINE 群組中：", value=template_text, height=220)

    # ----------------- Detailed Data For Leader -----------------
    st.write("---")
    st.markdown("### 📋 組員填報詳細內容清單")
    st.caption("依據組員號碼順序排列，方便組長閱讀心得與對法師的提問")
    
    # Filter detailed entries
    leader_df = all_data[(all_data["組別"] == leader_group) & (all_data["當週回報期間"].isin(selected_weeks_range))].copy()
    
    if not leader_df.empty:
        # Extract ID suffix for numerical sorting
        def get_sort_key(id_val):
            try:
                parts = str(id_val).split("-")
                return int(parts[-1])
            except Exception:
                return 999
                
        leader_df["_sort_key"] = leader_df["個人編號"].apply(get_sort_key)
        leader_df = leader_df.sort_values(by=["_sort_key", "當週回報期間"]).drop(columns=["_sort_key"])
        
        st.dataframe(
            leader_df[["個人編號", "個人姓名", "當週回報期間", "當週填表張數", "累積填表張數", "心得或對法師提問", "時間戳記"]],
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("💡 目前在此查詢期間內，該組尚無組員填報資料。")



    # ----------------- One-Click Weekly Thoughts Report (.docx) -----------------
    st.write("---")
    st.markdown("### 📝 一鍵生成每週心得報告 (.docx)")
    st.write("可為當前選取的組別生成一份格式精美的 Word (.docx) 文件，包含組員號碼、姓名、當週填表張數、累計張數以及心得或提問。適合直接呈報給法師或彙整關懷。")
    
    target_week_submissions = all_data[(all_data["組別"] == leader_group) & (all_data["當週回報期間"] == reminder_target_week)]
    
    st.write(f"📊 正在準備生成：**{leader_group}** 在 **{reminder_target_week.split(' (')[0]}** 的報告")
    
    doc_bio = generate_docx_report(
        group_name=leader_group,
        leader=groups_data[leader_group]["leader"],
        week=reminder_target_week,
        submissions=target_week_submissions,
        roster_list=roster_list
    )
    
    # Use columns to align button
    btn_col1, btn_col2 = st.columns([1, 2])
    with btn_col1:
        st.download_button(
            label="📥 一鍵導出 Word 格式心得報告",
            data=doc_bio,
            file_name=f"{leader_group}_情緒觀察每週心得報告_{reminder_target_week.split(' (')[0]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    with btn_col2:
        st.success("✅ 心得報告已生成！點擊左側按鈕即可下載 Word 檔。")


# ----------------- Tab 3: Data Management -----------------
with tab3:
    st.subheader("📊 填報紀錄資料庫 (最高權限管理員)")
    
    is_admin = st.checkbox("🔑 開啟完整資料庫檢視與匯出功能", value=True)
    
    if is_admin:
        df = load_data()
        
        if not df.empty:
            # Metrics
            total_records = len(df)
            total_weekly_sheets = df["當週填表張數"].sum()
            unique_members = df["個人姓名"].nunique()
            
            m1, m2, m3 = st.columns(3)
            with m1:
                st.metric("總回報筆數 (All Submissions)", f"{total_records} 筆")
            with m2:
                st.metric("當週累計填寫總張數", f"{total_weekly_sheets} 張")
            with m3:
                st.metric("全營不重複填報人數", f"{unique_members} 人")
            
            # Filters
            st.write("---")
            st.markdown("##### 🔍 系統全局數據篩選器")
            f_col1, f_col2 = st.columns(2)
            with f_col1:
                filter_group = st.multiselect("按組別篩選：", options=list(df["組別"].unique()))
            with f_col2:
                filter_week = st.multiselect("按回報期間篩選：", options=list(df["當週回報期間"].unique()))
                
            # Apply filters
            filtered_df = df.copy()
            if filter_group:
                filtered_df = filtered_df[filtered_df["組別"].isin(filter_group)]
            if filter_week:
                filtered_df = filtered_df[filtered_df["當週回報期間"].isin(filter_week)]
                
            st.write(f"📂 顯示篩選結果：共 {len(filtered_df)} 筆紀錄")
            st.dataframe(filtered_df, use_container_width=True)
            
            # Export data
            csv_data = filtered_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 匯出並下載全局 CSV 報表 (可用 Excel 直接開啟)",
                data=csv_data,
                file_name=f"all_groups_emotional_reports_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        else:
            st.warning("⚠️ 目前資料庫尚無任何填報紀錄。")
    else:
        st.info("請勾選「開啟完整資料庫檢視與匯出功能」以檢視數據庫。")
