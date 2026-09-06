import streamlit as st
import pandas as pd
import datetime
import os
import requests
import io
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Set page config for web app
st.set_page_config(
    page_title="情緒觀察營 — 申報與管理平台 (權限分流版)",
    page_icon="🌟",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    /* 全域清新水彩漸層背景與精緻字型 */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: #F6FAF4 !important;
        background-image: linear-gradient(135deg, #F3F8F2 0%, #FAF5EF 100%) !important;
    }
    
    /* 調整字型與主要內文顏色 */
    .stApp, p, span, label, li {
        font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif !important;
        color: #2E4436 !important; /* 深森林綠，確保手機上文字高度可讀 */
    }

    /* 網頁大標題：清新優雅水彩風 */
    .report-title {
        font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif;
        font-size: 2.3rem;
        font-weight: 800;
        color: #3C6E47; /* 溫和的水彩葉綠色 */
        text-align: center;
        margin-top: 15px;
        margin-bottom: 8px;
        text-shadow: 1px 1px 4px rgba(60, 110, 71, 0.15);
    }
    
    .report-subtitle {
        font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif;
        font-size: 1.1rem;
        color: #7A695A; /* 柔和的泥土暖褐 */
        text-align: center;
        margin-bottom: 30px;
        letter-spacing: 0.05em;
    }
    
    /* 提示區塊：淡淡的水彩綠葉底色 + 柔和左邊框 */
    .group-box {
        background-color: #EBF3EC !important;
        color: #2E5638 !important;
        padding: 16px;
        border-radius: 12px;
        border-left: 6px solid #8FA893 !important; /* 柔軟水彩綠 */
        box-shadow: 2px 4px 15px rgba(143, 168, 147, 0.12);
        margin-bottom: 25px;
        font-size: 1.05rem;
    }
    
    /* 提交成功區塊：粉嫩花瓣綠意感 */
    .success-box {
        background-color: #F0F6F2 !important;
        color: #2A4D3B !important;
        padding: 22px;
        border-radius: 14px;
        border: 1px solid #D5E5DE !important;
        box-shadow: 3px 6px 18px rgba(47, 82, 62, 0.08);
        margin-top: 15px;
    }
    
    /* LINE 關懷與郵件範本：微風暖沙水彩調 */
    .reminder-template {
        background-color: #FCF5EC !important;
        color: #6D5144 !important;
        padding: 18px;
        border-radius: 10px;
        border-left: 5px solid #E2A784 !important; /* 柔粉橘 */
        font-family: monospace;
        white-space: pre-wrap;
        margin-top: 10px;
        box-shadow: 1px 3px 10px rgba(226, 167, 132, 0.1);
    }
    
    .email-box {
        background-color: #F1F5F9 !important;
        color: #3B4B5E !important;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #8EA7BA !important; /* 水彩天藍 */
        font-family: monospace;
        white-space: pre-wrap;
        margin-top: 10px;
        box-shadow: 1px 3px 10px rgba(142, 167, 186, 0.1);
    }

    /* 調整 Streamlit 的按鈕，讓它符合水彩圓潤感 */
    div.stButton > button {
        background-color: #8FA893 !important; /* 水彩綠 */
        color: white !important;
        border-radius: 20px !important;
        border: none !important;
        box-shadow: 1px 3px 8px rgba(143, 168, 147, 0.25) !important;
        transition: all 0.3s ease !important;
        font-weight: bold !important;
        padding: 0.5rem 1.5rem !important;
    }
    div.stButton > button:hover {
        background-color: #76907B !important;
        box-shadow: 1px 4px 12px rgba(118, 144, 123, 0.35) !important;
        transform: translateY(-1px);
    }

    /* 頁面展開區塊 Header 圓角 */
    .streamlit-expanderHeader {
        background-color: #FAF5EF !important;
        border-radius: 10px !important;
        border: 1px solid #EAE0D5 !important;
    }

    /* 強制所有輸入框、文字區域、下拉選單、下拉清單外框使用白/淺色背景與深色文字，確保極高對比度且在手機或暗黑模式下清晰可讀 */
    input, textarea, select, 
    div[data-baseweb="select"] > div, 
    div[data-baseweb="textarea"] > div, 
    div[data-baseweb="input"] > div,
    [role="listbox"], [role="option"], 
    [data-baseweb="popover"] div, 
    [data-baseweb="popover"] ul, 
    [data-baseweb="popover"] li {
        background-color: #FFFFFF !important;
        color: #2E4436 !important;
        border-color: #D5E5DE !important;
    }
    /* 當焦點在輸入框或文字區時的文字顏色也必須保持深森林綠 */
    input:focus, textarea:focus {
        color: #2E4436 !important;
        background-color: #FFFFFF !important;
    }

    /* 確保側邊欄（Sidebar）在任何模式下都是溫暖的水彩粘土白，而非黑/深灰色 */
    section[data-testid="stSidebar"], [data-testid="stSidebarUserContent"], div[data-testid="stSidebarCollapseButton"] {
        background-color: #FAF5EF !important;
        background-image: linear-gradient(180deg, #F3F8F2 0%, #FAF5EF 100%) !important;
        color: #2E4436 !important;
    }
    section[data-testid="stSidebar"] p, 
    section[data-testid="stSidebar"] span, 
    section[data-testid="stSidebar"] label, 
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3 {
        color: #2E4436 !important;
    }

    /* 確保表格 Dataframe 在任何模式下皆為白色背景與深色文字 */
    div[data-testid="stTable"] table, div[data-testid="stDataFrame"] table {
        background-color: #FFFFFF !important;
        color: #2E4436 !important;
    }
    div[data-testid="stTable"] th, div[data-testid="stDataFrame"] th {
        background-color: #EBF3EC !important;
        color: #1B5E20 !important;
    }
    div[data-testid="stTable"] td, div[data-testid="stDataFrame"] td {
        background-color: #FFFFFF !important;
        color: #2E4436 !important;
    }

    /* 確保 Metric 數據指標卡片也有高對比度深色文字與白/淺色底色，在任何情況下都不會呈現黑底 */
    div[data-testid="stMetricValue"] > div {
        color: #1B5E20 !important; /* 清新綠數字 */
    }
    div[data-testid="stMetricLabel"] > div {
        color: #7A695A !important; /* 泥土腳標題 */
    }
    div[data-testid="stMetric"] {
        background-color: #FAF5EF !important;
        border-radius: 12px !important;
        padding: 12px !important;
        border: 1px solid #EAE0D5 !important;
        box-shadow: 1px 3px 10px rgba(122, 105, 90, 0.08) !important;
    }

</style>
""", unsafe_allow_html=True)

# Render Header
st.markdown('<div class="report-title">🌸 園區情緒觀察 週總結回報平台</div>', unsafe_allow_html=True)


# 1~18組完整分組名單 (來源自保溫分組名單)
# 1~18組完整分組名單 (嚴格依據 2026園區情緒觀察營-學員總名冊0805.pdf 整理，並在名字前方加上組內編號)
groups_data = {
    "第 1 組": {
        "leader": "曾旭正、李麗斐（副）",
        "members": [
            "1-1 曾旭正", "1-2 李麗斐", "1-3 蔡淵輝", "1-4 薛榕婷", "1-5 關婉玲", "1-6 周文祥", "1-7 馮豐隆", "1-8 郭擺灧", "1-9 鍾昆原", "1-10 張筱雯", "1-11 王坤德", "1-12 吳美毅", "1-13 曾美芳", "1-14 陳淑芬", "1-15 蔡沛妤", "1-16 黃淑清", "1-17 陳萬和", "1-18 梁庭", "1-19 嚴惠英"
        ]
    },
    "第 2 組": {
        "leader": "楊韻蓉",
        "members": [
            "2-1 林文華", "2-2 林佳慧", "2-3 楊韻蓉", "2-4 林翔均", "2-5 溫婷伊", "2-6 陳田恬", "2-7 廖袖婷", "2-8 王均晨", "2-9 蔡宜均", "2-10 吳煥崇", "2-11 黃曉鈺", "2-12 葉桂香", "2-13 陳秀花", "2-14 劉德政", "2-15 吳振仁"
        ]
    },
    "第 3 組": {
        "leader": "陳威廷",
        "members": [
            "3-1 陳威廷", "3-2 紀辰諭", "3-3 劉俊麟", "3-4 張又方", "3-5 謝耀德", "3-6 葉迷妮", "3-7 王韶怡", "3-8 郭芳佑", "3-9 王碧宏", "3-10 柯春黛", "3-11 黃玉宙", "3-12 何進德", "3-13 蕭名權", "3-14 戈德", "3-15 蘇献珍", "3-16 張顥嚴", "3-17 陳淑媛", "3-18 林秀婷", "3-19 柯春僖", "3-20 簡麗分", "3-21 賴彩鈴", "3-22 林耘庄", "3-23 黃如榛", "3-24 許媛婷", "3-25 陳薇莉", "3-26 張開明", "3-27 吳翊菱", "3-28 鍾亞昭", "3-29 林金泉"
        ]
    },
    "第 4 組": {
        "leader": "王國美",
        "members": [
            "4-1 王國美", "4-2 黃柏豪", "4-3 謝佳容", "4-4 劉以婕", "4-5 留千惠", "4-6 張玉芳", "4-7 吳念真", "4-8 陳芸詞", "4-9 陳揚旻", "4-10 鄭翊汝", "4-11 李濱如", "4-12 洪筱婷", "4-13 張麗卿", "4-14 陳聖昀", "4-15 曾麗珍", "4-16 賀正楨", "4-17 潘秋華", "4-18 鄭美華", "4-19 羅怡和", "4-20 周秀芬", "4-21 侯得正"
        ]
    },
    "第 5 組": {
        "leader": "郭宗川",
        "members": [
            "5-1 郭宗川", "5-2 劉馨檀", "5-3 黃瓊慧", "5-4 石俊傑", "5-5 張志豪", "5-6 林淑寧", "5-7 林麗雪", "5-8 陳毓旻", "5-9 楊舒涵", "5-10 李平裕", "5-11 李俊男", "5-12 尤躍勳", "5-13 戴世昌", "5-14 黃壬癸", "5-15 蔡宗佑", "5-16 邱淑春", "5-17 郭富貴", "5-18 李麗卿", "5-19 蕭玉卿", "5-20 陳秀美", "5-21 劉玉如", "5-22 賴幸絹"
        ]
    },
    "第 6 組": {
        "leader": "林家敬",
        "members": [
            "6-1 林家敬", "6-2 劉洹岑", "6-3 林德慶", "6-4 彭聖森", "6-5 粘智淳", "6-6 羅賢芳", "6-7 陳君妮", "6-8 徐立昇", "6-9 粘振清", "6-10 劉潓雪", "6-11 潘玟玲", "6-12 韓倩如", "6-13 黃乙娟", "6-14 廖鶴翔", "6-15 顏宗宏", "6-16 林獻崇", "6-17 張淑真", "6-18 許婉瑜", "6-19 劉淑美", "6-20 林俞彣", "6-21 許婷婷"
        ]
    },
    "第 7 組": {
        "leader": "陳定群",
        "members": [
            "7-1 陳定群", "7-2 吳庭妤", "7-3 吳幸蓉", "7-4 林伶雯", "7-5 林美芳", "7-6 林靖芳", "7-7 徐續仁", "7-8 張靜娟", "7-9 蕭淑勻", "7-10 吳璧合", "7-11 陳淑芳", "7-12 黃金輝", "7-13 鄭秀娟", "7-14 鄭宜涵", "7-15 鄭斐文", "7-16 謝萬蒲", "7-17 魏媛真", "7-18 楊美蘭"
        ]
    },
    "第 8 組": {
        "leader": "楊婉君",
        "members": [
            "8-1 楊婉君", "8-2 林宗平", "8-3 吳欣緯", "8-4 陳秋芬", "8-5 王怡誠", "8-6 王麗景", "8-7 章玲珠", "8-8 李瓊華", "8-9 劉燕冰", "8-10 盧菀萱", "8-11 李淑瑛", "8-12 楊素粉", "8-13 劉筱愛", "8-14 盧彥妤", "8-15 明子又", "8-16 黃翠月", "8-17 賴香雪", "8-18 蘇郁合", "8-19 蘇國樑"
        ]
    },
    "第 9 組": {
        "leader": "林如玉",
        "members": [
            "9-1 林如玉", "9-2 林淑娟", "9-3 余雅玲", "9-4 李慧瑛", "9-5 辛天送", "9-6 白秀清", "9-7 莊淑娟", "9-8 游建邦", "9-9 黃惠卿", "9-10 蔡純慧", "9-11 洪雅玲", "9-12 李宛融", "9-13 徐惠美", "9-14 周珈漩", "9-15 許湘惠", "9-16 周林美麗", "9-17 何苑色", "9-18 李俊宏", "9-19 喻榮華", "9-20 柳書玉", "9-21 蘇鴻濱"
        ]
    },
    "第 10 組": {
        "leader": "黃振育",
        "members": [
            "10-1 黃振育", "10-2 胡詠茹", "10-3 李宜晏", "10-4 楊崇誠", "10-5 陳世平", "10-6 張金爗", "10-7 陳麗華", "10-8 廖曉萍", "10-9 賴志雄", "10-10 吳東昱", "10-11 趙如薰", "10-12 叢思瑋", "10-13 李學璁", "10-14 黃品齊", "10-15 朱重信", "10-16 林明珠", "10-17 蔡春美", "10-18 楊昌學", "10-19 謝佩穎", "10-20 葉奕新"
        ]
    },
    "第 11 組": {
        "leader": "林美華",
        "members": [
            "11-1 賴素貞", "11-2 賴彥帛", "11-3 洪于晴", "11-4 何如甘", "11-5 李素茹", "11-6 林珍安", "11-7 吳核豫", "11-8 葉如玲", "11-9 林保秀", "11-10 陳品樺", "11-11 蘇相瑜", "11-12 江盈潔", "11-13 阮芬", "11-14 雙東華", "11-15 林美華", "11-16 蔡菁雯", "11-17 楊金珠", "11-18 蕭淑慧", "11-19 韋淑媛", "11-20 徐丞億"
        ]
    },
    "第 12 組": {
        "leader": "邱品瑄",
        "members": [
            "12-1 邱品瑄", "12-2 賴盈達", "12-3 陳雅惠", "12-4 江佳怡", "12-5 孫詩旻", "12-6 謝秀枝", "12-7 周素仰", "12-8 陳昭華", "12-9 簡達謙", "12-10 潘志滿", "12-11 童小鈴", "12-12 蔡宗涵", "12-13 蔡麗滿", "12-14 王冠蕙", "12-15 洪唯", "12-16 何淑琴", "12-17 顏妙如", "12-18 楊紫璿", "12-19 李淑芬", "12-20 洪金城"
        ]
    },
    "第 13 組": {
        "leader": "高韻筑",
        "members": [
            "13-1 陳重仰", "13-2 高韻筑", "13-3 王上苹", "13-4 陳珈暐", "13-5 林珮圻", "13-6 施瀞雅", "13-7 丁秌全", "13-8 黃蔚菁", "13-9 吳孟霖", "13-10 張慈方", "13-11 鄭鈞懿", "13-12 王慧玟", "13-13 彭美珠", "13-14 黃忠權", "13-15 陳彥伶", "13-16 卓卉綺", "13-17 林于煒", "13-18 張婉玲", "13-19 朱庭葦"
        ]
    },
    "第 14 組": {
        "leader": "蔡秉諺",
        "members": [
            "14-1 蔡秉諺", "14-2 高翊綺", "14-3 秦育曐", "14-4 莊媛婷", "14-5 蔡明真", "14-6 陳桂宸", "14-7 黃麗燕", "14-8 吳承璇", "14-9 葉美鴻", "14-10 張在增", "14-11 顏少萱", "14-12 王馨儀", "14-13 林奕瑄", "14-14 郭芷萱", "14-15 吳靜宜", "14-16 張容", "14-17 蕭米佋", "14-18 李岳璋", "14-19 廖育君", "14-20 謝孟君", "14-21 葉姿佑"
        ]
    },
    "第 15 組": {
        "leader": "高忠偉",
        "members": [
            "15-1 廖以萱", "15-2 高忠偉", "15-3 侯惠芸", "15-4 侯惠倫", "15-5 莊郁琳", "15-6 戴毓書", "15-7 張家瑄", "15-8 邵杰", "15-9 白豐銘", "15-10 黃沛瑋", "15-11 張雁菁", "15-12 蔡佩君", "15-13 鄭伃倢", "15-14 曾品綸", "15-15 張文馨", "15-16 劉思嫺", "15-17 黃姿蓉", "15-18 賴姳璇", "15-19 洪瑩蓁", "15-20 陳星妤"
        ]
    },
    "第 16 組": {
        "leader": "偕魁元",
        "members": [
            "16-1 偕魁元", "16-2 李庭安", "16-3 劉妙娟", "16-4 朱穆鳳", "16-5 沈承宗", "16-6 蔡雅雅", "16-7 凃文鳳", "16-8 莊雅惠", "16-9 劉巧領", "16-10 鄭雅文", "16-11 湯玉琦", "16-12 劉姿吟", "16-13 蕭茲方", "16-14 王芳珠", "16-15 陳冠同", "16-16 林秀菁", "16-17 黃政和", "16-18 徐友信", "16-19 吳莉惠", "16-20 林宜筠"
        ]
    },
    "第 17 組": {
        "leader": "蘇誌盈",
        "members": [
            "17-1 蘇誌盈", "17-2 王素敏", "17-3 蔡石福地", "17-4 李世林", "17-5 黃淑芳", "17-6 胡瑜玲", "17-7 柯孟宜", "17-8 林靜香", "17-9 葉瓊雅", "17-10 張秋美", "17-11 劉宇容", "17-12 黃克經", "17-13 林淑樺", "17-14 黃榮茂", "17-15 邵士誠", "17-16 鄭人豪", "17-17 姚翠華"
        ]
    },
    "第 18 組": {
        "leader": "陳麗華",
        "members": [
            "18-1 陳麗華", "18-2 廖述強", "18-3 黃淑梅", "18-4 林玲芬", "18-5 曹曼蓉", "18-6 柯錦鑫", "18-7 張湘瑜", "18-8 吳美蓉", "18-9 陳月美", "18-10 陳素葉", "18-11 郭宇家", "18-12 黃佩琪", "18-13 朱龍昌", "18-14 謝桂红", "18-15 阮善如", "18-16 陳丁清", "18-17 楊月娥", "18-18 謝秀錦", "18-19 陳慶旺", "18-20 陳錦鳳"
        ]
    }
}
# Fixed typos / mapping adjustments
groups_data["第 1 組"]["members"][7] = "1-8 郭瓈灧"
groups_data["第 11 組"]["members"][10] = "11-11 蘇相瑜"
groups_data["第 11 組"]["members"][13] = "11-14 許東華"


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
    "W1 ( 2026/08/31 ~ 2026/09/05 )",
    "W2 ( 2026/09/06 ~ 2026/09/12 )",
    "W3 ( 2026/09/13 ~ 2026/09/19 )",
    "W4 ( 2026/09/20 ~ 2026/09/26 )",
    "W5 ( 2026/09/27 ~ 2026/10/03 )",
    "W6 ( 2026/10/04 ~ 2026/10/10 )",
    "W7 ( 2026/10/11 ~ 2026/10/17 )",
    "W8 ( 2026/10/18 ~ 2026/10/24 )",
    "W9 ( 2026/10/25 ~ 2026/10/31 )",
    "W10 ( 2026/11/01 ~ 2026/11/07 )",
    "W11 ( 2026/11/08 ~ 2026/11/14 )",
    "W12 ( 2026/11/15 ~ 2026/11/21 )",
    "W13 ( 2026/11/22 ~ 2026/11/28 )",
    "W14 ( 2026/11/29 ~ 2026/12/05 )",
    "W15 ( 2026/12/06 ~ 2026/12/12 )",
    "W16 ( 2026/12/13 ~ 2026/12/19 )",
    "W17 ( 2026/12/20 ~ 2026/12/26 )",
    "W18 ( 2026/12/27 ~ 2027/01/02 )",
    "W19 ( 2027/01/03 ~ 2027/01/09 )",
    "W20 ( 2027/01/10 ~ 2027/01/16 )",
    "W21 ( 2027/01/17 ~ 2027/01/23 )",
    "W22 ( 2027/01/24 ~ 2027/01/30 )",
    "W23 ( 2027/01/31 ~ 2027/02/06 )",
    "W24 ( 2027/02/07 ~ 2027/02/13 )",
    "W25 ( 2027/02/14 ~ 2027/02/20 )",
    "W26 ( 2027/02/21 ~ 2027/02/27 )",
    "W27 ( 2027/02/28 ~ 2027/03/06 )",
    "W28 ( 2027/03/07 ~ 2027/03/13 )",
    "W29 ( 2027/03/14 ~ 2027/03/20 )",
    "W30 ( 2027/03/21 ~ 2027/03/27 )",
    "W31 ( 2027/03/28 ~ 2027/04/03 )",
    "W32 ( 2027/04/04 ~ 2027/04/10 )",
    "W33 ( 2027/04/11 ~ 2027/04/17 )",
    "W34 ( 2027/04/18 ~ 2027/04/24 )",
    "W35 ( 2027/04/25 ~ 2027/05/01 )",
    "W36 ( 2027/05/02 ~ 2027/05/08 )",
    "W37 ( 2027/05/09 ~ 2027/05/15 )",
    "W38 ( 2027/05/16 ~ 2027/05/22 )",
    "W39 ( 2027/05/23 ~ 2027/05/29 )",
    "W40 ( 2027/05/30 ~ 2027/06/05 )"
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
    df = load_data()
    new_df = pd.DataFrame([new_row])
    df = pd.concat([df, new_df], ignore_index=True)
    df.to_csv(DB_FILE, index=False, encoding='utf-8-sig')
    
    if gas_url and gas_url.strip():
        try:
            response = requests.post(gas_url.strip(), json=new_row, timeout=8)
            if response.status_code == 200:
                return True, "本機備份成功，且順利同步寫入 Google 雲端！"
            else:
                return True, f"本機備份成功，但雲端同步失敗 (狀態碼: {response.status_code})"
        except Exception as e:
            return True, f"本機備份成功，但雲端同步發生錯誤: {str(e)}"
    return True, "填報紀錄已成功儲存於本地伺服器。"

# Define passwords for group leaders (leader01 ~ leader18) and admin (admin2026)
leader_passwords = {f"第 {i} 組": f"leader{i:02d}" for i in range(1, 19)}
admin_password = "admin2026"

# Helper to generate beautiful Word Document (.docx) Weekly Thoughts Report
def generate_docx_report(group_name, leader, week, submissions, roster_list):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Microsoft JhengHei'
    font.size = Pt(11)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(f"📋 情緒觀察營 — {group_name} 每週心得報告")
    run_title.font.name = 'Microsoft JhengHei'
    run_title.font.size = Pt(18)
    run_title.bold = True
    run_title.font.color.rgb = RGBColor(27, 94, 32)
    
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_meta = p_meta.add_run(f"回報週別：{week}  |  保溫組長：{leader}  |  報表產出日期：{datetime.date.today().strftime('%Y/%m/%d')}")
    run_meta.font.size = Pt(10)
    run_meta.font.color.rgb = RGBColor(100, 110, 100)
    
    doc.add_paragraph()
    
    h_sum = doc.add_paragraph()
    run_h_sum = h_sum.add_run("📊 當週填表數據摘要表")
    run_h_sum.font.size = Pt(14)
    run_h_sum.bold = True
    run_h_sum.font.color.rgb = RGBColor(27, 94, 32)
    
    cols_data = [("個人編號", Inches(1.2)), ("組員姓名", Inches(1.2)), ("填報狀態", Inches(1.2)), ("當週張數", Inches(1.2)), ("累計張數", Inches(1.2))]
    table = doc.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    hdr_cells = table.rows[0].cells
    for i, (name, width) in enumerate(cols_data):
        hdr_cells[i].text = name
        hdr_cells[i].width = width
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.name = 'Microsoft JhengHei'
            r.font.size = Pt(10.5)
            
        tcPr = hdr_cells[i]._tc.get_or_add_tcPr()
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), '1B5E20')
        shading.set(qn('w:val'), 'clear')
        tcPr.append(shading)
        
    sub_map = {row["個人姓名"]: row for _, row in submissions.iterrows()}
    
    total_sheets = 0
    submitted_count = 0
    
    for row in roster_list:
        m_name = row["個人姓名"]
        m_id = row["組員號碼"]
        
        row_cells = table.add_row().cells
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
                    r.font.color.rgb = RGBColor(183, 28, 28)
                elif status == "✅ 已填" and i == 2:
                    r.font.color.rgb = RGBColor(46, 125, 50)
                    
    p_metrics = doc.add_paragraph()
    p_metrics.paragraph_format.space_before = Pt(12)
    p_metrics.paragraph_format.space_after = Pt(24)
    r_met = p_metrics.add_run(f"💡 本週小組統計：應填報 {len(roster_list)} 人，已填報 {submitted_count} 人，未填報 {len(roster_list) - submitted_count} 人。當週填表總張數：{total_sheets} 張。")
    r_met.font.size = Pt(10)
    r_met.font.bold = True
    r_met.font.color.rgb = RGBColor(100, 100, 100)
    
    h_detail = doc.add_paragraph()
    run_h_detail = h_detail.add_run("💬 組員心得與對法師提問明細")
    run_h_detail.font.size = Pt(14)
    run_h_detail.bold = True
    run_h_detail.font.color.rgb = RGBColor(27, 94, 32)
    h_detail.paragraph_format.space_after = Pt(12)
    
    for row in roster_list:
        m_name = row["個人姓名"]
        m_id = row["組員號碼"]
        
        p_mem = doc.add_paragraph()
        p_mem.paragraph_format.space_before = Pt(8)
        p_mem.paragraph_format.space_after = Pt(2)
        r_mem = p_mem.add_run(f"【{m_id}】 {m_name}")
        r_mem.bold = True
        r_mem.font.size = Pt(11.5)
        r_mem.font.color.rgb = RGBColor(78, 52, 46)
        
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
            
        p_div = doc.add_paragraph()
        p_div.paragraph_format.space_before = Pt(4)
        p_div.paragraph_format.space_after = Pt(4)
        r_div = p_div.add_run("―" * 50)
        r_div.font.color.rgb = RGBColor(230, 230, 230)
        r_div.font.size = Pt(8)
        
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# Initialize session state variables
if "is_authenticated" not in st.session_state:
    st.session_state["is_authenticated"] = False
if "role" not in st.session_state:
    st.session_state["role"] = "📝 一般學員填寫問卷"
if "logged_in_group" not in st.session_state:
    st.session_state["logged_in_group"] = None
if "submitted_successfully" not in st.session_state:
    st.session_state["submitted_successfully"] = False

# Sidebar Configuration (only visible when logged in)
if st.session_state["is_authenticated"]:
    with st.sidebar:
        st.image("https://img.icons8.com/clouds/200/database.png", width=80)
        st.markdown(f"### 👤 目前登入身分\n**{st.session_state['role']}**")
        if st.session_state["logged_in_group"]:
            st.markdown(f"**管理組別：{st.session_state['logged_in_group']}**")
            
        st.write("---")
        st.markdown("### ☁️ 雲端同步設定")
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
        
        st.write("---")
        # Logout button
        if st.button("🔒 登出管理系統", type="primary", key="sidebar_logout_btn"):
            st.session_state["is_authenticated"] = False
            st.session_state["role"] = "📝 一般學員填寫問卷"
            st.session_state["logged_in_group"] = None
            st.session_state["submitted_successfully"] = False
            # clear query params
            st.query_params.clear()
            st.rerun()
else:
    # If not authenticated, hide sidebar completely!
    st.markdown("""
    <style>
        [data-testid="collapsedControl"] {
            display: none !important;
        }
        section[data-testid="stSidebar"] {
            display: none !important;
        }
    
    /* 強制所有輸入框、文字區域、下拉選單、下拉清單外框使用白/淺色背景與深色文字，確保極高對比度且在手機或暗黑模式下清晰可讀 */
    input, textarea, select, 
    div[data-baseweb="select"] > div, 
    div[data-baseweb="textarea"] > div, 
    div[data-baseweb="input"] > div,
    [role="listbox"], [role="option"], 
    [data-baseweb="popover"] div, 
    [data-baseweb="popover"] ul, 
    [data-baseweb="popover"] li {
        background-color: #FFFFFF !important;
        color: #2E4436 !important;
        border-color: #D5E5DE !important;
    }
    /* 當焦點在輸入框或文字區時的文字顏色也必須保持深森林綠 */
    input:focus, textarea:focus {
        color: #2E4436 !important;
        background-color: #FFFFFF !important;
    }

    /* 確保側邊欄（Sidebar）在任何模式下都是溫暖的水彩粘土白，而非黑/深灰色 */
    section[data-testid="stSidebar"], [data-testid="stSidebarUserContent"], div[data-testid="stSidebarCollapseButton"] {
        background-color: #FAF5EF !important;
        background-image: linear-gradient(180deg, #F3F8F2 0%, #FAF5EF 100%) !important;
        color: #2E4436 !important;
    }
    section[data-testid="stSidebar"] p, 
    section[data-testid="stSidebar"] span, 
    section[data-testid="stSidebar"] label, 
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3 {
        color: #2E4436 !important;
    }

    /* 確保表格 Dataframe 在任何模式下皆為白色背景與深色文字 */
    div[data-testid="stTable"] table, div[data-testid="stDataFrame"] table {
        background-color: #FFFFFF !important;
        color: #2E4436 !important;
    }
    div[data-testid="stTable"] th, div[data-testid="stDataFrame"] th {
        background-color: #EBF3EC !important;
        color: #1B5E20 !important;
    }
    div[data-testid="stTable"] td, div[data-testid="stDataFrame"] td {
        background-color: #FFFFFF !important;
        color: #2E4436 !important;
    }

    /* 確保 Metric 數據指標卡片也有高對比度深色文字與白/淺色底色，在任何情況下都不會呈現黑底 */
    div[data-testid="stMetricValue"] > div {
        color: #1B5E20 !important; /* 清新綠數字 */
    }
    div[data-testid="stMetricLabel"] > div {
        color: #7A695A !important; /* 泥土腳標題 */
    }
    div[data-testid="stMetric"] {
        background-color: #FAF5EF !important;
        border-radius: 12px !important;
        padding: 12px !important;
        border: 1px solid #EAE0D5 !important;
        box-shadow: 1px 3px 10px rgba(122, 105, 90, 0.08) !important;
    }

</style>
    """, unsafe_allow_html=True)
    gas_api_url = st.session_state.get("gas_url", "")

# ----------------- Check Query Params for Admin Login -----------------
query_role = st.query_params.get("role", "")
query_manage = st.query_params.get("manage", "")

show_admin_login_screen = False
if (query_role in ["manage", "leader", "admin"] or query_manage == "true") and not st.session_state["is_authenticated"]:
    show_admin_login_screen = True

# Get current state variables for standard compatibility
is_authenticated = st.session_state["is_authenticated"]
role = st.session_state["role"]
logged_in_group = st.session_state["logged_in_group"]

# ----------------- Router Logic -----------------
if show_admin_login_screen:
    # 1. Dedicated separate Admin/Leader login view
    st.markdown("<div style='text-align: center; margin-top: 30px;'><img src='https://img.icons8.com/clouds/200/lock.png' width='120'></div>", unsafe_allow_html=True)
    st.subheader("🔐 情緒觀察營 — 管理者專屬登入頁面")
    st.info("💡 請輸入您的組長密碼（例如 leader10）或總負責人密碼（例如 admin2026）解鎖資料閱讀。")
    
    login_pwd = st.text_input("請輸入管理密碼：", type="password", key="dedicated_pwd_input")
    
    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        login_click = st.button("確認登入 🔓", type="primary", key="dedicated_login_btn")
    with col_btn2:
        back_to_form = st.button("返回學員填報頁面 📝", key="dedicated_back_btn")
        if back_to_form:
            st.query_params.clear()
            st.rerun()
            
    if login_click and login_pwd:
        if login_pwd == admin_password:
            st.session_state["is_authenticated"] = True
            st.session_state["role"] = "📊 總窗口後台管理"
            st.session_state["logged_in_group"] = None
            st.success("🔓 總窗口解鎖成功！您擁有全局最高權限。")
            st.query_params.clear()
            st.rerun()
        else:
            matched_group = None
            for g, p in leader_passwords.items():
                if login_pwd == p:
                    matched_group = g
                    break
            if matched_group:
                st.session_state["is_authenticated"] = True
                st.session_state["role"] = "👥 保溫組長專區"
                st.session_state["logged_in_group"] = matched_group
                st.success(f"🔓 登入成功！已解鎖【{matched_group}】權限。")
                st.query_params.clear()
                st.rerun()
            else:
                st.error("❌ 密碼錯誤，請重新輸入！ (提示：leader01 ~ leader18)")

elif not is_authenticated:
    # 2. Student Questionnaire View (Ultra-clean, No st.form, Zero mobile confirmation alerts)
    st.info("💡 敬請於亞洲時間每週六 23:00 前完成填寫並回傳，感謝您的配合。統計區間為「前週日至本週六」之合計填表張數。")
    
    st.subheader("第一步：選擇您的組別與姓名")
    col1, col2 = st.columns(2)
    with col1:
        group_selected = st.selectbox("請選擇您的組別 *", list(groups_data.keys()), index=9) # Default to Group 10
    
    with col2:
        names_list = groups_data[group_selected]["members"]
        name_selected = st.selectbox("請選擇您的姓名 *", names_list)
        
    leader = groups_data[group_selected]["leader"]
    
    group_num = group_selected.replace("第 ", "").replace(" 組", "")
    if name_selected and " " in name_selected:
        parts = name_selected.split(" ")
        auto_member_id = parts[0]
        actual_name = parts[1]
    else:
        auto_member_id = f"{group_num}-X"
        actual_name = name_selected
        
    st.markdown(f'<div class="group-box">ℹ️ 您選取的是 <b>{group_selected}</b>（保溫組長：{leader}） | 請確認已選擇編號為：<b>{auto_member_id}</b></div>', unsafe_allow_html=True)
    
    if st.session_state.get("submitted_successfully", False):
        st.markdown(f"""
        <div class="success-box">
            <h3 style="margin-top: 0; color: #0F5132;">🎉 提交成功！</h3>
            <p>感謝 <b>{st.session_state.get('last_submitted_name', '')}</b> 的回報！您的填表紀錄已妥善儲存。</p>
            <p>💡 {st.session_state.get('last_submitted_msg', '')}</p>
            <ul>
                <li><b>個人編號：</b>{st.session_state.get('last_submitted_id', '')}</li>
                <li><b>回報期間：</b>{st.session_state.get('last_submitted_period', '')}</li>
                <li><b>當週張數：</b>{st.session_state.get('last_submitted_weekly', 0)} 張</li>
                <li><b>累計張數：</b>{st.session_state.get('last_submitted_cum', 0)} 張</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("再填一筆 📝", type="primary", key="reset_submission_btn"):
            st.session_state["submitted_successfully"] = False
            st.rerun()
    else:
        st.subheader("第二步：填寫當週回報資訊")
        period = st.selectbox("請確認當週回報期間 *", weeks_list, index=0)
        
        st.write("---")
        st.markdown("#### 📝 請回答以下兩道題目：")
        
        weekly_sheets = st.number_input("【題目一】您本週（當週）合計填寫了幾張情緒觀察表？ *", min_value=0, max_value=100, value=0, step=1, key="student_weekly_sheets")
        cum_sheets = st.number_input("【題目二】您自 2026/8/17 起累計填寫了幾張情緒觀察表（含本週）？ *", min_value=0, max_value=9999, value=0, step=1, key="student_cum_sheets")
        
        st.write("---")
        feedback = st.text_area("✍️ 心得或對法師提問（選填）：", placeholder="請輸入您當週的觀察心得，或是想對法師提出的疑問...", key="student_feedback")
        
        submit_btn = st.button("提交問卷 📨", type="primary", key="student_submit_btn")
        
        if submit_btn:
            email = f"{actual_name}@example.com"
            female_names = [
                "謝佩穎", "謝佳容", "謝孟君", "蔡佩君", "許婷婷", "張靜娟", "李慧瑛", 
                "張慈方", "劉姿吟", "胡詠茹", "李宜晏", "陳麗華", "廖曉萍", "趙如薫", 
                "叢思瑋", "蔡春美", "曾美芳", "張筱雯", "關婉玲", "曾芳美", "吳美毅",                 "蔡沛妤", "李麗斐", "黃淑清", "梁庭", "嚴惠英", "陳淑芬", "郭瓈灧",                 "黃曉鈺", "陳秀花", "蔡宜均", "林佳慧", "廖袖婷", "葉桂香", "溫婷伊",                 "陳田恬", "張又方", "葉迷妮", "王韶怡", "柯春黛", "陳淑媛", "陳薇莉",                 "吳翊菱", "簡麗分", "蘇獻珍", "林秀婷", "柯春僖", "賴彩鈴", "林耘庄",                 "黃如榛", "許媛婷", "鄭翊汝", "李濱如", "洪筱婷", "劉以婕", "陳芸詞",                 "張麗卿", "潘秋華", "留千惠", "張玉芳", "曾麗珍", "鄭美華", "周秀芬", "鍾亞昭"
            ]
            gender = "女" if actual_name in female_names else "男"
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            new_row = {
                "時間戳記": timestamp,
                "組別": group_selected,
                "保溫組長": leader,
                "個人姓名": actual_name,
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
                st.session_state["submitted_successfully"] = True
                st.session_state["last_submitted_name"] = actual_name
                st.session_state["last_submitted_msg"] = msg
                st.session_state["last_submitted_id"] = auto_member_id
                st.session_state["last_submitted_period"] = period
                st.session_state["last_submitted_weekly"] = int(weekly_sheets)
                st.session_state["last_submitted_cum"] = int(cum_sheets)
                st.rerun()

    # Footer login expander for group leaders / admin
    st.write("---")
    with st.expander("🔑 關懷管理登入 (組長 / 總負責人專用)"):
        st.write("請輸入您的管理密碼解鎖權限：")
        login_pwd = st.text_input("輸入登入密碼", type="password", key="footer_pwd_input")
        if st.button("確認登入 🔓", key="footer_login_submit"):
            if login_pwd == admin_password:
                st.session_state["is_authenticated"] = True
                st.session_state["role"] = "📊 總窗口後台管理"
                st.session_state["logged_in_group"] = None
                st.success("🔓 總窗口解鎖成功！您擁有全局最高權限。")
                st.rerun()
            else:
                matched_group = None
                for g, p in leader_passwords.items():
                    if login_pwd == p:
                        matched_group = g
                        break
                if matched_group:
                    st.session_state["is_authenticated"] = True
                    st.session_state["role"] = "👥 保溫組長專區"
                    st.session_state["logged_in_group"] = matched_group
                    st.success(f"🔓 登入成功！您已解鎖為【{matched_group}】權限。")
                    st.rerun()
                else:
                    st.error("❌ 密碼錯誤，請重新輸入！ (提示：leader01 ~ leader18)")

elif role == "👥 保溫組長專區":
    if not is_authenticated:
        st.info("👈 請在左側控制面板輸入「組長密碼」解鎖您的保溫組別權限。 (預設為 leader01 至 leader18)")
        st.image("https://img.icons8.com/clouds/400/lock.png", width=250)
    else:
        st.subheader(f"👥 【{logged_in_group}】保溫組長一鍵關懷與心得導出專區")
        st.caption("🔒 系統已安全限制：您目前僅可查閱貴組的填報資料，無法查看其他組別。")
        
        # 1. Primary care selection - select EXACTLY ONE target week
        st.markdown("### 📅 第一步：選擇您要查詢與關懷的「目標週別」")
        reminder_target_week = st.selectbox("請選擇目標週別 *", weeks_list, index=0, key="leader_target_week")
        
        target_wk_label = reminder_target_week.split(' (')[0] # e.g. W1
        
        # Load data and build roster list
        all_data = load_data()
        group_roster = groups_data[logged_in_group]["members"]
        g_num = logged_in_group.replace("第 ", "").replace(" 組", "")
        
        roster_list = []
        for idx, m_full in enumerate(group_roster):
            if " " in m_full:
                parts = m_full.split(" ")
                m_id = parts[0]
                m_name = parts[1]
            else:
                m_id = f"{g_num}-{idx+1}"
                m_name = m_full
            roster_list.append({
                "組員號碼": m_id,
                "個人姓名": m_name,
                "編號排序": idx+1
            })
        roster_base_df = pd.DataFrame(roster_list)
        
        # Submissions for selected target week in logged in group
        target_wk_submissions = all_data[(all_data["組別"] == logged_in_group) & (all_data["當週回報期間"] == reminder_target_week)]
        submitted_names = target_wk_submissions["個人姓名"].tolist()
        
        # Unsubmitted members
        unsubmitted_members = []
        for row in roster_list:
            if row["個人姓名"] not in submitted_names:
                unsubmitted_members.append(row)
                
        # Simple statistics card
        total_group_m = len(roster_list)
        sub_count = len(submitted_names)
        unsub_count = len(unsubmitted_members)
        sub_rate = (sub_count / total_group_m) * 100 if total_group_m > 0 else 0
        
        st.write("---")
        st.markdown("### 📊 本週填報進度統計")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            st.metric("應填報人數", f"{total_group_m} 人")
        with col_m2:
            st.metric("已完成登記", f"{sub_count} 人")
        with col_m3:
            st.metric("尚未登記人數", f"{unsub_count} 人", delta=f"-{unsub_count}" if unsub_count > 0 else None, delta_color="inverse")
        with col_m4:
            st.metric("當週填報率", f"{sub_rate:.1f}%")
            
        st.write("---")
        
        # 2. Roster Matrix Expander (Hidden by default for simplification!)
        with st.expander("📅 展開歷次週別對照矩陣 (跨週對照表)"):
            st.markdown("您可以選擇自訂的回報週別範圍，以查看多週的對照狀態：")
            start_week = st.selectbox("開始週別：", weeks_list, index=0, key="matrix_start_wk")
            end_week = st.selectbox("結束週別：", weeks_list, index=0, key="matrix_end_wk")
            
            try:
                start_idx = weeks_list.index(start_week)
                end_idx = weeks_list.index(end_week)
                if start_idx <= end_idx:
                    selected_weeks_range = weeks_list[start_idx:end_idx+1]
                else:
                    selected_weeks_range = weeks_list[end_idx:start_idx+1]
            except Exception:
                selected_weeks_range = weeks_list[0:4]
                
            matrix_data = roster_base_df.copy()
            for wk in selected_weeks_range:
                wk_lbl = wk.split(" (")[0]
                wk_sub_names = all_data[(all_data["組別"] == logged_in_group) & (all_data["當週回報期間"] == wk)]
                status_col = []
                for m_full in group_roster:
                    m_name = m_full.split(" ")[1] if " " in m_full else m_full
                    user_record = wk_sub_names[wk_sub_names["個人姓名"] == m_name]
                    if not user_record.empty:
                        sheets = user_record["當週填表張數"].values[0]
                        status_col.append(f"✅ 已填 ({int(sheets)}張)")
                    else:
                        status_col.append("❌ 未填")
                matrix_data[wk_lbl] = status_col
            
            display_matrix = matrix_data.drop(columns=["編號排序"]).set_index("組員號碼")
            st.dataframe(display_matrix, use_container_width=True)
            
        st.write("---")
        
        # 3. Unsubmitted care list & LINE template
        st.markdown(f"### 🚨 {target_wk_label} 未填寫關懷專區")
        if unsub_count == 0:
            st.balloons()
            st.success(f"🎉 太棒了！【{logged_in_group}】在 {target_wk_label} 的全體組員皆已完成填報問卷！")
        else:
            col_u1, col_u2 = st.columns([1, 1])
            with col_u1:
                st.warning(f"⚠️ 尚有 **{unsub_count}** 位同修尚未完成本週填寫：")
                unsub_df = pd.DataFrame(unsubmitted_members).drop(columns=["編號排序"])
                st.dataframe(unsub_df, use_container_width=True, hide_index=True)
            with col_u2:
                st.markdown("💬 **LINE 溫馨關懷提醒範本：**")
                unsub_mentions = " ".join([f"@{m['個人姓名']}" for m in unsubmitted_members])
                template_text = f"""📢【情緒觀察表 — 每週填報溫馨提醒】

各位【{logged_in_group}】的同修師兄姐吉祥：

感恩大家一週以來精進修行與自我觀察。🙏
目前本週【{reminder_target_week}】之情緒觀察填表統計正在進行中。

若您已完成填表登記，非常感恩您的配合！
若本週尚未完成，請撥冗 1 分鐘點擊下方平台連結回報，感恩合十。❤️

📝 本週尚未完成登記的同修（請抽空上網補登喔）：
👉 {unsub_mentions}

（若您剛剛已提交，系統同步可能有些許延遲，請忽略此訊息。再次感恩大家一同增長智慧與慈悲！🌱）"""
                st.text_area("可以直接複製以下文字發到 LINE 群組中：", value=template_text, height=220, key="line_reminder_text_area")

        st.write("---")
        
        # 4. Detailed Data for selected target week
        st.markdown(f"### 📋 {target_wk_label} 組員填報心得與提問清單")
        st.caption("依據組員號碼順序排列，方便閱讀心得與對法師的提問")
        
        leader_df = target_wk_submissions.copy()
        if not leader_df.empty:
            def get_sort_key(id_val):
                try:
                    parts = str(id_val).split("-")
                    return int(parts[-1])
                except Exception:
                    return 999
            leader_df["_sort_key"] = leader_df["個人編號"].apply(get_sort_key)
            leader_df = leader_df.sort_values(by=["_sort_key"]).drop(columns=["_sort_key"])
            st.dataframe(
                leader_df[["個人編號", "個人姓名", "當週填表張數", "累積填表張數", "心得或對法師提問", "時間戳記"]],
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info(f"💡 目前在 {target_wk_label} 期間內，本小組尚無組員填報資料。")
            
        st.write("---")
        
        # 5. One-Click docx generator for target week
        st.markdown(f"### 📝 一鍵生成 {target_wk_label} 心得報告 (.docx)")
        doc_bio = generate_docx_report(
            group_name=logged_in_group,
            leader=groups_data[logged_in_group]["leader"],
            week=reminder_target_week,
            submissions=target_wk_submissions,
            roster_list=roster_list
        )
        btn_col1, btn_col2 = st.columns([1, 2])
        with btn_col1:
            st.download_button(
                label="📥 一鍵導出 Word 格式心得報告",
                data=doc_bio,
                file_name=f"{logged_in_group}_情緒觀察每週心得報告_{target_wk_label}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_docx_report_btn"
            )
        with btn_col2:
            st.success("✅ 當週心得報告已生成！點擊左側按鈕即可下載 Word 檔。")



elif role == "📊 總窗口後台管理":
    if not is_authenticated:
        st.info("👈 請在左側控制面板輸入「總窗口管理密碼」解鎖最高後台管理權限。")
        st.image("https://img.icons8.com/clouds/400/security-shield.png", width=250)
    else:
        st.subheader("📊 總窗口全局資料管理庫")
        
        # Load database
        df = load_data()
        
        # 1. Weekly Email status reporter (Monday Reminder)
        st.markdown("---")
        st.markdown("### 📧 【週一專用】一鍵生成 18 組長催繳通報郵件")
        st.write("每週一總窗口可以一鍵檢視 18 小組的完整進度，並能自動產生 18 封針對各組組長的通知信件。")
        
        col_em1, col_em2 = st.columns([1, 2])
        
        with col_em1:
            target_report_week = st.selectbox("請選擇您要彙整的「目標週別」 *", weeks_list, index=0)
            
            # Simple aggregation analysis
            st.write("---")
            st.markdown("**全營 18 組填寫率快速評估**")
            
            if not df.empty:
                # Submissions in selected week
                wk_sub = df[df["當週回報期間"] == target_report_week]
                
                rows_stats = []
                for g_name, g_info in groups_data.items():
                    m_list_full = g_info["members"]
                    m_list_names = [m.split(" ")[1] if " " in m else m for m in m_list_full]
                    ldr = g_info["leader"]
                    sub_names = wk_sub[wk_sub["組別"] == g_name]["個人姓名"].tolist()
                    sub_count = len(set(sub_names) & set(m_list_names))
                    unsub_count = len(m_list_full) - sub_count
                    rate = (sub_count / len(m_list_full)) * 100 if len(m_list_full) > 0 else 0
                    rows_stats.append({
                        "組別": g_name,
                        "保溫組長": ldr,
                        "應填人數": len(m_list_full),
                        "已填人數": sub_count,
                        "未填人數": unsub_count,
                        "填報率": f"{rate:.1f}%",
                        "unsub_names": [m for m in m_list_full if (m.split(" ")[1] if " " in m else m) not in sub_names],
                        "sub_names": [m for m in m_list_full if (m.split(" ")[1] if " " in m else m) in sub_names]
                    })
                
                stats_df = pd.DataFrame(rows_stats)
                st.dataframe(stats_df[["組別", "保溫組長", "已填人數", "未填人數", "填報率"]], hide_index=True, use_container_width=True)
            else:
                st.warning("⚠️ 尚無任何數據，無法生成催繳分析。")
                
        with col_em2:
            st.markdown("📬 **組長催繳郵件/訊息預覽與複製**")
            if not df.empty and 'rows_stats' in locals():
                gp_select_em = st.selectbox("選擇要發送的組別：", [r["組別"] for r in rows_stats])
                
                # Find matching group row
                matched_row = next(r for r in rows_stats if r["組別"] == gp_select_em)
                
                g_leader = matched_row["保溫組長"]
                m_total = matched_row["應填人數"]
                m_sub = matched_row["已填人數"]
                m_unsub = matched_row["未填人數"]
                r_rate = matched_row["填報率"]
                u_list = matched_row["unsub_names"]
                s_list = matched_row["sub_names"]
                
                unsub_str = "、".join(u_list) if u_list else "無（皆已填寫！）"
                sub_str = "、".join(s_list) if s_list else "無"
                
                email_subject = f"【情緒觀察營】{target_report_week.split(' (')[0]} 週總結填報狀況通報 — {gp_select_em}"
                email_body = f"""親愛的【{gp_select_em}】保溫組長 {g_leader} 吉祥：

以下是截至每週一為止，貴組在【{target_report_week.split(' (')[0]}】情緒觀察表填報統計的最新進度通報：

📊 填報數據統計：
- 應填寫人數：{m_total} 人
- 已完成人數：{m_sub} 人
- 未完成人數：{m_unsub} 人
- 當前填報率：{r_rate}

📝 已完成登記的同修：
👉 {sub_str}

🚨 尚未完成登記的同修：
👉 {unsub_str}

【一鍵關懷溫馨提醒】：
若貴組尚有組員未完成，再麻煩您點選下方平台連結登入組長專區，即可直接「一鍵複製 LINE 催繳文字」到您的組別群組進行溫馨關懷。

情緒觀察營管理平台連結：(請填寫您部署後的 Streamlit 連結)

感恩您的發心護持與關懷付出，祝願您福慧雙修、六時吉祥！🙏

情緒觀察大組總窗口 敬上"""
                
                st.markdown(f"**📩 郵件主旨：**")
                st.code(email_subject, language="text")
                st.markdown(f"**✉️ 郵件正文：**")
                st.text_area("點擊右上方複製按鈕，直接貼入您的 Outlook / Gmail 寄出：", value=email_body, height=330)
                
                # Consolidate download
                all_emails_text = ""
                for r in rows_stats:
                    e_sub = f"主旨：【情緒觀察營】{target_report_week.split(' (')[0]} 週總結填報狀況通報 — {r['組別']}"
                    u_str = "、".join(r["unsub_names"]) if r["unsub_names"] else "無（全體已填寫！）"
                    s_str = "、".join(r["sub_names"]) if r["sub_names"] else "無"
                    
                    e_body = f"""{e_sub}
                    
親愛的【{r['組別']}】保溫組長 {r['保溫組長']} 吉祥：

以下是截至每週一為止，貴組在【{target_report_week.split(' (')[0]}】情緒觀察表填報統計的最新進度通報：

📊 填報數據統計：
- 應填寫人數：{r['應填人數']} 人
- 已完成人數：{r['已填人數']} 人
- 未完成人數：{r['未填人數']} 人
- 當前填報率：{r['填報率']}

📝 已完成登記的同修：
👉 {s_str}

🚨 尚未完成登記的同修：
👉 {u_str}

【一鍵關懷溫馨提醒】：
再麻煩您協助撥冗催促未填寫之組員。感恩您的發心與承擔！🙏

情緒觀察大組總窗口 敬上
========================================================================
"""
                    all_emails_text += e_body
                    
                st.download_button(
                    label="📥 一鍵導出 18 組「所有組長催繳信件」純文字檔",
                    data=all_emails_text,
                    file_name=f"all_groups_monday_emails_{target_report_week.split(' (')[0]}.txt",
                    mime="text/plain"
                )
        
        # 2. Database View
        st.markdown("---")
        st.markdown("### 📁 完整資料庫檢視與全局數據篩選器")
        
        if not df.empty:
            total_records = len(df)
            total_weekly_sheets = df["當週填表張數"].sum()
            unique_members = df["個人姓名"].nunique()
            
            m1, m2, m3 = st.columns(3)
            with m1:
                st.metric("營隊總回報筆數", f"{total_records} 筆")
            with m2:
                st.metric("累計當週總填表張數", f"{total_weekly_sheets} 張")
            with m3:
                st.metric("全營不重複填報人數", f"{unique_members} 人")
            
            st.write("---")
            st.markdown("##### 🔍 系統全局數據篩選器")
            f_col1, f_col2 = st.columns(2)
            with f_col1:
                filter_group = st.multiselect("按組別篩選：", options=list(df["組別"].unique()))
            with f_col2:
                filter_week = st.multiselect("按回報期間篩選：", options=list(df["當週回報期間"].unique()))
                
            filtered_df = df.copy()
            if filter_group:
                filtered_df = filtered_df[filtered_df["組別"].isin(filter_group)]
            if filter_week:
                filtered_df = filtered_df[filtered_df["當週回報期間"].isin(filter_week)]
                
            st.write(f"📂 顯示篩選結果：共 {len(filtered_df)} 筆紀錄")
            st.dataframe(filtered_df, use_container_width=True)
            
            csv_data = filtered_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 匯出並下載全局 CSV 報表 (可用 Excel 直接開啟)",
                data=csv_data,
                file_name=f"all_groups_emotional_reports_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        else:
            st.warning("⚠️ 目前資料庫尚無任何填報紀錄。")
