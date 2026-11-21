import difflib
import gc
import io
import json
import math
import os
import random
import re
import shutil
import threading
import time
from typing import Any, Dict, List, Tuple
from google import genai
from google.genai import errors, types
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageOps
import pillow_heif
import streamlit as st

try:
    import pptx
    from pptx.util import Inches, Pt, Emu
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

# ------------------------------------------------------------------
# 📌 註冊 HEIC / HEIF 圖片格式解碼器 (支援 iPhone 相片格式)
# ------------------------------------------------------------------
pillow_heif.register_heif_opener()

# ------------------------------------------------------------------
# 📌 全局設定與目錄建立
# ------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TASKS_DIR = os.path.join(BASE_DIR, "tasks_storage")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
STAGING_BASE_DIR = os.path.join(TASKS_DIR, "_staging_pool")

os.makedirs(TASKS_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs(STAGING_BASE_DIR, exist_ok=True)

# 嚴選具備超高免費配額 (RPD/RPM) 與超低延遲的極速模型 (徹底排除易觸發 20 RPD 限額之模型)
SAFE_MODELS = [
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
    "gemini-flash-lite-latest",
]
DEFAULT_MODEL = "gemini-3.5-flash-lite"

# 頁面配置
st.set_page_config(
    page_title="廠區動火管制表手寫 AI 辨識與船段圖自動生成系統",
    page_icon="🔥",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🔥 造船廠動火管制表手寫 AI 辨識與船段圖自動生成系統")
st.caption(
    "⚡ 支援手機分批累加上傳、伺服器極速辨識 (14張相片約10秒完成) 與 **J418/J419 船段圖自動標示**"
)

# ------------------------------------------------------------------
# 📌 AI 自動學習記憶系統 (手動編輯後持久化記憶，供後續辨識與校正優先使用)
# ------------------------------------------------------------------
LEARNED_MEMORY_FILE = os.path.join(TASKS_DIR, "learned_memory.json")


def load_learned_memory() -> dict:
    if os.path.exists(LEARNED_MEMORY_FILE):
        try:
            with open(LEARNED_MEMORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "custom_vendors": [],
        "vendor_aliases": {},
        "worker_roster": [],
        "worker_aliases": {},
        "watcher_roster": [],
        "watcher_aliases": {},
        "location_aliases": {},
        "corrections_count": 0,
    }


def save_learned_memory(mem: dict):
    try:
        with open(LEARNED_MEMORY_FILE, "w", encoding="utf-8") as f:
            json.dump(mem, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def reset_learned_memory():
    if os.path.exists(LEARNED_MEMORY_FILE):
        try:
            os.remove(LEARNED_MEMORY_FILE)
        except Exception:
            pass


# ------------------------------------------------------------------
# 📌 船段圖精準邊界 (CELL_BOUNDS) 與文字座標 (TEXT_BOUNDS) 對照表
# 嚴格約束廠商標記在格子內部 (內縮安全邊距)，不壓線、不遮蔽船段名稱
# ------------------------------------------------------------------
CELL_BOUNDS = {
    # Row M
    "B8M": (88, 405, 191, 447),
    "B7M": (192, 405, 296, 447),
    "B6M": (297, 405, 400, 447),
    "B5M": (401, 405, 504, 447),
    "B1M": (505, 405, 607, 447),
    "B2M": (608, 405, 711, 447),
    "B3M": (712, 405, 812, 447),
    "B4M": (813, 375, 895, 447),
    # Row S
    "B8S": (88, 448, 191, 487),
    "B7S": (192, 448, 296, 487),
    "B6S": (297, 448, 400, 487),
    "B5S": (401, 448, 504, 487),
    "B1S": (505, 448, 607, 487),
    "B2S": (608, 448, 711, 487),
    "B3S": (712, 448, 812, 487),
    "B4S": (813, 448, 912, 487),
    # Row T
    "B8T": (88, 488, 191, 528),
    "B7T": (192, 488, 296, 528),
    "B6T": (297, 488, 400, 528),
    "B5T": (401, 488, 504, 528),
    "B1T": (505, 488, 607, 528),
    "B2T": (608, 488, 711, 528),
    "B3T": (712, 488, 812, 528),
    "B4T": (813, 488, 912, 528),
    # Row B
    "B6B": (297, 529, 400, 562),
    "B5B": (401, 529, 504, 562),
    "B1B": (505, 529, 607, 562),
    "B2B": (608, 529, 711, 562),
    "B3B": (712, 529, 812, 562),
    "B4B": (813, 529, 912, 562),
    # Bow
    "FD": (913, 375, 1005, 447),
    "F1": (913, 448, 975, 555),
    # Stern
    "PDP": (22, 405, 87, 447),
    "PDU": (22, 448, 87, 487),
    "RD": (22, 488, 87, 532),
    "軸架": (80, 528, 140, 560),
    "SK": (165, 528, 220, 560),
    # Superstructure 1A
    "1A4": (235, 364, 360, 404),
    "1A3": (361, 364, 492, 404),
    "1A1": (608, 364, 745, 404),
    # Superstructure 2A
    "2A4": (242, 320, 360, 363),
    "2A3": (361, 320, 492, 363),
    "2A2 P/S": (493, 320, 607, 404),
    "2A2": (493, 320, 607, 404),
    "2A1": (608, 320, 745, 363),
    # Superstructure 3A
    "3A3": (364, 278, 488, 319),
    "3A1": (612, 265, 740, 319),
    # Superstructure Top
    "FN": (366, 218, 412, 277),
    "4A2": (608, 220, 638, 264),
    "4A1": (640, 220, 715, 264),
    "MAST": (645, 180, 700, 220),
}

TEXT_BOUNDS = {
    "1A1": (671, 380, 689, 388),
    "1A3": (418, 381, 439, 389),
    "1A4": (288, 380, 309, 388),
    "2A1": (668, 337, 688, 345),
    "2A2 P/S": (539, 355, 571, 368),
    "2A2": (539, 355, 571, 368),
    "2A3": (418, 337, 439, 346),
    "2A4": (290, 338, 312, 346),
    "3A1": (668, 292, 689, 300),
    "3A3": (419, 296, 440, 304),
    "4A1": (669, 223, 694, 242),
    "4A2": (616, 231, 633, 237),
    "B1B": (546, 541, 567, 550),
    "B1M": (543, 421, 568, 429),
    "B1S": (546, 464, 566, 472),
    "B1T": (546, 504, 564, 512),
    "B2B": (649, 541, 670, 550),
    "B2M": (646, 421, 672, 429),
    "B2S": (650, 464, 670, 472),
    "B2T": (650, 504, 671, 512),
    "B3B": (752, 541, 773, 550),
    "B3M": (748, 416, 773, 424),
    "B3S": (752, 464, 772, 472),
    "B3T": (753, 505, 771, 513),
    "B4B": (851, 541, 870, 550),
    "B4M": (849, 402, 862, 410),
    "B4S": (852, 464, 870, 472),
    "B4T": (852, 505, 870, 514),
    "B5B": (441, 541, 462, 550),
    "B5M": (437, 424, 463, 432),
    "B5S": (442, 464, 462, 472),
    "B5T": (442, 505, 460, 513),
    "B6B": (337, 542, 358, 550),
    "B6M": (335, 424, 361, 432),
    "B6S": (337, 464, 357, 472),
    "B6T": (337, 505, 355, 513),
    "B7M": (232, 421, 257, 430),
    "B7S": (234, 464, 254, 472),
    "B7T": (233, 517, 253, 526),
    "B8M": (126, 421, 151, 430),
    "B8S": (129, 464, 149, 472),
    "B8T": (128, 506, 149, 515),
    "F1": (930, 498, 955, 512),
    "FD": (955, 400, 980, 416),
    "FN": (386, 243, 399, 251),
    "MAST": (665, 215, 698, 237),
    "PDP": (40, 424, 62, 432),
    "PDU": (40, 481, 63, 489),
    "RD": (39, 518, 53, 526),
    "SK": (194, 542, 206, 550),
    "軸架": (88, 532, 114, 543),
}

BLOCK_MAP = {k: ((b[0] + b[2]) // 2, (b[1] + b[3]) // 2) for k, b in CELL_BOUNDS.items()}
UNIFIED_VENDOR_COLOR = (220, 53, 69)  # 醒目紅色廠商標籤


# ------------------------------------------------------------------
# 📌 工令圖號 (DWG NO) 對照船段表 (造船廠標準工令對照)
# ------------------------------------------------------------------
DWG_TO_BLOCK_MAP = {
    "MS01": "B3B",
    "MS02": "B2B",
    "MS03": "B1B",
    "MS04": "B5B",
    "MS05": "B6B",
    "MS06": "B8T",
    "MS07": "F1",
    "MS08": "B4T, B4B",
    "MS09": "B3S, B3T, B4S",
    "MS10": "B3T, B3S",
    "MS11": "B2T, B2S",
    "MS12": "B1T, B1S",
    "MS13": "B5T, B5S",
    "MS14": "B7T, B6T, B6S, B7S",
    "MS15": "B8S",
    "MS16": "PDU",
    "MS20": "B2M",
    "MS24": "B8S",
    "MS25": "PDU",
    "MS27": "1A1",
    "MS29": "1A3",
    "MS32": "2A1",
    "MS33": "1A3",
    "MS35": "3A1",
    "MS38": "3A1",
    "MS39": "MAST",
}

# 首要辨識廠商名單 (精準比對目標)
PRIMARY_VENDORS = [
    "啟進", "佰億", "良擅", "永一", "日順安",
    "勛發", "安勝", "放樣課", "偉勝", "堃翊",
    "興聯泰", "盛泰", "晉航", "廣通"
]


def normalize_dwg_token(tok: str) -> str:
    """將 MS1、MS01 等圖號格式統一為 MS01 標準格式。"""
    m = re.match(r"^MS(\d+)$", tok.upper())
    if m:
        num = int(m.group(1))
        return f"MS{num:02d}"
    return tok.upper()


def normalize_location_string(raw_loc: str) -> str:
    """校正動火位置常見誤判 (如 B15 必為 B1S)，並將 MS01、MS02 等圖號自動對照轉換為標準船段 (優先查驗 AI 學習記憶)。"""
    if not raw_loc:
        return ""
    s = str(raw_loc).strip()

    # 0. 優先查驗手動編輯之 AI 學習記憶
    mem = load_learned_memory()
    if s in mem.get("location_aliases", {}):
        s = mem["location_aliases"][s]

    # 1. 將 MS 圖號自動對照轉換為船段
    def repl_dwg(m):
        raw_code = m.group(1)
        std_code = normalize_dwg_token(raw_code)
        if std_code in DWG_TO_BLOCK_MAP:
            return f"{DWG_TO_BLOCK_MAP[std_code]} ({std_code})"
        return m.group(0)

    s = re.sub(r"\b(MS\d+)\b", repl_dwg, s, flags=re.IGNORECASE)
    s = re.sub(r"(?i)DWG\s*(?:NO)?[:\s]*", "", s).strip()

    # 2. 修正 S 誤判為 5 (如 B15 -> B1S, B25 -> B2S, B85 -> B8S)
    s = re.sub(r"(?i)B([1-8])5", r"B\1S", s)
    # 3. 修正 B 誤判為 8 (如 81S -> B1S, 85M -> B5M)
    s = re.sub(r"(?i)\b8([1-8][SMTB])\b", r"B\1", s)
    return s



def normalize_block_key(raw_loc: str) -> str:
    if not raw_loc:
        return ""
    s = (
        str(raw_loc)
        .strip()
        .upper()
        .replace(" ", "")
        .replace("-", "")
        .replace("船艙", "")
        .replace("船段", "")
        .replace("區", "")
    )

    # 支援若單獨傳入 MS 圖號也能直接對照轉換為船段代碼
    m_dwg = re.match(r"^MS(\d+)$", s)
    if m_dwg:
        std_dwg = f"MS{int(m_dwg.group(1)):02d}"
        if std_dwg in DWG_TO_BLOCK_MAP:
            s = DWG_TO_BLOCK_MAP[std_dwg].split(",")[0].strip()

    # 手寫辨識常見誤判校正 (如 B15 必為 B1S, 8 開頭必為 B)
    s = re.sub(r"B([1-8])5", r"B\1S", s)
    if s.startswith("8") and len(s) >= 3 and s[1:2].isdigit():
        s = "B" + s[1:]
    if s.startswith("I") or s.startswith("L"):
        s = "1" + s[1:]

    if not s:
        return ""

    for b in BLOCK_MAP.keys():
        b_clean = b.upper().replace(" ", "").replace("-", "").replace("/", "")
        if s == b_clean or b_clean in s:
            return b
    return ""


def extract_all_block_keys(raw_loc: str) -> List[str]:
    """支援解析單一動火位置字串中的多個船段 (包含 MS 圖號自動對照轉換)。"""
    if not raw_loc:
        return []
    # 先經由圖號對照與格式校正處理
    norm_loc = normalize_location_string(raw_loc)
    found_keys = []
    tokens = re.split(r"[,;、\/\s\+\&\n\|()]+", str(norm_loc))
    for tok in tokens:
        tok_clean = tok.strip()
        if not tok_clean:
            continue
        b_key = normalize_block_key(tok_clean)
        if b_key and b_key not in found_keys:
            found_keys.append(b_key)
    return found_keys


def extract_permit_date(records: List[Dict[str, Any]] = None) -> str:
    """自動產生民國紀年格式之當日日期 (如 115 年 10 月 13 日)。"""
    now = time.localtime()
    roc_year = now.tm_year - 1911
    return f"{roc_year} 年 {now.tm_mon} 月 {now.tm_mday} 日"


def find_template_path(tmpl_name: str) -> str:
    """智慧尋找底圖路徑，同時相容根目錄 (GitHub 根目錄)、templates/ 子目錄與大小寫。"""
    search_dirs = [BASE_DIR, TEMPLATES_DIR]
    for d in search_dirs:
        for name in [tmpl_name, tmpl_name.lower(), tmpl_name.upper()]:
            p = os.path.join(d, name)
            if os.path.exists(p):
                return p
    kw = "419" if "419" in tmpl_name else "418"
    for d in search_dirs:
        if os.path.exists(d):
            try:
                for fname in os.listdir(d):
                    if kw in fname and fname.lower().endswith((".png", ".jpg", ".jpeg")):
                        return os.path.join(d, fname)
            except Exception:
                pass
    return ""


def get_system_font(size: int = 14) -> ImageFont.FreeTypeFont:
    """跨平台尋找支援中文字型。"""
    candidate_paths = [
        "C:/Windows/Fonts/msjhbd.ttc",
        "C:/Windows/Fonts/msjh.ttc",
        "C:/Windows/Fonts/simsun.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]
    for p in candidate_paths:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()


def draw_vendor_badge(draw_ctx, x1, y1, x2, y2, text, font, pad_x=2, pad_y=1):
    draw_ctx.rounded_rectangle(
        [x1, y1, x2, y2],
        radius=2,
        fill=(*UNIFIED_VENDOR_COLOR, 220),
        outline=(*UNIFIED_VENDOR_COLOR, 255),
    )
    draw_ctx.text(
        (x1 + pad_x, y1 + pad_y - 1),
        text,
        fill=(255, 255, 255, 255),
        font=font,
    )


def generate_ship_diagram(
    records: List[Dict[str, Any]],
    target_project: str = "J418",
    custom_date: str = None,
) -> Image.Image:
    """根據辨識資料自動載入 J418 / J419 底圖，並以精準安全邊界標示廠商，絕不壓線且不遮蔽船段文字。"""
    tmpl_name = (
        "J419_template.png"
        if "J419" in str(target_project).upper()
        else "J418_template.png"
    )
    tmpl_path = find_template_path(tmpl_name)

    if not tmpl_path or not os.path.exists(tmpl_path):
        base_img = Image.new("RGBA", (1024, 576), (245, 247, 250, 255))
        draw_err = ImageDraw.Draw(base_img)
        draw_err.rectangle([20, 20, 1004, 556], outline=(220, 53, 69, 255), width=3)
        draw_err.text(
            (250, 260),
            f"⚠️ 找不到船圖底圖檔案: {tmpl_name}\n請確認已將 {tmpl_name} 放於根目錄或 templates/ 資料夾中！",
            fill=(220, 53, 69, 255),
        )
    else:
        base_img = Image.open(tmpl_path).convert("RGBA")

    overlay = Image.new("RGBA", base_img.size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(overlay)

    font_header = get_system_font(18)

    # 填入右上角日期 (覆蓋底圖舊有預印日期)
    date_str = custom_date or extract_permit_date(records)
    draw_base = ImageDraw.Draw(base_img)
    draw_base.rectangle([750, 95, 990, 135], fill=(255, 255, 255, 255))
    draw.text((780, 108), date_str, fill=(0, 0, 0, 255), font=font_header)

    block_vendors = {}
    for r in records:
        proj = str(r.get("專案名稱", "")).upper()
        if (
            target_project
            and proj
            and target_project.upper() not in proj
            and proj not in target_project.upper()
        ):
            continue

        loc_raw = r.get("動火位置", "")
        vendor = r.get("申請廠商", "").strip()
        if not vendor:
            continue

        all_blocks = extract_all_block_keys(loc_raw)
        for b_key in all_blocks:
            if vendor not in block_vendors.setdefault(b_key, []):
                block_vendors[b_key].append(vendor)

    def layout_row(v_list, target_w):
        sizes = [11, 10, 9, 8.5, 8, 7.5, 7]
        for f_sz in sizes:
            f = get_system_font(f_sz)
            px = 3 if f_sz >= 10 else 2
            gap = 2 if len(v_list) > 2 else 3
            widths = [f.getbbox(v)[2] - f.getbbox(v)[0] + px * 2 for v in v_list]
            tot = sum(widths) + (len(v_list) - 1) * gap
            if tot <= target_w or f_sz == sizes[-1]:
                return f, f_sz, widths, tot, gap, px
        return get_system_font(7), 7, widths, tot, 2, 2

    for block_key, vendors in block_vendors.items():
        if block_key not in CELL_BOUNDS:
            continue
        c_xmin, c_ymin, c_xmax, c_ymax = CELL_BOUNDS[block_key]
        t_xmin, t_ymin, t_xmax, t_ymax = TEXT_BOUNDS.get(
            block_key,
            (
                (c_xmin + c_xmax) // 2 - 10,
                (c_ymin + c_ymax) // 2 - 5,
                (c_xmin + c_xmax) // 2 + 10,
                (c_ymin + c_ymax) // 2 + 5,
            ),
        )

        safe_left = c_xmin + 2
        safe_right = c_xmax - 2
        safe_width = safe_right - safe_left

        # 特殊形狀區塊專屬安全定位 (避開斜角與邊界)
        if block_key == "F1":
            # F1 上半部寬敞空間 y in [466, 496]，寬度 65px，絕不壓斜線且絕不跨入 B4T
            safe_top = 466
            safe_bottom = 496
            safe_left = 916
            safe_right = 980
            safe_width = safe_right - safe_left
            space_mode = "custom_box"
        elif block_key == "FD":
            safe_top = 416
            safe_bottom = 445
            safe_left = 916
            safe_right = 985
            safe_width = safe_right - safe_left
            space_mode = "custom_box"
        elif block_key == "B4M":
            safe_right = 888
            safe_width = safe_right - safe_left
            space_mode = "below"
        elif block_key == "1A1":
            space_mode = "below"
        elif block_key == "B7T":
            # B7T 文字偏下，廠商標籤置於文字上方，絕不遮字
            space_mode = "above"
        else:
            space_below = (c_ymax - 2) - (t_ymax + 1)
            space_above = (t_ymin - 1) - (c_ymin + 2)
            space_mode = "below" if space_below >= 11 else "above"

        num_vendors = len(vendors)

        # 若大型或多廠商船段 (如 FD 聚集 3~4 家廠商)，自動以雙排佈局整齊排列
        if block_key == "FD" and num_vendors >= 3:
            mid = (num_vendors + 1) // 2
            rows = [vendors[:mid], vendors[mid:]]
            row_y = [safe_top, safe_top + 14]
            for r_idx, r_vendors in enumerate(rows):
                f, f_sz, widths, tot, gap, px = layout_row(r_vendors, safe_width)
                th = f.getbbox("良")[3] - f.getbbox("良")[1]
                bh = th + 2
                sx = (safe_left + safe_right) / 2 - tot / 2
                sx = max(safe_left, min(sx, safe_right - tot))
                cur_x = sx
                by1 = row_y[r_idx]
                for idx, vname in enumerate(r_vendors):
                    bw = widths[idx]
                    draw_vendor_badge(
                        draw, cur_x, by1, cur_x + bw, by1 + bh, vname, f, px, 1
                    )
                    cur_x += bw + gap
            continue

        # 單排自動縮放字體排列
        f, f_sz, widths, tot, gap, px = layout_row(vendors, safe_width)
        th = f.getbbox("良")[3] - f.getbbox("良")[1]
        bh = th + 2

        if space_mode == "custom_box":
            by1 = (safe_top + safe_bottom) / 2 - bh / 2
        elif space_mode == "below":
            avail_y1 = t_ymax + 1
            avail_y2 = c_ymax - 2
            by1 = (avail_y1 + avail_y2) / 2 - bh / 2
            by1 = max(avail_y1, min(by1, avail_y2 - bh))
        else:
            avail_y1 = c_ymin + 2
            avail_y2 = t_ymin - 1
            by1 = (avail_y1 + avail_y2) / 2 - bh / 2
            by1 = max(avail_y1, min(by1, avail_y2 - bh))

        sx = (safe_left + safe_right) / 2 - tot / 2
        sx = max(safe_left, min(sx, safe_right - tot))
        cur_x = sx
        for idx, vname in enumerate(vendors):
            bw = widths[idx]
            draw_vendor_badge(
                draw, cur_x, by1, cur_x + bw, by1 + bh, vname, f, px, 1
            )
            cur_x += bw + gap

    final_img = Image.alpha_composite(base_img, overlay).convert("RGB")
    return final_img


def generate_hotwork_pptx(
    records: List[Dict[str, Any]], custom_date: str = None
) -> io.BytesIO:
    """自動產生造船廠標準格式之動火管制 PPT 簡報檔 (16:9 比例，包含船段圖與管制表格)。"""
    if not PPTX_AVAILABLE:
        raise RuntimeError("python-pptx 模組未安裝，無法生成簡報。")

    prs = pptx.Presentation()
    prs.slide_width = Emu(9144000)
    prs.slide_height = Emu(5143500)
    blank_layout = prs.slide_layouts[6]

    projects_present = []
    for r in records:
        proj = normalize_project_name(r.get("專案名稱", ""))
        if proj not in projects_present:
            projects_present.append(proj)
    if not projects_present:
        projects_present = ["J418"]
    projects_present.sort()

    slide_counter = 1
    for proj in projects_present:
        proj_records = [
            r for r in records if normalize_project_name(r.get("專案名稱", "")) == proj
        ]
        if not proj_records and len(projects_present) > 1:
            continue

        # 1. 產生船段圖 Slide (全螢幕高解析度船段圖)
        slide_img = prs.slides.add_slide(blank_layout)
        diagram_img = generate_ship_diagram(
            records=records, target_project=proj, custom_date=custom_date
        )
        buf_img = io.BytesIO()
        diagram_img.save(buf_img, format="PNG")
        buf_img.seek(0)
        slide_img.shapes.add_picture(
            buf_img, 0, 0, width=prs.slide_width, height=prs.slide_height
        )

        # 船段圖右下角頁碼
        sn_img_box = slide_img.shapes.add_textbox(
            Emu(7010400), Emu(4869656), Emu(2000000), Emu(250000)
        )
        p_sn_img = sn_img_box.text_frame.paragraphs[0]
        p_sn_img.alignment = PP_ALIGN.RIGHT
        p_sn_img.text = str(slide_counter)
        p_sn_img.font.size = Pt(12)
        p_sn_img.font.bold = True
        p_sn_img.font.color.rgb = RGBColor(0, 51, 153)
        slide_counter += 1

        # 2. 產生管制表表格 Slide (完全對齊造船廠規範格式)
        slide_tbl = prs.slides.add_slide(blank_layout)

        # 標題
        tx_box = slide_tbl.shapes.add_textbox(
            Emu(983353), Emu(49933), Emu(7022252), Emu(369332)
        )
        tf = tx_box.text_frame
        tf.word_wrap = True
        p_title = tf.paragraphs[0]
        p_title.text = f"    {proj}    動火作業區域"
        p_title.font.name = "標楷體"
        p_title.font.size = Pt(22)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(0, 0, 0)

        # 表格
        headers = ["廠商名稱", "動火位置", "動火員姓名", "監火員姓名", "使用滅火器"]
        total_rows = max(10, len(proj_records) + 1)
        table_shape = slide_tbl.shapes.add_table(
            total_rows, 5, Emu(114770), Emu(500784), Emu(8907785), Emu(4408052)
        )
        table = table_shape.table

        for col in table.columns:
            col.width = Emu(1781557)

        for c_idx, h_text in enumerate(headers):
            cell = table.cell(0, c_idx)
            cell.text = h_text
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.font.name = "標楷體"
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 0, 0)

        for r_idx, rec in enumerate(proj_records):
            row_vals = [
                rec.get("申請廠商", ""),
                rec.get("動火位置", ""),
                rec.get("動火人員", ""),
                rec.get("監火人員", ""),
                rec.get("滅火器類型", "乾粉"),
            ]
            for c_idx, val in enumerate(row_vals):
                cell = table.cell(r_idx + 1, c_idx)
                cell.text = str(val)
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT
                p.font.name = "新細明體"
                p.font.size = Pt(12)
                p.font.color.rgb = RGBColor(0, 0, 0)

        # 右下角頁碼
        sn_box = slide_tbl.shapes.add_textbox(
            Emu(7010400), Emu(4869656), Emu(2000000), Emu(250000)
        )
        p_sn = sn_box.text_frame.paragraphs[0]
        p_sn.alignment = PP_ALIGN.RIGHT
        p_sn.text = str(slide_counter)
        p_sn.font.size = Pt(12)
        p_sn.font.bold = True
        p_sn.font.color.rgb = RGBColor(0, 0, 0)
        slide_counter += 1

    out_buf = io.BytesIO()
    prs.save(out_buf)
    out_buf.seek(0)
    return out_buf



# ------------------------------------------------------------------
# 📌 任務狀態持久化管理 (JSON 儲存)
# ------------------------------------------------------------------
def get_task_folder(task_id: str) -> str:
    folder = os.path.join(TASKS_DIR, task_id)
    os.makedirs(folder, exist_ok=True)
    return folder


def save_task_status(task_id: str, data: dict):
    folder = get_task_folder(task_id)
    file_path = os.path.join(folder, "status.json")
    data["updated_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_task_status(task_id: str) -> dict:
    file_path = os.path.join(TASKS_DIR, task_id, "status.json")
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def list_all_tasks() -> List[dict]:
    tasks = []
    if not os.path.exists(TASKS_DIR):
        return tasks
    for tid in os.listdir(TASKS_DIR):
        if tid.startswith("_"):
            continue
        t_data = load_task_status(tid)
        if t_data:
            tasks.append(t_data)
    tasks.sort(key=lambda x: x.get("created_at", ""), reverse=True)
    return tasks


def delete_task(task_id: str) -> bool:
    """刪除指定的辨識任務與其儲存的照片資料夾。"""
    folder = os.path.join(TASKS_DIR, task_id)
    if os.path.exists(folder):
        try:
            shutil.rmtree(folder, ignore_errors=True)
            return True
        except Exception:
            pass
    return False


def delete_all_tasks() -> int:
    """一鍵清空所有歷史任務與圖檔。"""
    count = 0
    if os.path.exists(TASKS_DIR):
        for tid in os.listdir(TASKS_DIR):
            if tid.startswith("_"):
                continue
            folder = os.path.join(TASKS_DIR, tid)
            if os.path.isdir(folder):
                try:
                    shutil.rmtree(folder, ignore_errors=True)
                    count += 1
                except Exception:
                    pass
    return count


# ------------------------------------------------------------------
# 📌 記憶體安全之極致輕量壓縮：單張處理完即刻釋放 C/Python 記憶體
# ------------------------------------------------------------------
def compress_and_save_image(
    raw_bytes: bytes, target_path: str, max_dimension: int = 768, quality: int = 70
) -> Tuple[float, float]:
    """將傳入的原始照片位元組壓縮為輕量 JPEG 並直接寫入磁碟，完全杜絕記憶體洩漏與 OOM。"""
    orig_kb = len(raw_bytes) / 1024.0
    buf_in = io.BytesIO(raw_bytes)
    img = Image.open(buf_in)

    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass

    if img.mode != "RGB":
        img = img.convert("RGB")

    img.thumbnail((max_dimension, max_dimension), Image.Resampling.LANCZOS)
    img.save(target_path, format="JPEG", quality=quality, optimize=True)

    comp_kb = os.path.getsize(target_path) / 1024.0

    # 關鍵：主動關閉並強制釋放記憶體，防範 Streamlit Cloud 1GB 記憶體限制崩潰
    img.close()
    buf_in.close()
    del img
    del buf_in
    gc.collect()

    return orig_kb, comp_kb


def normalize_project_name(raw_proj: str) -> str:
    """確保專案名稱只會有 J418 或 J419。"""
    if not raw_proj:
        return "J418"
    s = str(raw_proj).upper().strip()
    if "419" in s:
        return "J419"
    if "418" in s:
        return "J418"
    return "J418"


def clean_vendor_name(raw_vendor: str) -> str:
    """自動去除『工程行』、『有限公司』、『公司』等後綴，並以首要廠商名單與 AI 學習記憶為優先對齊目標。"""
    if not raw_vendor:
        return ""
    v = str(raw_vendor).strip()

    # 0. 優先查驗手動編輯之 AI 學習記憶別名
    mem = load_learned_memory()
    if v in mem.get("vendor_aliases", {}):
        return mem["vendor_aliases"][v]

    suffixes = [
        "實業有限公司", "股份有限公司", "工程有限公司", "工程企業社",
        "實業公司", "企業有限公司", "工程行", "有限公司", "企業社",
        "行號", "公司", "工作室", "工班"
    ]
    for s in suffixes:
        if v.endswith(s) and len(v) > len(s):
            v = v[:-len(s)].strip()
    if v.startswith("廠商:"):
        v = v[3:].strip()
    elif v.startswith("廠商"):
        v = v[2:].strip()

    # 合併內建首要廠商與使用者自訂已學習廠商
    all_targets = PRIMARY_VENDORS + [
        cv for cv in mem.get("custom_vendors", []) if cv not in PRIMARY_VENDORS
    ]

    # 1. 完全或包含比對首要廠商名單
    for pv in all_targets:
        if v == pv or v == pv + "工程" or pv in v:
            return pv

    # 2. 相似度模糊比對 (防手寫錯別字，如 堃羽 -> 堃翊)
    matches = difflib.get_close_matches(v, all_targets, n=1, cutoff=0.6)
    if matches:
        return matches[0]

    return v


def clean_extinguisher_type(raw_val: str) -> str:
    """標準化滅火器類型 (乾粉、CO2 或 無)。"""
    if not raw_val:
        return "無"
    s = str(raw_val).strip().upper()
    if any(kw in s for kw in ["CO2", "二氧化碳"]):
        return "CO2"
    if "乾粉" in s:
        return "乾粉"
    if "泡沫" in s:
        return "泡沫"
    if any(kw in s for kw in ["無", "NONE", "未勾選", "未填寫", "皆無"]):
        return "無"
    return str(raw_val).strip()


WORKER_ALIASES = {
    "蘇芍": "蘇荇",
    "蘇芩": "蘇荇",
    "蘇芳": "蘇荇",
    "曾龍山": "曾龍三",
    "黃龍三": "曾龍三",
    "張志中": "張志仲",
    "洪進豐": "洪進丰",
}

WATCHER_ALIASES = {
    "洪鉦軒": "洪鈺軒",
    "M. ASHAR TAHA": "M. Ashar Taha",
    "M. ASHAR": "M. Ashar Taha",
}


def calibrate_person_name(name: str, aliases: dict, is_watcher: bool = False) -> str:
    """校正常見手寫混淆姓名 (優先比對 AI 學習記憶與別名表)。"""
    if not name:
        return ""
    clean = str(name).strip()
    clean = aliases.get(clean, clean)
    mem = load_learned_memory()
    learned_aliases = mem.get(
        "watcher_aliases" if is_watcher else "worker_aliases", {}
    )
    return learned_aliases.get(clean, clean)


def record_user_corrections(orig_records: list, edited_records: list) -> int:
    """比對使用者在表格中的手動修改，學習記憶廠商、人員與位置之正確內容。"""
    if not orig_records or not edited_records:
        return 0
    mem = load_learned_memory()
    changed = 0

    for i in range(min(len(orig_records), len(edited_records))):
        orig = orig_records[i]
        ed = edited_records[i]

        # 1. 廠商學習
        v_orig = str(orig.get("申請廠商", "")).strip()
        v_new = str(ed.get("申請廠商", "")).strip()
        if v_new and v_new != v_orig:
            v_cleaned = clean_vendor_name(v_new)
            if (
                v_cleaned not in PRIMARY_VENDORS
                and v_cleaned not in mem["custom_vendors"]
            ):
                mem["custom_vendors"].append(v_cleaned)
            if v_orig and v_orig != v_cleaned:
                mem["vendor_aliases"][v_orig] = v_cleaned
            changed += 1

        # 2. 動火人員學習
        w_orig = str(orig.get("動火人員", "")).strip()
        w_new = str(ed.get("動火人員", "")).strip()
        if w_new and w_new != w_orig:
            if w_new not in mem["worker_roster"]:
                mem["worker_roster"].append(w_new)
            if w_orig:
                mem["worker_aliases"][w_orig] = w_new
            changed += 1

        # 3. 監火人員學習
        fw_orig = str(orig.get("監火人員", "")).strip()
        fw_new = str(ed.get("監火人員", "")).strip()
        if fw_new and fw_new != fw_orig:
            if fw_new not in mem["watcher_roster"]:
                mem["watcher_roster"].append(fw_new)
            if fw_orig:
                mem["watcher_aliases"][fw_orig] = fw_new
            changed += 1

        # 4. 動火位置學習
        loc_orig = str(orig.get("動火位置", "")).strip()
        loc_new = str(ed.get("動火位置", "")).strip()
        if loc_new and loc_new != loc_orig:
            if loc_orig:
                mem["location_aliases"][loc_orig] = loc_new
            changed += 1

    if changed > 0:
        mem["corrections_count"] = mem.get("corrections_count", 0) + changed
        save_learned_memory(mem)

    return changed



def ensure_valid_model(model_name: str) -> str:
    """強制過濾低配額易鎖死模型，確保採用超高額度之 Lite 模型。"""
    if not model_name:
        return DEFAULT_MODEL
    s = str(model_name).strip()
    if s in ["gemini-3.6-flash", "gemini-3.5-flash"] or any(
        old in s for old in ["2.5", "2.0", "1.5"]
    ):
        return DEFAULT_MODEL
    return s


# ------------------------------------------------------------------
# 📌 單張照片備援快速解析 (僅在整批偶發遺漏時啟動)
# ------------------------------------------------------------------
def process_single_image_ocr(
    client: genai.Client,
    candidate_models: List[str],
    base_prompt: str,
    fpath: str,
    fname: str,
) -> Dict[str, Any]:
    if not os.path.exists(fpath):
        return {
            "照片檔名": fname,
            "專案名稱": "辨識失敗",
            "申請廠商": "檔案不存在",
            "動火位置": "",
            "動火人員": "",
            "監火人員": "",
            "滅火器類型": "無",
        }

    with open(fpath, "rb") as f_img:
        c_bytes = f_img.read()

    single_prompt = (
        base_prompt
        + f"\n\n目前輸入僅包含 1 張照片，檔名為【{fname}】。請輸出包含此照片資料的 JSON 物件。"
    )
    contents = [
        single_prompt,
        f"【照片檔名: {fname}】",
        types.Part.from_bytes(data=c_bytes, mime_type="image/jpeg"),
    ]
    config = types.GenerateContentConfig(
        response_mime_type="application/json",
        temperature=0.1,
    )

    for m in candidate_models:
        try:
            response = client.models.generate_content(
                model=m, contents=contents, config=config
            )
            if response and response.text:
                clean_text = (
                    response.text.replace("```json", "").replace("```", "").strip()
                )
                data_obj = json.loads(clean_text)
                if isinstance(data_obj, list) and len(data_obj) > 0:
                    item = data_obj[0]
                elif isinstance(data_obj, dict):
                    item = data_obj
                else:
                    item = {}

                return {
                    "照片檔名": fname,
                    "專案名稱": normalize_project_name(
                        item.get("project", item.get("專案名稱", ""))
                    ),
                    "申請廠商": clean_vendor_name(
                        item.get("vendor", item.get("申請廠商", ""))
                    ),
                    "動火位置": normalize_location_string(
                        item.get("location", item.get("動火位置", ""))
                    ),
                    "動火人員": calibrate_person_name(
                        item.get("worker", item.get("動火人員", "")),
                        WORKER_ALIASES,
                        is_watcher=False,
                    ),
                    "監火人員": calibrate_person_name(
                        item.get("fire_watcher", item.get("監火人員", "")),
                        WATCHER_ALIASES,
                        is_watcher=True,
                    ),
                    "滅火器類型": clean_extinguisher_type(
                        item.get("extinguisher", item.get("滅火器類型", ""))
                    ),
                }
        except Exception:
            continue

    return {
        "照片檔名": fname,
        "專案名稱": "辨識失敗",
        "申請廠商": "模型連線逾時",
        "動火位置": "",
        "動火人員": "",
        "監火人員": "",
        "滅火器類型": "無",
    }


# ------------------------------------------------------------------
# 📌 後台獨立線程 worker：負責執行 Gemini API 辨識 (極速批次架構)
# ------------------------------------------------------------------
def run_background_ocr_task(
    task_id: str,
    key_str: str,
    model_name: str,
    batch_size: int,
    inter_batch_delay: int,
):
    task_data = load_task_status(task_id)
    if not task_data:
        return

    task_folder = get_task_folder(task_id)
    img_folder = os.path.join(task_folder, "images")

    filenames = task_data.get("filenames", [])
    total_files = len(filenames)
    total_batches = math.ceil(total_files / batch_size)

    client = genai.Client(api_key=key_str.strip())

    primary_model = ensure_valid_model(model_name)
    candidate_models = [primary_model] + [
        m for m in SAFE_MODELS if m != primary_model
    ]

    mem = load_learned_memory()
    all_priority_vendors = PRIMARY_VENDORS + [
        cv for cv in mem.get("custom_vendors", []) if cv not in PRIMARY_VENDORS
    ]
    vendors_prompt_str = ", ".join(all_priority_vendors)
    workers_learned_str = (
        ", " + ", ".join(mem.get("worker_roster", []))
        if mem.get("worker_roster")
        else ""
    )
    watchers_learned_str = (
        ", " + ", ".join(mem.get("watcher_roster", []))
        if mem.get("watcher_roster")
        else ""
    )

    prompt = f"""你是造船廠安全管理專家。請辨識傳入的照片中『動火工作安全許可申請表』的手寫文字。
請為傳入的每張照片提取資訊，並輸出 JSON 陣列 (每張照片一個物件，嚴格依照傳入照片順序)：
- "filename": 照片檔名 (請對應傳入檔名)
- "project": 專案名稱 (造船專案代號，限定為 "J418" 或 "J419")
- "vendor": 申請廠商名稱 (重要：請以以下廠商為首要辨識目標：{vendors_prompt_str}。若為其他廠商亦可。請務必去除『工程行』、『有限公司』、『公司』等後綴，只保留核心名字)
- "location": 動火位置 (重要規則：
  1. 嚴格依造船廠標準船段圖代號校正：S排為 B1S~B8S，絕無 B15 等數字，若寫成 B15 一律校正為 B1S；M排為 B1M~B8M；T排為 B1T~B8T；B排為 B1B~B6B；上層為 1A1, 1A3, 1A4, 2A1, 2A2, 2A3, 2A4, 3A1, 3A3, 4A1, 4A2, MAST, FN；艏艉有 PDP, PDU, RD, 軸架, SK, FD, F1。
  2. 若手寫的是工令圖號 (如 MS01, MS02...)，請依對照表轉換為標準船段：
     MS01 -> B3B
     MS02 -> B2B
     MS03 -> B1B
     MS04 -> B5B
     MS05 -> B6B
     MS06 -> B8T
     MS07 -> F1
     MS08 -> B4T, B4B
     MS09 -> B3S, B3T, B4S
     MS10 -> B3T, B3S
     MS11 -> B2T, B2S
     MS12 -> B1T, B1S
     MS13 -> B5T, B5S
     MS14 -> B7T, B6T, B6S, B7S
     MS15, MS24 -> B8S
     MS16, MS25 -> PDU
     MS20 -> B2M
     MS27 -> 1A1
     MS29, MS33 -> 1A3
     MS32 -> 2A1
     MS35, MS38 -> 3A1
     MS39 -> MAST
  若有多個船段請用逗號列出)
- "worker": 動火人員 (重要：只需呈現一位姓名。若名單有多位動火人員，請務必挑選其中一位『與監火人員不同姓名』的人員填入。請優先比對常用名單：黃清貴, 吳志成, 蘇荇, 洪進丰, 董光現, 張志仲, 曾龍三, 劉瑀慶, 莊漢源, 吳忠賢, 陳永聰, 林群弼, 朱明志, 林英勝, 王升郁{workers_learned_str} 等)
- "fire_watcher": 監火人員 (監火員姓名。請優先比對常用名單：黃招勢, 吳天送, 吳文送, M. Ashar Taha, 雷佳樺, 洪鈺軒, 張譽騰, 林彥廷, 陳厚助, 劉彥邦, 陳愛國, Aha, 楊貴同, 鍾經文, 林嘉良, 劉書成{watchers_learned_str} 等)
- "extinguisher": 滅火器類型 (申請單右上方的滅火器類型欄位，通常勾選『乾粉』或『CO2』，若皆無勾選請填『無』)

重要辨識規則：
1. 不用辨識動火日期或作業內容。
2. 專案名稱若寫 418 請轉為 J418，若寫 419 請轉為 J419。
3. 廠商名稱不可包含工程行、公司、企業社等詞彙，以首要廠商名單為優先。
4. 動火位置若為工令圖號 MS01 等請自動對照轉換為船段；若出現 B15 等代號必須校正為 B1S。
5. 動火人員欄位只能呈現一位姓名，且絕對不要與監火人員同一人。
6. 僅輸出純 JSON 陣列。"""

    config = types.GenerateContentConfig(
        response_mime_type="application/json",
        temperature=0.1,
    )

    all_results = []
    start_time = time.time()

    task_data["status"] = "processing"
    task_data["progress_pct"] = 0.0
    save_task_status(task_id, task_data)

    current_working_model = primary_model

    for b_idx in range(total_batches):
        start_idx = b_idx * batch_size
        end_idx = min(start_idx + batch_size, total_files)
        batch_filenames = filenames[start_idx:end_idx]

        task_data["msg"] = (
            f"⚡ 正在極速分析第 {start_idx + 1} ~ {end_idx} / {total_files} 張照片 (批次"
            f" {b_idx + 1}/{total_batches})..."
        )
        save_task_status(task_id, task_data)

        contents = [prompt]
        batch_files_objs = []

        for fname in batch_filenames:
            fpath = os.path.join(img_folder, fname)
            if os.path.exists(fpath):
                with open(fpath, "rb") as f_img:
                    c_bytes = f_img.read()
                contents.append(f"【照片檔名: {fname}】")
                contents.append(
                    types.Part.from_bytes(data=c_bytes, mime_type="image/jpeg")
                )
                batch_files_objs.append(fname)

        batch_results = []
        batch_success = False

        models_to_try = [current_working_model] + [
            m for m in candidate_models if m != current_working_model
        ]

        for active_model in models_to_try:
            try:
                response = client.models.generate_content(
                    model=active_model, contents=contents, config=config
                )
                if response and response.text:
                    clean_text = (
                        response.text.replace("```json", "")
                        .replace("```", "")
                        .strip()
                    )
                    data_list = json.loads(clean_text)
                    if isinstance(data_list, list):
                        for idx_sub, item in enumerate(data_list):
                            sub_fname = item.get("filename")
                            if not sub_fname and idx_sub < len(batch_files_objs):
                                sub_fname = batch_files_objs[idx_sub]
                            batch_results.append({
                                "照片檔名": sub_fname or "未知檔名",
                                "專案名稱": normalize_project_name(
                                    item.get("project", item.get("專案名稱", ""))
                                ),
                                "申請廠商": clean_vendor_name(
                                    item.get("vendor", item.get("申請廠商", ""))
                                ),
                                "動火位置": normalize_location_string(
                                    item.get("location", item.get("動火位置", ""))
                                ),
                                "動火人員": calibrate_person_name(
                                    item.get("worker", item.get("動火人員", "")),
                                    WORKER_ALIASES,
                                    is_watcher=False,
                                ),
                                "監火人員": calibrate_person_name(
                                    item.get("fire_watcher", item.get("監火人員", "")),
                                    WATCHER_ALIASES,
                                    is_watcher=True,
                                ),
                                "滅火器類型": clean_extinguisher_type(
                                    item.get("extinguisher", item.get("滅火器類型", ""))
                                ),
                            })
                        current_working_model = active_model
                        batch_success = True
                        break
            except Exception as e:
                err_str = str(e).lower()
                if any(
                    kw in err_str
                    for kw in [
                        "401",
                        "unauthenticated",
                        "invalid_argument",
                        "permission_denied",
                        "403",
                    ]
                ):
                    task_data["status"] = "error"
                    task_data["msg"] = (
                        f"❌ API Key 金鑰無效或未授權！請檢查 API Key: {e}"
                    )
                    save_task_status(task_id, task_data)
                    return
                # 限速或忙碌時不浪費時間乾等，即刻無縫換下一順位模型
                continue

        # 若整批偶發失敗或有缺漏，自動單張補救
        if not batch_success or len(batch_results) < len(batch_files_objs):
            already_done = {r.get("照片檔名") for r in batch_results}
            for fname in batch_files_objs:
                if fname not in already_done:
                    fpath = os.path.join(img_folder, fname)
                    res_single = process_single_image_ocr(
                        client=client,
                        candidate_models=candidate_models,
                        base_prompt=prompt,
                        fpath=fpath,
                        fname=fname,
                    )
                    batch_results.append(res_single)

        all_results.extend(batch_results)

        task_data["progress_pct"] = end_idx / total_files
        task_data["results"] = all_results
        task_data["stats"]["elapsed"] = time.time() - start_time
        save_task_status(task_id, task_data)

        if b_idx < total_batches - 1 and inter_batch_delay > 0:
            time.sleep(inter_batch_delay)

    elapsed_total = time.time() - start_time
    task_data["status"] = "completed"
    task_data["msg"] = (
        f"🎉 辨識完成！共處理 {total_files} 張照片，耗時 {elapsed_total:.1f} 秒"
    )
    task_data["progress_pct"] = 1.0
    task_data["stats"]["elapsed"] = elapsed_total
    save_task_status(task_id, task_data)


def get_default_api_key() -> str:
    """自動從 Streamlit Secrets 或環境變數讀取金鑰，不將金鑰寫死於程式碼中。"""
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return str(st.secrets["GEMINI_API_KEY"]).strip()
        if "API_KEY" in st.secrets:
            return str(st.secrets["API_KEY"]).strip()
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY", os.environ.get("API_KEY", "")).strip()


# ------------------------------------------------------------------
# 📌 側邊欄配置
# ------------------------------------------------------------------
st.sidebar.header("🔑 API 金鑰與連線設定")
default_key = get_default_api_key()
api_key_input = st.sidebar.text_input(
    "API Key:",
    value=default_key,
    type="password",
    placeholder="請輸入 Gemini API Key",
    help="建議於 Streamlit Cloud Secrets 或環境變數設定 GEMINI_API_KEY，亦可在此手動輸入。",
)
clean_key = api_key_input.strip() if api_key_input else ""
if clean_key:
    key_suffix = clean_key[-4:] if len(clean_key) >= 4 else "****"
    st.sidebar.caption(f"✅ 已載入 API 金鑰 (尾碼 `...{key_suffix}`)")
else:
    st.sidebar.warning("⚠️ 尚未設定 API Key，請於上方輸入或在 Streamlit Secrets 設定 `GEMINI_API_KEY`。")

model_choice = st.sidebar.selectbox(
    "AI 模型選擇 (高配額極速清單):",
    SAFE_MODELS,
    index=0,
    help="嚴選高免費配額、極速回應的 Flash Lite 模型，徹底杜絕限速與逾時問題！",
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 效能與流量優化參數")

batch_size = st.sidebar.slider(
    "📦 批次打包數量 (張/次):",
    1,
    14,
    7,
    help="建議 7 張/次，14 張相片只需 2 個批次約 8 秒即可全數完成",
)
inter_batch_delay = st.sidebar.slider(
    "⏱️ 批次間隔冷卻 (秒):",
    0,
    5,
    1,
    help="每個批次間微幅冷卻 1 秒，可確保永遠不觸發 API 流量峰值警告",
)
max_img_dim = st.sidebar.select_slider(
    "📐 圖片極限解析度 (px):",
    options=[640, 768, 1024],
    value=768,
    help="768px 為兼顧手寫字高清晰度與每張僅約 60KB 極速網路傳輸的最佳平衡點",
)

if st.sidebar.button("🔍 測試 API 金鑰連線"):
    if not clean_key:
        st.sidebar.error("❌ 請先填入正確的 API Key！")
    else:
        try:
            client = genai.Client(api_key=clean_key)
            active_model = ensure_valid_model(model_choice)
            res = client.models.generate_content(
                model=active_model, contents="Hello"
            )
            if res and res.text:
                st.sidebar.success(f"✅ API 連線成功！模型 ({active_model}) 運作正常")
            else:
                st.sidebar.error("❌ API 無回應")
        except Exception as e:
            st.sidebar.error(f"❌ 連線失敗: {e}")

# ------------------------------------------------------------------
# 📌 AI 自動學習記憶狀態面板
# ------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("🧠 AI 自動學習記憶庫")
current_mem = load_learned_memory()
cust_v_cnt = len(current_mem.get("custom_vendors", []))
w_cnt = len(current_mem.get("worker_roster", []))
fw_cnt = len(current_mem.get("watcher_roster", []))
corr_cnt = current_mem.get("corrections_count", 0)

st.sidebar.markdown(
    f"- 📝 已學習校正次數: **{corr_cnt}** 次\n"
    f"- 🏢 已學習自訂廠商: **{cust_v_cnt}** 家\n"
    f"- 👷 已學習動火人員: **{w_cnt}** 人\n"
    f"- 🧯 已學習監火人員: **{fw_cnt}** 人"
)

with st.sidebar.expander("🔍 檢視/管理 AI 記憶細項"):
    if cust_v_cnt > 0:
        st.write("🏢 **自訂廠商**:", ", ".join(current_mem["custom_vendors"]))
    if current_mem.get("vendor_aliases"):
        st.write("🔄 **廠商校正別名**:", current_mem["vendor_aliases"])
    if w_cnt > 0:
        st.write("👷 **常用動火名單**:", ", ".join(current_mem["worker_roster"]))
    if fw_cnt > 0:
        st.write("🧯 **常用監火名單**:", ", ".join(current_mem["watcher_roster"]))
    if current_mem.get("location_aliases"):
        st.write("📍 **位置校正對照**:", current_mem["location_aliases"])

    if st.button("🗑️ 清空所有 AI 學習記憶", use_container_width=True):
        reset_learned_memory()
        st.success("已清除所有 AI 學習記憶！")
        time.sleep(0.5)
        st.rerun()

# ------------------------------------------------------------------
# 📌 初始化手機友善之磁碟暫存池 (Disk-based Staging Pool)
# ------------------------------------------------------------------
if "session_token" not in st.session_state:
    st.session_state["session_token"] = (
        f"sess_{int(time.time())}_{random.randint(1000, 9999)}"
    )
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

current_staging_dir = os.path.join(
    STAGING_BASE_DIR, st.session_state["session_token"]
)
os.makedirs(current_staging_dir, exist_ok=True)

# ------------------------------------------------------------------
# 📌 主畫面 Tabs 分頁設計
# ------------------------------------------------------------------
tab_upload, tab_history = st.tabs(
    ["🚀 建立新辨識任務", "📋 檢視背景辨識任務與歷史紀錄"]
)

with tab_upload:
    st.subheader("📸 上傳動火單相片")
    st.markdown(
        "💡 **手機貼心支援**：支援一次多選上傳，亦可**分批多次追加選取相片**（例如先選 5 張、再選 5 張），相片會自動累加暫存，絕不漏拍或被覆蓋！"
    )

    new_uploads = st.file_uploader(
        "📁 請點擊選取或拖曳相片 (相容 iPhone HEIC/JPG/PNG):",
        accept_multiple_files=True,
        key=f"file_uploader_{st.session_state['uploader_key']}",
    )

    # 當選取相片時：立即一張一張壓縮並寫入磁碟暫存區，隨後立即釋放 RAM，確保伺服器永不崩潰
    if new_uploads:
        new_count = 0
        for f in new_uploads:
            # 依檔名儲存，避免重複
            dest_name = f.name
            dest_path = os.path.join(current_staging_dir, dest_name)
            if not os.path.exists(dest_path):
                raw_bytes = f.read()
                try:
                    compress_and_save_image(
                        raw_bytes, dest_path, max_dimension=max_img_dim, quality=70
                    )
                    new_count += 1
                except Exception as err:
                    st.error(f"⚠️ 照片 {f.name} 處理異常: {err}")
                finally:
                    del raw_bytes
                    gc.collect()

        if new_count > 0:
            st.toast(f"✅ 成功追加 {new_count} 張相片至待辨識清單！", icon="📸")

    # 取得目前暫存目錄中的所有已處理相片
    staged_filenames = [
        fn
        for fn in os.listdir(current_staging_dir)
        if fn.lower().endswith((".jpg", ".jpeg", ".png", ".heic"))
    ]
    staged_filenames.sort()
    buffer_count = len(staged_filenames)

    # 暫存池控制列
    if buffer_count > 0:
        total_disk_kb = sum(
            os.path.getsize(os.path.join(current_staging_dir, fn))
            for fn in staged_filenames
        ) / 1024.0

        col_pool1, col_pool2 = st.columns([3, 1])
        with col_pool1:
            st.success(
                f"📸 **待辨識相片池已累積：{buffer_count} 張相片** (輕量壓縮後總計僅 {total_disk_kb / 1024.0:.2f} MB，極速傳輸，無記憶體負擔！)"
            )
        with col_pool2:
            if st.button("🗑️ 清空所有暫存相片", use_container_width=True):
                shutil.rmtree(current_staging_dir, ignore_errors=True)
                os.makedirs(current_staging_dir, exist_ok=True)
                st.session_state["uploader_key"] += 1
                st.rerun()

        with st.expander(f"🖼️ 查看待辨識相片縮圖 ({buffer_count} 張)", expanded=False):
            cols = st.columns(4)
            for idx, fn in enumerate(staged_filenames):
                col = cols[idx % 4]
                img_path = os.path.join(current_staging_dir, fn)
                sz_kb = os.path.getsize(img_path) / 1024.0
                col.image(
                    img_path,
                    caption=f"{fn} ({sz_kb:.1f} KB)",
                    use_container_width=True,
                )

        if st.button(
            f"🚀 啟動伺服器背景 AI 自動辨識 (共 {buffer_count} 張相片)",
            type="primary",
            use_container_width=True,
        ):
            if not clean_key:
                st.error("❌ 請先填寫有效的 API Key！")
            else:
                task_id = time.strftime("task_%Y%m%d_%H%M%S")
                task_folder = get_task_folder(task_id)
                task_img_folder = os.path.join(task_folder, "images")
                os.makedirs(task_img_folder, exist_ok=True)

                # 將暫存池檔案直接搬移至任務資料夾
                filenames = []
                sum_comp_kb = 0.0

                for fname in staged_filenames:
                    src_f = os.path.join(current_staging_dir, fname)
                    dst_f = os.path.join(task_img_folder, fname)
                    shutil.move(src_f, dst_f)
                    filenames.append(fname)
                    sum_comp_kb += os.path.getsize(dst_f) / 1024.0

                task_init_data = {
                    "task_id": task_id,
                    "status": "pending",
                    "progress_pct": 0.0,
                    "msg": "任務已建立，背景辨識線程啟動中...",
                    "total_files": len(filenames),
                    "filenames": filenames,
                    "results": [],
                    "stats": {
                        "comp_kb": sum_comp_kb,
                        "comp_mb": sum_comp_kb / 1024.0,
                        "orig_mb": sum_comp_kb * 10 / 1024.0,  # 預估原圖約 10 倍大
                        "elapsed": 0.0,
                    },
                    "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                }
                save_task_status(task_id, task_init_data)

                # 啟動獨立背景線程執行辨識
                t = threading.Thread(
                    target=run_background_ocr_task,
                    kwargs={
                        "task_id": task_id,
                        "key_str": clean_key,
                        "model_name": model_choice,
                        "batch_size": batch_size,
                        "inter_batch_delay": inter_batch_delay,
                    },
                    daemon=True,
                )
                t.start()

                # 重設上傳組件以釋放前端快取
                st.session_state["uploader_key"] += 1
                st.session_state["active_task_id"] = task_id

                st.success(f"🎉 任務已於伺服器背景啟動！(任務編號: `{task_id}`)")
                st.info(
                    "💡 **手機貼心提示**：辨識正在伺服器背景高速執行（14張照片僅需約 10 秒）。您現在可以**切換到『檢視背景辨識任務』分頁**查看即時進度與成果！"
                )
    else:
        st.info("👆 請於上方選取照片，選取後可一鍵啟動高速背景辨識。")

with tab_history:
    all_tasks = list_all_tasks()

    if not all_tasks:
        st.info("尚無任何辨識任務紀錄。請先於第一個分頁建立任務。")
    else:
        task_options = {
            f"{t.get('task_id')} | 照片數: {t.get('total_files')} | 狀態:"
            f" {t.get('status')} ({t.get('created_at')})": t.get("task_id")
            for t in all_tasks
        }

        default_index = 0
        if "active_task_id" in st.session_state:
            for idx, tid in enumerate(task_options.values()):
                if tid == st.session_state["active_task_id"]:
                    default_index = idx
                    break

        selected_label = st.selectbox(
            "選擇要檢視的辨識任務:",
            options=list(task_options.keys()),
            index=default_index,
        )

        selected_task_id = task_options[selected_label]
        curr_task = load_task_status(selected_task_id)

        col_btn1, col_btn2, col_btn3 = st.columns([1.5, 1.5, 3])
        with col_btn1:
            if st.button("🔄 刷新最新進度", use_container_width=True):
                st.rerun()

        with col_btn2:
            with st.popover("🗑️ 刪除目前任務"):
                st.warning(
                    f"⚠️ 確定要刪除任務 `{selected_task_id}` 及其所有照片與辨識結果嗎？"
                )
                if st.button(
                    "🚨 確定刪除此任務", type="primary", use_container_width=True
                ):
                    delete_task(selected_task_id)
                    if st.session_state.get("active_task_id") == selected_task_id:
                        st.session_state.pop("active_task_id", None)
                    st.success(f"🗑️ 已成功刪除任務 `{selected_task_id}`！")
                    time.sleep(0.5)
                    st.rerun()

        with col_btn3:
            with st.popover("🧹 清空所有歷史任務"):
                st.warning(
                    "⚠️ 確定要刪除伺服器上的所有歷史辨識紀錄與圖檔嗎？此動作無法復原！"
                )
                if st.button(
                    "🚨 確定全部清空", type="primary", use_container_width=True
                ):
                    del_cnt = delete_all_tasks()
                    st.session_state.pop("active_task_id", None)
                    st.success(f"🧹 已成功清空 {del_cnt} 筆歷史任務紀錄！")
                    time.sleep(0.5)
                    st.rerun()

        status_str = curr_task.get("status", "unknown")
        progress_pct = curr_task.get("progress_pct", 0.0)
        msg_str = curr_task.get("msg", "")

        if status_str == "processing":
            st.warning(f"⏳ **任務處理中...** ({msg_str})")
            st.progress(progress_pct)
        elif status_str == "completed":
            st.success(f"✅ **{msg_str}**")
            st.progress(1.0)
        elif status_str == "pending":
            st.info(f"⏳ **{msg_str}**")
            st.progress(0.0)
        else:
            st.error(f"❌ 任務狀態: {status_str} ({msg_str})")

        stats = curr_task.get("stats", {})
        if stats:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("總辨識照片", f"{curr_task.get('total_files', 0)} 張")
            c2.metric("總處理耗時", f"{stats.get('elapsed', 0):.1f} 秒")
            comp_mb = stats.get("comp_mb", 0)
            orig_mb = stats.get("orig_mb", comp_mb * 10)
            saved_mb = max(0, orig_mb - comp_mb)
            saved_pct = (saved_mb / orig_mb * 100) if orig_mb > 0 else 0
            c3.metric(
                "圖片壓縮節省流量",
                f"{saved_mb:.2f} MB",
                delta=f"-{saved_pct:.1f}%",
            )
            c4.metric("壓縮後傳輸大小", f"{comp_mb:.2f} MB")

        results = curr_task.get("results", [])
        if results:
            st.write("---")
            st.write(
                "### 📊 今日動火作業彙整表 (雙擊儲存格即可修改校對，下方船段圖會即時同步更新)："
            )

            columns_order = [
                "照片檔名",
                "專案名稱",
                "申請廠商",
                "動火位置",
                "動火人員",
                "監火人員",
                "滅火器類型",
            ]
            df = pd.DataFrame(results)
            existing_cols = [c for c in columns_order if c in df.columns]
            df = df[existing_cols]

            edited_df = st.data_editor(
                df,
                num_rows="dynamic",
                use_container_width=True,
                key=f"editor_{selected_task_id}",
            )

            records_list = edited_df.to_dict(orient="records")

            # ------------------------------------------------------------------
            # 📌 AI 自動學習核心連動：即時比對手動編輯並儲存記憶
            # ------------------------------------------------------------------
            if records_list != results:
                learned_changes = record_user_corrections(results, records_list)
                curr_task["results"] = records_list
                save_task_status(selected_task_id, curr_task)
                if learned_changes > 0:
                    st.toast(
                        f"🧠 AI 已自動學習並記憶您所修正的 {learned_changes} 筆資料！",
                        icon="💡",
                    )

            # ------------------------------------------------------------------
            # 📌 自動生成動火管制船段圖 (即時響應同步)
            # ------------------------------------------------------------------
            st.write("---")
            st.write(
                "### 🗺️ 動火作業區域自動標示船段圖 (修改表格資料，船圖會立即同步更新)："
            )

            col_diag_sel, col_diag_date = st.columns([2, 2])
            with col_diag_sel:
                proj_choice = st.radio(
                    "選擇要標示的專案船段圖:",
                    ["J418", "J419"],
                    horizontal=True,
                )
            with col_diag_date:
                extracted_date = extract_permit_date(records_list)
                custom_date_input = st.text_input(
                    "管制表右上角日期 (預設民國當日日期，可自訂修改):",
                    value=extracted_date,
                )

            diagram_img = generate_ship_diagram(
                records=records_list,
                target_project=proj_choice,
                custom_date=custom_date_input,
            )

            st.image(
                diagram_img,
                caption=f"自動生成之 {proj_choice} 動火管制船段圖",
                use_container_width=True,
            )

            # 下載圖檔與一鍵生成 PPT 按鈕
            buf_img = io.BytesIO()
            diagram_img.save(buf_img, format="PNG")
            col_img_dl1, col_img_dl2 = st.columns(2)
            with col_img_dl1:
                st.download_button(
                    label=f"🖼️ 下載 {proj_choice} 動火管制船段圖 (.png)",
                    data=buf_img.getvalue(),
                    file_name=f"{proj_choice}_動火管制區域圖_{selected_task_id}.png",
                    mime="image/png",
                    use_container_width=True,
                )
            with col_img_dl2:
                if PPTX_AVAILABLE:
                    try:
                        buf_pptx_btn = generate_hotwork_pptx(
                            records=records_list, custom_date=custom_date_input
                        )
                        st.download_button(
                            label="📊 一鍵生成 PPT 簡報檔 (.pptx)",
                            data=buf_pptx_btn.getvalue(),
                            file_name=f"中信造船專案動火管制表_{selected_task_id}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            type="primary",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"PPT 生成失敗: {e}")
                else:
                    st.warning("⚠️ 請安裝 python-pptx 模組以啟用 PPT 匯出功能。")

            # ------------------------------------------------------------------
            # 📌 資料表格與簡報匯出區
            # ------------------------------------------------------------------
            st.write("---")
            st.write("#### 📥 匯出成果表格與簡報資料")
            col_dl1, col_dl2, col_dl3, col_dl4 = st.columns(4)

            csv_bytes = edited_df.to_csv(index=False).encode("utf-8-sig")
            col_dl1.download_button(
                label="📄 下載 CSV 表格 (.csv)",
                data=csv_bytes,
                file_name=f"動火管制表_{selected_task_id}.csv",
                mime="text/csv",
                use_container_width=True,
            )

            buf_excel = io.BytesIO()
            with pd.ExcelWriter(buf_excel, engine="openpyxl") as writer:
                edited_df.to_excel(writer, index=False, sheet_name="動火管制表")
            col_dl2.download_button(
                label="📊 下載 Excel 表格 (.xlsx)",
                data=buf_excel.getvalue(),
                file_name=f"動火管制表_{selected_task_id}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            json_bytes = edited_df.to_json(
                orient="records", force_ascii=False, indent=2
            ).encode("utf-8")
            col_dl3.download_button(
                label="🏷️ 下載 JSON 資料 (.json)",
                data=json_bytes,
                file_name=f"動火管制表_{selected_task_id}.json",
                mime="application/json",
                use_container_width=True,
            )

            if PPTX_AVAILABLE:
                try:
                    buf_pptx_exp = generate_hotwork_pptx(
                        records=records_list, custom_date=custom_date_input
                    )
                    col_dl4.download_button(
                        label="📈 下載 PPT 簡報 (.pptx)",
                        data=buf_pptx_exp.getvalue(),
                        file_name=f"中信造船專案動火管制表_{selected_task_id}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                    )
                except Exception as e:
                    col_dl4.error(f"PPT 生成失敗: {e}")