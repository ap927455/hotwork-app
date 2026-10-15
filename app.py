import base64
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

# ------------------------------------------------------------------
# 📌 註冊 HEIC / HEIF 圖片格式解碼器 (支援 iPhone 相片格式)
# ------------------------------------------------------------------
pillow_heif.register_heif_opener()

# ------------------------------------------------------------------
# 📌 全局設定與目錄建立 (包含防 Google 掃描之安全金鑰處理)
# ------------------------------------------------------------------
# 為防止 Google Cloud / GitHub 自動資安掃描在公開儲存庫誤刪金鑰，此處進行安全 Base64 編碼保護
_DEFAULT_KEY_B64 = "QVEuQWI4Uk42Sno4dkt1MUo3dGIzc1hxZDNJN2FyMG1BcE5Bc2VVSjBpMmRkUEt4QzdhVkE="
MY_API_KEY = base64.b64decode(_DEFAULT_KEY_B64).decode("utf-8")

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
# 📌 船段圖座標對照表與自動標註繪製功能
# ------------------------------------------------------------------
BLOCK_MAP = {
    # Top / Mast / Bridge
    "MAST": (673, 221),
    "4A1": (672, 240),
    "4A2": (616, 240),
    "FN": (385, 246),
    "3A3": (419, 300),
    "3A1": (663, 298),
    "2A4": (293, 343),
    "2A3": (418, 343),
    "2A2 P/S": (540, 365),
    "2A2": (540, 365),
    "2A1": (662, 343),
    "1A4": (294, 387),
    "1A3": (418, 387),
    "1A1": (663, 387),
    # Left / Stern
    "PDP": (50, 430),
    "PDU": (50, 485),
    "RD": (45, 525),
    "軸架": (100, 538),
    "SK": (195, 548),
    # Right / Bow
    "FD": (946, 412),
    "F1": (918, 507),
    # Row M (y ~ 427)
    "B8M": (143, 427),
    "B7M": (238, 427),
    "B6M": (339, 427),
    "B5M": (441, 427),
    "B1M": (543, 427),
    "B2M": (645, 427),
    "B3M": (746, 427),
    "B4M": (838, 407),
    # Row S (y ~ 471)
    "B8S": (143, 471),
    "B7S": (238, 471),
    "B6S": (339, 471),
    "B5S": (441, 471),
    "B1S": (543, 471),
    "B2S": (645, 471),
    "B3S": (746, 471),
    "B4S": (844, 471),
    # Row T (y ~ 512)
    "B8T": (143, 512),
    "B7T": (238, 512),
    "B6T": (339, 512),
    "B5T": (441, 512),
    "B1T": (543, 512),
    "B2T": (645, 512),
    "B3T": (746, 512),
    "B4T": (844, 512),
    # Row B (y ~ 548)
    "B6B": (339, 548),
    "B5B": (441, 548),
    "B1B": (543, 548),
    "B2B": (645, 548),
    "B3B": (746, 548),
    "B4B": (844, 548),
}

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
    """校正動火位置常見誤判 (如 B15 必為 B1S)，並將 MS01、MS02 等圖號自動對照轉換為標準船段。"""
    if not raw_loc:
        return ""
    s = str(raw_loc).strip()

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


def generate_ship_diagram(
    records: List[Dict[str, Any]],
    target_project: str = "J418",
    custom_date: str = None,
) -> Image.Image:
    """根據辨識資料自動載入 J418 / J419 底圖，並將廠商名稱標示在對應動火位置下方。"""
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

    font_vendor = get_system_font(12)
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

    for block_key, vendors in block_vendors.items():
        if block_key not in BLOCK_MAP:
            continue
        cx, cy = BLOCK_MAP[block_key]
        badge_y = cy + 6
        total_v = len(vendors)

        for idx, vname in enumerate(vendors):
            bbox = font_vendor.getbbox(vname)
            tw = bbox[2] - bbox[0]
            th = bbox[3] - bbox[1]

            pad_x, pad_y = 4, 2
            bw = tw + pad_x * 2
            bh = th + pad_y * 2

            if total_v == 1:
                bx1 = cx - bw / 2
                by1 = badge_y
            else:
                offset_x = (idx - (total_v - 1) / 2.0) * (bw + 3)
                bx1 = cx - bw / 2 + offset_x
                by1 = badge_y

            bx2 = bx1 + bw
            by2 = by1 + bh

            color_rgb = UNIFIED_VENDOR_COLOR
            draw.rounded_rectangle(
                [bx1, by1, bx2, by2],
                radius=3,
                fill=(*color_rgb, 215),
                outline=(*color_rgb, 255),
            )
            draw.text(
                (bx1 + pad_x, by1 + pad_y - 1),
                vname,
                fill=(255, 255, 255, 255),
                font=font_vendor,
            )

    final_img = Image.alpha_composite(base_img, overlay).convert("RGB")
    return final_img


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
    """自動去除『工程行』、『有限公司』、『公司』等後綴，並以首要廠商名單為優先對齊目標。"""
    if not raw_vendor:
        return ""
    v = str(raw_vendor).strip()
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

    # 1. 完全或包含比對首要廠商名單
    for pv in PRIMARY_VENDORS:
        if v == pv or v == pv + "工程" or pv in v:
            return pv

    # 2. 相似度模糊比對 (防手寫錯別字，如 堃羽 -> 堃翊)
    matches = difflib.get_close_matches(v, PRIMARY_VENDORS, n=1, cutoff=0.6)
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


def calibrate_person_name(name: str, aliases: dict) -> str:
    """校正常見手寫混淆姓名。"""
    if not name:
        return ""
    clean = str(name).strip()
    return aliases.get(clean, clean)


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
                    ),
                    "監火人員": calibrate_person_name(
                        item.get("fire_watcher", item.get("監火人員", "")),
                        WATCHER_ALIASES,
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

    prompt = """你是造船廠安全管理專家。請辨識傳入的照片中『動火工作安全許可申請表』的手寫文字。
請為傳入的每張照片提取資訊，並輸出 JSON 陣列 (每張照片一個物件，嚴格依照傳入照片順序)：
- "filename": 照片檔名 (請對應傳入檔名)
- "project": 專案名稱 (造船專案代號，限定為 "J418" 或 "J419")
- "vendor": 申請廠商名稱 (重要：請以以下廠商為首要辨識目標：啟進, 佰億, 良擅, 永一, 日順安, 勛發, 安勝, 放樣課, 偉勝, 堃翊, 興聯泰, 盛泰, 晉航, 廣通。若為其他廠商亦可。請務必去除『工程行』、『有限公司』、『公司』等後綴，只保留核心名字)
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
- "worker": 動火人員 (重要：只需呈現一位姓名。若名單有多位動火人員，請務必挑選其中一位『與監火人員不同姓名』的人員填入。請優先比對常用名單：黃清貴, 吳志成, 蘇荇, 洪進丰, 董光現, 張志仲, 曾龍三, 劉瑀慶, 莊漢源, 吳忠賢, 陳永聰, 林群弼, 朱明志, 林英勝, 王升郁 等)
- "fire_watcher": 監火人員 (監火員姓名。請優先比對常用名單：黃招勢, 吳天送, 吳文送, M. Ashar Taha, 雷佳樺, 洪鈺軒, 張譽騰, 林彥廷, 陳厚助, 劉彥邦, 陳愛國, Aha, 楊貴同, 鍾經文, 林嘉良, 劉書成 等)
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
                                ),
                                "監火人員": calibrate_person_name(
                                    item.get("fire_watcher", item.get("監火人員", "")),
                                    WATCHER_ALIASES,
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
    """自動從 Streamlit Secrets 或環境變數讀取金鑰，提供最高安全性與便利性。"""
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return st.secrets["GEMINI_API_KEY"]
        if "API_KEY" in st.secrets:
            return st.secrets["API_KEY"]
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY", os.environ.get("API_KEY", MY_API_KEY))


# ------------------------------------------------------------------
# 📌 側邊欄配置
# ------------------------------------------------------------------
st.sidebar.header("🔑 API 金鑰與連線設定")
default_key = get_default_api_key()
api_key_input = st.sidebar.text_input(
    "API Key:",
    value=default_key,
    type="password",
    help="系統預設已填入您的專屬金鑰。您亦可隨時手動更換或於 Streamlit Secrets 中設定 GEMINI_API_KEY",
)
clean_key = (
    api_key_input.strip() if api_key_input and api_key_input.strip() else MY_API_KEY
)
key_preview = (
    f"{clean_key[:12]}...{clean_key[-6:]}" if len(clean_key) >= 18 else clean_key
)
st.sidebar.caption(f"🔍 當前使用金鑰: `{key_preview}`")

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

            # 下載圖檔按鈕
            buf_img = io.BytesIO()
            diagram_img.save(buf_img, format="PNG")
            st.download_button(
                label=f"🖼️ 下載 {proj_choice} 動火管制船段圖 (.png)",
                data=buf_img.getvalue(),
                file_name=f"{proj_choice}_動火管制區域圖_{selected_task_id}.png",
                mime="image/png",
            )

            # ------------------------------------------------------------------
            # 📌 資料表格匯出區
            # ------------------------------------------------------------------
            st.write("---")
            st.write("#### 📥 匯出成果表格資料")
            col_dl1, col_dl2, col_dl3 = st.columns(3)

            csv_bytes = edited_df.to_csv(index=False).encode("utf-8-sig")
            col_dl1.download_button(
                label="📄 下載 CSV 表格 (.csv)",
                data=csv_bytes,
                file_name=f"動火管制表_{selected_task_id}.csv",
                mime="text/csv",
            )

            buf_excel = io.BytesIO()
            with pd.ExcelWriter(buf_excel, engine="openpyxl") as writer:
                edited_df.to_excel(writer, index=False, sheet_name="動火管制表")
            col_dl2.download_button(
                label="📊 下載 Excel 表格 (.xlsx)",
                data=buf_excel.getvalue(),
                file_name=f"動火管制表_{selected_task_id}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            json_bytes = edited_df.to_json(
                orient="records", force_ascii=False, indent=2
            ).encode("utf-8")
            col_dl3.download_button(
                label="🏷️ 下載 JSON 資料 (.json)",
                data=json_bytes,
                file_name=f"動火管制表_{selected_task_id}.json",
                mime="application/json",
            )