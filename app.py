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
# 📌 全局設定與目錄建立
# ------------------------------------------------------------------
# 預設金鑰更新為使用者指定之最新金鑰
MY_API_KEY = "AQ.Ab8RN6Ket4k_wQVbI4-YMZ81aB2vVEaNvoVkSuQBOkqIYQaImA"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TASKS_DIR = os.path.join(BASE_DIR, "tasks_storage")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
os.makedirs(TASKS_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

# 具備超高免費配額 (RPD/RPM) 與亞秒級辨識的高效模型清單 (徹底排除容易觸發 20 RPD 限額的 3.6-flash 與 3.5-flash)
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
    "⚡ 支援伺服器背景極速辨識 (14張相片約10秒完成) 與 **J418/J419 船段圖自動標示**"
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

    # 手寫辨識常見誤判校正 (如數字 8 誤判為字母 B，或 B 誤判為 8)
    if s.startswith("8") and len(s) >= 3 and s[1:2].isdigit():
        s = "B" + s[1:]
    if s.startswith("I") or s.startswith("L"):
        s = "1" + s[1:]

    for b in BLOCK_MAP.keys():
        b_clean = b.upper().replace(" ", "").replace("-", "").replace("/", "")
        if s == b_clean or b_clean in s or s in b_clean:
            return b
    return ""


def extract_all_block_keys(raw_loc: str) -> List[str]:
    """支援解析單一動火位置字串中的多個船段 (如 'B1T、B5T、B5S' 或 'B1T, B5T, B5S船艙')。"""
    if not raw_loc:
        return []
    found_keys = []
    tokens = re.split(r"[,;、\/\s\+\&\n\|]+", str(raw_loc))
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
    # 1. 優先直接路徑比對
    for d in search_dirs:
        for name in [tmpl_name, tmpl_name.lower(), tmpl_name.upper()]:
            p = os.path.join(d, name)
            if os.path.exists(p):
                return p
    # 2. 關鍵字模糊比對 (419 或 418)
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
    """跨平台尋找支援中文字型 (相容 Windows 與 Linux/Streamlit Cloud)。"""
    candidate_paths = [
        # Windows 微軟正黑體
        "C:/Windows/Fonts/msjhbd.ttc",
        "C:/Windows/Fonts/msjh.ttc",
        "C:/Windows/Fonts/simsun.ttc",
        # Linux / Streamlit Cloud 開源中文字型
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
    """根據辨識資料自動載入 J418 / J419 底圖，並將廠商名稱標示在對應動火位置下方 (不遮擋船段名稱)。"""
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

    # 填入右上角日期 (先將原本底圖右上角預印日期區域填白覆蓋，再輸出動火日期)
    date_str = custom_date or extract_permit_date(records)
    draw_base = ImageDraw.Draw(base_img)
    draw_base.rectangle([750, 95, 990, 135], fill=(255, 255, 255, 255))
    draw.text((780, 108), date_str, fill=(0, 0, 0, 255), font=font_header)

    # 整理 (動火位置 -> 廠商名稱列表，支援同一列包含多個船段)
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

    # 在對應船段下方繪製統一顏色的標籤
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
            folder = os.path.join(TASKS_DIR, tid)
            if os.path.isdir(folder):
                try:
                    shutil.rmtree(folder, ignore_errors=True)
                    count += 1
                except Exception:
                    pass
    return count


# ------------------------------------------------------------------
# 📌 圖片壓縮處理 (極致輕量化：清晰手寫字 + 最小網路傳輸)
# ------------------------------------------------------------------
def compress_and_prep_image_bytes(
    file_bytes: bytes, max_dimension: int = 768, quality: int = 70
) -> Tuple[bytes, float, float]:
    orig_size_kb = len(file_bytes) / 1024.0
    buf_in = io.BytesIO(file_bytes)
    img = Image.open(buf_in)

    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass

    if img.mode != "RGB":
        img = img.convert("RGB")

    img.thumbnail((max_dimension, max_dimension), Image.Resampling.LANCZOS)
    buf_out = io.BytesIO()
    img.save(buf_out, format="JPEG", quality=quality, optimize=True)
    comp_bytes = buf_out.getvalue()
    comp_size_kb = len(comp_bytes) / 1024.0

    return comp_bytes, orig_size_kb, comp_size_kb


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


def ensure_valid_model(model_name: str) -> str:
    """強制確保採用具備超高免費配額之模型，若選擇了有20次限制的舊模型則自動切換。"""
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
    """針對單張照片進行快速解析，自動在可用模型間輪替。"""
    if not os.path.exists(fpath):
        return {
            "照片檔名": fname,
            "專案名稱": "辨識失敗",
            "申請廠商": "檔案不存在",
            "動火位置": "",
            "動火人員": "",
            "監火人員": "",
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
                    "申請廠商": item.get("vendor", item.get("申請廠商", "")),
                    "動火位置": item.get("location", item.get("動火位置", "")),
                    "動火人員": item.get("worker", item.get("動火人員", "")),
                    "監火人員": item.get("fire_watcher", item.get("監火人員", "")),
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

    # 備妥模型輪替清單 (若首選遇到忙碌，立即自動無縫切換到下一順位)
    primary_model = ensure_valid_model(model_name)
    candidate_models = [primary_model] + [
        m for m in SAFE_MODELS if m != primary_model
    ]

    prompt = """你是造船廠安全管理專家。請辨識傳入的照片中『動火工作安全許可申請表』的手寫文字。
請為傳入的每張照片提取資訊，並輸出 JSON 陣列 (每張照片一個物件，嚴格依照傳入照片順序)：
- "filename": 照片檔名 (請對應傳入檔名)
- "project": 專案名稱 (造船專案代號，限定為 "J418" 或 "J419")
- "vendor": 申請廠商名稱 (例如: 啟進, 良擅, 日順安, 勛發, 佰億, 永一, 廣通, 安勝, 佳勝, 強聯, 吳勝泰, 放樣課 等)
- "location": 動火位置 (船段代號，例如: B8S, 2A2 P/S, B5M, 1A1, SK, 軸架, B1T 等)
- "worker": 動火人員 (重要規則：只需呈現一位姓名。若表單上有多位動火人員，請務必挑選其中一位『與監火人員不同姓名』的人員填入；若只有一位且與監火人員同名，才可同名)
- "fire_watcher": 監火人員 (監火員姓名)

重要辨識規則：
1. 不用辨識動火日期或作業內容。
2. 專案名稱若寫 418 請轉為 J418，若寫 419 請轉為 J419。
3. 動火人員欄位只能呈現一位姓名，若有多位請務必與監火人員不要是同一位。
4. 僅輸出純 JSON 陣列。"""

    config = types.GenerateContentConfig(
        response_mime_type="application/json",
        temperature=0.1,
    )

    all_results = []
    start_time = time.time()
    total_orig_kb = task_data.get("stats", {}).get("orig_kb", 0.0)
    total_comp_kb = task_data.get("stats", {}).get("comp_kb", 0.0)

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

        # 優先使用當前穩定的模型，若遇到限速或異常立即嘗試其他安全模型
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
                        # 依照照片檔名或順序建立對應
                        for idx_sub, item in enumerate(data_list):
                            sub_fname = item.get("filename")
                            if not sub_fname and idx_sub < len(batch_files_objs):
                                sub_fname = batch_files_objs[idx_sub]
                            batch_results.append({
                                "照片檔名": sub_fname or "未知檔名",
                                "專案名稱": normalize_project_name(
                                    item.get("project", item.get("專案名稱", ""))
                                ),
                                "申請廠商": item.get(
                                    "vendor", item.get("申請廠商", "")
                                ),
                                "動火位置": item.get(
                                    "location", item.get("動火位置", "")
                                ),
                                "動火人員": item.get(
                                    "worker", item.get("動火人員", "")
                                ),
                                "監火人員": item.get(
                                    "fire_watcher", item.get("監火人員", "")
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
                # 限速或忙碌時不浪費時間睡眠，直接換下一個備援模型重試
                continue

        # 若整批偶發失敗或有缺漏，自動使用單張補救
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

        # 更新即時進度
        task_data["progress_pct"] = end_idx / total_files
        task_data["results"] = all_results
        task_data["stats"]["elapsed"] = time.time() - start_time
        save_task_status(task_id, task_data)

        if b_idx < total_batches - 1 and inter_batch_delay > 0:
            time.sleep(inter_batch_delay)

    # 任務完成
    elapsed_total = time.time() - start_time
    task_data["status"] = "completed"
    task_data["msg"] = (
        f"🎉 辨識完成！共處理 {total_files} 張照片，耗時 {elapsed_total:.1f} 秒"
    )
    task_data["progress_pct"] = 1.0
    task_data["stats"]["elapsed"] = elapsed_total
    save_task_status(task_id, task_data)


def get_default_api_key() -> str:
    """自動從 Streamlit Secrets 或環境變數讀取金鑰，若無則採用預設金鑰。"""
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
    help="系統預設已填入您的專屬金鑰。亦可隨時手動更換其他 AI Studio 金鑰",
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
    help="已為您嚴選具備超高免費額度 (RPD/RPM) 且辨識僅需 1~3 秒的極速模型，徹底解決限速與逾時問題！",
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
# 📌 初始化手機友善之累計上傳暫存池 (Session State)
# ------------------------------------------------------------------
if "upload_buffer" not in st.session_state:
    st.session_state["upload_buffer"] = {}
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

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

    # 當選取相片時，即時執行極致輕量壓縮並加入暫存池
    if new_uploads:
        added_count = 0
        for f in new_uploads:
            if f.name not in st.session_state["upload_buffer"]:
                raw_bytes = f.getvalue()
                try:
                    c_bytes, o_kb, c_kb = compress_and_prep_image_bytes(
                        raw_bytes, max_dimension=max_img_dim, quality=70
                    )
                    st.session_state["upload_buffer"][f.name] = {
                        "bytes": c_bytes,
                        "size_kb": c_kb,
                        "orig_kb": o_kb,
                    }
                    added_count += 1
                except Exception as err:
                    st.error(f"⚠️ 照片 {f.name} 壓縮處理異常: {err}")
        if added_count > 0:
            st.toast(f"✅ 成功追加 {added_count} 張相片至待辨識清單！", icon="📸")

    buffer_count = len(st.session_state["upload_buffer"])

    # 暫存池控制列
    if buffer_count > 0:
        col_pool1, col_pool2 = st.columns([3, 1])
        with col_pool1:
            total_c_mb = sum(
                v["size_kb"] for v in st.session_state["upload_buffer"].values()
            ) / 1024.0
            total_o_mb = sum(
                v["orig_kb"] for v in st.session_state["upload_buffer"].values()
            ) / 1024.0
            st.success(
                f"📸 **待辨識相片池已累積：{buffer_count} 張相片** (原圖 {total_o_mb:.1f} MB ➔ 壓縮後僅 {total_c_mb:.2f} MB，極速傳輸！)"
            )
        with col_pool2:
            if st.button("🗑️ 清空暫存相片", use_container_width=True):
                st.session_state["upload_buffer"] = {}
                st.session_state["uploader_key"] += 1
                st.rerun()

        with st.expander(f"🖼️ 查看待辨識相片縮圖 ({buffer_count} 張)", expanded=False):
            cols = st.columns(4)
            for idx, (fn, info) in enumerate(st.session_state["upload_buffer"].items()):
                col = cols[idx % 4]
                col.image(
                    info["bytes"],
                    caption=f"{fn} ({info['size_kb']:.1f} KB)",
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
                img_folder = os.path.join(task_folder, "images")
                os.makedirs(img_folder, exist_ok=True)

                filenames = []
                sum_orig_kb = 0.0
                sum_comp_kb = 0.0

                for fname, info in st.session_state["upload_buffer"].items():
                    save_path = os.path.join(img_folder, fname)
                    with open(save_path, "wb") as f_out:
                        f_out.write(info["bytes"])
                    filenames.append(fname)
                    sum_orig_kb += info["orig_kb"]
                    sum_comp_kb += info["size_kb"]

                task_init_data = {
                    "task_id": task_id,
                    "status": "pending",
                    "progress_pct": 0.0,
                    "msg": "任務已建立，背景辨識線程啟動中...",
                    "total_files": len(filenames),
                    "filenames": filenames,
                    "results": [],
                    "stats": {
                        "orig_kb": sum_orig_kb,
                        "comp_kb": sum_comp_kb,
                        "orig_mb": sum_orig_kb / 1024.0,
                        "comp_mb": sum_comp_kb / 1024.0,
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

                # 清空暫存池以便下次使用
                st.session_state["upload_buffer"] = {}
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

        # 數據指標看板
        stats = curr_task.get("stats", {})
        if stats:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("總辨識照片", f"{curr_task.get('total_files', 0)} 張")
            c2.metric("總處理耗時", f"{stats.get('elapsed', 0):.1f} 秒")
            orig_mb = stats.get("orig_mb", 0)
            comp_mb = stats.get("comp_mb", 0)
            saved_mb = orig_mb - comp_mb
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

            # 確保欄位順序乾淨整齊
            columns_order = [
                "照片檔名",
                "專案名稱",
                "申請廠商",
                "動火位置",
                "動火人員",
                "監火人員",
            ]
            df = pd.DataFrame(results)
            # 過濾並依指定順序呈現
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