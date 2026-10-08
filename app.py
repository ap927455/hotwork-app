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
MY_API_KEY = "AQ.Ab8RN6LmKR0pdFqz-jagUYkq6ep-koHyQaZwef9LFcJm0EDk2A"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TASKS_DIR = os.path.join(BASE_DIR, "tasks_storage")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
os.makedirs(TASKS_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

# 頁面配置
st.set_page_config(
    page_title="廠區動火管制表手寫 AI 辨識與船段圖自動生成系統",
    page_icon="🔥",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🔥 造船廠動火管制表手寫 AI 辨識與船段圖自動生成系統")
st.caption(
    "⚡ 支援伺服器背景持續辨識與 **J418/J419 船段圖自動標示** (完美不擋船段名稱)"
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

UNIFIED_VENDOR_COLOR = (220, 53, 69)  # 統一廠商名稱標籤顏色 (標準醒目深紅)


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


def extract_permit_date(records: List[Dict[str, Any]]) -> str:
  """從辨識結果中自動擷取動火當天日期 (如 115 年 9 月 10 日)。"""
  for r in records:
    raw_d = str(r.get("動火日期", "")).strip()
    if raw_d:
      clean_d = (
          raw_d.replace("年", " 年 ")
          .replace("月", " 月 ")
          .replace("日", " 日")
          .replace("/", " 年 ", 1)
          .replace("/", " 月 ")
          .replace(".", " 年 ", 1)
          .replace(".", " 月 ")
      )
      if "年" in clean_d:
        return clean_d
      return f"115 年 {clean_d}"
  return time.strftime("115 年 %m 月 %d 日").replace(" 0", " ")


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
  tmpl_path = os.path.join(TEMPLATES_DIR, tmpl_name)

  if not os.path.exists(tmpl_path):
    base_img = Image.new("RGBA", (1024, 576), (245, 247, 250, 255))
    draw_err = ImageDraw.Draw(base_img)
    draw_err.rectangle([20, 20, 1004, 556], outline=(220, 53, 69, 255), width=3)
    draw_err.text(
        (250, 260),
        f"⚠️ 找不到船圖底圖檔案: templates/{tmpl_name}\n請確認已將 templates/ 資料夾及其底圖檔案 Push 到 GitHub 儲存庫中！",
        fill=(220, 53, 69, 255),
    )
  else:
    base_img = Image.open(tmpl_path).convert("RGBA")

  overlay = Image.new("RGBA", base_img.size, (255, 255, 255, 0))
  draw = ImageDraw.Draw(overlay)

  font_path = "C:/Windows/Fonts/msjhbd.ttc"
  if not os.path.exists(font_path):
    font_path = "C:/Windows/Fonts/msjh.ttc"

  try:
    font_vendor = ImageFont.truetype(font_path, 12)
    font_header = ImageFont.truetype(font_path, 18)
  except Exception:
    font_vendor = ImageFont.load_default()
    font_header = ImageFont.load_default()

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
# 📌 圖片壓縮處理
# ------------------------------------------------------------------
def compress_and_prep_image_bytes(
    file_bytes: bytes, max_dimension: int = 1024, quality: int = 80
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


def is_rate_limit_error(err: Exception) -> bool:
  err_str = str(err).lower()
  return any(
      kw in err_str
      for kw in [
          "429",
          "resource_exhausted",
          "quota_exhausted",
          "rate limit",
          "too many requests",
      ]
  )


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


# ------------------------------------------------------------------
# 📌 單張照片備援辨識 (當整批辨識失敗時觸發單張逐一降級重試)
# ------------------------------------------------------------------
def process_single_image_ocr(
    client: genai.Client,
    model_name: str,
    base_prompt: str,
    fpath: str,
    fname: str,
    max_img_dim: int,
    max_retries: int,
    task_id: str,
    task_data: dict,
) -> Tuple[Dict[str, Any], float, float]:
  """當批次 API 呼叫失敗時，針對單張照片進行獨立重試與降級解析。"""
  if not os.path.exists(fpath):
    return {
        "照片檔名": fname,
        "專案名稱": "辨識失敗",
        "申請廠商": "檔案不存在",
        "動火位置": "",
        "動火人員": "",
        "監火人員": "",
        "動火日期": "",
        "作業內容": "",
    }, 0.0, 0.0

  with open(fpath, "rb") as f_img:
    raw_b = f_img.read()
  c_bytes, o_kb, c_kb = compress_and_prep_image_bytes(
      raw_b, max_dimension=max_img_dim, quality=80
  )

  single_prompt = (
      base_prompt
      + f"\n\n目前輸入僅包含 1 張照片，檔名為【{fname}】。請輸出包含此單張照片資料的 JSON 陣列或物件。"
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

  base_backoff = 3.0
  for attempt in range(max_retries):
    try:
      response = client.models.generate_content(
          model=model_name, contents=contents, config=config
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

        res_dict = {
            "照片檔名": fname,
            "專案名稱": normalize_project_name(
                item.get("project", item.get("專案名稱", ""))
            ),
            "申請廠商": item.get("vendor", item.get("申請廠商", "")),
            "動火位置": item.get("location", item.get("動火位置", "")),
            "動火人員": item.get("worker", item.get("動火人員", "")),
            "監火人員": item.get("fire_watcher", item.get("監火人員", "")),
            "動火日期": item.get("work_date", item.get("動火日期", "")),
            "作業內容": item.get("work_content", item.get("作業內容", "")),
        }
        return res_dict, o_kb, c_kb
    except Exception as e:
      err_str = str(e).lower()
      if any(kw in err_str for kw in ["401", "unauthenticated", "invalid_argument", "access_token_type_unsupported", "permission_denied", "403"]):
        task_data["msg"] = f"❌ API Key 金鑰無效 (401 未授權): {e}"
        save_task_status(task_id, task_data)
        return {
            "照片檔名": fname,
            "專案名稱": "API Key無效",
            "申請廠商": "請更換金鑰 (401)",
            "動火位置": "",
            "動火人員": "",
            "監火人員": "",
            "動火日期": "",
            "作業內容": "",
        }, o_kb, c_kb
      wait_sec = min(
          12.0, 3.0 + attempt * 2.0 + random.uniform(0.5, 1.5)
      )
      task_data["msg"] = (
          f"⏳ [單張重試] {fname} 觸發 API 限速 (第"
          f" {attempt + 1}/{max_retries} 次)，自動冷卻 {int(wait_sec)} 秒..."
      )
      save_task_status(task_id, task_data)
      time.sleep(wait_sec)

  return {
      "照片檔名": fname,
      "專案名稱": "辨識失敗",
      "申請廠商": "連線或流量限速失敗",
      "動火位置": "",
      "動火人員": "",
      "監火人員": "",
      "動火日期": "",
      "作業內容": "",
  }, o_kb, c_kb


# ------------------------------------------------------------------
# 📌 後台獨立線程 worker：負責執行 Gemini API 辨識
# ------------------------------------------------------------------
def run_background_ocr_task(
    task_id: str,
    key_str: str,
    model_name: str,
    batch_size: int,
    inter_batch_delay: int,
    max_img_dim: int,
    max_retries: int,
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
  prompt = """
你是造船廠安全管理專家。請仔細辨識傳入的照片中『動火工作安全許可申請表』的手寫文字。
請為傳入的每一張照片提取資訊，並嚴格輸出 JSON 陣列 (JSON Array)。

欄位辨識規則 (請嚴格遵守)：
- "filename": 照片檔名 (請精確對應輸入照片所標註的檔名)
- "project": 專案名稱 (重要：專案名稱只會有 "J418" 或 "J419"。請辨識並精確歸類為 "J418" 或 "J419"，不可輸出其他名稱)
- "vendor": 申請/施工廠商名稱 (如 啟進, 良擅, 日順安, 勛發, 佰億, 永一, 廣通 等)
- "location": 動火作業位置 (地點/區域/船艙如 B8S, 2A2 P/S, B5M, 1A1 等)
- "worker": 動火作業人員姓名 (動火員/施工人員)
- "fire_watcher": 監火人員姓名 (監火員)
- "work_date": 動火日期或時間 (如 115年9月10日 或 2026/09/10，若表格未填寫請給空字串 "")
- "work_content": 作業內容與方式 (重要：包含「整形」、「瓦斯切割」、「管路安裝」、「電焊」、「氣割」、「打磨」、「裝配」、「冷作」等所有作業內容或施工方式文字，務必歸在「作業內容」欄位)

請依照圖片順序輸出。字跡潦草時請依上下文合理推測，僅輸出純 JSON 資料。
"""
  config = types.GenerateContentConfig(
      response_mime_type="application/json",
      temperature=0.1,
  )

  all_results = []
  start_time = time.time()
  total_orig_kb = 0.0
  total_comp_kb = 0.0

  task_data["status"] = "processing"
  task_data["progress_pct"] = 0.0
  save_task_status(task_id, task_data)

  current_inter_batch_delay = inter_batch_delay

  for b_idx in range(total_batches):
    start_idx = b_idx * batch_size
    end_idx = min(start_idx + batch_size, total_files)
    batch_filenames = filenames[start_idx:end_idx]

    task_data["msg"] = (
        f"⚡ 正在分析第 {start_idx + 1} ~ {end_idx} / {total_files} 張照片 (批次"
        f" {b_idx + 1}/{total_batches})..."
    )
    save_task_status(task_id, task_data)

    contents = [prompt]
    batch_files_objs = []

    for fname in batch_filenames:
      fpath = os.path.join(img_folder, fname)
      if os.path.exists(fpath):
        with open(fpath, "rb") as f_img:
          raw_b = f_img.read()
        c_bytes, o_kb, c_kb = compress_and_prep_image_bytes(
            raw_b, max_dimension=max_img_dim, quality=80
        )
        total_orig_kb += o_kb
        total_comp_kb += c_kb
        contents.append(f"【照片檔名: {fname}】")
        contents.append(
            types.Part.from_bytes(data=c_bytes, mime_type="image/jpeg")
        )
        batch_files_objs.append(fname)

    # Gemini 呼叫與重試 (涵蓋 429 / 500 / 503 / 網路波動)
    batch_results = []
    base_backoff = 3.0
    batch_success = False

    for attempt in range(max_retries):
      try:
        response = client.models.generate_content(
            model=model_name, contents=contents, config=config
        )
        if response and response.text:
          clean_text = (
              response.text.replace("```json", "").replace("```", "").strip()
          )
          data_list = json.loads(clean_text)
          if isinstance(data_list, list):
            for item in data_list:
              batch_results.append({
                  "照片檔名": item.get(
                      "filename", item.get("照片檔名", "未知檔名")
                  ),
                  "專案名稱": normalize_project_name(
                      item.get("project", item.get("專案名稱", ""))
                  ),
                  "申請廠商": item.get("vendor", item.get("申請廠商", "")),
                  "動火位置": item.get("location", item.get("動火位置", "")),
                  "動火人員": item.get("worker", item.get("動火人員", "")),
                  "監火人員": item.get(
                      "fire_watcher", item.get("監火人員", "")
                  ),
                  "動火日期": item.get("work_date", item.get("動火日期", "")),
                  "作業內容": item.get(
                      "work_content", item.get("作業內容", "")
                  ),
              })
            batch_success = True
            break
      except Exception as e:
        err_str = str(e).lower()
        if any(kw in err_str for kw in ["401", "unauthenticated", "invalid_argument", "access_token_type_unsupported", "permission_denied", "403"]):
          task_data["status"] = "error"
          task_data["msg"] = f"❌ API Key 金鑰無效或未授權 (401 未授權)！請檢查 API Key 金鑰或權限設定: {e}"
          save_task_status(task_id, task_data)
          return
        # 動態增加批次冷卻時間以保護限速
        current_inter_batch_delay = min(15, current_inter_batch_delay + 1)
        wait_sec = min(
            15.0, 4.0 + attempt * 2.5 + random.uniform(0.5, 1.5)
        )
        task_data["msg"] = (
            f"⏳ 第 {start_idx + 1}~{end_idx} 張照片觸發 API 限速/網路問題"
            f" (第 {attempt + 1}/{max_retries} 次)，自動冷卻 {int(wait_sec)} 秒..."
        )
        save_task_status(task_id, task_data)
        time.sleep(wait_sec)

    # 若批次辨識失敗，自動啟動「單張降級備援機制」
    if not batch_success:
      task_data["msg"] = (
          f"⚠️ 第 {start_idx + 1}~{end_idx} 張批次辨識失敗，自動切換至「單張逐一降級辨識」..."
      )
      save_task_status(task_id, task_data)
      batch_results = []
      for fname in batch_files_objs:
        fpath = os.path.join(img_folder, fname)
        res_single, s_okb, s_ckb = process_single_image_ocr(
            client=client,
            model_name=model_name,
            base_prompt=prompt,
            fpath=fpath,
            fname=fname,
            max_img_dim=max_img_dim,
            max_retries=max_retries,
            task_id=task_id,
            task_data=task_data,
        )
        batch_results.append(res_single)
        time.sleep(2)

    all_results.extend(batch_results)

    # 更新進度
    task_data["progress_pct"] = end_idx / total_files
    task_data["results"] = all_results
    task_data["stats"] = {
        "orig_mb": total_orig_kb / 1024.0,
        "comp_mb": total_comp_kb / 1024.0,
        "elapsed": time.time() - start_time,
    }
    save_task_status(task_id, task_data)

    if b_idx < total_batches - 1 and current_inter_batch_delay > 0:
      time.sleep(current_inter_batch_delay)

  # 完成任務
  task_data["status"] = "completed"
  task_data["msg"] = (
      f"🎉 辨識完成！共處理 {total_files} 張照片，耗時"
      f" {time.time() - start_time:.1f} 秒"
  )
  task_data["progress_pct"] = 1.0
  save_task_status(task_id, task_data)


def get_default_api_key() -> str:
  """自動從 Streamlit Secrets 或環境變數讀取金鑰，支援雲端部署安全與方便性。"""
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
    help="自動帶入 Secrets/環境變數金鑰。亦可手動輸入 AI Studio 產生的金鑰",
)

model_choice = st.sidebar.selectbox(
    "AI 模型選擇:",
    ["gemini-3.6-flash", "gemini-3.8-flash", "gemini-flash-latest"],
    index=0,
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 效能與流量優化參數")

batch_size = st.sidebar.slider("📦 批次打包數量 (張/次):", 1, 6, 2)
inter_batch_delay = st.sidebar.slider("⏱️ 批次間隔冷卻 (秒):", 0, 15, 5)
max_img_dim = st.sidebar.select_slider(
    "📐 圖片極限解析度 (px):", options=[800, 1024, 1280, 1600], value=1024
)
max_retries = st.sidebar.number_input(
    "🔄 429 觸發最高重試次數:", 1, 10, 6
)

if st.sidebar.button("🔍 測試 API 金鑰連線"):
  if not api_key_input.strip() or "貼上你的" in api_key_input:
    st.sidebar.error("❌ 請先填入正確的 API Key！")
  else:
    try:
      client = genai.Client(api_key=api_key_input.strip())
      res = client.models.generate_content(model=model_choice, contents="Ping")
      if res and res.text:
        st.sidebar.success("✅ API 連線成功！模型運作正常")
      else:
        st.sidebar.error("❌ API 無回應")
    except Exception as e:
      st.sidebar.error(f"❌ 連線失敗: {e}")

# ------------------------------------------------------------------
# 📌 主畫面 Tabs 分頁設計
# ------------------------------------------------------------------
tab_upload, tab_history = st.tabs(
    ["🚀 建立新辨識任務", "📋 檢視背景辨識任務與歷史紀錄"]
)

with tab_upload:
  uploaded_files = st.file_uploader(
      "📁 請選擇或拖曳動火單相片 (支援全選照片，相容 HEIC/JPG/PNG):",
      accept_multiple_files=True,
  )

  if uploaded_files:
    st.info(f"📸 已選取 **{len(uploaded_files)}** 張照片檔案")

    with st.expander("🖼️ 檢視已上傳相片縮圖預覽"):
      cols = st.columns(4)
      for idx, f in enumerate(uploaded_files):
        col = cols[idx % 4]
        try:
          img_preview = Image.open(f)
          img_preview = ImageOps.exif_transpose(img_preview)
          col.image(
              img_preview,
              caption=f"{f.name} ({f.size / 1024:.1f} KB)",
              use_container_width=True,
          )
        except Exception:
          col.write(f"📄 {f.name} ({f.size / 1024:.1f} KB)")

    if st.button("🚀 啟動伺服器背景 AI 自動辨識"):
      if not api_key_input.strip():
        st.error("❌ 請先填寫有效的 API Key！")
      else:
        task_id = time.strftime("task_%Y%m%d_%H%M%S")
        task_folder = get_task_folder(task_id)
        img_folder = os.path.join(task_folder, "images")
        os.makedirs(img_folder, exist_ok=True)

        filenames = []
        for f in uploaded_files:
          save_path = os.path.join(img_folder, f.name)
          file_bytes = f.getvalue()
          # ✨ 上傳時立即執行超輕量化預壓縮 (解決 Streamlit Cloud 記憶體/CPU 瓶頸)
          try:
            comp_bytes, _, _ = compress_and_prep_image_bytes(
                file_bytes, max_dimension=max_img_dim, quality=80
            )
            with open(save_path, "wb") as f_out:
              f_out.write(comp_bytes)
          except Exception:
            with open(save_path, "wb") as f_out:
              f_out.write(file_bytes)
          filenames.append(f.name)

        task_init_data = {
            "task_id": task_id,
            "status": "pending",
            "progress_pct": 0.0,
            "msg": "任務已建立，等待伺服器背景處理...",
            "total_files": len(filenames),
            "filenames": filenames,
            "results": [],
            "stats": {},
            "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        }
        save_task_status(task_id, task_init_data)

        # 啟動獨立背景線程
        t = threading.Thread(
            target=run_background_ocr_task,
            kwargs={
                "task_id": task_id,
                "key_str": api_key_input,
                "model_name": model_choice,
                "batch_size": batch_size,
                "inter_batch_delay": inter_batch_delay,
                "max_img_dim": max_img_dim,
                "max_retries": max_retries,
            },
            daemon=True,
        )
        t.start()

        st.session_state["active_task_id"] = task_id
        st.success(f"🎉 任務已於伺服器背景啟動！(任務編號: `{task_id}`)")
        st.info(
            "💡 **手機貼心提示**：您現在可以**切換到其他 App、切換網頁或鎖定手機螢幕**，伺服器會持續辨識。稍後返回『檢視背景辨識任務』分頁即可查看結果！"
        )

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
        st.warning(f"⚠️ 確定要刪除任務 `{selected_task_id}` 及其所有照片與辨識結果嗎？")
        if st.button("🚨 確定刪除此任務", type="primary", use_container_width=True):
          delete_task(selected_task_id)
          if st.session_state.get("active_task_id") == selected_task_id:
            st.session_state.pop("active_task_id", None)
          st.success(f"🗑️ 已成功刪除任務 `{selected_task_id}`！")
          time.sleep(0.5)
          st.rerun()

    with col_btn3:
      with st.popover("🧹 清空所有歷史任務"):
        st.warning("⚠️ 確定要刪除伺服器上的所有歷史辨識紀錄與圖檔嗎？此動作無法復原！")
        if st.button("🚨 確定全部清空", type="primary", use_container_width=True):
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

    # 數據看板
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
          "圖片壓縮節省流量", f"{saved_mb:.2f} MB", delta=f"-{saved_pct:.1f}%"
      )
      c4.metric("壓縮後傳輸大小", f"{comp_mb:.2f} MB")

    results = curr_task.get("results", [])
    if results:
      st.write("---")
      st.write(
          "### 📊 今日動火作業彙整表 (雙擊儲存格即可修改校對，上方船段圖會即時同步更新)："
      )

      df = pd.DataFrame(results)
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
            "管制表右上角日期 (預設帶入動火單當天日期):",
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

