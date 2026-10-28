#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Thai ID Decoder — LINE Bot Edition
พิมพ์ "!id <เลขบัตร>" หรือพิมพ์เลขบัตร 13 หลักตรง ๆ ในแชทกับ LINE Official Account
รันด้วย Flask รับ webhook จาก LINE Messaging API
"""

import os
import io
import re
import json
import zipfile
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional, Dict
from datetime import datetime

from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError, LineBotApiError
from linebot.models import MessageEvent, TextMessage, TextSendMessage

# ============================================================
# ⚙️ CONFIGURATION (แก้ตรงนี้ หรือจะตั้งผ่าน environment variable ก็ได้)
# ============================================================

LINE_CHANNEL_ACCESS_TOKEN = os.environ.get(
    "LINE_CHANNEL_ACCESS_TOKEN",
    "cgTWaHGdOoyTc1a1fF9O2vRKyswwcfeT51xD3SDGIyZCQDMCgxhWYqC9OPSPqOaB4tO+l5eIf0eDfZ/jxp1rPTYng/kBO3Mb6izxRxd1xqPUTHiwhdGqE0h1lR6cbIOG6acshygU33CpoEvkcRkPTgdB04t89/1O/w1cDnyilFU=",
)
LINE_CHANNEL_SECRET = os.environ.get(
    "LINE_CHANNEL_SECRET",
    "0be217429b92f324572296d63430e052",
)

# ถ้าอยากจำกัดให้ใช้ได้เฉพาะ LINE user ID ของคุณ ใส่ user id ตรงนี้ (หาได้จาก event.source.user_id ที่ print ตอนทดสอบ)
# ปล่อยเป็นสตริงว่าง "" = ไม่ล็อค ใครทักมาก็ใช้คำสั่งได้หมด
OWNER_USER_ID = ""

PLACEHOLDER_TOKENS = {"", "ใส่_channel_access_token_ตรงนี้", "ใส่_channel_secret_ตรงนี้"}
# ⚠️ อย่าใส่ token/secret จริงลงในเซ็ตนี้ ไม่งั้นโปรแกรมจะคิดว่ามันเป็น placeholder แล้วปฏิเสธไม่รัน

# ============================================================
# CONSTANTS
# ============================================================
CACHE_FILE = "rcode.json"
SOURCE_XLSX_URL = "https://stat.bora.dopa.go.th/dload/rcode.xlsx"
SOURCE_PAGE_URL = "https://stat.bora.dopa.go.th/stat/statnew/statMenu/newStat/ccaa.php"
APP_VERSION = "1.0.0-line"

# ============================================================
# RCODE LOADER  (เหมือนเวอร์ชัน Discord ทุกอย่าง)
# ============================================================

def load_rcode_cache() -> Optional[Dict]:
    cache_path = Path(CACHE_FILE)
    if cache_path.exists():
        try:
            with open(cache_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                count = len(data.get("registries", {}))
                print(f"[RCODE] โหลดจาก cache: {count} รหัส")
                return data
        except Exception as e:
            print(f"[RCODE] cache error: {e}")
    return None


def save_rcode_cache(data: Dict):
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        count = len(data.get("registries", {}))
        print(f"[RCODE] บันทึก cache: {count} รหัส")
    except Exception as e:
        print(f"[RCODE] save cache error: {e}")


def parse_xlsx(content: bytes) -> Dict:
    registries = {}
    ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as zf:
            shared_strings = []
            try:
                ss_xml = zf.read("xl/sharedStrings.xml")
                ss_root = ET.fromstring(ss_xml)
                for si in ss_root.findall(".//main:si", ns):
                    text = "".join(t.text or "" for t in si.findall(".//main:t", ns))
                    shared_strings.append(text)
            except Exception:
                pass

            try:
                wb_xml = zf.read("xl/workbook.xml")
                wb_root = ET.fromstring(wb_xml)
                sheets = wb_root.findall(".//main:sheet", ns)
                if not sheets:
                    return registries
                rel_id = sheets[0].attrib.get(
                    "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "rId1"
                )
                sheet_num = rel_id.replace("rId", "")
                sheet_path = f"xl/worksheets/sheet{sheet_num}.xml"
            except Exception:
                sheet_path = "xl/worksheets/sheet1.xml"

            try:
                sheet_xml = zf.read(sheet_path)
                sheet_root = ET.fromstring(sheet_xml)
            except Exception:
                for i in range(1, 10):
                    try:
                        sheet_path = f"xl/worksheets/sheet{i}.xml"
                        sheet_xml = zf.read(sheet_path)
                        sheet_root = ET.fromstring(sheet_xml)
                        break
                    except Exception:
                        continue
                else:
                    return registries

            rows = sheet_root.findall(".//main:row", ns)
            for row in rows:
                cells = row.findall(".//main:c", ns)
                if len(cells) < 2:
                    continue

                code_val = None
                for child in cells[0]:
                    if child.tag.endswith("v"):
                        code_val = child.text
                        break
                if not code_val:
                    continue

                code = code_val.strip()
                if len(code) < 4 or not code.isdigit():
                    continue

                name_val = None
                for child in cells[1]:
                    if child.tag.endswith("v"):
                        name_val = child.text
                        if cells[1].attrib.get("t") == "s" and name_val:
                            try:
                                name_val = shared_strings[int(name_val)]
                            except Exception:
                                pass
                        break
                if not name_val:
                    continue

                name = name_val.strip()
                if not name:
                    continue

                status = "ใช้งาน"
                if len(cells) > 2:
                    status_val = None
                    for child in cells[2]:
                        if child.tag.endswith("v"):
                            status_val = child.text
                            if cells[2].attrib.get("t") == "s" and status_val:
                                try:
                                    status_val = shared_strings[int(status_val)]
                                except Exception:
                                    pass
                            break
                    if status_val:
                        status = status_val.strip()

                if code.endswith("00"):
                    prov_name = name.replace("จังหวัด", "").strip() or name
                    registries[code] = {
                        "name": name,
                        "subdistrict": prov_name,
                        "district": prov_name,
                        "province": prov_name,
                        "type": "จังหวัด",
                        "status": status,
                    }
                else:
                    prov_code = code[:2]
                    province_name = "ไม่ทราบ"
                    for rcode, info in registries.items():
                        if rcode.startswith(prov_code) and rcode.endswith("00"):
                            province_name = info.get("province", "ไม่ทราบ")
                            break

                    is_cancelled = name.startswith("*")
                    clean_name = name.lstrip("*").strip()

                    if "อำเภอ" in clean_name:
                        district_name = clean_name.replace("อำเภอ", "").strip()
                        office_type = "อำเภอ"
                    elif "สำนักงานเขต" in clean_name:
                        district_name = clean_name.replace("สำนักงานเขต", "").strip()
                        office_type = "เขต"
                    elif "เขต" in clean_name:
                        district_name = clean_name.replace("เขต", "").strip()
                        office_type = "เขต"
                    elif "กิ่งอำเภอ" in clean_name:
                        district_name = clean_name.replace("กิ่งอำเภอ", "").strip()
                        office_type = "กิ่งอำเภอ"
                    elif "เทศบาล" in clean_name:
                        district_name = clean_name
                        office_type = "เทศบาล"
                    else:
                        district_name = clean_name
                        office_type = "สำนักทะเบียน"

                    if not code.endswith("00"):
                        amp_code = code[:4] + "00"
                        amp_info = registries.get(amp_code, {})
                        if amp_info:
                            district_name = amp_info.get("district", district_name)
                            office_type = amp_info.get("type", office_type)

                    registries[code] = {
                        "name": clean_name,
                        "subdistrict": clean_name,
                        "district": district_name,
                        "province": province_name,
                        "type": office_type,
                        "status": "ยกเลิก" if is_cancelled or "ยกเลิก" in status else "ใช้งาน",
                    }
            return registries
    except Exception as e:
        print(f"[RCODE] parse xlsx error: {e}")
        return {}


def fetch_rcode_from_web() -> Optional[Dict]:
    print("[RCODE] กำลังดาวน์โหลดจาก DOPA...")
    try:
        req = urllib.request.Request(
            SOURCE_XLSX_URL,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            },
        )
        with urllib.request.urlopen(req, timeout=30) as response:
            content = response.read()
            if len(content) < 100 or content[:4] != b"PK\x03\x04":
                print("[RCODE] ไฟล์ที่ได้ไม่ใช่ XLSX")
                return None
            registries = parse_xlsx(content)
            if registries and len(registries) > 100:
                data = {
                    "registries": registries,
                    "metadata": {
                        "updated_at": datetime.now().isoformat(),
                        "record_count": len(registries),
                        "source_url": SOURCE_XLSX_URL,
                        "source_page": SOURCE_PAGE_URL,
                    },
                }
                save_rcode_cache(data)
                print(f"[RCODE] อัปเดตสำเร็จ: {len(registries)} รหัส")
                return data
            print(f"[RCODE] พบเพียง {len(registries)} รหัส")
            return None
    except urllib.error.URLError as e:
        print(f"[RCODE] URLError: {e}")
    except Exception as e:
        print(f"[RCODE] error: {e}")
    return None


def create_fallback_data() -> Dict:
    provinces = {
        "10": "กรุงเทพมหานคร", "11": "สมุทรปราการ", "12": "นนทบุรี", "13": "ปทุมธานี",
        "14": "พระนครศรีอยุธยา", "15": "อ่างทอง", "16": "ลพบุรี", "17": "สิงห์บุรี",
        "18": "ชัยนาท", "19": "สระบุรี", "20": "ชลบุรี", "21": "ระยอง",
        "22": "จันทบุรี", "23": "ตราด", "24": "ฉะเชิงเทรา", "25": "ปราจีนบุรี",
        "26": "นครนายก", "27": "สระแก้ว", "30": "นครราชสีมา", "31": "บุรีรัมย์",
        "32": "สุรินทร์", "33": "ศรีสะเกษ", "34": "อุบลราชธานี", "35": "ยโสธร",
        "36": "ชัยภูมิ", "37": "อำนาจเจริญ", "38": "บึงกาฬ", "39": "หนองบัวลำภู",
        "40": "ขอนแก่น", "41": "อุดรธานี", "42": "เลย", "43": "หนองคาย",
        "44": "มหาสารคาม", "45": "ร้อยเอ็ด", "46": "กาฬสินธุ์", "47": "สกลนคร",
        "48": "นครพนม", "49": "มุกดาหาร", "50": "เชียงใหม่", "51": "ลำพูน",
        "52": "ลำปาง", "53": "อุตรดิตถ์", "54": "แพร่", "55": "น่าน",
        "56": "พะเยา", "57": "เชียงราย", "58": "แม่ฮ่องสอน", "60": "นครสวรรค์",
        "61": "อุทัยธานี", "62": "กำแพงเพชร", "63": "ตาก", "64": "สุโขทัย",
        "65": "พิษณุโลก", "66": "พิจิตร", "67": "เพชรบูรณ์", "70": "ราชบุรี",
        "71": "กาญจนบุรี", "72": "สุพรรณบุรี", "73": "นครปฐม", "74": "สมุทรสาคร",
        "75": "สมุทรสงคราม", "76": "เพชรบุรี", "77": "ประจวบคีรีขันธ์",
        "80": "นครศรีธรรมราช", "81": "กระบี่", "82": "พังงา", "83": "ภูเก็ต",
        "84": "สุราษฎร์ธานี", "85": "ระนอง", "86": "ชุมพร", "90": "สงขลา",
        "91": "สตูล", "92": "ตรัง", "93": "พัทลุง", "94": "ปัตตานี",
        "95": "ยะลา", "96": "นราธิวาส",
    }
    registries = {}
    for code, name in provinces.items():
        registries[code + "00"] = {
            "name": f"จังหวัด{name}",
            "subdistrict": name,
            "district": name,
            "province": name,
            "type": "จังหวัด",
            "status": "ใช้งาน",
        }
    return {
        "registries": registries,
        "metadata": {
            "updated_at": datetime.now().isoformat(),
            "record_count": len(registries),
            "note": "ข้อมูลพื้นฐาน (เฉพาะจังหวัด)",
        },
    }


RCODE_DB: Dict = {}


def init_rcode():
    global RCODE_DB
    data = load_rcode_cache()
    if data:
        RCODE_DB = data
        return
    print("[RCODE] cache ไม่มี ลองดาวน์โหลด...")
    data = fetch_rcode_from_web()
    if data:
        RCODE_DB = data
        return
    print("[RCODE] ใช้ fallback")
    RCODE_DB = create_fallback_data()


# ============================================================
# ID DECODER  (เหมือนเวอร์ชัน Discord ทุกอย่าง)
# ============================================================

PERSON_TYPES = {
    0: {"title": "ไม่มีสถานะทางทะเบียน", "detail": "บุคคลที่ไม่มีสถานะทางทะเบียนในระบบ", "usage": "กลุ่มบุคคลที่ได้รับการสำรวจหรือจัดทำทะเบียนประวัติ"},
    1: {"title": "ผู้เกิดและได้สัญชาติไทย (แจ้งเกิดภายในกำหนด)", "detail": "แจ้งเกิดภายใน 15 วัน", "usage": "แจ้งเกิดตามกำหนดเวลากฎหมาย"},
    2: {"title": "ผู้เกิดและได้สัญชาติไทย (แจ้งเกิดเกินกำหนด)", "detail": "แจ้งเกิดหลัง 15 วัน", "usage": "แจ้งเกิดล่าช้า"},
    3: {"title": "มีชื่อในทะเบียนบ้านช่วงเริ่มระบบเลข 13 หลัก", "detail": "อยู่ในทะเบียนบ้านเดิมช่วงเริ่มระบบใหม่", "usage": "คนไทยที่เกิดก่อนระบบใหม่"},
    4: {"title": "บุคคลในทะเบียนเดิมที่ได้รับเลขภายหลัง", "detail": "มีทะเบียนเดิมแต่ยังไม่มีเลข", "usage": "ได้รับเลขภายหลัง"},
    5: {"title": "คนไทยที่เพิ่มชื่อในทะเบียนบ้านภายหลัง", "detail": "เพิ่มชื่อภายหลัง", "usage": "ตกสำรวจ / กรณีทางทะเบียน"},
    6: {"title": "บุคคลต่างด้าวหรือผู้มีสถานะการอยู่บางประเภท", "detail": "ผู้เข้าเมืองหรือผู้มีสถานะบางกลุ่ม", "usage": "ต่างด้าวที่ได้รับเลขประจำตัว"},
    7: {"title": "บุตรของบุคคลประเภท 6 ที่เกิดในประเทศไทย", "detail": "บุตรต่างด้าวที่เกิดในไทย", "usage": "บุตรที่เกิดในไทยของบุคคลประเภท 6"},
    8: {"title": "บุคคลที่ได้รับสถานะหรือสัญชาติภายหลัง", "detail": "ผู้มีถิ่นที่อยู่ / ผู้แปลงสัญชาติ", "usage": "ได้รับสัญชาติหรือสถานะภายหลัง"},
    9: {"title": "กลุ่มสำรอง (ไม่พบข้อมูล)", "detail": "หลัก 9 ยังไม่ใช้งาน", "usage": "รอการกำหนด"},
}


def normalize_id(raw: str) -> str:
    value = re.sub(r"[^\d]", "", raw.strip())
    if not value:
        raise ValueError("ไม่ได้กรอกเลขบัตร")
    if len(value) != 13:
        raise ValueError(f"เลขบัตรต้องมี 13 หลัก (ได้ {len(value)} หลัก)")
    return value


def format_id(pid: str) -> str:
    return f"{pid[0]}-{pid[1:5]}-{pid[5:10]}-{pid[10:12]}-{pid[12]}"


def calculate_checksum(pid: str) -> Dict:
    total = 0
    for i, ch in enumerate(pid[:12]):
        digit = int(ch)
        weight = 13 - i
        total += digit * weight
    remainder = total % 11
    expected = (11 - remainder) % 10
    actual = int(pid[12])
    return {"actual": actual, "expected": expected, "valid": actual == expected, "total": total, "remainder": remainder}


def decode_id(raw: str) -> Dict:
    try:
        pid = normalize_id(raw)
    except ValueError as e:
        return {"error": str(e)}
    checksum = calculate_checksum(pid)
    p_type = int(pid[0])
    person_type = PERSON_TYPES.get(p_type, {"title": "ไม่ทราบประเภท", "detail": "-", "usage": "-"})
    rcode = pid[1:5]
    reg = RCODE_DB.get("registries", {}).get(rcode, {})
    return {
        "formatted": format_id(pid),
        "valid": checksum["valid"],
        "person_type": {"code": p_type, **person_type},
        "registry": {
            "code": rcode,
            "name": reg.get("name", "ไม่พบข้อมูล"),
            "subdistrict": reg.get("subdistrict", "ไม่พบข้อมูล"),
            "district": reg.get("district", "ไม่พบข้อมูล"),
            "province": reg.get("province", "ไม่พบข้อมูล"),
            "type": reg.get("type", "สำนักทะเบียน"),
            "status": reg.get("status", "ไม่ทราบ"),
        },
        "serial": {"group": pid[5:10], "sequence": pid[10:12]},
        "checksum": checksum,
    }


def build_reply_text(result: Dict) -> str:
    """ประกอบผลลัพธ์เป็นข้อความธรรมดา (LINE text message ไม่มี embed แบบ Discord)"""
    if result.get("error"):
        return f"❌ {result['error']}"

    valid = result["valid"]
    status_line = "✅ ผ่านการตรวจสอบ" if valid else "❌ ไม่ผ่าน Checksum"
    pt = result["person_type"]
    reg = result["registry"]
    serial = result["serial"]
    checksum = result["checksum"]

    lines = [
        "🏛️ Thai ID Decoder",
        f"เลขบัตร: {result['formatted']}",
        f"สถานะ: {status_line}",
        f"Checksum: หลักสุดท้าย {checksum['actual']} / ควรเป็น {checksum['expected']}",
        "",
        f"👤 ประเภทบุคคล (หลัก 1: {pt['code']})",
        f"{pt['title']}",
        f"{pt['detail']}",
    ]

    lines.append("")
    if reg["name"] != "ไม่พบข้อมูล":
        lines += [
            "🏛️ สำนักทะเบียน",
            f"รหัส RCODE: {reg['code']}",
            f"ชื่อ: {reg['name']}",
            f"ประเภท: {reg['type']} | สถานะ: {reg['status']}",
            f"ตำบล/แขวง: {reg['subdistrict']}",
            f"อำเภอ/เขต: {reg['district']}",
            f"จังหวัด: {reg['province']}",
        ]
    else:
        lines += [f"🏛️ รหัส RCODE: {reg['code']} — ไม่พบข้อมูลในฐานข้อมูล"]

    lines += ["", f"📂 กลุ่มทะเบียน: {serial['group']} | ลำดับในกลุ่ม: {serial['sequence']}"]

    reg_meta = RCODE_DB.get("metadata", {})
    lines += ["", f"— v{APP_VERSION} | RCODE {reg_meta.get('record_count', '?')} รหัส"]

    return "\n".join(lines)


HELP_TEXT = (
    "📖 วิธีใช้\n"
    "พิมพ์: !id <เลขบัตรประชาชน 13 หลัก>\n"
    "ตัวอย่าง: !id 1234567890123\n"
    "หรือพิมพ์เลข 13 หลักตรง ๆ ก็ได้\n\n"
    "• ตรวจ Checksum ตามอัลกอริทึมกระทรวงมหาดไทย\n"
    "• ถอดรหัสสำนักทะเบียน (RCODE) จากฐาน DOPA"
)

# ============================================================
# LINE BOT (Flask webhook)
# ============================================================

app = Flask(__name__)
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)


def extract_id_query(text: str) -> Optional[str]:
    """รองรับทั้ง '!id 1234567890123' และพิมพ์เลข 13 หลักตรง ๆ"""
    text = text.strip()
    if text.startswith("!id"):
        return text[3:].strip()
    digits_only = re.sub(r"[^\d]", "", text)
    if len(digits_only) == 13:
        return digits_only
    return None


@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        print("[LINE] Invalid signature — ตรวจ LINE_CHANNEL_SECRET")
        abort(400)
    return "OK"


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event: MessageEvent):
    user_text = event.message.text or ""
    user_id = event.source.user_id if event.source else None
    print(f"[LINE] from={user_id} text={user_text!r}")

    if OWNER_USER_ID and user_id != OWNER_USER_ID:
        return  # เงียบไว้ ไม่ตอบถ้าไม่ใช่เจ้าของ (ล็อคไว้)

    if user_text.strip() in ("!id", "!id "):
        reply_text = HELP_TEXT
    else:
        query = extract_id_query(user_text)
        if query is None:
            return  # ไม่ใช่คำสั่งที่เกี่ยวข้อง ไม่ต้องตอบ
        result = decode_id(query)
        reply_text = build_reply_text(result)

    try:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))
    except LineBotApiError as e:
        print(f"[LINE] reply error: {e}")


@app.route("/health", methods=["GET"])
def health():
    reg_count = len(RCODE_DB.get("registries", {})) if RCODE_DB else 0
    return {"status": "ok", "version": APP_VERSION, "rcode_records": reg_count}


# ============================================================
# MAIN
# ============================================================

def main():
    if LINE_CHANNEL_ACCESS_TOKEN in PLACEHOLDER_TOKENS or LINE_CHANNEL_SECRET in PLACEHOLDER_TOKENS:
        print("❌ ยังไม่ได้ใส่ LINE_CHANNEL_ACCESS_TOKEN / LINE_CHANNEL_SECRET (หรือตั้งผ่าน env var)")
        return
    init_rcode()
    print("=" * 60)
    print(f"[BOT] Thai ID Decoder (LINE) v{APP_VERSION}")
    print(f"[BOT] Owner lock: {OWNER_USER_ID or 'OFF'}")
    reg_count = len(RCODE_DB.get("registries", {})) if RCODE_DB else 0
    print(f"[BOT] RCODE records: {reg_count}")
    print("[BOT] Webhook endpoint: POST /callback")
    print("=" * 60)
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)


if __name__ == "__main__":
    main()
