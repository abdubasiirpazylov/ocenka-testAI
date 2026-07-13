import streamlit as st
from docxtpl import DocxTemplate, RichText
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import os
import re 
import json
from num2words import num2words
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import requests 

import docx
import copy

try:
    import google.generativeai as genai
    from PIL import Image
    HAS_AI = True
except ImportError:
    HAS_AI = False

TEMPLATE_NAME = "образец отчета.docx"

st.set_page_config(page_title="Генератор Отчетов - Гарант Оценка", layout="wide")

if HAS_AI and "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    AI_READY = True
else:
    AI_READY = False

KG_REGIONS = {
    "г. Бишкек": ["Ленинский район", "Октябрьский район", "Первомайский район", "Свердловский район"],
    "г. Ош": ["Центральный", "Амир-Тимур", "Толойкон", "Керме-Тоо", "Жапалак"],
    "Чуйская область": ["Аламудунский район", "Ысык-Атинский район", "Сокулукский район", "Московский район", "Панфиловский район", "Жайылский район", "Кеминский район", "Чуйский район", "г. Токмок"],
    "Ошская область": ["Кара-Сууский район", "Ноокатский район", "Узгенский район", "Алайский район", "Араванский район", "Чон-Алайский район", "Кара-Кулджинский район"],
    "Джалал-Абадская область": ["Сузакский район", "Базар-Коргонский район", "Ноокенский район", "Аксыйский район", "Ала-Букинский район", "Чаткальский район", "Токтогульский район", "Тогуз-Тороуский район", "г. Джалал-Абад", "г. Кара-Куль", "г. Таш-Кумыр", "г. Майлуу-Суу"],
    "Иссык-Кульская область": ["Иссык-Кульский район", "Тюпский район", "Ак-Суйский район", "Джети-Огузский район", "Тонский район", "г. Каракол", "г. Балыкчы"],
    "Нарынская область": ["Нарынский район", "Ат-Башинский район", "Ак-Талинский район", "Жумгальский район", "Кочкорский район", "г. Нарын"],
    "Баткенская область": ["Баткенский район", "Кадамжайский район", "Лейлекский район", "г. Баткен", "г. Кызыл-Кыя", "г. Сулюкта"],
    "Таласская область": ["Таласский район", "Бакай-Атинский район", "Кара-Бууринский район", "Манасский район", "г. Талас"]
}

def upload_to_telegram(file_bytes, file_name):
    try:
        token = st.secrets.get("TELEGRAM_BOT_TOKEN", "")
        chat_id = st.secrets.get("TELEGRAM_CHAT_ID", "")
        if not token or not chat_id: return "Ошибка настроек"
        url = f"https://api.telegram.org/bot{token}/sendDocument"
        files = {"document": (file_name, io.BytesIO(file_bytes), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
        data = {"chat_id": chat_id, "caption": f"📄 Отчет: {file_name}"}
        response = requests.post(url, data=data, files=files)
        return "Отправлено" if response.json().get("ok") else "Ошибка"
    except: return "Сбой"

def get_google_sheets_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)

def append_to_google_sheets(boss_row, db_row):
    try:
        client = get_google_sheets_client()
        doc = client.open_by_key(st.secrets["spreadsheet_id"])
        doc.get_worksheet(0).append_row(boss_row)
        try: sheet_db = doc.worksheet("База_проверок")
        except: sheet_db = doc.add_worksheet(title="База_проверок", rows="1000", cols="10")
        sheet_db.append_row(db_row)
        return True
    except: return False

@st.cache_data(ttl=60)
def get_cached_preview():
    try: return pd.DataFrame(get_google_sheets_client().open_by_key(st.secrets["spreadsheet_id"]).get_worksheet(0).get_all_records())
    except: return None

@st.cache_data(ttl=60)
def get_cached_db():
    try: return pd.DataFrame(get_google_sheets_client().open_by_key(st.secrets["spreadsheet_id"]).worksheet("База_проверок").get_all_records())
    except: return pd.DataFrame()

def parse_sum_from_text(text):
    clean_text = text.replace('\xa0', ' ')
    match = re.search(r'(\d[\d\s]*[.,]\d+)', clean_text)
    if match:
        val_str = match.group(1).replace(' ', '').replace(',', '.')
        try: return float(val_str)
        except: return 0.0
    return 0.0

def format_sum(val):
    return f"{val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', ' ')

# =========================================================================
# ИДЕАЛЬНЫЙ ПАРСЕР
# =========================================================================
def process_smart_calc_tables(uploaded_file, ext_serv="", ext_parts="", ext_mat=""):
    try:
        doc_in = docx.Document(uploaded_file)
        doc_out = docx.Document()
        tables_found = {}
        
        def format_table_smart(tbl_element):
            tblPr = tbl_element.find(qn('w:tblPr'))
            if tblPr is None: tblPr = OxmlElement('w:tblPr'); tbl_element.insert(0, tblPr)
            tblW = tblPr.find(qn('w:tblW'))
            if tblW is None: tblW = OxmlElement('w:tblW'); tblPr.append(tblW)
            tblW.set(qn('w:w'), '5000'); tblW.set(qn('w:type'), 'pct')
            
            rows = tbl_element.findall(qn('w:tr'))
            if len(rows) > 0:
                for tc in rows[0].findall(qn('w:tc')):
                    for p in tc.findall(qn('w:p')):
                        for r in p.findall(qn('w:r')):
                            rPr = r.find(qn('w:rPr'))
                            if rPr is None: rPr = OxmlElement('w:rPr'); r.insert(0, rPr)
                            b = rPr.find(qn('w:b'))
                            if b is None: b = OxmlElement('w:b'); rPr.append(b)
                if len(rows) > 1: tbl_element.remove(rows[1])
            return tbl_element

        for table in doc_in.tables:
            prev_elm = table._element.getprevious()
            header_text = ""
            while prev_elm is not None:
                if prev_elm.tag.endswith('p'):
                    p = docx.text.paragraph.Paragraph(prev_elm, doc_in)
                    if p.text.strip(): header_text = p.text.lower(); break
                prev_elm = prev_elm.getprevious()
            
            col_headers = [cell.text.lower() for cell in table.rows[0].cells] if table.rows else []
            if "воздейств" in header_text or "затрат" in header_text or "услуг" in header_text or "нормо-час" in col_headers: tables_found['services'] = table
            elif "запасных" in header_text or "запчаст" in header_text: tables_found['parts'] = table
            elif "материал" in header_text: tables_found['materials'] = table

        approval_lines = []
        overall_total = 0.0
        
        if 'services' in tables_found:
            p_title = doc_out.add_paragraph()
            r = p_title.add_run("Перечень и стоимость затрат (услуг), необходимых для восстановления:"); r.bold = True
            p_title._p.addnext(format_table_smart(copy.deepcopy(tables_found['services']._element)))
            p_note = doc_out.add_paragraph()
            r = p_note.add_run("Примечание: "); r.bold = True
            p_note.add_run(f"стоимость нормо-часа... (см. настройки). {ext_serv.strip()}\n")
            overall_total += parse_sum_from_text(" ".join([c.text for c in tables_found['services'].rows[-1].cells]))
            approval_lines.append(f"Услуг – {format_sum(overall_total)} сом;")

        if 'parts' in tables_found:
            p_title = doc_out.add_paragraph()
            r = p_title.add_run("Стоимость запасных частей:"); r.bold = True
            p_title._p.addnext(format_table_smart(copy.deepcopy(tables_found['parts']._element)))
            p_note = doc_out.add_paragraph()
            r = p_note.add_run("Примечание: "); r.bold = True
            p_note.add_run(f"описание запчастей. {ext_parts.strip()}\n")
            overall_total += parse_sum_from_text(" ".join([c.text for c in tables_found['parts'].rows[-1].cells]))
            approval_lines.append(f"Запасных частей – {format_sum(overall_total - sum([parse_sum_from_text(c.text) for c in tables_found['services'].rows[-1].cells if 'services' in tables_found]))} сом;")

        if 'materials' in tables_found:
            p_title = doc_out.add_paragraph()
            r = p_title.add_run("Стоимость материалов:"); r.bold = True
            p_title._p.addnext(format_table_smart(copy.deepcopy(tables_found['materials']._element)))
            p_note = doc_out.add_paragraph()
            r = p_note.add_run("Примечание: "); r.bold = True
            p_note.add_run(f"информация о материалах. {ext_mat.strip()}\n")
            overall_total += parse_sum_from_text(" ".join([c.text for c in tables_found['materials'].rows[-1].cells]))
            approval_lines.append(f"Материалов – {format_sum(parse_sum_from_text(' '.join([c.text for c in tables_found['materials'].rows[-1].cells])))} сом;")
            
        buffer = io.BytesIO(); doc_out.save(buffer); buffer.seek(0)
        return buffer, "\n".join(approval_lines) + f"\nИтого: {format_sum(overall_total)} сом."
    except: return None, ""

# ... [Остальной код интерфейса Streamlit остается прежним] ...
