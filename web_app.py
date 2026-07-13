import traceback
import streamlit as st

# Глобальный перехватчик: чтобы приложение не вылетало в "черный экран", а показывало ошибку
try:
    st.set_page_config(page_title="Генератор Отчетов - Гарант Оценка", layout="wide")
except Exception:
    pass # Защита от двойного вызова

try:
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

    # Безопасная проверка ключей
    try:
        HAS_GEMINI_KEY = "GEMINI_API_KEY" in st.secrets
    except Exception:
        HAS_GEMINI_KEY = False

    if HAS_AI and HAS_GEMINI_KEY:
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
            
            if not token or not chat_id:
                st.error("❌ Ошибка: В Secrets не добавлены настройки Telegram!")
                return "Ошибка настроек"
                
            url = f"https://api.telegram.org/bot{token}/sendDocument"
            files = {"document": (file_name, io.BytesIO(file_bytes), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
            data = {
                "chat_id": chat_id, 
                "caption": f"📄 Сгенерирован новый отчет!\n🚗 Автомобиль: {file_name.replace('.docx', '')}\n📅 Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
            }
            
            response = requests.post(url, data=data, files=files)
            if response.json().get("ok"):
                st.info("🚀 Отчет успешно отправлен в Telegram-чат!")
                return "Отправлено в Telegram"
            else:
                return "Ошибка отправки"
        except Exception as e:
            return "Сбой системы"

    def get_google_sheets_client():
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
        return gspread.authorize(creds)

    def append_to_google_sheets(boss_row, db_row):
        try:
            client = get_google_sheets_client()
            doc = client.open_by_key(st.secrets["spreadsheet_id"])
            doc.get_worksheet(0).append_row(boss_row)
            try:
                sheet_db = doc.worksheet("База_проверок")
            except gspread.exceptions.WorksheetNotFound:
                sheet_db = doc.add_worksheet(title="База_проверок", rows="1000", cols="10")
                sheet_db.append_row(["Номер отчета", "Госномер", "VIN код", "Техпаспорт", "Дата отчета", "Статус файла"])
            sheet_db.append_row(db_row)
            return True
        except:
            return False

    @st.cache_data(ttl=60)
    def get_cached_preview():
        try:
            return pd.DataFrame(get_google_sheets_client().open_by_key(st.secrets["spreadsheet_id"]).get_worksheet(0).get_all_records())
        except: 
            return None

    @st.cache_data(ttl=60)
    def get_cached_db():
        try:
            return pd.DataFrame(get_google_sheets_client().open_by_key(st.secrets["spreadsheet_id"]).worksheet("База_проверок").get_all_records())
        except: 
            return pd.DataFrame()

    def parse_sum_from_text(text):
        clean_text = text.replace('\xa0', ' ')
        match = re.search(r'(\d[\d\s]*[.,]\d+)', clean_text)
        if match:
            val_str = match.group(1).replace(' ', '').replace(',', '.')
            try: 
                return float(val_str)
            except: 
                return 0.0
        return 0.0

    def format_sum(val):
        return f"{val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', ' ')

    # =========================================================================
    # БЕЗОПАСНЫЙ ПАРСЕР СМЕТЫ (Cambria + 100% Ширина)
    # =========================================================================
    def process_smart_calc_tables(uploaded_file, ext_serv="", ext_parts="", ext_mat=""):
        try:
            doc_in = docx.Document(uploaded_file)
            doc_out = docx.Document()
            tables_found = {}
            
            def format_table_smart(tbl_element):
                # Растягиваем на 100%
                tblPr = tbl_element.find(qn('w:tblPr'))
                if tblPr is None:
                    tblPr = OxmlElement('w:tblPr')
                    tbl_element.insert(0, tblPr)
                tblW = tblPr.find(qn('w:tblW'))
                if tblW is None:
                    tblW = OxmlElement('w:tblW')
                    tblPr.append(tblW)
                tblW.set(qn('w:w'), '5000')
                tblW.set(qn('w:type'), 'pct')
                
                rows = tbl_element.findall(qn('w:tr'))
                if len(rows) > 0:
                    for row_idx, row in enumerate(rows):
                        for tc in row.findall(qn('w:tc')):
                            for p in tc.findall(qn('w:p')):
                                for r in p.findall(qn('w:r')):
                                    rPr = r.find(qn('w:rPr'))
                                    if rPr is None:
                                        rPr = OxmlElement('w:rPr')
                                        r.insert(0, rPr)
                                    
                                    # Принудительно ставим Cambria
                                    rFonts = rPr.find(qn('w:rFonts'))
                                    if rFonts is None:
                                        rFonts = OxmlElement('w:rFonts')
                                        rPr.append(rFonts)
                                    rFonts.set(qn('w:ascii'), 'Cambria')
                                    rFonts.set(qn('w:hAnsi'), 'Cambria')
                                    rFonts.set(qn('w:cs'), 'Cambria')
                                    
                                    # Делаем жирным заголовки (первую строку)
                                    if row_idx == 0:
                                        b = rPr.find(qn('w:b'))
                                        if b is None:
                                            b = OxmlElement('w:b')
                                            rPr.append(b)
                    
                    # Удаляем нумерацию (2-ю строку)
                    if len(rows) > 1:
                        tbl_element.remove(rows[1])

                return tbl_element

            # Поиск таблиц
            for table in doc_in.tables:
                prev_elm = table._element.getprevious()
                header_text = ""
                while prev_elm is not None:
                    if prev_elm.tag.endswith('p'):
                        p = docx.text.paragraph.Paragraph(prev_elm, doc_in)
                        if p.text.strip():
                            header_text = p.text.lower()
                            break
                    prev_elm = prev_elm.getprevious()
                
                col_headers = [cell.text.lower() for cell in table.rows[0].cells] if table.rows else []
                
                if "воздейств" in header_text or "затрат" in header_text or "услуг" in header_text or "нормо-час" in col_headers:
                    tables_found['services'] = table
                elif "запасных" in header_text or "запчаст" in header_text:
                    tables_found['parts'] = table
                elif "материал" in header_text:
                    tables_found['materials'] = table

            approval_lines = []
            overall_total = 0.0
            
            def add_styled_run(paragraph, text, bold=False):
                r = paragraph.add_run(text)
                r.bold = bold
                r.font.name = 'Cambria'
                return r
            
            if 'services' in tables_found:
                p_title = doc_out.add_paragraph()
                add_styled_run(p_title, "Перечень и стоимость затрат (услуг), необходимых для восстановления:", bold=True)
                
                copied_tbl = copy.deepcopy(tables_found['services']._element)
                copied_tbl = format_table_smart(copied_tbl)
                p_title._p.addnext(copied_tbl)
                
                p_note = doc_out.add_paragraph()
                add_styled_run(p_note, "Примечание: ", bold=True)
                note_text = "стоимость нормо-часа ремонтно-восстановительных работ (1500,00 сом) определена согласно анализу стоимости услуг на станциях технического обслуживания: ИП «Сергей», +996702200885; +996553535533; +996550444488, +996559885102, +996550180555; +996555495545."
                if ext_serv.strip():
                    note_text += f"\nДополнительно: {ext_serv.strip()}"
                add_styled_run(p_note, note_text)
                doc_out.add_paragraph() 
                
                last_row_text = " ".join([cell.text for cell in tables_found['services'].rows[-1].cells])
                val = parse_sum_from_text(last_row_text)
                if val > 0:
                    approval_lines.append(f"Услуг – {format_sum(val)} сом;")
                    overall_total += val

            if 'parts' in tables_found:
                p_title = doc_out.add_paragraph()
                add_styled_run(p_title, "Стоимость запасных частей:", bold=True)
                
                copied_tbl = copy.deepcopy(tables_found['parts']._element)
                copied_tbl = format_table_smart(copied_tbl)
                p_title._p.addnext(copied_tbl)
                
                p_note = doc_out.add_paragraph()
                add_styled_run(p_note, "Примечание: ", bold=True)
                note_text = "указана средняя стоимость запасных частей, поддержанных оригинальных, дубликатов на основании анализа рынка ЕАЭС.\nСсылки: в качестве информации была использована база данных ОсОО «Гарант Оценка»; интернет-ресурсы: mashina.kg, lalafo.kg; +996 551 411 711; +996 504 386 999; +996 500 524 624; +996 556 522 516; +996 707 008 833; +996 707 380 001."
                if ext_parts.strip():
                    note_text += f"\nДополнительные ссылки: {ext_parts.strip()}"
                add_styled_run(p_note, note_text)
                doc_out.add_paragraph()
                
                last_row_text = " ".join([cell.text for cell in tables_found['parts'].rows[-1].cells])
                val = parse_sum_from_text(last_row_text)
                if val > 0:
                    approval_lines.append(f"Запасных частей – {format_sum(val)} сом;")
                    overall_total += val

            if 'materials' in tables_found:
                p_title = doc_out.add_paragraph()
                add_styled_run(p_title, "Стоимость материалов:", bold=True)
                
                copied_tbl = copy.deepcopy(tables_found['materials']._element)
                copied_tbl = format_table_smart(copied_tbl)
                p_title._p.addnext(copied_tbl)
                
                p_note = doc_out.add_paragraph()
                add_styled_run(p_note, "Примечание: ", bold=True)
                note_text = "указана средняя стоимость материалов, источники конъюнктурного анализа: +996 708 707 332; +996 13 54 46; +996 550 98 77 01; +996 553 40 03 98"
                if ext_mat.strip():
                    note_text += f"\nДополнительно: {ext_mat.strip()}"
                add_styled_run(p_note, note_text)
                
                last_row_text = " ".join([cell.text for cell in tables_found['materials'].rows[-1].cells])
                val = parse_sum_from_text(last_row_text)
                if val > 0:
                    approval_lines.append(f"Материалов – {format_sum(val)} сом;")
                    overall_total += val
                
            if len(tables_found) == 0:
                doc_out.add_paragraph("Таблицы с расчетами не найдены в загруженном документе.")
                
            buffer = io.BytesIO()
            doc_out.save(buffer)
            buffer.seek(0)
            
            approval_text = "\n".join(approval_lines)
            if overall_total > 0:
                approval_text += f"\nИтого: {format_sum(overall_total)} сом."
                
            return buffer, approval_text
        except Exception as e:
            return None, ""
    # =========================================================================

    DEFAULT_DAMAGE_SUFFIX = "Дефектный акт на транспортное средство на дату оценки не предоставлялся. Оценка технического состояния произведена без учёта скрытых дефектов."
    DEFAULT_REPAIR_SUFFIX = "После завершения ремонтно-восстановительных работ необходим контроль геометрии кузова, зазоров навесных элементов и качества ЛКП. Контроль выполняется организацией, осуществляющей ремонт."

    def clear_fields():
        fields_to_clear = [
            "report_num", "contract_num", "date_ocenki", "customer",
            "aymak_input", "street_address", "sum_num", "car_model", 
            "reg_num", "vin", "tech_passport", "year", "engine_vol", 
            "color", "body_type", "service_cost", "extra_services", 
            "extra_parts", "extra_materials"
        ]
        for f in fields_to_clear:
            if f in st.session_state:
                st.session_state[f] = ""
        
        if "steering" in st.session_state:
            st.session_state["steering"] = "Левый руль"
        if "date_otcheta" in st.session_state:
            st.session_state["date_otcheta"] = datetime.now().strftime("%d.%m.%Y")
        if "region_select" in st.session_state:
            st.session_state["region_select"] = list(KG_REGIONS.keys())[0]
        if "district_select" in st.session_state:
            st.session_state["district_select"] = KG_REGIONS[list(KG_REGIONS.keys())[0]][0]
            
        st.session_state.damage_text = DEFAULT_DAMAGE_SUFFIX
        st.session_state.repair_text = f"Для восстановления требуется выполнить комплекс слесарно-кузовных, рихтовочных и малярно-окрасочных работ с применением расходных материалов, с последующей сборкой и регулировкой навесных элементов.\n{DEFAULT_REPAIR_SUFFIX}"

    st.title("🚗 Главное рабочее место оценщика")

    if os.path.exists(TEMPLATE_NAME):
        st.success(f"✅ Базовый шаблон отчета (`{TEMPLATE_NAME}`) подключен.")
        template_source = TEMPLATE_NAME
    else:
        template_source = st.file_uploader("Загрузите шаблон отчета", type="docx")

    if AI_READY:
        with st.expander("🤖 Умный сканер техпаспорта", expanded=True):
            sts_images = st.file_uploader("Загрузить фото техпаспорта", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
            if sts_images and st.button("🔍 Распознать данные", type="primary"):
                with st.spinner("Изучаю документы..."):
                    try:
                        images_pil = [Image.open(img) for img in sts_images]
                        model = genai.GenerativeModel('gemini-2.5-flash')
                        prompt = f"Это фото техпаспорта КР. Верни СТРОГО в формате JSON без markdown ключи: customer, region, district, aymak, street_address, car_model, reg_num, vin, tech_passport, year, color, engine_vol, body_type. Справочник регионов: {json.dumps(KG_REGIONS, ensure_ascii=False)}"
                        response = model.generate_content([prompt] + images_pil)
                        raw_json = response.text.strip().strip("```json").strip("
