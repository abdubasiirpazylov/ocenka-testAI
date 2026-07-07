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

# --- ИМПОРТЫ ДЛЯ РАБОТЫ С GOOGLE DRIVE И ИИ ---
try:
    import google.generativeai as genai
    from PIL import Image
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload
    HAS_AI_AND_DRIVE = True
except ImportError:
    HAS_AI_AND_DRIVE = False

TEMPLATE_NAME = "образец отчета.docx"

st.set_page_config(page_title="Генератор Отчетов - Гарант Оценка", layout="wide")

# =========================================================
# НАСТРОЙКА ИИ GEMINI (ВЕРСИЯ 2.5 FLASH)
# =========================================================
if HAS_AI_AND_DRIVE and "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    AI_READY = True
else:
    AI_READY = False

# =========================================================
# БАЗА ДАННЫХ АДМИНИСТРАТИВНОГО ДЕЛЕНИЯ КЫРГЫЗСТАНА
# =========================================================
KG_REGIONS = {
    "г. Бишкек": ["Ленинский район", "Октябрьский район", "Первомайский район", "Свердловский район"],
    "г. Ош": ["Центральный", "Амир-Тимур", "Толойкон", "Керме-Тоо", "Жапалак"],
    "Чуйская область": ["Аlaмудунский район", "Ысык-Атинский район", "Сокулукский район", "Московский район", "Панфиловский район", "Жайылский район", "Кеминский район", "Чуйский район", "г. Токмок"],
    "Ошская область": ["Кара-Сууский район", "Ноокатский район", "Узгенский район", "Алайский район", "Араванский район", "Чон-Алайский район", "Кара-Кулджинский район"],
    "Джалал-Абадская область": ["Сузакский район", "Базар-Коргонский район", "Ноокенский район", "Аксыйский район", "Ала-Букинский район", "Чаткальский район", "Токтогульский район", "Тогуз-Тороуский район", "г. Джалал-Абад", "г. Кара-Куль", "г. Таш-Кумыр", "г. Майлуу-Суу"],
    "Иссык-Кульская область": ["Иссык-Кульский район", "Тюпский район", "Ак-Суйский район", "Джети-Огузский район", "Тонский район", "г. Каракол", "г. Балыкчы"],
    "Нарынская область": ["Нарынский район", "Ат-Башинский район", "Ак-Талинский район", "Жумгальский район", "Кочкорский район", "г. Нарын"],
    "Баткенская область": ["Баткенский район", "Кадамжайский район", "Лейлекский район", "г. Баткен", "г. Кызыл-Кыя", "г. Сулюкта"],
    "Таласская область": ["Таласский район", "Бакай-Атинский район", "Кара-Бууринский район", "Манасский район", "г. Талас"]
}

# --- ФУНКЦИЯ ЗАГРУЗКИ ФАЙЛА НА GOOGLE ДИСК ---
def upload_to_google_drive(file_bytes, file_name):
    if not HAS_AI_AND_DRIVE:
        return None
    try:
        scopes = ["https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
        service = build("drive", "v3", credentials=creds)
        
        folder_id = st.secrets.get("GOOGLE_DRIVE_FOLDER_ID", "")
        
        file_metadata = {
            "name": file_name,
            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        }
        if folder_id:
            file_metadata["parents"] = [folder_id]
            
        media = MediaIoBaseUpload(
            io.BytesIO(file_bytes), 
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
            resumable=True
        )
        
        uploaded_file = service.files().create(body=file_metadata, media_body=media, fields="id, webViewLink").execute()
        return uploaded_file.get("webViewLink")
    except Exception as e:
        st.error(f"❌ Ошибка отправки файла в Google Drive Cloud: {e}")
        return None

# --- ФУНКЦИИ ДЛЯ РАБОТЫ С GOOGLE SHEETS И КЭШЕМ ---
def get_google_sheets_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)

def append_to_google_sheets(boss_row, db_row):
    try:
        client = get_google_sheets_client()
        doc = client.open_by_key(st.secrets["spreadsheet_id"])
        
        sheet_boss = doc.get_worksheet(0)
        sheet_boss.append_row(boss_row)
        
        try:
            sheet_db = doc.worksheet("База_проверок")
        except gspread.exceptions.WorksheetNotFound:
            sheet_db = doc.add_worksheet(title="База_проверок", rows="1000", cols="10")
            sheet_db.append_row(["Номер отчета", "Госномер", "VIN код", "Техпаспорт", "Дата отчета", "Ссылка на файл в Облаке"])
            
        sheet_db.append_row(db_row)
        return True
    except Exception as e:
        st.error(f"❌ Ошибка записи в Google Sheets: {e}")
        return False

def get_google_sheets_preview():
    try:
        client = get_google_sheets_client()
        sheet = client.open_by_key(st.secrets["spreadsheet_id"]).get_worksheet(0)
        records = sheet.get_all_records()
        return pd.DataFrame(records)
    except Exception:
        return None

def get_google_sheets_database():
    try:
        client = get_google_sheets_client()
        sheet_db = client.open_by_key(st.secrets["spreadsheet_id"]).worksheet("База_проверок")
        records = sheet_db.get_all_records()
        return pd.DataFrame(records)
    except Exception:
        return pd.DataFrame()

@st.cache_data(ttl=60)
def get_cached_preview():
    return get_google_sheets_preview()

@st.cache_data(ttl=60)
def get_cached_db():
    return get_google_sheets_database()

# --- ФУНКЦИЯ ОЧИСТКИ ФОРМЫ ---
DEFAULT_DAMAGE_SUFFIX = "Дефектный акт на транспортное средство на дату оценки не предоставлялся. Оценка технического состояния произведена без учёта скрытых дефектов."
DEFAULT_REPAIR_SUFFIX = "После завершения ремонтно-восстановительных работ необходим контроль геометрии кузова, зазоров навесных элементов и качества ЛКП. Контроль выполняется организацией, осуществляющей ремонт."

def clear_fields():
    fields_to_clear = [
        "report_num", "contract_num", "date_ocenki", "customer",
        "aymak_input", "street_address", "sum_num", "car_model", 
        "reg_num", "vin", "tech_passport", "year", "engine_vol", 
        "color", "body_type", "service_cost"
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

# =========================================================
# ИНТЕРФЕЙС ПРИЛОЖЕНИЯ
# =========================================================

st.title("🚗 Главное рабочее место оценщика")
st.markdown("Заполните данные, загрузите техпаспорт и прикрепите фотоотчет.")

if os.path.exists(TEMPLATE_NAME):
    st.success(f"✅ Базовый шаблон отчета (`{TEMPLATE_NAME}`) успешно подключен автоматически.")
    template_source = TEMPLATE_NAME
else:
    st.warning(f"⚠️ Файл `{TEMPLATE_NAME}` не найден. Загрузите его вручную ниже:")
    template_source = st.file_uploader("Загрузите шаблон отчета", type="docx")

# =========================================================
# БЛОК МУЛЬТИ-СКАНИРОВАНИЯ ТЕХПАСПОРТА (GEMINI 2.5 FLASH)
# =========================================================
if AI_READY:
    with st.expander("🤖 Умный сканер техпаспорта (Полное распознавание авто + хозяина)", expanded=True):
        st.info("💡 Загрузите фото техпаспорта (можно обе стороны сразу). ИИ сам проверит адреса и исправит опечатки.")
        sts_images = st.file_uploader("Загрузить фото техпаспорта (одно или несколько)", type=["jpg", "jpeg", "png"], key="sts_uploader", accept_multiple_files=True)
        
        if sts_images:
            if st.button("🔍 Распознать данные", type="primary"):
                with st.spinner(f"Движок Gemini 2.5 Flash изучает документы ({len(sts_images)} шт.)..."):
                    try:
                        images_pil = []
                        for img_file in sts_images:
                            img_file.seek(0)
                            images_pil.append(Image.open(img_file))
                        
                        model = genai.GenerativeModel('gemini-2.5-flash')
                        kg_regions_json = json.dumps(KG_REGIONS, ensure_ascii=False)
                        
                        prompt = f"""
                        Это фотографии техпаспорта (СТС) транспортного средства Кыргызской Республики. 
                        Здесь могут быть обе стороны документа (лицевая и оборотная) либо листы электронного варианта.
                        Твоя задача — внимательно изучить ВСЕ фотографии, сопоставить информацию и извлечь данные автомобиля, а также данные О СОБСТВЕННИКЕ (ХОЗЯИНЕ).
                        
                        ОБРАТИ ВНИМАНИЕ НА АДРЕС! В документе могут быть опечатки (например, "Леннинский" вместо "Ленинский") или сокращения (пропуск "Кыргызская Республика").
                        Я передаю тебе строгий системный справочник регионов и районов. Ты должен найти в адресе владельца область/город и район, и сопоставить их со справочником. 
                        Выдай ТОЧНОЕ совпадение из справочника. Улицу, номер дома и квартиры оставь строго как написано в документе (например: "ул. Ахунбаева, дом 34").

                        СПРАВОЧНИК ОБЛАСТЕЙ И РАЙОНОВ:
                        {kg_regions_json}
                        
                        Верни ответ СТРОГО в формате JSON, без каких-либо дополнительных слов, комментариев или markdown-разметки (никаких ```json). 
                        Если каких-то данных не видно ни на одной из фотографий, верни пустую строку "".
                        
                        Формат JSON:
                        {{
                            "customer": "ФИО Собственника/Владельца авто полностью",
                            "region": "ТОЧНОЕ название ключа (области/города) из предоставленного Справочника",
                            "district": "ТОЧНОЕ название района из предоставленного Справочника, соответствующее выбранной области",
                            "aymak": "Село или Айыл аймагы (если есть, оставь как в документе)",
                            "street_address": "Улица, дом, квартира (строго как в документе, например: ул. Токтогула, д. 42)",
                            "car_model": "Марка и модель авто",
                            "reg_num": "Государственный номер (все буквы и цифры слитно в верхнем регистре)",
                            "vin": "Идентификационный номер (VIN)",
                            "tech_passport": "Номер бланка техпаспорта (буквы и цифры)",
                            "year": "Год выпуска",
                            "color": "Цвет кузова",
                            "engine_vol": "Объем двигателя (только цифры, например: 2.5 или 2000)",
                            "body_type": "Тип кузова"
                        }}
                        """
                        
                        request_content = [prompt] + images_pil
                        response = model.generate_content(request_content)
                        
                        raw_json = response.text.strip()
                        if raw_json.startswith("```json"):
                            raw_json = raw_json[7:]
                        if raw_json.startswith("```"):
                            raw_json = raw_json[3:]
                        if raw_json.endswith("```"):
                            raw_json = raw_json[:-3]
                            
                        extracted_data = json.loads(raw_json)
                        
                        ai_region = extracted_data.get("region", "")
                        ai_district = extracted_data.get("district", "")
                        
                        if ai_region in KG_REGIONS:
                            st.session_state["region_select"] = ai_region
                            if ai_district in KG_REGIONS[ai_region]:
                                st.session_state["district_select"] = ai_district
                            else:
                                st.session_state["district_select"] = KG_REGIONS[ai_region][0]
                        
                        st.session_state["customer"] = extracted_data.get("customer", "")
                        st.session_state["aymak_input"] = extracted_data.get("aymak", "")
                        st.session_state["street_address"] = extracted_data.get("street_address", "")
                        
                        st.session_state["car_model"] = extracted_data.get("car_model", "")
                        st.session_state["reg_num"] = extracted_data.get("reg_num", "")
                        st.session_state["vin"] = extracted_data.get("vin", "")
                        st.session_state["tech_passport"] = extracted_data.get("tech_passport", "")
                        st.session_state["year"] = extracted_data.get("year", "")
                        st.session_state["color"] = extracted_data.get("color", "")
                        st.session_state["engine_vol"] = extracted_data.get("engine_vol", "")
                        st.session_state["body_type"] = extracted_data.get("body_type", "")
                        
                        st.success("✅ Все данные успешно распознаны! Опечатки в адресах исправлены, дома и улицы сохранены.")
                        st.rerun() 
                        
                    except Exception as e:
                        st.error(f"❌ Ошибка распознавания: {e}. Заполните поля вручную.")
else:
    st.info("⚠️ Сканер техпаспорта недоступен. Проверьте requirements.txt и Secrets.")

# =========================================================

col_hdr1, col_hdr2 = st.columns([4, 1])
with col_hdr1:
    st.header("1. Ввод данных")
with col_hdr2:
    st.write("") 
    st.button("🧹 Очистить форму", on_click=clear_fields, use_container_width=True, type="secondary")

df_preview = get_cached_preview()
df_db = get_cached_db()

col1, col2 = st.columns(2)

with col1:
    st.subheader("Общие данные")
    report_num = st.text_input("Номер отчета:", key="report_num")
    contract_num = st.text_input("Номер договора:", key="contract_num")
    date_ocenki = st.text_input("Дата оценки:", key="date_ocenki") 
    
    today_str = datetime.now().strftime("%d.%m.%Y")
    if "date_otcheta" not in st.session_state:
        st.session_state.date_otcheta = today_str
    date_otcheta = st.text_input("Дата отчета (только для реестра):", key="date_otcheta")
    
    customer = st.text_input("ФИО Заказчика:", key="customer")
    
    st.markdown("**Адрес регистрации**")
    c_geo1, c_geo2, c_geo3 = st.columns(3)
    
    with c_geo1:
        selected_region = st.selectbox("Область / Город:", list(KG_REGIONS.keys()), key="region_select")
    with c_geo2:
        selected_district = st.selectbox("Район / Округ:", KG_REGIONS[selected_region], key="district_select")
    with c_geo3:
        aymak = st.text_input("Село / Айыл аймагы:", placeholder="Например: с. Ленинское", key="aymak_input")
        
    street_detail = st.text_input("Улица, дом, квартира:", placeholder="Например: ул. Токтогула, д. 42, кв. 5", key="street_address")
    
    if aymak.strip():
        full_address = f"Кыргызская Республика, {selected_region}, {selected_district}, {aymak.strip()}, {street_detail.strip()}"
    else:
        full_address = f"Кыргызская Республика, {selected_region}, {selected_district}, {street_detail.strip()}"
        
    full_address = full_address.strip(", ")
    st.caption(f"**Итоговый адрес для отчета:** {full_address}")
    
    sum_num = st.text_input("Сумма ущерба цифрами:", placeholder="Например: 247300", key="sum_num")
    
    generated_sum_words = ""
    if sum_num:
        try:
            clean_num_str = "".join(c for c in sum_num if c.isdigit() or c in ".,").replace(",", ".")
            if "." in clean_num_str:
                number_val = float(clean_num_str)
                integer_part = int(number_val)
                generated_sum_words = num2words(integer_part, lang='ru').lower()
            elif clean_num_str:
                number_val = int(clean_num_str)
                generated_sum_words = num2words(number_val, lang='ru').lower()
                
            generated_sum_words = re.sub(r'[a-zA-Z|]', '', generated_sum_words)
            generated_sum_words = " ".join(generated_sum_words.split())
        except ValueError:
            generated_sum_words = ""

    sum_words = st.text_input("Сумма ущерба прописью:", value=generated_sum_words)

with col2:
    st.subheader("Данные автомобиля и услуги")
    st.caption("✨ *Эти поля заполняются автоматически при сканировании техпаспорта*")
    car_model = st.text_input("Марка, модель:", key="car_model")
    reg_num = st.text_input("Гос. номер:", key="reg_num")
    vin = st.text_input("VIN код:", key="vin")
    tech_passport = st.text_input("Тех. паспорт №:", key="tech_passport")
    year = st.text_input("Год выпуска:", key="year")
    engine_vol = st.text_input("Объем ДВС:", key="engine_vol")
    color = st.text_input("Цвет кузова:", key="color")
    
    col_inner1, col_inner2 = st.columns(2)
    with col_inner1:
        body_type = st.text_input("Тип кузова:", key="body_type")
    with col_inner2:
        steering = st.selectbox("Положение руля:", ["Левый руль", "Правый руль"], key="steering")
        
    st.divider()
    service_cost = st.text_input("💰 Стоимость услуги (заработок, для отчета шефу):", placeholder="Например: 5000", key="service_cost")

# --- БРОНЕБОЙНАЯ СИСТЕМА ПРОВЕРКИ ПО ОБЕИМ ТАБЛИЦАМ ---
has_duplicates = False 
warnings_list = []

existing_reports = set()
existing_regs = set()
existing_vins = set()
existing_passports = set()

if df_preview is not None and not df_preview.empty:
    actual_cols_1 = {str(col).strip().lower(): col for col in df_preview.columns}
    col_rep1 = actual_cols_1.get("номер отчета")
    col_reg1 = actual_cols_1.get("госномер", actual_cols_1.get("гос. номер"))
    
    if col_rep1: 
        existing_reports.update([str(x).strip().lower() for x in df_preview[col_rep1].dropna() if str(x).strip()])
    if col_reg1: 
        existing_regs.update([str(x).strip().lower() for x in df_preview[col_reg1].dropna() if str(x).strip()])

if df_db is not None and not df_db.empty:
    actual_cols_2 = {str(col).strip().lower(): col for col in df_db.columns}
    col_rep2 = actual_cols_2.get("номер отчета")
    col_reg2 = actual_cols_2.get("госномер", actual_cols_2.get("гос. номер"))
    col_vin2 = actual_cols_2.get("vin код", actual_cols_2.get("vin-код", actual_cols_2.get("vin")))
    col_pass2 = actual_cols_2.get("техпаспорт", actual_cols_2.get("тех. паспорт", actual_cols_2.get("тех паспорт")))
    
    if col_rep2: 
        existing_reports.update([str(x).strip().lower() for x in df_db[col_rep2].dropna() if str(x).strip()])
    if col_reg2: 
        existing_regs.update([str(x).strip().lower() for x in df_db[col_reg2].dropna() if str(x).strip()])
    if col_vin2: 
        existing_vins.update([str(x).strip().lower() for x in df_db[col_vin2].dropna() if str(x).strip()])
    if col_pass2: 
        existing_passports.update([str(x).strip().lower() for x in df_db[col_pass2].dropna() if str(x).strip()])

current_report = report_num.strip().lower() if report_num else ""
current_reg = reg_num.strip().lower() if reg_num else ""
current_vin = vin.strip().lower() if vin else ""
current_passport = tech_passport.strip().lower() if tech_passport else ""

if current_report and current_report in existing_reports:
    warnings_list.append(f"Отчет № **{report_num}**")
if current_reg and current_reg in existing_regs:
    warnings_list.append(f"Госномер **{reg_num}**")
if current_vin and current_vin in existing_vins:
    warnings_list.append(f"VIN-код **{vin}**")
if current_passport and current_passport in existing_passports:
    warnings_list.append(f"Тех. паспорт **{tech_passport}**")
    
if warnings_list:
    has_duplicates = True
    st.error(f"⛔ **ГЕНЕРАЦИЯ ЗАБЛОКИРОВАНА!**\n\nДанные: {', '.join(warnings_list)} уже числятся в базе (найдено совпадение в Google Sheets)!\n\nОбязательно нажмите серую кнопку **«🧹 Очистить форму»** в самом верху.")

st.header("2. Описание повреждений и ремонта")

DAMAGE_TEMPLATES = {
    "--- Выберите шаблон ---": "",
    "[Кузов] Передняя часть": "При осмотре установлены повреждения передней части кузова: деформация бампера, повреждение облицовочных элементов, смещение/деформация навесных деталей, нарушение ЛКП.",
    "[Кузов] Задняя часть": "Выявлены повреждения задней части кузова: деформация бампера, повреждение крышки багажника/фонарей, нарушение геометрии сопряжений, повреждение ЛКП.",
    "[Кузов] Боковая часть": "Установлены повреждения боковой части кузова: деформация дверей/крыльев, повреждение навесных элементов, нарушение ЛКП.",
    "[Кузов] Силовые элементы": "Имеются признаки деформации силовых элементов кузова (лонжерон/панель), требующие восстановительных работ с последующим контролем геометрии.",
    "[Оптика] Фара (трещина/разрушение)": "Блок-фара передняя (указать сторону): сквозное разрушение (трещина) рассеивателя.",
    "[Оптика] Фара (царапины)": "Блок-фара передняя (указать сторону): глубокие царапины и потертости рассеивателя.",
    "[Оптика] Фара (крепления)": "Блок-фара передняя (указать сторону): излом элементов крепления корпуса.",
    "[Стекла] Лобовое (трещина)": "Стекло ветровое: линейная трещина в зоне видимости водителя (или: в зоне работы стеклоочистителей).",
    "[Стекла] Лобовое (скол)": "Стекло ветровое: скол типа «звезда» (или «бычий глаз») с развивающимися трещинами.",
    "[Стекла] Боковое (царапины)": "Стекло передней/задней двери (указать сторону): царапины (задиры) на внешней поверхности.",
    "[Стекла] Боковое (разрушение)": "Стекло передней/задней двери (указать сторону): разрушение элемента (отсутствует).",
    "[Стекла] Заднее (седан)": "Стекло задка: разрушение элемента / глубокие царапины.",
    "[Стекла] Заднее (хэтчбек/внедорожник)": "Стекло двери задка (крышки багажника): повреждение нитей обогрева / разрушение.",
}

REPAIR_TEMPLATES = {
    "--- Выберите шаблон ---": "",
    "[Кузов] Стандартные работы": "Для восстановления требуется выполнить комплекс слесарно-кузовных, рихтовочных и малярно-окрасочных работ с применением расходных материалов, с последующей сборкой и регулировкой навесных элементов.",
    "[Оптика] Замена фары": "Демонтаж, монтаж (замена) блок-фары передней (указать сторону) в сборе.",
    "[Стекла] Лобовое стекло (база)": "Замена стекла ветрового (вклейка) с использованием комплекта однокомпонентного полиуретанового клея.",
    "[Стекла] Лобовое стекло (+датчики)": "Замена стекла ветрового (вклейка) с использованием комплекта однокомпонентного полиуретанового клея и переустановкой датчика дождя/камеры слежения.",
    "[Стекла] Боковое стекло": "Снятие обивки двери, очистка внутренней полости от осколков, замена стекла двери.",
    "[Стекла] Заднее стекло": "Замена стекла задка (вклейка) с подключением элементов обогрева."
}

if "damage_text" not in st.session_state:
    st.session_state.damage_text = DEFAULT_DAMAGE_SUFFIX

if "repair_text" not in st.session_state:
    st.session_state.repair_text = f"Для восстановления требуется выполнить комплекс слесарно-кузовных, рихтовочных и малярно-окрасочных работ с применением расходных материалов, с последующей сборкой и регулировкой навесных элементов.\n{DEFAULT_REPAIR_SUFFIX}"

def add_to_damage():
    selected = st.session_state.dmg_selector
    if selected and DAMAGE_TEMPLATES[selected]:
        current = st.session_state.damage_text
        new_phrase = DAMAGE_TEMPLATES[selected]
        if DEFAULT_DAMAGE_SUFFIX in current:
            st.session_state.damage_text = current.replace(DEFAULT_DAMAGE_SUFFIX, new_phrase + "\n" + DEFAULT_DAMAGE_SUFFIX)
        else:
            st.session_state.damage_text = current + "\n" + new_phrase if current else new_phrase

def add_to_repair():
    selected = st.session_state.rep_selector
    if selected and REPAIR_TEMPLATES[selected]:
        current = st.session_state.repair_text
        new_phrase = REPAIR_TEMPLATES[selected]
        if DEFAULT_REPAIR_SUFFIX in current:
            st.session_state.repair_text = current.replace(DEFAULT_REPAIR_SUFFIX, new_phrase + "\n" + DEFAULT_REPAIR_SUFFIX)
        else:
            st.session_state.repair_text = current + "\n" + new_phrase if current else new_phrase

col_dmg, col_rep = st.columns(2)
with col_dmg:
    st.selectbox("Конструктор осмотра:", list(DAMAGE_TEMPLATES.keys()), key="dmg_selector")
    st.button("➕ Добавить в осмотр", on_click=add_to_damage, use_container_width=True)

with col_rep:
    st.selectbox("Конструктор ремонта:", list(REPAIR_TEMPLATES.keys()), key="rep_selector")
    st.button("➕ Добавить в ремонт", on_click=add_to_repair, use_container_width=True)

damage_desc = st.text_area("Характеристика повреждений (при осмотре установлено):", key="damage_text", height=200)
repair_desc = st.text_area("Требуемый ремонт (для восстановления требуется):", key="repair_text", height=200)

def format_text_with_newlines(text):
    rt = RichText()
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    for i, line in enumerate(lines):
        if i < len(lines) - 1:
            rt.add(line + '\n')
        else:
            rt.add(line)
    return rt

st.header("3. Приложение: Фотоотчет")
st.info("💡 Загрузите сюда файл .docx, который был сгенерирован экспертом в мобильном приложении.")
photo_report_doc = st.file_uploader("Загрузите готовый Фотоотчет (.docx)", type="docx")

if template_source is not None:
    button_clicked = st.button("СГЕНЕРИРОВАТЬ ИТОГОВЫЙ ОТЧЕТ", type="primary", use_container_width=True, disabled=has_duplicates)
    
    if button_clicked:
        if has_duplicates:
            st.error("❌ Системная блокировка! Очистите форму перед генерацией.")
            st.stop()
            
        try:
            if not isinstance(template_source, str):
                template_source.seek(0)
            doc = DocxTemplate(template_source)
            
            if photo_report_doc is not None:
                photo_report_doc.seek(0)
                subdoc_photo = doc.new_subdoc(photo_report_doc)
            else:
                subdoc_photo = "Таблица с фотографиями не была приложена к отчету."

            context = {
                "REPORT_NUM": report_num,
                "CONTRACT_NUM": contract_num,
                "DATE": date_ocenki,
                "CUSTOMER_NAME": customer,
                "ADDRESS": full_address, 
                "CAR_MODEL": car_model,
                "REG_NUM": reg_num,
                "VIN": vin,
                "TECH_PASSPORT": tech_passport,
                "YEAR": year,
                "ENGINE_VOL": engine_vol,
                "COLOR": color,
                "BODY_TYPE": body_type,
                "STEERING": steering,
                "TOTAL_SUM_NUM": sum_num,
                "TOTAL_SUM_WORDS": sum_words,
                "DAMAGE_DESC": format_text_with_newlines(damage_desc),
                "REPAIR_DESC": format_text_with_newlines(repair_desc),
                "PHOTO_TABLE": subdoc_photo 
            }
            
            doc.render(context)
            
            settings = doc.docx.settings.element
            update_fields = OxmlElement('w:updateFields')
            update_fields.set(qn('w:val'), 'true')
            settings.append(update_fields)
            
            buffer = io.BytesIO()
            doc.save(buffer)
            file_bytes = buffer.getvalue()
            
            safe_reg_num = reg_num.strip() if reg_num.strip() else "Без_номера"
            file_name = f"{safe_reg_num}.docx"
            
            # --- АВТОМАТИЧЕСКАЯ ОТПРАВКА В GOOGLE DRIVE CLOUD ---
            with st.spinner("Загрузка отчета в облако Google Drive..."):
                drive_link = upload_to_google_drive(file_bytes, file_name)
                
            if drive_link:
                cloud_status_text = drive_link
                st.info(f"☁️ Файл успешно сохранен в облаке Google Drive!")
            else:
                cloud_status_text = "Ошибка загрузки файла в облако"
            
            # Добавляем ссылку на файл в отчет шефу и базу проверок
            row_boss = [report_num, car_model, reg_num, date_ocenki, date_otcheta, service_cost, cloud_status_text]
            row_db = [report_num, reg_num, vin, tech_passport, date_otcheta, cloud_status_text]
            
            success = append_to_google_sheets(row_boss, row_db)
            
            if success:
                get_cached_preview.clear() 
                get_cached_db.clear() 
                st.success("✅ Отчет создан! Данные и ссылка на файл мгновенно улетели в Google Таблицы.")
            else:
                st.warning("⚠️ Отчет Word создан, но не удалось записать данные в Google Sheets.")
            
            st.download_button(
                label=f"📥 СКАЧАТЬ ИТОГОВЫЙ ОТЧЕТ С ЛОКАЛЬНОГО КОМПЬЮТЕРА ({file_name})",
                data=file_bytes,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"Произошла ошибка при обработке файла: {e}")

st.sidebar.title("📊 Живой отчет для шефа")
st.sidebar.markdown("Данные подгружаются напрямую из облака Google Sheets.")

df_preview = get_cached_preview()
if df_preview is not None and not df_preview.empty:
    st.sidebar.dataframe(df_preview, use_container_width=True)
    st.sidebar.success("🟢 Синхронизация с облаком активна")
    
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df_preview.to_excel(writer, index=False, sheet_name='Реестр')
    excel_buffer.seek(0)
    
    st.sidebar.download_button(
        label="📥 Скачать реестр Excel (.xlsx)",
        data=excel_buffer,
        file_name=f"Отчет_шефу_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.sidebar.info("Таблица пуста или еще не подключена в Secrets.")
