import os
import json
import streamlit as st
from docx_generator import generate_tariff_docx, combine_images_vertically

st.set_page_config(
    page_title="TCF Tariff Sheet System",
    layout="wide",
    initial_sidebar_state="collapsed"
)

DB_PATH = "elements_db.json"

@st.cache_data
def load_db():
    if not os.path.exists(DB_PATH):
        return {}
    with open(DB_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

db = load_db()

# Полностью скрываем весь мусор Streamlit и настраиваем чистый современный дизайн
st.markdown("""
<style>
    #MainMenu, footer, header {visibility: hidden;}
    .stAppDeployButton {display: none;}
    
    .stApp {
        background-color: #f4f6f9;
        color: #2c3e50;
    }
    /* Карточки боксов с мягкой тенью */
    div[data-testid="stVerticalBlock"] > div[style*="border"] {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 12px !important;
        padding: 12px !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
    }
    .soft-panel {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 16px;
        margin-bottom: 14px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    .status-ok {
        color: #0d6efd;
        font-weight: 600;
        padding: 6px 0;
        border-bottom: 1px solid #e2e8f0;
    }
    .status-warn {
        color: #fd7e14;
        font-weight: 600;
        padding: 6px 0;
        border-bottom: 1px solid #e2e8f0;
    }
    .preview-card {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 8px;
        text-align: center;
        margin-bottom: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04);
    }
    .preview-hdr {
        display: flex;
        justify-content: space-between;
        font-size: 12px;
        font-weight: 700;
        color: #64748b;
        border-bottom: 1px solid #e2e8f0;
        padding-bottom: 4px;
        margin-bottom: 6px;
    }
    /* Поля ввода и селекты */
    .stTextInput input, .stSelectbox select, .stNumberInput input {
        background-color: #ffffff !important;
        color: #1e293b !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 6px !important;
    }
    /* Заголовок ряда боксов в стиле сайта */
    .row-banner {
        background-color: #e2e8f0;
        color: #1e293b;
        font-weight: 700;
        padding: 8px 16px;
        border-radius: 8px;
        border: 1px solid #cbd5e1;
        margin-top: 8px;
        margin-bottom: 12px;
        font-size: 15px;
    }
    /* Все кнопки в едином стиле с красивым ховером */
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        border: 1px solid #cbd5e1;
        background-color: #ffffff;
        color: #1e293b;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        transition: all 0.15s ease-in-out;
    }
    .stButton > button:hover {
        border-color: #334155;
        background-color: #f8fafc;
        color: #334155;
        box-shadow: 0 4px 6px rgba(0,0,0,0.08);
        transform: translateY(-1px);
    }
</style>
""", unsafe_allow_html=True)

UI = {
    "TR": {
        "lang_label": "Dil",
        "app_title": "Akrobatik Cimnastik Tariff Sheet Sistemi",
        "comp_label": "Yarışma Türü",
        "comp_tr": "Ulusal Yarışma",
        "comp_int": "Uluslararası Yarışma",
        "team_header": "TAKIM VE SPORCU BİLGİLERİ",
        "age_group": "Yaş Grubu",
        "category": "Yarışma Kategorisi",
        "athletes_hdr": "Sporcuların Adı ve Soyadı",
        "athletes_int_hint": "Uluslararası: İsimleri Latin harfleriyle giriniz.",
        "athletes_tr_hint": "Ulusal: Belge Türkçe oluşturulacaktır.",
        "club_tr": "Kulübün Adı",
        "club_int": "Kulüp / Federasyon Adı",
        "phone_tr": "Antrenörün Telefon NO",
        "phone_int": "Antrenörün Telefon NO",
        "team_no_tr": "Takım NO",
        "team_no_int": "Takım NO",
        "clear_all": "Tüm Kutuları Temizle",
        "boxes_hdr": "ELEMENT KUTULARI (1 — 14)",
        "boxes_sub": "Element seçmek için kutudaki butona tıklayın.",
        "row1": "1. Sıra (Kutular 1 – 7)",
        "row2": "2. Sıra (Kutular 8 – 14)",
        "box_lbl": "Kutu",
        "mode_empty": "Boş",
        "mode_team": "Çift / Grup",
        "mode_indiv": "Bireysel",
        "mode_sep": "Ayırıcı //",
        "sep_sum_lbl": "Seri Ara Toplamı (// X)",
        "pick_el_btn": "Görsel Katalogdan Seç",
        "val_lbl": "Zorluk Değeri (V)",
        "hold_lbl": "Süre (Saniye)",
        "el_no_lbl": "Element NO",
        "same_all_lbl": "Tüm partnerler aynı",
        "calc_hdr": "ZORLUK HESAPLAMASI",
        "team_pts": "Çift / Grup Toplamı",
        "indiv_pts": "Bireysel Puanı",
        "rules_hdr": "KURALLAR VE KONTROL (TCF 2026)",
        "dl_hdr": "BELGEYİ İNDİR",
        "dl_btn": "TARIFF SHEET İNDİR (.DOCX)",
        "dl_note": "Word belgesi tek sayfa A4 yatay olarak hazırlandı.",
        "preview_hdr": "TARIFF SHEET ÖNİZLEME",
        "dialog_hdr": "Element Kataloğu",
        "modal_title": "Görsel Element Kataloğu — Kutu",
        "search_no": "Element Numarası Ara (Örn: 053)",
        "filter_type": "Element Türü",
        "filter_all": "Tümü",
        "filter_bal": "Balans (3\")",
        "filter_dyn": "Dinamik",
        "filter_trans": "Alt Kişi Geçişi (+V)",
        "select_btn": "Seç",
        "roles_3": ["Üst", "Orta", "Alt"],
        "roles_4": ["Üst", "Orta 1", "Orta 2", "Alt"],
        "roles_2": ["Üst", "Alt"],
        "cats_display": {
            "Çift Erkek": "Çift Erkek",
            "Çift Kadın": "Çift Kadın",
            "Çift Mix": "Çift Mix",
            "3'lü Kadın": "3'lü Kadın",
            "4'lü Erkek": "4'lü Erkek"
        },
        "rule_count_ok": "Çift/Grup element sayısı tamamlandı",
        "rule_count_warn": "Çift/Grup element sayısı eksik/fazla. Gereken:",
        "rule_bd_ok": "Balans ve Dinamik dağılımı uygun",
        "rule_bd_warn": "En az 2 Balans ve 2 Dinamik olmalıdır",
        "rule_ind_ok": "Bireysel elementler tamamlandı (3/3)",
        "rule_ind_warn": "Bireysel element sayısı (Gereken: 3):"
    },
    "RU": {
        "lang_label": "Язык",
        "app_title": "Системный конструктор Tariff Sheet",
        "comp_label": "Тип соревнований",
        "comp_tr": "Национальные соревнования",
        "comp_int": "Международные соревнования",
        "team_header": "ДАННЫЕ КОМАНДЫ И СПОРТСМЕНОВ",
        "age_group": "Возрастная группа",
        "category": "Категория",
        "athletes_hdr": "Имя и Фамилия спортсменов",
        "athletes_int_hint": "Международный: вводите имена латиницей.",
        "athletes_tr_hint": "Национальный: бланк будет сформирован на турецком языке.",
        "club_tr": "Название клуба",
        "club_int": "Клуб / Федерация",
        "phone_tr": "Телефон тренера",
        "phone_int": "Телефон тренера",
        "team_no_tr": "Номер команды",
        "team_no_int": "Номер команды",
        "clear_all": "Очистить все боксы",
        "boxes_hdr": "ЗАПОЛНЕНИЕ БОКСОВ (1 — 14)",
        "boxes_sub": "Нажмите кнопку в боксе для открытия визуального каталога.",
        "row1": "Ряд 1 (Боксы 1 – 7)",
        "row2": "Ряд 2 (Боксы 8 – 14)",
        "box_lbl": "Бокс",
        "mode_empty": "Пусто",
        "mode_team": "Парный / Групповой",
        "mode_indiv": "Индивидуальный",
        "mode_sep": "Разделитель //",
        "sep_sum_lbl": "Сумма связки (// X)",
        "pick_el_btn": "Выбрать элемент (Каталог)",
        "val_lbl": "Трудность (V)",
        "hold_lbl": "Секунды статики",
        "el_no_lbl": "Номер элемента",
        "same_all_lbl": "Одинаковый у всех",
        "calc_hdr": "РАСЧЕТ СЛОЖНОСТИ",
        "team_pts": "Парные / Групповые",
        "indiv_pts": "Индивидуальные",
        "rules_hdr": "ПРОВЕРКА ПРАВИЛ (TCF 2026)",
        "dl_hdr": "СКАЧАТЬ ДОКУМЕНТ",
        "dl_btn": "СКАЧАТЬ TARIFF SHEET (.DOCX)",
        "dl_note": "Документ полностью редактируемый в Word на 1 листе А4.",
        "preview_hdr": "ПРЕВЬЮ БЛАНКА TARIFF SHEET",
        "dialog_hdr": "Каталог элементов",
        "modal_title": "Визуальный каталог элементов — Бокс",
        "search_no": "Поиск по номеру (например: 053)",
        "filter_type": "Тип элемента",
        "filter_all": "Все",
        "filter_bal": "Баланс (3\")",
        "filter_dyn": "Динамика",
        "filter_trans": "Переход (+V)",
        "select_btn": "Выбрать",
        "roles_3": ["Верхняя", "Средняя", "Нижняя"],
        "roles_4": ["Верхний", "Средний 1", "Средний 2", "Нижний"],
        "roles_2": ["Верхний", "Нижний"],
        "cats_display": {
            "Çift Erkek": "Мужская пара",
            "Çift Kadın": "Женская пара",
            "Çift Mix": "Смешанная пара",
            "3'lü Kadın": "Женская тройка",
            "4'lü Erkek": "Мужская четверка"
        },
        "rule_count_ok": "Количество парных/групповых элементов в норме",
        "rule_count_warn": "Количество парных/групповых не совпадает. Требуется:",
        "rule_bd_ok": "Баланс и Динамика распределены по правилам",
        "rule_bd_warn": "Минимум 2 баланса и 2 динамики обязательны",
        "rule_ind_ok": "Все 3 бокса индивидуальных заполнены (3/3)",
        "rule_ind_warn": "Заполнено индивидуальных боксов (Требуется 3):"
    },
    "EN": {
        "lang_label": "Language",
        "app_title": "Acrobatic Gymnastics Tariff Sheet System",
        "comp_label": "Competition Type",
        "comp_tr": "National Competition",
        "comp_int": "International Competition",
        "team_header": "TEAM AND GYMNAST INFORMATION",
        "age_group": "Age Group",
        "category": "Competition Category",
        "athletes_hdr": "Gymnasts' First and Last Names",
        "athletes_int_hint": "International: Enter names in Latin characters.",
        "athletes_tr_hint": "National: Document will be generated in Turkish format.",
        "club_tr": "Club Name",
        "club_int": "Club / Federation Name",
        "phone_tr": "Coach Phone NO",
        "phone_int": "Coach Phone NO",
        "team_no_tr": "Team NO",
        "team_no_int": "Team NO",
        "clear_all": "Clear All Boxes",
        "boxes_hdr": "ELEMENT BOXES (1 — 14)",
        "boxes_sub": "Click the button inside a box to open the visual catalog.",
        "row1": "Row 1 (Boxes 1 – 7)",
        "row2": "Row 2 (Boxes 8 – 14)",
        "box_lbl": "Box",
        "mode_empty": "Empty",
        "mode_team": "Pair / Group",
        "mode_indiv": "Individual",
        "mode_sep": "Separator //",
        "sep_sum_lbl": "Series Subtotal (// X)",
        "pick_el_btn": "Select from Catalog",
        "val_lbl": "Difficulty Value (V)",
        "hold_lbl": "Hold Time (Seconds)",
        "el_no_lbl": "Element NO",
        "same_all_lbl": "Same for all partners",
        "calc_hdr": "DIFFICULTY CALCULATION",
        "team_pts": "Pair / Group Total",
        "indiv_pts": "Individual Points",
        "rules_hdr": "RULES AND VALIDATION (TCF 2026)",
        "dl_hdr": "DOWNLOAD DOCUMENT",
        "dl_btn": "DOWNLOAD TARIFF SHEET (.DOCX)",
        "dl_note": "Word document is formatted on a single landscape A4 page.",
        "preview_hdr": "TARIFF SHEET PREVIEW",
        "dialog_hdr": "Element Catalog",
        "modal_title": "Visual Element Catalog — Box",
        "search_no": "Search by Element Number (e.g. 053)",
        "filter_type": "Element Type",
        "filter_all": "All",
        "filter_bal": "Balance (3\")",
        "filter_dyn": "Dynamic",
        "filter_trans": "Base Transition (+V)",
        "select_btn": "Select",
        "roles_3": ["Top", "Middle", "Base"],
        "roles_4": ["Top", "Middle 1", "Middle 2", "Base"],
        "roles_2": ["Top", "Base"],
        "cats_display": {
            "Çift Erkek": "Men's Pair",
            "Çift Kadın": "Women's Pair",
            "Çift Mix": "Mixed Pair",
            "3'lü Kadın": "Women's Group",
            "4'lü Erkek": "Men's Group"
        },
        "rule_count_ok": "Pair/Group element count is complete",
        "rule_count_warn": "Pair/Group element count mismatch. Required:",
        "rule_bd_ok": "Balance and Dynamic distribution is valid",
        "rule_bd_warn": "At least 2 Balance and 2 Dynamic elements required",
        "rule_ind_ok": "Individual elements completed (3/3)",
        "rule_ind_warn": "Individual element boxes filled (Required: 3):"
    },
    "FR": {
        "lang_label": "Langue",
        "app_title": "Système de Feuille de Tarif de Gymnastique Acrobatique",
        "comp_label": "Type de Compétition",
        "comp_tr": "Compétition Nationale",
        "comp_int": "Compétition Internationale",
        "team_header": "INFORMATIONS SUR L'ÉQUIPE ET LES GYMNASTES",
        "age_group": "Catégorie d'Âge",
        "category": "Catégorie de Compétition",
        "athletes_hdr": "Nom et Prénom des Gymnastes",
        "athletes_int_hint": "International : Saisissez les noms en caractères latins.",
        "athletes_tr_hint": "National : Le document sera généré au format turc.",
        "club_tr": "Nom du Club",
        "club_int": "Nom du Club / Fédération",
        "phone_tr": "Téléphone de l'Entraîneur",
        "phone_int": "Téléphone de l'Entraîneur",
        "team_no_tr": "N° d'Équipe",
        "team_no_int": "N° d'Équipe",
        "clear_all": "Effacer Toutes les Cases",
        "boxes_hdr": "CASES D'ÉLÉMENTS (1 — 14)",
        "boxes_sub": "Cliquez sur le bouton dans une case pour ouvrir le catalogue visuel.",
        "row1": "Rangée 1 (Cases 1 – 7)",
        "row2": "Rangée 2 (Cases 8 – 14)",
        "box_lbl": "Case",
        "mode_empty": "Vide",
        "mode_team": "Duo / Groupe",
        "mode_indiv": "Individuel",
        "mode_sep": "Séparateur //",
        "sep_sum_lbl": "Sous-total de la Série (// X)",
        "pick_el_btn": "Choisir dans le Catalogue",
        "val_lbl": "Valeur de Difficulté (V)",
        "hold_lbl": "Temps de Maintien (Secondes)",
        "el_no_lbl": "N° d'Élément",
        "same_all_lbl": "Identique pour tous les partenaires",
        "calc_hdr": "CALCUL DE LA DIFFICULTÉ",
        "team_pts": "Total Duo / Groupe",
        "indiv_pts": "Points Individuels",
        "rules_hdr": "RÈGLES ET CONTRÔLE (TCF 2026)",
        "dl_hdr": "TÉLÉCHARGER LE DOCUMENT",
        "dl_btn": "TÉLÉCHARGER LA FEUILLE DE TARIF (.DOCX)",
        "dl_note": "Le document Word est préparé sur une seule page A4 paysage.",
        "preview_hdr": "APERÇU DE LA FEUILLE DE TARIF",
        "dialog_hdr": "Catalogue d'Éléments",
        "modal_title": "Catalogue Visuel d'Éléments — Case",
        "search_no": "Rechercher par N° d'Élément (ex: 053)",
        "filter_type": "Type d'Élément",
        "filter_all": "Tous",
        "filter_bal": "Statique (3\")",
        "filter_dyn": "Dynamique",
        "filter_trans": "Transition du Porteur (+V)",
        "select_btn": "Choisir",
        "roles_3": ["Voltigeur", "Milieu", "Porteur"],
        "roles_4": ["Voltigeur", "Milieu 1", "Milieu 2", "Porteur"],
        "roles_2": ["Voltigeur", "Porteur"],
        "cats_display": {
            "Çift Erkek": "Duo Masculin",
            "Çift Kadın": "Duo Féminin",
            "Çift Mix": "Duo Mixte",
            "3'lü Kadın": "Trio Féminin",
            "4'lü Erkek": "Quatuor Masculin"
        },
        "rule_count_ok": "Nombre d'éléments Duo/Groupe complet",
        "rule_count_warn": "Nombre d'éléments Duo/Groupe incorrect. Requis :",
        "rule_bd_ok": "Répartition Statique et Dynamique conforme",
        "rule_bd_warn": "Au moins 2 Statiques et 2 Dynamiques requis",
        "rule_ind_ok": "Éléments individuels complets (3/3)",
        "rule_ind_warn": "Cases d'éléments individuels remplies (Requis : 3) :"
    }
}

LANG_NAMES = {
    "TR": "Türkçe",
    "RU": "Русский",
    "EN": "English",
    "FR": "Français"
}

if "ui_lang" not in st.session_state:
    st.session_state["ui_lang"] = "TR"

# Верхняя шапка: Заголовок, выбор типа соревнований и языка
top_c1, top_c2, top_c3 = st.columns([2.0, 1.4, 0.6])

with top_c3:
    current_lang = st.session_state["ui_lang"]
    ui_lang = st.selectbox(
        UI[current_lang]["lang_label"],
        options=["TR", "RU", "EN", "FR"],
        format_func=lambda x: LANG_NAMES[x],
        key="ui_lang"
    )

T = UI[ui_lang]

with top_c1:
    st.markdown(f"### {T['app_title']}")

with top_c2:
    comp_type = st.selectbox(
        T["comp_label"],
        options=["TR", "INT"],
        format_func=lambda x: T["comp_tr"] if x == "TR" else T["comp_int"]
    )

if not db:
    st.error("elements_db.json not found.")
    st.stop()


@st.dialog(T["dialog_hdr"], width="large")
def open_element_picker(state_key, allowed_codes, is_indiv=False, box_num=1, role_title=""):
    st.markdown(f"#### {T['modal_title']} {box_num} {f'({role_title})' if role_title else ''}")
    f_col1, f_col2 = st.columns([1.2, 1.5])
    with f_col1:
        search_q = st.text_input(T["search_no"], key=f"search_{state_key}").strip()
    with f_col2:
        if not is_indiv:
            type_filter = st.selectbox(
                T["filter_type"],
                options=["ALL", "balance", "dynamic", "base_transition"],
                format_func=lambda x: {
                    "ALL": T["filter_all"],
                    "balance": T["filter_bal"],
                    "dynamic": T["filter_dyn"],
                    "base_transition": T["filter_trans"]
                }[x],
                key=f"flt_{state_key}"
            )
        else:
            type_filter = "ALL"

    filtered = []
    for c in allowed_codes:
        info = db[c]
        if type_filter != "ALL" and info["type"] != type_filter:
            continue
        if search_q:
            if search_q.lstrip("0") not in str(info["num"]) and search_q not in c:
                continue
        filtered.append(c)

    cols_per_row = 5
    for i in range(0, len(filtered), cols_per_row):
        row_slice = filtered[i:i + cols_per_row]
        grid_cols = st.columns(cols_per_row)
        for j, code in enumerate(row_slice):
            info = db[code]
            with grid_cols[j]:
                with st.container(border=True):
                    hold_badge = f" | {info['hold']}" if info["hold"] else ""
                    st.caption(f"**No {code}** | **V {info['value']}**{hold_badge}")
                    if os.path.exists(info["image"]):
                        st.image(info["image"], use_container_width=True)
                    if st.button(f"{T['select_btn']} {code}", key=f"pick_{state_key}_{code}", use_container_width=True):
                        st.session_state[state_key] = code
                        st.rerun()


# --- КОМПАКТНЫЙ БЛОК ДАННЫХ КОМАНДЫ В ВЕРХНЕЙ ЧАСТИ ЭКРАНА ---
with st.container(border=True):
    st.markdown(f"##### {T['team_header']}")
    
    col_meta1, col_meta2 = st.columns(2)
    with col_meta1:
        age_group = st.selectbox(T["age_group"], ["6-12", "8-14", "9-15"])
    with col_meta2:
        category = st.selectbox(
            T["category"],
            options=["Çift Erkek", "Çift Kadın", "Çift Mix", "3'lü Kadın", "4'lü Erkek"],
            format_func=lambda x: T["cats_display"][x]
        )

    if category == "3'lü Kadın":
        partner_roles = T["roles_3"]
    elif category == "4'lü Erkek":
        partner_roles = T["roles_4"]
    else:
        partner_roles = T["roles_2"]

    num_partners = len(partner_roles)

    st.markdown(f"**{T['athletes_hdr']}**")
    if comp_type == "INT":
        st.caption(T["athletes_int_hint"])
    else:
        st.caption(T["athletes_tr_hint"])

    ath_cols = st.columns(num_partners)
    athletes = []
    for idx_r, role in enumerate(partner_roles):
        with ath_cols[idx_r]:
            athletes.append(st.text_input(role, key=f"ath_{idx_r}"))

    club_label = T["club_int"] if comp_type == "INT" else T["club_tr"]
    phone_label = T["phone_int"] if comp_type == "INT" else T["phone_tr"]
    team_label = T["team_no_int"] if comp_type == "INT" else T["team_no_tr"]

    inf_cols = st.columns(3)
    with inf_cols[0]:
        club_name = st.text_input(club_label, value="")
    with inf_cols[1]:
        coach_phone = st.text_input(phone_label, value="")
    with inf_cols[2]:
        team_no = st.text_input(team_label, value="")

    if st.button(T["clear_all"], use_container_width=False):
        for k in list(st.session_state.keys()):
            if k.startswith("b_") or k.startswith("sel_"):
                del st.session_state[k]
        st.rerun()

st.divider()


team_codes = []
indiv_codes = []
for code, info in sorted(db.items(), key=lambda x: x[1]["num"]):
    if category in info["allowed_categories"]:
        if info["type"] == "individual":
            indiv_codes.append(code)
        else:
            team_codes.append(code)


st.markdown(f"#### {T['boxes_hdr']}")
st.caption(T["boxes_sub"])

boxes_data = {}
running_series_sum = 0
team_balance_count = 0
team_dynamic_count = 0
indiv_box_count = 0

total_team_points = 0
total_indiv_sum_all_partners = 0

MODE_KEYS = ["empty", "team", "indiv", "sep"]

# --- ОТОБРАЖЕНИЕ БОКСОВ БЕЗ ВКЛАДОК: РЯД 1 (1-7) И ПОД НИМ РЯД 2 (8-14) ---
row_titles = [T["row1"], T["row2"]]

for row_idx, (start_b, end_b) in enumerate([(1, 7), (8, 14)]):
    st.markdown(f"<div class='row-banner'>{row_titles[row_idx]}</div>", unsafe_allow_html=True)
    cols = st.columns(7)
    for idx, box_num in enumerate(range(start_b, end_b + 1)):
        with cols[idx]:
            with st.container(border=True):
                st.markdown(f"**{T['box_lbl']} {box_num}**")
                mode = st.selectbox(
                    "Mode",
                    options=MODE_KEYS,
                    format_func=lambda m: {
                        "empty": T["mode_empty"],
                        "team": T["mode_team"],
                        "indiv": T["mode_indiv"],
                        "sep": T["mode_sep"]
                    }[m],
                    key=f"b_mode_{box_num}",
                    label_visibility="collapsed"
                )

                if mode == "empty":
                    boxes_data[box_num] = {
                        "mode": "empty",
                        "hold": "",
                        "d_cells": ["", "", "", ""],
                        "element_no": "",
                        "images": []
                    }

                elif mode == "sep":
                    default_sum = str(running_series_sum) if running_series_sum > 0 else ""
                    custom_sum = st.text_input(T["sep_sum_lbl"], value=default_sum, key=f"b_sep_{box_num}")
                    boxes_data[box_num] = {
                        "mode": "separator",
                        "sep_sum": custom_sum,
                        "hold": "",
                        "d_cells": ["", "", "", ""],
                        "element_no": "",
                        "images": []
                    }
                    running_series_sum = 0

                elif mode == "team":
                    sel_key = f"sel_team_{box_num}"
                    if sel_key not in st.session_state or st.session_state[sel_key] not in team_codes:
                        st.session_state[sel_key] = team_codes[0]

                    el_code = st.session_state[sel_key]
                    el_info = db[el_code]

                    if os.path.exists(el_info["image"]):
                        st.image(el_info["image"], use_container_width=True)
                    st.caption(f"**No {el_code}** | **V {el_info['value']}**")

                    if st.button(T["pick_el_btn"], key=f"btn_t_{box_num}", use_container_width=True):
                        open_element_picker(sel_key, team_codes, is_indiv=False, box_num=box_num)

                    val = st.number_input(
                        T["val_lbl"],
                        min_value=0,
                        max_value=30,
                        value=int(el_info["value"]),
                        key=f"b_tval_{box_num}_{el_code}"
                    )
                    hold = st.text_input(
                        T["hold_lbl"],
                        value=el_info["hold"],
                        key=f"b_thold_{box_num}_{el_code}"
                    )
                    custom_el_no = st.text_input(
                        T["el_no_lbl"],
                        value=str(int(el_code)),
                        key=f"b_telno_{box_num}_{el_code}"
                    )

                    running_series_sum += val
                    total_team_points += val

                    if el_info["type"] == "balance":
                        team_balance_count += 1
                    elif el_info["type"] == "dynamic":
                        team_dynamic_count += 1

                    # Трудность парного/группового элемента строго во 2-й клеточке
                    boxes_data[box_num] = {
                        "mode": "team",
                        "hold": hold,
                        "d_cells": ["", str(val), "", ""],
                        "element_no": custom_el_no,
                        "images": [el_info["image"]]
                    }

                elif mode == "indiv":
                    indiv_box_count += 1
                    running_series_sum = 0
                    same_for_all = st.checkbox(
                        T["same_all_lbl"],
                        value=(num_partners > 2),
                        key=f"b_isame_{box_num}"
                    )

                    chosen_codes = []
                    chosen_vals = []
                    chosen_el_nos = []

                    if same_for_all:
                        sel_key = f"sel_iall_{box_num}"
                        if sel_key not in st.session_state or st.session_state[sel_key] not in indiv_codes:
                            st.session_state[sel_key] = indiv_codes[0]
                        c_code = st.session_state[sel_key]
                        c_info = db[c_code]

                        if os.path.exists(c_info["image"]):
                            st.image(c_info["image"], use_container_width=True)
                        st.caption(f"**No {c_code}** | **V {c_info['value']}**")

                        if st.button(T["pick_el_btn"], key=f"btn_ia_{box_num}", use_container_width=True):
                            open_element_picker(sel_key, indiv_codes, is_indiv=True, box_num=box_num)

                        val_ind = st.number_input(
                            T["val_lbl"],
                            min_value=0,
                            max_value=30,
                            value=int(c_info["value"]),
                            key=f"b_ival_all_{box_num}_{c_code}"
                        )
                        custom_el_no = st.text_input(
                            T["el_no_lbl"],
                            value=str(int(c_code)),
                            key=f"b_ielno_all_{box_num}_{c_code}"
                        )

                        chosen_codes = [c_code] * num_partners
                        chosen_vals = [val_ind] * num_partners
                        chosen_el_nos = [custom_el_no] * num_partners
                    else:
                        for p_i, role in enumerate(partner_roles):
                            sel_key = f"sel_ip_{box_num}_{p_i}"
                            if sel_key not in st.session_state or st.session_state[sel_key] not in indiv_codes:
                                st.session_state[sel_key] = indiv_codes[0]
                            c_code = st.session_state[sel_key]
                            c_info = db[c_code]

                            st.caption(f"**{role}**: No {c_code} (V {c_info['value']})")
                            if st.button(f"{role}: No {c_code}", key=f"btn_ip_{box_num}_{p_i}", use_container_width=True):
                                open_element_picker(sel_key, indiv_codes, is_indiv=True, box_num=box_num, role_title=role)

                            val_p = st.number_input(
                                f"{T['val_lbl']} ({role})",
                                min_value=0,
                                max_value=30,
                                value=int(c_info["value"]),
                                key=f"b_ival_p_{box_num}_{p_i}_{c_code}"
                            )
                            el_no_p = st.text_input(
                                f"{T['el_no_lbl']} ({role})",
                                value=str(int(c_code)),
                                key=f"b_ielno_p_{box_num}_{p_i}_{c_code}"
                            )

                            chosen_codes.append(c_code)
                            chosen_vals.append(val_p)
                            chosen_el_nos.append(el_no_p)

                    total_indiv_sum_all_partners += sum(chosen_vals)

                    d_cells = ["", "", "", ""]
                    for p_i in range(num_partners):
                        d_cells[p_i] = str(chosen_vals[p_i])

                    if len(set(chosen_codes)) == 1:
                        imgs = [db[chosen_codes[0]]["image"]]
                    else:
                        imgs = [db[c]["image"] for c in chosen_codes]

                    if len(set(chosen_el_nos)) == 1:
                        el_no_str = chosen_el_nos[0]
                    else:
                        el_no_str = " / ".join(chosen_el_nos)

                    holds = [db[c]["hold"] for c in chosen_codes if db[c]["hold"]]
                    default_hold_str = holds[0] if holds else ""
                    hold_str = st.text_input(
                        T["hold_lbl"],
                        value=default_hold_str,
                        key=f"b_ihold_{box_num}_{'_'.join(chosen_codes)}"
                    )

                    boxes_data[box_num] = {
                        "mode": "individual",
                        "hold": hold_str,
                        "d_cells": d_cells,
                        "element_no": el_no_str,
                        "images": imgs
                    }


indiv_total_points = round(total_indiv_sum_all_partners / num_partners)
total_d_puani = total_team_points + indiv_total_points
total_d_skoru = total_d_puani / 100.0

st.divider()

col_stat1, col_stat2, col_stat3 = st.columns([1.2, 1.5, 1.3])

with col_stat1:
    st.markdown(f"<div class='soft-panel'><h4>{T['calc_hdr']}</h4>"
                f"<p>{T['team_pts']}: <b>{total_team_points}</b></p>"
                f"<p>{T['indiv_pts']} ({total_indiv_sum_all_partners} / {num_partners}): <b>{indiv_total_points}</b></p>"
                f"<h3>D-Puanı: {total_d_puani} &nbsp;|&nbsp; D-Skoru: {total_d_skoru:.2f}</h3></div>",
                unsafe_allow_html=True)

with col_stat2:
    req_team = 4 if age_group == "6-12" else (5 if age_group.startswith("9-15") and "Grubu" in category else 6)
    total_team_done = team_balance_count + team_dynamic_count

    c1_cls = "status-ok" if total_team_done == req_team else "status-warn"
    c1_txt = f"{T['rule_count_ok']} ({total_team_done}/{req_team})" if total_team_done == req_team else f"{T['rule_count_warn']} {req_team} ({total_team_done})"

    c2_cls = "status-ok" if (team_balance_count >= 2 and team_dynamic_count >= 2) else "status-warn"
    c2_txt = f"{T['rule_bd_ok']} ({team_balance_count}B + {team_dynamic_count}D)" if (team_balance_count >= 2 and team_dynamic_count >= 2) else f"{T['rule_bd_warn']} ({team_balance_count}B + {team_dynamic_count}D)"

    c3_cls = "status-ok" if indiv_box_count == 3 else "status-warn"
    c3_txt = T["rule_ind_ok"] if indiv_box_count == 3 else f"{T['rule_ind_warn']} {indiv_box_count}/3"

    st.markdown(
        f"<div class='soft-panel'><h4>{T['rules_hdr']}</h4>"
        f"<div class='{c1_cls}'>{c1_txt}</div>"
        f"<div class='{c2_cls}'>{c2_txt}</div>"
        f"<div class='{c3_cls}'>{c3_txt}</div></div>",
        unsafe_allow_html=True
    )

with col_stat3:
    with st.container(border=True):
        st.markdown(f"#### {T['dl_hdr']}")
        meta_dict = {
            "comp_type": comp_type,
            "age_group": age_group,
            "category": category,
            "athletes": athletes,
            "club_name": club_name,
            "coach_phone": coach_phone,
            "team_no": team_no
        }
        docx_stream = generate_tariff_docx(meta_dict, boxes_data, total_d_puani, total_d_skoru)
        safe_club = club_name.replace(" ", "_") if club_name else "Team"
        file_name = f"Tariff_Sheet_{comp_type}_{category}_{age_group}_{safe_club}.docx"

        st.download_button(
            label=T["dl_btn"],
            data=docx_stream,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        st.caption(T["dl_note"])

st.divider()

st.markdown(f"#### {T['preview_hdr']}")
for row_start in (1, 8):
    p_cols = st.columns(7)
    for i in range(7):
        b_num = row_start + i
        b_info = boxes_data[b_num]
        with p_cols[i]:
            st.markdown(
                f"<div class='preview-card'>"
                f"<div class='preview-hdr'><span>{b_num}</span><span>{b_info['hold']}</span></div>",
                unsafe_allow_html=True
            )
            if b_info["mode"] == "empty":
                st.caption(T["mode_empty"])
            elif b_info["mode"] == "separator":
                st.markdown(f"### // {b_info.get('sep_sum', '')}")
            else:
                buf = combine_images_vertically(b_info["images"], target_width=240, target_height=290)
                if buf:
                    st.image(buf, use_container_width=True)
                d_str = " | ".join([x if x != "" else "·" for x in b_info["d_cells"]])
                st.caption(f"D: {d_str}  \nNO: {b_info['element_no']}")
            st.markdown("</div>", unsafe_allow_html=True)
