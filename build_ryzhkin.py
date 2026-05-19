# -*- coding: utf-8 -*-
"""
Сборка БП Рыжкина А.А. (двери, Рязань, 350к грант) на template Воробьёва.
Запуск:
  python3 build_ryzhkin.py
Создаёт файл: "Бизнес план Рыжкин.docx" в корне репо.

Подход: in-place edit document.xml Воробьёва — заменяем тексты
параграфов и ячеек таблиц, добавляем строки где нужно.
Все стили (шрифты, размеры, отступы, ширины колонок) сохраняются.
"""
import os, shutil, sys, json, zipfile, copy, re
import xml.etree.ElementTree as ET

ROOT = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = "/tmp/vor_tpl"  # распакованный Воробьёв
TEMPLATE_DOCX = os.path.join(ROOT, "Бизнес план Воробьев Павел Алексеевич.docx")
OUT_FILE = os.path.join(ROOT, "Бизнес план Рыжкин.docx")
WORK_DIR = "/tmp/ryzhkin_build"

# Гарантируем, что распакованный Воробьёв есть на месте
if not os.path.isdir(TEMPLATE_DIR):
    os.makedirs(TEMPLATE_DIR, exist_ok=True)
    with zipfile.ZipFile(TEMPLATE_DOCX) as z:
        z.extractall(TEMPLATE_DIR)

# ============================================================
# 0. NAMESPACES
# ============================================================
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + NS_W + "}"
# Регистрируем все часто используемые namespaces, чтобы ET сохранял префиксы
NAMESPACES = {
    'wpc':  "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    'cx':   "http://schemas.microsoft.com/office/drawing/2014/chartex",
    'mc':   "http://schemas.openxmlformats.org/markup-compatibility/2006",
    'o':    "urn:schemas-microsoft-com:office:office",
    'r':    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    'm':    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    'v':    "urn:schemas-microsoft-com:vml",
    'wp14': "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    'wp':   "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    'w10':  "urn:schemas-microsoft-com:office:word",
    'w':    NS_W,
    'w14':  "http://schemas.microsoft.com/office/word/2010/wordml",
    'w15':  "http://schemas.microsoft.com/office/word/2012/wordml",
    'w16cex': "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    'w16cid': "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    'w16':    "http://schemas.microsoft.com/office/word/2018/wordml",
    'w16sdtdh': "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    'w16se': "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    'wpg':  "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    'wpi':  "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    'wne':  "http://schemas.microsoft.com/office/word/2006/wordml",
    'wps':  "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


# ============================================================
# 1. ВХОДНЫЕ ДАННЫЕ
# ============================================================
PERSON = {
    "fio_full": "Рыжкин Артемий Алексеевич",
    "fio_short": "А.А. Рыжкин",
    "birth": "27.12.2004 г.",
    "addr_reg": "390011, г. Рязань, ул. Старообрядческий проезд, д. 11, кв. 33",
    "addr_fact": "г. Рязань, ул. Московская, д. 35, кв. 84",
    "phone": "89308891018",
    "email": "artemryzkin80@gmail.com",
    "inn": "622402008161",
    "okved": "43.32 - Работы столярные и плотничные.",
}

# ============================================================
# 2. СМЕТА (15 позиций оборудования + 1 продвижение)
# ============================================================
EQUIPMENT = [
    # (наименование, кол, цена_за_ед, поставщик, источник)
    ("Фрезер Makita RT0702CX2", 1, 49607, "Все инструменты", "grant"),
    ("Торцовочная пила Makita LS0714N", 1, 55900, "Все инструменты", "grant"),
    ("Перфоратор Makita HR2470", 1, 22490, "Все инструменты", "grant"),
    ("Дрель-шуруповёрт DeWalt DCD771C2", 1, 42347, "Все инструменты", "grant"),
    ("Лобзик Bosch GST 90 BE", 1, 29993, "Все инструменты", "grant"),
    ("Угловая шлифмашина Makita 9558HN", 1, 22260, "Все инструменты", "grant"),
    ("Лазерный нивелир ADA Cube 360 Basic", 1, 14482, "Все инструменты", "grant"),
    ("Шаблон для врезки петель и замков «Платинум Павла Солдатова»", 1, 22466, "Все инструменты", "grant"),
    ("Стусло Stanley поворотное", 1, 7307, "Все инструменты", "grant"),
    ("Уровень магнитный Stabila Type R-Tec 196 (1 м)", 1, 19750, "Все инструменты", "grant"),
    ("Промышленный пылесос Makita VC2512L", 1, 24037, "Все инструменты", "grant"),
    ("Комплект аккумуляторов и ЗУ (Ryobi 18 В, Makita 12 В, DeWalt 18 В)", 1, 16081, "Все инструменты", "grant"),
    ("Комплект пильных дисков CMT 190×30 (по дереву и металлу)", 1, 21231, "Все инструменты", "grant"),
    ("Кейсы для инструмента систейнер Tanos (2 шт.)", 1, 7652, "Все инструменты", "mixed"),  # переходная
    ("Стартовый запас расходников (монтажная пена, метизы, анкера, монтажный клей, затирка)", 1, 10000, "Леруа Мерлен", "own"),
]
PROMO_ITEM = ("Авито (платное продвижение объявлений на старте)", 1, 6000, "Авито", "own")

GRANT = 350000
SMETA_TOTAL = sum(c * p for _, c, p, _, _ in EQUIPMENT) + PROMO_ITEM[1] * PROMO_ITEM[2]
EQUIP_TOTAL = sum(c * p for _, c, p, _, _ in EQUIPMENT)
PROMO_TOTAL = PROMO_ITEM[1] * PROMO_ITEM[2]
OWN = SMETA_TOTAL - GRANT

sum_full_grant = sum(c * p for _, c, p, _, src in EQUIPMENT[:13])
mixed_pos_total = EQUIPMENT[13][1] * EQUIPMENT[13][2]
mixed_grant = GRANT - sum_full_grant
mixed_own = mixed_pos_total - mixed_grant

# ============================================================
# 3. ФИНМОДЕЛЬ — 5.1 (помесячно по услугам)
# ============================================================
# Порядок месяцев у Воробьёва: май, июн, июл, авг, сен, окт, ноя, дек, янв, фев, мар, апр
# Старт в мае (закупка), пик в сентябре-октябре и марте-апреле,
# спад в январе-феврале.
SERVICES = [
    # (название_длинное_для_2.1.1, название_короткое_для_5.1, время, цена, [12 значений])
    ("Установка межкомнатной двери под ключ (МДФ/экошпон): врезка петель и магнитного замка, монтаж наличников и доборов",
     "Установка межкомнатной двери под ключ", "3 ч. 30 мин.", 3500,
     [0, 9, 11, 12, 14, 13, 10, 8, 5, 7, 14, 12]),
    ("Монтаж скрытой двери Invisible: закладка алюминиевого короба + чистовая навеска полотна (двухэтапный)",
     "Монтаж скрытой двери Invisible", "6 часов", 6500,
     [0, 1, 2, 2, 3, 3, 1, 1, 0, 1, 3, 2]),
    ("Установка входной металлической двери: анкерный крепёж, регулировка уплотнителей и доводчика",
     "Установка входной металлической двери", "3 часа", 3500,
     [0, 2, 3, 3, 4, 3, 2, 2, 1, 2, 4, 3]),
    ("Установка комплекта доборных элементов и расширение/сужение проёма",
     "Установка доборов и расширение проёма", "1.5 часа", 1500,
     [0, 7, 9, 10, 11, 10, 7, 6, 3, 5, 12, 10]),
    ("Монтаж скрытого (теневого) алюминиевого плинтуса в квартире до 30 пог. м",
     "Монтаж скрытого плинтуса до 30 пог. м", "5 часов", 7500,
     [0, 1, 2, 2, 3, 2, 1, 1, 0, 1, 3, 2]),
]

REV_MONTHS = [0] * 12
for _, _, _, price, qty in SERVICES:
    for i, q in enumerate(qty):
        REV_MONTHS[i] += q * price
REV_TOTAL = sum(REV_MONTHS)

# ============================================================
# 4. ФИНМОДЕЛЬ — 5.2 (расходы / прибыль / ЧДП)
# ============================================================
EXP_MATERIALS = [0, 5400, 7200, 8100, 9300, 8800, 6700, 7400, 4500, 6200, 9700, 8200]
EXP_TRANSPORT = [0, 1500, 1700, 1800, 2200, 2000, 1500, 1500, 1200, 1300, 2200, 1800]
EXP_COMMS     = [0, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800]
EXP_PROMO     = [0, 2000, 2200, 2200, 2800, 2500, 1800, 1800, 1500, 1700, 2800, 2200]
EXP_UNFORESEEN= [0, 300, 400, 400, 500, 400, 300, 300, 200, 300, 500, 400]
EXP_TAX       = [round(r * 0.04) for r in REV_MONTHS]
EXP_EQUIPMENT = [SMETA_TOTAL] + [0] * 11
EXP_RENT      = [0] * 12

EXPENSES_BY_MONTH = [
    EXP_EQUIPMENT[i] + EXP_MATERIALS[i] + EXP_RENT[i] + EXP_TRANSPORT[i] + EXP_COMMS[i]
    + EXP_PROMO[i] + EXP_UNFORESEEN[i] + EXP_TAX[i]
    for i in range(12)
]
PROFIT_BY_MONTH = [REV_MONTHS[i] - EXPENSES_BY_MONTH[i] for i in range(12)]

EXTERNAL = [GRANT] + [0] * 11
OWNFUND  = [OWN] + [0] * 11

CFLOW = [PROFIT_BY_MONTH[i] + EXTERNAL[i] + OWNFUND[i] for i in range(12)]
CFLOW_CUM = []
acc = 0
for v in CFLOW:
    acc += v
    CFLOW_CUM.append(acc)

EXP_TOTAL = sum(EXPENSES_BY_MONTH)
PROFIT_TOTAL = REV_TOTAL - EXP_TOTAL
RENTABILITY = PROFIT_TOTAL / EXP_TOTAL * 100

PAYBACK = 12
for i, v in enumerate(CFLOW_CUM):
    if v >= SMETA_TOTAL:
        PAYBACK = i + 1
        break

YEAR2_REVENUE = 1020000
YEAR2_EXPENSES = 290000
YEAR2_PROFIT = YEAR2_REVENUE - YEAR2_EXPENSES
YEAR2_RENT = YEAR2_PROFIT / YEAR2_EXPENSES * 100

# ============================================================
# 5. АУДИТ
# ============================================================
def audit():
    err = []
    s_check = sum(c * p for _, c, p, _, _ in EQUIPMENT) + PROMO_ITEM[1] * PROMO_ITEM[2]
    if s_check != SMETA_TOTAL: err.append(f"AUDIT-1: сумма {s_check} != ИТОГО {SMETA_TOTAL}")
    if GRANT + OWN != SMETA_TOTAL: err.append("AUDIT-2: grant+own != total")
    if GRANT != 350000: err.append("AUDIT-2b: grant != 350 000")
    if mixed_grant <= 0 or mixed_own <= 0: err.append("AUDIT-3: переходная позиция")
    if mixed_grant + mixed_own != mixed_pos_total: err.append("AUDIT-3b")
    for i, (_, _, _, _, src) in enumerate(EQUIPMENT[:13]):
        if src != "grant": err.append(f"AUDIT-4: 1.{i+1}")
    if EQUIPMENT[14][4] != "own": err.append("AUDIT-4b")
    if not (368000 <= SMETA_TOTAL <= 372000): err.append(f"X1: {SMETA_TOTAL}")
    if not (850000 <= REV_TOTAL <= 1150000): err.append(f"X2: выручка {REV_TOTAL}")
    if EXP_EQUIPMENT[0] != SMETA_TOTAL: err.append("X4")
    if sum(EXTERNAL) != GRANT: err.append("X5")
    if sum(OWNFUND) != OWN: err.append("X5b")
    if CFLOW[0] != 0: err.append(f"X6: ЧДП м1 {CFLOW[0]}")
    if PROFIT_TOTAL != sum(PROFIT_BY_MONTH): err.append("X7")
    if not (40 <= RENTABILITY <= 70): err.append(f"X8: {RENTABILITY:.1f}%")
    for m in range(12):
        hrs = (SERVICES[0][4][m]*3.5 + SERVICES[1][4][m]*6 + SERVICES[2][4][m]*3
               + SERVICES[3][4][m]*1.5 + SERVICES[4][4][m]*5)
        if hrs > 160: err.append(f"X9: мес {m+1} ч {hrs}")
    if not (5 <= PAYBACK <= 9): err.append(f"X10: {PAYBACK}")
    for m in range(12):
        s = (EXP_EQUIPMENT[m] + EXP_MATERIALS[m] + EXP_RENT[m] + EXP_TRANSPORT[m]
             + EXP_COMMS[m] + EXP_PROMO[m] + EXP_UNFORESEEN[m] + EXP_TAX[m])
        if s != EXPENSES_BY_MONTH[m]: err.append(f"X11: мес {m+1}")
    if CFLOW_CUM[-1] != PROFIT_TOTAL + GRANT + OWN: err.append("X12")
    return err

errors = audit()
print("=" * 60)
print(f"СМЕТА: {SMETA_TOTAL:,} ₽ (грант {GRANT:,} + свои {OWN:,})".replace(",", " "))
print(f"Переходная 1.14: грант {mixed_grant} + свои {mixed_own} = {mixed_pos_total}")
print(f"ВЫРУЧКА: {REV_TOTAL:,} ₽".replace(",", " "))
print(f"РАСХОДЫ: {EXP_TOTAL:,} ₽".replace(",", " "))
print(f"ПРИБЫЛЬ: {PROFIT_TOTAL:,} ₽".replace(",", " "))
print(f"ЧДП нараст м12: {CFLOW_CUM[-1]:,} ₽".replace(",", " "))
print(f"Рентабельность: {RENTABILITY:.1f}%, окупаемость: {PAYBACK} мес")
print("=" * 60)
if errors:
    print("ОШИБКИ АУДИТА:")
    for e in errors: print(" -", e)
    sys.exit(1)
print("АУДИТ: OK")
print("=" * 60)


# ============================================================
# 6. РАСПАКОВКА TEMPLATE И ПОДГОТОВКА К ЗАМЕНЕ
# ============================================================
if os.path.exists(WORK_DIR):
    shutil.rmtree(WORK_DIR)
shutil.copytree(TEMPLATE_DIR, WORK_DIR)

DOC_PATH = os.path.join(WORK_DIR, "word", "document.xml")

# Регистрируем mc:Ignorable перед парсингом
tree = ET.parse(DOC_PATH)
root = tree.getroot()
body = root.find(f"{W}body")


# ============================================================
# 7. ХЕЛПЕРЫ ДЛЯ ПРАВКИ XML
# ============================================================
def fmt(n):
    return f"{n:,}".replace(",", " ")


def get_runs(p_el):
    """Все <w:r> на верхнем уровне параграфа."""
    return p_el.findall(f"{W}r")


def get_first_rpr(p_el):
    """Найти первый rPr параграфа (либо в первом r, либо в pPr/rPr) для копирования стилей."""
    runs = get_runs(p_el)
    for r in runs:
        rpr = r.find(f"{W}rPr")
        if rpr is not None:
            return rpr
    pPr = p_el.find(f"{W}pPr")
    if pPr is not None:
        rpr = pPr.find(f"{W}rPr")
        if rpr is not None:
            return rpr
    return None


def set_para_text(p_el, new_text):
    """Заменить текст параграфа, сохраняя rPr первого run и pPr.

    Алгоритм:
      1. Запомнить rPr из первого <w:r> (или из pPr).
      2. Удалить все <w:r> и <w:hyperlink> в параграфе.
      3. Создать один новый <w:r> с тем же rPr и одним <w:t> = new_text.
    """
    # сохраняем pPr на месте, удаляем всё остальное
    pPr = p_el.find(f"{W}pPr")
    rpr_template = get_first_rpr(p_el)

    # удалить всё кроме pPr
    for child in list(p_el):
        if child.tag != f"{W}pPr":
            p_el.remove(child)

    # создать run
    r = ET.SubElement(p_el, f"{W}r")
    if rpr_template is not None:
        r.append(copy.deepcopy(rpr_template))
    t = ET.SubElement(r, f"{W}t")
    t.text = new_text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def get_first_para(tc_el):
    """Первый <w:p> внутри ячейки."""
    return tc_el.find(f"{W}p")


def set_cell_text(tc_el, new_text, *, bold=None):
    """Заменить текст ячейки. Сохраняет первый параграф, удаляет остальные.

    bold=True/False/None — управление жирным; None = оставить как у template.
    """
    paras = tc_el.findall(f"{W}p")
    if not paras:
        # создать пустой параграф
        p = ET.SubElement(tc_el, f"{W}p")
        paras = [p]

    # очистить все, кроме первого
    for p in paras[1:]:
        tc_el.remove(p)

    p = paras[0]
    set_para_text(p, new_text)
    # bold
    if bold is True:
        r = p.find(f"{W}r")
        if r is not None:
            rpr = r.find(f"{W}rPr")
            if rpr is None:
                rpr = ET.SubElement(r, f"{W}rPr")
                # переместить в начало run
                r.remove(rpr)
                r.insert(0, rpr)
            if rpr.find(f"{W}b") is None:
                ET.SubElement(rpr, f"{W}b")
    elif bold is False:
        r = p.find(f"{W}r")
        if r is not None:
            rpr = r.find(f"{W}rPr")
            if rpr is not None:
                b = rpr.find(f"{W}b")
                if b is not None:
                    rpr.remove(b)


def get_cells(row_el):
    return row_el.findall(f"{W}tc")


def clone_row(row_el):
    """Глубокая копия строки (для добавления новых строк в таблицу)."""
    return copy.deepcopy(row_el)


def find_p_by_prefix(body, prefix):
    """Найти первый <w:p>, текст которого начинается с prefix."""
    for el in body.iter(f"{W}p"):
        txt = ''.join(t.text or '' for t in el.iter(f"{W}t"))
        if txt.startswith(prefix):
            return el
    return None


# ============================================================
# 8. ПРАВКА ПАРАГРАФОВ ТИТУЛА И РЕЗЮМЕ (по индексу + по префиксу)
# ============================================================
# Берём элементы тела по индексу как у Воробьёва (сверка структурно совпадает)
elems = list(body)

def by_idx(i):
    return elems[i]

# ТИТУЛ
set_para_text(by_idx(17), "Проект «Профессиональная установка межкомнатных и входных дверных блоков, монтаж скрытых дверных систем Invisible и теневых алюминиевых плинтусов»")
set_para_text(by_idx(28), "                               " + PERSON["fio_full"])
set_para_text(by_idx(29), "Адрес: " + PERSON["addr_reg"])
set_para_text(by_idx(30), "Телефон: " + PERSON["phone"])
set_para_text(by_idx(31), "                                                                                      E-mail: " + PERSON["email"])

# РЕЗЮМЕ 1.1—1.15
set_para_text(by_idx(46), "1.1. " + PERSON["fio_full"])
set_para_text(by_idx(47), "1.2. Дата рождения: " + PERSON["birth"])
set_para_text(by_idx(48), "1.3. Адрес регистрации: " + PERSON["addr_reg"] + ".")
set_para_text(by_idx(49), "1.4. Адрес фактического проживания: " + PERSON["addr_fact"] + ".")
set_para_text(by_idx(50), "1.5. Контактные данные: Телефон: " + PERSON["phone"] + ", E-mail: " + PERSON["email"] + ".")
set_para_text(by_idx(51), "1.6. Семейное положение: не женат.")
set_para_text(by_idx(52), "1.7. Образование: Неоконченное высшее (студент 4 курса РГРТУ им. В.Ф. Уткина, направление 15.03.04). (Приложение 1)")
set_para_text(by_idx(53), "1.8. Общий стаж работы, трудовая деятельность в течение последних 3 лет: в период с 2023 по 2026 гг. официально не трудоустроен в связи с прохождением обучения по очной форме в высшем учебном заведении.")
set_para_text(by_idx(54), "1.9. Опыт работы (предпринимательской деятельности) по выбранному направлению: имею профильное техническое образование и наработанный неофициальный опыт участия в отделочных работах в составе ремонтных бригад г. Рязани — точная разметка проёмов лазерным нивелиром, врезка петель и замков по шаблону, установка межкомнатных и входных дверных блоков, устройство напольных плинтусов.")
set_para_text(by_idx(55), "1.10. Потребность в обучении/повышении квалификации с обоснованием: потребность в обучении отсутствует. Имеются необходимые практические навыки.")
set_para_text(by_idx(56), "1.11. Текущий статус автора проекта: Самозанятый.")
set_para_text(by_idx(57), "1.12. Организационно-правовая форма ведения организуемого бизнеса и система налогообложения: Самозанятость, система налогообложения — Налог на профессиональный доход (НПД). ИНН " + PERSON["inn"] + ".")
set_para_text(by_idx(58), "1.13. Виды деятельности (ОКВЭД): " + PERSON["okved"])
set_para_text(by_idx(59), "1.14. Наличие или необходимость лицензий на соответствующие виды деятельности, патентов, сертификатов, авторских прав, медицинской книжки и т.п. - не требуется.")
set_para_text(by_idx(60), "1.15. Месторасположение: Деятельность носит выездной характер. Хранение электроинструмента, прецизионной оснастки и расходных материалов осуществляется в имеющемся в собственности гараже. Аренда отдельного офиса не требуется.")

# 1.16. — параграфы заголовка остаются как у Воробьёва (62, 63)

# ============================================================
# T1 (1.16): 7 строк × 2 колонки
# ============================================================
t1 = by_idx(64)
t1_rows = t1.findall(f"{W}tr")
# row0 = шапка (Показатель | Значение) — оставляем
# row1—6 = значения
values_116 = [
    fmt(REV_TOTAL) + " руб.",
    fmt(EXP_TOTAL) + " руб.",
    fmt(PROFIT_TOTAL) + " руб.",
    fmt(CFLOW_CUM[-1]) + " руб.",
    f"{RENTABILITY:.1f}%".replace(".", ","),
    f"{PAYBACK} месяцев",
]
for i, v in enumerate(values_116, start=1):
    cells = get_cells(t1_rows[i])
    set_cell_text(cells[1], v)

# ============================================================
# 2.1 — текст
# ============================================================
set_para_text(by_idx(66),
    "2.1 Услуги по профессиональной установке межкомнатных и входных дверных блоков, монтажу скрытых дверных систем Invisible и устройству скрытого (теневого) алюминиевого плинтуса. Технологический процесс выстроен в три самостоятельных этапа: подготовка проёма с лазерной нивелировкой и контролем геометрии, прецизионная врезка петель и магнитных замков по шаблону «Платинум» с минимальным люфтом и идентичной посадкой фурнитуры на каждом полотне, чистовой монтаж полотен, доборных элементов и наличников с финальной отделкой стыков. Для скрытых дверных систем работа разбита на закладку алюминиевого короба на этапе чернового ремонта и чистовую навеску полотна после отделки стен. Все работы выполняются на объекте заказчика, аренда отдельного помещения не требуется."
)

# ============================================================
# T2 (2.1.1): 4 строки в template (1 заголовок + 3 услуги).
# Нам нужно 6 строк (1 + 5 услуг) → добавить 2 строки клонированием row1.
# ============================================================
t2 = by_idx(69)
t2_rows = t2.findall(f"{W}tr")
template_row = t2_rows[1]  # template для услуги
# Добавим столько строк, чтобы суммарно стало 1 + len(SERVICES) = 6
need_extra = (1 + len(SERVICES)) - len(t2_rows)
for _ in range(need_extra):
    new_row = clone_row(template_row)
    t2.append(new_row)
# Теперь заполним строки услуг (индексы 1..len(SERVICES))
t2_rows = t2.findall(f"{W}tr")
for idx_s, (full_name, _, time_text, price, _) in enumerate(SERVICES, start=1):
    cells = get_cells(t2_rows[idx_s])
    set_cell_text(cells[0], full_name)
    set_cell_text(cells[1], time_text)
    set_cell_text(cells[2], fmt(price))

# Текст после T2 (P70 = idx 70 пустой, P71 = idx 71 — описание услуг)
set_para_text(by_idx(71),
    "В состав каждой услуги входит выезд мастера на объект, защита прилегающих поверхностей плёнкой, разметка лазерным нивелиром, контроль геометрии проёма уровнем, врезка фурнитуры по шаблону «Платинум», финальная очистка пылеудалителем и сдача результата заказчику с фотоотчётом «до/после»."
)

# 2.2
set_para_text(by_idx(72),
    "2.2 Целевая аудитория: основная аудитория — собственники квартир в новостройках и вторичном жилом фонде г. Рязани и области, заходящие в этап чистового ремонта и заказывающие комплект из 4–6 межкомнатных дверей одновременно. Дополнительный сегмент — владельцы квартир в премиальных жилых комплексах с европланировкой, выбирающие скрытые дверные системы Invisible и теневые алюминиевые плинтуса как часть единого дизайн-кода. Третий сегмент — частные дизайнеры интерьеров и небольшие ремонтные бригады, нуждающиеся в подрядчике на узкий технологический участок «двери + плинтус» с фиксированным сроком и собственной прецизионной оснасткой."
)

# 2.3 — детальный, три уровня по оснащению
set_para_text(by_idx(73),
    "2.3. Конкурентная среда, конкурентная позиция бизнеса: проанализировав объявления сайта «Авито» по г. Рязани (Приложение 3), я выделил три уровня предложения. Первый — бригады при магазинах и салонах межкомнатных дверей: стоимость монтажа от 3 500 до 5 500 рублей за стандартное полотно, при этом скрытые системы и теневые плинтуса в их перечне отсутствуют. Второй — частные мастера с базовым набором инструмента: стоимость от 2 800 до 3 500 рублей за дверь, однако без шаблона врезки фурнитуры и лазерной нивелировки точность монтажа полотен с магнитными замками и скрытыми петлями нестабильна. Третий уровень — профессиональные установщики со специализированной оснасткой: фрезер, торцовочная пила, шаблон «Платинум», лазерный нивелир; их предложение составляет 4 500–6 500 рублей за стандартную дверь и 5 500–8 000 за скрытую систему. Моё преимущество — полный комплект прецизионной оснастки и компетенция по скрытым дверным системам Invisible и теневым алюминиевым плинтусам по средней рыночной цене."
)

# 2.4
set_para_text(by_idx(74),
    "2.4. Каналы реализации товара: платное продвижение профильных объявлений на платформах «Авито» и «Юла» в категории «Двери, окна» по г. Рязани и Рязанской области, профиль на сервисе «Профи.ру» с накоплением отзывов и портфолио, страница в социальных сетях с фотоотчётами «до/после», а также личные договорённости с менеджерами рязанских салонов межкомнатных дверей и подрядных ремонтных бригад."
)

# 2.8
set_para_text(by_idx(75),
    "2.8. Продвижение. Поиск клиентов будет вестись через прямую коллаборацию с дизайнерами интерьеров и ремонтными бригадами г. Рязани, работающими в сегменте премиальной отделки. Уникальное торговое предложение проекта — пакет «Скрытая дверь Invisible под ключ»: закладка короба на черновом этапе и чистовая навеска полотна после отделки в фиксированный срок без переделок и доборов за счёт мастера. Каждый завершённый объект сопровождается фотоотчётом и короткой видеовизитой, формирующими портфолио для дальнейшего органического трафика."
)

# 3.1 — 3.3 (находим по префиксу, чтобы не зависеть от точного индекса)
def replace_by_prefix(prefix, new_text):
    el = find_p_by_prefix(body, prefix)
    if el is not None:
        set_para_text(el, new_text)
    else:
        print(f"WARN: не найден параграф с префиксом '{prefix}'")

replace_by_prefix("3.1.",
    "3.1. с 09:00 до 21:00 без выходных. Выезд на объекты осуществляется по предварительной договорённости с заказчиком."
)
replace_by_prefix("3.2.",
    "3.2. Штатное расписание: дополнительные сотрудники не требуются."
)
replace_by_prefix("3.3.",
    "3.3. Требования к персоналу: требуются практические навыки прецизионной разметки проёмов, врезки петель и замков по шаблону, монтажа скрытых дверных систем Invisible и устройства алюминиевых плинтусов, владение электроинструментом и лазерной нивелировкой, что у меня имеется."
)

# ============================================================
# T3 (4.1): 6 строк × 3 колонки
# ============================================================
t3 = by_idx(83)
t3_rows = t3.findall(f"{W}tr")
# row1: грант
set_cell_text(get_cells(t3_rows[1])[2], fmt(GRANT))
# row2: свои
set_cell_text(get_cells(t3_rows[2])[2], fmt(OWN))
# row3, 4 — заёмные/иное оставляем 0
# row5: ИТОГО (gridSpan=2 первая ячейка → итого 2 cells)
set_cell_text(get_cells(t3_rows[5])[1], fmt(SMETA_TOTAL), bold=True)

# ============================================================
# T4 (4.2): 22 строки × 7 колонок
# row0 = шапка
# row1 = "1.* Основные средства..." — заголовок
# row2..16 = 1.1—1.15 (15 позиций)
# row17 = "2.*"
# row18 = "3.*"
# row19 = "4.*" — продвижение (у нас 6 000)
# row20 = "5.*"
# row21 = ИТОГО
# ============================================================
t4 = by_idx(87)
t4_rows = t4.findall(f"{W}tr")
SRC_LABEL = {
    "grant": "Средства соцконтракта",
    "own": "Собственные средства",
    "mixed": "Средства соцконтракта (" + fmt(mixed_grant) + " руб.) и собственные средства (" + fmt(mixed_own) + " руб.)",
}
# row1 (1.*) — заголовок раздела, обновим название
cells = get_cells(t4_rows[1])
set_cell_text(cells[1], "Основные средства (оборудование, инструменты, оснастка):", bold=True)

# row2..16 — позиции 1.1—1.15
for i, (name, qty, price, supplier, src) in enumerate(EQUIPMENT):
    row = t4_rows[2 + i]
    cells = get_cells(row)
    summa = qty * price
    set_cell_text(cells[0], f" 1.{i+1}")
    set_cell_text(cells[1], name)
    set_cell_text(cells[2], str(qty))
    set_cell_text(cells[3], fmt(price))
    set_cell_text(cells[4], fmt(summa))
    set_cell_text(cells[5], supplier)
    set_cell_text(cells[6], SRC_LABEL[src])

# row17 = "2.*"
cells = get_cells(t4_rows[17])
set_cell_text(cells[1], "Материально-производственные запасы (комплектующие, сырье и т.д.):", bold=True)
for i in (2, 3, 4): set_cell_text(cells[i], "")
set_cell_text(cells[5], "")
set_cell_text(cells[6], "")

# row18 = "3.*"
cells = get_cells(t4_rows[18])
set_cell_text(cells[1], "Имущественные обязательства (аренда, не более 15% общей суммы)", bold=True)
for i in (2, 3, 4): set_cell_text(cells[i], "")

# row19 = "4.*" — Продвижение, у нас 6 000
cells = get_cells(t4_rows[19])
set_cell_text(cells[1], "Продвижение (не более 5% общей суммы): " + PROMO_ITEM[0], bold=True)
set_cell_text(cells[2], str(PROMO_ITEM[1]))
set_cell_text(cells[3], fmt(PROMO_ITEM[2]))
set_cell_text(cells[4], fmt(PROMO_TOTAL))
set_cell_text(cells[5], PROMO_ITEM[3])
set_cell_text(cells[6], "Собственные средства")

# row20 = "5.*"
cells = get_cells(t4_rows[20])
set_cell_text(cells[1], "Подготовка разрешительной документации, приобретение ПО, электронной подписи (не более 10% общей суммы)", bold=True)
for i in (2, 3, 4): set_cell_text(cells[i], "")

# row21 = ИТОГО (gridSpan=2 первая, итого 6 cells: 0=ИТОГО, 1=Кол-во, 2=Цена, 3=Сумма, 4=Поставщик, 5=Источник)
cells = get_cells(t4_rows[21])
set_cell_text(cells[0], "ИТОГО", bold=True)
set_cell_text(cells[3], fmt(SMETA_TOTAL), bold=True)

# ============================================================
# 4.3
# ============================================================
set_para_text(by_idx(92),
    "В течение одного месяца с момента поступления денежных средств в рамках социального контракта проводится закупка полного спектра необходимого инструмента и оснастки. Порядок закупки: основной электроинструмент (фрезер, торцовочная пила, перфоратор, шуруповёрт, лобзик, угловая шлифмашина), прецизионная оснастка (лазерный нивелир, шаблон врезки «Платинум», уровень, стусло, промышленный пылесос), комплекты аккумуляторов, пильных дисков, систейнеров и стартовый запас расходных материалов; параллельно — запуск рекламной кампании на платформе «Авито»."
)

# ============================================================
# T5 (5.1): 18 строк × 16 колонок
# row0 = шапка (№п/п, Наименование, 1мес...12мес, ИТОГО, Доля%)
# row1 = месяцы (май..апр)
# row2 = "Доходы всего, руб., в т.ч.:"
# row3 = услуга 1: "1." | name | м1..м12 | ИТОГО | Доля
# row4 = "количество"
# row5 = "стоимость"
# row6, 9, 12, 15 — другие услуги
# row7, 10, 13, 16 — количества
# row8, 11, 14, 17 — стоимости
# ============================================================
t5 = by_idx(97)
t5_rows = t5.findall(f"{W}tr")

# row2: Доходы всего по месяцам
cells = get_cells(t5_rows[2])
# Cells: 0=№п/п(пусто), 1=label, 2..13=м1..м12, 14=ИТОГО, 15=Доля
for i in range(12):
    set_cell_text(cells[2 + i], fmt(REV_MONTHS[i]) if REV_MONTHS[i] else "0", bold=True)
set_cell_text(cells[14], fmt(REV_TOTAL), bold=True)
set_cell_text(cells[15], "100%", bold=True)

# Услуги
for s_idx, (full_name, short_name, time_text, price, qty) in enumerate(SERVICES):
    base_row = 3 + s_idx * 3
    if base_row + 2 >= len(t5_rows):
        break  # на всякий случай
    # Главная строка услуги: № | name | revenues | total | share
    cells_srv = get_cells(t5_rows[base_row])
    set_cell_text(cells_srv[0], f"{s_idx + 1}.")
    set_cell_text(cells_srv[1], short_name)
    rev_total = sum(qty) * price
    share = rev_total / REV_TOTAL * 100
    for i in range(12):
        rev_m = qty[i] * price
        set_cell_text(cells_srv[2 + i], fmt(rev_m) if rev_m else "0")
    set_cell_text(cells_srv[14], fmt(rev_total))
    set_cell_text(cells_srv[15], f"{share:.1f}%".replace(".", ","))
    # Строка "количество"
    cells_qty = get_cells(t5_rows[base_row + 1])
    # row[1] = "количество" — оставляем
    for i in range(12):
        set_cell_text(cells_qty[2 + i], str(qty[i]) if qty[i] else "0")
    set_cell_text(cells_qty[14], str(sum(qty)))
    # Строка "стоимость"
    cells_cost = get_cells(t5_rows[base_row + 2])
    for i in range(12):
        set_cell_text(cells_cost[2 + i], fmt(price))

# Текст после T5 (idx 99)
set_para_text(by_idx(99),
    "На доходы от услуг по установке дверей и теневых плинтусов влияет сезонный ремонтный цикл: пик активности приходится на сентябрь–октябрь и март–апрель, когда массово сдаются новостройки и собственники переходят к чистовой отделке; в декабре–феврале наблюдается умеренный спад. Первый месяц проекта (май) целиком отводится под закупку инструмента и оснастки, заявки начинают поступать со второго месяца после публикации профильных объявлений и завершения первых эталонных объектов."
)

# ============================================================
# T6 (5.2): 18 строк × 15 колонок
# row0 = шапка, row1 = месяцы
# row2 = "1 Доходы"
# row3 = "2 Расходы"
# row4..11 = "2.1 .. 2.8"
# row12 = "3 Прибыль"
# row13 = "4 ЧДП"
# row14 = "4.1 нараст"
# row15 = "5 Рентабельность"
# row16 = "6 Стороннее"
# row17 = "7 Собственные"
# ============================================================
t6 = by_idx(104)
t6_rows = t6.findall(f"{W}tr")

def fill_row(row_idx, vals, total, *, bold=False):
    cells = get_cells(t6_rows[row_idx])
    for i in range(12):
        v = vals[i]
        text = fmt(v) if v != 0 else "0"
        set_cell_text(cells[2 + i], text, bold=bold)
    set_cell_text(cells[14], fmt(total), bold=True)

fill_row(2, REV_MONTHS, REV_TOTAL, bold=True)
fill_row(3, EXPENSES_BY_MONTH, EXP_TOTAL, bold=True)
fill_row(4, EXP_EQUIPMENT, sum(EXP_EQUIPMENT))
fill_row(5, EXP_MATERIALS, sum(EXP_MATERIALS))
fill_row(6, EXP_RENT, 0)
fill_row(7, EXP_TRANSPORT, sum(EXP_TRANSPORT))
fill_row(8, EXP_COMMS, sum(EXP_COMMS))
fill_row(9, EXP_PROMO, sum(EXP_PROMO))
fill_row(10, EXP_UNFORESEEN, sum(EXP_UNFORESEEN))
fill_row(11, EXP_TAX, sum(EXP_TAX))
fill_row(12, PROFIT_BY_MONTH, PROFIT_TOTAL, bold=True)
fill_row(13, CFLOW, sum(CFLOW), bold=True)
fill_row(14, CFLOW_CUM, CFLOW_CUM[-1])
# Рентабельность — только в ИТОГО
cells_rent = get_cells(t6_rows[15])
for i in range(12):
    set_cell_text(cells_rent[2 + i], "")
set_cell_text(cells_rent[14], f"{RENTABILITY:.1f}%".replace(".", ","), bold=True)
fill_row(16, EXTERNAL, sum(EXTERNAL))
fill_row(17, OWNFUND, sum(OWNFUND))

# Текст после T6 (idx 106)
set_para_text(by_idx(106),
    f"В первый месяц расходы составляют {fmt(SMETA_TOTAL)} руб. (закупка всего необходимого инструмента, оснастки и стартовая реклама), которые полностью покрываются средствами социального контракта (350 000 руб.) и собственными средствами ({fmt(OWN)} руб.), что обеспечивает нулевой чистый денежный поток в первом месяце. В расходные материалы включены монтажная пена, метизы и анкера, клей-герметик для скрытых дверных систем, сменные пильные диски и абразивы, наличники и доборы для образцов, средства индивидуальной защиты."
)

# 5.3 (idx 109)
months_word = ['', 'один месяц', 'два месяца', 'три месяца', 'четыре месяца', 'пять месяцев',
               'шесть месяцев', 'семь месяцев', 'восемь месяцев', 'девять месяцев',
               'десять месяцев', 'одиннадцать месяцев', 'двенадцать месяцев']
set_para_text(by_idx(109), f"5.3. Срок окупаемости проекта: {months_word[PAYBACK]}.")

# 5.4 заголовок (idx 111) оставляем
# T7 (5.4): 5 строк × 4 колонки
t7 = by_idx(112)
t7_rows = t7.findall(f"{W}tr")
# row1: Доходы | 1й | 2й
set_cell_text(get_cells(t7_rows[1])[2], fmt(REV_TOTAL) + " руб.")
set_cell_text(get_cells(t7_rows[1])[3], fmt(YEAR2_REVENUE) + " руб.")
# row2: Расходы
set_cell_text(get_cells(t7_rows[2])[2], fmt(EXP_TOTAL) + " руб.")
set_cell_text(get_cells(t7_rows[2])[3], fmt(YEAR2_EXPENSES) + " руб.")
# row3: Прибыль
set_cell_text(get_cells(t7_rows[3])[2], fmt(PROFIT_TOTAL) + " руб.")
set_cell_text(get_cells(t7_rows[3])[3], fmt(YEAR2_PROFIT) + " руб.")
# row4: Рентабельность
set_cell_text(get_cells(t7_rows[4])[2], f"{RENTABILITY:.1f}%".replace(".", ","))
set_cell_text(get_cells(t7_rows[4])[3], f"{YEAR2_RENT:.0f}%")

# ============================================================
# T8 (6): 8 строк × 3 колонки
# ============================================================
RISKS = [
    ("Снижение платежеспособности потребителей", "Низкая",
     "Комплектные пакеты «дверь + доборы + плинтус под ключ» дешевле раздельного заказа у разных мастеров."),
    ("Эпидемиологическая обстановка. Риск заболеваемости", "Низкая",
     "Вакцинация и работа в защитных перчатках и респираторе при пылеобразующих операциях."),
    ("Уменьшение потока клиентов (сезонность)", "Низкая",
     "В межсезонье акцент на коммерческих помещениях и скрытых дверных системах Invisible с длинным циклом продажи."),
    ("Несоответствие ожиданиям клиента", "Низкая",
     "Согласование контрольных точек по фотографиям после закладки короба и до чистового монтажа."),
    ("Болезнь", "Низкая",
     "Поддерживаю режим, регулярные осмотры, страховой запас по срокам сдачи объектов."),
    ("Появление конкурентов", "Средняя",
     "Удержание качества за счёт прецизионной оснастки и накопленного портфолио скрытых дверей."),
    ("Поломка оборудования", "Низкая",
     "Транспортировка фрезера и торцовочной пилы в защитных кейсах систейнер, ежемесячная очистка коллектора пылеудаления."),
]
t8 = by_idx(115)
t8_rows = t8.findall(f"{W}tr")
for i, (risk, prob, sol) in enumerate(RISKS, start=1):
    cells = get_cells(t8_rows[i])
    set_cell_text(cells[0], risk)
    set_cell_text(cells[1], prob)
    set_cell_text(cells[2], sol)

# ============================================================
# 7. Подпись (idx 123) — у Воробьёва "Воробь..." (длинный пробел)
# Ищем параграф с "Воробь" и заменяем на "А.А. Рыжкин"
# ============================================================
for el in body.iter(f"{W}p"):
    txt = ''.join(t.text or '' for t in el.iter(f"{W}t"))
    if "Воробь" in txt:
        # Сохраняем правый отступ — оставляем хвост пробелов перед именем
        # Но проще: заменяем весь текст на "А.А. Рыжкин"
        set_para_text(el, "                                                                                                                  " + PERSON["fio_short"])
        break

# ============================================================
# 9. СОХРАНИТЬ document.xml
# ============================================================
# ET не пишет standalone, делаем вручную через tostring
xml_bytes = ET.tostring(root, encoding='utf-8', xml_declaration=False)
with open(DOC_PATH, 'wb') as f:
    f.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    f.write(xml_bytes)

# Удаляем header1.xml (там подпись Воробьёва как watermark) - не трогаем, у Воробьёва пустой
# Просто оставляем как есть

# ============================================================
# 10. УПАКОВКА В DOCX
# ============================================================
if os.path.exists(OUT_FILE):
    os.remove(OUT_FILE)


def add_to_zip(zf, root, base):
    for name in sorted(os.listdir(root)):
        full = os.path.join(root, name)
        if os.path.isdir(full):
            add_to_zip(zf, full, base)
        else:
            rel = os.path.relpath(full, base)
            zf.write(full, rel)


with zipfile.ZipFile(OUT_FILE, "w", zipfile.ZIP_DEFLATED) as zf:
    add_to_zip(zf, WORK_DIR, WORK_DIR)

print(f"\nГотово: {OUT_FILE}")
print(f"Размер: {os.path.getsize(OUT_FILE)} байт")
