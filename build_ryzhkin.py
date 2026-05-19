# -*- coding: utf-8 -*-
"""
Сборка БП Рыжкина А.А. (двери, Рязань, 350к грант).
Запуск:
  python3 build_ryzhkin.py
Создаёт файл: "Бизнес план Рыжкин.docx" в корне репо.
"""
import os, shutil, sys, json, zipfile
from xml.sax.saxutils import escape

ROOT = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = "/tmp/docx_extract/Бизнес план Куксов испр"
OUT_FILE = os.path.join(ROOT, "Бизнес план Рыжкин.docx")
WORK_DIR = "/tmp/ryzhkin_build"

# ============================================================
# 1. ВХОДНЫЕ ДАННЫЕ
# ============================================================
PERSON = {
    "fio_full": "Рыжкин Артемий Алексеевич",
    "fio_short": "А.А. Рыжкин",
    "birth": "27.12.2004 г.",
    "addr_reg": "390011, г. Рязань, ул. Старообрядческий проезд, д. 11, кв. 33",
    "addr_fact": "г. Рязань, ул. Московская, д. 35, кв. 84",
    "loc_short": "г. Рязань, ул. Московская, д. 35",
    "phone": "+79308891018",
    "email": "artemryzkin80@gmail.com",
    "inn": "622402008161",
    "okved": "43.32 Работы столярные и плотничные",
}

# ============================================================
# 2. СМЕТА (15 позиций оборудования + 1 продвижение)
# ============================================================
# Позиции 1.1—1.13 идут на грант полностью (347 951).
# Позиция 1.14 (Кейсы Tanos) — переходная: грант 2 049, свои 5 603.
# Позиции 1.15 (расходники) и 4.* (продвижение) — полностью свои.
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
    ("Комплект аккумуляторов и ЗУ (Ryobi 18 В 2.5 Ач, Makita 12 В 1.5 Ач, DeWalt 18 В 5.0 Ач)", 1, 16081, "Все инструменты", "grant"),
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

# Расчёт распределения грант/свои на переходной позиции 1.14
sum_full_grant = sum(c * p for _, c, p, _, src in EQUIPMENT[:13])  # позиции 1.1—1.13
mixed_pos_total = EQUIPMENT[13][1] * EQUIPMENT[13][2]
mixed_grant = GRANT - sum_full_grant
mixed_own = mixed_pos_total - mixed_grant
own_pos_15 = EQUIPMENT[14][1] * EQUIPMENT[14][2]

# ============================================================
# 3. ФИНМОДЕЛЬ — таблица 5.1 (помесячно по услугам)
# ============================================================
# Сезонность дверной ниши: пик март-октябрь, спад ноябрь-февраль,
# м1 = 0 (закупка инструмента и реклама).
SERVICES = [
    # (название_длинное, название_короткое_для_5.1, время_текст, цена, [12 значений])
    ("Установка межкомнатной двери под ключ (МДФ/экошпон): врезка петель и магнитного замка, монтаж наличников и доборов",
     "Установка межкомнатной двери под ключ", "3 ч. 30 мин.", 3500,
     [0, 6, 8, 10, 11, 10, 9, 10, 12, 14, 9, 8]),
    ("Монтаж скрытой двери Invisible: закладка алюминиевого короба + чистовая навеска полотна (двухэтапный)",
     "Монтаж скрытой двери Invisible", "6 ч.", 6500,
     [0, 1, 2, 2, 2, 2, 1, 2, 3, 3, 1, 2]),
    ("Установка входной металлической двери: анкерный крепёж, регулировка уплотнителей и доводчика",
     "Установка входной металлической двери", "3 ч.", 3500,
     [0, 2, 2, 2, 3, 2, 2, 2, 3, 3, 2, 2]),
    ("Установка комплекта доборных элементов и расширение/сужение проёма",
     "Установка доборов и расширение проёма", "1 ч. 30 мин.", 1500,
     [0, 5, 7, 9, 10, 9, 8, 9, 11, 12, 8, 7]),
    ("Монтаж скрытого (теневого) алюминиевого плинтуса в квартире до 30 пог. м",
     "Монтаж скрытого (теневого) плинтуса до 30 пог. м", "5 ч.", 7500,
     [0, 1, 2, 2, 2, 2, 1, 2, 2, 3, 2, 2]),
]

def revenue_by_month():
    """Возвращает [12] помесячной выручки и общую."""
    months = [0] * 12
    for _, _, _, price, qty in SERVICES:
        for i, q in enumerate(qty):
            months[i] += q * price
    return months, sum(months)

REV_MONTHS, REV_TOTAL = revenue_by_month()

# ============================================================
# 4. ФИНМОДЕЛЬ — таблица 5.2 (расходы / прибыль / ЧДП)
# ============================================================
# Месяц 1: 2.1 = вся смета (371 603), остальное = 0
# Месяцы 2-12:
#   2.2 расходные ~9% выручки (рваные)
#   2.4 транспорт без авто (такси/каршеринг) 1500-2500
#   2.5 связь/интернет 800
#   2.6 реклама 1800-3000 (рваные, постоянные расходы на Авито PRO)
#   2.7 непредвиденные/амортизация 300-500
#   2.8 НПД 4% от выручки
EXP_MATERIALS = [0, 4500, 6500, 7400, 8000, 7400, 5800, 7300, 9000, 10500, 6300, 6500]  # 79 200
EXP_TRANSPORT = [0, 1500, 1700, 1800, 2000, 1800, 1500, 1800, 2200, 2500, 1500, 1600]   # 19 900
EXP_COMMS     = [0, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800]               # 8 800
EXP_PROMO     = [0, 1800, 2200, 2200, 2500, 2500, 2000, 2200, 2800, 3000, 2200, 2000]   # 25 400
EXP_UNFORESEEN= [0, 300, 300, 400, 400, 400, 300, 400, 500, 500, 300, 300]               # 4 100
EXP_TAX       = [round(r * 0.04) for r in REV_MONTHS]                                    # 4% от выручки

# Оборудование в м1 = вся смета (включая стартовое продвижение из раздела 4*)
EXP_EQUIPMENT = [SMETA_TOTAL] + [0] * 11
EXP_RENT      = [0] * 12

EXPENSES_BY_MONTH = []
for i in range(12):
    EXPENSES_BY_MONTH.append(
        EXP_EQUIPMENT[i] + 0 + EXP_RENT[i] + EXP_TRANSPORT[i] + EXP_COMMS[i] +
        EXP_PROMO[i] + EXP_UNFORESEEN[i] + EXP_TAX[i] + EXP_MATERIALS[i]
    )
PROFIT_BY_MONTH = [REV_MONTHS[i] - EXPENSES_BY_MONTH[i] for i in range(12)]

# Стороннее финансирование (грант 350к в м1) и собственные (21 603 в м1)
EXTERNAL = [GRANT] + [0] * 11
OWNFUND  = [OWN] + [0] * 11

CFLOW = [PROFIT_BY_MONTH[i] + EXTERNAL[i] + OWNFUND[i] for i in range(12)]
CFLOW_CUM = []
acc = 0
for v in CFLOW:
    acc += v
    CFLOW_CUM.append(acc)

REV_TOTAL_CHK = sum(REV_MONTHS)
EXP_TOTAL = sum(EXPENSES_BY_MONTH)
PROFIT_TOTAL = REV_TOTAL_CHK - EXP_TOTAL
RENTABILITY = PROFIT_TOTAL / EXP_TOTAL * 100

# Срок окупаемости: первый месяц, когда CFLOW_CUM >= SMETA_TOTAL
PAYBACK = 12
for i, v in enumerate(CFLOW_CUM):
    if v >= SMETA_TOTAL:
        PAYBACK = i + 1
        break

# 5.4: 2-й год (без оборудования, рост выручки и расходов)
YEAR2_REVENUE = 1020000
YEAR2_EXPENSES = 290000
YEAR2_PROFIT = YEAR2_REVENUE - YEAR2_EXPENSES
YEAR2_RENT = YEAR2_PROFIT / YEAR2_EXPENSES * 100

# ============================================================
# 5. АУДИТ (4 проверки сметы + 12 перекрёстных)
# ============================================================
def audit():
    errors = []
    # 1. Σ всех позиций 4.2 == ИТОГО
    s_check = sum(c * p for _, c, p, _, _ in EQUIPMENT) + PROMO_ITEM[1] * PROMO_ITEM[2]
    if s_check != SMETA_TOTAL:
        errors.append(f"AUDIT-1: сумма позиций {s_check} != ИТОГО {SMETA_TOTAL}")
    # 2. Грант + Свои == ИТОГО, грант ровно 350 000
    if GRANT + OWN != SMETA_TOTAL:
        errors.append(f"AUDIT-2: grant({GRANT}) + own({OWN}) != total({SMETA_TOTAL})")
    if GRANT != 350000:
        errors.append(f"AUDIT-2b: грант {GRANT} != 350 000")
    # 3. Переходная позиция корректна
    if mixed_grant <= 0 or mixed_own <= 0:
        errors.append(f"AUDIT-3: переходная позиция распределена некорректно: grant={mixed_grant}, own={mixed_own}")
    if mixed_grant + mixed_own != mixed_pos_total:
        errors.append(f"AUDIT-3b: переходная позиция: {mixed_grant}+{mixed_own} != {mixed_pos_total}")
    # 4. Все позиции выше переходной = grant, ниже = own
    for i, (_, _, _, _, src) in enumerate(EQUIPMENT[:13]):
        if src != "grant":
            errors.append(f"AUDIT-4: поз 1.{i+1} должна быть grant, имеет {src}")
    if EQUIPMENT[14][4] != "own":
        errors.append(f"AUDIT-4b: поз 1.15 должна быть own")

    # Перекрёстные проверки
    # X1. ИТОГО смета в коридоре 368 000 — 372 000
    if not (368000 <= SMETA_TOTAL <= 372000):
        errors.append(f"X1: ИТОГО {SMETA_TOTAL} вне коридора 368-372к")
    # X2. Выручка год 1 в коридоре 850-1150к
    if not (850000 <= REV_TOTAL <= 1150000):
        errors.append(f"X2: выручка {REV_TOTAL} вне коридора 850-1150к")
    # X3. 5.1 ИТОГО == 5.2 строка 1 ИТОГО
    if REV_TOTAL != sum(REV_MONTHS):
        errors.append(f"X3: 5.1 vs 5.2 рассинхрон")
    # X4. 5.2 строка 2.1 ИТОГО == ИТОГО смета 4.2 == 1.16 закупка
    if EXP_EQUIPMENT[0] != SMETA_TOTAL:
        errors.append(f"X4: 5.2/2.1 ({EXP_EQUIPMENT[0]}) != смета 4.2 ({SMETA_TOTAL})")
    # X5. 5.2 строка 6 = 350 000, строка 7 = OWN
    if sum(EXTERNAL) != GRANT:
        errors.append(f"X5: стороннее {sum(EXTERNAL)} != 350 000")
    if sum(OWNFUND) != OWN:
        errors.append(f"X5b: свои {sum(OWNFUND)} != {OWN}")
    # X6. ЧДП м1 == 0
    if CFLOW[0] != 0:
        errors.append(f"X6: ЧДП м1 = {CFLOW[0]} (должен быть 0)")
    # X7. Прибыль ИТОГО == Σ прибыль помесячно
    if PROFIT_TOTAL != sum(PROFIT_BY_MONTH):
        errors.append(f"X7: прибыль ИТОГО рассинхрон")
    # X8. Рентабельность 40-70%
    if not (40 <= RENTABILITY <= 70):
        errors.append(f"X8: рентабельность {RENTABILITY:.1f}% вне коридора 40-70%")
    # X9. Помесячные часы ≤ 160
    for m in range(12):
        hrs = (qty := SERVICES[0][4][m]) * 3.5 + SERVICES[1][4][m] * 6 + SERVICES[2][4][m] * 3 + SERVICES[3][4][m] * 1.5 + SERVICES[4][4][m] * 5
        if hrs > 160:
            errors.append(f"X9: мес {m+1} часов {hrs} > 160")
    # X10. Срок окупаемости 5-9 мес
    if not (5 <= PAYBACK <= 9):
        errors.append(f"X10: срок окупаемости {PAYBACK} вне диапазона 5-9 мес")
    # X11. Расходы по статьям складываются
    for m in range(12):
        s = EXP_EQUIPMENT[m] + EXP_MATERIALS[m] + EXP_RENT[m] + EXP_TRANSPORT[m] + EXP_COMMS[m] + EXP_PROMO[m] + EXP_UNFORESEEN[m] + EXP_TAX[m]
        if s != EXPENSES_BY_MONTH[m]:
            errors.append(f"X11: мес {m+1} расходы {s} != {EXPENSES_BY_MONTH[m]}")
    # X12. Нараст. ЧДП в м12 == Σ ЧДП == прибыль_итого + грант + свои
    if CFLOW_CUM[-1] != sum(CFLOW):
        errors.append(f"X12: нараст м12 != Σ ЧДП")
    if CFLOW_CUM[-1] != PROFIT_TOTAL + GRANT + OWN:
        errors.append(f"X12b: нараст м12 != прибыль + грант + свои")
    return errors

ERRORS = audit()
print("=" * 60)
print(f"СМЕТА: {SMETA_TOTAL:,} ₽ (грант {GRANT:,} + свои {OWN:,})".replace(",", " "))
print(f"Оборудование: {EQUIP_TOTAL:,} ₽, Продвижение: {PROMO_TOTAL:,} ₽".replace(",", " "))
print(f"Переходная позиция 1.14 (Кейсы Tanos): грант {mixed_grant} + свои {mixed_own} = {mixed_pos_total}")
print(f"ВЫРУЧКА год 1: {REV_TOTAL:,} ₽".replace(",", " "))
print(f"РАСХОДЫ год 1: {EXP_TOTAL:,} ₽".replace(",", " "))
print(f"ПРИБЫЛЬ год 1: {PROFIT_TOTAL:,} ₽".replace(",", " "))
print(f"ЧДП нарастающим м12: {CFLOW_CUM[-1]:,} ₽".replace(",", " "))
print(f"Рентабельность: {RENTABILITY:.1f}%")
print(f"Срок окупаемости: {PAYBACK} мес")
print("=" * 60)
if ERRORS:
    print("ОШИБКИ АУДИТА:")
    for e in ERRORS:
        print(" -", e)
    sys.exit(1)
else:
    print("АУДИТ: OK (4+12 проверок пройдено)")
print("=" * 60)


# ============================================================
# 6. ГЕНЕРАЦИЯ document.xml
# ============================================================

def fmt_num(n):
    """123456 -> '123 456'"""
    return f"{n:,}".replace(",", " ")

def x(s):
    return escape(str(s), {'"': '&quot;', "'": '&apos;'})

# ----------------------------------------------------------
# Helper builders
# ----------------------------------------------------------
def run(text, *, bold=False, size=28, italic=False):
    """Run with TNR font, размер в half-points (28 = 14pt, 22 = 11pt)."""
    rpr = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/>'
    if bold:
        rpr += '<w:b/><w:bCs/>'
    if italic:
        rpr += '<w:i/><w:iCs/>'
    rpr += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>'
    # preserve inner whitespace
    return f'<w:r>{rpr}<w:t xml:space="preserve">{x(text)}</w:t></w:r>'

def para(text="", *, bold=False, size=28, align="both", italic=False, ind_first=0, before=0, after=0):
    ppr = '<w:pPr>'
    spacing = ''
    if before:
        spacing += f' w:before="{before}"'
    if after:
        spacing += f' w:after="{after}"'
    if spacing:
        ppr += f'<w:spacing{spacing} w:line="276" w:lineRule="auto"/>'
    if ind_first:
        ppr += f'<w:ind w:firstLine="{ind_first}"/>'
    if align != "both":
        ppr += f'<w:jc w:val="{align}"/>'
    else:
        ppr += '<w:jc w:val="both"/>'
    ppr += f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="{size}"/></w:rPr>'
    ppr += '</w:pPr>'
    if text == "":
        return f'<w:p>{ppr}</w:p>'
    return f'<w:p>{ppr}{run(text, bold=bold, size=size, italic=italic)}</w:p>'

def para_runs(runs_xml, *, align="both", before=0, after=0, ind_first=0, size=28):
    ppr = '<w:pPr>'
    spacing = ''
    if before:
        spacing += f' w:before="{before}"'
    if after:
        spacing += f' w:after="{after}"'
    if spacing:
        ppr += f'<w:spacing{spacing} w:line="276" w:lineRule="auto"/>'
    if ind_first:
        ppr += f'<w:ind w:firstLine="{ind_first}"/>'
    if align != "both":
        ppr += f'<w:jc w:val="{align}"/>'
    else:
        ppr += '<w:jc w:val="both"/>'
    ppr += '</w:pPr>'
    return f'<w:p>{ppr}{runs_xml}</w:p>'

def cell(content_xml, *, w=2000, valign="center"):
    """content_xml — XML параграфов для ячейки."""
    return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/><w:vAlign w:val="{valign}"/></w:tcPr>{content_xml}</w:tc>'

def cell_text(text, *, w=2000, bold=False, align="center", size=22, valign="center"):
    """Удобная ячейка с одним параграфом."""
    p = para(text, bold=bold, size=size, align=align)
    return cell(p, w=w, valign=valign)

def table(rows_xml, widths, *, header_borders=True):
    """rows_xml — список <w:tr>...</w:tr>; widths — список ширин."""
    total_w = sum(widths)
    grid = ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    borders = '''<w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tblBorders>'''
    return f'''<w:tbl>
        <w:tblPr>
            <w:tblW w:w="{total_w}" w:type="dxa"/>
            <w:jc w:val="center"/>
            {borders}
            <w:tblLayout w:type="fixed"/>
            <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
        </w:tblPr>
        <w:tblGrid>{grid}</w:tblGrid>
        {''.join(rows_xml)}
    </w:tbl>'''

def row(cells_xml):
    return f'<w:tr>{cells_xml}</w:tr>'

def row_merged(cells_xml):
    return f'<w:tr>{cells_xml}</w:tr>'

# ============================================================
# Сборка тела документа
# ============================================================
body_parts = []
B = body_parts.append

# ---------- ТИТУЛ ----------
B(para("Бизнес-план", bold=True, size=32, align="center", before=240, after=120))
B(para("по реализации мероприятий социальной адаптации", bold=True, size=28, align="center", after=60))
B(para("по осуществлению индивидуальной предпринимательской деятельности", bold=True, size=28, align="center", after=240))
B(para("Проект «Профессиональная установка межкомнатных и входных дверей, монтаж скрытых дверных систем Invisible и теневых алюминиевых плинтусов»", bold=True, italic=True, size=28, align="center", after=480))
B(para("Разработал:", size=28, align="center", after=60))
B(para(PERSON["fio_full"], bold=True, size=28, align="center", after=480))
B(para(f"Адрес: {PERSON['addr_reg']}", size=28, align="center", after=60))
B(para(f"Телефон: {PERSON['phone']}", size=28, align="center", after=60))
B(para(f"E-mail: {PERSON['email']}", size=28, align="center", after=480))
B(para("г. Рязань", size=28, align="center", after=60))
B(para("2026", size=28, align="center", after=240))
# Разрыв страницы
B('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

# ---------- 1. РЕЗЮМЕ ----------
B(para("1. Резюме", bold=True, size=32, align="left", before=120, after=120))
B(para(f"1.1. {PERSON['fio_full']}.", size=28, align="both"))
B(para(f"1.2. Дата рождения: {PERSON['birth']}", size=28))
B(para(f"1.3. Адрес регистрации: {PERSON['addr_reg']}.", size=28))
B(para(f"1.4. Адрес фактического проживания: {PERSON['addr_fact']}.", size=28))
B(para(f"1.5. Контактные данные: Телефон: {PERSON['phone']}, E-mail: {PERSON['email']}.", size=28))
B(para("1.6. Семейное положение: не женат.", size=28))
B(para("1.7. Образование: получаю высшее техническое в ФГБОУ ВО «Рязанский государственный радиотехнический университет имени В.Ф. Уткина» (РГРТУ) по направлению 15.03.04 «Автоматизация технологических процессов и производств». В настоящее время обучаюсь на 4 курсе очной формы. (Приложение 1)", size=28))
B(para("1.8. Общий стаж работы, трудовая деятельность в течение последних 3 лет: в связи с прохождением обучения по очной форме в высшем учебном заведении официальная трудовая деятельность в период с 2023 по 2026 гг. не осуществлялась.", size=28))
B(para("1.9. Опыт работы (предпринимательской деятельности) по выбранному направлению: владею практическими навыками работы со столярным и монтажным электроинструментом, разметки проёмов и устройства напольных покрытий, наработанными при участии в отделочных и ремонтных работах в качестве помощника установщика. Профильное техническое образование (РГРТУ, направление 15.03.04 «Автоматизация технологических процессов и производств») обеспечивает уверенное чтение чертёжной документации, лазерную нивелировку и точную разметку проёмов под скрытые дверные системы и теневые алюминиевые плинтуса.", size=28))
B(para("1.10. Потребность в обучении/повышении квалификации с обоснованием: потребность в обучении отсутствует. Уровень практической подготовки и базового инженерного образования в полной мере соответствует требованиям технологии монтажа межкомнатных и входных дверных блоков, скрытых дверных систем Invisible и алюминиевых теневых плинтусов.", size=28))
B(para("1.11. Текущий статус автора проекта: обучающийся, применяющий специальный налоговый режим «Налог на профессиональный доход» (самозанятый).", size=28))
B(para(f"1.12. Организационно-правовая форма ведения организуемого бизнеса и система налогообложения: самозанятость, система налогообложения — Налог на профессиональный доход (НПД). ИНН {PERSON['inn']}.", size=28))
B(para(f"1.13. Виды деятельности (ОКВЭД): {PERSON['okved']}.", size=28))
B(para("1.14. Наличие или необходимость лицензий на соответствующие виды деятельности, патентов, сертификатов, авторских прав, медицинской книжки и т.п. - не требуется.", size=28))
B(para(f"1.15. Месторасположение: {PERSON['loc_short']}. Имеется собственный гараж рядом с домом для хранения инструмента и оснастки.", size=28))
B(para("1.16. Показатели экономической эффективности", size=28))
B(para("По итогам первого года реализации проекта:", size=28))

# Таблица 1.16
rows_116 = []
rows_116.append(row(
    cell_text("Показатель", w=4500, bold=True, size=22) +
    cell_text("Значение", w=3000, bold=True, size=22)
))
rows_116.append(row(cell_text("доходы, руб.", w=4500, align="left", size=22) + cell_text(fmt_num(REV_TOTAL), w=3000, size=22)))
rows_116.append(row(cell_text("расходы, руб.", w=4500, align="left", size=22) + cell_text(fmt_num(EXP_TOTAL), w=3000, size=22)))
rows_116.append(row(cell_text("прибыль, руб.", w=4500, align="left", size=22) + cell_text(fmt_num(PROFIT_TOTAL), w=3000, size=22)))
rows_116.append(row(cell_text("чистый денежный поток, руб.", w=4500, align="left", size=22) + cell_text(fmt_num(CFLOW_CUM[-1]), w=3000, size=22)))
rows_116.append(row(cell_text("рентабельность, %", w=4500, align="left", size=22) + cell_text(f"{RENTABILITY:.1f}%".replace(".", ","), w=3000, size=22)))
rows_116.append(row(cell_text("срок окупаемости, мес.", w=4500, align="left", size=22) + cell_text(f"{PAYBACK} месяцев", w=3000, size=22)))
B(table(rows_116, [4500, 3000]))
B(para(""))

# ---------- 2. ОПИСАНИЕ БИЗНЕСА ----------
B(para("2. Описание бизнеса и план маркетинга", bold=True, size=32, align="left", before=120, after=120))

B(para_runs(
    run("2.1 Краткое описание бизнеса по проекту: ", bold=True, size=28) +
    run("запуск выездной мастерской по профессиональной установке межкомнатных и входных дверных блоков, монтажу скрытых дверных систем Invisible и устройству скрытого (теневого) алюминиевого плинтуса в квартирах, частных домах и коммерческих помещениях г. Рязани. Технологический процесс выстроен в три самостоятельных этапа: подготовка проёма с лазерной нивелировкой и контролем геометрии, прецизионная врезка петель и магнитных замков по шаблону «Платинум» (минимальный люфт и идентичная посадка фурнитуры на каждом полотне), чистовой монтаж полотен, доборных элементов и наличников с финальной отделкой стыков. Для скрытых дверных систем работа разбита на закладку алюминиевого короба на этапе чернового ремонта и чистовую навеску полотна после отделки стен — это критическая компетенция, отсутствующая у большинства частных мастеров г. Рязани. Деятельность не требует аренды помещения: все работы выполняются на объекте заказчика, инструмент и расходные материалы хранятся в собственном гараже. Бизнес имеет круглогодичный характер с устойчивыми сезонными пиками в марте–октябре, обусловленными активным ремонтным циклом и сдачей новостроек.", size=28)
))

B(para("Таблица 2.1.1 План продаж товаров, работ, услуг в месяц", size=28, before=120, after=60))

# Таблица 2.1.1
rows_211 = []
rows_211.append(row(
    cell_text("Виды работ/услуг", w=6500, bold=True, size=22) +
    cell_text("Затрачиваемое время", w=2200, bold=True, size=22) +
    cell_text("Стоимость (руб.)", w=1800, bold=True, size=22)
))
for full_name, _, time_text, price, _ in SERVICES:
    rows_211.append(row(
        cell_text(full_name, w=6500, align="left", size=22) +
        cell_text(time_text, w=2200, size=22) +
        cell_text(fmt_num(price), w=1800, size=22)
    ))
B(table(rows_211, [6500, 2200, 1800]))

B(para("В состав каждой услуги входит выезд мастера на объект, защита прилегающих поверхностей плёнкой, разметка лазерным нивелиром, контроль геометрии проёма уровнем Stabila, врезка фурнитуры по шаблону «Платинум», финальная очистка пылеудалителем Makita и сдача результата заказчику с фотоотчётом «до/после». Цены сформированы исходя из медианных рыночных показателей в г. Рязани по данным анализа объявлений Авито (Приложение 3).", size=28, before=60))

B(para_runs(
    run("2.2 Целевая аудитория: ", bold=True, size=28) +
    run("основной целевой аудиторией являются собственники квартир в новостройках и вторичном жилье г. Рязани и пригородных микрорайонов в радиусе до 30 км, заходящие в этап чистового ремонта и заказывающие комплект из 4–6 межкомнатных дверей одновременно. Второй сегмент — владельцы квартир в премиальных жилых комплексах с панорамным остеклением и европланировкой, выбирающие скрытые дверные системы Invisible и теневые алюминиевые плинтуса как часть единого дизайн-кода. Третий сегмент — частные дизайнеры интерьеров и небольшие ремонтные бригады г. Рязани, нуждающиеся в подрядчике на узкий технологический участок «двери + плинтус» с фиксированным сроком и собственной прецизионной оснасткой.", size=28)
))

B(para_runs(
    run("2.3. Конкурентная среда, конкурентная позиция бизнеса: ", size=28) +
    run("проанализировав объявления сайта «Авито» по г. Рязани (Приложение 3), я выделил три уровня предложения. Первый — бригады при магазинах и салонах межкомнатных дверей, привязанные к продаже своей продукции; стоимость монтажа от 3 500 до 5 500 рублей за стандартное полотно, при этом скрытые системы и теневые плинтуса в их перечне отсутствуют. Второй — частные мастера с базовым набором инструмента (перфоратор, шуруповёрт, ручная стусло-пила); стоимость от 2 800 до 3 500 рублей за дверь, однако без шаблона врезки фурнитуры и лазерной нивелировки точность монтажа полотен с магнитными замками и скрытыми петлями нестабильна. Третий уровень — профессиональные установщики со специализированной оснасткой (фрезер, торцовочная пила, шаблон «Платинум», лазерный нивелир); их предложение составляет 4 500–6 500 рублей за стандартную дверь и 5 500–8 000 за скрытую систему. Спрос на услуги стабилен и растёт за счёт массового ввода новостроек в Рязани и популярности скрытых дверных систем в проектах европланировки.", size=28)
))
B(para("Конкурентными преимуществами проекта являются: полный комплект прецизионной оснастки (фрезер Makita RT0702CX2, торцовочная пила Makita LS0714N, шаблон «Платинум Павла Солдатова», лазерный нивелир ADA Cube 360); владение технологией монтажа скрытых дверных систем Invisible в два этапа; сопутствующая компетенция по теневым алюминиевым плинтусам, востребованным в премиальном сегменте; гарантия на монтаж 12 месяцев и фотоотчёт по каждому объекту.", size=28))

B(para_runs(
    run("2.4. Каналы реализации товара: ", bold=True, size=28) +
    run("основным каналом привлечения заявок будет являться платное продвижение профильных объявлений на платформах «Авито» и «Юла» в категории «Двери, окна» по г. Рязани и Рязанской области. Параллельно ведётся профиль на сервисе «Профи.ру» с накоплением отзывов и портфолио, а также страница в социальных сетях с фотоотчётами «до/после» и короткими роликами этапов монтажа. Офлайн-канал — личные договорённости с менеджерами рязанских салонов межкомнатных дверей и подрядных ремонтных бригад на постоянное направление клиентов в обмен на партнёрский процент.", size=28)
))

B(para_runs(
    run("2.8. Продвижение. ", bold=True, size=28) +
    run("Поиск клиентов будет вестись через прямую коллаборацию с дизайнерами интерьеров и ремонтными бригадами г. Рязани, работающими в сегменте премиальной отделки и европланировки. Уникальное торговое предложение проекта — пакет «Скрытая дверь Invisible под ключ»: закладка короба на черновом этапе и чистовая навеска полотна после отделки в фиксированный срок без переделок и доборов за счёт мастера. Для повышения конверсии заявок будет использоваться платное продвижение объявлений и таргетинг по запросам «установка двери Рязань», «скрытая дверь монтаж», «теневой плинтус». Каждый завершённый объект сопровождается фотоотчётом и короткой видеовизитой, формирующими портфолио для дальнейшего органического трафика.", size=28)
))

# ---------- 3. ОРГАНИЗАЦИОННЫЙ ПЛАН ----------
B(para("3. Организационный план", bold=True, size=32, align="left", before=120, after=120))
B(para("3.1. Планируемый график работы - с 09-00 до 21-00 без выходных. Выезд на объекты осуществляется по предварительной договорённости с заказчиком.", size=28))
B(para("3.2. Штатное расписание дополнительные сотрудники не требуются.", size=28))
B(para("3.3. Требования к персоналу: иметь практические навыки прецизионной разметки проёмов, врезки петель и замков по шаблону, монтажа скрытых дверных систем и устройства алюминиевых плинтусов, которые у меня имеются.", size=28))

# ---------- 4. ПРОИЗВОДСТВЕННЫЙ ПЛАН ----------
B(para("4. Производственный план", bold=True, size=32, align="left", before=120, after=120))
B(para("4.1. Источники финансирования", size=28))

# Таблица 4.1
rows_41 = []
rows_41.append(row(
    cell_text("№ п/п", w=900, bold=True, size=22) +
    cell_text("Источник финансирования", w=6800, bold=True, size=22) +
    cell_text("Сумма (тыс. руб.)", w=2300, bold=True, size=22)
))
rows_41.append(row(
    cell_text("1.", w=900, size=22) +
    cell_text("Средства государственной социальной помощи в рамках заключенного социального контракта на осуществление индивидуальной предпринимательской деятельности*", w=6800, align="left", size=22) +
    cell_text(fmt_num(GRANT), w=2300, size=22)
))
rows_41.append(row(
    cell_text("2.", w=900, size=22) +
    cell_text("Собственные средства", w=6800, align="left", size=22) +
    cell_text(fmt_num(OWN), w=2300, size=22)
))
rows_41.append(row(
    cell_text("3.", w=900, size=22) +
    cell_text("Заемные средства", w=6800, align="left", size=22) +
    cell_text("0", w=2300, size=22)
))
rows_41.append(row(
    cell_text("4.", w=900, size=22) +
    cell_text("Иное", w=6800, align="left", size=22) +
    cell_text("0", w=2300, size=22)
))
rows_41.append(row(
    cell_text("ИТОГО", w=900, bold=True, size=22) +
    cell_text("", w=6800, size=22) +
    cell_text(fmt_num(SMETA_TOTAL), w=2300, bold=True, size=22)
))
B(table(rows_41, [900, 6800, 2300]))
B(para(""))

B(para("4.2. Перечень необходимых для бизнеса затрат с указанием источников финансирования (в необходимых случаях – обоснование закупок).", size=28, before=120))

# Таблица 4.2
rows_42 = []
W42 = [700, 4800, 700, 1300, 1300, 1500, 2000]
rows_42.append(row(
    cell_text("№ п/п", w=W42[0], bold=True, size=22) +
    cell_text("Наименование", w=W42[1], bold=True, size=22) +
    cell_text("Кол-во", w=W42[2], bold=True, size=22) +
    cell_text("Цена за ед.", w=W42[3], bold=True, size=22) +
    cell_text("Сумма", w=W42[4], bold=True, size=22) +
    cell_text("Поставщик", w=W42[5], bold=True, size=22) +
    cell_text("Источник финансирования", w=W42[6], bold=True, size=22)
))
# Раздел 1.*
rows_42.append(row(
    cell_text("1.*", w=W42[0], bold=True, size=22) +
    cell_text("Основные средства (оборудование, инструменты, оснастка):", w=W42[1], bold=True, align="left", size=22) +
    cell_text("", w=W42[2], size=22) +
    cell_text("", w=W42[3], size=22) +
    cell_text(fmt_num(EQUIP_TOTAL), w=W42[4], bold=True, size=22) +
    cell_text("", w=W42[5], size=22) +
    cell_text("", w=W42[6], size=22)
))
def src_label(src):
    return {
        "grant": "Средства соцконтракта",
        "own": "Собственные средства",
        "mixed": "Средства соцконтракта, собственные средства",
    }[src]
for i, (name, qty, price, supplier, src) in enumerate(EQUIPMENT, start=1):
    summa = qty * price
    rows_42.append(row(
        cell_text(f"1.{i}", w=W42[0], size=22) +
        cell_text(name, w=W42[1], align="left", size=22) +
        cell_text(str(qty), w=W42[2], size=22) +
        cell_text(fmt_num(price), w=W42[3], size=22) +
        cell_text(fmt_num(summa), w=W42[4], size=22) +
        cell_text(supplier, w=W42[5], size=22) +
        cell_text(src_label(src), w=W42[6], align="left", size=22)
    ))
# Раздел 2.* — 0
rows_42.append(row(
    cell_text("2.*", w=W42[0], bold=True, size=22) +
    cell_text("Материально-производственные запасы (комплектующие, сырье и т.д.):", w=W42[1], bold=True, align="left", size=22) +
    cell_text("", w=W42[2], size=22) +
    cell_text("", w=W42[3], size=22) +
    cell_text("0", w=W42[4], bold=True, size=22) +
    cell_text("", w=W42[5], size=22) +
    cell_text("", w=W42[6], size=22)
))
# Раздел 3.* — 0
rows_42.append(row(
    cell_text("3.*", w=W42[0], bold=True, size=22) +
    cell_text("Имущественные обязательства (аренда, не более 15% общей суммы)", w=W42[1], bold=True, align="left", size=22) +
    cell_text("", w=W42[2], size=22) +
    cell_text("", w=W42[3], size=22) +
    cell_text("0", w=W42[4], bold=True, size=22) +
    cell_text("", w=W42[5], size=22) +
    cell_text("", w=W42[6], size=22)
))
# Раздел 4.* — продвижение
rows_42.append(row(
    cell_text("4.*", w=W42[0], bold=True, size=22) +
    cell_text("Продвижение (не более 5% общей суммы)", w=W42[1], bold=True, align="left", size=22) +
    cell_text(str(PROMO_ITEM[1]), w=W42[2], size=22) +
    cell_text(fmt_num(PROMO_ITEM[2]), w=W42[3], size=22) +
    cell_text(fmt_num(PROMO_TOTAL), w=W42[4], bold=True, size=22) +
    cell_text(PROMO_ITEM[3], w=W42[5], size=22) +
    cell_text("Собственные средства", w=W42[6], align="left", size=22)
))
# Раздел 5.* — 0
rows_42.append(row(
    cell_text("5.*", w=W42[0], bold=True, size=22) +
    cell_text("Подготовка разрешительной документации, приобретение ПО, электронной подписи (не более 10% общей суммы)", w=W42[1], bold=True, align="left", size=22) +
    cell_text("", w=W42[2], size=22) +
    cell_text("", w=W42[3], size=22) +
    cell_text("0", w=W42[4], bold=True, size=22) +
    cell_text("", w=W42[5], size=22) +
    cell_text("", w=W42[6], size=22)
))
# ИТОГО
rows_42.append(row(
    cell_text("ИТОГО", w=W42[0], bold=True, size=22) +
    cell_text("", w=W42[1], size=22) +
    cell_text("", w=W42[2], size=22) +
    cell_text("", w=W42[3], size=22) +
    cell_text(fmt_num(SMETA_TOTAL), w=W42[4], bold=True, size=22) +
    cell_text("", w=W42[5], size=22) +
    cell_text("", w=W42[6], size=22)
))
B(table(rows_42, W42))
B(para("(Приложение 2)", size=28, align="right"))

B(para("4.3. Помесячный план закупок с указанием наименования затрат и необходимой суммой.", size=28, before=120))
B(para("В течение одного месяца с момента поступления денежных средств в рамках социального контракта проводится закупка полного спектра необходимого инструмента и оснастки. Порядок закупки: заказ и оплата основного электроинструмента (фрезер, торцовочная пила, перфоратор, шуруповёрт, лобзик, угловая шлифмашина); приобретение прецизионной оснастки (лазерный нивелир, шаблон врезки «Платинум», уровень, стусло, промышленный пылесос); закупка комплектов аккумуляторов, пильных дисков, систейнеров и стартового запаса расходных материалов; запуск рекламной кампании на платформе «Авито».", size=28))

# ---------- 5. ФИНАНСОВЫЙ ПЛАН ----------
B(para("5. Финансовый план и показатели эффективности проекта", bold=True, size=32, align="left", before=120, after=120))
B(para("5.1. Прогноз предполагаемых доходов по проекту", size=28))

# Таблица 5.1 — сложная: 12 столбцов мес + ИТОГО + Доля
# Столбцы: № п/п | Наименование | м1..м12 | ИТОГО | Доля
W51 = [400, 2400] + [550] * 12 + [900, 700]
HEAD51 = ["№ п/п", "Наименование месяца / товара/услуги"] + [f"{i+1} мес." for i in range(12)] + ["ИТОГО", "Доля, %"]

rows_51 = []
# Шапка
rows_51.append(row(
    ''.join(cell_text(t, w=W51[i], bold=True, size=20) for i, t in enumerate(HEAD51))
))
# Строка "Доходы всего, руб."
months_total = REV_MONTHS
total_rev = sum(months_total)
rows_51.append(row(
    cell_text("", w=W51[0], size=20) +
    cell_text("Доходы всего, руб., в т.ч.:", w=W51[1], bold=True, align="left", size=20) +
    ''.join(cell_text(fmt_num(m) if m else "0", w=W51[2 + i], bold=True, size=20) for i, m in enumerate(months_total)) +
    cell_text(fmt_num(total_rev), w=W51[14], bold=True, size=20) +
    cell_text("100%", w=W51[15], bold=True, size=20)
))
# Строки услуг
for idx, (full_name, short_name, time_text, price, qty) in enumerate(SERVICES, start=1):
    revenue = sum(qty) * price
    share = revenue / total_rev * 100
    # Строка с ценой за заказ и доходом по месяцам
    rows_51.append(row(
        cell_text(str(idx), w=W51[0], size=20) +
        cell_text(short_name, w=W51[1], align="left", size=20) +
        ''.join(cell_text(fmt_num(qty[i] * price) if qty[i] else "0", w=W51[2 + i], size=20) for i in range(12)) +
        cell_text(fmt_num(revenue), w=W51[14], size=20) +
        cell_text(f"{share:.1f}%".replace(".", ","), w=W51[15], size=20)
    ))
    # Строка "количество"
    rows_51.append(row(
        cell_text("", w=W51[0], size=20) +
        cell_text("количество", w=W51[1], align="left", size=20) +
        ''.join(cell_text(str(q), w=W51[2 + i], size=20) for i, q in enumerate(qty)) +
        cell_text(str(sum(qty)), w=W51[14], size=20) +
        cell_text("", w=W51[15], size=20)
    ))
    # Строка "стоимость"
    rows_51.append(row(
        cell_text("", w=W51[0], size=20) +
        cell_text("стоимость", w=W51[1], align="left", size=20) +
        ''.join(cell_text(fmt_num(price), w=W51[2 + i], size=20) for i in range(12)) +
        cell_text("", w=W51[14], size=20) +
        cell_text("", w=W51[15], size=20)
    ))
B(table(rows_51, W51))
B(para(""))

B(para("На доходы от услуг по установке дверей и теневых плинтусов влияет сезонный ремонтный цикл: пик активности приходится на март–октябрь, когда массово сдаются новостройки и собственники переходят к чистовой отделке; в ноябре–феврале наблюдается умеренный спад с локальным предновогодним подъёмом в декабре. Первый месяц проекта целиком отводится под закупку инструмента и оснастки, заявки начинают поступать со второго месяца после публикации профильных объявлений и завершения первых эталонных объектов.", size=28))

B(para("5.2. Показатели экономической эффективности проекта", size=28, before=120))

# Таблица 5.2
W52 = [400, 2400] + [600] * 12 + [900]
HEAD52 = ["№ п/п", "Показатель"] + [f"{i+1} мес." for i in range(12)] + ["ИТОГО"]
rows_52 = []
rows_52.append(row(''.join(cell_text(t, w=W52[i], bold=True, size=20) for i, t in enumerate(HEAD52))))
def row52(num, label, vals, total, *, bold=False):
    return row(
        cell_text(num, w=W52[0], size=20, bold=bold) +
        cell_text(label, w=W52[1], align="left", size=20, bold=bold) +
        ''.join(cell_text(fmt_num(v) if v else "0", w=W52[2 + i], size=20, bold=bold) for i, v in enumerate(vals)) +
        cell_text(fmt_num(total), w=W52[14], bold=True, size=20)
    )
rows_52.append(row52("1", "Доходы, руб.", REV_MONTHS, sum(REV_MONTHS), bold=True))
rows_52.append(row52("2", "Расходы, руб.:", EXPENSES_BY_MONTH, sum(EXPENSES_BY_MONTH), bold=True))
rows_52.append(row52("2.1", "Оборудование", EXP_EQUIPMENT, sum(EXP_EQUIPMENT)))
rows_52.append(row52("2.2", "Расходные материалы", EXP_MATERIALS, sum(EXP_MATERIALS)))
rows_52.append(row52("2.3", "Аренда", EXP_RENT, sum(EXP_RENT)))
rows_52.append(row52("2.4", "Транспортные расходы", EXP_TRANSPORT, sum(EXP_TRANSPORT)))
rows_52.append(row52("2.5", "Телефон, интернет", EXP_COMMS, sum(EXP_COMMS)))
rows_52.append(row52("2.6", "Реклама", EXP_PROMO, sum(EXP_PROMO)))
rows_52.append(row52("2.7", "Непредвиденные расходы, ремонт, амортизация", EXP_UNFORESEEN, sum(EXP_UNFORESEEN)))
rows_52.append(row52("2.8", "Налог (НПД 4%)", EXP_TAX, sum(EXP_TAX)))
rows_52.append(row52("3", "Прибыль, руб. (стр. 1 – стр. 2)", PROFIT_BY_MONTH, PROFIT_TOTAL, bold=True))
rows_52.append(row52("4", "Чистый денежный поток, руб. (стр.3 + стр.6 + стр.7)", CFLOW, sum(CFLOW), bold=True))
rows_52.append(row52("4.1", "Чистый поток нарастающим итогом", CFLOW_CUM, CFLOW_CUM[-1]))
# Строка 5 — рентабельность только в ИТОГО
rows_52.append(row(
    cell_text("5", w=W52[0], size=20) +
    cell_text("Рентабельность, % (стр.3/стр.2)*100%)", w=W52[1], align="left", size=20) +
    ''.join(cell_text("", w=W52[2 + i], size=20) for i in range(12)) +
    cell_text(f"{RENTABILITY:.1f}%".replace(".", ","), w=W52[14], bold=True, size=20)
))
rows_52.append(row52("6", "Стороннее финансирование", EXTERNAL, sum(EXTERNAL)))
rows_52.append(row52("7", "Собственные средства", OWNFUND, sum(OWNFUND)))
B(table(rows_52, W52))

B(para("В расходные материалы включены монтажная пена, метизы и анкера, клей-герметик для скрытых дверных систем, сменные пильные диски и абразивы, наличники и доборы для образцов, средства индивидуальной защиты. Объём расходных материалов изменяется пропорционально количеству выездов в месяц.", size=28, before=120))

B(para(f"5.3. Срок окупаемости проекта {['','один месяц','два месяца','три месяца','четыре месяца','пять месяцев','шесть месяцев','семь месяцев','восемь месяцев','девять месяцев','десять месяцев','одиннадцать месяцев','двенадцать месяцев'][PAYBACK]}.", size=28))
B(para("5.4. Показатели экономической эффективности за 2 года.", size=28))

# Таблица 5.4
rows_54 = []
rows_54.append(row(
    cell_text("№", w=600, bold=True, size=22) +
    cell_text("Показатель", w=4400, bold=True, size=22) +
    cell_text("1-й год", w=2200, bold=True, size=22) +
    cell_text("2-й год", w=2200, bold=True, size=22)
))
rows_54.append(row(
    cell_text("1.", w=600, size=22) +
    cell_text("Доходы, руб.", w=4400, align="left", size=22) +
    cell_text(fmt_num(REV_TOTAL), w=2200, size=22) +
    cell_text(fmt_num(YEAR2_REVENUE), w=2200, size=22)
))
rows_54.append(row(
    cell_text("2.", w=600, size=22) +
    cell_text("Расходы, руб.", w=4400, align="left", size=22) +
    cell_text(fmt_num(EXP_TOTAL), w=2200, size=22) +
    cell_text(fmt_num(YEAR2_EXPENSES), w=2200, size=22)
))
rows_54.append(row(
    cell_text("3.", w=600, size=22) +
    cell_text("Прибыль, руб.", w=4400, align="left", size=22) +
    cell_text(fmt_num(PROFIT_TOTAL), w=2200, size=22) +
    cell_text(fmt_num(YEAR2_PROFIT), w=2200, size=22)
))
rows_54.append(row(
    cell_text("4.", w=600, size=22) +
    cell_text("Рентабельность, %", w=4400, align="left", size=22) +
    cell_text(f"{RENTABILITY:.1f}%".replace(".", ","), w=2200, size=22) +
    cell_text(f"{YEAR2_RENT:.0f}%", w=2200, size=22)
))
B(table(rows_54, [600, 4400, 2200, 2200]))

# ---------- 6. РИСКИ ----------
B(para("6. Описание рисков и мер по их минимизации", bold=True, size=32, align="left", before=120, after=120))

risks = [
    ("Снижение платежеспособности потребителей", "низкая",
     "Комплектные пакеты «дверь + доборы + плинтус под ключ» дешевле раздельного заказа у разных мастеров."),
    ("Эпидемиологическая обстановка. Риск заболеваемости", "низкая",
     "Вакцинация и работа в защитных перчатках и респираторе при пылеобразующих операциях."),
    ("Уменьшение потока клиентов (сезонность)", "низкая",
     "В межсезонье — акцент на коммерческих помещениях и скрытых дверных системах Invisible с длинным циклом продажи."),
    ("Несоответствие ожиданиям клиента", "низкая",
     "Согласование контрольных точек по фотографиям после закладки короба и до чистового монтажа."),
    ("Болезнь", "низкая",
     "Поддерживаю режим, регулярные осмотры, страховой запас по срокам сдачи объектов."),
    ("Появление конкурентов", "средняя",
     "Удержание качества за счёт прецизионной оснастки и накопленного портфолио скрытых дверей."),
    ("Поломка оборудования", "низкая",
     "Транспортировка фрезера и торцовочной пилы в защитных кейсах систейнер, ежемесячная очистка коллектора пылеудаления."),
]
W6 = [3500, 2200, 4500]
rows_6 = []
rows_6.append(row(
    cell_text("Риск", w=W6[0], bold=True, size=22) +
    cell_text("Вероятность реализации (Низкая/средняя/высокая)", w=W6[1], bold=True, size=22) +
    cell_text("Решение (должно быть понятным и конкретным)", w=W6[2], bold=True, size=22)
))
for r, p, s in risks:
    rows_6.append(row(
        cell_text(r, w=W6[0], align="left", size=22) +
        cell_text(p, w=W6[1], size=22) +
        cell_text(s, w=W6[2], align="left", size=22)
    ))
B(table(rows_6, W6))

# ---------- 7. ПРИЛОЖЕНИЯ ----------
B(para("7. Приложение. Личные документы; требуемое оборудование; конкуренты и цены", bold=True, size=32, align="left", before=240, after=240))
B(para(PERSON["fio_short"], size=28, align="right", after=240))
B(para("Приложение 1: Документы, подтверждающие квалификацию", size=28, before=120))
B(para("Приложение 2: Обоснование стоимости оборудования", size=28))
B(para("Приложение 3: Анализ рыночных цен и конкурентов", size=28))

# ---------- SECT PR ----------
sect = '''<w:sectPr>
    <w:pgSz w:w="11906" w:h="16838"/>
    <w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>
    <w:cols w:space="708"/>
    <w:docGrid w:linePitch="360"/>
</w:sectPr>'''

DOCUMENT_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14">
<w:body>
{''.join(body_parts)}
{sect}
</w:body>
</w:document>'''

# ============================================================
# 7. УПАКОВКА В DOCX
# ============================================================
if os.path.exists(WORK_DIR):
    shutil.rmtree(WORK_DIR)
shutil.copytree(TEMPLATE_DIR, WORK_DIR)

# Записываем новый document.xml
with open(os.path.join(WORK_DIR, "word", "document.xml"), "w", encoding="utf-8") as f:
    f.write(DOCUMENT_XML)

# Удаляем header1.xml и медиа — мы на них не ссылаемся
header_path = os.path.join(WORK_DIR, "word", "header1.xml")
if os.path.exists(header_path):
    os.remove(header_path)
media_dir = os.path.join(WORK_DIR, "word", "media")
if os.path.isdir(media_dir):
    shutil.rmtree(media_dir)

# Чистим document.xml.rels от ссылок на header и image
rels_path = os.path.join(WORK_DIR, "word", "_rels", "document.xml.rels")
with open(rels_path, "r", encoding="utf-8") as f:
    rels = f.read()
import re
rels = re.sub(r'<Relationship[^/]+Type="[^"]*relationships/header"[^/]*/>', '', rels)
rels = re.sub(r'<Relationship[^/]+Type="[^"]*relationships/image"[^/]*/>', '', rels)
with open(rels_path, "w", encoding="utf-8") as f:
    f.write(rels)

# Чистим [Content_Types].xml от ссылок на header и image
ct_path = os.path.join(WORK_DIR, "[Content_Types].xml")
with open(ct_path, "r", encoding="utf-8") as f:
    ct = f.read()
ct = re.sub(r'<Override[^/]+PartName="/word/header1\.xml"[^/]*/>', '', ct)
ct = re.sub(r'<Default[^/]+Extension="png"[^/]*/>', '', ct)
with open(ct_path, "w", encoding="utf-8") as f:
    f.write(ct)

# Запаковываем в docx
if os.path.exists(OUT_FILE):
    os.remove(OUT_FILE)

def add_to_zip(zf, root, base):
    for name in os.listdir(root):
        full = os.path.join(root, name)
        rel = os.path.relpath(full, base)
        if os.path.isdir(full):
            add_to_zip(zf, full, base)
        else:
            zf.write(full, rel)

with zipfile.ZipFile(OUT_FILE, "w", zipfile.ZIP_DEFLATED) as zf:
    add_to_zip(zf, WORK_DIR, WORK_DIR)

print(f"\nГотово: {OUT_FILE}")
print(f"Размер: {os.path.getsize(OUT_FILE)} байт")
