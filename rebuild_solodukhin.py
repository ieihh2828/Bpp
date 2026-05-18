#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Реконструкция БП Солодухина А.Р. под жёсткие критерии Рязанской комиссии.
Ниша: профессиональный клининг и комплексная уборка.

Берём БП Воробьева как базу (та же структура 8 таблиц), переделываем под Солодухина:
- Все данные клиента
- Все 5 услуг
- Все таблицы 1.16, 4.1, 4.2, 5.1, 5.2, 5.4, риски
- Все тексты разделов
"""
from __future__ import annotations
import os, shutil, zipfile, copy
import xml.etree.ElementTree as ET

SRC = '/projects/sandbox/Bpp/Бизнес план Воробьев Павел Алексеевич.docx'  # база
WORK = '/tmp/solo_build'
DST = '/projects/sandbox/Bpp/Бизнес план Солодухин Арсений Романович.docx'

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = '{' + W + '}'
ET.register_namespace('w', W)
NS_ALL = {
    'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'cx':  'http://schemas.microsoft.com/office/drawing/2014/chartex',
    'mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o':   'urn:schemas-microsoft-com:office:office',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm':   'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v':   'urn:schemas-microsoft-com:vml',
    'wp14':'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'w':   W,
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
}
for prefix, uri in NS_ALL.items():
    ET.register_namespace(prefix, uri)

# =============================================================================
# ДАННЫЕ КЛИЕНТА
# =============================================================================
CLIENT_FIO = 'Солодухин Арсений Романович'
CLIENT_FIO_SHORT = 'Солодухин А. Р.'
CLIENT_BIRTH = '14.04.2003'
ADDRESS_REG = 'Рязанская обл., г. Рязань, ул. Подгорная, д. 3, кв. 33'
ADDRESS_FACT = 'Рязанская обл., г. Рязань, проезд Щедрина, д. 13А, кв. 3'
PHONE = '+7 (920) 951-86-20'
EMAIL = 'rpriaa365@gmail.com'
INN = '623407078326'
PROJECT_NAME = 'Профессиональный клининг и комплексная уборка квартир и помещений'

MONTHS = ['май','июн','июл','авг','сен','окт','ноя','дек','янв','фев','мар','апр']

SERVICES = [
    ('Генеральная уборка квартиры до 50 кв.м', '4 часа', 5000),
    ('Уборка после ремонта (комплексная) до 50 кв.м', '6 часов', 12000),
    ('Поддерживающая уборка квартиры до 50 кв.м', '2 часа', 3000),
    ('Мытьё окон и остекления (до 4 окон / балкон)', '2 часа', 3500),
    ('Химчистка мягкой мебели или ковра на дому', '2 часа', 4500),
]

QTY = {
    'gen': [0, 5,  4,  4,  4,  5,  5,  6,  1,  2,  4,  5 ],
    'rem': [0, 3,  3,  4,  5,  5,  4,  3,  0,  1,  5,  4 ],
    'sup': [0, 3,  3,  3,  3,  2,  3,  3,  1,  2,  3,  3 ],
    'win': [0, 1,  2,  1,  2,  2,  1,  2,  0,  0,  3,  3 ],
    'him': [0, 1,  1,  2,  2,  2,  2,  2,  1,  1,  1,  1 ],
}
PRICES = {'gen': 5000, 'rem': 12000, 'sup': 3000, 'win': 3500, 'him': 4500}
SVC_KEYS = ['gen','rem','sup','win','him']

REC =  [0, 1500, 1500, 2000, 2500, 3000, 3500, 4500, 2000, 2500, 4000, 4500]
NEPR = [0, 600,  700,  800,  800,  1200, 1200, 1500, 500,  600,  1200, 1500]
PHONE_COST = 650
MAT_PER_ORDER = {'gen': 800, 'rem': 1800, 'sup': 300, 'win': 200, 'him': 400}

# Смета (пересчитываем под клининг)
ESTIMATE = [
    # (item_id, name, qty, price_total, supplier, fund_type)
    ('1.1', 'Экстрактор моющий Santoemma Sabrina SW15', '1', '93700', 'vseinstrumenti.ru', 'grant'),
    ('1.2', 'Пароочиститель профессиональный Karcher SG 4/4', '1', '134800', 'cleaning-city.ru', 'grant'),
    ('1.3', 'Пылесос для сухой/влажной уборки Karcher NT 30/1 Ap L', '1', '54044', 'Яндекс Маркет', 'grant'),
    ('1.4', 'Профессиональный озонатор Otridar', '1', '12436', 'Озон', 'grant'),
    ('1.5', 'Комплект для мытья окон Unger ErgoTec', '1', '9900', 'Seilor', 'grant'),
    ('1.6', 'Набор основных чистящих средств CHEMSPEC (профессиональная химия)', '1', '28750', 'cleanshop77.ru', 'grant'),
    ('1.7', 'Набор уборочный Vileda Professional УльтраСпид Мини (комплект, 2 шт.)', '1', '12198', 'vseinstrumenti.ru', 'transition'),
    ('1.8', 'Комплект спецодежды для клининга У26 КБР (комплект, 2 шт.)', '1', '9242', 'Яндекс Маркет', 'own'),
    ('1.9', 'Стремянка алюминиевая Алюмет 4 ст. с органайзером', '1', '3931', 'vseinstrumenti.ru', 'own'),
]
PROMO_LINE = ('4.1', 'Стартовый рекламный пакет Авито Pro (Турбо-продвижение, 1 мес.)', '1', '5000', 'Avito.ru', 'own')

EQUIP_TOTAL_FIRST_MONTH = 369000  # сумма всех позиций (округлённо)

GRANT = 350000
OWN = 19000

# Расчёт переходной позиции
def calc_transition():
    s = 0
    for i, e in enumerate(ESTIMATE):
        price = int(e[3])
        if s + price <= GRANT:
            s += price
        else:
            return i, GRANT - s, price - (GRANT - s)
    return None, 0, 0

TRANS_IDX, TRANS_GRANT, TRANS_OWN = calc_transition()

# =============================================================================
# УТИЛИТЫ
# =============================================================================
def fmt_rub(n):
    return f'{n:,}'.replace(',', ' ')

def calc_revenue(i):
    return sum(QTY[k][i] * PRICES[k] for k in SVC_KEYS)

def calc_orders(i):
    return sum(QTY[k][i] for k in SVC_KEYS)

def calc_materials(i):
    raw = sum(QTY[k][i] * MAT_PER_ORDER[k] for k in SVC_KEYS)
    return round(raw / 100) * 100 if raw else 0

def calc_transport(i):
    if i == 0: return 0
    n = calc_orders(i)
    return round((1500 + 200 * n) / 100) * 100

def calc_npd(i):
    return round(calc_revenue(i) * 0.04 / 10) * 10

def calc_expenses_breakdown(i):
    osn = EQUIP_TOTAL_FIRST_MONTH if i == 0 else 0
    mat = calc_materials(i)
    arenda = 0
    trans = calc_transport(i)
    phone = 0 if i == 0 else PHONE_COST
    rec = REC[i]
    nepr = NEPR[i]
    npd = calc_npd(i)
    total = osn + mat + arenda + trans + phone + rec + nepr + npd
    return {'osn':osn,'mat':mat,'arenda':arenda,'trans':trans,'phone':phone,'rec':rec,'nepr':nepr,'npd':npd,'total':total}

# =============================================================================
# DOCX I/O
# =============================================================================
def unpack_docx(src, work):
    if os.path.exists(work): shutil.rmtree(work)
    os.makedirs(work)
    with zipfile.ZipFile(src) as z:
        z.extractall(work)

def repack_docx(work, dst):
    if os.path.exists(dst): os.remove(dst)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(work):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, work).replace(os.sep, '/')
                z.write(full, rel)

# =============================================================================
# XML helpers
# =============================================================================
def cell_text(cell):
    return ''.join((t.text or '') for t in cell.iter(NS+'t'))

def first_run_props(cell):
    for r in cell.iter(NS+'r'):
        rpr = r.find(NS+'rPr')
        if rpr is not None: return rpr
    return None

def first_par_props(cell):
    p = cell.find(NS+'p')
    if p is None: return None
    return p.find(NS+'pPr')

def set_cell_text(cell, text, *, bold=None):
    saved_rpr = first_run_props(cell)
    saved_ppr = first_par_props(cell)
    for p in list(cell.findall(NS+'p')):
        cell.remove(p)
    p = ET.SubElement(cell, NS+'p')
    if saved_ppr is not None:
        p.insert(0, copy.deepcopy(saved_ppr))
    r = ET.SubElement(p, NS+'r')
    if saved_rpr is not None:
        r.append(copy.deepcopy(saved_rpr))
    if bold is not None:
        rpr = r.find(NS+'rPr')
        if rpr is None:
            rpr = ET.SubElement(r, NS+'rPr'); r.insert(0, rpr)
        for old in rpr.findall(NS+'b'): rpr.remove(old)
        if bold:
            ET.SubElement(rpr, NS+'b')
    t = ET.SubElement(r, NS+'t')
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

def get_tables(root):
    body = root.find(NS+'body')
    return body.findall('.//' + NS + 'tbl')

def get_paragraphs(root):
    body = root.find(NS+'body')
    return list(body.findall(NS+'p'))

def para_text(p):
    return ''.join((t.text or '') for t in p.iter(NS+'t'))

def replace_paragraph_text(p, new_text):
    saved_rpr = None
    for r in p.findall(NS+'r'):
        rpr = r.find(NS+'rPr')
        if rpr is not None:
            saved_rpr = copy.deepcopy(rpr); break
    for child in list(p):
        if child.tag != NS+'pPr':
            p.remove(child)
    r = ET.SubElement(p, NS+'r')
    if saved_rpr is not None:
        r.append(saved_rpr)
    t = ET.SubElement(r, NS+'t')
    t.text = new_text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

def find_para_by_substr(paragraphs, needle):
    for i, p in enumerate(paragraphs):
        if needle in para_text(p):
            return i
    return -1

def insert_paragraph_after(parent, ref_p, text):
    new_p = ET.Element(NS+'p')
    ppr = ref_p.find(NS+'pPr')
    if ppr is not None:
        new_p.append(copy.deepcopy(ppr))
    saved_rpr = None
    for r in ref_p.findall(NS+'r'):
        rpr = r.find(NS+'rPr')
        if rpr is not None:
            saved_rpr = copy.deepcopy(rpr); break
    r = ET.SubElement(new_p, NS+'r')
    if saved_rpr is not None:
        r.append(saved_rpr)
    t = ET.SubElement(r, NS+'t')
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    children = list(parent)
    idx = children.index(ref_p)
    parent.insert(idx + 1, new_p)
    return new_p

def clone_row(row):
    return copy.deepcopy(row)

# =============================================================================
# ОСНОВНАЯ ЛОГИКА
# =============================================================================
def main():
    print(f'>> Распаковка БП Воробьева как базы -> {WORK}')
    unpack_docx(SRC, WORK)
    doc_xml_path = os.path.join(WORK, 'word', 'document.xml')
    tree = ET.parse(doc_xml_path)
    root = tree.getroot()
    body = root.find(NS+'body')
    
    tables = get_tables(root)
    print(f'>> Таблиц найдено: {len(tables)}')
    assert len(tables) == 8

    # Финмодель
    annual_rev = sum(calc_revenue(i) for i in range(12))
    annual_exp = sum(calc_expenses_breakdown(i)['total'] for i in range(12))
    annual_profit = annual_rev - annual_exp
    annual_cashflow = annual_profit + GRANT + OWN
    rentab = annual_profit / annual_exp * 100
    cum = 0; payback = 0
    for i in range(12):
        cum += calc_revenue(i) - calc_expenses_breakdown(i)['total']
        if cum > 0:
            payback = i + 1
            break
    print(f'>> Выр={annual_rev:,} Расх={annual_exp:,} Прб={annual_profit:,} ЧДП={annual_cashflow:,} Рент={rentab:.1f}% Окуп={payback}')

    # =========================================================================
    # СНАЧАЛА: ГЛОБАЛЬНАЯ ЗАМЕНА ВСЕХ УПОМИНАНИЙ ВОРОБЬЕВА НА СОЛОДУХИНА
    # =========================================================================
    print('>> Глобальная замена данных клиента')
    REPLACEMENTS = [
        # ФИО
        ('Воробьев Павел Алексеевич', CLIENT_FIO),
        ('Воробьев П. А.', CLIENT_FIO_SHORT),
        ('Воробьев П.А.', CLIENT_FIO_SHORT),
        ('Воробьева Павла Алексеевича', 'Солодухина Арсения Романовича'),
        # Адреса
        ('Рязанская обл., г. Рязань, ул. Гайдара (СОЛОТЧА), д. 3, кв. 8', ADDRESS_REG),
        ('Рязанская обл., г. Рязань, ул. Гайдара (Солотча), д. 3, кв. 8', ADDRESS_REG),
        # Телефон, email
        ('89156053803', PHONE),
        ('vorobevpavel695@gmail.com', EMAIL),
        # ИНН
        ('623402569401', INN),
        # Дата рождения
        ('25.06.2004', CLIENT_BIRTH),
        # Название проекта
        ('Монтаж и обслуживание кондиционеров', PROJECT_NAME),
        ('Монтаж кондиционеров', PROJECT_NAME),
    ]
    
    # Применяем ко всем <w:t> элементам
    for elem in root.iter(NS+'t'):
        if elem.text:
            for old, new in REPLACEMENTS:
                if old in elem.text:
                    elem.text = elem.text.replace(old, new)

    # =========================================================================
    # ТАБЛИЦА #0: 1.16 Резюме
    # =========================================================================
    print('>> [#0] 1.16 Резюме')
    t0 = tables[0]
    rows0 = t0.findall(NS+'tr')
    set_cell_text(rows0[1].findall(NS+'tc')[1], f'{fmt_rub(annual_rev)} руб.')
    set_cell_text(rows0[2].findall(NS+'tc')[1], f'{fmt_rub(annual_exp)} руб.')
    set_cell_text(rows0[3].findall(NS+'tc')[1], f'{fmt_rub(annual_profit)} руб.')
    set_cell_text(rows0[4].findall(NS+'tc')[1], f'{fmt_rub(annual_cashflow)} руб.')
    set_cell_text(rows0[5].findall(NS+'tc')[1], f'{rentab:.1f}%'.replace('.', ','))
    set_cell_text(rows0[6].findall(NS+'tc')[1], f'{payback} месяцев')

    # =========================================================================
    # ТАБЛИЦА #1: Прайс
    # =========================================================================
    print('>> [#1] Прайс — 5 услуг')
    t1 = tables[1]
    rows1 = t1.findall(NS+'tr')
    template_row = clone_row(rows1[1])
    for r in rows1[1:]:
        t1.remove(r)
    for name, time_str, price in SERVICES:
        nr = clone_row(template_row)
        cells = nr.findall(NS+'tc')
        set_cell_text(cells[0], name)
        set_cell_text(cells[1], time_str)
        set_cell_text(cells[2], str(price))
        t1.append(nr)

    # =========================================================================
    # ТАБЛИЦА #2: 4.1 Источники
    # =========================================================================
    print('>> [#2] 4.1 Источники')
    t2 = tables[2]
    rows2 = t2.findall(NS+'tr')
    set_cell_text(rows2[1].findall(NS+'tc')[2], '350 000')
    set_cell_text(rows2[2].findall(NS+'tc')[2], fmt_rub(OWN))
    set_cell_text(rows2[3].findall(NS+'tc')[2], '0')
    set_cell_text(rows2[4].findall(NS+'tc')[2], '0')
    set_cell_text(rows2[5].findall(NS+'tc')[-1], fmt_rub(GRANT + OWN))

    # =========================================================================
    # ТАБЛИЦА #3: 4.2 Смета — переписываем целиком
    # =========================================================================
    print('>> [#3] 4.2 Смета — переписываем')
    t3 = tables[3]
    rows3 = t3.findall(NS+'tr')
    # Структура: row[0]=шапка, row[1]="1.* Основные средства", row[2]..row[16]=1.1..1.15 (15 позиций),
    # row[17]="2.* Запасы", row[18]="3.* Аренда", row[19]="4.* Продвижение", row[20]="5.* Документы", row[21]=ИТОГО
    
    # У нас 9 позиций основных средств + 1 в разделе "Продвижение"
    # Удалим лишние строки (с 1.10 по 1.15 в шаблоне Воробьева)
    # rows3[2] = 1.1, rows3[3] = 1.2, ... rows3[16] = 1.15
    # Нам нужно оставить только 9 (rows3[2..10]) и удалить rows3[11..16]
    # Затем заменить содержимое 1.1..1.9 на наши данные
    
    # Запомним шаблоны строк
    item_template = clone_row(rows3[2])  # шаблон строки 1.1
    
    # Удалим все 15 строк 1.1..1.15 (rows3[2..16])
    rows_to_remove = rows3[2:17]
    for r in rows_to_remove:
        t3.remove(r)
    
    # Перечитаем строки
    rows3 = t3.findall(NS+'tr')
    # Теперь rows3[0]=шапка, rows3[1]="1.*", rows3[2]="2.*", rows3[3]="3.*", rows3[4]="4.*", rows3[5]="5.*", rows3[6]=ИТОГО
    
    # Вставим 9 строк ОСНОВНЫХ СРЕДСТВ после "1.* Основные средства"
    section1_row = rows3[1]
    children = list(t3)
    insert_at = children.index(section1_row) + 1
    
    for idx, (item_id, name, qty, price, sup, fund_type) in enumerate(ESTIMATE):
        new_row = clone_row(item_template)
        cells = new_row.findall(NS+'tc')
        set_cell_text(cells[0], item_id)
        set_cell_text(cells[1], name)
        set_cell_text(cells[2], qty)
        set_cell_text(cells[3], price)
        set_cell_text(cells[4], price)
        set_cell_text(cells[5], sup)
        # fund text
        if fund_type == 'grant':
            fund_text = 'Средства соцконтракта'
        elif fund_type == 'transition':
            fund_text = 'Средства соцконтракта, собственные средства'
        else:
            fund_text = 'Собственные средства'
        set_cell_text(cells[6], fund_text)
        t3.insert(insert_at + idx, new_row)
    
    # Перечитаем
    rows3 = t3.findall(NS+'tr')
    # Найдём строку "4.* Продвижение" — ищем по началу первой ячейки "4.*"
    section4_idx = None
    for i, r in enumerate(rows3):
        cells = r.findall(NS+'tc')
        if cells:
            cell0_text = cell_text(cells[0]).strip()
            cell1_text = cell_text(cells[1])
            if cell0_text == '4.*' and 'Продвижение' in cell1_text:
                section4_idx = i
                break
    
    if section4_idx:
        # Вставим строку 4.1 (Авито)
        new_row = clone_row(item_template)
        cells = new_row.findall(NS+'tc')
        set_cell_text(cells[0], PROMO_LINE[0])
        set_cell_text(cells[1], PROMO_LINE[1])
        set_cell_text(cells[2], PROMO_LINE[2])
        set_cell_text(cells[3], PROMO_LINE[3])
        set_cell_text(cells[4], PROMO_LINE[3])
        set_cell_text(cells[5], PROMO_LINE[4])
        set_cell_text(cells[6], 'Собственные средства')
        t3.insert(section4_idx + 1, new_row)
    
    # ИТОГО
    rows3 = t3.findall(NS+'tr')
    itogo_row = rows3[-1]
    itogo_cells = itogo_row.findall(NS+'tc')
    if len(itogo_cells) >= 5:
        set_cell_text(itogo_cells[4], fmt_rub(GRANT + OWN), bold=True)

    # =========================================================================
    # ТАБЛИЦА #4: 5.1 Доходы
    # =========================================================================
    print('>> [#4] 5.1 Доходы')
    t4 = tables[4]
    rows4 = t4.findall(NS+'tr')
    
    # Строка "Доходы всего" (rows4[2])
    row_total = rows4[2]
    cells = row_total.findall(NS+'tc')
    annual = 0
    for i in range(12):
        rev_i = calc_revenue(i)
        annual += rev_i
        set_cell_text(cells[2 + i], fmt_rub(rev_i) if rev_i else '0')
    set_cell_text(cells[14], fmt_rub(annual), bold=True)
    set_cell_text(cells[15], '100')
    
    # 5 услуг, каждая 3 строки (название/количество/стоимость)
    SVC_ROW_LAYOUT = [('gen', 3), ('rem', 6), ('sup', 9), ('win', 12), ('him', 15)]
    for key, idx in SVC_ROW_LAYOUT:
        name_row = rows4[idx]
        qty_row = rows4[idx + 1]
        sum_row = rows4[idx + 2]
        svc_idx = SVC_KEYS.index(key)
        svc_name = SERVICES[svc_idx][0]
        price = PRICES[key]
        nrcells = name_row.findall(NS+'tc')
        qcells = qty_row.findall(NS+'tc')
        scells = sum_row.findall(NS+'tc')
        set_cell_text(nrcells[1], svc_name)
        total_qty = 0; total_sum = 0
        for i in range(12):
            q = QTY[key][i]
            revenue_i = q * price
            total_qty += q
            total_sum += revenue_i
            set_cell_text(nrcells[2 + i], fmt_rub(revenue_i) if revenue_i else '0')
            set_cell_text(qcells[2 + i], str(q) if q else '0')
            set_cell_text(scells[2 + i], fmt_rub(price))
        set_cell_text(nrcells[14], fmt_rub(total_sum), bold=True)
        set_cell_text(qcells[14], str(total_qty))
        set_cell_text(scells[14], '')
        share = total_sum / annual * 100 if annual else 0
        set_cell_text(nrcells[15], f'{share:.1f}'.replace('.', ','))
        set_cell_text(qcells[15], '')
        set_cell_text(scells[15], '')

    # =========================================================================
    # ТАБЛИЦА #5: 5.2 Расходы
    # =========================================================================
    print('>> [#5] 5.2 Расходы')
    t5 = tables[5]
    rows5 = t5.findall(NS+'tr')
    
    def fill_row(row_idx, monthly_values, bold_total=False):
        row = rows5[row_idx]
        cells = row.findall(NS+'tc')
        for i in range(12):
            v = monthly_values[i]
            if v == 0:
                set_cell_text(cells[2 + i], '0')
            else:
                set_cell_text(cells[2 + i], fmt_rub(v))
        total = sum(monthly_values)
        set_cell_text(cells[14], fmt_rub(total), bold=bold_total)
    
    inc_arr = [calc_revenue(i) for i in range(12)]
    exp_breakdown = [calc_expenses_breakdown(i) for i in range(12)]
    
    fill_row(2, inc_arr, bold_total=True)
    fill_row(4, [b['osn'] for b in exp_breakdown])
    fill_row(5, [b['mat'] for b in exp_breakdown])
    fill_row(6, [b['arenda'] for b in exp_breakdown])
    fill_row(7, [b['trans'] for b in exp_breakdown])
    fill_row(8, [b['phone'] for b in exp_breakdown])
    fill_row(9, [b['rec'] for b in exp_breakdown])
    fill_row(10, [b['nepr'] for b in exp_breakdown])
    fill_row(11, [b['npd'] for b in exp_breakdown])
    
    exp_arr = [b['total'] for b in exp_breakdown]
    fill_row(3, exp_arr, bold_total=True)
    
    profit_arr = [inc_arr[i] - exp_arr[i] for i in range(12)]
    fill_row(12, profit_arr, bold_total=True)
    
    grant_arr = [GRANT] + [0]*11
    own_arr = [OWN] + [0]*11
    cf_arr = [profit_arr[i] + grant_arr[i] + own_arr[i] for i in range(12)]
    fill_row(13, cf_arr, bold_total=True)
    
    # 4.1 нарастающим итогом
    cum2 = 0
    cum_arr = []
    for v in cf_arr:
        cum2 += v
        cum_arr.append(cum2)
    row14 = rows5[14]; cells14 = row14.findall(NS+'tc')
    for i in range(12):
        set_cell_text(cells14[2 + i], fmt_rub(cum_arr[i]))
    set_cell_text(cells14[14], fmt_rub(cum_arr[-1]), bold=True)
    
    # Рентабельность — пусто помесячно, ИТОГО годовая
    row15 = rows5[15]; cells15 = row15.findall(NS+'tc')
    for i in range(12):
        set_cell_text(cells15[2 + i], '')
    annual_rent = sum(profit_arr) / sum(exp_arr) * 100
    set_cell_text(cells15[14], f'{annual_rent:.1f}%'.replace('.', ','), bold=True)
    
    fill_row(16, grant_arr)
    fill_row(17, own_arr)

    # =========================================================================
    # ТАБЛИЦА #6: 5.4 Год 2
    # =========================================================================
    print('>> [#6] 5.4 Год 2')
    annual_rev_y2 = 970000
    annual_exp_y2 = sum(b['total'] - b['osn'] for b in exp_breakdown) + 80000
    annual_exp_y2 = int(annual_exp_y2 * 1.05)
    annual_profit_y2 = annual_rev_y2 - annual_exp_y2
    rent_y2 = annual_profit_y2 / annual_exp_y2 * 100
    
    t6 = tables[6]
    rows6 = t6.findall(NS+'tr')
    set_cell_text(rows6[1].findall(NS+'tc')[2], f'{fmt_rub(annual_rev)} руб.')
    set_cell_text(rows6[1].findall(NS+'tc')[3], f'{fmt_rub(annual_rev_y2)} руб.')
    set_cell_text(rows6[2].findall(NS+'tc')[2], f'{fmt_rub(annual_exp)} руб.')
    set_cell_text(rows6[2].findall(NS+'tc')[3], f'{fmt_rub(annual_exp_y2)} руб.')
    set_cell_text(rows6[3].findall(NS+'tc')[2], f'{fmt_rub(annual_profit)} руб.')
    set_cell_text(rows6[3].findall(NS+'tc')[3], f'{fmt_rub(annual_profit_y2)} руб.')
    set_cell_text(rows6[4].findall(NS+'tc')[2], f'{rentab:.1f}%'.replace('.', ','))
    set_cell_text(rows6[4].findall(NS+'tc')[3], f'{rent_y2:.0f}%')

    # =========================================================================
    # ТАБЛИЦА #7: Риски — переписываем
    # =========================================================================
    print('>> [#7] Риски')
    RISKS = [
        ('Снижение платежеспособности потребителей', 'средняя', 'Поддерживать пакет «эконом-уборка» с минимальной комплектацией работ.'),
        ('Эпидемиологическая обстановка. Риск заболеваемости', 'низкая', 'Соблюдать санитарные нормы при работе в помещении заказчика.'),
        ('Уменьшение потока клиентов (сезонность)', 'средняя', 'Зимой переключаться на B2B-сегмент и поддерживающую уборку.'),
        ('Несоответствие ожиданиям клиента', 'средняя', 'Заключать договор с фотоотчётом по каждому этапу работ.'),
        ('Болезнь', 'низкая', 'Вести здоровый образ жизни и поддерживать резервный фонд.'),
        ('Появление конкурентов', 'средняя', 'Удерживать рейтинг 4,9+ на Авито за счёт качества и фотоотчётов.'),
        ('Поломка оборудования', 'средняя', 'Иметь резервный комплект ключевого инструмента.'),
    ]
    t7 = tables[7]
    rows7 = t7.findall(NS+'tr')
    for i, (risk, prob, sol) in enumerate(RISKS):
        cells = rows7[i+1].findall(NS+'tc')
        set_cell_text(cells[0], risk)
        set_cell_text(cells[1], prob.capitalize())
        set_cell_text(cells[2], sol)

    # =========================================================================
    # ПАРАГРАФЫ — ТЕКСТОВЫЕ ПРАВКИ
    # =========================================================================
    print('>> Текстовые правки в параграфах')
    paragraphs = get_paragraphs(root)
    
    # 1.1 ФИО полностью (если есть)
    # 1.7 Образование
    idx = find_para_by_substr(paragraphs, '1.7. Образование')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], '1.7. Образование: Среднее общее (11 классов).')
    
    # 1.8 Стаж
    idx = find_para_by_substr(paragraphs, '1.8.')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], '1.8. Общий стаж работы, трудовая деятельность в течение последних 3 лет: имел опыт работы в крупных логистических компаниях ООО «РВБ», ООО «Вайлдберриз», ООО «ТТМ Центр» (работа на складских комплексах). В настоящее время официально не трудоустроен.')
    
    # 1.9 Опыт по направлению
    idx = find_para_by_substr(paragraphs, '1.9.')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], '1.9. Опыт работы (предпринимательской деятельности) по выбранному направлению: опыт работы в сфере профессиональной уборки имеется более 2 лет. Наработан практический опыт в составе клининговых бригад при выполнении генеральных уборок и послестроительных работ на объектах в г. Рязани, освоено профессиональное оборудование (моющие пылесосы, парогенераторы, оборудование для мытья окон), отработаны технологии работы с профессиональной химией и обработки различных типов поверхностей.')
    
    # 1.10 Обучение
    idx = find_para_by_substr(paragraphs, '1.10.')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], '1.10. Потребность в обучении/повышении квалификации с обоснованием: Потребность в обучении отсутствует. Имеются необходимые практические навыки.')
    
    # 1.13 ОКВЭД
    idx = find_para_by_substr(paragraphs, '1.13. Виды деятельности (ОКВЭД)')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], '1.13. Виды деятельности (ОКВЭД): 81.21 — Деятельность по общей уборке зданий.')
    elif (idx := find_para_by_substr(paragraphs, '43.22')) >= 0:
        # У Воробьева был 43.22, заменим
        for t in paragraphs[idx].iter(NS+'t'):
            if t.text and '43.22' in t.text:
                t.text = '81.21 — Деятельность по общей уборке зданий.'
                # удалим всё что после "43.22"
                paragraphs[idx]
                break
    
    # 1.15 Месторасположение
    idx = find_para_by_substr(paragraphs, '1.15. Месторасположение')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], f'1.15. Месторасположение: {ADDRESS_FACT}. Деятельность носит выездной характер. Для транспортировки оборудования и выезда на объекты имеется легковой автомобиль. Хранение инструмента и расходных материалов — в собственном гараже рядом с домом, аренда помещения не требуется.')
    
    # Адреса 1.3 / 1.4
    idx = find_para_by_substr(paragraphs, '1.3. Адрес регистрации')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], f'1.3. Адрес регистрации: {ADDRESS_REG}.')
    idx = find_para_by_substr(paragraphs, '1.4. Адрес фактического')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], f'1.4. Адрес фактического проживания: {ADDRESS_FACT}.')
    
    # 2.1 — Краткое описание бизнеса по проекту (стандартный заголовок шаблона)
    NEW_2_1 = (
        '2.1 Краткое описание бизнеса по проекту: Организация мобильной службы профессиональной уборки квартир, '
        'частных домов и коммерческих помещений на территории г. Рязани и Рязанской области. '
        'Деятельность направлена на оказание широкого спектра клининговых услуг с применением профессионального '
        'оборудования и сертифицированной химии: генеральная уборка, уборка после ремонта, регулярная поддерживающая '
        'уборка квартир и небольших офисов, мытьё окон и оконных блоков (включая безразводную мойку), химчистка '
        'мягкой мебели, ковров и текстильных покрытий, дезодорирование помещений с помощью озонатора. Все работы '
        'выполняются с выездом к заказчику в согласованное время. Парк инструмента включает моющий экстрактор для '
        'глубокой чистки тканей и ковров, парогенератор для удаления стойких загрязнений и работы после ремонта, '
        'промышленный пылесос для сухой и влажной уборки, профессиональный комплект для мытья окон и набор '
        'специализированной химии под разные типы поверхностей. Такой комплект позволяет выполнять задачи, '
        'недоступные исполнителям с бытовым инвентарём, и обеспечивать качество, сопоставимое с крупными клининговыми '
        'компаниями, при заметно меньшей итоговой цене за счёт работы в режиме самозанятого без офисных накладных расходов.'
    )
    idx = find_para_by_substr(paragraphs, '2.1 Услуги')
    if idx < 0:
        idx = find_para_by_substr(paragraphs, '2.1 Краткое')
    if idx < 0:
        idx = find_para_by_substr(paragraphs, '2.1 ')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], NEW_2_1)
    
    # Описание услуг (5 пунктов) — найти параграф с "В перечень услуг"
    idx = find_para_by_substr(paragraphs, 'В перечень услуг')
    if idx < 0:
        idx = find_para_by_substr(paragraphs, 'перечень услуг входят')
    if idx >= 0:
        new_text = (
            'Перечень оказываемых услуг включает пять направлений. '
            'Генеральная уборка предполагает глубокую мойку всех поверхностей, чистку сантехники, мытьё окон, '
            'обеспыливание мебели и обработку напольных покрытий моющим экстрактором. '
            'Уборка после ремонта рассчитана на удаление строительной пыли, остатков клея, краски, затирки и шпатлёвки '
            'с использованием парогенератора и промышленного пылесоса; включает мойку оконных блоков и подоконников, '
            'обеспыливание стен и потолков. '
            'Поддерживающая уборка — регулярная (еженедельная или раз в две недели) уборка квартир для постоянных клиентов. '
            'Мытьё окон и остекления выполняется профессиональным комплектом для безразводной мойки и охватывает '
            'до четырёх оконных блоков либо балкон. '
            'Химчистка мягкой мебели и ковров производится моющим экстрактором с пенной обработкой и щелочной химией; '
            'при необходимости — обработка озонатором для нейтрализации запахов.'
        )
        replace_paragraph_text(paragraphs[idx], new_text)
    
    # 2.2 Целевая аудитория (стандартный заголовок шаблона)
    idx = find_para_by_substr(paragraphs, '2.2 Целевая аудитория')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], '2.2 Целевая аудитория: проект ориентирован на два основных сегмента. Базовый сегмент — частные клиенты: владельцы квартир в новостройках, заказывающие уборку после ремонта; собственники жилья на вторичном рынке, обращающиеся за генеральной и регулярной уборкой; а также семьи с детьми, для которых важна безопасная профессиональная химия и качественная чистка мягкой мебели и ковров. Дополнительный сегмент — небольшие коммерческие помещения (офисы, магазины и салоны), заинтересованные в регулярной уборке по договору. Со временем планируется наращивание доли постоянных клиентов на поддерживающей уборке.')
    
    # 2.3 Конкуренция — переписываем (без брендов, без "120 объявлений", без "письменной гарантии")
    # Стандартный заголовок шаблона: "2.3. Конкурентная среда, конкурентная позиция бизнеса:"
    NEW_KONK_LINES = [
        '2.3. Конкурентная среда, конкурентная позиция бизнеса:',
        'Анализ Авито показал, что в Рязани услуги уборки квартир и помещений предлагают исполнители трёх типов, существенно различающихся по цене и подходу к работе.',
        'Первый тип — клининговые компании с офисом и штатом сотрудников. Стоимость генеральной уборки квартиры 1–2 комнат у них находится в диапазоне 6 000–10 000 руб., уборки после ремонта — от 12 000 до 18 000 руб. Сильная сторона таких компаний — возможность работать с юридическими лицами по безналу и выезд бригады из нескольких человек. Слабая — высокий чек за счёт содержания офиса и постоянного штата, а также длительное ожидание свободной даты в пиковые месяцы (1–2 недели).',
        'Второй тип — самозанятые и частные исполнители с бытовым инвентарём. На них приходится примерно половина объявлений в категории; цены, как правило, считаются за квадратный метр и стартуют от 120 руб./кв.м, без фиксированной стоимости заказа. Сильная сторона — низкая цена. Слабые: бытовое оборудование вместо профессионального (обычные пылесосы, моющие средства из розничной сети), отсутствие парогенератора и моющего экстрактора, нет официального чека и договора. Для задач после ремонта или химчистки мебели такие исполнители, как правило, не подходят.',
        'Третий тип — самозанятые мастера, которые работают с полноценным профессиональным оборудованием. Это и есть рыночная ниша проекта. Стоимость генеральной уборки в этом сегменте — 4 500–6 500 руб., после ремонта — 10 000–15 000 руб. Цены проекта установлены по медиане сегмента.',
        'Позиционирование проекта строится на сочетании трёх элементов: профессиональный инвентарь (моющий экстрактор, парогенератор, промышленный пылесос для сухой и влажной уборки, средства для мытья окон без разводов, специализированная химия для разных типов поверхностей); работа в режиме самозанятого с официальным чеком и договором на каждую услугу; цена на уровне середины рынка, а не премиального сегмента. Такая комбинация даёт клиенту качество, сопоставимое с крупным клинингом, при стоимости частного мастера. Подробное сравнение цен и предложений по конкурентам — в Приложении 2.'
    ]
    idx = find_para_by_substr(paragraphs, 'Крупные фирмы')
    if idx == -1:
        idx = find_para_by_substr(paragraphs, '2.3. Конкурентная')
    if idx == -1:
        idx = find_para_by_substr(paragraphs, 'Конкурентная среда')
    if idx >= 0:
        target_p = paragraphs[idx]
        replace_paragraph_text(target_p, NEW_KONK_LINES[0])
        # снимаем bold с заголовка 2.3 (по требованию)
        for r in target_p.findall(NS+'r'):
            rpr = r.find(NS+'rPr')
            if rpr is not None:
                for old in rpr.findall(NS+'b'):
                    rpr.remove(old)
                for old in rpr.findall(NS+'bCs'):
                    rpr.remove(old)
        prev = target_p
        # Перед вставкой удалим следующие параграфы конкурентного блока (от Воробьева)
        # Но безопаснее: просто вставляем после
        for line in NEW_KONK_LINES[1:]:
            prev = insert_paragraph_after(body, prev, line)
        # Удалим следующие параграфы которые остались от воробьевского текста
        # Найдём параграфы между нашим блоком и "2.4 Каналы реализации"
        paragraphs_now = get_paragraphs(root)
        last_inserted_idx = paragraphs_now.index(prev)
        # Удаляем всё между last_inserted и до "2.4"
        i = last_inserted_idx + 1
        while i < len(paragraphs_now):
            txt = para_text(paragraphs_now[i])
            if '2.4' in txt[:6] or 'Каналы реализации' in txt or '3.' in txt[:4]:
                break
            # Удалим если это продолжение старого конкурентного текста
            if any(kw in txt for kw in ['Воробьев', 'кондицион', 'сплит', 'монтаж', '146', 'Value', 'Testo', 'Bosch', 'фреон', 'вакуум']):
                body.remove(paragraphs_now[i])
                paragraphs_now = get_paragraphs(root)  # перечитать
                continue
            i += 1
    
    # Перечитать параграфы
    paragraphs = get_paragraphs(root)
    
    # 5.1 текст — кратко
    NEW_5_1 = (
        'Первый месяц (май) выручки не даёт — он уходит на закупку оборудования, оформление самозанятости и '
        'запуск рекламы на Авито. С июня начинается основная работа: летом стабильное плато 8–14 заказов в месяц, '
        'осенью спрос растёт за счёт сдачи новостроек, в ноябре–декабре — предновогодний пик. Январь традиционно '
        'слабый, с февраля рынок оживает, а в марте–апреле наступает весенний всплеск (мойка окон, генеральные '
        'перед дачным сезоном).'
    )
    idx = find_para_by_substr(paragraphs, 'Выручка в первый месяц')
    if idx < 0:
        idx = find_para_by_substr(paragraphs, 'Первый месяц проекта')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], NEW_5_1)
    
    # 5.2 текст — кратко
    NEW_5_2 = (
        'В первом месяце расходы — ' + fmt_rub(EQUIP_TOTAL_FIRST_MONTH) + ' руб. (закупка оборудования и стартовой химии). '
        'Они полностью покрываются средствами соцконтракта (350 000 руб.) и собственными вложениями (' + fmt_rub(OWN) + ' руб.), '
        'кассового разрыва нет. Со второго месяца расходы переменные: химия и транспорт зависят от количества заказов, '
        'налог НПД 4% уплачивается с выручки, реклама на Авито — от 1 500 до 4 500 руб. в зависимости от сезона.'
    )
    idx = find_para_by_substr(paragraphs, 'В первый месяц расходы составляют')
    if idx < 0:
        idx = find_para_by_substr(paragraphs, 'Расходная часть первого месяца')
    if idx < 0:
        idx = find_para_by_substr(paragraphs, 'расходы составляют')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], NEW_5_2)
    
    # 5.3 окупаемость
    idx = find_para_by_substr(paragraphs, 'Срок окупаемости')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], f'5.3. Срок окупаемости проекта: {payback} мес.')
    
    # 3.3 СИЗ — поскольку у клининга нет промальпа, упрощаем
    idx = find_para_by_substr(paragraphs, '3.3. Требования к персоналу')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], '3.3. Требования к персоналу: требуется уверенное знание технологий профессиональной уборки, владение профессиональным оборудованием (моющий экстрактор, парогенератор, промышленный пылесос), знание свойств щёлочной и кислотной химии, навыки безопасной работы со стремянкой при мытье окон. Автор проекта этими навыками уже обладает.')

    # ---- ДОПОЛНИТЕЛЬНЫЕ ТОЧЕЧНЫЕ ПРАВКИ ----

    # 1.2. Дата рождения — убрать «г.р.» (получалось «14.04.2003 г.р. г.»)
    idx = find_para_by_substr(paragraphs, '1.2. Дата рождения')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], f'1.2. Дата рождения: {CLIENT_BIRTH} г.')

    # 2.8. Продвижение — переписать под клининг (был воробьёвский «чистый монтаж, аккуратные штробы»)
    idx = find_para_by_substr(paragraphs, '2.8. Продвижение')
    if idx < 0:
        idx = find_para_by_substr(paragraphs, 'аккуратные штробы')
    if idx < 0:
        idx = find_para_by_substr(paragraphs, 'чистый монтаж')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx],
            '2.8. Продвижение. Размещение объявлений на Авито с использованием платных пакетов продвижения '
            'для попадания в топ выдачи. Публикация фотоотчётов «до/после» (отмытые окна, очищенные ковры, '
            'квартиры после ремонта) в социальных сетях. Программа лояльности — скидка на следующую уборку '
            'за оставленный отзыв с фото.'
        )

    # 3.1 Планируемый график работы (стандартный заголовок шаблона)
    idx = find_para_by_substr(paragraphs, '3.1. с 09:00')
    if idx < 0:
        idx = find_para_by_substr(paragraphs, '3.1. Планируемый график')
    if idx < 0:
        idx = find_para_by_substr(paragraphs, 'без выходных')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx],
            '3.1. Планируемый график работы — с 09:00 до 21:00 без выходных. Выезд на объекты осуществляется по '
            'предварительной договорённости с заказчиком.'
        )

    # 3.2 Штатное расписание (стандартный заголовок шаблона — без двоеточия)
    idx = find_para_by_substr(paragraphs, '3.2. Штатное расписание')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx],
            '3.2. Штатное расписание дополнительные сотрудники не требуются.'
        )

    # 4.3. Помесячный план закупок — был воробьёвский текст про вакуумный насос, фреон и т.д.
    idx = find_para_by_substr(paragraphs, 'вакуумный насос')
    if idx < 0:
        idx = find_para_by_substr(paragraphs, 'парогенератором')
    if idx < 0:
        idx = find_para_by_substr(paragraphs, 'В течение одного месяца с момента поступления')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx],
            'В течение одного месяца с момента поступления средств социального контракта проводится закупка '
            'всего необходимого оборудования и инвентаря. В первую очередь приобретается дорогостоящее основное '
            'оборудование, обеспечивающее старт работы: моющий экстрактор, парогенератор и промышленный пылесос '
            'для сухой и влажной уборки. Параллельно закупаются озонатор, комплект для мытья окон и стартовый '
            'набор профессиональной химии. В завершение — расходный инвентарь, спецодежда и стремянка. Закупка '
            'товара поэтапная, под фактическую потребность каждого этапа работ.'
        )

    # 7. Приложение — упростить (без перечисления, как у Воробьёва)
    idx = find_para_by_substr(paragraphs, '7. Приложение')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], '7. Приложение')

    # Приложение 1 — удалить (документы квалификации)
    idx = find_para_by_substr(paragraphs, 'Приложение 1: Документы')
    if idx >= 0:
        body.remove(paragraphs[idx])
    
    # Приложение 2 → Приложение 1
    paragraphs = get_paragraphs(root)
    idx = find_para_by_substr(paragraphs, 'Приложение 2: Обоснование')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], 'Приложение 1: Обоснование стоимости оборудования')
    
    # Приложение 3 → Приложение 2
    paragraphs = get_paragraphs(root)
    idx = find_para_by_substr(paragraphs, 'Приложение 3: Анализ')
    if idx >= 0:
        replace_paragraph_text(paragraphs[idx], 'Приложение 2: Анализ рыночных цен и конкурентов')
    
    # Также заменим "(Приложение 2)" на "(Приложение 1)" в тексте после сметы
    for elem in root.iter(NS+'t'):
        if elem.text:
            if '(Приложение 2)' in elem.text:
                elem.text = elem.text.replace('(Приложение 2)', '(Приложение 1)')
            if 'Приложении 3' in elem.text:
                elem.text = elem.text.replace('Приложении 3', 'Приложении 2')
            if 'Приложение 3' in elem.text:
                elem.text = elem.text.replace('Приложение 3', 'Приложение 2')

    # =========================================================================
    # СОХРАНЕНИЕ
    # =========================================================================
    print('>> Сохранение')
    tree.write(doc_xml_path, xml_declaration=True, encoding='UTF-8', method='xml')
    repack_docx(WORK, DST)
    sz = os.path.getsize(DST)
    print(f'>> Готово: {DST} ({sz} байт)')
    
    print('\n' + '='*60)
    print('ИТОГИ ФИНМОДЕЛИ')
    print('='*60)
    print(f'Выручка год:    {annual_rev:>10,} руб.')
    print(f'Расходы год:    {annual_exp:>10,} руб.')
    print(f'Прибыль год:    {annual_profit:>10,} руб.')
    print(f'ЧДП:            {annual_cashflow:>10,} руб.')
    print(f'Рентабельность: {rentab:>9.1f}%')
    print(f'Окупаемость:    {payback:>9} мес.')
    print(f'Год 2: {annual_rev_y2:,} / {annual_exp_y2:,} / {annual_profit_y2:,} / {rent_y2:.0f}%')

if __name__ == '__main__':
    main()
