#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Реконструкция БП Воробьева П.А. под жёсткие критерии Рязанской комиссии.
Работает напрямую с document.xml (без python-docx, чтобы не зависеть от PyPI).
Вход:  Бизнес план Воробьев Павел Алексеевич.docx (исходник в репо)
Выход: Бизнес план Воробьев Павел Алексеевич FIX.docx
"""
from __future__ import annotations
import os, shutil, zipfile, copy, re
import xml.etree.ElementTree as ET

# -----------------------------------------------------------------------------
# КОНСТАНТЫ И ИСХОДНЫЕ ДАННЫЕ
# -----------------------------------------------------------------------------
SRC = '/projects/sandbox/Bpp/Бизнес план Воробьев Павел Алексеевич.docx'
DST = '/projects/sandbox/Bpp/Бизнес план Воробьев Павел Алексеевич FIX.docx'
WORK = '/tmp/vor_build'

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = '{' + W + '}'
ET.register_namespace('w', W)
# Все namespaces из document.xml — копируем при распаковке
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

# -----------------------------------------------------------------------------
# ФИНМОДЕЛЬ — фиксированные значения
# -----------------------------------------------------------------------------
MONTHS = ['май','июн','июл','авг','сен','окт','ноя','дек','янв','фев','мар','апр']

# 5 услуг прайса
SERVICES = [
    ('Стандартный монтаж сплит-системы под ключ (настенного типа 07–09)', '3 часа',   10000),
    ('Двухэтапная установка кондиционера с закладкой трассы в новостройке', '5 часов',  14000),
    ('Комплексное сезонное техническое обслуживание (чистка, мойка, дозаправка)', '1.5 часа', 5000),
    ('Заправка фреоном / устранение утечки', '1 час', 3500),
    ('Демонтаж и сезонная консервация кондиционера', '1.5 часа', 3000),
]

# Помесячные количества заказов по услугам (5 строк)
# 0=май, 11=апр
QTY = {
    'std': [0, 9, 11, 8, 5, 2, 1, 0, 0, 0, 4, 5],   # 10000
    'two': [0, 1, 1, 2, 2, 2, 2, 2, 1, 1, 2, 2],   # 14000
    'to':  [0, 2, 2, 2, 2, 2, 1, 1, 1, 1, 3, 3],   # 5000
    'fr':  [0, 1, 1, 1, 1, 1, 1, 1, 0, 1, 2, 1],   # 3500
    'dem': [0, 0, 0, 0, 1, 2, 2, 1, 1, 1, 1, 3],   # 3000
}
PRICES = {'std': 10000, 'two': 14000, 'to': 5000, 'fr': 3500, 'dem': 3000}
SVC_KEYS = ['std','two','to','fr','dem']

# Расходы помесячно (рассчитаны заранее, см. финмодель)
# Реклама плавающая / Непредвиденные плавающие / Телефон стабильно 650
REC =  [5000, 1500, 1500, 1800, 2200, 3000, 3200, 2500, 2000, 3500, 4500, 5000]
NEPR = [   0, 1200, 1800, 1500, 1200,  800,  800,  600,  500,  800, 1500, 1800]
PHONE = 650

# Смета: добавляется новая позиция 4.1 (Авито Турбо 5000 ₽), грант 350 000 + свои 19 896 = 369 896
EQUIP_TOTAL_FIRST_MONTH = 369896  # Все 1.1-1.15 + 4.1 расход вытаскивается из 5.2 строки 2.1

# -----------------------------------------------------------------------------
# УТИЛИТЫ
# -----------------------------------------------------------------------------
def fmt_rub(n: int) -> str:
    """1234567 -> '1 234 567'"""
    s = f'{n:,}'.replace(',', ' ')
    return s

def calc_revenue_month(i: int) -> int:
    return sum(QTY[k][i] * PRICES[k] for k in SVC_KEYS)

def calc_orders_month(i: int) -> int:
    return sum(QTY[k][i] for k in SVC_KEYS)

def calc_materials_month(i: int) -> int:
    """Расходники: стд 500, 2эт 1500, ТО 200, заправка 200, демонтаж 200; округлено до сотен."""
    raw = QTY['std'][i]*500 + QTY['two'][i]*1500 + QTY['to'][i]*200 + QTY['fr'][i]*200 + QTY['dem'][i]*200
    if raw == 0: return 0
    return int(round(raw / 100.0)) * 100

def calc_transport_month(i: int) -> int:
    """Транспорт: 1-й мес = 0 (орг.период), далее 1500 база + 150*заказ; округление до сотен."""
    if i == 0:
        return 0  # организационный месяц, выездов нет
    n = calc_orders_month(i)
    raw = 1500 + 150 * n
    return int(round(raw / 100.0)) * 100

def calc_npd_month(i: int) -> int:
    """НПД 4% от выручки, округление до 10."""
    rev = calc_revenue_month(i)
    return int(round(rev * 0.04 / 10.0)) * 10

def calc_expenses_breakdown(i: int) -> dict:
    osn = EQUIP_TOTAL_FIRST_MONTH if i == 0 else 0
    mat = calc_materials_month(i)
    arenda = 0
    trans = calc_transport_month(i)
    phone = 0 if i == 0 else PHONE  # 1-й мес организационный, телефон с 2-го
    rec = REC[i]
    nepr = NEPR[i]
    npd = calc_npd_month(i)
    total = osn + mat + arenda + trans + phone + rec + nepr + npd
    return {
        'osn': osn, 'mat': mat, 'arenda': arenda, 'trans': trans,
        'phone': phone, 'rec': rec, 'nepr': nepr, 'npd': npd, 'total': total
    }

# -----------------------------------------------------------------------------
# DOCX I/O
# -----------------------------------------------------------------------------
def unpack_docx(src: str, work_dir: str):
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)
    with zipfile.ZipFile(src, 'r') as z:
        z.extractall(work_dir)

def repack_docx(work_dir: str, dst: str):
    if os.path.exists(dst):
        os.remove(dst)
    # Word требует чтобы [Content_Types].xml шёл первым и был uncompressed (необязательно, но безопаснее)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(work_dir):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, work_dir).replace(os.sep, '/')
                z.write(full, rel)

# -----------------------------------------------------------------------------
# XML helpers
# -----------------------------------------------------------------------------
def cell_text(cell) -> str:
    return ''.join((t.text or '') for t in cell.iter(NS+'t'))

def get_cell_paragraphs(cell):
    return cell.findall(NS+'p')

def first_run_props(cell):
    """Достать первый <w:rPr> из ячейки (для сохранения шрифта при перезаписи)."""
    for r in cell.iter(NS+'r'):
        rpr = r.find(NS+'rPr')
        if rpr is not None:
            return rpr
    return None

def first_par_props(cell):
    p = cell.find(NS+'p')
    if p is None: return None
    return p.find(NS+'pPr')

def set_cell_text(cell, text: str, *, bold: bool|None=None, align: str|None=None):
    """Полностью заменить содержимое ячейки одним абзацем с указанным текстом.
    Сохраняем стиль шрифта (rPr) и абзаца (pPr) из первого исходного абзаца ячейки."""
    saved_rpr = first_run_props(cell)
    saved_ppr = first_par_props(cell)
    # удалить все <w:p> в ячейке
    for p in list(cell.findall(NS+'p')):
        cell.remove(p)
    # вставить новый <w:p>
    p = ET.SubElement(cell, NS+'p')
    if saved_ppr is not None:
        p.insert(0, copy.deepcopy(saved_ppr))
    else:
        # минимальный pPr
        pass
    if align:
        ppr = p.find(NS+'pPr')
        if ppr is None:
            ppr = ET.SubElement(p, NS+'pPr')
            p.insert(0, ppr)
        # удалить старый jc
        for old in ppr.findall(NS+'jc'):
            ppr.remove(old)
        jc = ET.SubElement(ppr, NS+'jc')
        jc.set(NS+'val', align)
    r = ET.SubElement(p, NS+'r')
    if saved_rpr is not None:
        r.append(copy.deepcopy(saved_rpr))
    if bold is not None:
        rpr = r.find(NS+'rPr')
        if rpr is None:
            rpr = ET.SubElement(r, NS+'rPr')
            r.insert(0, rpr)
        for old in rpr.findall(NS+'b'):
            rpr.remove(old)
        if bold:
            b = ET.SubElement(rpr, NS+'b')
    t = ET.SubElement(r, NS+'t')
    t.text = text
    # preserve spaces в начале/конце
    if text != text.strip() or '  ' in text:
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return p

def set_cell_lines(cell, lines: list[str], *, bold: bool|None=None, align: str|None=None):
    """Несколько абзацев в ячейке."""
    saved_rpr = first_run_props(cell)
    saved_ppr = first_par_props(cell)
    for p in list(cell.findall(NS+'p')):
        cell.remove(p)
    for line in lines:
        p = ET.SubElement(cell, NS+'p')
        if saved_ppr is not None:
            p.insert(0, copy.deepcopy(saved_ppr))
        if align:
            ppr = p.find(NS+'pPr')
            if ppr is None:
                ppr = ET.SubElement(p, NS+'pPr'); p.insert(0, ppr)
            for old in ppr.findall(NS+'jc'): ppr.remove(old)
            jc = ET.SubElement(ppr, NS+'jc'); jc.set(NS+'val', align)
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
        t.text = line
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

def get_tables(root):
    body = root.find(NS+'body')
    return body.findall('.//' + NS + 'tbl')

def get_paragraphs(root):
    body = root.find(NS+'body')
    return list(body.findall(NS+'p'))

def para_text(p) -> str:
    return ''.join((t.text or '') for t in p.iter(NS+'t'))

def replace_paragraph_text(p, new_text: str):
    """Заменить весь текст параграфа одной строкой, сохранив pPr и rPr первого run."""
    saved_ppr = p.find(NS+'pPr')
    saved_rpr = None
    for r in p.findall(NS+'r'):
        rpr = r.find(NS+'rPr')
        if rpr is not None:
            saved_rpr = copy.deepcopy(rpr); break
    # удалить все <w:r>, <w:hyperlink>, <w:fldSimple>, <w:smartTag> и пр. оставить только pPr
    for child in list(p):
        if child.tag != NS+'pPr':
            p.remove(child)
    r = ET.SubElement(p, NS+'r')
    if saved_rpr is not None:
        r.append(saved_rpr)
    t = ET.SubElement(r, NS+'t')
    t.text = new_text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

def replace_paragraph_text_multiline(p, lines: list[str]):
    """Заменить текст параграфа на несколько строк (через мягкие переносы <w:br/>)."""
    saved_ppr = p.find(NS+'pPr')
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
    for idx, line in enumerate(lines):
        if idx > 0:
            ET.SubElement(r, NS+'br')
        t = ET.SubElement(r, NS+'t')
        t.text = line
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

def find_para_by_substr(paragraphs, needle: str):
    """Вернуть индекс первого параграфа содержащего подстроку (без учёта переносов в runs)."""
    for i, p in enumerate(paragraphs):
        if needle in para_text(p):
            return i
    return -1

def insert_paragraph_after(parent, ref_p, text: str):
    """Вставить новый <w:p> после ref_p, копируя pPr и rPr из ref_p."""
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
    # вставить после ref_p
    children = list(parent)
    idx = children.index(ref_p)
    parent.insert(idx + 1, new_p)
    return new_p

def clone_row(row):
    return copy.deepcopy(row)

# -----------------------------------------------------------------------------
# ОСНОВНАЯ ЛОГИКА
# -----------------------------------------------------------------------------
def main():
    print(f'>> Распаковка {SRC} -> {WORK}')
    unpack_docx(SRC, WORK)

    doc_xml_path = os.path.join(WORK, 'word', 'document.xml')
    tree = ET.parse(doc_xml_path)
    root = tree.getroot()
    body = root.find(NS+'body')

    tables = get_tables(root)
    print(f'>> Найдено таблиц: {len(tables)}')
    assert len(tables) == 8, 'Ожидалось 8 таблиц в исходном БП'

    # =========================================================================
    # ТАБЛИЦА #0: 1.16 Показатели резюме
    # =========================================================================
    print('>> [#0] Раздел 1.16 — пересчёт показателей резюме')
    annual_rev = sum(calc_revenue_month(i) for i in range(12))
    annual_exp = sum(calc_expenses_breakdown(i)['total'] for i in range(12))
    annual_profit = annual_rev - annual_exp
    annual_cashflow = annual_profit + 350000 + 19896  # + грант + свои
    rentab = annual_profit / annual_exp * 100

    # Найти месяц где накоп. прибыль перешла в плюс
    cum = 0; payback = 0
    for i in range(12):
        cum += calc_revenue_month(i) - calc_expenses_breakdown(i)['total']
        if cum > 0:
            payback = i + 1
            break

    print(f'   Выручка год: {annual_rev:,} | Расходы: {annual_exp:,} | Прибыль: {annual_profit:,}')
    print(f'   ЧДП: {annual_cashflow:,} | Рент: {rentab:.1f}% | Окупаемость: {payback} мес.')

    t0 = tables[0]
    rows0 = t0.findall(NS+'tr')
    # rows0[0] - заголовок; 1..6 - данные
    set_cell_text(rows0[1].findall(NS+'tc')[1], f'{fmt_rub(annual_rev)} руб.')
    set_cell_text(rows0[2].findall(NS+'tc')[1], f'{fmt_rub(annual_exp)} руб.')
    set_cell_text(rows0[3].findall(NS+'tc')[1], f'{fmt_rub(annual_profit)} руб.')
    set_cell_text(rows0[4].findall(NS+'tc')[1], f'{fmt_rub(annual_cashflow)} руб.')
    set_cell_text(rows0[5].findall(NS+'tc')[1], f'{rentab:.1f}%'.replace('.', ','))
    set_cell_text(rows0[6].findall(NS+'tc')[1], f'{payback} месяцев')

    # =========================================================================
    # ТАБЛИЦА #1: 2.1.1 Прайс — 5 услуг (расширяем с 3 до 5 строк)
    # =========================================================================
    print('>> [#1] Прайс 2.1.1 — 5 услуг')
    t1 = tables[1]
    rows1 = t1.findall(NS+'tr')
    # rows1[0] — заголовок, 1..3 — 3 услуги (в исходнике)
    template_row = clone_row(rows1[1])  # шаблон строки услуги
    # Удалить старые строки данных
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
    # ТАБЛИЦА #2: 4.1 Источники финансирования — пересчёт ИТОГО
    # =========================================================================
    print('>> [#2] 4.1 Источники финансирования')
    t2 = tables[2]
    rows2 = t2.findall(NS+'tr')
    # rows2[0] заголовок, 1=грант, 2=свои, 3=заёмные, 4=иное, 5=ИТОГО
    set_cell_text(rows2[1].findall(NS+'tc')[2], '350 000')
    set_cell_text(rows2[2].findall(NS+'tc')[2], '19 896')
    set_cell_text(rows2[3].findall(NS+'tc')[2], '0')
    set_cell_text(rows2[4].findall(NS+'tc')[2], '0')
    # ИТОГО ячейка
    set_cell_text(rows2[5].findall(NS+'tc')[-1], '369 896')

    # =========================================================================
    # ТАБЛИЦА #3: 4.2 Смета — добавляем 4.1 Авито Турбо, расставляем фонды
    # =========================================================================
    print('>> [#3] 4.2 Смета — переходная позиция 1.13, добавление п.4.1')
    t3 = tables[3]
    rows3 = t3.findall(NS+'tr')
    # rows3[0] заголовок; rows3[1] = "1.* Основные средства"; rows3[2..16] = 1.1..1.15
    # rows3[17] = "2.* Запасы"; rows3[18] = "3.* Аренда"; rows3[19] = "4.* Продвижение";
    # rows3[20] = "5.* Документация"; rows3[21] = "ИТОГО"

    # Поправить позицию 1.13 (Фреон) — переходная: 13 189 грант + 431 свои
    cells_113 = rows3[14].findall(NS+'tc')
    set_cell_text(cells_113[6], 'Средства соцконтракта (13 189 ₽), собственные средства (431 ₽)')

    # 1.14 и 1.15 — собственные средства (это уже так в исходнике, но переподтвердим)
    set_cell_text(rows3[15].findall(NS+'tc')[6], 'Собственные средства')
    set_cell_text(rows3[16].findall(NS+'tc')[6], 'Собственные средства')

    # Под раздел "4.* Продвижение" (rows3[19]) добавить новую строку 4.1
    section4_row = rows3[19]  # сама строка-заголовок раздела 4
    # Создать новую строку как клон 1.1 (rows3[2]) с правильным форматированием
    new_row_4_1 = clone_row(rows3[2])
    new_cells = new_row_4_1.findall(NS+'tc')
    set_cell_text(new_cells[0], ' 4.1')
    set_cell_text(new_cells[1], 'Стартовый рекламный пакет Авито Pro (Турбо-продвижение, 1 мес.)')
    set_cell_text(new_cells[2], '1')
    set_cell_text(new_cells[3], '5000')
    set_cell_text(new_cells[4], '5000')
    set_cell_text(new_cells[5], 'Avito.ru')
    set_cell_text(new_cells[6], 'Собственные средства')
    # Вставить после section4_row (т.е. перед rows3[20])
    children = list(t3)
    idx = children.index(section4_row)
    t3.insert(idx + 1, new_row_4_1)

    # ИТОГО — последняя строка. Посмотрим её ячейки:
    rows3 = t3.findall(NS+'tr')  # перечитываем после вставки
    itogo_row = rows3[-1]
    itogo_cells = itogo_row.findall(NS+'tc')
    # 6 ячеек: ИТОГО / / / / / (Приложение 2)
    # Обычно ИТОГО справа от названия. Поставим сумму в предпоследнюю или последнюю.
    # Структура: [ИТОГО][пусто][пусто][пусто][сумма?][Приложение 2]
    # Запишем сумму в ту ячейку, где должна стоять цифра (4-я по счёту, индекс 4 если 6 ячеек)
    # Из дампа: row 21 (6 ячеек): ['ИТОГО', '', '', '', '', '']  — все пустые кроме первой
    # Поставим сумму в индекс 4 (под колонкой "Сумма")
    if len(itogo_cells) >= 5:
        set_cell_text(itogo_cells[4], '369 896', bold=True)

    # =========================================================================
    # ТАБЛИЦА #4: 5.1 Доходы — заполнение по месяцам
    # =========================================================================
    print('>> [#4] 5.1 Доходы — заполнение помесячно')
    t4 = tables[4]
    rows4 = t4.findall(NS+'tr')
    # rows4[0] = заголовок "1 мес. ... 12 мес. ИТОГО Доля,%"
    # rows4[1] = "май июн ... апр"
    # rows4[2] = "Доходы всего, руб., в т.ч.:"
    # rows4[3] = "1." (название услуги в [1])
    # rows4[4] = "количество"
    # rows4[5] = "стоимость"
    # rows4[6..17] = "2." "количество" "стоимость" — итого 5 услуг по 3 строки

    # Строка 2: суммарные доходы по всем услугам помесячно
    row_total = rows4[2]
    cells = row_total.findall(NS+'tc')
    # cells[0]=№, cells[1]=Доходы всего..., cells[2..13]=12 месяцев, cells[14]=ИТОГО, cells[15]=Доля%
    annual = 0
    for i in range(12):
        rev_i = calc_revenue_month(i)
        annual += rev_i
        set_cell_text(cells[2 + i], fmt_rub(rev_i) if rev_i else '0')
    set_cell_text(cells[14], fmt_rub(annual), bold=True)
    set_cell_text(cells[15], '100')

    # Услуги 1..5 — по 3 строки на услугу
    # СТРУКТУРА (по требованию Рязанской комиссии):
    #   Строка А (name_row): [название услуги] + в ячейках месяцев — ВЫРУЧКА за услугу (Q*P)
    #   Строка Б (qty_row):  "количество"     + число заказов в месяц (рваное)
    #   Строка В (sum_row):  "стоимость"      + ЦЕНА ЗА ЕДИНИЦУ из прайса (фиксированная константа)
    # ИТОГО:
    #   - для name_row    = сумма выручки за 12 мес.  (жирным)
    #   - для qty_row     = сумма заказов за 12 мес.
    #   - для sum_row     = ПУСТО (это константа цены, не имеет ИТОГО)
    # Доля % — только в name_row (выручка услуги / годовая выручка)
    SVC_ROW_LAYOUT = [
        # (key, idx_name_row)
        ('std', 3), ('two', 6), ('to', 9), ('fr', 12), ('dem', 15),
    ]
    for key, idx in SVC_ROW_LAYOUT:
        name_row = rows4[idx]      # строка с № и названием — сюда же выручка по месяцам
        qty_row = rows4[idx + 1]   # "количество"
        sum_row = rows4[idx + 2]   # "стоимость" — цена за единицу
        svc_idx = SVC_KEYS.index(key)
        svc_name = SERVICES[svc_idx][0]
        price = PRICES[key]

        nrcells = name_row.findall(NS+'tc')
        qcells = qty_row.findall(NS+'tc')
        scells = sum_row.findall(NS+'tc')

        # Имя услуги — в name_row, ячейка [1]
        set_cell_text(nrcells[1], svc_name)

        total_qty = 0
        total_sum = 0
        for i in range(12):
            q = QTY[key][i]
            revenue_i = q * price
            total_qty += q
            total_sum += revenue_i
            # А) выручка за услугу в месяце — в name_row
            set_cell_text(nrcells[2 + i], fmt_rub(revenue_i) if revenue_i else '0')
            # Б) количество заказов
            set_cell_text(qcells[2 + i], str(q) if q else '0')
            # В) стоимость = цена за единицу (фиксированная константа)
            set_cell_text(scells[2 + i], fmt_rub(price))

        # ИТОГО:
        set_cell_text(nrcells[14], fmt_rub(total_sum), bold=True)  # годовая выручка услуги
        set_cell_text(qcells[14], str(total_qty))                  # годовое кол-во заказов
        set_cell_text(scells[14], '')                              # стоимость — пусто

        # Доля % — только в name_row (доля выручки этой услуги от всей выручки)
        share = total_sum / annual * 100 if annual else 0
        set_cell_text(nrcells[15], f'{share:.1f}'.replace('.', ','))
        set_cell_text(qcells[15], '')
        set_cell_text(scells[15], '')

    # =========================================================================
    # ТАБЛИЦА #5: 5.2 Расходы — заполнение полностью
    # =========================================================================
    print('>> [#5] 5.2 Расходы — заполнение помесячно')
    t5 = tables[5]
    rows5 = t5.findall(NS+'tr')
    # 0=заголовок, 1=месяцы, 2=Доходы, 3=Расходы итого, 4..11 = 2.1..2.8,
    # 12=Прибыль, 13=ЧДП, 14=Нараст, 15=Рент%, 16=Стороннее, 17=Свои

    def fill_row(row_idx: int, monthly_values: list[int], total: int|None=None,
                 bold_total: bool=False, fmt=fmt_rub, label_col_filled=False):
        row = rows5[row_idx]
        cells = row.findall(NS+'tc')
        for i in range(12):
            v = monthly_values[i]
            if v == 0:
                set_cell_text(cells[2 + i], '0')
            else:
                set_cell_text(cells[2 + i], fmt(v) if isinstance(v, int) else str(v))
        if total is None:
            total = sum(monthly_values)
        set_cell_text(cells[14], fmt(total) if isinstance(total, int) else str(total), bold=bold_total)

    inc_arr = [calc_revenue_month(i) for i in range(12)]
    exp_breakdown = [calc_expenses_breakdown(i) for i in range(12)]

    # 1. Доходы
    fill_row(2, inc_arr, bold_total=True)

    # 2.1 Оборудование
    fill_row(4,  [b['osn']    for b in exp_breakdown])
    # 2.2 Расходные материалы
    fill_row(5,  [b['mat']    for b in exp_breakdown])
    # 2.3 Аренда
    fill_row(6,  [b['arenda'] for b in exp_breakdown])
    # 2.4 Транспорт
    fill_row(7,  [b['trans']  for b in exp_breakdown])
    # 2.5 Телефон
    fill_row(8,  [b['phone']  for b in exp_breakdown])
    # 2.6 Реклама
    fill_row(9,  [b['rec']    for b in exp_breakdown])
    # 2.7 Непредвиденные
    fill_row(10, [b['nepr']   for b in exp_breakdown])
    # 2.8 НПД
    fill_row(11, [b['npd']    for b in exp_breakdown])

    # 2 Расходы итого (сумма всех 2.1..2.8)
    exp_arr = [b['total'] for b in exp_breakdown]
    fill_row(3, exp_arr, bold_total=True)

    # 3 Прибыль
    profit_arr = [inc_arr[i] - exp_arr[i] for i in range(12)]
    fill_row(12, profit_arr, bold_total=True)

    # 4 Чистый денежный поток (стр.3 + грант_месяц + свои_месяц)
    # В мае приток = 350000 + 19896 = 369896, в остальных = 0
    grant_arr = [350000] + [0]*11
    own_arr = [19896] + [0]*11
    cf_arr = [profit_arr[i] + grant_arr[i] + own_arr[i] for i in range(12)]
    fill_row(13, cf_arr, bold_total=True)

    # 4.1 Чистый поток нарастающим итогом
    cum = 0
    cum_arr = []
    for v in cf_arr:
        cum += v
        cum_arr.append(cum)
    # последняя ячейка нарастающего = итоговая = cum
    row14 = rows5[14]; cells14 = row14.findall(NS+'tc')
    for i in range(12):
        set_cell_text(cells14[2 + i], fmt_rub(cum_arr[i]))
    set_cell_text(cells14[14], fmt_rub(cum_arr[-1]), bold=True)

    # 5 Рентабельность — у Куксова/Сизова помесячная рентабельность = ПУСТО, только ИТОГО годовой
    row15 = rows5[15]; cells15 = row15.findall(NS+'tc')
    for i in range(12):
        set_cell_text(cells15[2 + i], '')  # пусто по месяцам (как у Куксова)
    annual_rent = sum(profit_arr) / sum(exp_arr) * 100
    set_cell_text(cells15[14], f'{annual_rent:.1f}%'.replace('.', ','), bold=True)

    # 6 Стороннее финансирование (грант)
    fill_row(16, grant_arr)

    # 7 Собственные средства
    fill_row(17, own_arr)

    # =========================================================================
    # ТАБЛИЦА #6: 5.4 Год 2
    # =========================================================================
    print('>> [#6] 5.4 Показатели за 2 года')
    # Год 2 без оборудования + резерв 80 000 на модернизацию
    annual_rev_y2 = 950000  # +8.4% к году 1
    # Расходы год 2: всё кроме 2.1 + 80k резерв на новое оборудование
    annual_exp_y2 = sum(exp_breakdown[i]['total'] - exp_breakdown[i]['osn'] for i in range(12)) + 80000
    # +10% к расходникам/транспорту за рост заказов
    annual_exp_y2 = int(annual_exp_y2 * 1.05)
    annual_profit_y2 = annual_rev_y2 - annual_exp_y2
    rent_y2 = annual_profit_y2 / annual_exp_y2 * 100

    print(f'   Год 2: Доход {annual_rev_y2:,} | Расход {annual_exp_y2:,} | Прибыль {annual_profit_y2:,} | Рент {rent_y2:.0f}%')

    t6 = tables[6]
    rows6 = t6.findall(NS+'tr')
    # 0=заголовок, 1=Доходы, 2=Расходы, 3=Прибыль, 4=Рент
    set_cell_text(rows6[1].findall(NS+'tc')[2], f'{fmt_rub(annual_rev)} руб.')
    set_cell_text(rows6[1].findall(NS+'tc')[3], f'{fmt_rub(annual_rev_y2)} руб.')
    set_cell_text(rows6[2].findall(NS+'tc')[2], f'{fmt_rub(annual_exp)} руб.')
    set_cell_text(rows6[2].findall(NS+'tc')[3], f'{fmt_rub(annual_exp_y2)} руб.')
    set_cell_text(rows6[3].findall(NS+'tc')[2], f'{fmt_rub(annual_profit)} руб.')
    set_cell_text(rows6[3].findall(NS+'tc')[3], f'{fmt_rub(annual_profit_y2)} руб.')
    set_cell_text(rows6[4].findall(NS+'tc')[2], f'{rentab:.1f}%'.replace('.', ','))
    set_cell_text(rows6[4].findall(NS+'tc')[3], f'{rent_y2:.0f}%')

    # =========================================================================
    # ТАБЛИЦА #7: Риски — короткие решения по 1 предложению
    # =========================================================================
    print('>> [#7] Риски — короткие решения')
    RISKS = [
        ('Снижение платежеспособности потребителей',
         'средняя',
         'Поддерживать пакет «эконом-монтаж» с базовой комплектацией.'),
        ('Эпидемиологическая обстановка. Риск заболеваемости',
         'низкая',
         'Соблюдать СИЗ при работе на территории заказчика.'),
        ('Уменьшение потока клиентов (сезонность)',
         'высокая',
         'Зимой переключаться на закладку трасс в новостройках.'),
        ('Несоответствие ожиданиям клиента',
         'средняя',
         'Заключать договор с фотоотчётом по каждому этапу работ.'),
        ('Болезнь',
         'низкая',
         'Вести здоровый образ жизни и поддерживать резервный фонд.'),
        ('Появление конкурентов',
         'средняя',
         'Удерживать рейтинг 4,9+ на Авито за счёт качества и гарантии.'),
        ('Поломка оборудования',
         'средняя',
         'Иметь резервный комплект ключевого инструмента.'),
    ]
    t7 = tables[7]
    rows7 = t7.findall(NS+'tr')
    for i, (risk, prob, sol) in enumerate(RISKS):
        row = rows7[i + 1]
        cells = row.findall(NS+'tc')
        set_cell_text(cells[0], risk)
        set_cell_text(cells[1], prob.capitalize())
        set_cell_text(cells[2], sol)

    # =========================================================================
    # ПАРАГРАФЫ — текстовые правки
    # =========================================================================
    print('>> Текстовые правки в параграфах')
    paragraphs = get_paragraphs(root)

    # 2.3. Конкуренция — заменить весь параграф "Крупные фирмы — дорого..."
    # Сначала заменим параграф с "Крупные фирмы" на новый длинный текст из нескольких параграфов
    NEW_CONKURENT_LINES = [
        '2.3. Конкурентная среда, конкурентная позиция бизнеса.',
        'Анализ объявлений на Авито в категории «Монтаж и обслуживание кондиционеров» по г. Рязани (выборка 146 активных предложений на май 2026 г.) показал, что рынок работает в трёх ценовых сегментах.',
        'Климатические компании и магазины с собственными бригадами ставят сплит-систему за 12 500–18 000 руб. Это премиальный сегмент: гарантия от юридического лица, рассрочка, выезд по записи. Слабая сторона — высокая итоговая стоимость работ за счёт офисных накладных и очередь до 14–20 дней в пиковые месяцы (июнь–август), что приводит к потере «горячего» клиента.',
        'Частники без оборудования занимают около 60% объявлений Авито, цена монтажа 3 000–7 000 руб. Слабая сторона — отсутствие профессионального инструмента (продувка контура фреоном из баллона вместо вакуумирования), отсутствие алмазного бурения с пылеудалением, отсутствие официального чека и письменной гарантии. По данным форумов производителей, срок безотказной службы кондиционеров после такого монтажа сокращается на 30–40%.',
        'Частные мастера-самозанятые с полным комплектом профессионального инструмента — целевой сегмент проекта. Цена монтажа в этом сегменте — 9 000–12 000 руб., медиана 10 000 руб.',
        'Конкурентная позиция проекта. Цена 10 000 руб. за стандартный монтаж — медиана профессионального сегмента. Ключевые преимущества: профессиональный вакуумный насос Value VRP-8DV и цифровой коллектор Testo 550s с фотоотчётом клиенту, алмазное бурение Bosch с пылеудалением, официальный чек самозанятого, договор на каждую работу, письменная гарантия 12 месяцев на монтажные работы. Это позволяет занять позицию «премиум за цену середины рынка» без перехода в дорогой сегмент. Полный сравнительный анализ цен 6 ключевых конкурентов — в Приложении 3.'
    ]

    # Найти параграф с "Крупные фирмы — дорого"
    idx_konk = find_para_by_substr(paragraphs, 'Крупные фирмы')
    if idx_konk == -1:
        # альтернативный поиск
        idx_konk = find_para_by_substr(paragraphs, 'без проф. инструмента')
    if idx_konk >= 0:
        target_p = paragraphs[idx_konk]
        # Первый абзац — заголовок (жирный, как и был в оригинале)
        replace_paragraph_text(target_p, NEW_CONKURENT_LINES[0])
        # Принудительно сделать жирным первый run заголовка
        for r in target_p.findall(NS+'r'):
            rpr = r.find(NS+'rPr')
            if rpr is None:
                rpr = ET.SubElement(r, NS+'rPr'); r.insert(0, rpr)
            for old_b in rpr.findall(NS+'b'): rpr.remove(old_b)
            ET.SubElement(rpr, NS+'b')
            break
        # Затем вставить остальные параграфы после
        prev = target_p
        for line in NEW_CONKURENT_LINES[1:]:
            prev = insert_paragraph_after(body, prev, line)
        print(f'   2.3 Конкуренция — заменён параграф #{idx_konk}, добавлено {len(NEW_CONKURENT_LINES)-1} абзацев')

    # Перечитать параграфы (структура изменилась)
    paragraphs = get_paragraphs(root)

    # 2.4 + 2.8 → перенумеровать в 2.4 и 2.5 (вместо 2.4 и 2.8)
    # Параграф начинающийся с "2.8." → заменить на "2.5."
    for p in paragraphs:
        txt = para_text(p)
        if txt.startswith('2.8.') or txt.lstrip().startswith('2.8.'):
            # Найти первый <w:t> и заменить начало
            for t in p.iter(NS+'t'):
                if t.text and '2.8' in t.text:
                    t.text = t.text.replace('2.8.', '2.5.', 1)
                    print(f'   2.8 → 2.5 (нумерация)')
                    break
            break

    # 5.1 — текст после таблицы доходов (про сезонность)
    # Найдём параграф "Выручка в первый месяц (май)" и заменим целиком
    NEW_5_1_TEXT = (
        'Выручка в первый месяц (май) отсутствует — этот период полностью посвящён закупке оборудования, '
        'оформлению самозанятости, размещению объявлений на Авито и формированию портфолио. '
        'Со второго месяца начинается активная работа в высокий сезон: в июне–августе спрос на стандартный монтаж сплит-систем '
        'максимален (8–11 заказов в месяц). С сентября по ноябрь идёт плавный спад; структура заказов смещается '
        'на закладку трасс в новостройках. Январь — самый низкий месяц после праздников (3 заказа), '
        'но кассовых разрывов нет: накопленный за лето чистый поток обеспечивает запас. '
        'С марта начинается второй пик (ТО, демонтаж) — клиенты заказывают заранее.'
    )

    idx_5_1 = find_para_by_substr(paragraphs, 'Выручка в первый месяц')
    if idx_5_1 >= 0:
        replace_paragraph_text(paragraphs[idx_5_1], NEW_5_1_TEXT)
        print('   5.1 текст обновлён')

    # 5.2 — текст после таблицы расходов
    NEW_5_2_TEXT = (
        'В первый месяц (май) расходы составляют ' + fmt_rub(exp_breakdown[0]['total']) + ' руб. — это закупка всего необходимого инструмента '
        'и стартовый рекламный пакет Авито. Расходы первого месяца полностью покрываются средствами социального контракта '
        '(350 000 руб.) и собственными вложениями автора (19 896 руб.), кассового разрыва нет. '
        'Со второго месяца расходы носят преимущественно динамический характер: расходные материалы (медные трубы, изоляция, '
        'химия для мойки) и транспортные расходы растут пропорционально количеству заказов, налог НПД 4% уплачивается по факту '
        'поступления выручки. Реклама плавающая по сезону: в летние пиковые месяцы достаточно базовой подписки на Авито (1 500 руб.), '
        'осенью и в предсезонье объём увеличивается до 3 500–5 000 руб.'
    )

    idx_5_2 = find_para_by_substr(paragraphs, 'В первый месяц расходы составляют')
    if idx_5_2 >= 0:
        replace_paragraph_text(paragraphs[idx_5_2], NEW_5_2_TEXT)
        print('   5.2 текст обновлён')

    # 5.3 Срок окупаемости — заменить "4 месяца" на актуальный
    idx_5_3 = find_para_by_substr(paragraphs, 'Срок окупаемости')
    if idx_5_3 >= 0:
        # Целиком заменить
        replace_paragraph_text(paragraphs[idx_5_3], f'5.3. Срок окупаемости проекта: {payback} мес.')
        print(f'   5.3 окупаемость → {payback} мес.')

    # =========================================================================
    # СОХРАНЕНИЕ
    # =========================================================================
    print('>> Сохранение document.xml')
    tree.write(doc_xml_path, xml_declaration=True, encoding='UTF-8', method='xml')

    print(f'>> Упаковка в {DST}')
    repack_docx(WORK, DST)

    # Контроль
    sz = os.path.getsize(DST)
    print(f'>> Готово. Размер файла: {sz} байт')

    # Финальная сводка
    print('\n' + '='*70)
    print('ИТОГОВАЯ ФИНМОДЕЛЬ')
    print('='*70)
    print(f'Выручка год:       {annual_rev:>10,} руб.')
    print(f'Расходы год:       {annual_exp:>10,} руб.')
    print(f'Прибыль год:       {annual_profit:>10,} руб.')
    print(f'Чистый поток:      {annual_cashflow:>10,} руб.')
    print(f'Рентабельность:    {rentab:>9.1f}%')
    print(f'Окупаемость:       {payback:>9} мес.')
    print('-'*70)
    print(f'Год 2: доход {annual_rev_y2:,} / расход {annual_exp_y2:,} / прибыль {annual_profit_y2:,} / рент {rent_y2:.0f}%')

if __name__ == '__main__':
    main()
