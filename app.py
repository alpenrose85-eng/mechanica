import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

# Настройка страницы
st.set_page_config(
    page_title="Обработчик протоколов механических испытаний",
    page_icon="📊",
    layout="wide"
)

# Стили
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 2rem;
    }
    .info-box {
        background-color: #f0f7ff;
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #1E3A8A;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def clean_number(text):
    """Очистка и преобразование чисел из текста"""
    if not text:
        return 0
    
    # Убираем пробелы в числах (например, "3 363" -> "3363")
    text = str(text).replace(' ', '')
    
    # Заменяем запятые на точки для десятичных чисел
    text = text.replace(',', '.')
    
    # Убираем все нечисловые символы, кроме точки и цифр
    text = re.sub(r'[^\d.]', '', text)
    
    try:
        return float(text) if '.' in text else int(text)
    except:
        return 0

def parse_protocol_from_docx(file_content):
    """Парсинг данных из DOCX файла с протоколом"""
    doc = Document(BytesIO(file_content))
    
    data_rows = []
    
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            
            # Ищем строку с клеймом образца
            for i, cell_text in enumerate(cells):
                if re.match(r'^\d+-\d+$', cell_text):
                    try:
                        sample_mark = cell_text
                        
                        # Извлекаем данные из строки
                        # В таблице 14 колонок, данные находятся на определенных позициях
                        if len(cells) >= 14:
                            # Температура - 5-я колонка (индекс 5 в 0-based)
                            temp_text = cells[5]
                            temp_match = re.search(r'(\d+)', temp_text)
                            temperature = int(temp_match.group(1)) if temp_match else 20
                            
                            # Предел прочности - 10-я колонка (индекс 10)
                            strength_text = cells[10]
                            strength = clean_number(strength_text)
                            
                            # Предел текучести - 11-я колонка (индекс 11)
                            yield_text = cells[11]
                            yield_strength = clean_number(yield_text)
                            
                            # Относительное сужение - 12-я колонка (индекс 12)
                            reduction_text = cells[12]
                            reduction = clean_number(reduction_text)
                            
                            # Относительное удлинение - 13-я колонка (индекс 13)
                            elongation_text = cells[13]
                            elongation = clean_number(elongation_text)
                            
                            data_rows.append({
                                'Клеймо': sample_mark,
                                'Температура': temperature,
                                'Предел прочности': strength,
                                'Предел текучести': yield_strength,
                                'Отн. удл.': elongation,
                                'Отн. суж.': reduction
                            })
                            
                    except Exception as e:
                        continue
    
    # Если не нашли данные в таблицах, пробуем парсить текст
    if not data_rows:
        return parse_protocol_from_text('\n'.join([p.text for p in doc.paragraphs]))
    
    return pd.DataFrame(data_rows)

def parse_protocol_from_text(text):
    """Парсинг данных из текста протокола"""
    lines = text.split('\n')
    data_rows = []
    
    for line in lines:
        # Ищем строки с клеймом образца
        if re.search(r'\d+-\d+', line) and any(x in line for x in ['МПа', '485', '297', '57', '30']):
            # Убираем лишние пробелы
            line_clean = re.sub(r'\s+', ' ', line.strip())
            
            # Разбиваем строку на части
            parts = line_clean.split()
            
            # Ищем клеймо
            for i, part in enumerate(parts):
                if re.match(r'^\d+-\d+$', part):
                    try:
                        sample_mark = part
                        
                        # Ищем числовые значения после клейма
                        numbers = []
                        for j in range(i+1, len(parts)):
                            # Очищаем каждое значение
                            cleaned = clean_number(parts[j])
                            if cleaned != 0:
                                numbers.append(cleaned)
                        
                        # В таблице должно быть минимум 12 чисел после клейма
                        if len(numbers) >= 12:
                            # Температура - 3-е число после клейма
                            temperature = int(numbers[2]) if len(numbers) > 2 else 20
                            
                            # Предел прочности - 8-е число после клейма
                            strength = numbers[7] if len(numbers) > 7 else 0
                            
                            # Предел текучести - 9-е число после клейма
                            yield_strength = numbers[8] if len(numbers) > 8 else 0
                            
                            # Относительное сужение - 10-е число после клейма
                            reduction = numbers[9] if len(numbers) > 9 else 0
                            
                            # Относительное удлинение - 11-е число после клейма
                            elongation = numbers[10] if len(numbers) > 10 else 0
                            
                            data_rows.append({
                                'Клеймо': sample_mark,
                                'Температура': temperature,
                                'Предел прочности': strength,
                                'Предел текучести': yield_strength,
                                'Отн. удл.': elongation,
                                'Отн. суж.': reduction
                            })
                            
                    except Exception as e:
                        continue
    
    return pd.DataFrame(data_rows)

def interpolate_yield_strength(temp):
    """Линейная интерполяция нормативного предела текучести для стали марки 20"""
    known_points = [
        (20, 216),
        (250, 196),
        (400, 137),
        (450, 127)
    ]
    
    for t, value in known_points:
        if temp == t:
            return value
    
    if temp < 20:
        return 216
    elif 20 < temp <= 250:
        x1, y1 = 20, 216
        x2, y2 = 250, 196
    elif 250 < temp <= 400:
        x1, y1 = 250, 196
        x2, y2 = 400, 137
    elif 400 < temp <= 450:
        x1, y1 = 400, 137
        x2, y2 = 450, 127
    else:
        return 127
    
    result = y1 + (y2 - y1) * (temp - x1) / (x2 - x1)
    return round(result)

def parse_mapping_file(mapping_file):
    """Парсинг файла соответствия названий образцов"""
    try:
        if mapping_file.name.endswith('.xlsx'):
            df_mapping = pd.read_excel(mapping_file, header=None)
        else:
            return {}
        
        mapping = {}
        
        # Создаем список для сохранения порядка строк
        rows = []
        
        for idx, row in df_mapping.iterrows():
            if len(row) >= 2 and pd.notna(row[0]) and pd.notna(row[1]):
                new_name = str(row[0]).strip()
                lab_number = str(row[1]).strip()
                
                # Извлекаем числовую часть из лабораторного номера
                try:
                    numbers = re.findall(r'\d+', lab_number)
                    if numbers:
                        pipe_num = int(numbers[0])
                        rows.append({
                            'index': idx,
                            'pipe_num': pipe_num,
                            'new_name': new_name
                        })
                except ValueError:
                    continue
        
        # Сортируем строки по индексу в порядке возрастания (сверху вниз)
        rows.sort(key=lambda x: x['index'])
        
        # Присваиваем порядок от 1 до N (сохраняя порядок из файла)
        for order, row in enumerate(rows, 1):
            mapping[row['pipe_num']] = {
                'new_name': row['new_name'],
                'order': order
            }
        
        return mapping
    except Exception as e:
        st.error(f"Ошибка при чтении файла соответствия: {str(e)}")
        return {}

def get_test_data():
    """Возвращает тестовые данные из примера протокола"""
    test_data = [
        {'Клеймо': '1-1', 'Температура': 20, 'Предел прочности': 485, 'Предел текучести': 297, 'Отн. удл.': 30, 'Отн. суж.': 57},
        {'Клеймо': '1-2', 'Температура': 20, 'Предел прочности': 481, 'Предел текучести': 295, 'Отн. удл.': 33, 'Отн. суж.': 61},
        {'Клеймо': '1-3', 'Температура': 403, 'Предел прочности': 478, 'Предел текучести': 214, 'Отн. удл.': 28, 'Отн. суж.': 63},
        {'Клеймо': '1-4', 'Температура': 403, 'Предел прочности': 483, 'Предел текучести': 289, 'Отн. удл.': 24, 'Отн. суж.': 58},
        {'Клеймо': '2-1', 'Температура': 20, 'Предел прочности': 474, 'Предел текучести': 300, 'Отн. удл.': 36, 'Отн. суж.': 61},
        {'Клеймо': '2-2', 'Температура': 20, 'Предел прочности': 466, 'Предел текучести': 290, 'Отн. удл.': 37, 'Отн. суж.': 63},
        {'Клеймо': '2-3', 'Температура': 403, 'Предел прочности': 443, 'Предел текучести': 264, 'Отн. удл.': 27, 'Отн. суж.': 65},
        {'Клеймо': '2-4', 'Температура': 403, 'Предел прочности': 444, 'Предел текучести': 305, 'Отн. удл.': 25, 'Отн. суж.': 62},
        {'Клеймо': '3-1', 'Температура': 20, 'Предел прочности': 488, 'Предел текучести': 301, 'Отн. удл.': 30, 'Отн. суж.': 60},
        {'Клеймо': '3-2', 'Температура': 20, 'Предел прочности': 487, 'Предел текучести': 305, 'Отн. удл.': 34, 'Отн. суж.': 60},
        {'Клеймо': '3-3', 'Температура': 403, 'Предел прочности': 428, 'Предел текучести': 250, 'Отн. удл.': 31, 'Отн. суж.': 65},
        {'Клеймо': '3-4', 'Температура': 403, 'Предел прочности': 427, 'Предел текучести': 249, 'Отн. удл.': 32, 'Отн. суж.': 63},
        {'Клеймо': '4-1', 'Температура': 20, 'Предел прочности': 525, 'Предел текучести': 401, 'Отн. удл.': 28, 'Отн. суж.': 59},
        {'Клеймо': '4-2', 'Температура': 20, 'Предел прочности': 520, 'Предел текучести': 336, 'Отн. удл.': 35, 'Отн. суж.': 60},
        {'Клеймо': '4-3', 'Температура': 403, 'Предел прочности': 450, 'Предел текучести': 242, 'Отн. удл.': 28, 'Отн. суж.': 60},
        {'Клеймо': '4-4', 'Температура': 403, 'Предел прочности': 447, 'Предел текучести': 246, 'Отн. удл.': 29, 'Отн. суж.': 62},
        {'Клеймо': '5-1', 'Температура': 20, 'Предел прочности': 494, 'Предел текучести': 266, 'Отн. удл.': 39, 'Отн. суж.': 60},
        {'Клеймо': '5-2', 'Температура': 20, 'Предел прочности': 496, 'Предел текучести': 273, 'Отн. удл.': 35, 'Отн. суж.': 59},
        {'Клеймо': '5-3', 'Температура': 403, 'Предел прочности': 430, 'Предел текучести': 232, 'Отн. удл.': 31, 'Отн. суж.': 64},
        {'Клеймо': '5-4', 'Температура': 403, 'Предел прочности': 436, 'Предел текучести': 224, 'Отн. удл.': 28, 'Отн. суж.': 68},
        {'Клеймо': '6-1', 'Температура': 20, 'Предел прочности': 502, 'Предел текучести': 295, 'Отн. удл.': 31, 'Отн. суж.': 59},
        {'Клеймо': '6-2', 'Температура': 20, 'Предел прочности': 503, 'Предел текучести': 294, 'Отн. удл.': 34, 'Отн. суж.': 55},
        {'Клеймо': '6-3', 'Температура': 403, 'Предел прочности': 469, 'Предел текучести': 254, 'Отн. удл.': 27, 'Отн. суж.': 64},
        {'Клеймо': '6-4', 'Температура': 403, 'Предел прочности': 454, 'Предел текучести': 223, 'Отн. удл.': 24, 'Отн. суж.': 65},
        {'Клеймо': '7-1', 'Температура': 20, 'Предел прочности': 504, 'Предел текучести': 329, 'Отн. удл.': 28, 'Отн. суж.': 58},
        {'Клеймо': '7-2', 'Температура': 20, 'Предел прочности': 499, 'Предел текучести': 314, 'Отн. удл.': 35, 'Отн. суж.': 57},
        {'Клеймо': '7-3', 'Температура': 403, 'Предел прочности': 459, 'Предел текучести': 278, 'Отн. удл.': 28, 'Отн. суж.': 67},
        {'Клеймо': '7-4', 'Температура': 403, 'Предел прочности': 457, 'Предел текучести': 264, 'Отн. удл.': 24, 'Отн. суж.': 63},
    ]
    
    return pd.DataFrame(test_data)

def create_detailed_dataframe(df, mapping=None):
    """Создание детализированной таблицы с добавлением нормативных значений"""
    if df.empty:
        return pd.DataFrame()
    
    # Извлекаем номер трубы из клейма
    df['Номер трубы'] = df['Клеймо'].apply(lambda x: int(x.split('-')[0]) if '-' in str(x) else 0)
    df['Номер образца'] = df['Клеймо'].apply(lambda x: int(x.split('-')[1]) if '-' in str(x) else 0)
    
    # Определяем порядок следования образцов
    if mapping:
        # Создаем список номеров труб в порядке из mapping
        sorted_pipes = []
        other_pipes = []
        
        for pipe_num in df['Номер трубы'].unique():
            if pipe_num in mapping:
                sorted_pipes.append(pipe_num)
            else:
                other_pipes.append(pipe_num)
        
        # Сортируем по порядку из mapping
        sorted_pipes.sort(key=lambda x: mapping[x]['order'])
        # Сортируем остальные по возрастанию
        other_pipes.sort()
        
        # Объединяем списки
        ordered_pipes = sorted_pipes + other_pipes
        
        # Создаем столбец с порядком для сортировки
        def get_order(pipe_num):
            if pipe_num in mapping:
                return mapping[pipe_num]['order']
            else:
                return 999 + pipe_num
        
        df['Порядок'] = df['Номер трубы'].apply(get_order)
        df['Новое название'] = df['Номер трубы'].apply(
            lambda x: mapping.get(x, {}).get('new_name', f"Труба {x}")
        )
        
        # Сортируем по порядку, затем по температуре, затем по номеру образца
        df = df.sort_values(['Порядок', 'Температура', 'Номер образца'])
    else:
        df['Новое название'] = df['Номер трубы'].apply(lambda x: f"Труба {x}")
        df = df.sort_values(['Номер трубы', 'Температура', 'Номер образца'])
        ordered_pipes = sorted(df['Номер трубы'].unique())
    
    detailed_rows = []
    
    # Проходим по трубам в нужном порядке
    for pipe_num in ordered_pipes:
        pipe_data = df[df['Номер трубы'] == pipe_num]
        
        # Определяем название образца
        if mapping and pipe_num in mapping:
            pipe_name = mapping[pipe_num]['new_name']
        else:
            pipe_name = f"Труба {pipe_num}"
        
        # Группируем по температуре
        for temp in sorted(pipe_data['Температура'].unique()):
            temp_data = pipe_data[pipe_data['Температура'] == temp]
            
            # Добавляем строки для каждого образца
            for _, row in temp_data.iterrows():
                detailed_rows.append({
                    'Образец': pipe_name,
                    'Клеймо образца (лаборатория)': row['Клеймо'],
                    'Температура, °C': temp,
                    'Предел прочности, МПа': int(round(row['Предел прочности'])),
                    'Предел текучести, МПа': int(round(row['Предел текучести'])),
                    'Отн. удл., %': int(round(row['Отн. удл.'])),
                    'Отн. суж., %': int(round(row['Отн. суж.']))
                })
            
            # Добавляем строку со средними значениями
            if len(temp_data) > 0:
                detailed_rows.append({
                    'Образец': pipe_name,
                    'Клеймо образца (лаборатория)': 'Среднее',
                    'Температура, °C': temp,
                    'Предел прочности, МПа': int(round(temp_data['Предел прочности'].mean())),
                    'Предел текучести, МПа': int(round(temp_data['Предел текучести'].mean())),
                    'Отн. удл., %': int(round(temp_data['Отн. удл.'].mean())),
                    'Отн. суж., %': int(round(temp_data['Отн. суж.'].mean()))
                })
        
        # Добавляем пустую строку между образцами
        detailed_rows.append({
            'Образец': '',
            'Клеймо образца (лаборатория)': '',
            'Температура, °C': '',
            'Предел прочности, МПа': '',
            'Предел текучести, МПа': '',
            'Отн. удл., %': '',
            'Отн. суж., %': ''
        })
    
    # Удаляем последнюю пустую строку
    if detailed_rows and all(v == '' for v in detailed_rows[-1].values()):
        detailed_rows.pop()
    
    # Добавляем нормативные значения
    detailed_rows.append({
        'Образец': 'Требования [3] для стали марки 20',
        'Клеймо образца (лаборатория)': '',
        'Температура, °C': 20,
        'Предел прочности, МПа': '412-549',
        'Предел текучести, МПа': 216,
        'Отн. удл., %': 24,
        'Отн. суж., %': 45
    })
    
    # Добавляем нормативные значения для повышенных температур, которые есть в данных
    unique_temps = sorted([t for t in df['Температура'].unique() if t > 20])
    
    for temp in unique_temps:
        normative_yield = interpolate_yield_strength(temp)
        
        detailed_rows.append({
            'Образец': 'Требования [3] для стали марки 20',
            'Клеймо образца (лаборатория)': '',
            'Температура, °C': temp,
            'Предел прочности, МПа': '-',
            'Предел текучести, МПа': normative_yield,
            'Отн. удл., %': '-',
            'Отн. суж., %': '-'
        })
    
    detailed_df = pd.DataFrame(detailed_rows)
    return detailed_df

def create_summary_table(df, mapping=None):
    """Создание сводной таблицы"""
    if df.empty:
        return pd.DataFrame(), []
    
    # Извлекаем номер трубы
    df['Номер трубы'] = df['Клеймо'].apply(lambda x: int(x.split('-')[0]) if '-' in str(x) else 0)
    
    # Определяем порядок
    if mapping:
        summary_rows = []
        for pipe_num in df['Номер трубы'].unique():
            pipe_data = df[df['Номер трубы'] == pipe_num]
            high_temp_data = pipe_data[pipe_data['Температура'] > 20]
            
            if not high_temp_data.empty:
                avg_yield = int(round(high_temp_data['Предел текучести'].mean()))
                
                if pipe_num in mapping:
                    pipe_name = mapping[pipe_num]['new_name']
                    order = mapping[pipe_num]['order']
                else:
                    pipe_name = f"Труба {pipe_num}"
                    order = 999 + pipe_num
                
                summary_rows.append({
                    'Порядок': order,
                    'Образец': pipe_name,
                    'Средний предел текучести, МПа': avg_yield
                })
        
        # Сортируем по порядку
        summary_df = pd.DataFrame(summary_rows)
        if not summary_df.empty:
            summary_df = summary_df.sort_values('Порядок').drop('Порядок', axis=1)
    else:
        summary_rows = []
        for pipe_num in sorted(df['Номер трубы'].unique()):
            pipe_data = df[df['Номер трубы'] == pipe_num]
            high_temp_data = pipe_data[pipe_data['Температура'] > 20]
            
            if not high_temp_data.empty:
                avg_yield = int(round(high_temp_data['Предел текучести'].mean()))
                
                summary_rows.append({
                    'Образец': f"Труба {pipe_num}",
                    'Средний предел текучести, МПа': avg_yield
                })
        
        summary_df = pd.DataFrame(summary_rows)
    
    temperatures_above_20 = sorted([t for t in df['Температура'].unique() if t > 20])
    return summary_df, temperatures_above_20

def create_word_report(detailed_df, summary_df, high_temps):
    """Создание Word документа с таблицами"""
    doc = Document()
    
    # Настройка стилей
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    # Заголовок
    title = doc.add_paragraph('Таблица механических свойств')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].font.size = Pt(14)
    title.runs[0].bold = True
    
    # Дата
    date_para = doc.add_paragraph()
    date_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    date_run = date_para.add_run(f"Дата формирования: {datetime.now().strftime('%d.%m.%Y')}")
    date_run.font.size = Pt(10)
    
    doc.add_paragraph()
    
    # Таблица 1
    doc.add_paragraph('1. Результаты механических испытаний образцов')
    doc.paragraphs[-1].runs[0].bold = True
    
    # Создаем таблицу
    if not detailed_df.empty:
        table1 = doc.add_table(rows=len(detailed_df)+1, cols=len(detailed_df.columns))
        table1.style = 'Table Grid'
        table1.autofit = False
        
        # Заголовки
        headers = detailed_df.columns.tolist()
        for i, header in enumerate(headers):
            cell = table1.cell(0, i)
            cell.text = str(header)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].runs[0].font.bold = True
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        
        # Данные
        for i, row in detailed_df.iterrows():
            for j, col in enumerate(headers):
                cell = table1.cell(i+1, j)
                value = str(row[col]) if pd.notna(row[col]) else ''
                cell.text = value
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                
                # Жирный шрифт для средних значений и нормативных строк
                if 'Среднее' in value or 'Требования' in str(row.get('Образец', '')):
                    cell.paragraphs[0].runs[0].font.bold = True
    
    doc.add_page_break()
    
    # Таблица 2
    if not summary_df.empty:
        if high_temps:
            temp_str = ", ".join(map(str, high_temps))
            title2 = doc.add_paragraph(f'2. Средние пределы текучести при повышенной температуре ({temp_str}°C)')
        else:
            title2 = doc.add_paragraph('2. Средние пределы текучести при повышенной температуре')
        title2.runs[0].bold = True
        
        table2 = doc.add_table(rows=len(summary_df)+1, cols=len(summary_df.columns))
        table2.style = 'Table Grid'
        
        # Заголовки
        headers2 = summary_df.columns.tolist()
        for i, header in enumerate(headers2):
            cell = table2.cell(0, i)
            cell.text = str(header)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].runs[0].font.bold = True
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        
        # Данные
        for i, row in summary_df.iterrows():
            for j, col in enumerate(headers2):
                cell = table2.cell(i+1, j)
                value = str(row[col]) if pd.notna(row[col]) else ''
                cell.text = value
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    
    # Сохраняем в BytesIO
    doc_bytes = BytesIO()
    doc.save(doc_bytes)
    doc_bytes.seek(0)
    
    return doc_bytes

def main():
    """Основная функция"""
    st.markdown('<h1 class="main-header">📊 Обработчик протоколов механических испытаний</h1>', unsafe_allow_html=True)
    
    # Информационный блок
    st.markdown("""
    <div class="info-box">
    <h4>📁 Загрузите файлы для обработки</h4>
    <p>1. Протокол испытаний (DOCX) - обязательный<br>
    2. Файл соответствия названий (Excel) - опционально, для переименования образцов</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Два загрузчика файлов
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Протокол испытаний")
        uploaded_protocol = st.file_uploader(
            "Загрузите протокол испытаний (DOCX)",
            type=['docx'],
            key="protocol",
            help="Основной файл с результатами механических испытаний"
        )
    
    with col2:
        st.subheader("📊 Файл соответствия")
        uploaded_mapping = st.file_uploader(
            "Загрузите файл соответствия названий (Excel)",
            type=['xlsx', 'xls'],
            key="mapping",
            help="Excel файл с двумя столбцами: новое название и номер из протокола"
        )
    
    # Боковая панель
    with st.sidebar:
        st.header("⚙️ Настройки")
        st.markdown("---")
        
        st.subheader("Параметры обработки")
        use_test_data = st.checkbox("Использовать тестовые данные", value=True,
                                   help="Использовать примерные данные для демонстрации")
        
        st.subheader("Нормативные значения")
        st.markdown("""
        **Сталь марки 20:**
        - 20°C: 216 МПа
        - 250°C: 196 МПа
        - 400°C: 137 МПа
        - 450°C: 127 МПа
        """)
    
    # Обработка файлов
    if uploaded_protocol is not None or use_test_data:
        try:
            with st.spinner("📊 Обработка данных..."):
                # Парсим файл соответствия если есть
                mapping = {}
                if uploaded_mapping is not None:
                    mapping = parse_mapping_file(uploaded_mapping)
                    if mapping:
                        st.success(f"✅ Загружено {len(mapping)} соответствий названий")
                
                # Получаем данные протокола
                if use_test_data:
                    df = get_test_data()
                    file_source = "тестовые данные"
                else:
                    file_content = uploaded_protocol.read()
                    df = parse_protocol_from_docx(file_content)
                    file_source = uploaded_protocol.name
                
                if df.empty:
                    st.error("Не удалось извлечь данные из файла.")
                    st.info("Попробуйте включить опцию 'Использовать тестовые данные'")
                    return
                
                # Создаем таблицы
                detailed_df = create_detailed_dataframe(df, mapping)
                summary_df, high_temps = create_summary_table(df, mapping)
                
                # Показываем статистику
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Обработано образцов", len(df))
                with col2:
                    unique_pipes = df['Клеймо'].apply(lambda x: str(x).split('-')[0] if '-' in str(x) else '0').nunique()
                    st.metric("Количество труб", unique_pipes)
                with col3:
                    temps = sorted(df['Температура'].unique())
                    st.metric("Температуры испытаний", f"{len(temps)} видов")
                
                # Предпросмотр
                st.subheader("📋 Предпросмотр основной таблицы")
                st.dataframe(detailed_df, use_container_width=True, hide_index=True)
                
                if not summary_df.empty:
                    st.subheader("📊 Предпросмотр сводной таблицы")
                    st.dataframe(summary_df, use_container_width=True, hide_index=True)
                
                # Создание Word документа
                st.subheader("📥 Скачать отчет")
                
                doc_bytes = create_word_report(detailed_df, summary_df, high_temps)
                
                # Кнопка скачивания
                filename = f"Таблица_механических_свойств_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                
                st.download_button(
                    label="⬇️ Скачать отчет в Word",
                    data=doc_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
        except Exception as e:
            st.error(f"Ошибка при обработке: {str(e)}")
    
    else:
        # Инструкция
        st.info("👈 Загрузите протокол испытаний (DOCX файл) для начала обработки")

if __name__ == "__main__":
    main()
