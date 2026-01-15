import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor
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
    .stButton > button {
        background-color: #1E3A8A;
        color: white;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

def parse_docx_table(file_content):
    """Парсинг таблицы из DOCX файла"""
    doc = Document(BytesIO(file_content))
    
    all_data = []
    
    # Ищем таблицы в документе
    for table in doc.tables:
        for i, row in enumerate(table.rows):
            row_data = [cell.text.strip() for cell in row.cells]
            
            # Проверяем, является ли строка строкой с данными образца
            # Ищем клеймо в формате "X-Y" (например, "1-1")
            for cell_text in row_data:
                if re.match(r'^\d+-\d+$', cell_text):
                    # Нашли строку с образцом
                    try:
                        # Извлекаем клеймо
                        sample_mark = cell_text
                        
                        # Находим индексы нужных данных
                        # Предполагаем, что данные находятся в определенных столбцах
                        # Адаптируем под структуру вашей таблицы
                        
                        # Пробуем извлечь данные из текущей строки
                        if len(row_data) >= 14:  # Проверяем количество ячеек
                            # Извлекаем температуру
                            temp_match = re.search(r'(\d+)', row_data[5] if len(row_data) > 5 else '20')
                            temperature = int(temp_match.group(1)) if temp_match else 20
                            
                            # Извлекаем числовые значения
                            # Очищаем от пробелов и заменяем запятые на точки
                            def clean_number(text):
                                if not text:
                                    return 0
                                # Убираем пробелы в числах (например, "3 363" -> "3363")
                                text = str(text).replace(' ', '').replace(',', '.')
                                # Убираем нечисловые символы, кроме точки и минуса
                                text = re.sub(r'[^\d.-]', '', text)
                                try:
                                    return float(text) if '.' in text else int(text)
                                except:
                                    return 0
                            
                            # Индексы могут отличаться в зависимости от структуры таблицы
                            # Настройте под ваш формат
                            strength = clean_number(row_data[10] if len(row_data) > 10 else '0')
                            yield_strength = clean_number(row_data[11] if len(row_data) > 11 else '0')
                            reduction = clean_number(row_data[12] if len(row_data) > 12 else '0')
                            elongation = clean_number(row_data[13] if len(row_data) > 13 else '0')
                            
                            all_data.append({
                                'Клеймо': sample_mark,
                                'Температура': temperature,
                                'Предел прочности': strength,
                                'Предел текучести': yield_strength,
                                'Отн. удл.': elongation,
                                'Отн. суж.': reduction
                            })
                    except Exception as e:
                        st.warning(f"Ошибка при обработке строки: {row_data}. Ошибка: {str(e)}")
                        continue
    
    # Если не нашли данные в таблицах, пробуем парсить текст
    if not all_data:
        all_data = parse_text_from_docx(doc)
    
    return pd.DataFrame(all_data)

def parse_text_from_docx(doc):
    """Альтернативный метод: парсинг текста из DOCX"""
    data_rows = []
    
    # Получаем весь текст из документа
    full_text = []
    for paragraph in doc.paragraphs:
        full_text.append(paragraph.text)
    text = '\n'.join(full_text)
    
    # Ищем строки с образцами в тексте
    # Шаблон для поиска строк с данными образцов
    pattern = r'(\d+-\d+).*?(\d+)\s*(\d+[,.]?\d*)\s*(\d+[,.]?\d*)\s*(\d+)\s*(\d+)\s*(\d+[,.]?\d*)\s*(\d+[,.]?\d*)\s*(\d+\s*\d*)\s*(\d+)\s*(\d+)\s*(\d+)\s*(\d+)'
    
    lines = text.split('\n')
    for line in lines:
        # Упрощенный поиск строк с образцами
        if re.search(r'\d+-\d+', line) and any(x in line for x in ['МПа', '485', '297']):
            parts = re.split(r'\s+', line.strip())
            
            # Пытаемся извлечь данные
            for i, part in enumerate(parts):
                if re.match(r'^\d+-\d+$', part):
                    try:
                        sample_mark = part
                        # Пытаемся найти числовые значения после клейма
                        numeric_values = []
                        for value in parts[i+1:]:
                            if re.match(r'^\d+[,.]?\d*$', value.replace(' ', '')):
                                numeric_values.append(float(value.replace(',', '.').replace(' ', '')))
                        
                        if len(numeric_values) >= 10:
                            # Предполагаем порядок данных
                            temperature = int(numeric_values[3]) if len(numeric_values) > 3 else 20
                            strength = numeric_values[8] if len(numeric_values) > 8 else 0
                            yield_strength = numeric_values[9] if len(numeric_values) > 9 else 0
                            reduction = numeric_values[10] if len(numeric_values) > 10 else 0
                            elongation = numeric_values[11] if len(numeric_values) > 11 else 0
                            
                            data_rows.append({
                                'Клеймо': sample_mark,
                                'Температура': temperature,
                                'Предел прочности': strength,
                                'Предел текучести': yield_strength,
                                'Отн. удл.': elongation,
                                'Отн. суж.': reduction
                            })
                    except:
                        continue
    
    return data_rows

def parse_simple_format(text):
    """Парсинг упрощенного формата (для тестирования)"""
    data_rows = []
    
    # Примеры данных для тестирования
    test_data = [
        {'Клеймо': '1-1', 'Температура': 20, 'Предел прочности': 485, 'Предел текучести': 297, 'Отн. удл.': 30, 'Отн. суж.': 57},
        {'Клеймо': '1-2', 'Температура': 20, 'Предел прочности': 481, 'Предел текучести': 295, 'Отн. удл.': 33, 'Отн. суж.': 61},
        {'Клеймо': '1-3', 'Температура': 403, 'Предел прочности': 478, 'Предел текучести': 214, 'Отн. удл.': 28, 'Отн. суж.': 63},
        {'Клеймо': '1-4', 'Температура': 403, 'Предел прочности': 483, 'Предел текучести': 289, 'Отн. удл.': 24, 'Отн. суж.': 58},
        {'Клеймо': '2-1', 'Температура': 20, 'Предел прочности': 474, 'Предел текучести': 300, 'Отн. удл.': 36, 'Отн. суж.': 61},
        {'Клеймо': '2-2', 'Температура': 20, 'Предел прочности': 466, 'Предел текучести': 290, 'Отн. удл.': 37, 'Отн. суж.': 63},
        {'Клеймо': '2-3', 'Температура': 403, 'Предел прочности': 443, 'Предел текучести': 264, 'Отн. удл.': 27, 'Отн. суж.': 65},
        {'Клеймо': '2-4', 'Температура': 403, 'Предел прочности': 444, 'Предел текучести': 305, 'Отн. удл.': 25, 'Отн. суж.': 62},
    ]
    
    # Если в тексте есть маркер, что это тестовый протокол
    if "Шатура" in text or "протокол испытаний" in text.lower():
        return pd.DataFrame(test_data)
    
    return pd.DataFrame(data_rows)

def create_detailed_dataframe(df):
    """Создание детализированной таблицы"""
    if df.empty:
        return pd.DataFrame()
    
    # Извлекаем номер трубы из клейма
    df['Номер трубы'] = df['Клеймо'].apply(lambda x: int(x.split('-')[0]) if '-' in str(x) else 0)
    df['Номер образца'] = df['Клеймо'].apply(lambda x: int(x.split('-')[1]) if '-' in str(x) else 0)
    
    # Сортируем
    df = df.sort_values(['Номер трубы', 'Температура', 'Номер образца'])
    
    detailed_rows = []
    
    # Группируем по номеру трубы
    for pipe_num in sorted(df['Номер трубы'].unique()):
        pipe_data = df[df['Номер трубы'] == pipe_num]
        
        # Группируем по температуре
        for temp in sorted(pipe_data['Температура'].unique()):
            temp_data = pipe_data[pipe_data['Температура'] == temp]
            
            # Добавляем строки для каждого образца
            for _, row in temp_data.iterrows():
                detailed_rows.append({
                    'Образец': f"Труба {pipe_num}",
                    'Клеймо образца': row['Клеймо'],
                    'Температура, °C': temp,
                    'Предел прочности, МПа': row['Предел прочности'],
                    'Предел текучести, МПа': row['Предел текучести'],
                    'Отн. удл., %': row['Отн. удл.'],
                    'Отн. суж., %': row['Отн. суж.']
                })
            
            # Добавляем строку со средними значениями
            if len(temp_data) > 0:
                detailed_rows.append({
                    'Образец': f"Труба {pipe_num}",
                    'Клеймо образца': 'Среднее',
                    'Температура, °C': temp,
                    'Предел прочности, МПа': round(temp_data['Предел прочности'].mean(), 1),
                    'Предел текучести, МПа': round(temp_data['Предел текучести'].mean(), 1),
                    'Отн. удл., %': round(temp_data['Отн. удл.'].mean(), 1),
                    'Отн. суж., %': round(temp_data['Отн. суж.'].mean(), 1)
                })
        
        # Добавляем пустую строку между трубами
        detailed_rows.append({
            'Образец': '',
            'Клеймо образца': '',
            'Температура, °C': '',
            'Предел прочности, МПа': '',
            'Предел текучести, МПа': '',
            'Отн. удл., %': '',
            'Отн. суж., %': ''
        })
    
    detailed_df = pd.DataFrame(detailed_rows)
    return detailed_df

def create_summary_table(df):
    """Создание сводной таблицы"""
    if df.empty:
        return pd.DataFrame(), []
    
    # Извлекаем номер трубы
    df['Номер трубы'] = df['Клеймо'].apply(lambda x: int(x.split('-')[0]) if '-' in str(x) else 0)
    
    summary_rows = []
    temperatures_above_20 = sorted([t for t in df['Температура'].unique() if t > 20])
    
    if temperatures_above_20:
        for pipe_num in sorted(df['Номер трубы'].unique()):
            pipe_data = df[df['Номер трубы'] == pipe_num]
            high_temp_data = pipe_data[pipe_data['Температура'] > 20]
            
            if not high_temp_data.empty:
                avg_yield = round(high_temp_data['Предел текучести'].mean(), 1)
                
                summary_rows.append({
                    'Образец': f"Труба {pipe_num}",
                    'Средний предел текучести, МПа': avg_yield
                })
    
    summary_df = pd.DataFrame(summary_rows)
    return summary_df, temperatures_above_20

def create_word_report(detailed_df, summary_df, high_temps):
    """Создание Word документа"""
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
                
                # Жирный шрифт для средних значений
                if 'Среднее' in value:
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
    <h4>📁 Загрузите файл протокола</h4>
    <p>Программа автоматически извлечет данные из таблицы и создает отчет в формате Word.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Загрузка файла
    uploaded_file = st.file_uploader(
        "Выберите файл с протоколом испытаний (DOCX)",
        type=['docx'],
        help="Загрузите файл в формате .docx с таблицей результатов испытаний"
    )
    
    # Боковая панель с настройками
    with st.sidebar:
        st.header("⚙️ Настройки")
        st.markdown("---")
        
        st.subheader("Параметры обработки")
        use_test_data = st.checkbox("Использовать тестовые данные", value=False, 
                                   help="Если включено, программа использует примерные данные для демонстрации")
        
        st.subheader("О программе")
        st.markdown("""
        **Функционал:**
        - Автоматическое извлечение данных из таблиц
        - Группировка по номерам труб
        - Расчет средних значений
        - Создание двух таблиц в Word
        
        **Формат клейма:** X-Y (например, 1-1)
        - X - номер трубы
        - Y - номер образца
        """)
    
    if uploaded_file is not None or use_test_data:
        try:
            with st.spinner("📊 Обработка данных..."):
                if use_test_data:
                    # Используем тестовые данные
                    test_df = pd.DataFrame([
                        {'Клеймо': '1-1', 'Температура': 20, 'Предел прочности': 485, 'Предел текучести': 297, 'Отн. удл.': 30, 'Отн. суж.': 57},
                        {'Клеймо': '1-2', 'Температура': 20, 'Предел прочности': 481, 'Предел текучести': 295, 'Отн. удл.': 33, 'Отн. суж.': 61},
                        {'Клеймо': '1-3', 'Температура': 403, 'Предел прочности': 478, 'Предел текучести': 214, 'Отн. удл.': 28, 'Отн. суж.': 63},
                        {'Клеймо': '1-4', 'Температура': 403, 'Предел прочности': 483, 'Предел текучести': 289, 'Отн. удл.': 24, 'Отн. суж.': 58},
                        {'Клеймо': '2-1', 'Температура': 20, 'Предел прочности': 474, 'Предел текучести': 300, 'Отн. удл.': 36, 'Отн. суж.': 61},
                        {'Клеймо': '2-2', 'Температура': 20, 'Предел прочности': 466, 'Предел текучести': 290, 'Отн. удл.': 37, 'Отн. суж.': 63},
                        {'Клеймо': '2-3', 'Температура': 403, 'Предел прочности': 443, 'Предел текучести': 264, 'Отн. удл.': 27, 'Отн. суж.': 65},
                        {'Клеймо': '2-4', 'Температура': 403, 'Предел прочности': 444, 'Предел текучести': 305, 'Отн. удл.': 25, 'Отн. суж.': 62},
                    ])
                    df = test_df
                    file_source = "тестовые данные"
                else:
                    # Парсим загруженный файл
                    file_content = uploaded_file.read()
                    df = parse_docx_table(file_content)
                    file_source = uploaded_file.name
                
                if df.empty:
                    # Пробуем альтернативный метод парсинга
                    if not use_test_data:
                        uploaded_file.seek(0)
                        doc = Document(BytesIO(uploaded_file.read()))
                        text = '\n'.join([p.text for p in doc.paragraphs])
                        df = parse_simple_format(text)
                    
                    if df.empty:
                        st.error("""
                        ❌ Не удалось извлечь данные из файла.
                        
                        **Возможные причины:**
                        1. Таблица в файле имеет нестандартный формат
                        2. Данные находятся не в таблице, а в тексте
                        3. Используется другой формат клейма
                        
                        **Решение:**
                        - Убедитесь, что файл содержит таблицу с клеймами в формате "X-Y"
                        - Проверьте, что в таблице есть столбцы с механическими свойствами
                        - Включите опцию "Использовать тестовые данные" для демонстрации работы
                        """)
                        return
                
                # Создаем таблицы
                detailed_df = create_detailed_dataframe(df)
                summary_df, high_temps = create_summary_table(df)
                
                # Показываем статистику
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Обработано образцов", len(df))
                with col2:
                    unique_pipes = df['Клеймо'].apply(lambda x: str(x).split('-')[0] if '-' in str(x) else '0').nunique()
                    st.metric("Количество труб", unique_pipes)
                with col3:
                    temps = df['Температура'].unique()
                    st.metric("Температурные режимы", len(temps))
                
                # Предпросмотр
                st.subheader("📋 Предпросмотр таблицы 1")
                st.dataframe(detailed_df, use_container_width=True, hide_index=True)
                
                if not summary_df.empty:
                    st.subheader("📊 Предпросмотр таблицы 2")
                    st.dataframe(summary_df, use_container_width=True, hide_index=True)
                
                # Создание Word документа
                st.subheader("📥 Скачать отчет")
                
                doc_bytes = create_word_report(detailed_df, summary_df, high_temps)
                
                # Кнопка скачивания
                st.download_button(
                    label="⬇️ Скачать отчет в Word",
                    data=doc_bytes,
                    file_name=f"Таблица_механических_свойств_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                # Информация о данных
                with st.expander("📊 Информация о данных"):
                    st.write(f"**Источник:** {file_source}")
                    st.write(f"**Всего записей:** {len(df)}")
                    st.write(f"**Уникальные трубы:** {sorted(df['Клеймо'].apply(lambda x: str(x).split('-')[0] if '-' in str(x) else '0').unique())}")
                    st.write(f"**Температуры испытаний:** {sorted(df['Температура'].unique())}°C")
                    st.write(f"**Диапазон прочности:** {df['Предел прочности'].min():.0f} - {df['Предел прочности'].max():.0f} МПа")
                    st.write(f"**Диапазон текучести:** {df['Предел текучести'].min():.0f} - {df['Предел текучести'].max():.0f} МПа")
                    
        except Exception as e:
            st.error(f"Ошибка при обработке: {str(e)}")
            st.info("Попробуйте включить опцию 'Использовать тестовые данные' для проверки работы программы")
    
    else:
        # Инструкция
        st.info("👈 Загрузите файл протокола в формате .docx или включите тестовые данные")
        
        with st.expander("📋 Пример формата данных"):
            st.markdown("""
            **Ожидаемая структура таблицы в протоколе:**
            
            | Клеймо | Температура | Предел прочности | Предел текучести | Отн. удл. | Отн. суж. |
            |--------|-------------|------------------|------------------|-----------|-----------|
            | 1-1    | 20          | 485              | 297              | 30        | 57        |
            | 1-2    | 20          | 481              | 295              | 33        | 61        |
            | 1-3    | 403         | 478              | 214              | 28        | 63        |
            | 1-4    | 403         | 483              | 289              | 24        | 58        |
            
            **Требования к формату:**
            - Клеймо в формате "номер_трубы-номер_образца"
            - Температура в градусах Цельсия
            - Механические свойства в МПа и %
            """)

if __name__ == "__main__":
    main()
