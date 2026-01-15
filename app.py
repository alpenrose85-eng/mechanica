import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime
import tempfile
import os
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

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
    .download-btn {
        background-color: #1E3A8A;
        color: white;
        padding: 10px 20px;
        border-radius: 5px;
        text-decoration: none;
        display: inline-block;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

def parse_protocol_from_text(text):
    """Парсинг данных из текста протокола"""
    # Поиск таблицы с данными
    lines = text.split('\n')
    data_rows = []
    in_table = False
    
    for line in lines:
        if '|' in line and 'Клеймо' in line:
            in_table = True
            continue
        if in_table and '|' in line and '------' not in line:
            # Обрабатываем строку таблицы
            parts = [p.strip() for p in line.split('|') if p.strip()]
            if len(parts) >= 13:  # Проверяем, что строка содержит нужные данные
                try:
                    # Извлекаем данные
                    sample_mark = parts[1]
                    strength = float(parts[9].replace(' ', ''))
                    yield_strength = float(parts[10].replace(' ', ''))
                    elongation = float(parts[11].replace(' ', ''))
                    reduction = float(parts[12].replace(' ', ''))
                    temp_match = re.search(r'(\d+)', parts[5])
                    temperature = int(temp_match.group(1)) if temp_match else 20
                    
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
    
    return pd.DataFrame(data_rows)

def create_detailed_dataframe(df):
    """Создание детализированной таблицы с группировкой по трубам и температурам"""
    # Извлекаем номер трубы из клейма
    df['Номер трубы'] = df['Клеймо'].apply(lambda x: int(x.split('-')[0]))
    df['Номер образца'] = df['Клеймо'].apply(lambda x: int(x.split('-')[1]))
    
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
            if len(temp_data) > 1:
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
    """Создание сводной таблицы со средними значениями при повышенной температуре"""
    # Извлекаем номер трубы
    df['Номер трубы'] = df['Клеймо'].apply(lambda x: int(x.split('-')[0]))
    
    summary_rows = []
    temperatures_above_20 = sorted(df[df['Температура'] > 20]['Температура'].unique())
    
    if temperatures_above_20:
        for pipe_num in sorted(df['Номер трубы'].unique()):
            pipe_data = df[df['Номер трубы'] == pipe_num]
            high_temp_data = pipe_data[pipe_data['Температура'] > 20]
            
            if not high_temp_data.empty:
                avg_yield = round(high_temp_data['Предел текучести'].mean(), 1)
                
                # Для каждой повышенной температуры
                for temp in temperatures_above_20:
                    temp_data = pipe_data[pipe_data['Температура'] == temp]
                    if not temp_data.empty:
                        summary_rows.append({
                            'Образец': f"Труба {pipe_num}",
                            'Температура, °C': temp,
                            'Средний предел текучести, МПа': avg_yield
                        })
    
    summary_df = pd.DataFrame(summary_rows)
    if not summary_df.empty:
        summary_df = summary_df.drop_duplicates(subset=['Образец'])
    
    return summary_df, temperatures_above_20

def create_word_report(detailed_df, summary_df, high_temps):
    """Создание Word документа с таблицами"""
    doc = Document()
    
    # Настройка стилей
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    
    # Заголовок
    title = doc.add_heading('Таблица механических свойств', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Добавляем дату
    date_para = doc.add_paragraph()
    date_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    date_run = date_para.add_run(f"Дата формирования: {datetime.now().strftime('%d.%m.%Y')}")
    date_run.font.size = Pt(10)
    
    doc.add_paragraph()  # Пустая строка
    
    # ТАБЛИЦА 1: Детализированная таблица
    doc.add_heading('1. Результаты механических испытаний образцов', level=1)
    
    # Создаем таблицу
    num_rows = len(detailed_df) + 1  + 1# +1 для заголовков
    num_cols = len(detailed_df.columns)
    
    table1 = doc.add_table(rows=num_rows, cols=num_cols)
    table1.style = 'Table Grid'
    table1.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Заголовки таблицы
    headers = detailed_df.columns.tolist()
    for i, header in enumerate(headers):
        cell = table1.cell(0, i)
        cell.text = str(header)
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.runs[0].font.bold = True
    
    # Заполняем таблицу данными
    for i, row in detailed_df.iterrows():
        for j, col in enumerate(headers):
            cell = table1.cell(i + 1, j)
            cell.text = str(row[col]) if pd.notna(row[col]) else ''
            paragraph = cell.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Выделяем строки со средними значениями
            if row['Клеймо образца'] == 'Среднее':
                for run in paragraph.runs:
                    run.font.bold = True
    
    doc.add_page_break()
    
    # ТАБЛИЦА 2: Сводная таблица
    if not summary_df.empty:
        if high_temps:
            temp_str = ", ".join(map(str, high_temps))
            doc.add_heading(f'2. Средние пределы текучести при повышенной температуре ({temp_str}°C)', level=1)
        else:
            doc.add_heading('2. Средние пределы текучести при повышенной температуре', level=1)
        
        # Создаем таблицу
        num_rows_summary = len(summary_df) + 1
        num_cols_summary = len(summary_df.columns)
        
        table2 = doc.add_table(rows=num_rows_summary, cols=num_cols_summary)
        table2.style = 'Table Grid'
        table2.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Заголовки
        summary_headers = summary_df.columns.tolist()
        for i, header in enumerate(summary_headers):
            cell = table2.cell(0, i)
            cell.text = str(header)
            paragraph = cell.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.runs[0].font.bold = True
        
        # Данные
        for i, row in summary_df.iterrows():
            for j, col in enumerate(summary_headers):
                cell = table2.cell(i + 1, j)
                cell.text = str(row[col]) if pd.notna(row[col]) else ''
                paragraph = cell.paragraphs[0]
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Сохраняем в BytesIO
    doc_bytes = BytesIO()
    doc.save(doc_bytes)
    doc_bytes.seek(0)
    
    return doc_bytes

def main():
    """Основная функция"""
    st.markdown('<h1 class="main-header">📊 Обработчик протоколов механических испытаний</h1>', unsafe_allow_html=True)
    
    # Загрузка файла
    st.markdown('<div class="info-box">', unsafe_allow_html=True)
    st.subheader("📁 Загрузите файл протокола")
    
    uploaded_file = st.file_uploader(
        "Выберите файл с протоколом испытаний",
        type=['docx', 'txt'],
        help="Поддерживаемые форматы: DOCX, TXT"
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_file is not None:
        try:
            # Чтение файла
            if uploaded_file.name.endswith('.docx'):
                doc = Document(uploaded_file)
                text = '\n'.join([paragraph.text for paragraph in doc.paragraphs])
            else:
                text = uploaded_file.getvalue().decode('utf-8')
            
            # Парсинг данных
            with st.spinner("📊 Обработка протокола..."):
                df = parse_protocol_from_text(text)
                
                if df.empty:
                    st.error("Не удалось извлечь данные из файла. Проверьте формат протокола.")
                    return
                
                # Создаем таблицы
                detailed_df = create_detailed_dataframe(df)
                summary_df, high_temps = create_summary_table(df)
                
                # Показываем статистику
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Обработано образцов", len(df))
                with col2:
                    st.metric("Количество труб", df['Клеймо'].apply(lambda x: x.split('-')[0]).nunique())
                with col3:
                    temps = df['Температура'].unique()
                    st.metric("Температурные режимы", f"{len(temps)} ({', '.join(map(str, sorted(temps)))})")
                
                # Предпросмотр таблиц
                st.subheader("📋 Предпросмотр таблицы 1 (детализированной)")
                st.dataframe(
                    detailed_df,
                    use_container_width=True,
                    hide_index=True
                )
                
                if not summary_df.empty:
                    st.subheader("📊 Предпросмотр таблицы 2 (сводной)")
                    st.dataframe(
                        summary_df,
                        use_container_width=True,
                        hide_index=True
                    )
                
                # Создание Word документа
                st.subheader("📥 Скачать отчет")
                
                with st.spinner("Формирование Word документа..."):
                    doc_bytes = create_word_report(detailed_df, summary_df, high_temps)
                    
                    # Кнопка скачивания
                    st.download_button(
                        label="⬇️ Скачать отчет в Word (Таблица механических свойств.docx)",
                        data=doc_bytes,
                        file_name=f"Таблица_механических_свойств_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        help="Нажмите, чтобы скачать отчет в формате Word с двумя таблицами"
                    )
                
                # Информация о данных
                with st.expander("📝 Информация об обработанных данных"):
                    st.write("**Структура данных:**")
                    st.write(f"- Всего строк в протоколе: {len(df)}")
                    st.write(f"- Уникальные номера труб: {sorted(df['Клеймо'].apply(lambda x: int(x.split('-')[0])).unique())}")
                    st.write(f"- Температуры испытаний: {sorted(df['Температура'].unique())}°C")
                    
                    st.write("\n**Поведение программы:**")
                    st.write("- Автоматически определяет количество образцов для каждой температуры")
                    st.write("- Рассчитывает средние значения для каждой группы образцов")
                    st.write("- Создает отдельную таблицу для повышенных температур (>20°C)")
                    
        except Exception as e:
            st.error(f"Ошибка при обработке файла: {str(e)}")
            st.info("Пожалуйста, убедитесь, что файл соответствует формату протокола испытаний")
    
    else:
        # Инструкция
        st.info("👈 Загрузите файл протокола для начала обработки")
        
        with st.expander("ℹ️ Как подготовить файл"):
            st.markdown("""
            **Требования к формату протокола:**
            1. Файл должен содержать таблицу с результатами испытаний
            2. В таблице должны быть колонки:
               - Клеймо образца (формат "X-Y", где X - номер трубы)
               - Температура испытания
               - Предел прочности (МПа)
               - Предел текучести (МПа)
               - Относительное удлинение (%)
               - Относительное сужение (%)
            
            **Пример клейма:**
            - "1-1" - труба 1, образец 1
            - "1-2" - труба 1, образец 2
            - "2-1" - труба 2, образец 1
            
            **Что делает программа:**
            1. Автоматически группирует образцы по номерам труб
            2. Для каждой температуры создает отдельные строки
            3. Рассчитывает средние значения для каждой группы
            4. Формирует две таблицы в Word документе
            """)
        
        # Пример данных
        st.subheader("📋 Пример структуры данных")
        example_data = pd.DataFrame({
            'Клеймо': ['1-1', '1-2', '1-3', '1-4', '2-1', '2-2'],
            'Температура': [20, 20, 403, 403, 20, 403],
            'Предел прочности': [485, 481, 478, 483, 474, 443],
            'Предел текучести': [297, 295, 214, 289, 300, 264],
            'Отн. удл.': [30, 33, 28, 24, 36, 27],
            'Отн. суж.': [57, 61, 63, 58, 61, 65]
        })
        st.dataframe(example_data, use_container_width=True)

if __name__ == "__main__":
    main()
