import streamlit as st
import pandas as pd
import numpy as np
import json
from datetime import datetime, timedelta
import re

# Проверяем и импортируем Plotly
try:
    import plotly.express as px
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
    PLOTLY_AVAILABLE = True
except ImportError:
    st.error("Plotly не установлен. Установите с помощью: pip install plotly")
    PLOTLY_AVAILABLE = False

# Настройка страницы
st.set_page_config(
    page_title="📊 Дашборд успеваемости",
    page_icon="📊",
    layout="wide"
)

# Инициализация session state для пресетов
if 'filter_presets' not in st.session_state:
    st.session_state.filter_presets = {
        "Все данные": {
            "classes": [],
            "parallel_mode": False,
            "parallels": [],
            "subjects": [],
            "grade_range": [0, 100],
            "top_n": 10,
            "date_range": None
        },
        "2-е параллель": {
            "classes": [],
            "parallel_mode": True,
            "parallels": ["2"],
            "subjects": [],
            "grade_range": [0, 100],
            "top_n": 10,
            "date_range": None
        },
        "Старшие классы (10-11)": {
            "classes": [],
            "parallel_mode": True,
            "parallels": ["10", "11"],
            "subjects": [],
            "grade_range": [0, 100],
            "top_n": 15,
            "date_range": None
        },
        "Средние классы (7-9)": {
            "classes": [],
            "parallel_mode": True,
            "parallels": ["7", "8", "9"],
            "subjects": [],
            "grade_range": [0, 100],
            "top_n": 12,
            "date_range": None
        },
        "Точные науки": {
            "classes": [],
            "parallel_mode": False,
            "parallels": [],
            "subjects": ["Math", "Physics", "Chemistry", "CS", "Calc", "Further Math"],
            "grade_range": [0, 100],
            "top_n": 6,
            "date_range": None
        },
        "Языки": {
            "classes": [],
            "parallel_mode": False,
            "parallels": [],
            "subjects": ["English", "ESL", "Rus", "Kaz", "RusLit", "KazLit"],
            "grade_range": [0, 100],
            "top_n": 8,
            "date_range": None
        }
    }

if 'current_filters' not in st.session_state:
    st.session_state.current_filters = {
        "classes": [],
        "parallel_mode": False,
        "parallels": [],
        "subjects": [],
        "grade_range": [0, 100],
        "top_n": 10,
        "date_range": None,
        "selected_dates": []  # Новое: для чекбоксов дат
    }

# Инициализация фильтров таблицы
if 'table_filters' not in st.session_state:
    st.session_state.table_filters = {
        'Student': [],
        'Class': [],
        'Subject': [],
        'Date': []
    }


@st.cache_data
def load_data():
    """Загружает данные из Excel файла"""
    try:
        # Путь к файлу (в корне репозитория)
        df = pd.read_excel('Marks 2526.xlsx', sheet_name='Average year, no teacher')
        
        # Проверяем наличие нужных колонок
        required_cols = ['Student', 'Class', 'Subject', 'Average']
        missing = [col for col in required_cols if col not in df.columns]
        
        if missing:
            st.error(f"❌ Отсутствуют колонки: {missing}")
            st.write("Найденные колонки:", list(df.columns))
            return create_demo_data()
        
        # Проверяем наличие колонки Date
        if 'Date' not in df.columns:
            st.warning("⚠️ Колонка 'Date' не найдена. Графики трендов будут недоступны.")
            # Добавляем фиктивную дату (сегодня) для совместимости
            df['Date'] = pd.Timestamp.now()
        else:
            # Конвертируем дату
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        
        st.success(f"✅ Загружено {len(df):,} записей из Excel")
        return df
        
    except FileNotFoundError:
        st.error("❌ Файл 'Marks 2526.xlsx' не найден!")
        st.info("Используются демо-данные для демонстрации")
        return create_demo_data()
        
    except Exception as e:
        st.error(f"❌ Ошибка загрузки: {e}")
        st.info("Используются демо-данные для демонстрации")
        return create_demo_data()


def create_demo_data():
    """Создает демо-данные для демонстрации с датами"""
    np.random.seed(42)
    
    students = [f"Студент_{i}" for i in range(1, 101)]
    classes = ['2A', '2B', '3A', '3B', '4A', '4B', '5A', '5B', '6A', '6B', 
               '7A', '7B', '8A', '8B', '9A', '9B', '10A', '10B', '11A', '11B']
    subjects = ['Math', 'Physics', 'Chemistry', 'Biology', 'English', 'History', 
                'Geography', 'Literature', 'CS', 'Art', 'PE', 'Music']
    
    # Генерируем даты за последние 12 месяцев
    end_date = datetime.now()
    start_date = end_date - timedelta(days=365)
    
    data = []
    for student in students:
        class_name = np.random.choice(classes)
        student_subjects = np.random.choice(subjects, size=np.random.randint(6, 10), replace=False)
        
        # Несколько оценок для каждого предмета в разные даты
        for subject in student_subjects:
            num_grades = np.random.randint(3, 8)  # 3-7 оценок за год
            
            for _ in range(num_grades):
                # Случайная дата
                random_days = np.random.randint(0, 365)
                grade_date = start_date + timedelta(days=random_days)
                
                # Создаем более реалистичное распределение оценок с трендом
                base_score = 75
                if subject in ['Math', 'Physics', 'Chemistry']:
                    base_score = 70
                elif subject in ['Art', 'PE', 'Music']:
                    base_score = 85
                
                # Добавляем небольшой положительный тренд со временем
                time_factor = random_days / 365 * 3  # До +3 баллов за год
                
                score = np.random.normal(base_score + time_factor, 12)
                score = max(30, min(100, score))
                
                data.append({
                    'Student': student,
                    'Class': class_name,
                    'Subject': subject,
                    'Average': round(score, 1),
                    'Date': grade_date.date()
                })
    
    return pd.DataFrame(data)


def extract_parallel_from_class(class_name):
    """Извлекает номер параллели из названия класса"""
    match = re.search(r'^(\d+)', str(class_name))
    return match.group(1) if match else 'Другие'


def get_available_parallels(df):
    """Получает список доступных параллелей"""
    parallels = set()
    for class_name in df['Class'].unique():
        parallel = extract_parallel_from_class(class_name)
        if parallel != 'Другие':
            parallels.add(parallel)
    return sorted(list(parallels), key=lambda x: int(x) if x.isdigit() else 999)


def get_classes_by_parallels(df, selected_parallels):
    """Возвращает классы для выбранных параллелей"""
    if not selected_parallels:
        return []
    
    classes = []
    for class_name in df['Class'].unique():
        parallel = extract_parallel_from_class(class_name)
        if parallel in selected_parallels:
            classes.append(class_name)
    
    return sorted(classes)


def create_plotly_chart(chart_type, data, **kwargs):
    """Создает графики Plotly с обработкой ошибок"""
    if not PLOTLY_AVAILABLE:
        st.error("Plotly не доступен для создания интерактивных графиков")
        return None
    
    try:
        if chart_type == 'bar':
            return px.bar(data, **kwargs)
        elif chart_type == 'scatter':
            return px.scatter(data, **kwargs)
        elif chart_type == 'histogram':
            return px.histogram(data, **kwargs)
        elif chart_type == 'box':
            return px.box(data, **kwargs)
        elif chart_type == 'heatmap':
            return px.imshow(data, **kwargs)
        elif chart_type == 'line':
            return px.line(data, **kwargs)
        elif chart_type == 'area':
            return px.area(data, **kwargs)
    except Exception as e:
        st.error(f"Ошибка создания графика: {e}")
        return None


def save_preset():
    """Сохраняет текущие фильтры как пресет"""
    preset_name = st.session_state.get('new_preset_name', '').strip()
    
    if preset_name and preset_name not in st.session_state.filter_presets:
        st.session_state.filter_presets[preset_name] = st.session_state.current_filters.copy()
        st.success(f"✅ Пресет '{preset_name}' сохранен!")
        st.session_state.new_preset_name = ""
    elif preset_name in st.session_state.filter_presets:
        st.warning(f"⚠️ Пресет '{preset_name}' уже существует!")
    else:
        st.warning("⚠️ Введите название пресета!")


def load_preset(preset_name):
    """Загружает пресет фильтров"""
    if preset_name in st.session_state.filter_presets:
        st.session_state.current_filters = st.session_state.filter_presets[preset_name].copy()
        st.success(f"✅ Пресет '{preset_name}' загружен!")


def delete_preset(preset_name):
    """Удаляет пресет"""
    if preset_name in st.session_state.filter_presets and preset_name != "Все данные":
        del st.session_state.filter_presets[preset_name]
        st.success(f"✅ Пресет '{preset_name}' удален!")


def export_presets():
    """Экспортирует пресеты в JSON"""
    return json.dumps(st.session_state.filter_presets, indent=2, ensure_ascii=False, default=str)


def import_presets(json_data):
    """Импортирует пресеты из JSON"""
    try:
        imported_presets = json.loads(json_data)
        st.session_state.filter_presets.update(imported_presets)
        st.success(f"✅ Импортировано {len(imported_presets)} пресетов!")
    except json.JSONDecodeError:
        st.error("❌ Ошибка: неверный формат JSON!")


def render_excel_like_filter(column_name, df, filter_key):
    """
    Рендерит Excel-подобный фильтр с чекбоксами для колонки
    
    Args:
        column_name: Название колонки для фильтрации
        df: DataFrame с данными
        filter_key: Ключ для хранения выбранных значений в session_state
    """
    # Получаем уникальные значения для колонки
    unique_values = sorted(df[column_name].unique().tolist())
    
    # Для дат конвертируем в строковый формат для отображения
    if column_name == 'Date':
        unique_values = [pd.to_datetime(d).strftime('%Y-%m-%d') if pd.notna(d) else 'N/A' 
                        for d in unique_values]
        unique_values = sorted(list(set(unique_values)))
    
    # Получаем текущие выбранные значения
    current_selection = st.session_state.table_filters.get(filter_key, [])
    
    # Создаем чекбокс "Выбрать все"
    select_all = st.checkbox(
        f"✅ Выбрать все ({len(unique_values)})",
        value=len(current_selection) == 0,
        key=f"select_all_{filter_key}"
    )
    
    if select_all:
        st.session_state.table_filters[filter_key] = []
    else:
        # Показываем чекбоксы для каждого значения
        selected_values = []
        
        # Ограничиваем высоту для большого количества значений
        max_display = 15
        if len(unique_values) > max_display:
            st.info(f"Показаны первые {max_display} из {len(unique_values)} значений")
            display_values = unique_values[:max_display]
        else:
            display_values = unique_values
        
        for value in display_values:
            if st.checkbox(
                str(value), 
                value=str(value) in [str(v) for v in current_selection],
                key=f"filter_{filter_key}_{value}"
            ):
                selected_values.append(value)
        
        st.session_state.table_filters[filter_key] = selected_values


def apply_table_filters(df):
    """Применяет фильтры таблицы к DataFrame"""
    filtered_df = df.copy()
    
    for column, selected_values in st.session_state.table_filters.items():
        if selected_values:  # Если есть выбранные значения
            if column == 'Date':
                # Конвертируем даты для фильтрации
                filtered_df['Date_str'] = pd.to_datetime(filtered_df['Date']).dt.strftime('%Y-%m-%d')
                filtered_df = filtered_df[filtered_df['Date_str'].isin(selected_values)]
                filtered_df = filtered_df.drop('Date_str', axis=1)
            else:
                filtered_df = filtered_df[filtered_df[column].isin(selected_values)]
    
    return filtered_df


def render_filter_sidebar(df):
    """Отображает боковую панель с фильтрами и пресетами"""
    st.sidebar.header("🔧 Фильтры и Пресеты")
    
    # --- СЕКЦИЯ ПРЕСЕТОВ ---
    st.sidebar.subheader("📋 Управление пресетами")
    
    # Выбор существующего пресета
    preset_names = list(st.session_state.filter_presets.keys())
    selected_preset = st.sidebar.selectbox(
        "Выберите пресет:",
        [""] + preset_names,
        format_func=lambda x: "Выберите пресет..." if x == "" else x
    )
    
    col1, col2 = st.sidebar.columns(2)
    with col1:
        if st.button("🔥 Загрузить", disabled=not selected_preset):
            load_preset(selected_preset)
            st.rerun()
    
    with col2:
        if st.button("🗑️ Удалить", disabled=not selected_preset or selected_preset == "Все данные"):
            delete_preset(selected_preset)
            st.rerun()
    
    # Создание нового пресета
    st.sidebar.text_input(
        "Название нового пресета:",
        key="new_preset_name",
        placeholder="Введите название..."
    )
    
    if st.sidebar.button("💾 Сохранить текущие фильтры"):
        save_preset()
        st.rerun()
    
    # Экспорт/Импорт пресетов
    with st.sidebar.expander("📤 Экспорт/Импорт пресетов"):
        st.download_button(
            label="📤 Экспорт пресетов (JSON)",
            data=export_presets(),
            file_name=f"filter_presets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            mime="application/json"
        )
        
        imported_json = st.text_area(
            "📥 Импорт пресетов (вставьте JSON):",
            height=100,
            placeholder='{"Мой пресет": {...}}'
        )
        
        if st.button("📥 Импортировать") and imported_json.strip():
            import_presets(imported_json)
            st.rerun()
    
    st.sidebar.markdown("---")
    
    # --- СЕКЦИЯ ФИЛЬТРОВ ---
    st.sidebar.subheader("🎛️ Фильтры")
    
    # Получаем текущие значения из session_state
    current_filters = st.session_state.current_filters
    
    # === НОВЫЙ ФИЛЬТР ПО ДАТАМ С ЧЕКБОКСАМИ ===
    st.sidebar.markdown("**📅 Период данных:**")
    
    # Получаем уникальные даты из данных
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'])
        unique_dates = sorted(df['Date'].dt.date.unique().tolist())
        
        # Опции быстрого выбора
        date_filter_mode = st.sidebar.radio(
            "Режим выбора дат:",
            ["📆 Все даты", "📋 Выбор чекбоксами", "🔍 Быстрый период"],
            index=0
        )
        
        if date_filter_mode == "📋 Выбор чекбоксами":
            st.sidebar.info(f"Всего дат: {len(unique_dates)}")
            
            # Показываем группировку по месяцам для удобства
            dates_by_month = {}
            for date in unique_dates:
                month_key = date.strftime('%Y-%m')
                if month_key not in dates_by_month:
                    dates_by_month[month_key] = []
                dates_by_month[month_key].append(date)
            
            selected_dates = []
            
            # Показываем по месяцам
            for month_key in sorted(dates_by_month.keys(), reverse=True):
                month_dates = dates_by_month[month_key]
                
                with st.sidebar.expander(f"📅 {month_key} ({len(month_dates)} дат)"):
                    # Чекбокс для выбора всего месяца
                    select_month = st.checkbox(
                        f"Выбрать весь месяц",
                        key=f"month_{month_key}"
                    )
                    
                    if select_month:
                        selected_dates.extend(month_dates)
                    else:
                        # Показываем отдельные даты
                        for date in sorted(month_dates, reverse=True):
                            if st.checkbox(
                                date.strftime('%Y-%m-%d'),
                                key=f"date_{date}"
                            ):
                                selected_dates.append(date)
            
            st.session_state.current_filters['selected_dates'] = selected_dates
            date_range = None
            
        elif date_filter_mode == "🔍 Быстрый период":
            # Опции быстрого выбора периода
            date_preset = st.sidebar.selectbox(
                "Быстрый выбор периода:",
                ["Последний месяц", "Последние 3 месяца", 
                 "Последние 6 месяцев", "Последний год", "Весь период"]
            )
            
            today = datetime.now().date()
            max_date = max(unique_dates)
            min_date = min(unique_dates)
            
            if date_preset == "Последний месяц":
                start_date = today - timedelta(days=30)
            elif date_preset == "Последние 3 месяца":
                start_date = today - timedelta(days=90)
            elif date_preset == "Последние 6 месяцев":
                start_date = today - timedelta(days=180)
            elif date_preset == "Последний год":
                start_date = today - timedelta(days=365)
            else:  # Весь период
                start_date = min_date
            
            start_date = max(start_date, min_date)
            end_date = max_date
            
            date_range = (start_date, end_date)
            st.session_state.current_filters['selected_dates'] = []
            st.sidebar.success(f"📅 {start_date} — {end_date}")
            
        else:  # Все даты
            date_range = None
            st.session_state.current_filters['selected_dates'] = []
            
    else:
        date_range = None
        st.sidebar.info("📅 Даты не найдены в данных")
    
    st.sidebar.markdown("---")
    
    # Переключатель режима фильтрации
    parallel_mode = st.sidebar.checkbox(
        "📢 Фильтр по параллелям",
        value=current_filters.get('parallel_mode', False),
        help="Включите для выбора целых параллелей (2, 3, 4... классы)"
    )
    
    if parallel_mode:
        # Фильтр по параллелям
        available_parallels = get_available_parallels(df)
        selected_parallels = st.sidebar.multiselect(
            "📊 Выберите параллели:",
            options=available_parallels,
            default=current_filters.get('parallels', []),
            help="Выберите номера параллелей (например, 10 = все 10-е классы)"
        )
        
        # Показываем какие классы включены
        if selected_parallels:
            included_classes = get_classes_by_parallels(df, selected_parallels)
            with st.sidebar.expander(f"📋 Включенные классы ({len(included_classes)})"):
                st.write(", ".join(included_classes))
        
        selected_classes = get_classes_by_parallels(df, selected_parallels)
        
    else:
        # Обычный мультиселект для классов
        all_classes = sorted(df['Class'].unique().tolist())
        selected_classes = st.sidebar.multiselect(
            "📚 Выберите классы:",
            options=all_classes,
            default=current_filters.get('classes', []),
            help="Оставьте пустым для выбора всех классов"
        )
        selected_parallels = []
    
    # Мультиселект для предметов
    all_subjects = sorted(df['Subject'].unique().tolist())
    selected_subjects = st.sidebar.multiselect(
        "📖 Выберите предметы:",
        options=all_subjects,
        default=current_filters.get('subjects', []),
        help="Оставьте пустым для выбора всех предметов"
    )
    
    # Слайдер для диапазона оценок
    min_possible = int(df['Average'].min())
    max_possible = int(df['Average'].max())
    current_range = current_filters.get('grade_range', [min_possible, max_possible])
    
    grade_range = st.sidebar.slider(
        "📊 Диапазон оценок:",
        min_value=min_possible,
        max_value=max_possible,
        value=(
            max(min_possible, current_range[0]), 
            min(max_possible, current_range[1])
        ),
        help="Выберите минимальную и максимальную оценку"
    )
    
    # Количество топ предметов
    top_n = st.sidebar.slider(
        "🏆 Количество предметов в рейтинге:",
        min_value=5,
        max_value=25,
        value=current_filters.get('top_n', 10),
        help="Количество предметов для отображения в топе"
    )
    
    # Обновляем current_filters
    st.session_state.current_filters.update({
        'classes': selected_classes,
        'parallel_mode': parallel_mode,
        'parallels': selected_parallels,
        'subjects': selected_subjects,
        'grade_range': list(grade_range),
        'top_n': top_n,
        'date_range': date_range
    })
    
    # Кнопка сброса фильтров
    if st.sidebar.button("🔄 Сбросить все фильтры"):
        st.session_state.current_filters = {
            'classes': [],
            'parallel_mode': False,
            'parallels': [],
            'subjects': [],
            'grade_range': [min_possible, max_possible],
            'top_n': 10,
            'date_range': None,
            'selected_dates': []
        }
        st.session_state.table_filters = {
            'Student': [],
            'Class': [],
            'Subject': [],
            'Date': []
        }
        st.rerun()
    
    return selected_classes, selected_subjects, grade_range, date_range, top_n


def apply_filters(df, selected_classes, selected_subjects, grade_range, date_range):
    """Применяет фильтры к данным"""
    filtered_df = df.copy()
    
    # Фильтр по классам
    if selected_classes:
        filtered_df = filtered_df[filtered_df['Class'].isin(selected_classes)]
    
    # Фильтр по предметам
    if selected_subjects:
        filtered_df = filtered_df[filtered_df['Subject'].isin(selected_subjects)]
    
    # Фильтр по диапазону оценок
    filtered_df = filtered_df[
        (filtered_df['Average'] >= grade_range[0]) &
        (filtered_df['Average'] <= grade_range[1])
    ]
    
    # Фильтр по датам (из бокового меню)
    if date_range:
        start_date, end_date = date_range
        filtered_df = filtered_df[
            (pd.to_datetime(filtered_df['Date']).dt.date >= start_date) &
            (pd.to_datetime(filtered_df['Date']).dt.date <= end_date)
        ]
    
    # Фильтр по выбранным датам (чекбоксы)
    selected_dates = st.session_state.current_filters.get('selected_dates', [])
    if selected_dates:
        filtered_df = filtered_df[
            pd.to_datetime(filtered_df['Date']).dt.date.isin(selected_dates)
        ]
    
    return filtered_df


def apply_manual_grade_filter(df, min_grade, max_grade):
    """Применяет дополнительный ручной фильтр по оценкам"""
    filtered_df = df.copy()
    
    if min_grade is not None:
        filtered_df = filtered_df[filtered_df['Average'] >= min_grade]
    
    if max_grade is not None:
        filtered_df = filtered_df[filtered_df['Average'] <= max_grade]
    
    return filtered_df


def main():
    st.title("📊 Дашборд успеваемости учеников")
    st.markdown("---")
    
    # Загружаем данные
    df = load_data()
    
    if df.empty:
        st.error("❌ Данные не загружены!")
        return
    
    # Отображаем боковую панель и получаем фильтры
    selected_classes, selected_subjects, grade_range, date_range, top_n = render_filter_sidebar(df)
    
    # Применяем фильтры
    filtered_df = apply_filters(df, selected_classes, selected_subjects, grade_range, date_range)
    
    if not filtered_df.empty:
        # === ОСНОВНЫЕ МЕТРИКИ ===
        st.header("📈 Ключевые метрики")
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            st.metric("📚 Всего записей", f"{len(filtered_df):,}")
        with col2:
            st.metric("🎓 Средний балл", f"{filtered_df['Average'].mean():.1f}")
        with col3:
            st.metric("📊 Медиана", f"{filtered_df['Average'].median():.1f}")
        with col4:
            st.metric("👥 Уникальных учеников", f"{filtered_df['Student'].nunique():,}")
        with col5:
            st.metric("📖 Уникальных предметов", f"{filtered_df['Subject'].nunique():,}")
        
        st.markdown("---")
        
        # Детальная таблица с EXCEL-ПОДОБНЫМИ ФИЛЬТРАМИ
        with st.expander("📋 Детальные данные и экспорт", expanded=True):
            st.markdown("### 🔍 Excel-подобные фильтры")
            
            # Создаем колонки для фильтров
            filter_cols = st.columns(4)
            
            with filter_cols[0]:
                with st.popover("🎓 Фильтр: Студент"):
                    render_excel_like_filter('Student', filtered_df, 'Student')
            
            with filter_cols[1]:
                with st.popover("📚 Фильтр: Класс"):
                    render_excel_like_filter('Class', filtered_df, 'Class')
            
            with filter_cols[2]:
                with st.popover("📖 Фильтр: Предмет"):
                    render_excel_like_filter('Subject', filtered_df, 'Subject')
            
            with filter_cols[3]:
                with st.popover("📅 Фильтр: Дата"):
                    render_excel_like_filter('Date', filtered_df, 'Date')
            
            # Применяем фильтры таблицы
            table_filtered_df = apply_table_filters(filtered_df)
            
            st.markdown("---")
            
            # Опции сортировки и отображения
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                sort_by = st.selectbox(
                    "Сортировать по:",
                    ["Average", "Student", "Class", "Subject", "Date"],
                    index=0
                )
            
            with col2:
                sort_order = st.selectbox(
                    "Порядок:",
                    ["По убыванию", "По возрастанию"],
                    index=0
                )
            
            with col3:
                show_rows = st.selectbox(
                    "Показать строк:",
                    [50, 100, 200, 500, "Все"],
                    index=0
                )
            
            with col4:
                show_stats = st.checkbox("Показать статистику", value=True)
            
            # Ручной фильтр по диапазону оценок
            st.markdown("**🎯 Дополнительный фильтр по диапазону оценок:**")
            col1, col2 = st.columns(2)
            
            with col1:
                manual_min_grade = st.number_input(
                    "Минимальная оценка:",
                    min_value=0.0,
                    max_value=100.0,
                    value=None,
                    step=0.1,
                    placeholder="Например, 75.0",
                    help="Оставьте пустым для отключения фильтра"
                )
            
            with col2:
                manual_max_grade = st.number_input(
                    "Максимальная оценка:",
                    min_value=0.0,
                    max_value=100.0,
                    value=None,
                    step=0.1,
                    placeholder="Например, 95.0",
                    help="Оставьте пустым для отключения фильтра"
                )
            
            display_filtered_df = apply_manual_grade_filter(table_filtered_df, manual_min_grade, manual_max_grade)
            
            ascending = sort_order == "По возрастанию"
            sorted_df = display_filtered_df.sort_values(sort_by, ascending=ascending)
            
            if show_rows != "Все":
                display_df = sorted_df.head(show_rows)
            else:
                display_df = sorted_df
            
            if show_stats:
                st.markdown("**📊 Статистика отображаемых данных:**")
                col1, col2, col3, col4, col5 = st.columns(5)
                
                with col1:
                    st.metric("Записей", len(display_df))
                with col2:
                    st.metric("Средняя оценка", f"{display_df['Average'].mean():.1f}")
                with col3:
                    st.metric("Медиана", f"{display_df['Average'].median():.1f}")
                with col4:
                    st.metric("Станд. отклонение", f"{display_df['Average'].std():.1f}")
                with col5:
                    st.metric("Диапазон", f"{display_df['Average'].min():.1f} - {display_df['Average'].max():.1f}")
                
                # Информация о применённых фильтрах
                active_filters = []
                for col, vals in st.session_state.table_filters.items():
                    if vals:
                        active_filters.append(f"{col}: {len(vals)} выбрано")
                
                if active_filters or manual_min_grade is not None or manual_max_grade is not None:
                    filter_info = []
                    if active_filters:
                        filter_info.append("Excel-фильтры: " + ", ".join(active_filters))
                    if manual_min_grade is not None:
                        filter_info.append(f"Оценка ≥ {manual_min_grade}")
                    if manual_max_grade is not None:
                        filter_info.append(f"Оценка ≤ {manual_max_grade}")
                    
                    st.info(f"🎯 Применены фильтры: {' | '.join(filter_info)}. "
                           f"Показано {len(display_filtered_df)} из {len(filtered_df)} записей.")
            
            st.dataframe(
                display_df,
                use_container_width=True,
                height=400
            )
            
            # Кнопки для скачивания
            st.markdown("**📥 Экспорт данных:**")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                csv = display_filtered_df.to_csv(index=False)
                st.download_button(
                    label="📄 Полные данные (CSV)",
                    data=csv,
                    file_name=f'grades_data_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
                    mime='text/csv'
                )
            
            with col2:
                summary_stats = display_filtered_df.groupby(['Class', 'Subject'])['Average'].agg(['mean', 'count', 'std']).round(2)
                st.download_button(
                    label="📊 Сводная статистика (CSV)",
                    data=summary_stats.to_csv(),
                    file_name=f'summary_stats_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
                    mime='text/csv'
                )
            
            with col3:
                if len(display_filtered_df) > 0:
                    subject_avg = display_filtered_df.groupby('Subject')['Average'].agg(['mean', 'count']).round(2)
                    subject_ranking = subject_avg.reset_index()
                    subject_ranking.columns = ['Предмет', 'Средняя_оценка', 'Количество_оценок']
                    st.download_button(
                        label="🏆 Рейтинг предметов (CSV)",
                        data=subject_ranking.to_csv(index=False),
                        file_name=f'subject_ranking_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
                        mime='text/csv'
                    )
    
    else:
        st.warning("⚠️ Нет данных для выбранных фильтров!")
        st.info("💡 Попробуйте:")
        st.markdown("""
        - Изменить критерии фильтрации
        - Загрузить пресет 'Все данные'
        - Расширить диапазон оценок
        - Выбрать другие параллели или классы
        - Расширить диапазон дат
        """)
    
    # Дополнительная информация в сайдбаре
    with st.sidebar:
        st.markdown("---")
        st.markdown("### 📈 Статистика")
        if len(filtered_df) > 0:
            st.write(f"**Записей:** {len(filtered_df):,}")
            st.write(f"**Медиана:** {filtered_df['Average'].median():.1f}")
            st.write(f"**Станд. отклонение:** {filtered_df['Average'].std():.1f}")
            st.write(f"**Мин. оценка:** {filtered_df['Average'].min():.1f}")
            st.write(f"**Макс. оценка:** {filtered_df['Average'].max():.1f}")
        
        # Информация о параллелях
        if not filtered_df.empty:
            st.markdown("### 📢 Доступные параллели")
            available_parallels = get_available_parallels(df)
            st.write(", ".join(available_parallels))


if __name__ == "__main__":
    main()
