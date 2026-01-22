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
        "date_range": None
    }


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
    
    # === ФИЛЬТР ПО ДАТАМ ===
    st.sidebar.markdown("**📅 Период данных:**")
    
    # Получаем диапазон дат из данных
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'])
        min_date = df['Date'].min().date()
        max_date = df['Date'].max().date()
        
        # Опции быстрого выбора периода
        date_preset = st.sidebar.selectbox(
            "Быстрый выбор периода:",
            ["Произвольный период", "Последний месяц", "Последние 3 месяца", 
             "Последние 6 месяцев", "Последний год", "Весь период"]
        )
        
        today = datetime.now().date()
        
        if date_preset == "Последний месяц":
            default_start = today - timedelta(days=30)
            default_end = today
        elif date_preset == "Последние 3 месяца":
            default_start = today - timedelta(days=90)
            default_end = today
        elif date_preset == "Последние 6 месяцев":
            default_start = today - timedelta(days=180)
            default_end = today
        elif date_preset == "Последний год":
            default_start = today - timedelta(days=365)
            default_end = today
        elif date_preset == "Весь период":
            default_start = min_date
            default_end = max_date
        else:  # Произвольный период
            default_start = min_date
            default_end = max_date
        
        # Ограничиваем даты доступным диапазоном
        default_start = max(default_start, min_date)
        default_end = min(default_end, max_date)
        
        col1, col2 = st.sidebar.columns(2)
        with col1:
            start_date = st.date_input(
                "С:",
                value=default_start,
                min_value=min_date,
                max_value=max_date
            )
        with col2:
            end_date = st.date_input(
                "По:",
                value=default_end,
                min_value=min_date,
                max_value=max_date
            )
        
        date_range = (start_date, end_date)
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
    st.session_state.current_filters = {
        'classes': selected_classes,
        'parallel_mode': parallel_mode,
        'parallels': selected_parallels,
        'subjects': selected_subjects,
        'grade_range': list(grade_range),
        'top_n': top_n,
        'date_range': date_range
    }
    
    # Кнопка сброса фильтров
    if st.sidebar.button("🔄 Сбросить все фильтры"):
        st.session_state.current_filters = {
            'classes': [],
            'parallel_mode': False,
            'parallels': [],
            'subjects': [],
            'grade_range': [min_possible, max_possible],
            'top_n': 10,
            'date_range': None
        }
        st.rerun()
    
    return selected_classes, selected_subjects, grade_range, top_n, parallel_mode, selected_parallels, date_range


def apply_filters(df, selected_classes, selected_subjects, grade_range, date_range=None):
    """Применяет фильтры к данным"""
    filtered_df = df.copy()
    
    # Фильтр по датам
    if date_range and 'Date' in filtered_df.columns:
        start_date, end_date = date_range
        filtered_df['Date'] = pd.to_datetime(filtered_df['Date'])
        filtered_df = filtered_df[
            (filtered_df['Date'].dt.date >= start_date) & 
            (filtered_df['Date'].dt.date <= end_date)
        ]
    
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
    
    return filtered_df


def apply_manual_grade_filter(df, min_grade, max_grade):
    """Применяет ручной фильтр по диапазону оценок"""
    if min_grade is not None and max_grade is not None:
        return df[(df['Average'] >= min_grade) & (df['Average'] <= max_grade)]
    elif min_grade is not None:
        return df[df['Average'] >= min_grade]
    elif max_grade is not None:
        return df[df['Average'] <= max_grade]
    return df


def render_filter_summary(selected_classes, selected_subjects, grade_range, original_df, filtered_df, parallel_mode, selected_parallels, date_range):
    """Отображает сводку примененных фильтров"""
    with st.expander("🔍 Примененные фильтры", expanded=False):
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            if parallel_mode and selected_parallels:
                st.write("**📢 Параллели:**")
                st.write(f"• {', '.join(selected_parallels)} классы")
                st.write(f"• *({len(selected_classes)} классов)*")
            elif selected_classes:
                st.write("**📚 Классы:**")
                if len(selected_classes) <= 3:
                    st.write(f"• {', '.join(selected_classes)}")
                else:
                    st.write(f"• {len(selected_classes)} классов")
            else:
                st.write("**📚 Классы:**")
                st.write("• Все классы")
        
        with col2:
            st.write("**📖 Предметы:**")
            if selected_subjects:
                if len(selected_subjects) <= 3:
                    st.write(f"• {', '.join(selected_subjects)}")
                else:
                    st.write(f"• {len(selected_subjects)} предметов")
            else:
                st.write("• Все предметы")
        
        with col3:
            st.write("**📊 Диапазон оценок:**")
            st.write(f"• {grade_range[0]} - {grade_range[1]}")
        
        with col4:
            st.write("**📅 Период:**")
            if date_range:
                st.write(f"• {date_range[0]} — {date_range[1]}")
            else:
                st.write("• Весь период")
        
        # Статистика фильтрации
        original_count = len(original_df)
        filtered_count = len(filtered_df)
        percentage = (filtered_count / original_count * 100) if original_count > 0 else 0
        
        st.info(f"📈 Отображено {filtered_count:,} из {original_count:,} записей ({percentage:.1f}%)")


def create_subject_ranking_charts(filtered_df, top_n):
    """Создает графики рейтинга лучших и худших предметов"""
    subject_avg = filtered_df.groupby('Subject')['Average'].agg(['mean', 'count']).reset_index()
    subject_avg['mean'] = subject_avg['mean'].round(1)
    subject_avg = subject_avg.sort_values('mean', ascending=False)
    
    # Топ лучших и худших
    top_best = subject_avg.head(top_n)
    top_worst = subject_avg.tail(top_n).sort_values('mean', ascending=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader(f"🏆 Топ-{top_n} лучших предметов")
        
        if PLOTLY_AVAILABLE and len(top_best) > 0:
            fig_best = create_plotly_chart(
                'bar',
                top_best,
                x='mean',
                y='Subject',
                orientation='h',
                color='mean',
                color_continuous_scale='Greens',
                labels={'mean': 'Средняя оценка', 'Subject': 'Предмет'},
                text='mean'
            )
            if fig_best:
                fig_best.update_layout(
                    height=400,
                    yaxis={'categoryorder': 'total ascending'},
                    showlegend=False
                )
                fig_best.update_traces(texttemplate='%{text:.1f}', textposition='outside')
                st.plotly_chart(fig_best, use_container_width=True)
        else:
            st.dataframe(
                top_best[['Subject', 'mean', 'count']].rename(columns={
                    'Subject': 'Предмет',
                    'mean': 'Средняя оценка', 
                    'count': 'Количество'
                }),
                use_container_width=True
            )
    
    with col2:
        st.subheader(f"⚠️ Топ-{top_n} предметов для внимания")
        
        if PLOTLY_AVAILABLE and len(top_worst) > 0:
            fig_worst = create_plotly_chart(
                'bar',
                top_worst,
                x='mean',
                y='Subject',
                orientation='h',
                color='mean',
                color_continuous_scale='Reds_r',
                labels={'mean': 'Средняя оценка', 'Subject': 'Предмет'},
                text='mean'
            )
            if fig_worst:
                fig_worst.update_layout(
                    height=400,
                    yaxis={'categoryorder': 'total ascending'},
                    showlegend=False
                )
                fig_worst.update_traces(texttemplate='%{text:.1f}', textposition='outside')
                st.plotly_chart(fig_worst, use_container_width=True)
        else:
            st.dataframe(
                top_worst[['Subject', 'mean', 'count']].rename(columns={
                    'Subject': 'Предмет',
                    'mean': 'Средняя оценка', 
                    'count': 'Количество'
                }),
                use_container_width=True
            )
    
    return subject_avg


def create_trend_charts(filtered_df):
    """Создает графики трендов по времени"""
    if 'Date' not in filtered_df.columns or filtered_df.empty:
        st.info("📅 Нет данных о датах для построения трендов")
        return
    
    st.subheader("📈 Анализ трендов")
    
    # Преобразуем даты
    df_trends = filtered_df.copy()
    df_trends['Date'] = pd.to_datetime(df_trends['Date'])
    
    # Определяем оптимальную группировку по дате
    date_range_days = (df_trends['Date'].max() - df_trends['Date'].min()).days
    
    if date_range_days <= 31:
        freq = 'D'
        freq_label = 'день'
    elif date_range_days <= 90:
        freq = 'W'
        freq_label = 'неделя'
    else:
        freq = 'M'
        freq_label = 'месяц'
    
    # Выбор типа тренда
    trend_type = st.radio(
        "Выберите тип анализа тренда:",
        ["Общий тренд", "По предметам", "По классам", "По параллелям"],
        horizontal=True
    )
    
    if trend_type == "Общий тренд":
        # Агрегация по периодам
        df_trends['Period'] = df_trends['Date'].dt.to_period(freq).dt.to_timestamp()
        trend_data = df_trends.groupby('Period').agg({
            'Average': ['mean', 'std', 'count']
        }).reset_index()
        trend_data.columns = ['Period', 'Mean', 'Std', 'Count']
        trend_data['Mean'] = trend_data['Mean'].round(1)
        
        if PLOTLY_AVAILABLE and len(trend_data) > 1:
            # Создаем график с трендлайном и доверительным интервалом
            fig = go.Figure()
            
            # Основная линия
            fig.add_trace(go.Scatter(
                x=trend_data['Period'],
                y=trend_data['Mean'],
                mode='lines+markers',
                name='Средний балл',
                line=dict(color='#636EFA', width=3),
                marker=dict(size=8)
            ))
            
            # Добавляем скользящее среднее (тренд)
            if len(trend_data) >= 3:
                trend_data['MA'] = trend_data['Mean'].rolling(window=3, min_periods=1).mean()
                fig.add_trace(go.Scatter(
                    x=trend_data['Period'],
                    y=trend_data['MA'],
                    mode='lines',
                    name='Скользящее среднее (тренд)',
                    line=dict(color='#EF553B', width=2, dash='dash')
                ))
            
            # Доверительный интервал (±1 std)
            if trend_data['Std'].notna().any():
                fig.add_trace(go.Scatter(
                    x=pd.concat([trend_data['Period'], trend_data['Period'][::-1]]),
                    y=pd.concat([trend_data['Mean'] + trend_data['Std'], 
                                (trend_data['Mean'] - trend_data['Std'])[::-1]]),
                    fill='toself',
                    fillcolor='rgba(99, 110, 250, 0.2)',
                    line=dict(color='rgba(255,255,255,0)'),
                    name='±1 станд. отклонение',
                    showlegend=True
                ))
            
            fig.update_layout(
                title=f'Динамика среднего балла (группировка по {freq_label})',
                xaxis_title='Период',
                yaxis_title='Средний балл',
                height=450,
                hovermode='x unified'
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Статистика тренда
            col1, col2, col3, col4 = st.columns(4)
            
            first_period = trend_data['Mean'].iloc[0]
            last_period = trend_data['Mean'].iloc[-1]
            change = last_period - first_period
            pct_change = (change / first_period * 100) if first_period > 0 else 0
            
            with col1:
                st.metric("Начало периода", f"{first_period:.1f}")
            with col2:
                st.metric("Конец периода", f"{last_period:.1f}")
            with col3:
                st.metric("Изменение", f"{change:+.1f}", delta=f"{pct_change:+.1f}%")
            with col4:
                trend_direction = "📈 Рост" if change > 0.5 else "📉 Снижение" if change < -0.5 else "➡️ Стабильно"
                st.metric("Тренд", trend_direction)
        else:
            st.dataframe(trend_data)
    
    elif trend_type == "По предметам":
        # Выбор предметов для сравнения
        available_subjects = df_trends['Subject'].unique().tolist()
        
        default_subjects = available_subjects[:5] if len(available_subjects) > 5 else available_subjects
        selected_subjects_trend = st.multiselect(
            "Выберите предметы для сравнения:",
            options=available_subjects,
            default=default_subjects,
            max_selections=8
        )
        
        if selected_subjects_trend:
            df_subjects = df_trends[df_trends['Subject'].isin(selected_subjects_trend)]
            df_subjects['Period'] = df_subjects['Date'].dt.to_period(freq).dt.to_timestamp()
            
            trend_by_subject = df_subjects.groupby(['Period', 'Subject'])['Average'].mean().reset_index()
            trend_by_subject['Average'] = trend_by_subject['Average'].round(1)
            
            if PLOTLY_AVAILABLE and len(trend_by_subject) > 0:
                fig = create_plotly_chart(
                    'line',
                    trend_by_subject,
                    x='Period',
                    y='Average',
                    color='Subject',
                    markers=True,
                    labels={'Average': 'Средний балл', 'Period': 'Период', 'Subject': 'Предмет'}
                )
                if fig:
                    fig.update_layout(
                        title=f'Динамика по предметам (группировка по {freq_label})',
                        height=500,
                        hovermode='x unified'
                    )
                    st.plotly_chart(fig, use_container_width=True)
    
    elif trend_type == "По классам":
        # Выбор классов для сравнения
        available_classes = df_trends['Class'].unique().tolist()
        
        default_classes = available_classes[:5] if len(available_classes) > 5 else available_classes
        selected_classes_trend = st.multiselect(
            "Выберите классы для сравнения:",
            options=available_classes,
            default=default_classes,
            max_selections=8
        )
        
        if selected_classes_trend:
            df_classes = df_trends[df_trends['Class'].isin(selected_classes_trend)]
            df_classes['Period'] = df_classes['Date'].dt.to_period(freq).dt.to_timestamp()
            
            trend_by_class = df_classes.groupby(['Period', 'Class'])['Average'].mean().reset_index()
            trend_by_class['Average'] = trend_by_class['Average'].round(1)
            
            if PLOTLY_AVAILABLE and len(trend_by_class) > 0:
                fig = create_plotly_chart(
                    'line',
                    trend_by_class,
                    x='Period',
                    y='Average',
                    color='Class',
                    markers=True,
                    labels={'Average': 'Средний балл', 'Period': 'Период', 'Class': 'Класс'}
                )
                if fig:
                    fig.update_layout(
                        title=f'Динамика по классам (группировка по {freq_label})',
                        height=500,
                        hovermode='x unified'
                    )
                    st.plotly_chart(fig, use_container_width=True)
    
    else:  # По параллелям
        df_trends['Parallel'] = df_trends['Class'].apply(extract_parallel_from_class)
        df_trends['Period'] = df_trends['Date'].dt.to_period(freq).dt.to_timestamp()
        
        trend_by_parallel = df_trends.groupby(['Period', 'Parallel'])['Average'].mean().reset_index()
        trend_by_parallel['Average'] = trend_by_parallel['Average'].round(1)
        
        # Сортируем параллели по числовому значению
        trend_by_parallel['Parallel_num'] = trend_by_parallel['Parallel'].apply(
            lambda x: int(x) if x.isdigit() else 999
        )
        trend_by_parallel = trend_by_parallel.sort_values(['Period', 'Parallel_num'])
        
        if PLOTLY_AVAILABLE and len(trend_by_parallel) > 0:
            fig = create_plotly_chart(
                'line',
                trend_by_parallel,
                x='Period',
                y='Average',
                color='Parallel',
                markers=True,
                labels={'Average': 'Средний балл', 'Period': 'Период', 'Parallel': 'Параллель'}
            )
            if fig:
                fig.update_layout(
                    title=f'Динамика по параллелям (группировка по {freq_label})',
                    height=500,
                    hovermode='x unified'
                )
                st.plotly_chart(fig, use_container_width=True)


def create_heatmap_calendar(filtered_df):
    """Создает тепловую карту по дням/неделям"""
    if 'Date' not in filtered_df.columns or filtered_df.empty:
        return
    
    st.subheader("🗓️ Тепловая карта активности")
    
    df_heat = filtered_df.copy()
    df_heat['Date'] = pd.to_datetime(df_heat['Date'])
    df_heat['DayOfWeek'] = df_heat['Date'].dt.day_name()
    df_heat['Week'] = df_heat['Date'].dt.isocalendar().week
    
    # Агрегация по дню недели и неделе
    heat_data = df_heat.groupby(['Week', 'DayOfWeek'])['Average'].mean().reset_index()
    heat_pivot = heat_data.pivot(index='DayOfWeek', columns='Week', values='Average')
    
    # Сортируем дни недели
    day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    heat_pivot = heat_pivot.reindex([d for d in day_order if d in heat_pivot.index])
    
    if PLOTLY_AVAILABLE and not heat_pivot.empty:
        fig = px.imshow(
            heat_pivot.values,
            x=heat_pivot.columns.tolist(),
            y=heat_pivot.index.tolist(),
            color_continuous_scale='RdYlGn',
            labels={'x': 'Неделя', 'y': 'День недели', 'color': 'Ср. балл'},
            aspect='auto'
        )
        fig.update_layout(
            title='Средний балл по дням недели и неделям года',
            height=300
        )
        st.plotly_chart(fig, use_container_width=True)


def main():
    st.title("📊 Дашборд успеваемости школы")
    st.markdown("*Интерактивная аналитика с фильтрами по датам, параллелям и анализом трендов*")
    st.markdown("---")
    
    # Загрузка демо-данных
    df = create_demo_data()
    
    if df is None or len(df) == 0:
        st.error("Не удалось загрузить данные")
        return
    
    # Отображение фильтров и получение выбранных значений
    selected_classes, selected_subjects, grade_range, top_n, parallel_mode, selected_parallels, date_range = render_filter_sidebar(df)
    
    # Применение фильтров
    filtered_df = apply_filters(df, selected_classes, selected_subjects, grade_range, date_range)
    
    # Отображение сводки фильтров
    render_filter_summary(selected_classes, selected_subjects, grade_range, df, filtered_df, parallel_mode, selected_parallels, date_range)
    
    # Основная статистика
    if len(filtered_df) > 0:
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                label="👥 Студентов",
                value=filtered_df['Student'].nunique(),
                delta=f"{filtered_df['Student'].nunique() - df['Student'].nunique():+d}"
            )
        
        with col2:
            st.metric(
                label="🏫 Классов",
                value=filtered_df['Class'].nunique(),
                delta=f"{filtered_df['Class'].nunique() - df['Class'].nunique():+d}"
            )
        
        with col3:
            st.metric(
                label="📚 Предметов",
                value=filtered_df['Subject'].nunique(),
                delta=f"{filtered_df['Subject'].nunique() - df['Subject'].nunique():+d}"
            )
        
        with col4:
            avg_score = filtered_df['Average'].mean()
            total_avg = df['Average'].mean()
            delta = avg_score - total_avg
            st.metric(
                label="📈 Средний балл",
                value=f"{avg_score:.1f}",
                delta=f"{delta:+.1f}"
            )
        
        st.markdown("---")
        
        # === НОВЫЙ РАЗДЕЛ: ТРЕНДЫ ===
        create_trend_charts(filtered_df)
        
        st.markdown("---")
        
        # Рейтинг предметов (лучшие и худшие)
        subject_avg = create_subject_ranking_charts(filtered_df, top_n)
        
        st.markdown("---")
        
        # Остальные графики
        col1, col2 = st.columns(2)
        
        with col1:
            # Распределение оценок
            st.subheader("📊 Распределение оценок")
            if PLOTLY_AVAILABLE:
                fig_dist = create_plotly_chart(
                    'histogram',
                    filtered_df,
                    x='Average',
                    nbins=20,
                    color_discrete_sequence=['#636EFA']
                )
                if fig_dist:
                    fig_dist.add_vline(
                        x=filtered_df['Average'].mean(),
                        line_dash="dash",
                        line_color="red",
                        annotation_text=f"Среднее: {filtered_df['Average'].mean():.1f}"
                    )
                    st.plotly_chart(fig_dist, use_container_width=True)
            else:
                st.write("Статистика распределения:")
                st.write(filtered_df['Average'].describe())
        
        with col2:
            # Сравнение параллелей (если применен фильтр по параллелям)
            if parallel_mode and selected_parallels and len(selected_parallels) > 1:
                st.subheader("📢 Сравнение параллелей")
                
                filtered_df_with_parallel = filtered_df.copy()
                filtered_df_with_parallel['Parallel'] = filtered_df_with_parallel['Class'].apply(extract_parallel_from_class)
                
                parallel_stats = filtered_df_with_parallel.groupby('Parallel')['Average'].agg(['mean', 'count']).reset_index()
                parallel_stats['mean'] = parallel_stats['mean'].round(1)
                
                if PLOTLY_AVAILABLE:
                    fig_parallels = create_plotly_chart(
                        'bar',
                        parallel_stats,
                        x='Parallel',
                        y='mean',
                        color='mean',
                        color_continuous_scale='Viridis',
                        labels={'mean': 'Средняя оценка', 'Parallel': 'Параллель'},
                        text='mean'
                    )
                    if fig_parallels:
                        fig_parallels.update_traces(texttemplate='%{text:.1f}', textposition='outside')
                        st.plotly_chart(fig_parallels, use_container_width=True)
                else:
                    st.dataframe(parallel_stats.rename(columns={'Parallel': 'Параллель', 'mean': 'Средняя оценка'}))
            else:
                # Статистика по категориям оценок
                st.subheader("📈 Категории успеваемости")
                
                bins = [0, 40, 65, 85, 100]
                labels = ['Неуд. (0-39)', 'Удовл. (40-64)', 'Хорошо (65-84)', 'Отлично (85-100)']
                filtered_df_copy = filtered_df.copy()
                filtered_df_copy['Grade_Category'] = pd.cut(filtered_df_copy['Average'], bins=bins, labels=labels, include_lowest=True)
                
                category_counts = filtered_df_copy['Grade_Category'].value_counts()
                
                if PLOTLY_AVAILABLE and len(category_counts) > 0:
                    fig_pie = go.Figure(data=[go.Pie(
                        labels=category_counts.index,
                        values=category_counts.values,
                        hole=0.3,
                        marker_colors=['#ff4444', '#ff8800', '#88dd00', '#44dd44']
                    )])
                    fig_pie.update_layout(height=400)
                    st.plotly_chart(fig_pie, use_container_width=True)
                else:
                    st.dataframe(category_counts.reset_index().rename(columns={'index': 'Категория', 'Grade_Category': 'Количество'}))
        
        # Сравнение классов
        if len(filtered_df['Class'].unique()) > 1:
            st.subheader("🏫 Анализ по классам")
            
            class_avg = filtered_df.groupby('Class')['Average'].agg(['mean', 'count']).reset_index()
            class_avg['mean'] = class_avg['mean'].round(1)
            class_avg = class_avg.sort_values('mean', ascending=False)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**🏆 Лучшие классы:**")
                top_classes = class_avg.head(5)
                for idx, row in top_classes.iterrows():
                    st.write(f"• {row['Class']}: {row['mean']:.1f} ({row['count']} оценок)")
            
            with col2:
                st.write("**⚠️ Классы для внимания:**")
                bottom_classes = class_avg.tail(5)
                for idx, row in bottom_classes.iterrows():
                    st.write(f"• {row['Class']}: {row['mean']:.1f} ({row['count']} оценок)")
            
            if PLOTLY_AVAILABLE and len(class_avg) > 0:
                fig_classes = create_plotly_chart(
                    'scatter',
                    class_avg,
                    x='Class',
                    y='mean',
                    size='count',
                    color='mean',
                    color_continuous_scale='RdYlGn',
                    labels={'mean': 'Средняя оценка', 'count': 'Количество оценок'},
                    hover_data=['count'],
                    title="Успеваемость по классам"
                )
                if fig_classes:
                    fig_classes.update_layout(height=400)
                    st.plotly_chart(fig_classes, use_container_width=True)
        
        # Box plot анализ
        if len(filtered_df) > 50:
            st.subheader("📈 Детальный анализ распределения")
            
            analysis_type = st.radio(
                "Выберите тип анализа:",
                ["По классам", "По предметам", "По параллелям"] if parallel_mode and len(selected_parallels) > 1 else ["По классам", "По предметам"],
                horizontal=True
            )
            
            if PLOTLY_AVAILABLE:
                if analysis_type == "По классам":
                    top_classes_for_box = filtered_df.groupby('Class')['Average'].count().nlargest(15).index
                    df_for_box = filtered_df[filtered_df['Class'].isin(top_classes_for_box)]
                    
                    fig_box = create_plotly_chart(
                        'box',
                        df_for_box,
                        x='Class',
                        y='Average',
                        labels={'Average': 'Оценка', 'Class': 'Класс'}
                    )
                elif analysis_type == "По предметам":
                    top_subjects_for_box = filtered_df.groupby('Subject')['Average'].count().nlargest(12).index
                    df_for_box = filtered_df[filtered_df['Subject'].isin(top_subjects_for_box)]
                    
                    fig_box = create_plotly_chart(
                        'box',
                        df_for_box,
                        x='Subject',
                        y='Average',
                        labels={'Average': 'Оценка', 'Subject': 'Предмет'}
                    )
                else:
                    filtered_df_with_parallel = filtered_df.copy()
                    filtered_df_with_parallel['Parallel'] = filtered_df_with_parallel['Class'].apply(extract_parallel_from_class)
                    
                    fig_box = create_plotly_chart(
                        'box',
                        filtered_df_with_parallel,
                        x='Parallel',
                        y='Average',
                        labels={'Average': 'Оценка', 'Parallel': 'Параллель'}
                    )
                
                if fig_box:
                    fig_box.update_layout(height=450)
                    fig_box.update_xaxes(tickangle=-45)
                    st.plotly_chart(fig_box, use_container_width=True)
            else:
                if analysis_type == "По классам":
                    box_stats = filtered_df.groupby('Class')['Average'].agg(['min', 'max', 'mean', 'median']).round(1)
                elif analysis_type == "По предметам":
                    box_stats = filtered_df.groupby('Subject')['Average'].agg(['min', 'max', 'mean', 'median']).round(1)
                else:
                    filtered_df_with_parallel = filtered_df.copy()
                    filtered_df_with_parallel['Parallel'] = filtered_df_with_parallel['Class'].apply(extract_parallel_from_class)
                    box_stats = filtered_df_with_parallel.groupby('Parallel')['Average'].agg(['min', 'max', 'mean', 'median']).round(1)
                
                st.dataframe(box_stats, use_container_width=True)
        
        # Детальная таблица с улучшенными опциями
        with st.expander("📋 Детальные данные и экспорт"):
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
            
            display_filtered_df = apply_manual_grade_filter(filtered_df, manual_min_grade, manual_max_grade)
            
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
                
                if manual_min_grade is not None or manual_max_grade is not None:
                    filter_info = []
                    if manual_min_grade is not None:
                        filter_info.append(f"≥ {manual_min_grade}")
                    if manual_max_grade is not None:
                        filter_info.append(f"≤ {manual_max_grade}")
                    
                    st.info(f"🎯 Применен дополнительный фильтр оценок: {' и '.join(filter_info)}. "
                           f"Отфильтровано {len(display_filtered_df)} из {len(filtered_df)} записей.")
            
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
                if 'subject_avg' in locals():
                    subject_ranking = subject_avg[['Subject', 'mean', 'count']].round(2)
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
