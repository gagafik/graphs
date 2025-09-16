import streamlit as st
import pandas as pd
import numpy as np
import json
from datetime import datetime
import re

# Проверяем и импортируем Plotly
try:
    import plotly.express as px
    import plotly.graph_objects as go
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
            "stats_grade_range": [0, 100],
            "top_n": 10
        },
        "2-е параллель": {
            "classes": [],
            "parallel_mode": True,
            "parallels": ["2"],
            "subjects": [],
            "grade_range": [0, 100],
            "stats_grade_range": [0, 100],
            "top_n": 10
        },
        "Старшие классы (10-11)": {
            "classes": [],
            "parallel_mode": True,
            "parallels": ["10", "11"],
            "subjects": [],
            "grade_range": [0, 100],
            "stats_grade_range": [0, 100],
            "top_n": 15
        },
        "Средние классы (7-9)": {
            "classes": [],
            "parallel_mode": True,
            "parallels": ["7", "8", "9"],
            "subjects": [],
            "grade_range": [0, 100],
            "stats_grade_range": [0, 100],
            "top_n": 12
        },
        "Точные науки": {
            "classes": [],
            "parallel_mode": False,
            "parallels": [],
            "subjects": ["Math", "Physics", "Chemistry", "CS", "Calc", "Further Math"],
            "grade_range": [0, 100],
            "stats_grade_range": [0, 100],
            "top_n": 6
        },
        "Языки": {
            "classes": [],
            "parallel_mode": False,
            "parallels": [],
            "subjects": ["English", "ESL", "Rus", "Kaz", "RusLit", "KazLit"],
            "grade_range": [0, 100],
            "stats_grade_range": [0, 100],
            "top_n": 8
        }
    }

if 'current_filters' not in st.session_state:
    st.session_state.current_filters = {
        "classes": [],
        "parallel_mode": False,
        "parallels": [],
        "subjects": [],
        "grade_range": [0, 100],
        "stats_grade_range": [0, 100],
        "top_n": 10
    }

@st.cache_data
def load_data():
    """Загружает и кэширует данные"""
    try:
        possible_files = [
            'Marks 2526.xlsx',
            'marks_2526.xlsx', 
            'data/Marks 2526.xlsx',
            'data/marks_2526.xlsx'
        ]
        
        df = None
        for file_path in possible_files:
            try:
                df = pd.read_excel(file_path, sheet_name='Average year, no teacher')
                st.success(f"✅ Данные загружены из {file_path}")
                break
            except FileNotFoundError:
                continue
        
        if df is None:
            st.warning("⚠️ Файл Excel не найден. Используются демо-данные.")
            df = create_demo_data()
        
        return df
        
    except Exception as e:
        st.error(f"Ошибка загрузки данных: {e}")
        st.info("Используются демо-данные для демонстрации")
        return create_demo_data()

def create_demo_data():
    """Создает демо-данные для демонстрации"""
    np.random.seed(42)
    
    students = [f"Студент_{i}" for i in range(1, 101)]
    classes = ['2A', '2B', '3A', '3B', '4A', '4B', '5A', '5B', '6A', '6B', 
               '7A', '7B', '8A', '8B', '9A', '9B', '10A', '10B', '11A', '11B']
    subjects = ['Math', 'Physics', 'Chemistry', 'Biology', 'English', 'History', 
                'Geography', 'Literature', 'CS', 'Art', 'PE', 'Music']
    
    data = []
    for student in students:
        class_name = np.random.choice(classes)
        for subject in np.random.choice(subjects, size=np.random.randint(6, 10), replace=False):
            # Создаем более реалистичное распределение оценок
            base_score = 75
            if subject in ['Math', 'Physics', 'Chemistry']:
                base_score = 70  # Точные науки чуть сложнее
            elif subject in ['Art', 'PE', 'Music']:
                base_score = 85  # Творческие предметы легче
            
            score = np.random.normal(base_score, 12)
            score = max(30, min(100, score))
            data.append({
                'Student': student,
                'Class': class_name,
                'Subject': subject,
                'Average': round(score, 1)
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
    return json.dumps(st.session_state.filter_presets, indent=2, ensure_ascii=False)

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
    
    # Слайдер для диапазона оценок (основной фильтр)
    min_possible = int(df['Average'].min())
    max_possible = int(df['Average'].max())
    current_range = current_filters.get('grade_range', [min_possible, max_possible])
    
    grade_range = st.sidebar.slider(
        "📊 Диапазон оценок (фильтр данных):",
        min_value=min_possible,
        max_value=max_possible,
        value=(
            max(min_possible, current_range[0]), 
            min(max_possible, current_range[1])
        ),
        help="Фильтрует данные по указанному диапазону оценок"
    )
    
    # Новый слайдер для статистики
    st.sidebar.markdown("---")
    st.sidebar.subheader("📈 Статистика отображаемых данных")
    
    current_stats_range = current_filters.get('stats_grade_range', [min_possible, max_possible])
    
    stats_grade_range = st.sidebar.slider(
        "📊 Диапазон для статистики:",
        min_value=min_possible,
        max_value=max_possible,
        value=(
            max(min_possible, current_stats_range[0]), 
            min(max_possible, current_stats_range[1])
        ),
        help="Показывает статистику только для оценок в указанном диапазоне"
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
        'stats_grade_range': list(stats_grade_range),
        'top_n': top_n
    }
    
    # Кнопка сброса фильтров
    if st.sidebar.button("🔄 Сбросить все фильтры"):
        st.session_state.current_filters = {
            'classes': [],
            'parallel_mode': False,
            'parallels': [],
            'subjects': [],
            'grade_range': [min_possible, max_possible],
            'stats_grade_range': [min_possible, max_possible],
            'top_n': 10
        }
        st.rerun()
    
    return selected_classes, selected_subjects, grade_range, stats_grade_range, top_n, parallel_mode, selected_parallels

def apply_filters(df, selected_classes, selected_subjects, grade_range):
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
    
    return filtered_df

def render_filter_summary(selected_classes, selected_subjects, grade_range, stats_grade_range, original_df, filtered_df, parallel_mode, selected_parallels):
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
            st.write("**📊 Диапазон оценок (фильтр):**")
            st.write(f"• {grade_range[0]} - {grade_range[1]}")
        
        with col4:
            st.write("**📈 Диапазон для статистики:**")
            st.write(f"• {stats_grade_range[0]} - {stats_grade_range[1]}")
        
        # Статистика фильтрации
        original_count = len(original_df)
        filtered_count = len(filtered_df)
        percentage = (filtered_count / original_count * 100) if original_count > 0 else 0
        
        st.info(f"📈 Отображено {filtered_count:,} из {original_count:,} записей ({percentage:.1f}%)")

def create_subject_ranking_charts(filtered_df, stats_grade_range, top_n):
    """Создает графики рейтинга лучших и худших предметов"""
    # Применяем фильтр статистики
    stats_df = filtered_df[
        (filtered_df['Average'] >= stats_grade_range[0]) & 
        (filtered_df['Average'] <= stats_grade_range[1])
    ]
    
    if len(stats_df) == 0:
        st.warning("⚠️ Нет данных для выбранного диапазона статистики!")
        return None
    
    subject_avg = stats_df.groupby('Subject')['Average'].agg(['mean', 'count']).reset_index()
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
                color_continuous_scale='Reds_r',  # Обратная красная шкала
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

def main():
    st.title("📊 Дашборд успеваемости школы")
    st.markdown("*Интерактивная аналитика с фильтрами по параллелям и рейтингом предметов*")
    st.markdown("---")
    
    # Загрузка данных
    df = load_data()
    
    if df is None or len(df) == 0:
        st.error("Не удалось загрузить данные")
        return
    
    # Отображение фильтров и получение выбранных значений
    selected_classes, selected_subjects, grade_range, stats_grade_range, top_n, parallel_mode, selected_parallels = render_filter_sidebar(df)
    
    # Применение фильтров
    filtered_df = apply_filters(df, selected_classes, selected_subjects, grade_range)
    
    # Отображение сводки фильтров
    render_filter_summary(selected_classes, selected_subjects, grade_range, stats_grade_range, df, filtered_df, parallel_mode, selected_parallels)
    
    # Основная статистика
    if len(filtered_df) > 0:
        # Применяем фильтр статистики для метрик
        stats_df = filtered_df[
            (filtered_df['Average'] >= stats_grade_range[0]) & 
            (filtered_df['Average'] <= stats_grade_range[1])
        ]
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                label="👥 Студентов",
                value=stats_df['Student'].nunique(),
                delta=f"{stats_df['Student'].nunique() - df['Student'].nunique():+d}"
            )
        
        with col2:
            st.metric(
                label="🏫 Классов",
                value=stats_df['Class'].nunique(),
                delta=f"{stats_df['Class'].nunique() - df['Class'].nunique():+d}"
            )
        
        with col3:
            st.metric(
                label="📚 Предметов",
                value=stats_df['Subject'].nunique(),
                delta=f"{stats_df['Subject'].nunique() - df['Subject'].nunique():+d}"
            )
        
        with col4:
            if len(stats_df) > 0:
                avg_score = stats_df['Average'].mean()
                total_avg = df['Average'].mean()
                delta = avg_score - total_avg
                st.metric(
                    label="📈 Средний балл",
                    value=f"{avg_score:.1f}",
                    delta=f"{delta:+.1f}"
                )
            else:
                st.metric(
                    label="📈 Средний балл",
                    value="Нет данных",
                    delta=None
                )
        
        st.markdown("---")
        
        # Рейтинг предметов (лучшие и худшие)
        subject_avg = create_subject_ranking_charts(filtered_df, stats_grade_range, top_n)
        
        st.markdown("---")
        
        # Остальные графики
        col1, col2 = st.columns(2)
        
        with col1:
            # Распределение оценок (используем stats_df)
            st.subheader("📊 Распределение оценок")
            if len(stats_df) > 0:
                if PLOTLY_AVAILABLE:
                    fig_dist = create_plotly_chart(
                        'histogram',
                        stats_df,
                        x='Average',
                        nbins=20,
                        color_discrete_sequence=['#636EFA']
                    )
                    if fig_dist:
                        fig_dist.add_vline(
                            x=stats_df['Average'].mean(),
                            line_dash="dash",
                            line_color="red",
                            annotation_text=f"Среднее: {stats_df['Average'].mean():.1f}"
                        )
                        st.plotly_chart(fig_dist, use_container_width=True)
                else:
                    st.write("Статистика распределения:")
                    st.write(stats_df['Average'].describe())
            else:
                st.info("Нет данных для выбранного диапазона статистики")
        
        with col2:
            # Категории успеваемости (используем stats_df)
            st.subheader("📈 Категории успеваемости")
            
            if len(stats_df) > 0:
                # Обновленные категории
                bins = [0, 40, 65, 85, 100]
                labels = ['Неуд [0;40)', 'Удовл [40;65)', 'Хорошо [65;85)', 'Отлично [85;100]']
                stats_df_copy = stats_df.copy()
                stats_df_copy['Grade_Category'] = pd.cut(
                    stats_df_copy['Average'], 
                    bins=bins, 
                    labels=labels, 
                    include_lowest=True,
                    right=False  # [a;b) интервалы
                )
                
                category_counts = stats_df_copy['Grade_Category'].value_counts()
                
                if PLOTLY_AVAILABLE and len(category_counts) > 0:
                    fig_pie = go.Figure(data=[go.Pie(
                        labels=category_counts.index,
                        values=category_counts.values,
                        hole=0.3,
                        marker_colors=['#ff4444', '#ff8800', '#ffdd00', '#44dd44']
                    )])
                    fig_pie.update_layout(height=400)
                    st.plotly_chart(fig_pie, use_container_width=True)
                else:
                    st.dataframe(category_counts.reset_index().rename(columns={'index': 'Категория', 'Grade_Category': 'Количество'}))
            else:
                st.info("Нет данных для выбранного диапазона статистики")

        # Сравнение параллелей (если применен фильтр по параллелям)
        if parallel_mode and selected_parallels and len(selected_parallels) > 1:
            st.subheader("📢 Сравнение параллелей")
            
            # Добавляем колонку с параллелью для анализа
            stats_df_with_parallel = stats_df.copy() if len(stats_df) > 0 else pd.DataFrame()
            if len(stats_df_with_parallel) > 0:
                stats_df_with_parallel['Parallel'] = stats_df_with_parallel['Class'].apply(extract_parallel_from_class)
                
                parallel_stats = stats_df_with_parallel.groupby('Parallel')['Average'].agg(['mean', 'count']).reset_index()
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
        
        # Сравнение классов
        if len(stats_df) > 0 and stats_df['Class'].nunique() > 1:
            st.subheader("🏫 Анализ по классам")
            
            class_avg = stats_df.groupby('Class')['Average'].agg(['mean', 'count']).reset_index()
            class_avg['mean'] = class_avg['mean'].round(1)
            class_avg = class_avg.sort_values('mean', ascending=False)
            
            # Показываем топ и аутсайдеров
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
            
            # Scatter plot успеваемости классов
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
        if len(stats_df) > 50:  # Показываем только если достаточно данных
            st.subheader("📈 Детальный анализ распределения")
            
            analysis_type = st.radio(
                "Выберите тип анализа:",
                ["По классам", "По предметам", "По параллелям"] if parallel_mode and len(selected_parallels) > 1 else ["По классам", "По предметам"],
                horizontal=True
            )
            
            if PLOTLY_AVAILABLE:
                if analysis_type == "По классам":
                    # Ограничиваем количество классов для читаемости
                    top_classes_for_box = stats_df.groupby('Class')['Average'].count().nlargest(15).index
                    df_for_box = stats_df[stats_df['Class'].isin(top_classes_for_box)]
                    
                    fig_box = create_plotly_chart(
                        'box',
                        df_for_box,
                        x='Class',
                        y='Average',
                        labels={'Average': 'Оценка', 'Class': 'Класс'}
                    )
                elif analysis_type == "По предметам":
                    # Ограничиваем количество предметов для читаемости
                    top_subjects_for_box = stats_df.groupby('Subject')['Average'].count().nlargest(12).index
                    df_for_box = stats_df[stats_df['Subject'].isin(top_subjects_for_box)]
                    
                    fig_box = create_plotly_chart(
                        'box',
                        df_for_box,
                        x='Subject',
                        y='Average',
                        labels={'Average': 'Оценка', 'Subject': 'Предмет'}
                    )
                else:  # По параллелям
                    stats_df_with_parallel = stats_df.copy()
                    stats_df_with_parallel['Parallel'] = stats_df_with_parallel['Class'].apply(extract_parallel_from_class)
                    
                    fig_box = create_plotly_chart(
                        'box',
                        stats_df_with_parallel,
                        x='Parallel',
                        y='Average',
                        labels={'Average': 'Оценка', 'Parallel': 'Параллель'}
                    )
                
                if fig_box:
                    fig_box.update_layout(height=450)
                    fig_box.update_xaxes(tickangle=-45)
                    st.plotly_chart(fig_box, use_container_width=True)
            else:
                # Fallback: статистика
                if analysis_type == "По классам":
                    box_stats = stats_df.groupby('Class')['Average'].agg(['min', 'max', 'mean', 'median']).round(1)
                elif analysis_type == "По предметам":
                    box_stats = stats_df.groupby('Subject')['Average'].agg(['min', 'max', 'mean', 'median']).round(1)
                else:
                    stats_df_with_parallel = stats_df.copy()
                    stats_df_with_parallel['Parallel'] = stats_df_with_parallel['Class'].apply(extract_parallel_from_class)
                    box_stats = stats_df_with_parallel.groupby('Parallel')['Average'].agg(['min', 'max', 'mean', 'median']).round(1)
                
                st.dataframe(box_stats, use_container_width=True)
        
        # Детальная таблица с улучшенными опциями
        with st.expander("📋 Детальные данные и экспорт"):
            # Дополнительные опции для таблицы
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                sort_by = st.selectbox(
                    "Сортировать по:",
                    ["Average", "Student", "Class", "Subject"],
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
            
            # Используем stats_df для отображения таблицы
            display_df_source = stats_df if len(stats_df) > 0 else filtered_df
            
            # Применяем сортировку
            ascending = sort_order == "По возрастанию"
            sorted_df = display_df_source.sort_values(sort_by, ascending=ascending)
            
            # Ограничиваем количество строк
            if show_rows != "Все":
                display_df = sorted_df.head(show_rows)
            else:
                display_df = sorted_df
            
            if show_stats and len(display_df) > 0:
                # Статистика по отображаемым данным
                st.markdown("**📊 Статистика отображаемых данных:**")
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Записей", len(display_df))
                with col2:
                    st.metric("Средняя оценка", f"{display_df['Average'].mean():.1f}")
                with col3:
                    st.metric("Медиана", f"{display_df['Average'].median():.1f}")
                with col4:
                    st.metric("Станд. отклонение", f"{display_df['Average'].std():.1f}")
            
            if len(display_df) > 0:
                st.dataframe(
                    display_df,
                    use_container_width=True,
                    height=400
                )
            else:
                st.info("Нет данных для отображения")
            
            # Кнопки для скачивания
            if len(display_df_source) > 0:
                st.markdown("**📥 Экспорт данных:**")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    csv = display_df_source.to_csv(index=False)
                    st.download_button(
                        label="📄 Полные данные (CSV)",
                        data=csv,
                        file_name=f'grades_data_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
                        mime='text/csv'
                    )
                
                with col2:
                    summary_stats = display_df_source.groupby(['Class', 'Subject'])['Average'].agg(['mean', 'count', 'std']).round(2)
                    st.download_button(
                        label="📊 Сводная статистика (CSV)",
                        data=summary_stats.to_csv(),
                        file_name=f'summary_stats_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
                        mime='text/csv'
                    )
                
                with col3:
                    # Экспорт рейтинга предметов
                    if subject_avg is not None:
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
        """)
    
    # Дополнительная информация в сайдбаре
    with st.sidebar:
        st.markdown("---")
        st.markdown("### 📈 Статистика")
        if len(filtered_df) > 0:
            # Используем stats_df для статистики в сайдбаре
            stats_df = filtered_df[
                (filtered_df['Average'] >= stats_grade_range[0]) & 
                (filtered_df['Average'] <= stats_grade_range[1])
            ] if len(filtered_df) > 0 else pd.DataFrame()
            
            if len(stats_df) > 0:
                st.write(f"**Записей:** {len(stats_df):,}")
                st.write(f"**Медиана:** {stats_df['Average'].median():.1f}")
                st.write(f"**Станд. отклонение:** {stats_df['Average'].std():.1f}")
                st.write(f"**Мин. оценка:** {stats_df['Average'].min():.1f}")
                st.write(f"**Макс. оценка:** {stats_df['Average'].max():.1f}")
            else:
                st.write("**Нет данных в диапазоне статистики**")
        
        # Информация о параллелях
        if not filtered_df.empty:
            st.markdown("### 📢 Доступные параллели")
            available_parallels = get_available_parallels(df)
            st.write(", ".join(available_parallels))
        
        st.markdown("---")
        st.markdown("### 📁 Загрузка данных")
        uploaded_file = st.file_uploader(
            "Загрузите Excel файл",
            type=['xlsx', 'xls'],
            help="Выберите файл с данными об оценках"
        )
        
        if uploaded_file is not None:
            try:
                df_uploaded = pd.read_excel(uploaded_file, sheet_name=0)
                st.success(f"✅ Файл загружен: {len(df_uploaded)} записей")
                
                if all(col in df_uploaded.columns for col in ['Student', 'Class', 'Subject', 'Average']):
                    st.success("✅ Структура данных корректная")
                else:
                    st.warning("⚠️ Проверьте названия столбцов: Student, Class, Subject, Average")
                    st.write("Найденные столбцы:", list(df_uploaded.columns))
                    
            except Exception as e:
                st.error(f"❌ Ошибка загрузки: {e}")
        
        st.markdown("---")
        st.markdown("### 💡 Подсказки")
        st.info("""
        **📢 Фильтр по параллелям:**
        • Включите для выбора целых параллелей
        • Например: 10 = все 10-е классы (10A, 10B, 10C...)
        
        **📊 Диапазоны оценок:**
        • Фильтр данных - отбирает записи
        • Диапазон статистики - показывает статистику
        
        **🏆 Рейтинг предметов:**
        • Зеленые - лучшие предметы
        • Красные - требуют внимания
        
        **📋 Пресеты:**
        • Сохраняйте часто используемые фильтры
        • Используйте готовые пресеты по параллелям
        • Делитесь настройками с коллегами
        
        **📈 Категории оценок:**
        • [0;40) - Неудовлетворительно
        • [40;65) - Удовлетворительно  
        • [65;85) - Хорошо
        • [85;100] - Отлично
        """)

if __name__ == "__main__":
    main()
