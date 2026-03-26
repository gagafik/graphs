import streamlit as st
import pandas as pd
import numpy as np
import json
import os
import io
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

# --- Путь к файлу пресетов на диске ---
PRESETS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "filter_presets.json")

DEFAULT_PRESETS = {
    "Все данные": {
        "classes": [],
        "parallel_mode": False,
        "parallels": [],
        "subjects": [],
        "grade_range": [0, 100],
        "top_n": 10
    },
    "2-е параллель": {
        "classes": [],
        "parallel_mode": True,
        "parallels": ["2"],
        "subjects": [],
        "grade_range": [0, 100],
        "top_n": 10
    },
    "Старшие классы (10-11)": {
        "classes": [],
        "parallel_mode": True,
        "parallels": ["10", "11"],
        "subjects": [],
        "grade_range": [0, 100],
        "top_n": 15
    },
    "Средние классы (7-9)": {
        "classes": [],
        "parallel_mode": True,
        "parallels": ["7", "8", "9"],
        "subjects": [],
        "grade_range": [0, 100],
        "top_n": 12
    },
    "Точные науки": {
        "classes": [],
        "parallel_mode": False,
        "parallels": [],
        "subjects": ["Math", "Physics", "Chemistry", "CS", "Calc", "Further Math"],
        "grade_range": [0, 100],
        "top_n": 6
    },
    "Языки": {
        "classes": [],
        "parallel_mode": False,
        "parallels": [],
        "subjects": ["English", "ESL", "Rus", "Kaz", "RusLit", "KazLit"],
        "grade_range": [0, 100],
        "top_n": 8
    }
}


def load_presets_from_disk():
    """Загружает пресеты из JSON-файла на диске"""
    try:
        if os.path.exists(PRESETS_FILE):
            with open(PRESETS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except (json.JSONDecodeError, IOError):
        pass
    return DEFAULT_PRESETS.copy()


def save_presets_to_disk(presets):
    """Сохраняет пресеты в JSON-файл на диске"""
    try:
        with open(PRESETS_FILE, 'w', encoding='utf-8') as f:
            json.dump(presets, f, indent=2, ensure_ascii=False)
    except IOError as e:
        st.error(f"Ошибка сохранения пресетов: {e}")


# Инициализация session state
if 'filter_presets' not in st.session_state:
    st.session_state.filter_presets = load_presets_from_disk()

if 'current_filters' not in st.session_state:
    st.session_state.current_filters = {
        "classes": [],
        "parallel_mode": False,
        "parallels": [],
        "subjects": [],
        "grade_range": [0, 100],
        "top_n": 10
    }

if 'uploaded_df' not in st.session_state:
    st.session_state.uploaded_df = None


@st.cache_data
def load_data_from_file():
    """Загружает данные из файла на диске"""
    possible_files = [
        'Marks 2526.xlsx',
        'marks_2526.xlsx',
        'data/Marks 2526.xlsx',
        'data/marks_2526.xlsx'
    ]

    for file_path in possible_files:
        try:
            df = pd.read_excel(file_path, sheet_name='Average year, no teacher')
            return df
        except (FileNotFoundError, ValueError):
            continue

    return None


def load_data():
    """Загружает данные: из загруженного файла или с диска"""
    if st.session_state.uploaded_df is not None:
        return st.session_state.uploaded_df

    df = load_data_from_file()
    if df is not None:
        return df

    st.warning("⚠️ Файл Excel не найден. Используются демо-данные.")
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
            base_score = 75
            if subject in ['Math', 'Physics', 'Chemistry']:
                base_score = 70
            elif subject in ['Art', 'PE', 'Music']:
                base_score = 85

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
        save_presets_to_disk(st.session_state.filter_presets)
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


def delete_preset(preset_name):
    """Удаляет пресет"""
    if preset_name in st.session_state.filter_presets and preset_name != "Все данные":
        del st.session_state.filter_presets[preset_name]
        save_presets_to_disk(st.session_state.filter_presets)


def export_presets():
    """Экспортирует пресеты в JSON"""
    return json.dumps(st.session_state.filter_presets, indent=2, ensure_ascii=False)


def import_presets(json_data):
    """Импортирует пресеты из JSON"""
    try:
        imported_presets = json.loads(json_data)
        st.session_state.filter_presets.update(imported_presets)
        save_presets_to_disk(st.session_state.filter_presets)
        st.success(f"✅ Импортировано {len(imported_presets)} пресетов!")
    except json.JSONDecodeError:
        st.error("❌ Ошибка: неверный формат JSON!")


def render_filter_sidebar(df):
    """Отображает боковую панель с фильтрами и пресетами"""
    st.sidebar.header("🔧 Фильтры и Пресеты")

    # --- СЕКЦИЯ ПРЕСЕТОВ (свёрнута по умолчанию) ---
    with st.sidebar.expander("📋 Управление пресетами"):
        preset_names = list(st.session_state.filter_presets.keys())
        selected_preset = st.selectbox(
            "Выберите пресет:",
            [""] + preset_names,
            format_func=lambda x: "Выберите пресет..." if x == "" else x
        )

        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔥 Загрузить", disabled=not selected_preset):
                load_preset(selected_preset)
                st.rerun()

        with col2:
            if st.button("🗑️ Удалить", disabled=not selected_preset or selected_preset == "Все данные"):
                delete_preset(selected_preset)
                st.rerun()

        st.text_input(
            "Название нового пресета:",
            key="new_preset_name",
            placeholder="Введите название..."
        )

        if st.button("💾 Сохранить текущие фильтры"):
            save_preset()
            st.rerun()

        # Экспорт/Импорт пресетов
        st.markdown("---")
        st.download_button(
            label="📤 Экспорт пресетов (JSON)",
            data=export_presets(),
            file_name=f"filter_presets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            mime="application/json"
        )

        imported_json = st.text_area(
            "📥 Импорт (вставьте JSON):",
            height=80,
            placeholder='{"Мой пресет": {...}}'
        )

        if st.button("📥 Импортировать") and imported_json.strip():
            import_presets(imported_json)
            st.rerun()

    st.sidebar.markdown("---")

    # --- СЕКЦИЯ ФИЛЬТРОВ ---
    st.sidebar.subheader("🎛️ Фильтры")

    current_filters = st.session_state.current_filters

    parallel_mode = st.sidebar.checkbox(
        "📢 Фильтр по параллелям",
        value=current_filters.get('parallel_mode', False),
        help="Включите для выбора целых параллелей (2, 3, 4... классы)"
    )

    if parallel_mode:
        available_parallels = get_available_parallels(df)
        selected_parallels = st.sidebar.multiselect(
            "📊 Выберите параллели:",
            options=available_parallels,
            default=current_filters.get('parallels', []),
            help="Выберите номера параллелей (например, 10 = все 10-е классы)"
        )

        if selected_parallels:
            included_classes = get_classes_by_parallels(df, selected_parallels)
            with st.sidebar.expander(f"📋 Включенные классы ({len(included_classes)})"):
                st.write(", ".join(included_classes))

        selected_classes = get_classes_by_parallels(df, selected_parallels)
    else:
        all_classes = sorted(df['Class'].unique().tolist())
        selected_classes = st.sidebar.multiselect(
            "📚 Выберите классы:",
            options=all_classes,
            default=current_filters.get('classes', []),
            help="Оставьте пустым для выбора всех классов"
        )
        selected_parallels = []

    all_subjects = sorted(df['Subject'].unique().tolist())
    selected_subjects = st.sidebar.multiselect(
        "📖 Выберите предметы:",
        options=all_subjects,
        default=current_filters.get('subjects', []),
        help="Оставьте пустым для выбора всех предметов"
    )

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

    top_n = st.sidebar.slider(
        "🏆 Количество предметов в рейтинге:",
        min_value=5,
        max_value=25,
        value=current_filters.get('top_n', 10),
        help="Количество предметов для отображения в топе"
    )

    st.session_state.current_filters = {
        'classes': selected_classes,
        'parallel_mode': parallel_mode,
        'parallels': selected_parallels,
        'subjects': selected_subjects,
        'grade_range': list(grade_range),
        'top_n': top_n
    }

    if st.sidebar.button("🔄 Сбросить все фильтры"):
        st.session_state.current_filters = {
            'classes': [],
            'parallel_mode': False,
            'parallels': [],
            'subjects': [],
            'grade_range': [min_possible, max_possible],
            'top_n': 10
        }
        st.rerun()

    # --- Загрузка файла ---
    st.sidebar.markdown("---")
    with st.sidebar.expander("📁 Загрузка данных"):
        uploaded_file = st.file_uploader(
            "Загрузите Excel файл",
            type=['xlsx', 'xls'],
            help="Файл должен содержать столбцы: Student, Class, Subject, Average"
        )

        if uploaded_file is not None:
            try:
                df_uploaded = pd.read_excel(uploaded_file, sheet_name=0)
                required_cols = ['Student', 'Class', 'Subject', 'Average']

                if all(col in df_uploaded.columns for col in required_cols):
                    st.session_state.uploaded_df = df_uploaded
                    st.success(f"✅ Загружено {len(df_uploaded)} записей")
                    st.rerun()
                else:
                    missing = [c for c in required_cols if c not in df_uploaded.columns]
                    st.error(f"❌ Не хватает столбцов: {', '.join(missing)}")
                    st.write("Найденные столбцы:", list(df_uploaded.columns))
            except Exception as e:
                st.error(f"❌ Ошибка загрузки: {e}")

        if st.session_state.uploaded_df is not None:
            if st.button("🔄 Вернуться к исходным данным"):
                st.session_state.uploaded_df = None
                st.rerun()

    # --- Подсказки ---
    with st.sidebar.expander("💡 Подсказки"):
        st.info("""
        **📢 Фильтр по параллелям:**
        • Включите для выбора целых параллелей

        **🏆 Рейтинг предметов:**
        • Зеленые — лучшие предметы
        • Красные — требуют внимания

        **📋 Пресеты:**
        • Сохраняются на диск автоматически

        **🔍 Поиск ученика:**
        • Используйте вкладку «Профиль ученика»
        """)

    return selected_classes, selected_subjects, grade_range, top_n, parallel_mode, selected_parallels


def apply_filters(df, selected_classes, selected_subjects, grade_range):
    """Применяет фильтры к данным"""
    filtered_df = df.copy()

    if selected_classes:
        filtered_df = filtered_df[filtered_df['Class'].isin(selected_classes)]

    if selected_subjects:
        filtered_df = filtered_df[filtered_df['Subject'].isin(selected_subjects)]

    filtered_df = filtered_df[
        (filtered_df['Average'] >= grade_range[0]) &
        (filtered_df['Average'] <= grade_range[1])
    ]

    return filtered_df


def render_filter_summary(selected_classes, selected_subjects, grade_range, original_df, filtered_df, parallel_mode, selected_parallels):
    """Отображает сводку примененных фильтров"""
    with st.expander("🔍 Примененные фильтры", expanded=False):
        col1, col2, col3 = st.columns(3)

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
                st.write("**📚 Классы:** Все")

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
            st.write(f"• {grade_range[0]} – {grade_range[1]}")

        original_count = len(original_df)
        filtered_count = len(filtered_df)
        percentage = (filtered_count / original_count * 100) if original_count > 0 else 0
        st.info(f"📈 Отображено {filtered_count:,} из {original_count:,} записей ({percentage:.1f}%)")


def create_subject_ranking_charts(filtered_df, top_n):
    """Создает графики рейтинга лучших и худших предметов. Возвращает subject_avg."""
    subject_avg = filtered_df.groupby('Subject')['Average'].agg(['mean', 'count']).reset_index()
    subject_avg['mean'] = subject_avg['mean'].round(1)
    subject_avg = subject_avg.sort_values('mean', ascending=False)

    top_best = subject_avg.head(top_n)
    top_worst = subject_avg.tail(top_n).sort_values('mean', ascending=True)

    col1, col2 = st.columns(2)

    with col1:
        st.subheader(f"🏆 Топ-{top_n} лучших предметов")
        if PLOTLY_AVAILABLE and len(top_best) > 0:
            fig_best = create_plotly_chart(
                'bar', top_best,
                x='mean', y='Subject', orientation='h',
                color='mean', color_continuous_scale='Greens',
                labels={'mean': 'Средняя оценка', 'Subject': 'Предмет'},
                text='mean'
            )
            if fig_best:
                fig_best.update_layout(height=400, yaxis={'categoryorder': 'total ascending'}, showlegend=False)
                fig_best.update_traces(texttemplate='%{text:.1f}', textposition='outside')
                st.plotly_chart(fig_best, use_container_width=True)
        else:
            st.dataframe(top_best[['Subject', 'mean', 'count']].rename(
                columns={'Subject': 'Предмет', 'mean': 'Средняя оценка', 'count': 'Количество'}
            ), use_container_width=True)

    with col2:
        st.subheader(f"⚠️ Топ-{top_n} предметов для внимания")
        if PLOTLY_AVAILABLE and len(top_worst) > 0:
            fig_worst = create_plotly_chart(
                'bar', top_worst,
                x='mean', y='Subject', orientation='h',
                color='mean', color_continuous_scale='Reds_r',
                labels={'mean': 'Средняя оценка', 'Subject': 'Предмет'},
                text='mean'
            )
            if fig_worst:
                fig_worst.update_layout(height=400, yaxis={'categoryorder': 'total ascending'}, showlegend=False)
                fig_worst.update_traces(texttemplate='%{text:.1f}', textposition='outside')
                st.plotly_chart(fig_worst, use_container_width=True)
        else:
            st.dataframe(top_worst[['Subject', 'mean', 'count']].rename(
                columns={'Subject': 'Предмет', 'mean': 'Средняя оценка', 'count': 'Количество'}
            ), use_container_width=True)

    return subject_avg


def render_heatmap(filtered_df):
    """Отображает тепловую карту: класс x предмет"""
    st.subheader("🗺️ Тепловая карта: Класс × Предмет")

    pivot = filtered_df.pivot_table(
        values='Average', index='Class', columns='Subject', aggfunc='mean'
    ).round(1)

    if pivot.empty or pivot.shape[0] < 2 or pivot.shape[1] < 2:
        st.info("Недостаточно данных для тепловой карты. Выберите больше классов/предметов.")
        return

    # Сортируем классы по номеру параллели
    sorted_classes = sorted(pivot.index, key=lambda x: (int(re.search(r'^(\d+)', str(x)).group(1)) if re.search(r'^(\d+)', str(x)) else 999, str(x)))
    pivot = pivot.reindex(sorted_classes)

    if PLOTLY_AVAILABLE:
        fig = px.imshow(
            pivot,
            color_continuous_scale='RdYlGn',
            aspect='auto',
            labels=dict(x='Предмет', y='Класс', color='Средний балл'),
            text_auto='.1f'
        )
        fig.update_layout(height=max(400, len(pivot) * 30 + 100))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.dataframe(pivot, use_container_width=True)


def render_student_profile(df):
    """Отображает профиль ученика с поиском"""
    st.subheader("🔍 Профиль ученика")

    all_students = sorted(df['Student'].unique().tolist())

    search_query = st.text_input(
        "Поиск ученика:",
        placeholder="Начните вводить имя..."
    )

    if search_query:
        matched = [s for s in all_students if search_query.lower() in s.lower()]
    else:
        matched = all_students

    if not matched:
        st.warning("Ученик не найден.")
        return

    selected_student = st.selectbox(
        "Выберите ученика:",
        options=matched,
        index=0
    )

    if selected_student:
        student_data = df[df['Student'] == selected_student].copy()

        if student_data.empty:
            st.warning("Нет данных для этого ученика.")
            return

        student_class = student_data['Class'].iloc[0]
        student_avg = student_data['Average'].mean()
        total_avg = df['Average'].mean()

        # Карточка ученика
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("👤 Ученик", selected_student)
        with col2:
            st.metric("🏫 Класс", student_class)
        with col3:
            st.metric("📈 Средний балл", f"{student_avg:.1f}")
        with col4:
            # Ранг в классе
            class_students = df[df['Class'] == student_class].groupby('Student')['Average'].mean()
            rank = (class_students > student_avg).sum() + 1
            total_in_class = len(class_students)
            st.metric("🏅 Место в классе", f"{rank}/{total_in_class}")

        # Таблица оценок
        student_display = student_data[['Subject', 'Average']].sort_values('Average', ascending=False)
        student_display = student_display.rename(columns={'Subject': 'Предмет', 'Average': 'Оценка'})

        col1, col2 = st.columns([1, 1])

        with col1:
            st.dataframe(student_display, use_container_width=True, hide_index=True)

        with col2:
            if PLOTLY_AVAILABLE and len(student_data) > 0:
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=student_data['Subject'],
                    y=student_data['Average'],
                    marker_color=[
                        '#44dd44' if v >= 85 else '#88dd00' if v >= 65 else '#ff8800' if v >= 40 else '#ff4444'
                        for v in student_data['Average']
                    ],
                    text=student_data['Average'].round(1),
                    textposition='outside'
                ))
                fig.add_hline(y=total_avg, line_dash="dash", line_color="gray",
                              annotation_text=f"Среднее по школе: {total_avg:.1f}")
                fig.update_layout(
                    height=350,
                    xaxis_title="Предмет",
                    yaxis_title="Оценка",
                    showlegend=False
                )
                fig.update_xaxes(tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)


def render_overview_metrics(df, filtered_df):
    """Отображает основные метрики без дельт"""
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(label="👥 Учеников", value=filtered_df['Student'].nunique())
    with col2:
        st.metric(label="🏫 Классов", value=filtered_df['Class'].nunique())
    with col3:
        st.metric(label="📚 Предметов", value=filtered_df['Subject'].nunique())
    with col4:
        avg_score = filtered_df['Average'].mean()
        st.metric(label="📈 Средний балл", value=f"{avg_score:.1f}")


def render_distribution_and_categories(filtered_df, parallel_mode, selected_parallels):
    """Распределение оценок и категории успеваемости"""
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📊 Распределение оценок")
        if PLOTLY_AVAILABLE:
            fig_dist = create_plotly_chart(
                'histogram', filtered_df,
                x='Average', nbins=20,
                color_discrete_sequence=['#636EFA']
            )
            if fig_dist:
                fig_dist.add_vline(
                    x=filtered_df['Average'].mean(),
                    line_dash="dash", line_color="red",
                    annotation_text=f"Среднее: {filtered_df['Average'].mean():.1f}"
                )
                st.plotly_chart(fig_dist, use_container_width=True)
        else:
            st.write(filtered_df['Average'].describe())

    with col2:
        if parallel_mode and selected_parallels and len(selected_parallels) > 1:
            st.subheader("📢 Сравнение параллелей")
            temp_df = filtered_df.copy()
            temp_df['Parallel'] = temp_df['Class'].apply(extract_parallel_from_class)
            parallel_stats = temp_df.groupby('Parallel')['Average'].agg(['mean', 'count']).reset_index()
            parallel_stats['mean'] = parallel_stats['mean'].round(1)

            if PLOTLY_AVAILABLE:
                fig_parallels = create_plotly_chart(
                    'bar', parallel_stats,
                    x='Parallel', y='mean',
                    color='mean', color_continuous_scale='Viridis',
                    labels={'mean': 'Средняя оценка', 'Parallel': 'Параллель'},
                    text='mean'
                )
                if fig_parallels:
                    fig_parallels.update_traces(texttemplate='%{text:.1f}', textposition='outside')
                    st.plotly_chart(fig_parallels, use_container_width=True)
            else:
                st.dataframe(parallel_stats.rename(columns={'Parallel': 'Параллель', 'mean': 'Средняя оценка'}))
        else:
            st.subheader("📈 Категории успеваемости")
            temp_df = filtered_df.copy()
            bins = [0, 40, 65, 85, 100]
            labels = ['Неуд. (0-39)', 'Удовл. (40-64)', 'Хорошо (65-84)', 'Отлично (85-100)']
            temp_df['Grade_Category'] = pd.cut(temp_df['Average'], bins=bins, labels=labels, include_lowest=True)
            category_counts = temp_df['Grade_Category'].value_counts()

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
                st.dataframe(category_counts)


def render_class_analysis(filtered_df):
    """Анализ по классам"""
    if len(filtered_df['Class'].unique()) <= 1:
        return

    st.subheader("🏫 Анализ по классам")

    class_avg = filtered_df.groupby('Class')['Average'].agg(['mean', 'count']).reset_index()
    class_avg['mean'] = class_avg['mean'].round(1)
    class_avg = class_avg.sort_values('mean', ascending=False)

    col1, col2 = st.columns(2)
    with col1:
        st.write("**🏆 Лучшие классы:**")
        for _, row in class_avg.head(5).iterrows():
            st.write(f"• {row['Class']}: {row['mean']:.1f} ({row['count']} оценок)")
    with col2:
        st.write("**⚠️ Классы для внимания:**")
        for _, row in class_avg.tail(5).iterrows():
            st.write(f"• {row['Class']}: {row['mean']:.1f} ({row['count']} оценок)")

    if PLOTLY_AVAILABLE and len(class_avg) > 0:
        fig_classes = create_plotly_chart(
            'scatter', class_avg,
            x='Class', y='mean', size='count',
            color='mean', color_continuous_scale='RdYlGn',
            labels={'mean': 'Средняя оценка', 'count': 'Количество оценок'},
            hover_data=['count'],
            title="Успеваемость по классам"
        )
        if fig_classes:
            fig_classes.update_layout(height=400)
            st.plotly_chart(fig_classes, use_container_width=True)


def render_box_analysis(filtered_df, parallel_mode, selected_parallels):
    """Box plot анализ"""
    if len(filtered_df) <= 50:
        return

    st.subheader("📈 Детальный анализ распределения")

    options = ["По классам", "По предметам"]
    if parallel_mode and len(selected_parallels) > 1:
        options.append("По параллелям")

    analysis_type = st.radio("Выберите тип анализа:", options, horizontal=True)

    if PLOTLY_AVAILABLE:
        fig_box = None
        if analysis_type == "По классам":
            top_classes_for_box = filtered_df.groupby('Class')['Average'].count().nlargest(15).index
            df_for_box = filtered_df[filtered_df['Class'].isin(top_classes_for_box)]
            fig_box = create_plotly_chart('box', df_for_box, x='Class', y='Average',
                                          labels={'Average': 'Оценка', 'Class': 'Класс'})
        elif analysis_type == "По предметам":
            top_subjects_for_box = filtered_df.groupby('Subject')['Average'].count().nlargest(12).index
            df_for_box = filtered_df[filtered_df['Subject'].isin(top_subjects_for_box)]
            fig_box = create_plotly_chart('box', df_for_box, x='Subject', y='Average',
                                          labels={'Average': 'Оценка', 'Subject': 'Предмет'})
        else:
            temp_df = filtered_df.copy()
            temp_df['Parallel'] = temp_df['Class'].apply(extract_parallel_from_class)
            fig_box = create_plotly_chart('box', temp_df, x='Parallel', y='Average',
                                          labels={'Average': 'Оценка', 'Parallel': 'Параллель'})

        if fig_box:
            fig_box.update_layout(height=450)
            fig_box.update_xaxes(tickangle=-45)
            st.plotly_chart(fig_box, use_container_width=True)
    else:
        if analysis_type == "По классам":
            stats = filtered_df.groupby('Class')['Average'].agg(['min', 'max', 'mean', 'median']).round(1)
        elif analysis_type == "По предметам":
            stats = filtered_df.groupby('Subject')['Average'].agg(['min', 'max', 'mean', 'median']).round(1)
        else:
            temp_df = filtered_df.copy()
            temp_df['Parallel'] = temp_df['Class'].apply(extract_parallel_from_class)
            stats = temp_df.groupby('Parallel')['Average'].agg(['min', 'max', 'mean', 'median']).round(1)
        st.dataframe(stats, use_container_width=True)


def render_detail_table(filtered_df, subject_avg):
    """Детальная таблица с экспортом"""
    with st.expander("📋 Детальные данные и экспорт"):
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            sort_by = st.selectbox("Сортировать по:", ["Average", "Student", "Class", "Subject"], index=0)
        with col2:
            sort_order = st.selectbox("Порядок:", ["По убыванию", "По возрастанию"], index=0)
        with col3:
            show_rows = st.selectbox("Показать строк:", [50, 100, 200, 500, "Все"], index=0)
        with col4:
            show_stats = st.checkbox("Показать статистику", value=True)

        # Ручной фильтр по диапазону оценок
        st.markdown("**🎯 Дополнительный фильтр по диапазону оценок:**")
        col1, col2 = st.columns(2)
        with col1:
            manual_min_grade = st.number_input(
                "Минимальная оценка:", min_value=0.0, max_value=100.0,
                value=None, step=0.1, placeholder="Например, 75.0"
            )
        with col2:
            manual_max_grade = st.number_input(
                "Максимальная оценка:", min_value=0.0, max_value=100.0,
                value=None, step=0.1, placeholder="Например, 95.0"
            )

        display_filtered_df = filtered_df.copy()
        if manual_min_grade is not None:
            display_filtered_df = display_filtered_df[display_filtered_df['Average'] >= manual_min_grade]
        if manual_max_grade is not None:
            display_filtered_df = display_filtered_df[display_filtered_df['Average'] <= manual_max_grade]

        ascending = sort_order == "По возрастанию"
        sorted_df = display_filtered_df.sort_values(sort_by, ascending=ascending)

        if show_rows != "Все":
            display_df = sorted_df.head(show_rows)
        else:
            display_df = sorted_df

        if show_stats:
            st.markdown("**📊 Статистика отображаемых данных:**")
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                st.metric("Записей", len(display_df))
            with c2:
                st.metric("Средняя", f"{display_df['Average'].mean():.1f}" if len(display_df) > 0 else "—")
            with c3:
                st.metric("Медиана", f"{display_df['Average'].median():.1f}" if len(display_df) > 0 else "—")
            with c4:
                st.metric("Ст. откл.", f"{display_df['Average'].std():.1f}" if len(display_df) > 1 else "—")
            with c5:
                if len(display_df) > 0:
                    st.metric("Диапазон", f"{display_df['Average'].min():.1f}–{display_df['Average'].max():.1f}")

            if manual_min_grade is not None or manual_max_grade is not None:
                parts = []
                if manual_min_grade is not None:
                    parts.append(f"≥ {manual_min_grade}")
                if manual_max_grade is not None:
                    parts.append(f"≤ {manual_max_grade}")
                st.info(f"🎯 Доп. фильтр: {' и '.join(parts)}. "
                        f"Отфильтровано {len(display_filtered_df)} из {len(filtered_df)} записей.")

        st.dataframe(display_df, use_container_width=True, height=400)

        # Кнопки экспорта
        st.markdown("**📥 Экспорт данных:**")
        col1, col2, col3 = st.columns(3)

        with col1:
            csv = display_filtered_df.to_csv(index=False)
            st.download_button(
                label="📄 CSV",
                data=csv,
                file_name=f'grades_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
                mime='text/csv'
            )

        with col2:
            # Excel-экспорт
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                display_filtered_df.to_excel(writer, sheet_name='Данные', index=False)
                summary = display_filtered_df.groupby(['Class', 'Subject'])['Average'].agg(
                    ['mean', 'count', 'std']).round(2)
                summary.to_excel(writer, sheet_name='Сводка')
                if subject_avg is not None:
                    ranking = subject_avg[['Subject', 'mean', 'count']].round(2)
                    ranking.columns = ['Предмет', 'Средняя_оценка', 'Количество']
                    ranking.to_excel(writer, sheet_name='Рейтинг предметов', index=False)
            excel_buffer.seek(0)

            st.download_button(
                label="📊 Excel",
                data=excel_buffer,
                file_name=f'grades_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        with col3:
            summary_csv = display_filtered_df.groupby(['Class', 'Subject'])['Average'].agg(
                ['mean', 'count', 'std']).round(2).to_csv()
            st.download_button(
                label="📊 Сводка (CSV)",
                data=summary_csv,
                file_name=f'summary_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
                mime='text/csv'
            )


def main():
    st.title("📊 Дашборд успеваемости школы")
    st.markdown("*Интерактивная аналитика с фильтрами по параллелям и рейтингом предметов*")
    st.markdown("---")

    # Загрузка данных
    df = load_data()

    if df is None or len(df) == 0:
        st.error("Не удалось загрузить данные")
        return

    if st.session_state.uploaded_df is not None:
        st.success("📂 Используются загруженные данные")

    # Фильтры
    selected_classes, selected_subjects, grade_range, top_n, parallel_mode, selected_parallels = render_filter_sidebar(df)
    filtered_df = apply_filters(df, selected_classes, selected_subjects, grade_range)

    render_filter_summary(selected_classes, selected_subjects, grade_range, df, filtered_df, parallel_mode, selected_parallels)

    if len(filtered_df) == 0:
        st.warning("⚠️ Нет данных для выбранных фильтров!")
        st.info("💡 Попробуйте изменить критерии, загрузить пресет «Все данные» или расширить диапазон.")
        return

    # Вкладки
    tab_overview, tab_heatmap, tab_student = st.tabs(["📊 Обзор", "🗺️ Тепловая карта", "👤 Профиль ученика"])

    with tab_overview:
        render_overview_metrics(df, filtered_df)
        st.markdown("---")

        # Рейтинг предметов
        subject_avg = create_subject_ranking_charts(filtered_df, top_n)
        st.markdown("---")

        render_distribution_and_categories(filtered_df, parallel_mode, selected_parallels)
        st.markdown("---")

        render_class_analysis(filtered_df)
        render_box_analysis(filtered_df, parallel_mode, selected_parallels)
        st.markdown("---")

        render_detail_table(filtered_df, subject_avg)

    with tab_heatmap:
        render_heatmap(filtered_df)

    with tab_student:
        render_student_profile(df)

    # Статистика в сайдбаре
    with st.sidebar:
        st.markdown("---")
        st.markdown("### 📈 Статистика")
        st.write(f"**Записей:** {len(filtered_df):,}")
        st.write(f"**Медиана:** {filtered_df['Average'].median():.1f}")
        st.write(f"**Ст. откл.:** {filtered_df['Average'].std():.1f}")
        st.write(f"**Мин.:** {filtered_df['Average'].min():.1f}")
        st.write(f"**Макс.:** {filtered_df['Average'].max():.1f}")

        if not filtered_df.empty:
            st.markdown("### 📢 Параллели")
            st.write(", ".join(get_available_parallels(df)))


if __name__ == "__main__":
    main()
