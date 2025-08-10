import streamlit as st
import pandas as pd
import numpy as np
import json
from datetime import datetime

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
            "subjects": [],
            "grade_range": [0, 100],
            "top_n": 10
        },
        "Старшие классы": {
            "classes": ["10A", "10B", "10C", "11A", "11B", "11C"],
            "subjects": [],
            "grade_range": [0, 100],
            "top_n": 15
        },
        "Точные науки": {
            "classes": [],
            "subjects": ["Math", "Physics", "Chemistry", "CS"],
            "grade_range": [0, 100],
            "top_n": 5
        }
    }

if 'current_filters' not in st.session_state:
    st.session_state.current_filters = {
        "classes": [],
        "subjects": [],
        "grade_range": [0, 100],
        "top_n": 10
    }

@st.cache_data
def load_data():
    """Загружает и кэширует данные"""
    try:
        possible_files = [
            'Marks 2425.xlsx',
            'marks_2425.xlsx', 
            'data/Marks 2425.xlsx',
            'data/marks_2425.xlsx'
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
    classes = ['10A', '10B', '10C', '11A', '11B', '11C', '9A', '9B', '9C']
    subjects = ['Math', 'Physics', 'Chemistry', 'Biology', 'English', 'History', 'Geography', 'Literature', 'CS', 'Art']
    
    data = []
    for student in students:
        class_name = np.random.choice(classes)
        for subject in np.random.choice(subjects, size=np.random.randint(5, 8), replace=False):
            score = np.random.normal(75, 12)
            score = max(40, min(100, score))
            data.append({
                'Student': student,
                'Class': class_name,
                'Subject': subject,
                'Average': round(score, 1)
            })
    
    return pd.DataFrame(data)

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
        if st.button("📥 Загрузить", disabled=not selected_preset):
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
    
    # Мультиселект для классов
    all_classes = sorted(df['Class'].unique().tolist())
    selected_classes = st.sidebar.multiselect(
        "📚 Выберите классы:",
        options=all_classes,
        default=current_filters.get('classes', []),
        help="Оставьте пустым для выбора всех классов"
    )
    
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
        "🏆 Количество топ предметов:",
        min_value=5,
        max_value=25,
        value=current_filters.get('top_n', 10),
        help="Количество лучших предметов для отображения"
    )
    
    # Обновляем current_filters
    st.session_state.current_filters = {
        'classes': selected_classes,
        'subjects': selected_subjects,
        'grade_range': list(grade_range),
        'top_n': top_n
    }
    
    # Кнопка сброса фильтров
    if st.sidebar.button("🔄 Сбросить все фильтры"):
        st.session_state.current_filters = {
            'classes': [],
            'subjects': [],
            'grade_range': [min_possible, max_possible],
            'top_n': 10
        }
        st.rerun()
    
    return selected_classes, selected_subjects, grade_range, top_n

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

def render_filter_summary(selected_classes, selected_subjects, grade_range, original_df, filtered_df):
    """Отображает сводку примененных фильтров"""
    with st.expander("🔍 Примененные фильтры", expanded=False):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.write("**📚 Классы:**")
            if selected_classes:
                st.write(f"• {', '.join(selected_classes)}")
            else:
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
        
        # Статистика фильтрации
        original_count = len(original_df)
        filtered_count = len(filtered_df)
        percentage = (filtered_count / original_count * 100) if original_count > 0 else 0
        
        st.info(f"📈 Отображено {filtered_count:,} из {original_count:,} записей ({percentage:.1f}%)")

def main():
    st.title("📊 Дашборд успеваемости школы")
    st.markdown("*Интерактивная аналитика с мультиселектом и пресетами фильтров*")
    st.markdown("---")
    
    # Загрузка данных
    df = load_data()
    
    if df is None or len(df) == 0:
        st.error("Не удалось загрузить данные")
        return
    
    # Отображение фильтров и получение выбранных значений
    selected_classes, selected_subjects, grade_range, top_n = render_filter_sidebar(df)
    
    # Применение фильтров
    filtered_df = apply_filters(df, selected_classes, selected_subjects, grade_range)
    
    # Отображение сводки фильтров
    render_filter_summary(selected_classes, selected_subjects, grade_range, df, filtered_df)
    
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
        
        # Графики
        col1, col2 = st.columns(2)
        
        with col1:
            # Рейтинг предметов
            st.subheader(f"🏆 Топ-{top_n} предметов")
            subject_avg = filtered_df.groupby('Subject')['Average'].agg(['mean', 'count']).reset_index()
            subject_avg['mean'] = subject_avg['mean'].round(1)
            subject_avg = subject_avg.sort_values('mean', ascending=False).head(top_n)
            
            if PLOTLY_AVAILABLE and len(subject_avg) > 0:
                fig_subjects = create_plotly_chart(
                    'bar',
                    subject_avg,
                    x='mean',
                    y='Subject',
                    orientation='h',
                    color='mean',
                    color_continuous_scale='Viridis',
                    labels={'mean': 'Средняя оценка', 'Subject': 'Предмет'},
                    text='mean'
                )
                if fig_subjects:
                    fig_subjects.update_layout(
                        height=400,
                        yaxis={'categoryorder': 'total ascending'},
                        showlegend=False
                    )
                    fig_subjects.update_traces(texttemplate='%{text:.1f}', textposition='outside')
                    st.plotly_chart(fig_subjects, use_container_width=True)
            else:
                st.dataframe(
                    subject_avg[['Subject', 'mean', 'count']].rename(columns={
                        'Subject': 'Предмет',
                        'mean': 'Средняя оценка', 
                        'count': 'Количество'
                    })
                )
        
        with col2:
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
        
        # Сравнение классов
        if len(filtered_df['Class'].unique()) > 1:
            st.subheader("🏫 Сравнение классов")
            class_avg = filtered_df.groupby('Class')['Average'].agg(['mean', 'count']).reset_index()
            class_avg['mean'] = class_avg['mean'].round(1)
            
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
                    hover_data=['count']
                )
                if fig_classes:
                    fig_classes.update_layout(height=400)
                    st.plotly_chart(fig_classes, use_container_width=True)
            else:
                st.dataframe(class_avg.rename(columns={'mean': 'Средняя оценка', 'count': 'Количество'}))
        
        # Box plot по классам или предметам
        if len(filtered_df['Class'].unique()) > 1:
            st.subheader("📈 Детальный анализ распределения")
            
            analysis_type = st.radio(
                "Выберите тип анализа:",
                ["По классам", "По предметам"],
                horizontal=True
            )
            
            if PLOTLY_AVAILABLE:
                if analysis_type == "По классам":
                    fig_box = create_plotly_chart(
                        'box',
                        filtered_df,
                        x='Class',
                        y='Average',
                        labels={'Average': 'Оценка', 'Class': 'Класс'}
                    )
                else:
                    # Ограничиваем количество предметов для читаемости
                    top_subjects_for_box = filtered_df.groupby('Subject')['Average'].count().nlargest(10).index
                    df_for_box = filtered_df[filtered_df['Subject'].isin(top_subjects_for_box)]
                    
                    fig_box = create_plotly_chart(
                        'box',
                        df_for_box,
                        x='Subject',
                        y='Average',
                        labels={'Average': 'Оценка', 'Subject': 'Предмет'}
                    )
                
                if fig_box:
                    fig_box.update_layout(height=400)
                    fig_box.update_xaxes(tickangle=-45)
                    st.plotly_chart(fig_box, use_container_width=True)
            else:
                if analysis_type == "По классам":
                    box_stats = filtered_df.groupby('Class')['Average'].agg(['min', 'max', 'mean', 'median']).round(1)
                else:
                    box_stats = filtered_df.groupby('Subject')['Average'].agg(['min', 'max', 'mean', 'median']).round(1)
                st.dataframe(box_stats)
        
        # Детальная таблица
        with st.expander("📋 Детальные данные"):
            # Дополнительные опции для таблицы
            col1, col2, col3 = st.columns(3)
            
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
                    [50, 100, 200, "Все"],
                    index=0
                )
            
            # Применяем сортировку
            ascending = sort_order == "По возрастанию"
            sorted_df = filtered_df.sort_values(sort_by, ascending=ascending)
            
            # Ограничиваем количество строк
            if show_rows != "Все":
                display_df = sorted_df.head(show_rows)
            else:
                display_df = sorted_df
            
            st.dataframe(
                display_df,
                use_container_width=True,
                height=400
            )
            
            # Кнопки для скачивания
            col1, col2 = st.columns(2)
            
            with col1:
                csv = filtered_df.to_csv(index=False)
                st.download_button(
                    label="📥 Скачать отфильтрованные данные (CSV)",
                    data=csv,
                    file_name=f'filtered_grades_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
                    mime='text/csv'
                )
            
            with col2:
                excel_data = filtered_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Скачать сводную статистику",
                    data=filtered_df.groupby(['Class', 'Subject'])['Average'].agg(['mean', 'count', 'std']).round(2).to_csv(),
                    file_name=f'summary_stats_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
                    mime='text/csv'
                )
    
    else:
        st.warning("⚠️ Нет данных для выбранных фильтров!")
        st.info("💡 Попробуйте изменить критерии фильтрации или загрузить пресет 'Все данные'")
    
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
        **Мультиселект фильтры:**
        • Выберите несколько классов/предметов
        • Пустой выбор = все элементы
        
        **Пресеты:**
        • Сохраняйте часто используемые фильтры
        • Экспортируйте/импортируйте настройки
        • Делитесь пресетами с коллегами
        """)

if __name__ == "__main__":
    main()
