import gradio as gr
import pandas as pd
from pathlib import Path
import json
import tempfile
import zipfile
import shutil
import os
from modules.standardizer import ValueStandardizer
from modules.replacer import DataReplacer

# Глобальные переменные для хранения состояния
current_standardizer = None
standardization_rules = {}

def load_standardization_project(zip_file):
    """Загружает проект для стандартизации"""
    global current_standardizer
    
    if not zip_file:
        return None, "❌ Пожалуйста, загрузите ZIP-архив с результатами обработки"
    
    try:
        # Создаем временную директорию
        temp_dir = Path(tempfile.mkdtemp())
        
        # Распаковываем архив
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Инициализируем стандартизатор
        current_standardizer = ValueStandardizer(temp_dir)
        
        # Находим файлы атрибутов
        attribute_files = current_standardizer.find_attribute_files()
        
        if not attribute_files:
            return None, "❌ В архиве не найдены файлы атрибутов. Сначала выполните обработку онтологии."
        
        # Создаем список атрибутов для выбора
        attributes_list = list(attribute_files.keys())
        
        report = f"""
✅ Проект загружен успешно!

📊 Статистика:
• Найдено атрибутов: {len(attributes_list)}
• Файлы: {', '.join(list(attribute_files.keys())[:5])}{'...' if len(attributes_list) > 5 else ''}

🎯 Выберите атрибут для стандартизации из списка ниже.
        """
        
        # Сохраняем временную папку
        gr.Info("Проект загружен. Временные файлы сохранены.")
        
        return gr.Dropdown(choices=attributes_list, value=attributes_list[0] if attributes_list else None), report, temp_dir
        
    except Exception as e:
        return None, f"❌ Ошибка загрузки проекта: {str(e)}", None

def load_attribute_values(attribute_name):
    """Загружает значения выбранного атрибута"""
    global current_standardizer
    
    if not current_standardizer or not attribute_name:
        return None, "❌ Сначала загрузите проект"
    
    try:
        # Находим файл атрибута
        attribute_files = current_standardizer.find_attribute_files()
        file_path = attribute_files.get(attribute_name)
        
        if not file_path:
            return None, f"❌ Файл для атрибута '{attribute_name}' не найден"
        
        # Загружаем значения
        values = current_standardizer.load_attribute_values(file_path)
        
        if not values:
            return None, f"❌ В атрибуте '{attribute_name}' нет значений"
        
        # Группируем похожие значения
        groups = current_standardizer.group_similar_values(values)
        
        # Создаем DataFrame для отображения
        display_data = []
        for i, group in enumerate(groups, 1):
            suggested = current_standardizer.suggest_standard_value(group)
            for value in group:
                display_data.append({
                    'Группа': f"Группа {i}",
                    'Исходное значение': value,
                    'Предложенное стандартное': suggested,
                    'Выбранное стандартное': suggested  # по умолчанию = предложенное
                })
        
        df = pd.DataFrame(display_data)
        
        report = f"""
📊 Атрибут: **{attribute_name}**
• Всего уникальных значений: {len(values)}
• Сгруппировано в: {len(groups)} групп
• Рекомендация: проверьте предложенные стандартные значения и при необходимости измените их.
        """
        
        return df, report, groups
        
    except Exception as e:
        return None, f"❌ Ошибка загрузки значений: {str(e)}", None

def update_standard_values(df, groups, attribute_name):
    """Обновляет правила стандартизации на основе пользовательских правок"""
    global current_standardizer, standardization_rules
    
    if df is None or not groups:
        return "❌ Нет данных для обновления"
    
    try:
        # Преобразуем DataFrame обратно в правила
        rules = {}
        
        # Создаем маппинг из DataFrame
        for idx, row in df.iterrows():
            original = row['Исходное значение']
            standard = row['Выбранное стандартное']
            
            if pd.notna(original) and pd.notna(standard):
                rules[str(original).strip()] = str(standard).strip()
        
        # Сохраняем правила
        if attribute_name not in standardization_rules:
            standardization_rules[attribute_name] = {}
        
        standardization_rules[attribute_name].update(rules)
        
        # Обновляем в стандартизаторе
        current_standardizer.create_standardization_map(attribute_name, rules)
        
        return f"✅ Правила для атрибута '{attribute_name}' обновлены. Изменено: {len(rules)} значений."
        
    except Exception as e:
        return f"❌ Ошибка обновления правил: {str(e)}"

def save_standardization_rules(output_dir=None):
    """Сохраняет все правила стандартизации"""
    global current_standardizer, standardization_rules
    
    if not standardization_rules:
        return None, "❌ Нет правил для сохранения"
    
    try:
        if output_dir is None:
            # Создаем временную директорию
            output_dir = Path(tempfile.mkdtemp())
        
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Сохраняем через стандартизатор
        current_standardizer.standardization_rules = standardization_rules
        current_standardizer.save_standardization_rules(output_dir)
        
        # Создаем ZIP архив
        zip_path = output_dir.parent / "правила_стандартизации.zip"
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(output_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(output_dir)
                    zipf.write(file_path, arcname)
        
        report = f"""
✅ Правила стандартизации сохранены!

📁 Сохраненные файлы:
• standardization_rules.json - все правила в JSON
• стандартизация_*.xlsx - таблицы для каждого атрибута
• правила_стандартизации.zip - архив со всеми файлами

🎯 Теперь вы можете использовать эти правила для обратной замены в исходных файлах.
        """
        
        return zip_path, report
        
    except Exception as e:
        return None, f"❌ Ошибка сохранения: {str(e)}"

def apply_standardization_to_original(original_zip, rules_zip):
    """Применяет стандартизацию к исходному архиву"""
    if not original_zip or not rules_zip:
        return None, "❌ Загрузите оба архива"
    
    try:
        # Временные директории
        temp_dir = Path(tempfile.mkdtemp())
        rules_dir = temp_dir / "rules"
        output_dir = temp_dir / "output"
        
        # Распаковываем правила
        with zipfile.ZipFile(rules_zip, 'r') as zip_ref:
            zip_ref.extractall(rules_dir)
        
        # Загружаем правила из JSON
        rules_json_path = rules_dir / "standardization_rules.json"
        if not rules_json_path.exists():
            return None, "❌ В архиве правил не найден файл standardization_rules.json"
        
        with open(rules_json_path, 'r', encoding='utf-8') as f:
            rules = json.load(f)
        
        # Применяем замену
        replacer = DataReplacer(rules)
        result_zip = replacer.process_archive(
            Path(original_zip),
            rules,
            output_dir
        )
        
        report = f"""
✅ Стандартизация применена успешно!

📊 Результат:
• Создан архив: {result_zip.name}
• Все исходные файлы обновлены
• Логи изменений сохранены в соответствующих папках

⚠️ ВНИМАНИЕ: В новом архиве все значения заменены на стандартизированные.
        """
        
        # Очистка временных файлов
        shutil.rmtree(rules_dir, ignore_errors=True)
        
        return result_zip, report
        
    except Exception as e:
        return None, f"❌ Ошибка применения стандартизации: {str(e)}"

# Создаем интерфейс
with gr.Blocks(title="🔄 Стандартизатор значений онтологии") as demo:
    gr.Markdown("# 🔄 Стандартизатор значений онтологии ГРМ")
    gr.Markdown("""
    Этот инструмент позволяет:
    1. **Загрузить результаты обработки** онтологии
    2. **Сгруппировать похожие значения** атрибутов
    3. **Выбрать стандартные значения** для каждой группы
    4. **Применить стандартизацию** к исходным файлам
    """)
    
    with gr.Tabs():
        with gr.TabItem("📥 1. Загрузка проекта"):
            gr.Markdown("### Загрузите ZIP-архив с результатами обработки онтологии")
            
            with gr.Row():
                project_zip = gr.File(
                    label="ZIP с результатами обработки",
                    file_types=[".zip"],
                    type="filepath"
                )
                
                load_btn = gr.Button("📂 Загрузить проект", variant="primary")
            
            attribute_dropdown = gr.Dropdown(
                label="Выберите атрибут для стандартизации",
                interactive=True
            )
            
            project_report = gr.Textbox(label="Отчет", lines=5, interactive=False)
            
            # Скрытое поле для хранения временной папки
            temp_dir_state = gr.State()
        
        with gr.TabItem("🎯 2. Стандартизация значений"):
            gr.Markdown("### Выберите стандартные значения для групп")
            
            attribute_name_display = gr.Textbox(
                label="Текущий атрибут",
                interactive=False
            )
            
            values_df = gr.Dataframe(
                label="Значения атрибута",
                headers=["Группа", "Исходное значение", "Предложенное стандартное", "Выбранное стандартное"],
                datatype=["str", "str", "str", "str"],
                col_count=(4, "fixed"),
                interactive=True,
                wrap=True
            )
            
            with gr.Row():
                load_values_btn = gr.Button("📊 Загрузить значения", variant="primary")
                save_rules_btn = gr.Button("💾 Сохранить правила", variant="secondary")
            
            standardization_report = gr.Textbox(label="Отчет", lines=3, interactive=False)
            
            # Скрытое состояние для групп
            groups_state = gr.State()
        
        with gr.TabItem("💾 3. Сохранение правил"):
            gr.Markdown("### Сохраните все правила стандартизации")
            
            save_all_btn = gr.Button("📦 Сохранить все правила", variant="primary", size="lg")
            
            rules_output = gr.File(label="Скачать правила стандартизации")
            save_report = gr.Textbox(label="Отчет", lines=5, interactive=False)
        
        with gr.TabItem("🔄 4. Применение к исходным данным"):
            gr.Markdown("### Примените стандартизацию к исходному архиву")
            
            with gr.Row():
                original_archive = gr.File(
                    label="Исходный ZIP-архив с онтологией",
                    file_types=[".zip"],
                    type="filepath"
                )
                
                standardization_rules_file = gr.File(
                    label="ZIP с правилами стандартизации",
                    file_types=[".zip"],
                    type="filepath"
                )
            
            apply_btn = gr.Button("🔄 Применить стандартизацию", variant="primary", size="lg")
            
            result_archive = gr.File(label="Стандартизированный архив")
            apply_report = gr.Textbox(label="Отчет", lines=5, interactive=False)
    
    # Обработчики событий
    load_btn.click(
        fn=load_standardization_project,
        inputs=[project_zip],
        outputs=[attribute_dropdown, project_report, temp_dir_state]
    )
    
    load_values_btn.click(
        fn=lambda attr: (attr, load_attribute_values(attr)),
        inputs=[attribute_dropdown],
        outputs=[attribute_name_display, values_df, groups_state]
    )
    
    save_rules_btn.click(
        fn=lambda df, groups, attr: update_standard_values(df, groups, attr),
        inputs=[values_df, groups_state, attribute_name_display],
        outputs=[standardization_report]
    )
    
    save_all_btn.click(
        fn=save_standardization_rules,
        inputs=[],
        outputs=[rules_output, save_report]
    )
    
    apply_btn.click(
        fn=apply_standardization_to_original,
        inputs=[original_archive, standardization_rules_file],
        outputs=[result_archive, apply_report]
    )

if __name__ == "__main__":
    demo.launch(
        server_name="0.0.0.0",
        server_port=7861,  # Другой порт, чтобы не конфликтовал с основным app.py
        share=False
    )