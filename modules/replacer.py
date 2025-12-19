import pandas as pd
from pathlib import Path
import json
from typing import Dict, List
import zipfile
import tempfile
import shutil

class DataReplacer:
    """
    Класс для обратной замены значений в исходных файлах ПредЗап
    """
    
    def __init__(self, standardization_rules: Dict):
        self.standardization_rules = standardization_rules
    
    def find_record_files(self, extracted_dir: Path) -> List[Path]:
        """Находит все файлы ПредЗап в распакованной структуре"""
        record_files = []
        
        for pattern in ["*ПредЗап*.xlsx", "*предзап*.xlsx", "*записи*.xlsx"]:
            found_files = list(extracted_dir.rglob(pattern))
            record_files.extend(found_files)
        
        return record_files
    
    def replace_values_in_file(self, file_path: Path, 
                             standardization_rules: Dict) -> bool:
        """
        Заменяет значения в файле ПредЗап согласно правилам стандартизации
        """
        try:
            # Читаем Excel файл
            df = pd.read_excel(file_path)
            
            changes_made = False
            changes_log = []
            
            # Для каждого столбца проверяем, есть ли правила для этого атрибута
            for column in df.columns:
                column_str = str(column).strip()
                
                # Ищем правила для этого атрибута
                rules_for_column = None
                for attr_name, rules in standardization_rules.items():
                    if attr_name in column_str or column_str in attr_name:
                        rules_for_column = rules
                        break
                
                if rules_for_column:
                    # Заменяем значения в столбце
                    original_values = df[column].astype(str).tolist()
                    df[column] = df[column].apply(
                        lambda x: self._replace_single_value(x, rules_for_column)
                    )
                    
                    # Логируем изменения
                    new_values = df[column].astype(str).tolist()
                    for i, (orig, new) in enumerate(zip(original_values, new_values)):
                        if orig != new:
                            changes_made = True
                            changes_log.append({
                                'file': file_path.name,
                                'column': column,
                                'row': i + 2,  # +2 т.к. Excel нумерует с 1 и есть заголовок
                                'original': orig,
                                'standardized': new
                            })
            
            if changes_made:
                # Сохраняем измененный файл
                df.to_excel(file_path, index=False)
                print(f"✅ Изменения сохранены в: {file_path.name}")
                
                # Сохраняем лог изменений
                log_file = file_path.parent / f"log_изменений_{file_path.stem}.xlsx"
                pd.DataFrame(changes_log).to_excel(log_file, index=False)
            
            return changes_made
            
        except Exception as e:
            print(f"❌ Ошибка обработки файла {file_path}: {e}")
            return False
    
    def _replace_single_value(self, value, rules: Dict[str, str]) -> str:
        """Заменяет одно значение согласно правилам"""
        if pd.isna(value):
            return value
        
        value_str = str(value).strip()
        
        # Прямое совпадение
        if value_str in rules:
            return rules[value_str]
        
        # Частичное совпадение (если значение содержит ключ)
        for original, standard in rules.items():
            if original.lower() in value_str.lower():
                # Заменяем только совпадающую часть
                # Это сложная логика, можно упростить
                return standard
        
        # Если не нашли совпадение, возвращаем оригинал
        return value_str
    
    def process_archive(self, input_zip_path: Path, 
                       standardization_rules: Dict,
                       output_dir: Path) -> Path:
        """
        Обрабатывает исходный архив: заменяет значения и создает новый архив
        """
        # Создаем временную директорию
        temp_dir = Path(tempfile.mkdtemp())
        extracted_dir = temp_dir / "extracted"
        
        try:
            # Распаковываем архив
            with zipfile.ZipFile(input_zip_path, 'r') as zip_ref:
                zip_ref.extractall(extracted_dir)
            
            # Находим и обрабатываем все файлы ПредЗап
            record_files = self.find_record_files(extracted_dir)
            total_changes = 0
            
            print(f"🔍 Найдено файлов для обработки: {len(record_files)}")
            
            for file_path in record_files:
                print(f"🔄 Обработка: {file_path.relative_to(extracted_dir)}")
                if self.replace_values_in_file(file_path, standardization_rules):
                    total_changes += 1
            
            # Создаем новый архив с измененными файлами
            output_zip_path = output_dir / f"стандартизированный_{input_zip_path.name}"
            
            with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(extracted_dir):
                    for file in files:
                        file_path = Path(root) / file
                        arcname = file_path.relative_to(extracted_dir)
                        zipf.write(file_path, arcname)
            
            print(f"✅ Создан стандартизированный архив: {output_zip_path}")
            print(f"📊 Изменено файлов: {total_changes}/{len(record_files)}")
            
            return output_zip_path
            
        finally:
            # Очищаем временные файлы
            shutil.rmtree(temp_dir, ignore_errors=True)