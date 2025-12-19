import pandas as pd
import json
from pathlib import Path
from typing import Dict, List, Tuple, Set
import re
from collections import defaultdict

class ValueStandardizer:
    """
    Класс для стандартизации значений атрибутов
    """
    
    def __init__(self, input_dir: Path):
        self.input_dir = Path(input_dir)
        self.standardization_rules = {}
        self.attribute_files = {}
        
    def find_attribute_files(self) -> Dict[str, Path]:
        """Находит все файлы атрибутов в папке 'Вспомогательные справочники'"""
        справочники_dir = self.input_dir / "Вспомогательные справочники"
        attribute_files = {}
        
        if справочники_dir.exists():
            for file_path in справочники_dir.glob("*.xlsx"):
                attribute_name = file_path.stem
                attribute_files[attribute_name] = file_path
        
        return attribute_files
    
    def load_attribute_values(self, file_path: Path) -> List[str]:
        """Загружает значения из файла атрибута"""
        try:
            df = pd.read_excel(file_path)
            if 'Значение' in df.columns:
                values = df['Значение'].dropna().astype(str).unique().tolist()
                return sorted([v.strip() for v in values if v.strip()])
            return []
        except Exception as e:
            print(f"Ошибка загрузки файла {file_path}: {e}")
            return []
    
    def group_similar_values(self, values: List[str], 
                           similarity_threshold: float = 0.7) -> List[List[str]]:
        """
        Группирует похожие значения для удобства стандартизации
        Использует простую эвристику для поиска похожих строк
        """
        groups = []
        used_values = set()
        
        # Сначала группируем по первым словам
        first_word_groups = defaultdict(list)
        for value in values:
            if pd.isna(value):
                continue
            value_str = str(value).strip().lower()
            first_word = value_str.split()[0] if value_str.split() else ""
            first_word_groups[first_word].append(value)
        
        # Затем внутри групп ищем еще более похожие
        for first_word, group_values in first_word_groups.items():
            if len(group_values) == 1:
                groups.append(group_values)
                continue
            
            # Сортируем по длине для удобства
            group_values_sorted = sorted(group_values, key=len)
            current_group = [group_values_sorted[0]]
            
            for value in group_values_sorted[1:]:
                # Проверяем похожесть
                if self._are_values_similar(current_group[0], value, similarity_threshold):
                    current_group.append(value)
                else:
                    groups.append(current_group)
                    current_group = [value]
            
            if current_group:
                groups.append(current_group)
        
        return groups
    
    def _are_values_similar(self, value1: str, value2: str, 
                          threshold: float = 0.7) -> bool:
        """Проверяет, похожи ли два значения"""
        val1 = str(value1).lower().strip()
        val2 = str(value2).lower().strip()
        
        # Если одно значение содержит другое
        if val1 in val2 or val2 in val1:
            return True
        
        # Считаем процент совпадающих слов
        words1 = set(re.findall(r'\w+', val1))
        words2 = set(re.findall(r'\w+', val2))
        
        if not words1 or not words2:
            return False
        
        intersection = words1.intersection(words2)
        similarity = len(intersection) / max(len(words1), len(words2))
        
        return similarity >= threshold
    
    def suggest_standard_value(self, values_group: List[str]) -> str:
        """Предлагает стандартное значение для группы"""
        if not values_group:
            return ""
        
        # 1. Ищем самый короткий вариант
        shortest = min(values_group, key=len)
        
        # 2. Ищем вариант с наибольшим количеством заглавных букв (возможно, аббревиатура)
        def count_uppercase(s):
            return sum(1 for c in s if c.isupper())
        
        most_uppercase = max(values_group, key=count_uppercase)
        
        # 3. Предпочитаем вариант без лишних символов
        clean_values = []
        for value in values_group:
            clean_value = re.sub(r'[^\w\s]', '', value).strip()
            if clean_value:
                clean_values.append(clean_value)
        
        # Возвращаем самый частый вариант или самый короткий
        if clean_values:
            from collections import Counter
            most_common = Counter(clean_values).most_common(1)[0][0]
            return most_common
        
        return shortest
    
    def create_standardization_map(self, attribute_name: str, 
                                 original_to_standard: Dict[str, str]) -> None:
        """Создает карту стандартизации для атрибута"""
        if attribute_name not in self.standardization_rules:
            self.standardization_rules[attribute_name] = {}
        
        self.standardization_rules[attribute_name].update(original_to_standard)
    
    def save_standardization_rules(self, output_dir: Path) -> None:
        """Сохраняет правила стандартизации"""
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Сохраняем в JSON
        json_path = output_dir / "standardization_rules.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(self.standardization_rules, f, ensure_ascii=False, indent=2)
        
        # Сохраняем в Excel для каждого атрибута
        for attr_name, rules in self.standardization_rules.items():
            df = pd.DataFrame(list(rules.items()), 
                            columns=['Исходное значение', 'Стандартное значение'])
            excel_path = output_dir / f"стандартизация_{attr_name}.xlsx"
            df.to_excel(excel_path, index=False)
        
        print(f"✅ Правила стандартизации сохранены в: {output_dir}")
    
    def load_standardization_rules(self, rules_file: Path) -> None:
        """Загружает правила стандартизации из файла"""
        with open(rules_file, 'r', encoding='utf-8') as f:
            self.standardization_rules = json.load(f)