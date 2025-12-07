import pandas as pd
from datetime import datetime, timedelta
import sys

def get_week_start():
    """Возвращает дату начала текущей недели (понедельник)"""
    today = datetime.now().date()
    # weekday() возвращает 0 для понедельника, 6 для воскресенья
    days_since_monday = today.weekday()
    week_start = today - timedelta(days=days_since_monday)
    return week_start

def is_current_week(date_value):
    """Проверяет, попадает ли дата в текущую неделю"""
    if pd.isna(date_value):
        return False
    
    try:
        date_obj = pd.to_datetime(date_value).date()
        week_start = get_week_start()
        today = datetime.now().date()
        return week_start <= date_obj <= today
    except:
        return False

def process_contractors(main_file, operations_file, groups_file, output_file=None):
    """
    Обрабатывает таблицы контрагентов и создает итоговую таблицу
    
    Параметры:
    - main_file: путь к файлу с основной таблицей (Дебит, Кредит)
    - operations_file: путь к файлу с операциями (приход/расход)
    - groups_file: путь к файлу с группами контрагентов
    - output_file: путь к выходному файлу (по умолчанию 'result.xlsx')
    """
    
    # Загрузка таблиц
    print("Загрузка таблиц...")
    try:
        df_main = pd.read_excel(main_file) if main_file.endswith('.xlsx') or main_file.endswith('.xls') else pd.read_csv(main_file)
        df_ops = pd.read_excel(operations_file) if operations_file.endswith('.xlsx') or operations_file.endswith('.xls') else pd.read_csv(operations_file)
        df_groups = pd.read_excel(groups_file) if groups_file.endswith('.xlsx') or groups_file.endswith('.xls') else pd.read_csv(groups_file)
    except Exception as e:
        print(f"Ошибка при загрузке файлов: {e}")
        return
    
    # Нормализация названий колонок (убираем лишние пробелы)
    df_main.columns = df_main.columns.str.strip()
    df_ops.columns = df_ops.columns.str.strip()
    df_groups.columns = df_groups.columns.str.strip()
    
    print(f"Начало недели: {get_week_start()}")
    print(f"Сегодня: {datetime.now().date()}")
    
    # Фильтрация по текущей неделе
    print("\nФильтрация данных по текущей неделе...")
    df_main_filtered = df_main[df_main["Дата"].apply(is_current_week)].copy()
    df_ops_filtered = df_ops[df_ops["Дата"].apply(is_current_week)].copy()
    
    print(f"Строк в основной таблице (текущая неделя): {len(df_main_filtered)}")
    print(f"Строк в таблице операций (текущая неделя): {len(df_ops_filtered)}")
    
    # Расчет сумм по основной таблице (Дебит - Кредит)
    print("\nРасчет сумм по основной таблице...")
    if len(df_main_filtered) > 0:
        main_sums = df_main_filtered.groupby("Id контаргента").apply(
            lambda x: x["Дебит"].fillna(0).sum() - x["Кредит"].fillna(0).sum()
        ).to_dict()
    else:
        main_sums = {}
    
    # Расчет сумм по операциям (Приход - Расход)
    print("Расчет сумм по операциям...")
    if len(df_ops_filtered) > 0:
        def calculate_operation_sum(group):
            pr = group[group["Операция"] == "приход"]["Сумма"].fillna(0).sum()
            rs = group[group["Операция"] == "расход"]["Сумма"].fillna(0).sum()
            return pr - rs
        
        ops_sums = df_ops_filtered.groupby("Id контаргента").apply(calculate_operation_sum).to_dict()
    else:
        ops_sums = {}
    
    # Объединение всех Id контрагентов
    all_contractor_ids = set(list(main_sums.keys()) + list(ops_sums.keys()))
    
    # Создание словаря для итоговых сумм по контрагентам
    contractor_totals = {}
    for contractor_id in all_contractor_ids:
        main_sum = main_sums.get(contractor_id, 0)
        ops_sum = ops_sums.get(contractor_id, 0)
        contractor_totals[contractor_id] = main_sum + ops_sum
    
    # Создание словаря Id контрагента -> Группа
    print("\nСоздание маппинга контрагентов к группам...")
    id_to_group = {}
    id_to_name = {}
    
    for _, row in df_groups.iterrows():
        contractor_id = row["Id контаргента"]
        group = row["Группа"]
        name = row["Имя контрагента"]
        id_to_group[contractor_id] = group
        id_to_name[contractor_id] = name
    
    # Группировка по группам
    print("Группировка контрагентов по группам...")
    group_data = {}
    
    for contractor_id, total_sum in contractor_totals.items():
        group = id_to_group.get(contractor_id, "Без группы")
        name = id_to_name.get(contractor_id, f"Контрагент {contractor_id}")
        
        if group not in group_data:
            group_data[group] = {
                'sum': 0,
                'contractors': []
            }
        
        group_data[group]['sum'] += total_sum
        group_data[group]['contractors'].append(name)
    
    # Создание итоговой таблицы
    print("\nСоздание итоговой таблицы...")
    result_data = []
    
    for group, data in sorted(group_data.items()):
        result_data.append({
            'Группа': group,
            'Итоговая сумма': round(data['sum'], 2),
            'Контрагенты': ', '.join(sorted(set(data['contractors'])))
        })
    
    df_result = pd.DataFrame(result_data)
    
    # Сохранение результата
    if output_file is None:
        output_file = 'result_contractors.xlsx'
    
    print(f"\nСохранение результата в {output_file}...")
    df_result.to_excel(output_file, index=False)
    
    print(f"\nГотово! Создано групп: {len(df_result)}")
    print("\nРезультат:")
    print(df_result.to_string(index=False))
    
    return df_result

if __name__ == "__main__":
    # Если файлы переданы как аргументы командной строки
    if len(sys.argv) >= 4:
        main_file = sys.argv[1]
        operations_file = sys.argv[2]
        groups_file = sys.argv[3]
        output_file = sys.argv[4] if len(sys.argv) > 4 else None
        process_contractors(main_file, operations_file, groups_file, output_file)
    else:
        # Интерактивный режим
        print("Обработка таблиц контрагентов")
        print("=" * 50)
        
        main_file = input("Введите путь к основной таблице (с Дебит/Кредит): ").strip()
        operations_file = input("Введите путь к таблице операций (приход/расход): ").strip()
        groups_file = input("Введите путь к таблице групп: ").strip()
        output_file = input("Введите путь к выходному файлу (Enter для result_contractors.xlsx): ").strip()
        
        if not output_file:
            output_file = None
        
        process_contractors(main_file, operations_file, groups_file, output_file)

