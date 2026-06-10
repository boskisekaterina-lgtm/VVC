"""
Оркестратор мультиагентной системы ВВС.
Координирует запуск трёх агентов и сохраняет результаты.
"""

import os
import datetime
from .analyst_agent import run_analyst
from .designer_agent import run_designer
from .sales_assistant import run_sales_assistant


def save_output(name: str, content: str, output_dir: str = "output") -> str:
    """Сохраняет результат агента в файл и возвращает путь."""
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{output_dir}/{name}_{timestamp}.md"
    with open(filename, "w", encoding="utf-8") as f:
        f.write(content)
    return filename


def run_all(
    networks_data: str | None = None,
    designer_focus: str | None = None,
    supplier_profile: str | None = None,
    save: bool = True,
) -> dict[str, str]:
    """
    Запускает всех трёх агентов последовательно.

    Args:
        networks_data: Данные о сетях для агента-аналитика.
        designer_focus: Фокус дизайнера — 'site', 'presentation' или None (оба).
        supplier_profile: Профиль поставщика для адаптации скрипта продаж.
        save: Сохранять ли результаты в файлы.

    Returns:
        Словарь с результатами каждого агента.
    """
    results = {}

    print("=" * 60)
    print("МУЛЬТИАГЕНТНАЯ СИСТЕМА ВВС")
    print("=" * 60)

    print("\n[1/3] Агент-аналитик: анализ передачи клиентов")
    print("-" * 40)
    analyst_result = run_analyst(networks_data)
    results["analyst"] = analyst_result
    if save:
        path = save_output("analyst_transfer_table", analyst_result)
        print(f"Сохранено: {path}")

    print("\n[2/3] Агент-дизайнер: сайт и презентация")
    print("-" * 40)
    designer_result = run_designer(designer_focus)
    results["designer"] = designer_result
    if save:
        path = save_output("designer_content", designer_result)
        print(f"Сохранено: {path}")

    print("\n[3/3] Агент-ассистент по продажам: скрипт звонка")
    print("-" * 40)
    sales_result = run_sales_assistant(supplier_profile)
    results["sales"] = sales_result
    if save:
        path = save_output("sales_script", sales_result)
        print(f"Сохранено: {path}")

    print("\n" + "=" * 60)
    print("Все агенты завершили работу.")
    print("=" * 60)

    return results


def run_single(agent: str, **kwargs) -> str:
    """
    Запускает одного агента по имени.

    Args:
        agent: 'analyst', 'designer' или 'sales'.
        **kwargs: Аргументы для соответствующего агента.

    Returns:
        Результат агента.
    """
    if agent == "analyst":
        return run_analyst(kwargs.get("networks_data"))
    elif agent == "designer":
        return run_designer(kwargs.get("focus"))
    elif agent == "sales":
        return run_sales_assistant(kwargs.get("supplier_profile"))
    else:
        raise ValueError(f"Неизвестный агент: '{agent}'. Допустимые: analyst, designer, sales")
