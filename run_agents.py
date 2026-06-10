"""
Точка входа для мультиагентной системы ВВС.

Использование:
  python run_agents.py                        # запустить всех агентов
  python run_agents.py analyst                # только аналитик
  python run_agents.py designer               # только дизайнер
  python run_agents.py designer site          # только сайт
  python run_agents.py designer presentation  # только презентация
  python run_agents.py sales                  # только ассистент (универсальный скрипт)
  python run_agents.py sales "профиль..."     # скрипт под конкретного поставщика

Требует: ANTHROPIC_API_KEY в переменных среды или в файле .env
"""

import sys
import os

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from agents.orchestrator import run_all, run_single


def print_result(title: str, content: str) -> None:
    print(f"\n{'=' * 60}")
    print(f"  {title}")
    print(f"{'=' * 60}\n")
    print(content)


def main():
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("Ошибка: переменная ANTHROPIC_API_KEY не задана.")
        print("Создайте файл .env с содержимым:")
        print("  ANTHROPIC_API_KEY=sk-ant-...")
        sys.exit(1)

    args = sys.argv[1:]

    if not args:
        results = run_all()
        print_result("АНАЛИТИК — Таблица передачи клиентов", results["analyst"])
        print_result("ДИЗАЙНЕР — Сайт и презентация", results["designer"])
        print_result("АССИСТЕНТ ПО ПРОДАЖАМ — Скрипт звонка", results["sales"])

    elif args[0] == "analyst":
        result = run_single("analyst")
        print_result("АНАЛИТИК — Таблица передачи клиентов", result)

    elif args[0] == "designer":
        focus = args[1] if len(args) > 1 else None
        if focus and focus not in ("site", "presentation"):
            print(f"Ошибка: для designer допустимо 'site' или 'presentation', получено: '{focus}'")
            sys.exit(1)
        result = run_single("designer", focus=focus)
        print_result("ДИЗАЙНЕР — Сайт и презентация", result)

    elif args[0] == "sales":
        profile = args[1] if len(args) > 1 else None
        result = run_single("sales", supplier_profile=profile)
        print_result("АССИСТЕНТ ПО ПРОДАЖАМ — Скрипт звонка", result)

    else:
        print(f"Неизвестная команда: '{args[0]}'")
        print("Допустимые команды: analyst, designer, sales (или без аргументов — все агенты)")
        sys.exit(1)


if __name__ == "__main__":
    main()
