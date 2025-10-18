#!/usr/bin/env python3
"""
Простой скрипт для запуска бота на Render
"""

import os
import logging

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

def main():
    try:
        from bot import ActGeneratorBot
        bot = ActGeneratorBot()
        print("✅ Бот успешно импортирован")
        bot.run()
    except ImportError as e:
        print(f"❌ Ошибка импорта: {e}")
        print("Проверьте установлены ли все зависимости")
    except Exception as e:
        print(f"❌ Ошибка запуска бота: {e}")

if __name__ == "__main__":
    print("🚀 Запуск бота...")
    main()