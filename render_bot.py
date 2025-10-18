#!/usr/bin/env python3
"""
Упрощенный запуск бота для Render
"""

import os
import sys
import logging

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO,
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)

def check_dependencies():
    """Проверяем наличие всех зависимостей"""
    dependencies = [
        ('python-telegram-bot', 'telegram'),
        ('python-docx', 'docx'),
        ('PIL', 'PIL'),
        ('dotenv', 'dotenv'),
        ('requests', 'requests')
    ]
    
    missing = []
    for package, import_name in dependencies:
        try:
            if import_name == 'telegram':
                import telegram
            elif import_name == 'docx':
                from docx import Document
            elif import_name == 'PIL':
                from PIL import Image
            elif import_name == 'dotenv':
                import dotenv
            elif import_name == 'requests':
                import requests
            print(f"✅ {package} успешно импортирован")
        except ImportError as e:
            print(f"❌ Ошибка импорта {package}: {e}")
            missing.append(package)
    
    return missing

def main():
    print("🚀 Проверка зависимостей...")
    missing = check_dependencies()
    
    if missing:
        print(f"❌ Отсутствуют зависимости: {missing}")
        print("Установите их: pip install " + " ".join(missing))
        return
    
    try:
        print("✅ Все зависимости установлены, запускаем бота...")
        from bot import ActGeneratorBot
        bot = ActGeneratorBot()
        print("🤖 Бот успешно создан, начинаем работу...")
        bot.run()
    except Exception as e:
        print(f"💥 Критическая ошибка: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()