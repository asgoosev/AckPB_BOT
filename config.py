# Конфигурационные настройки бота
import os
from dotenv import load_dotenv

# Загрузка переменных окружения из .env файла
load_dotenv()

# Настройки Telegram бота
BOT_TOKEN = os.getenv('BOT_TOKEN', 'YOUR_BOT_TOKEN_HERE')

# Настройки email
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', '587'))
EMAIL = os.getenv('EMAIL', 'asgoosev2@gmail.com')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD', 'your_password')

# Основной email
DEFAULT_EMAIL = "asgoosev@gmail.com"

# Данные подразделений (теперь с разделением на должность и ФИО)
departments = {
    "КТЦ": {"head": "начальник КТЦ", "name": "Маркелов А.П."},
    "ЭЦ": {"head": "и.о. начальника ЭЦ", "name": "Сидоров А.П."}, 
    "ЦТАИ": {"head": "начальник ЦТАИ", "name": "Носов Н.А."},
    "ХЦ": {"head": "начальник ХЦ", "name": "Рубинчик С.В."},
    "ТТЦ": {"head": "начальник ТТЦ", "name": "Першин А.А."},
    "ЦЦР": {"head": "начальник ЦЦР", "name": "Гусаков С.Г."},
    "Пом дир": {"head": "Помощник директора", "name": "Перевязкин М.А."}
}

# Типы нарушений
violation_types = [
    "[дата] при проведении обхода территории электростанции, выявлены грубые нарушения правил пожарной безопасности:",
    "[дата] при проведении обхода зданий и сооружений, выявлены грубые нарушения правил пожарной безопасности:", 
    "[дата] при проверке проведения работ повышенной опасности, выявлены грубые нарушения правил пожарной безопасности:"
]

# Константы для состояний разговора
SELECTING_DEPARTMENT, SELECTING_TIME, INPUT_PLACE, SELECTING_VIOLATION_TYPE, \
INPUT_VIOLATIONS, INPUT_WITNESS, UPLOAD_PHOTOS, SELECT_EMAIL, INPUT_EMAIL = range(9)

# Настройки документа
DOCUMENT_SETTINGS = {
    'font_name': 'Arial',
    'font_size': 12,  # pt
    'left_indent': 1.0,  # cm
    'photo_height': 11.0,  # см - уменьшено для размещения 2 фото на странице
    'photos_per_page': 2
}
