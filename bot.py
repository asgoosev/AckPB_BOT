import os
import logging
import smtplib
import requests
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from typing import Dict, List
import tempfile

from telegram import (
    Update, 
    ReplyKeyboardMarkup, 
    InlineKeyboardMarkup, 
    InlineKeyboardButton,
    ReplyKeyboardRemove
)
from telegram.ext import (
    Application, 
    CommandHandler, 
    MessageHandler, 
    ContextTypes, 
    ConversationHandler, 
    CallbackQueryHandler,
    filters
)

try:
    from docx import Document
    from docx.shared import Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    print("Ошибка: python-docx не установлен. Установите его: pip install python-docx")
    DOCX_AVAILABLE = False

try:
    from PIL import Image
    PILLOW_AVAILABLE = True
except ImportError:
    print("Ошибка: Pillow не установлен. Установите его: pip install Pillow")
    PILLOW_AVAILABLE = False

# Импорт конфигурации
import config

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Словарь для месяцев на русском
RUSSIAN_MONTHS = {
    1: "января", 2: "февраля", 3: "марта", 4: "апреля", 
    5: "мая", 6: "июня", 7: "июля", 8: "августа", 
    9: "сентября", 10: "октября", 11: "ноября", 12: "декабря"
}

class ActGeneratorBot:
    def __init__(self):
        self.token = config.BOT_TOKEN
        self.smtp_server = config.SMTP_SERVER
        self.smtp_port = config.SMTP_PORT
        self.email = config.EMAIL
        self.password = config.EMAIL_PASSWORD
        
        if not self.token or self.token == 'YOUR_BOT_TOKEN_HERE':
            raise ValueError("Пожалуйста, установите BOT_TOKEN в файле .env")
        
        # Инициализация приложения
        self.application = Application.builder().token(self.token).build()
        self.setup_handlers()
        
    def setup_handlers(self):
        """Настройка обработчиков команд - ВАЖНО: правильный порядок!"""
        
        # Создаем ConversationHandler как отдельный объект
        conv_handler = ConversationHandler(
            entry_points=[CommandHandler('start', self.start)],
            states={
                config.SELECTING_DEPARTMENT: [
                    CallbackQueryHandler(self.department_selected, pattern='^dept_')
                ],
                config.SELECTING_TIME: [
                    MessageHandler(filters.TEXT & ~filters.COMMAND, self.time_selected)
                ],
                config.INPUT_PLACE: [
                    MessageHandler(filters.TEXT & ~filters.COMMAND, self.place_selected)
                ],
                config.SELECTING_VIOLATION_TYPE: [
                    CallbackQueryHandler(self.violation_type_selected, pattern='^viol_type_'),
                    CallbackQueryHandler(self.navigate_violations, pattern='^nav_')
                ],
                config.INPUT_VIOLATIONS: [
                    MessageHandler(filters.TEXT & ~filters.COMMAND, self.violations_input),
                    CallbackQueryHandler(self.violations_done_callback, pattern='^violations_done$')
                ],
                config.INPUT_WITNESS: [
                    MessageHandler(filters.TEXT & ~filters.COMMAND, self.witness_input)
                ],
                config.UPLOAD_PHOTOS: [
                    MessageHandler(filters.PHOTO, self.photo_uploaded),
                    CallbackQueryHandler(self.photos_finished, pattern='^photos_done$')
                ],
                config.SELECT_EMAIL: [
                    CallbackQueryHandler(self.email_selected, pattern='^email_'),
                    CallbackQueryHandler(self.input_new_email, pattern='^new_email$')
                ],
                config.INPUT_EMAIL: [
                    MessageHandler(filters.TEXT & ~filters.COMMAND, self.email_input)
                ]
            },
            fallbacks=[CommandHandler('cancel', self.cancel)],
            allow_reentry=True
        )
        
        # ВАЖНО: ConversationHandler должен быть добавлен ПЕРВЫМ
        self.application.add_handler(conv_handler)
        
        # Затем добавляем остальные обработчики команд
        self.application.add_handler(CommandHandler('help', self.help_command))
        self.application.add_handler(CommandHandler('newact', self.start))
        self.application.add_handler(CommandHandler('done', self.violations_done))
        self.application.add_handler(CommandHandler('cancel', self.cancel))
        
        # Обработчик для текстовых сообщений вне ConversationHandler
        self.application.add_handler(MessageHandler(
            filters.TEXT & ~filters.COMMAND, 
            self.handle_text_messages
        ))
        
        # Добавим обработчик для неизвестных callback-запросов для отладки
        self.application.add_handler(CallbackQueryHandler(self.unknown_callback, pattern='.*'))
    
    async def unknown_callback(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработчик для неизвестных callback-запросов"""
        query = update.callback_query
        await query.answer()
        logger.warning(f"Неизвестный callback: {query.data}")
        await query.edit_message_text("❌ Произошла ошибка. Пожалуйста, начните заново с /start")
    
    async def handle_text_messages(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработка текстовых сообщений вне ConversationHandler"""
        # Если есть активный разговор, игнорируем сообщение
        if context.user_data.get('active'):
            await update.message.reply_text(
                "Пожалуйста, завершите текущее создание акта или отправьте /cancel для отмены"
            )
            return
        
        await update.message.reply_text(
            "Для начала создания акта отправьте /start или /newact\n"
            "Для справки отправьте /help"
        )
    
    async def help_command(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Справка по командам"""
        await update.message.reply_text(
            "📋 <b>Доступные команды:</b>\n\n"
            "/start - начать создание нового акта\n"
            "/newact - создать новый акт (аналогично /start)\n"
            "/cancel - отменить текущую операцию\n"
            "/done - завершить ввод нарушений\n"
            "/help - показать эту справку",
            parse_mode='HTML'
        )
    
    async def start(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Начало создания нового акта"""
        user = update.message.from_user
        logger.info(f"Пользователь {user.first_name} начал создание акта")
        
        # Проверка зависимостей
        if not DOCX_AVAILABLE:
            await update.message.reply_text(
                "❌ <b>Ошибка!</b>\n\n"
                "Библиотека python-docx не установлена.\n"
                "Запустите в командной строке:\n"
                "<code>pip install python-docx</code>",
                parse_mode='HTML'
            )
            return ConversationHandler.END
        
        if not PILLOW_AVAILABLE:
            await update.message.reply_text(
                "❌ <b>Ошибка!</b>\n\n"
                "Библиотека Pillow не установлена.\n"
                "Запустите в командной строке:\n"
                "<code>pip install Pillow</code>",
                parse_mode='HTML'
            )
            return ConversationHandler.END
        
        # Инициализация данных пользователя
        context.user_data.clear()
        context.user_data['violations'] = []
        context.user_data['photos'] = []
        context.user_data['email_history'] = []
        context.user_data['active'] = True
        
        # Клавиатура для выбора подразделения
        keyboard = []
        for dept in config.departments.keys():
            keyboard.append([InlineKeyboardButton(dept, callback_data=f"dept_{dept}")])
        
        reply_markup = InlineKeyboardMarkup(keyboard)
        
        await update.message.reply_text(
            "🎯 <b>Создание акта о нарушении ПБ</b>\n\n"
            "Выберите подразделение:",
            reply_markup=reply_markup,
            parse_mode='HTML'
        )
        
        return config.SELECTING_DEPARTMENT
    
    async def department_selected(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Обработка выбора подразделения"""
        query = update.callback_query
        await query.answer()
        
        department = query.data.replace('dept_', '')
        context.user_data['department'] = department
        context.user_data['department_head'] = config.departments[department]['head']
        context.user_data['department_head_name'] = config.departments[department]['name']
        
        logger.info(f"Выбрано подразделение: {department}")
        
        # Клавиатура для выбора времени
        time_keyboard = [
            ["09:00-10:00", "10:00-11:00", "11:00-12:00"],
            ["12:00-13:00", "13:00-14:00", "14:00-15:00"],
            ["15:00-16:00", "16:00-17:00", "17:00-18:00"],
            ["Отмена ❌"]
        ]
        
        reply_markup = ReplyKeyboardMarkup(
            time_keyboard, 
            one_time_keyboard=True,
            resize_keyboard=True
        )
        
        await query.edit_message_text(
            f"✅ Выбрано: <b>{department}</b>\n\n"
            "⏰ Выберите время проведения проверки:",
            parse_mode='HTML'
        )
        
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text="Используйте клавиатуру ниже для выбора времени:",
            reply_markup=reply_markup
        )
        
        return config.SELECTING_TIME
    
    async def time_selected(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Обработка выбора времени"""
        time_range = update.message.text
        
        if time_range == "Отмена ❌":
            return await self.cancel(update, context)
        
        context.user_data['time_range'] = time_range
        
        logger.info(f"Выбрано время: {time_range}")
        
        await update.message.reply_text(
            f"✅ Время проверки: <b>{time_range}</b>\n\n"
            "📍 Введите место проведения проверки:\n\n"
            "<i>Например: Главный корпус ПСО, отм. 8,0, ряд А-Б, оси 15-16</i>",
            reply_markup=ReplyKeyboardRemove(),
            parse_mode='HTML'
        )
        
        return config.INPUT_PLACE
    
    async def place_selected(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Обработка ввода места проведения"""
        place = update.message.text
        context.user_data['place'] = place
        
        logger.info(f"Введено место: {place}")
        
        # Показ карусели типов нарушений
        return await self.show_violation_types(update, context)
    
    async def show_violation_types(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Показать карусель типов нарушений"""
        current_index = context.user_data.get('violation_type_index', 0)
        
        keyboard = [
            [
                InlineKeyboardButton("⬅️", callback_data="nav_prev"),
                InlineKeyboardButton(f"{current_index + 1}/{len(config.violation_types)}", callback_data="ignore"),
                InlineKeyboardButton("➡️", callback_data="nav_next")
            ],
            [InlineKeyboardButton("✅ Выбрать", callback_data=f"viol_type_{current_index}")]
        ]
        
        reply_markup = InlineKeyboardMarkup(keyboard)
        
        violation_text = config.violation_types[current_index].replace(
            '[дата]', 
            datetime.now().strftime('%d.%m.%Y')
        )
        
        text = (
            f"✅ Место: <b>{context.user_data['place']}</b>\n\n"
            "📝 <b>Выберите тип нарушения:</b>\n\n"
            f"{violation_text}"
        )
        
        if update.callback_query:
            await update.callback_query.edit_message_text(
                text=text,
                reply_markup=reply_markup,
                parse_mode='HTML'
            )
        else:
            await update.message.reply_text(
                text=text,
                reply_markup=reply_markup,
                parse_mode='HTML'
            )
        
        return config.SELECTING_VIOLATION_TYPE
    
    async def navigate_violations(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Навигация по типам нарушений"""
        query = update.callback_query
        await query.answer()
        
        current_index = context.user_data.get('violation_type_index', 0)
        
        if query.data == "nav_prev":
            current_index = (current_index - 1) % len(config.violation_types)
        elif query.data == "nav_next":
            current_index = (current_index + 1) % len(config.violation_types)
        
        context.user_data['violation_type_index'] = current_index
        
        return await self.show_violation_types(update, context)
    
    async def violation_type_selected(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Обработка выбора типа нарушения"""
        query = update.callback_query
        await query.answer()
        
        violation_index = int(query.data.replace('viol_type_', ''))
        context.user_data['violation_type'] = config.violation_types[violation_index]
        
        logger.info(f"Выбран тип нарушения: {violation_index}")
        
        keyboard = InlineKeyboardMarkup([[
            InlineKeyboardButton("✅ Завершить ввод нарушений", callback_data="violations_done")
        ]])
        
        await query.edit_message_text(
            f"✅ Выбран тип нарушения\n\n"
            "📝 Введите конкретные нарушения (по одному за сообщение).\n\n"
            "<i>Когда закончите, нажмите кнопку ниже или отправьте </i><b>/done</b>",
            reply_markup=keyboard,
            parse_mode='HTML'
        )
        
        return config.INPUT_VIOLATIONS
    
    async def violations_input(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Обработка ввода нарушений"""
        violation_text = update.message.text
        
        # Проверяем, не является ли сообщение командой /done
        if violation_text == '/done':
            return await self.violations_done(update, context)
        
        context.user_data['violations'].append(violation_text)
        
        keyboard = InlineKeyboardMarkup([[
            InlineKeyboardButton("✅ Завершить ввод нарушений", callback_data="violations_done")
        ]])
        
        await update.message.reply_text(
            f"✅ Нарушение добавлено (<b>{len(context.user_data['violations'])}</b>)\n"
            "Введите следующее нарушение или нажмите кнопку для продолжения",
            reply_markup=keyboard,
            parse_mode='HTML'
        )
        
        return config.INPUT_VIOLATIONS
    
    async def violations_done(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Обработка команды /done для завершения ввода нарушений"""
        if not context.user_data.get('violations'):
            await update.message.reply_text(
                "❌ Вы не ввели ни одного нарушения. Пожалуйста, введите хотя бы одно нарушение."
            )
            return config.INPUT_VIOLATIONS
        
        await update.message.reply_text(
            f"✅ Введено нарушений: <b>{len(context.user_data['violations'])}</b>\n\n"
            "👤 Введите данные лица, в присутствии которого проводилась проверка:\n"
            "<b>Формат: Организация, Должность, Фамилия И.О.</b>\n\n"
            "<i>Пример: ООО 'Стройсервис', Инженер, Иванов А.С.</i>",
            parse_mode='HTML'
        )
        
        return config.INPUT_WITNESS

    async def violations_done_callback(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Обработка callback кнопки завершения ввода нарушений"""
        query = update.callback_query
        await query.answer()
        
        if not context.user_data.get('violations'):
            await query.edit_message_text(
                "❌ Вы не ввели ни одного нарушения. Пожалуйста, введите хотя бы одно нарушение."
            )
            return config.INPUT_VIOLATIONS
        
        await query.edit_message_text(
            f"✅ Введено нарушений: <b>{len(context.user_data['violations'])}</b>\n\n"
            "👤 Введите данные лица, в присутствии которого проводилась проверка:\n"
            "<b>Формат: Организация, Должность, Фамилия И.О.</b>\n\n"
            "<i>Пример: ООО 'Стройсервис', Инженер, Иванов А.С.</i>",
            parse_mode='HTML'
        )
        
        return config.INPUT_WITNESS
    
    async def witness_input(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Обработка ввода данных свидетеля"""
        witness_data = update.message.text
        
        try:
            org, position, name = [x.strip() for x in witness_data.split(',', 2)]
            context.user_data['witness_org'] = org
            context.user_data['witness_position'] = position
            context.user_data['witness_name'] = name
        except ValueError:
            await update.message.reply_text(
                "❌ Неверный формат. Используйте: <b>Организация, Должность, Фамилия И.О.</b>\n"
                "Повторите ввод:",
                parse_mode='HTML'
            )
            return config.INPUT_WITNESS
        
        logger.info(f"Введены данные свидетеля: {witness_data}")
        
        keyboard = InlineKeyboardMarkup([[
            InlineKeyboardButton("✅ Завершить загрузку", callback_data="photos_done")
        ]])
        
        await update.message.reply_text(
            f"✅ Данные свидетеля сохранены\n\n"
            "📸 Теперь загрузите фотографии нарушений.\n\n"
            "<b>Вы можете:</b>\n"
            "• Отправлять фото по одному\n"
            "• Отправлять несколько фото сразу (альбомом)\n"
            "• Нажать кнопку для завершения\n\n"
            "Фото будут вставлены в акт с подписями!",
            reply_markup=keyboard,
            parse_mode='HTML'
        )
        
        return config.UPLOAD_PHOTOS
    
    async def photo_uploaded(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Обработка загруженных фотографий"""
        # Обрабатываем multiple photos (альбом)
        if update.message.media_group_id:
            # Это альбом - обрабатываем все фото в сообщении
            if update.message.photo:
                photo = update.message.photo[-1]  # Берем фото наивысшего качества
                context.user_data['photos'].append({
                    'file_id': photo.file_id,
                    'file_unique_id': photo.file_unique_id
                })
            # Для альбомов не отправляем ответ сразу, ждем все фото
            return config.UPLOAD_PHOTOS
        else:
            # Одиночное фото
            photo = update.message.photo[-1]
            context.user_data['photos'].append({
                'file_id': photo.file_id,
                'file_unique_id': photo.file_unique_id
            })
        
        keyboard = InlineKeyboardMarkup([[
            InlineKeyboardButton("✅ Завершить загрузку", callback_data="photos_done")
        ]])
        
        await update.message.reply_text(
            f"✅ Фото сохранено (<b>{len(context.user_data['photos'])}</b>)\n"
            "Продолжайте загружать фото или нажмите <b>Завершить загрузку</b>",
            reply_markup=keyboard,
            parse_mode='HTML'
        )
        
        return config.UPLOAD_PHOTOS
    
    async def photos_finished(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Завершение загрузки фотографий"""
        query = update.callback_query
        await query.answer()
        
        logger.info(f"Загружено фото: {len(context.user_data['photos'])}")
        
        # Показ истории email
        email_history = context.user_data.get('email_history', [])
        
        # Всегда показываем основной email и историю
        keyboard = []
        
        # Основной email
        keyboard.append([InlineKeyboardButton(
            f"📧 {config.DEFAULT_EMAIL} [Основной]", 
            callback_data=f"email_{config.DEFAULT_EMAIL}"
        )])
        
        # История email (последние 5)
        for email in email_history[-5:]:
            keyboard.append([InlineKeyboardButton(email, callback_data=f"email_{email}")])
        
        keyboard.append([InlineKeyboardButton("✉️ Ввести новый email", callback_data="new_email")])
        
        reply_markup = InlineKeyboardMarkup(keyboard)
        
        await query.edit_message_text(
            f"✅ Загружено фото: <b>{len(context.user_data['photos'])}</b>\n\n"
            "📧 Выберите email для отправки акта:",
            reply_markup=reply_markup,
            parse_mode='HTML'
        )
        
        return config.SELECT_EMAIL
    
    async def email_selected(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Обработка выбора email из истории"""
        query = update.callback_query
        await query.answer()
        
        email = query.data.replace('email_', '')
        
        # Обновляем историю email (если это не основной email)
        if email != config.DEFAULT_EMAIL:
            email_history = context.user_data.get('email_history', [])
            if email in email_history:
                email_history.remove(email)
            email_history.append(email)
            context.user_data['email_history'] = email_history[-5:]  # Храним только последние 5
        
        return await self.finalize_act(update, context, email)
    
    async def input_new_email(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Запрос нового email"""
        query = update.callback_query
        await query.answer()
        
        await query.edit_message_text(
            "📧 Введите email для отправки акта:"
        )
        
        return config.INPUT_EMAIL
    
    async def email_input(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Обработка ввода email"""
        email = update.message.text
        
        # Простая валидация email
        if '@' not in email or '.' not in email:
            await update.message.reply_text(
                "❌ Неверный формат email. Пожалуйста, введите корректный email:"
            )
            return config.INPUT_EMAIL
        
        # Обновление истории email (если это не основной email)
        if email != config.DEFAULT_EMAIL:
            email_history = context.user_data.get('email_history', [])
            if email not in email_history:
                email_history.append(email)
                context.user_data['email_history'] = email_history[-5:]  # Храним только последние 5
        
        return await self.finalize_act(update, context, email)
    
    async def finalize_act(self, update: Update, context: ContextTypes.DEFAULT_TYPE, email: str) -> int:
        """Финальная обработка и отправка акта"""
        if update.callback_query:
            message = update.callback_query.message
            await update.callback_query.answer()
        else:
            message = update.message
        
        await message.reply_text("🔄 Формирую DOCX акт с фотографиями...")
        
        try:
            # Создание документа с фото
            doc_path = await self.create_act_document_with_photos(context.user_data, context)
            
            # Отправка по email
            await self.send_email(email, doc_path, context.user_data)
            
            # Очистка временных файлов
            if os.path.exists(doc_path):
                os.remove(doc_path)
            
            # Очистка данных пользователя
            context.user_data.clear()
            
            # Клавиатура для создания нового акта
            keyboard = ReplyKeyboardMarkup([
                ["/newact - Создать новый акт"],
                ["/help - Помощь"]
            ], resize_keyboard=True, one_time_keyboard=True)
            
            await message.reply_text(
                f"✅ <b>DOCX акт с фотографиями успешно отправлен на {email}</b>\n\n"
                "Для создания нового акта используйте кнопку ниже или отправьте /newact",
                reply_markup=keyboard,
                parse_mode='HTML'
            )
            
            logger.info(f"Акт успешно отправлен на {email}")
            
        except Exception as e:
            logger.error(f"Ошибка при создании/отправке акта: {e}")
            await message.reply_text(
                "❌ <b>Произошла ошибка при формировании акта</b>\n\n"
                "Пожалуйста, попробуйте снова, отправив /start\n"
                f"Ошибка: {str(e)}",
                parse_mode='HTML'
            )
        
        return ConversationHandler.END
    
    async def create_act_document_with_photos(self, user_data: Dict, context: ContextTypes.DEFAULT_TYPE) -> str:
        """Создание документа акта с вставленными фотографиями"""
        if not DOCX_AVAILABLE:
            raise ImportError("python-docx не установлен")
        
        doc = Document()
        
        # Настройка стилей
        style = doc.styles['Normal']
        style.font.name = config.DOCUMENT_SETTINGS['font_name']
        style.font.size = Inches(config.DOCUMENT_SETTINGS['font_size'] / 72)
        
        # Настройка полей страницы
        sections = doc.sections
        for section in sections:
            section.top_margin = Cm(2)
            section.bottom_margin = Cm(2)
            section.left_margin = Cm(2)
            section.right_margin = Cm(1)  # Правое поле 1 см
        
        # Получаем текущую дату с русским месяцем
        current_date = datetime.now()
        day = current_date.day
        month = RUSSIAN_MONTHS[current_date.month]
        year = current_date.year
        
        # Страница 1 - заголовок
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("АКТ")
        run.bold = True
        
        # Тип акта
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(self.get_act_type(user_data['violation_type']))
        run.bold = True
        
        # Место
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(user_data['place'])
        run.bold = True
        
        # Дата и организация
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.add_run(f"{current_date.strftime('%d.%m.%Y')}\t\t\tТЭЦ-16 филиал ПАО «Мосэнерго»")
        
        doc.add_paragraph()  # Пустая строка
        
        # Состав комиссии - исправлено: теперь содержит и должность и ФИО
        p = doc.add_paragraph("Комиссия филиала ТЭЦ 16 ПАО «Мосэнерго» в составе:")
        p.paragraph_format.left_indent = Cm(config.DOCUMENT_SETTINGS['left_indent'])
        
        members = [
            "Заместитель ГИ - Гусев А.С.",
            "Инспектор по пожарной безопасности - Кретсу В.",
            f"{user_data['department_head']} - {user_data['department_head_name']}"
        ]
        
        for member in members:
            p = doc.add_paragraph(member)
            p.paragraph_format.left_indent = Cm(config.DOCUMENT_SETTINGS['left_indent'])
        
        # Основной текст
        doc.add_paragraph()
        
        # Абзац 1
        time_range = user_data['time_range']
        p1_text = (
            f"В период с {time_range.split('-')[0]} до {time_range.split('-')[1]} часов "
            f"{day} {month} {year} года провела проверку "
            f"{self.get_inspection_type(user_data['violation_type'])}, {user_data['place']}."
        )
        p = doc.add_paragraph(p1_text)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(config.DOCUMENT_SETTINGS['left_indent'])
        
        # Абзац 2
        p = doc.add_paragraph("В результате осмотра было выявлено грубое нарушение требований пожарной безопасности:")
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(config.DOCUMENT_SETTINGS['left_indent'])
        
        # Нарушения
        for i, violation in enumerate(user_data['violations'], 1):
            p = doc.add_paragraph(f"{i}. {violation}")
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(config.DOCUMENT_SETTINGS['left_indent'])
        
        # Заключительная часть
        witness_full_info = f"{user_data['witness_org']}, {user_data['witness_position']}, {user_data['witness_name']}"
        p = doc.add_paragraph(
            f"Проверка проведена в присутствии {witness_full_info}.\n"
            "По факту нарушений составлен настоящий АКТ."
        )
        p.paragraph_format.left_indent = Cm(config.DOCUMENT_SETTINGS['left_indent'])
        
        # Визовый блок - исправлено: теперь содержит и должность и ФИО
        doc.add_paragraph()
        p = doc.add_paragraph("Комиссия филиала ТЭЦ 16:")
        
        # Подписи - исправлено: теперь содержит и должность и ФИО
        signatures = [
            ("Заместитель главного инженера", "__________________", "Гусев А.С."),
            ("Инспектор по пожарной безопасности", "__________________", "Кретсу В."),
            (f"{user_data['department_head']}", "__________________", f"{user_data['department_head_name']}"),
            (f"Представитель {user_data['witness_org']}:", "", ""),
            (user_data['witness_position'], "__________________", user_data['witness_name'])
        ]
        
        for title, sign, name in signatures:
            p = doc.add_paragraph()
            p.add_run(f"{title}\t\t{sign}\t\t{name}")
        
        # Добавляем фото в приложение
        if user_data['photos']:
            await self.add_photos_to_document(doc, user_data['photos'], context)
        
        # Сохранение документа с правильным названием на латинице
        current_datetime = datetime.now()
        filename = f"act_{current_datetime.strftime('%d.%m.%Y_%H.%M')}.docx"
        doc.save(filename)
        
        return filename

    async def add_photos_to_document(self, doc: Document, photos: List[Dict], context: ContextTypes.DEFAULT_TYPE):
        """Добавление фотографий в документ с подписями"""
        if not photos:
            return
            
        # Добавляем разрыв страницы для приложения
        doc.add_page_break()
        
        # Настройка полей для страниц с фото
        sections = doc.sections
        for section in sections:
            section.top_margin = Cm(1)  # Верхнее поле 1 см
            section.bottom_margin = Cm(1)  # Нижнее поле 1 см
        
        # Заголовок приложения - БЕЗ ПУСТОЙ СТРОКИ ПОСЛЕ
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run("Приложение")
        run.bold = True
        
        # Скачиваем и обрабатываем фото
        temp_files = []
        try:
            for i, photo_info in enumerate(photos, 1):
                file_id = photo_info['file_id']
                
                # Получаем файл фото
                photo_file = await context.bot.get_file(file_id)
                file_url = photo_file.file_path
                
                # Скачиваем фото
                response = requests.get(file_url)
                if response.status_code == 200:
                    # Создаем временный файл
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.jpg') as temp_file:
                        temp_file.write(response.content)
                        temp_files.append((temp_file.name, i))
                else:
                    logger.error(f"Ошибка скачивания фото {i}")
            
            # Добавляем фото в документ по 2 на страницу с уменьшенной высотой
            for i in range(0, len(temp_files), 2):
                # Первое фото
                if i < len(temp_files):
                    temp_file1, photo_num1 = temp_files[i]
                    
                    # Добавляем фото с высотой 11 см
                    try:
                        doc.add_picture(temp_file1, height=Cm(11.0))
                    except Exception as e:
                        logger.error(f"Ошибка вставки фото {photo_num1}: {e}")
                        p = doc.add_paragraph(f"[Фото {photo_num1} - ошибка загрузки]")
                    
                    # Добавляем подпись
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.add_run(f"Фото {photo_num1}")
                
                # Второе фото
                if i + 1 < len(temp_files):
                    doc.add_paragraph()  # Пустой абзац между фото
                    
                    temp_file2, photo_num2 = temp_files[i + 1]
                    
                    # Добавляем фото с высотой 11 см
                    try:
                        doc.add_picture(temp_file2, height=Cm(11.0))
                    except Exception as e:
                        logger.error(f"Ошибка вставки фото {photo_num2}: {e}")
                        p = doc.add_paragraph(f"[Фото {photo_num2} - ошибка загрузки]")
                    
                    # Добавляем подпись
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.add_run(f"Фото {photo_num2}")
                
                # Добавляем разрыв страницы после каждой пары фото (кроме последней)
                if i + 2 < len(temp_files):
                    doc.add_page_break()
                    # Заголовок приложения на новой странице - БЕЗ ПУСТОЙ СТРОКИ
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    run = p.add_run("Приложение")
                    run.bold = True
                    
        finally:
            # Удаляем временные файлы
            for temp_file, _ in temp_files:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
    
    def get_act_type(self, violation_type: str) -> str:
        if "территории электростанции" in violation_type:
            return "Комиссионного осмотра территории электростанции"
        elif "зданий и сооружений" in violation_type:
            return "Комиссионного осмотра зданий и сооружений"
        elif "работ повышенной опасности" in violation_type:
            return "Комиссионной проверки проведения работ повышенной опасности"
        return "Комиссионного осмотра"
    
    def get_inspection_type(self, violation_type: str) -> str:
        if "территории электростанции" in violation_type:
            return "территории электростанции"
        elif "зданий и сооружений" in violation_type:
            return "зданий и сооружений электростанции"
        elif "работ повышенной опасности" in violation_type:  # ИСПРАВЛЕНО: заменил "в" на "in"
            return "проведения работ повышенной опасности"
        return ""
    
    async def send_email(self, to_email: str, file_path: str, user_data: Dict):
        """Отправка акта по email"""
        msg = MIMEMultipart()
        msg['From'] = self.email
        msg['To'] = to_email
        msg['Subject'] = f"Акт проверки {datetime.now().strftime('%d.%m.%Y')}"
        
        body = f"Акт проверки от {datetime.now().strftime('%d.%m.%Y')}"
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # Прикрепление файла
        with open(file_path, "rb") as attachment:
            part = MIMEBase('application', 'vnd.openxmlformats-officedocument.wordprocessingml.document')
            part.set_payload(attachment.read())
        
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename="{os.path.basename(file_path)}"'
        )
        msg.attach(part)
        
        # Отправка
        try:
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.email, self.password)
            server.send_message(msg)
            server.quit()
        except Exception as e:
            logger.error(f"Ошибка отправки email: {e}")
            raise
    
    async def cancel(self, update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
        """Отмена текущей операции"""
        await update.message.reply_text(
            'Операция отменена. Для начала нового акта отправьте /start или /newact',
            reply_markup=ReplyKeyboardRemove()
        )
        context.user_data.clear()
        return ConversationHandler.END
    
    def run(self):
        """Запуск бота"""
        logger.info("Бот запущен")
        self.application.run_polling()

if __name__ == "__main__":
    try:
        bot = ActGeneratorBot()
        bot.run()
    except Exception as e:
        logger.error(f"Ошибка запуска бота: {e}")
        print(f"Ошибка: {e}")
