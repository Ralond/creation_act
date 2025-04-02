import os
from pathlib import Path

# Базовые пути
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / 'data'
OUTPUT_DIR = BASE_DIR / 'output'

# Создаем папки если их нет
DATA_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# Настройки Excel
DEFAULT_TEMPLATE = DATA_DIR / 'Шаблон.xlsx'
DEFAULT_SOURCE = DATA_DIR / 'Исходные.xlsx'
DEFAULT_REGISTER = DATA_DIR / 'Реестр.xlsx'

# Форматы сертификатов
CERTIFICATE_EXTENSIONS = ('.pdf', '.jpg', '.jpeg', '.png')

# Проверка существования файлов
REQUIRED_FILES = {
    'template': DEFAULT_TEMPLATE,
    'source': DEFAULT_SOURCE,
    'register': DEFAULT_REGISTER
}

for name, path in REQUIRED_FILES.items():
    if not path.exists():
        raise FileNotFoundError(f"Не найден обязательный файл: {path}")