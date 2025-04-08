from pathlib import Path

# Базовые пути
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / 'data'
OUTPUT_DIR = BASE_DIR / 'output'

# Создаем папки если их нет
DATA_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# Форматы сертификатов
CERTIFICATE_EXTENSIONS = ('.pdf', '.jpg', '.jpeg', '.png')

# Проверка существования папок
if not DATA_DIR.exists():
    raise FileNotFoundError(f"Папка с данными не найдена: {DATA_DIR}")

if not OUTPUT_DIR.exists():
    OUTPUT_DIR.mkdir(parents=True)