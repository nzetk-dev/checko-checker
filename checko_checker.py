import json
import logging
import sys
import time
import sqlite3
import csv
import hashlib
from typing import Optional, Dict, List, Any
from dataclasses import dataclass, asdict
from pathlib import Path

try:
    import requests
except ImportError:
    print("Необходимо установить библиотеку requests: pip install requests")
    sys.exit(1)

try:
    from cryptography.fernet import Fernet
except ImportError:
    print("Необходимо установить библиотеку cryptography: pip install cryptography")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
except ImportError:
    openpyxl = None

# ==================== КОНФИГУРАЦИЯ ====================
CONFIG_FILE = "config.json"
LOG_FILE = "checko_checker.log"
CACHE_DB = "cache.db"
HISTORY_DB = "history.db"
KEY_FILE = "secret.key"
LAST_COMPANY_RESPONSE = "last_company_response.json"
LAST_SEARCH_RESPONSE = "last_search_response.json"
SHOW_ZERO_FINANCES = False  # True - показывать все строки финансовой отчётности, включая нули

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


# ==================== ШИФРОВАНИЕ ====================
class KeyManager:
    """Управление шифрованием API-ключа"""

    def __init__(self):
        self.key_file = Path(KEY_FILE)
        self.key = None

    def generate_key(self):
        self.key = Fernet.generate_key()
        with open(self.key_file, 'wb') as f:
            f.write(self.key)
        logger.info("Создан новый ключ шифрования")

    def load_key(self):
        if not self.key_file.exists():
            self.generate_key()
        else:
            with open(self.key_file, 'rb') as f:
                self.key = f.read()

    def encrypt(self, data: str) -> str:
        self.load_key()
        return Fernet(self.key).encrypt(data.encode()).decode()

    def decrypt(self, encrypted_data: str) -> str:
        self.load_key()
        return Fernet(self.key).decrypt(encrypted_data.encode()).decode()


# ==================== КЭШ ====================
class CacheManager:
    """Кэширование запросов для экономии лимитов"""

    def __init__(self):
        self.conn = sqlite3.connect(CACHE_DB)
        self._create_table()

    def _create_table(self):
        self.conn.execute('''
            CREATE TABLE IF NOT EXISTS cache (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                request_hash TEXT UNIQUE,
                endpoint TEXT,
                response TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        self.conn.commit()

    @staticmethod
    def _get_hash(endpoint: str, params: Dict) -> str:
        # API-ключ намеренно исключён из хэша
        clean = {k: v for k, v in params.items() if k != 'key'}
        request_str = f"{endpoint}:{str(sorted(clean.items()))}"
        return hashlib.md5(request_str.encode()).hexdigest()

    def get(self, endpoint: str, params: Dict) -> Optional[Dict]:
        cursor = self.conn.execute(
            'SELECT response FROM cache WHERE request_hash = ?',
            (self._get_hash(endpoint, params),)
        )
        result = cursor.fetchone()
        if result:
            logger.info(f"Данные получены из кэша: {endpoint}")
            return json.loads(result[0])
        return None

    def set(self, endpoint: str, params: Dict, response: Dict):
        self.conn.execute(
            'INSERT OR REPLACE INTO cache (request_hash, endpoint, response) VALUES (?, ?, ?)',
            (self._get_hash(endpoint, params), endpoint, json.dumps(response, ensure_ascii=False))
        )
        self.conn.commit()

    def clear_old(self, days: int = 7):
        self.conn.execute(
            'DELETE FROM cache WHERE created_at < datetime("now", "-{} days")'.format(days)
        )
        self.conn.commit()

    def close(self):
        self.conn.close()


# ==================== ИСТОРИЯ ====================
class HistoryManager:
    """Хранение истории проверок"""

    def __init__(self):
        self.conn = sqlite3.connect(HISTORY_DB)
        self._create_table()

    def _create_table(self):
        self.conn.execute('''
            CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                inn TEXT,
                company_name TEXT,
                ogrn TEXT,
                status TEXT,
                director TEXT,
                address TEXT,
                okved TEXT,
                checked_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        self.conn.commit()

    def add_record(self, company_data: Dict):
        self.conn.execute('''
            INSERT INTO history (inn, company_name, ogrn, status, director, address, okved)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (
            company_data.get('inn', ''),
            company_data.get('name', ''),
            company_data.get('ogrn', ''),
            company_data.get('status', ''),
            company_data.get('director', ''),
            company_data.get('address', ''),
            json.dumps(company_data.get('okved_list', []), ensure_ascii=False)
        ))
        self.conn.commit()

    def get_all(self) -> List[Dict]:
        cursor = self.conn.execute(
            'SELECT inn, company_name, ogrn, status, director, address, okved, checked_at '
            'FROM history ORDER BY checked_at DESC'
        )
        columns = ['inn', 'company_name', 'ogrn', 'status', 'director', 'address', 'okved', 'checked_at']
        return [dict(zip(columns, row)) for row in cursor.fetchall()]

    def close(self):
        self.conn.close()


# ==================== ЭКСПОРТ ====================
class ExportManager:
    """Экспорт данных в CSV и Excel"""

    @staticmethod
    def to_csv(data: List[Dict], filename: str = "export.csv"):
        if not data:
            logger.warning("Нет данных для экспорта")
            return
        fieldnames = ['inn', 'company_name', 'ogrn', 'status', 'director', 'address', 'okved', 'checked_at']
        with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(data)
        logger.info(f"Данные экспортированы в {filename}")
        print(f"Данные экспортированы в {filename}")

    @staticmethod
    def to_excel(data: List[Dict], filename: str = "export.xlsx"):
        if not data:
            logger.warning("Нет данных для экспорта")
            return
        if openpyxl is None:
            print("Библиотека openpyxl не установлена. Экспорт в Excel невозможен.")
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "История проверок"
        headers = ['ИНН', 'Название', 'ОГРН', 'Статус', 'Директор', 'Адрес', 'ОКВЭД', 'Дата проверки']
        ws.append(headers)
        header_font = Font(bold=True)
        header_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
        for record in data:
            okved = record.get('okved', '')
            if isinstance(okved, str):
                try:
                    okved = ', '.join(json.loads(okved))
                except Exception:
                    pass
            ws.append([
                record.get('inn', ''), record.get('company_name', ''), record.get('ogrn', ''),
                record.get('status', ''), record.get('director', ''), record.get('address', ''),
                okved, record.get('checked_at', '')
            ])
        for column in ws.columns:
            max_length = max((len(str(cell.value)) for cell in column if cell.value is not None), default=0)
            ws.column_dimensions[column[0].column_letter].width = min(max_length + 2, 50)
        wb.save(filename)
        logger.info(f"Данные экспортированы в {filename}")
        print(f"Данные экспортированы в {filename}")


# ==================== МОДЕЛЬ ДАННЫХ ====================
@dataclass
class CompanyData:
    inn: str
    ogrn: str
    name: str
    kpp: str
    address: str
    director: str
    okved_list: List[str]
    registration_date: str
    status: str
    contacts: Dict[str, Any]
    finances: Optional[Dict[str, Any]] = None


# ==================== API ====================
class CheckoAPI:
    """Класс для работы с API Checko.ru"""

    BASE_URL = "https://api.checko.ru/v2"

    def __init__(self, api_key: str, cache: CacheManager):
        self.api_key = api_key
        self.cache = cache
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'CheckoChecker/2.1',
            'Accept': 'application/json'
        })

    # ---------- служебные ----------
    @staticmethod
    def _dump_raw(response: Any, filename: str):
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(response, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.debug(f"Не удалось сохранить сырой ответ: {e}")

    @staticmethod
    def _deep_find(data: Any, key: str, _depth: int = 0) -> Any:
        """Рекурсивный поиск первого вхождения ключа в dict/list любой вложенности"""
        if _depth > 5 or data is None:
            return None
        if isinstance(data, dict):
            if key in data and data[key] not in (None, '', [], {}):
                return data[key]
            for v in data.values():
                found = CheckoAPI._deep_find(v, key, _depth + 1)
                if found is not None:
                    return found
        elif isinstance(data, list):
            for v in data:
                found = CheckoAPI._deep_find(v, key, _depth + 1)
                if found is not None:
                    return found
        return None

    def _make_request(self, endpoint: str, params: Dict[str, Any]) -> Dict[str, Any]:
        cached = self.cache.get(endpoint, params)
        if cached:
            return cached

        url = f"{self.BASE_URL}/{endpoint}"
        full_params = dict(params)
        full_params['key'] = self.api_key

        try:
            response = self.session.get(url, params=full_params, timeout=30)

            if response.status_code == 429:
                logger.warning("Превышен лимит запросов. Ожидание 1 секунды...")
                time.sleep(1)
                return self._make_request(endpoint, params)

            response.raise_for_status()
            data = response.json()

            if isinstance(data, dict) and 'error' in data:
                logger.error(f"Ошибка API: {data['error']}")
                raise Exception(f"Ошибка API: {data['error']}")

            self.cache.set(endpoint, params, data)
            return data

        except requests.exceptions.Timeout:
            logger.error("Превышено время ожидания запроса")
            raise
        except requests.exceptions.ConnectionError:
            logger.error("Ошибка подключения к серверу")
            raise
        except requests.exceptions.RequestException as e:
            logger.error(f"Ошибка запроса: {e}")
            raise

    # ---------- парсинг карточки компании ----------
    def _parse_company(self, data: Dict[str, Any]) -> CompanyData:
        containers = [data]
        for key in ('org', 'company', 'organization', 'ul', 'info', 'legal', 'requisites', 'data'):
            if isinstance(data.get(key), dict):
                containers.append(data[key])

        def pick(*names):
            for c in containers:
                for n in names:
                    v = c.get(n)
                    if v not in (None, '', [], {}):
                        return v
            for n in names:
                v = self._deep_find(data, n)
                if v not in (None, '', [], {}):
                    return v
            return ''

        inn = str(pick('inn'))
        ogrn = str(pick('ogrn'))
        kpp = str(pick('kpp'))
        name = str(pick('name', 'name_full', 'fullname', 'full_name', 'title', 'short_name'))
        status = str(pick('status', 'state', 'activity', 'state_status'))
        reg_date = str(pick('reg_date', 'registration_date', 'ogrn_date', 'date_registration', 'created'))

        addr_raw = pick('address', 'address_full', 'addr')
        if isinstance(addr_raw, dict):
            address = str(addr_raw.get('value') or addr_raw.get('text') or addr_raw.get('address') or '')
        else:
            address = str(addr_raw)

        mgr_raw = pick('manager', 'director', 'head', 'ceo', 'manager_name')
        if isinstance(mgr_raw, dict):
            director = str(mgr_raw.get('name') or mgr_raw.get('fio') or '')
        elif isinstance(mgr_raw, list) and mgr_raw:
            first = mgr_raw[0]
            director = str(first.get('name') or first.get('fio')) if isinstance(first, dict) else str(first)
        else:
            director = str(mgr_raw)

        okved_raw = pick('okved', 'okveds', 'activities')
        okved_list: List[str] = []
        if isinstance(okved_raw, list):
            for item in okved_raw:
                if isinstance(item, dict):
                    code = item.get('code') or item.get('okved') or ''
                    nm = item.get('name') or item.get('title') or item.get('text') or ''
                    okved_list.append(f"{code} {nm}".strip())
                elif isinstance(item, str):
                    okved_list.append(item)
        elif isinstance(okved_raw, dict):
            okved_list.append(f"{okved_raw.get('code', '')} {okved_raw.get('name', '')}".strip())

        contacts_raw = pick('contacts', 'phones', 'emails')
        contacts = contacts_raw if isinstance(contacts_raw, dict) else {}

        if not inn and not name:
            logger.warning(
                "Не удалось извлечь базовые поля. Сырой ответ сохранён в "
                f"{LAST_COMPANY_RESPONSE} - пришлите этот файл для точной настройки парсинга."
            )

        return CompanyData(
            inn=inn, ogrn=ogrn, name=name, kpp=kpp, address=address,
            director=director, okved_list=okved_list, registration_date=reg_date,
            status=status, contacts=contacts
        )

    # ---------- публичные методы ----------
    def get_company(self, inn: str) -> Optional[CompanyData]:
        logger.info(f"Запрос данных компании с ИНН: {inn}")
        try:
            company_data = self._make_request('company', {'inn': inn})
            self._dump_raw(company_data, LAST_COMPANY_RESPONSE)

            if not isinstance(company_data, dict) or 'data' not in company_data:
                logger.warning(f"Компания с ИНН {inn} не найдена")
                return None

            company = self._parse_company(company_data['data'])

            try:
                finances_data = self._make_request('finances', {'inn': inn})
                if isinstance(finances_data, dict) and 'data' in finances_data:
                    company.finances = finances_data['data']
            except Exception as e:
                logger.warning(f"Не удалось получить финансовые данные: {e}")

            return company
        except Exception as e:
            logger.error(f"Ошибка при получении данных компании {inn}: {e}")
            return None

    def search_company(self, query: str) -> List[Dict[str, Any]]:
        logger.info(f"Поиск компаний по запросу: {query}")
        try:
            response = self._make_request('search', {'by': 'name', 'obj': 'org', 'query': query})
            self._dump_raw(response, LAST_SEARCH_RESPONSE)
            return self._normalize_search_data(response.get('data') if isinstance(response, dict) else None)
        except Exception as e:
            logger.error(f"Ошибка при поиске: {e}")
            return []

    @staticmethod
    def _normalize_search_data(data: Any) -> List[Dict[str, Any]]:
        if isinstance(data, list):
            return data
        if isinstance(data, dict):
            for key in ('items', 'results', 'list', 'orgs', 'organizations', 'data'):
                value = data.get(key)
                if isinstance(value, list):
                    return value
            if data.get('inn'):
                return [data]
            return []
        return []

    @staticmethod
    def validate_inn(inn: str) -> bool:
        if not inn.isdigit() or len(inn) != 10:
            return False
        weights = [2, 4, 10, 3, 5, 9, 4, 6, 8]
        checksum = sum(int(inn[i]) * weights[i] for i in range(9)) % 11 % 10
        return checksum == int(inn[9])


# ==================== ПРИЛОЖЕНИЕ ====================
class CheckoChecker:
    def __init__(self):
        self.api: Optional[CheckoAPI] = None
        self.cache = CacheManager()
        self.history = HistoryManager()
        self.export = ExportManager()
        self.key_manager = KeyManager()
        self.load_config()

    def load_config(self):
        config_path = Path(CONFIG_FILE)
        if not config_path.exists():
            logger.info("Конфиг-файл не найден. Создаю новый...")
            api_key = input("Введите ваш API-ключ Checko.ru: ").strip()
            config = {
                'api_key_encrypted': self.key_manager.encrypt(api_key),
                'default_mode': 'single'
            }
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            logger.info(f"Конфиг сохранён в {CONFIG_FILE}")

        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)

        encrypted_key = config.get('api_key_encrypted')
        if not encrypted_key:
            logger.error("API-ключ не настроен. Отредактируйте config.json")
            sys.exit(1)
        try:
            api_key = self.key_manager.decrypt(encrypted_key)
        except Exception as e:
            logger.error(f"Ошибка расшифровки API-ключа: {e}")
            logger.error("Удалите secret.key и config.json для пересоздания.")
            sys.exit(1)

        self.api = CheckoAPI(api_key, self.cache)

    def print_company_info(self, company: CompanyData):
        print("\n" + "=" * 80)
        print(f"ИНН: {company.inn}")
        print(f"ОГРН: {company.ogrn}")
        print(f"Название: {company.name}")
        print(f"КПП: {company.kpp}")
        print(f"Статус: {company.status}")
        print(f"Дата регистрации: {company.registration_date}")
        print(f"Директор: {company.director}")
        print(f"Адрес: {company.address}")

        print("\nВиды деятельности (ОКВЭД):")
        if company.okved_list:
            for okved in company.okved_list:
                print(f"  - {okved}")
        else:
            print("  (нет данных)")

        if company.contacts:
            print("\nКонтакты:")
            for key, value in company.contacts.items():
                if value:
                    print(f"  {key}: {value}")

        if company.finances:
            print("\nФинансовая отчётность (тыс. руб., ненулевые показатели):"
                  if not SHOW_ZERO_FINANCES else "\nФинансовая отчётность:")
            for year in sorted(company.finances, key=str):
                block = company.finances[year]
                if not isinstance(block, dict):
                    continue
                items = block if SHOW_ZERO_FINANCES else {k: v for k, v in block.items() if v not in (0, '0', None)}
                if not items:
                    continue
                print(f"  {year}:")
                for k, v in items.items():
                    print(f"    {k}: {v}")
        print("=" * 80 + "\n")

    def single_check(self):
        while True:
            inn = input("\nВведите ИНН (10 цифр) или 'q' для выхода: ").strip()
            if inn.lower() == 'q':
                break
            if not self.api.validate_inn(inn):
                logger.error("Некорректный ИНН. Должно быть 10 цифр.")
                continue
            company = self.api.get_company(inn)
            if company:
                self.print_company_info(company)
                self.history.add_record(asdict(company))
                logger.info("Результат сохранён в историю")
            else:
                print(f"Компания с ИНН {inn} не найдена или произошла ошибка.")

    def search_mode(self):
        while True:
            query = input("\nВведите название компании или 'q' для выхода: ").strip()
            if query.lower() == 'q':
                break
            results = self.api.search_company(query)
            if not isinstance(results, list):
                logger.error(f"Поиск вернул неожиданный тип: {type(results)}")
                results = []
            if not results:
                print("Ничего не найдено.")
                continue
            print(f"\nНайдено компаний: {len(results)}")
            for i, result in enumerate(results[:10], 1):
                if isinstance(result, dict):
                    print(f"{i}. {result.get('name', 'N/A')} (ИНН: {result.get('inn', 'N/A')})")
                else:
                    print(f"{i}. {result}")
            choice = input("\nВведите номер для подробной информации или 'q' для выхода: ").strip()
            if choice.lower() == 'q':
                break
            try:
                idx = int(choice) - 1
                if 0 <= idx < len(results) and isinstance(results[idx], dict):
                    inn = results[idx].get('inn')
                    if inn:
                        company = self.api.get_company(str(inn))
                        if company:
                            self.print_company_info(company)
                            self.history.add_record(asdict(company))
            except ValueError:
                print("Некорректный ввод.")

    def batch_check(self):
        file_path = input("Введите путь к файлу со списком ИНН (по одному в строке): ").strip()
        if not Path(file_path).exists():
            logger.error(f"Файл {file_path} не найден")
            return
        with open(file_path, 'r', encoding='utf-8') as f:
            inns = [line.strip() for line in f if line.strip()]
        logger.info(f"Загружено ИНН для проверки: {len(inns)}")

        output_file = "results.txt"
        with open(output_file, 'w', encoding='utf-8') as out:
            out.write("Результаты проверки компаний\n" + "=" * 80 + "\n\n")
            for i, inn in enumerate(inns, 1):
                logger.info(f"Обработка {i}/{len(inns)}: ИНН {inn}")
                if not self.api.validate_inn(inn):
                    out.write(f"ИНН {inn}: Некорректный формат\n\n")
                    continue
                company = self.api.get_company(inn)
                if company:
                    out.write(
                        f"ИНН: {company.inn}\nНазвание: {company.name}\nОГРН: {company.ogrn}\n"
                        f"Статус: {company.status}\nДиректор: {company.director}\n"
                        f"Адрес: {company.address}\n{'-' * 80}\n\n"
                    )
                    self.history.add_record(asdict(company))
                    time.sleep(0.5)
                else:
                    out.write(f"ИНН {inn}: Не найдена\n\n")
        logger.info(f"Результаты сохранены в {output_file}")
        print(f"\nРезультаты сохранены в {output_file}")

    def show_history(self):
        history = self.history.get_all()
        if not history:
            print("История проверок пуста.")
            return
        print(f"\nИстория проверок ({len(history)} записей):")
        print("-" * 80)
        for i, record in enumerate(history[:20], 1):
            print(f"{i}. {record['checked_at']} | ИНН: {record['inn']} | {record['company_name']}")
        if len(history) > 20:
            print(f"... и ещё {len(history) - 20} записей")

    def export_data(self):
        history = self.history.get_all()
        if not history:
            print("Нет данных для экспорта.")
            return
        print("\nВыберите формат экспорта:")
        print("1. CSV")
        print("2. Excel")
        choice = input("Ваш выбор (1-2): ").strip()
        if choice == '1':
            self.export.to_csv(history)
        elif choice == '2':
            self.export.to_excel(history)
        else:
            print("Некорректный выбор.")

    def run(self):
        print("\n" + "=" * 80)
        print("CHECKO.RU API CHECKER v2.1")
        print("=" * 80)
        self.cache.clear_old(days=7)
        while True:
            print("\nВыберите режим работы:")
            print("1. Проверка по ИНН (одиночная)")
            print("2. Поиск по названию")
            print("3. Массовая проверка из файла")
            print("4. История проверок")
            print("5. Экспорт данных (CSV/Excel)")
            print("6. Выход")
            choice = input("\nВаш выбор (1-6): ").strip()
            try:
                if choice == '1':
                    self.single_check()
                elif choice == '2':
                    self.search_mode()
                elif choice == '3':
                    self.batch_check()
                elif choice == '4':
                    self.show_history()
                elif choice == '5':
                    self.export_data()
                elif choice == '6':
                    print("До свидания!")
                    break
                else:
                    print("Некорректный выбор. Введите 1-6.")
            except KeyboardInterrupt:
                print("\n\nПрервано пользователем. До свидания!")
                break
            except Exception as e:
                logger.error(f"Неожиданная ошибка: {e}")
        self.cache.close()
        self.history.close()


def main():
    try:
        app = CheckoChecker()
        app.run()
    except Exception as e:
        logger.critical(f"Критическая ошибка: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
