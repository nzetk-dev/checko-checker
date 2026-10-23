import json
import logging
import sys
import time
import sqlite3
import csv
import hashlib
from typing import Optional, Dict, List, Any, Tuple
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

# Расшифровка кодов строк бухгалтерской отчётности (Ф1 — Баланс, Ф2 — ОФР)
FIN_CODE_NAMES = {
    "1100": "Итого внеоборотных активов",
    "1110": "Нематериальные поисковые активы",
    "1120": "Результаты исследований и разработок",
    "1130": "Нематериальные активы",
    "1140": "Основные средства",
    "1150": "Доходные вложения в материальные ценности",
    "1160": "Отложенные налоговые активы",
    "1170": "Прочие внеоборотные активы",
    "1180": "Финансовые вложения (неликвидные)",
    "1190": "Прочие внеоборотные активы",
    "1200": "Итого оборотных активов",
    "1210": "Запасы",
    "1220": "НДС по приобретённым ценностям",
    "1230": "Дебиторская задолженность",
    "1240": "Финансовые вложения (ликвидные)",
    "1250": "Денежные средства и денежные эквиваленты",
    "1260": "Прочие оборотные активы",
    "1300": "Итого капитал",
    "1310": "Уставный капитал (складочный капитал, уставный фонд)",
    "1320": "Собственные акции, выкупленные у акционеров",
    "1340": "Переоценка внеоборотных активов",
    "1350": "Добавочный капитал (без переоценки)",
    "1360": "Резервный капитал",
    "1370": "Нераспределённая прибыль (непокрытый убыток)",
    "1400": "Итого долгосрочных обязательств",
    "1410": "Заемные средства (долгосрочные)",
    "1420": "Отложенные налоговые обязательства",
    "1430": "Оценочные обязательства (долгосрочные)",
    "1450": "Прочие долгосрочные обязательства",
    "1500": "Итого краткосрочных обязательств",
    "1510": "Заемные средства (краткосрочные)",
    "1520": "Кредиторская задолженность",
    "1530": "Доходы будущих периодов",
    "1540": "Оценочные обязательства (краткосрочные)",
    "1550": "Прочие краткосрочные обязательства",
    "1600": "Баланс (актив)",
    "1700": "Баланс (пассив)",
    "2100": "Валовая прибыль (убыток)",
    "2110": "Выручка",
    "2120": "Себестоимость продаж",
    "2200": "Прибыль (убыток) от продаж",
    "2210": "Коммерческие расходы",
    "2220": "Управленческие расходы",
    "2300": "Прибыль (убыток) до налогообложения",
    "2310": "Доходы от участия в других организациях",
    "2320": "Проценты к получению",
    "2330": "Проценты к уплате",
    "2340": "Прочие доходы",
    "2350": "Прочие расходы",
    "2400": "Чистая прибыль (убыток)",
    "2410": "Совокупный финансовый результат",
    "2421": "Результат от переоценки внеоборотных активов",
    "2430": "Результат от прочих операций",
    "2450": "Изменение добавочного капитала",
    "2460": "Изменение резервного капитала",
    "2500": "Совокупный финансовый результат периода",
    "2510": "Результат от переоценки внеоборотных активов, не включаемый в чистую прибыль",
    "2520": "Результат от прочих операций, не включаемый в чистую прибыль",
    "3200": "Величина капитала на 31 декабря года, предшествовавшего отчётному",
    "3300": "Величина капитала на 31 декабря отчётного года",
    "3310": "Поступления — всего",
    "3311": "В результате выпуска акций",
    "3312": "В результате увеличения номинальной стоимости акций",
    "3313": "В результате реорганизации",
    "3314": "Изменения добавочного капитала",
    "3315": "Изменения резервного капитала",
    "3316": "Изменения нераспределённой прибыли (непокрытого убытка)",
    "3320": "Изъятия — всего",
    "3321": "В результате уменьшения номинальной стоимости акций",
    "3322": "В результате уменьшения числа акций",
    "3323": "В результате реорганизации",
    "3324": "Изменения добавочного капитала",
    "3325": "Изменения резервного капитала",
    "3326": "Изменения нераспределённой прибыли (непокрытого убытка)",
    "3600": "Чистые активы",
    "4100": "Чистые финансовые вложения (за исключением денежных эквивалентов)",
    "4110": "Поступления — всего",
    "4111": "От продажи акций других организаций",
    "4112": "От возврата займов, процентов по займам",
    "4113": "Иные платежи",
    "4119": "Иные платежи",
    "4120": "Платежи — всего",
    "4121": "В связи с приобретением акций, долей",
    "4122": "В связи с предоставлением займов",
    "4123": "Иные платежи",
    "4124": "Иные платежи",
    "4129": "Иные платежи",
    "4200": "Денежные потоки от текущих операций",
    "4210": "Поступления — всего",
    "4211": "От продажи продукции, товаров, работ, услуг",
    "4212": "Арендные платежи, лицензионные платежи, роялти",
    "4213": "От перепродажи финансовых активов",
    "4214": "От процентов на остатки денежных средств",
    "4219": "Прочие поступления",
    "4220": "Платежи — всего",
    "4221": "Поставщикам (подрядчикам)",
    "4222": "Работникам",
    "4223": "Проценты по долговым обязательствам",
    "4224": "Налоги на прибыль",
    "4229": "Прочие платежи",
    "4300": "Денежные потоки от финансовых операций",
    "4310": "Поступления — всего",
    "4311": "От выпуска акций, долей",
    "4312": "От выпуска облигаций, векселей и т.д.",
    "4313": "От предоставления займов третьими лицами",
    "4314": "Прочие поступления",
    "4320": "Платежи — всего",
    "4321": "В связи с погашением (выкупом) акций, долей",
    "4322": "В связи с погашением облигаций, векселей",
    "4323": "В связи с погашением займов и кредитов",
    "4324": "В связи с уплатой дивидендов, процентов",
    "4329": "Прочие платежи",
    "4400": "Величина изменения денежных средств и денежных эквивалентов",
    "4450": "Денежные средства и денежные эквиваленты на начало отчётного периода",
    "4490": "Денежные средства и денежные эквиваленты на конец отчётного периода",
    "4500": "Денежные эквиваленты на конец отчётного периода",
    "6100": "Разница между величиной чистых активов и уставным капиталом на конец года",
    "6200": "Разница между величиной чистых активов и уставным капиталом на начало года",
    "6210": "Принятая учётная политика (чистые активы меньше УК)",
    "6215": "Решение об уменьшении УК",
    "6220": "Увеличение УК",
    "6230": "Уменьшение УК",
    "6240": "Вклад в имущество общества",
    "6250": "Прочие изменения",
    "6300": "Покрытие за счёт имущества",
    "6310": "Изменение добавочного капитала",
    "6311": "Результат переоценки",
    "6312": "Прочее",
    "6313": "Прочее",
    "6320": "Изменение резервного капитала",
    "6321": "Отчисления",
    "6322": "Использование",
    "6323": "Прочее",
    "6324": "Прочее",
    "6325": "Прочее",
    "6326": "Прочее",
    "6330": "Изменение нераспределённой прибыли",
    "6350": "Изменение собственных акций",
    "6400": "Прочие изменения",
}


# ==================== ЛОГИРОВАНИЕ ====================
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
        print(f"✓ Данные экспортированы в {filename}")

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
            ws.column_dimensions[column[0].column_letter].width = min(max_length + 2, 60)
        wb.save(filename)
        logger.info(f"Данные экспортированы в {filename}")
        print(f"✓ Данные экспортированы в {filename}")


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
    BASE_URL = "https://api.checko.ru/v2"

    def __init__(self, api_key: str, cache: CacheManager):
        self.api_key = api_key
        self.cache = cache
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'CheckoChecker/2.2',
            'Accept': 'application/json'
        })

    @staticmethod
    def _dump_raw(response: Any, filename: str):
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(response, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.debug(f"Не удалось сохранить сырой ответ: {e}")

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

    def _parse_company(self, data: Dict[str, Any]) -> CompanyData:
        """Парсинг ответа с русскоязычными ключами (схема ofdata/Checko v2)"""
        # Защитные геттеры
        def g(d, *keys, default=''):
            if not isinstance(d, dict):
                return default
            for k in keys:
                v = d.get(k)
                if v not in (None, '', [], {}):
                    return v
            return default

        inn = str(g(data, 'ИНН', 'inn'))
        ogrn = str(g(data, 'ОГРН', 'ogrn'))
        kpp = str(g(data, 'КПП', 'kpp'))
        name = str(g(data, 'НаимПолн', 'НаимСокр', 'НаимАнгл', 'name', 'name_full'))

        status_raw = g(data, 'Статус', 'status')
        if isinstance(status_raw, dict):
            status = str(status_raw.get('Наим') or status_raw.get('name') or '')
        else:
            status = str(status_raw)

        reg_date = str(g(data, 'ДатаРег', 'ДатаОГРН', 'reg_date', 'registration_date'))

        # Юридический адрес
        addr_raw = g(data, 'ЮрАдрес', 'address')
        if isinstance(addr_raw, dict):
            address = str(
                addr_raw.get('АдресРФ')
                or addr_raw.get('Адрес')
                or addr_raw.get('value')
                or addr_raw.get('text')
                or ''
            )
        else:
            address = str(addr_raw)

        # Руководитель (может быть список)
        mgr_raw = g(data, 'Руковод', 'manager', 'director', 'head')
        if isinstance(mgr_raw, list) and mgr_raw:
            first = mgr_raw[0]
            if isinstance(first, dict):
                director = str(first.get('ФИО') or first.get('name') or first.get('fio') or '')
            else:
                director = str(first)
        elif isinstance(mgr_raw, dict):
            director = str(mgr_raw.get('ФИО') or mgr_raw.get('name') or mgr_raw.get('fio') or '')
        else:
            director = str(mgr_raw)

        # ОКВЭД: основной + дополнительные
        okved_list: List[str] = []
        main_okved = g(data, 'ОКВЭД', 'okved')
        if isinstance(main_okved, dict):
            code = main_okved.get('Код') or main_okved.get('code') or ''
            nm = main_okved.get('Наим') or main_okved.get('name') or main_okved.get('title') or ''
            if code or nm:
                okved_list.append(f"{code} {nm}".strip())
        extra = g(data, 'ОКВЭДДоп', 'okved_extra')
        if isinstance(extra, list):
            for item in extra:
                if isinstance(item, dict):
                    code = item.get('Код') or item.get('code') or ''
                    nm = item.get('Наим') or item.get('name') or item.get('title') or ''
                    if code or nm:
                        okved_list.append(f"{code} {nm}".strip())
                elif isinstance(item, str):
                    okved_list.append(item)

        # Контакты
        contacts_raw = g(data, 'Контакты', 'contacts')
        contacts: Dict[str, Any] = {}
        if isinstance(contacts_raw, dict):
            for key, label in [('Тел', 'Телефоны'), ('Емэйл', 'Email'), ('ВебСайт', 'Веб-сайт')]:
                val = contacts_raw.get(key)
                if val:
                    if isinstance(val, list):
                        contacts[label] = ', '.join(str(x) for x in val if x)
                    else:
                        contacts[label] = str(val)

        return CompanyData(
            inn=inn, ogrn=ogrn, name=name, kpp=kpp, address=address,
            director=director, okved_list=okved_list, registration_date=reg_date,
            status=status, contacts=contacts
        )

    def get_company(self, inn: str) -> Optional[CompanyData]:
        logger.info(f"Запрос данных компании с ИНН: {inn}")
        try:
            response = self._make_request('company', {'inn': inn})
            self._dump_raw(response, LAST_COMPANY_RESPONSE)

            if not isinstance(response, dict) or 'data' not in response:
                logger.warning(f"Компания с ИНН {inn} не найдена")
                return None

            data = response['data']
            if not isinstance(data, dict):
                logger.warning(f"Ожидался словарь в data для ИНН {inn}")
                return None

            company = self._parse_company(data)

            # Получение финансовой отчётности
            try:
                fin_response = self._make_request('finances', {'inn': inn})
                if isinstance(fin_response, dict) and 'data' in fin_response:
                    company.finances = fin_response['data']
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
            if data.get('inn') or data.get('ИНН'):
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

    def format_finance(self, block: Dict[str, Any], show_zero: bool = False) -> List[Tuple[str, str, Any]]:
        """Возвращает список кортежей (код, расшифровка, значение) без нулей"""
        out = []
        for code, value in block.items():
            code_str = str(code)
            if not show_zero and value in (0, '0', None, ''):
                continue
            label = FIN_CODE_NAMES.get(code_str, "")
            if label:
                out.append((code_str, label, value))
            else:
                out.append((code_str, f"Строка {code_str}", value))
        return out

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
            print("\nФинансовая отчётность:")
            for year in sorted(company.finances, key=str):
                block = company.finances[year]
                if not isinstance(block, dict):
                    continue
                items = self.format_finance(block, show_zero=False)
                if not items:
                    continue
                print(f"\n  Год {year}:")
                # Сортируем по коду, чтобы было логично
                for code, label, value in sorted(items, key=lambda x: x[0]):
                    try:
                        value_fmt = f"{int(value):,}".replace(',', ' ')
                    except (ValueError, TypeError):
                        value_fmt = str(value)
                    print(f"    {code} {label}: {value_fmt} тыс. руб.")
        print("=" * 80 + "\n")

    def single_check(self, inn: Optional[str] = None):
        """Одиночная проверка. Если inn=None — запрашивается у пользователя."""
        while True:
            if inn is None:
                raw = input("\nВведите ИНН (10 цифр) или 'q' для выхода: ").strip()
            else:
                raw = inn
                inn = None  # для следующих итераций запрашиваем у пользователя

            if raw.lower() == 'q':
                break
            if not self.api.validate_inn(raw):
                logger.error("Некорректный ИНН. Должно быть 10 цифр.")
                continue

            company = self.api.get_company(raw)
            if company:
                # Сохраняем в историю только если есть хоть какие-то данные
                if company.inn or company.name:
                    self.print_company_info(company)
                    self.history.add_record(asdict(company))
                    logger.info("Результат сохранён в историю")
                else:
                    logger.warning(
                        f"По ИНН {raw} получены пустые данные. "
                        f"Проверьте файл {LAST_COMPANY_RESPONSE}."
                    )
                    print(f"Компания с ИНН {raw} не найдена или API вернул пустой ответ.")
            else:
                print(f"Компания с ИНН {raw} не найдена или произошла ошибка.")

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
                    # Поддержка и русских, и латинских ключей
                    name = result.get('НаимПолн') or result.get('НаимСокр') or result.get('name') or 'N/A'
                    inn_val = result.get('ИНН') or result.get('inn') or 'N/A'
                    print(f"{i}. {name} (ИНН: {inn_val})")
                else:
                    print(f"{i}. {result}")
            choice = input("\nВведите номер для подробной информации или 'q' для выхода: ").strip()
            if choice.lower() == 'q':
                break
            try:
                idx = int(choice) - 1
                if 0 <= idx < len(results) and isinstance(results[idx], dict):
                    inn_val = results[idx].get('ИНН') or results[idx].get('inn')
                    if inn_val:
                        company = self.api.get_company(str(inn_val))
                        if company and (company.inn or company.name):
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
                if company and (company.inn or company.name):
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

    def show_raw_response(self):
        """Пункт меню: показать последний сырой ответ API"""
        for path, label in [
            (LAST_COMPANY_RESPONSE, "Компания"),
            (LAST_SEARCH_RESPONSE, "Поиск")
        ]:
            p = Path(path)
            if p.exists():
                print(f"\n--- Последний ответ ({label}) — {p} ---")
                try:
                    with open(p, 'r', encoding='utf-8') as f:
                        print(f.read()[:5000])
                except Exception as e:
                    print(f"Ошибка чтения: {e}")
            else:
                print(f"\nФайл {p} отсутствует (ещё не выполняли запросов типа «{label}»).")

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
        print("CHECKO.RU API CHECKER v2.2")
        print("=" * 80)
        print("Совет: ИНН (10 цифр) можно вводить прямо в главном меню.")
        self.cache.clear_old(days=7)
        while True:
            print("\nВыберите режим работы:")
            print("1. Проверка по ИНН (одиночная)")
            print("2. Поиск по названию")
            print("3. Массовая проверка из файла")
            print("4. История проверок")
            print("5. Экспорт данных (CSV/Excel)")
            print("6. Показать последний сырой ответ API (JSON)")
            print("7. Выход")
            raw = input("\nВаш выбор (1-7) или ИНН: ").strip()

            # Если пользователь ввёл 10-значное число — сразу проверка по ИНН
            if raw.isdigit() and len(raw) == 10:
                self.single_check(inn=raw)
                continue

            try:
                if raw == '1':
                    self.single_check()
                elif raw == '2':
                    self.search_mode()
                elif raw == '3':
                    self.batch_check()
                elif raw == '4':
                    self.show_history()
                elif raw == '5':
                    self.export_data()
                elif raw == '6':
                    self.show_raw_response()
                elif raw == '7':
                    print("До свидания!")
                    break
                else:
                    print("Некорректный выбор. Введите 1-7 или ИНН (10 цифр).")
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
