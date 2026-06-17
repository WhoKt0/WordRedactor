# RenSer — бот коммерческих писем

Production-ready бот для автоматической генерации персонализированных коммерческих писем по Word-шаблону, конвертации в PDF и отправки по email.

## 1. Что делает бот

1. Читает список банков из Excel (`data/banks.xlsx`).
2. Для каждой строки подставляет данные в Word-шаблон (безопасная замена плейсхолдеров с сохранением форматирования).
3. Сохраняет персональный DOCX в `output/docx/`.
4. Конвертирует DOCX в PDF через Microsoft Word COM (Windows).
5. Отправляет email с PDF-вложением (или только генерирует файлы в режиме `dry_run`).
6. Пишет подробный лог и отчёт `output/reports/status_YYYYMMDD_HHMMSS.xlsx`.

## 2. Структура проекта

```
project/
  templates/
    letter_template.docx      # Word-шаблон с плейсхолдерами
    email_template.txt        # Текст письма
  data/
    banks.xlsx                # Входные данные
  output/
    docx/                     # Сгенерированные Word-файлы
    pdf/                      # PDF-файлы
    reports/                  # Отчёты о статусе отправки
  logs/
    app.log                   # Лог приложения
  src/
    main.py                   # Точка входа
    config.py
    models.py
    excel_reader.py
    document_generator.py
    pdf_converter.py
    email_sender.py
    status_report.py
    validators.py
    logger_setup.py
  scripts/
    create_sample_assets.py   # Создание примера шаблона и Excel
  .env.example
  config.yaml
  requirements.txt
  README.md
```

## 3. Как подготовить Word-шаблон

1. Создайте документ в Microsoft Word с нужным оформлением (шрифты, отступы, логотип).
2. В местах подстановки вставьте **цельные** плейсхолдеры одним фрагментом текста (не разбивайте `{{OUT_NUMBER}}` на несколько runs вручную — бот умеет объединять runs, но надёжнее держать плейсхолдер цельным).
3. Сохраните файл как `templates/letter_template.docx`.

Пример готового шаблона можно сгенерировать:

```bash
python scripts/create_sample_assets.py
```

## 4. Плейсхолдеры в Word-шаблоне

| Плейсхолдер | Пример результата |
|---|---|
| `{{OUT_NUMBER}}` | `315` |
| `{{DATE}}` | `02.06.2026` |
| `{{BANK_LEGAL_NAME}}` | `АК "Алокабанк"` |
| `{{MR_MS}}` | `г-же` / `г-ну` |
| `{{CHAIR_SHORT_DATIVE}}` | `Ирисбековой К. Н.` |
| `{{GREETING_WORD}}` | `Уважаемая` / `Уважаемый` |
| `{{GREETING_NAME}}` | `Каммуна Наринбаевна` |

Пример блока в документе:

```
Исх № {{OUT_NUMBER}} от {{DATE}}.

Председателю правления
{{BANK_LEGAL_NAME}}
{{MR_MS}} {{CHAIR_SHORT_DATIVE}}

{{GREETING_WORD}} {{GREETING_NAME}},
```

Правила по полу (`gender` в Excel):

- `female` → `г-же`, `Уважаемая`
- `male` → `г-ну`, `Уважаемый`

## 5. Как подготовить Excel

Файл: `data/banks.xlsx`, лист с данными (первая строка — заголовки).

| Колонка | Описание | Пример |
|---|---|---|
| `bank_name` | Краткое имя банка | Алокабанк |
| `bank_legal_name` | Юр. название в письме | АК "Алокабанк" |
| `recipient_email` | Email получателя | info@example.uz |
| `cc_email` | Копия (необязательно) | |
| `bcc_email` | Скрытая копия (необязательно) | |
| `chair_full_name` | Полное ФИО председателя | Ирисбекова Каммуна Наринбаевна |
| `chair_short_dative` | ФИО в дательном падеже, кратко | Ирисбековой К. Н. |
| `gender` | `male` или `female` | female |
| `greeting_name` | Имя и отчество для обращения | Каммуна Наринбаевна |
| `pdf_filename` | Имя PDF-файла | Предложение_по_картам_Алокабанк.pdf |
| `email_bank_name` | Название банка в теме/теле email | АК "Алокабанк" |

## 6. Настройка `.env`

Скопируйте пример:

```bash
copy .env.example .env
```

Заполните:

```env
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=your@gmail.com
SMTP_PASSWORD=your_app_password
FROM_EMAIL=your@gmail.com
FROM_NAME=RenSer Technologies
START_OUT_NUMBER=315
```

- `START_OUT_NUMBER` — исходящий номер для первой **успешно** обработанной строки; для каждой следующей успешной строки +1.
- Дата письма — сегодняшняя, либо задаётся в `config.yaml` → `app.explicit_date`.

## 7. Настройка `config.yaml`

Ключевые параметры:

```yaml
app:
  dry_run: true              # true = не отправлять email
  stop_on_error: false       # true = остановиться при ошибке строки
  delay_between_emails_seconds: 15
  explicit_date: null        # например "02.06.2026" или "2026-06-02"
  date_format: "%d.%m.%Y"
```

## 8. Установка и запуск

### Требования

- Windows 10/11
- Python 3.10+
- **Microsoft Word** (для конвертации PDF)
- Доступ к SMTP-серверу

### Установка

```bash
cd project
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python scripts/create_sample_assets.py
copy .env.example .env
```

### Первый тест без отправки email

В `config.yaml` оставьте:

```yaml
app:
  dry_run: true
```

Запуск:

```bash
python -m src.main
```

Бот создаст DOCX и PDF, в логе появится строка `DRY_RUN email would be sent to ...`, статус в отчёте — `generated_not_sent`.

### Проверка PDF перед отправкой

1. Откройте файлы в `output/pdf/`.
2. Проверьте номер, дату, адресата, обращение и вёрстку.
3. Убедитесь, что в документе нет оставшихся `{{...}}`.

### Реальная отправка email

1. Заполните `.env` (SMTP, FROM_EMAIL, пароль).
2. В `config.yaml` установите:

```yaml
app:
  dry_run: false
```

3. Замените тестовые email в Excel на реальные.
4. Запустите:

```bash
python -m src.main
```

## 9. Логи

- Консоль — основные события.
- Файл `logs/app.log` — полный лог с ошибками и traceback.

Уровень логирования: `config.yaml` → `logging.level` (`DEBUG`, `INFO`, `WARNING`, `ERROR`).

## 10. Status report

После каждого запуска: `output/reports/status_YYYYMMDD_HHMMSS.xlsx`

| Колонка | Описание |
|---|---|
| `row_number` | Номер строки Excel |
| `bank_name` | Банк |
| `out_number` | Исходящий номер |
| `date` | Дата в письме |
| `docx_path` | Путь к DOCX |
| `pdf_path` | Путь к PDF |
| `status` | Итоговый статус |
| `error_message` | Текст ошибки |
| `sent_at` | Время отправки |

Статусы:

| Статус | Значение |
|---|---|
| `validation_failed` | Ошибка валидации, письмо не создавалось |
| `docx_generated` | Промежуточный (обычно не финальный) |
| `pdf_generated` | Промежуточный |
| `generated_not_sent` | DOCX+PDF готовы, email не отправлен (dry_run) |
| `sent` | Email отправлен |
| `failed` | Ошибка на этапе генерации/конвертации/отправки |

## 11. Частые ошибки

| Проблема | Решение |
|---|---|
| Неправильный email | Проверьте `recipient_email` в Excel |
| Microsoft Word не установлен | Установите Word; без него PDF не создаётся |
| Шаблон не найден | Положите `templates/letter_template.docx` |
| Остались `{{...}}` в Word | Проверьте плейсхолдеры в шаблоне и данные в Excel |
| SMTP auth failed | Проверьте логин/пароль SMTP |
| Gmail блокирует вход | Используйте [пароль приложения](https://support.google.com/accounts/answer/185833), не обычный пароль |
| `stop_on_error: true` | Бот остановится на первой ошибке после валидации |

## 12. Плейсхолдеры в email

Файл `templates/email_template.txt`:

- `{{BANK_NAME}}`
- `{{BANK_LEGAL_NAME}}`

Тема письма настраивается в `config.yaml` → `email.subject_template` (поддерживает `{{BANK_NAME}}`).

## Быстрый старт (чеклист)

1. `pip install -r requirements.txt`
2. `python scripts/create_sample_assets.py`
3. Положите свой `templates/letter_template.docx` (или используйте сгенерированный пример).
4. Заполните `data/banks.xlsx`.
5. `copy .env.example .env` и при необходимости настройте SMTP.
6. `dry_run: true` → `python -m src.main` → проверьте `output/pdf/`.
7. `dry_run: false` → реальная отправка.
