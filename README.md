# RenSer — бот коммерческих писем

Production-ready бот для автоматической генерации персонализированных коммерческих писем по Word-шаблону, конвертации в PDF и отдельной рассылки по email.

## 1. Что делает бот

1. Читает список банков из Excel (`data/banks.xlsx`).
2. Для каждой строки подставляет данные в Word-шаблон (безопасная замена плейсхолдеров с сохранением форматирования).
3. Сохраняет персональный DOCX и PDF.
4. Конвертирует DOCX в PDF через Microsoft Word COM (Windows).
5. Создаёт manifest успешных PDF для последующей рассылки.
6. Пишет подробный лог, Excel status report, manifest и TXT-файл ошибок (если ошибки есть).

**DOCX — промежуточный исходник для PDF.** Официальным документом считается PDF. СВХИ фиксируется только после успешного создания PDF.

**Email — отдельный этап** после финальной генерации. Рассылка не выполняется из `preview.py` и не запускается автоматически после `final_generate.py`.

## 2. Режимы генерации и исходящие номера (СВХИ)

Исходящий номер **больше не увеличивается при каждом запуске из `.env`**. Фактический последний зафиксированный номер хранится в `state/generation_state.json`.

`START_OUT_NUMBER` в `.env` используется **только при первом создании** `state/generation_state.json`. Если `START_OUT_NUMBER=315`, state создаётся с `last_committed_out_number=314`, и следующий номер будет `315`.

### Preview — безопасная тестовая генерация (по умолчанию)

```bash
python preview.py
```

или:

```bash
python run.py
python run.py --preview
```

- Обычная безопасная генерация при правках шаблона, Excel и текста.
- Перед запуском **очищает** старые файлы в `output/preview/docx/`, `pdf/`, `reports/` (`.gitkeep` сохраняется).
- Сначала выполняется **preflight validation** всех строк Excel.
- Документы создаются с номерами, начиная с `last_committed_out_number + 1`.
- **Номера не тратятся** — после завершения `last_committed_out_number` не меняется.
- Повторный preview снова начнёт с того же следующего номера.
- Файлы складываются в `output/preview/docx/`, `output/preview/pdf/`, `output/preview/reports/`.
- Создаётся manifest успешных PDF: `ген_N_manifest.xlsx`.

### Final / Commit — финальная генерация

```bash
python final_generate.py
```

или:

```bash
python run.py --commit-numbers
python run.py --commit-numbers --yes
```

- Финальная генерация, когда документы уже готовы к отправке.
- Сначала выполняется **preflight validation** и показывается summary с диапазоном СВХИ.
- Только после ввода `YES` начинается генерация (до этого номера не фиксируются).
- **СВХИ фиксируется только после успешного PDF**, не после одного DOCX.
- Если DOCX создан, но PDF не создан — номер не тратится.
- Если PDF создан, но email не отправился — **номер уже считается использованным** (СВХИ не откатывается). Повторить отправку можно через `send_emails.py`.
- Ошибочные строки preflight пропускаются без траты номера.
- Файлы складываются в `output/docx/`, `output/pdf/`, `output/reports/`.
- Создаётся manifest успешных PDF: `ген_N_manifest.xlsx`.

### Проверка и ручная корректировка счётчика

```bash
python confirm_counter.py
```

- Показывает последний зафиксированный исходящий номер.
- Показывает следующий финальный номер.
- Позволяет вручную изменить `last_committed_out_number`.

## 3. Обращение после «Уважаемый» / «Уважаемая»

В шаблоне используется:

```text
{{GREETING_WORD}} {{GREETING_NAME}},
```

После «Уважаемый/Уважаемая» должно быть **минимум два слова**:

- Плохо: `Уважаемый Шухрат,`
- Хорошо: `Уважаемый Шухрат Атабаев,` или `Уважаемая Каммуна Наринбаевна,`

Если в Excel `greeting_name` содержит только одно слово, бот пытается восстановить второе слово из `chair_full_name`. Служебные слова (`перепроверь`, `проверь`, `уточнить`, `уточни`) игнорируются.

Если восстановить корректное обращение невозможно — строка пропускается, документ не создаётся, номер не тратится, ошибка попадает в `output/.../reports/ген_N_ошибки.txt`.

## 4. Структура проекта

```
project/
  preview.py                  # Preview-генерация (по умолчанию)
  final_generate.py           # Финальная генерация с фиксацией номеров
  confirm_counter.py          # Просмотр/изменение счётчика номеров
  email_preview.py            # Просмотр писем по manifest (без отправки)
  email_test_send.py          # Одно тестовое письмо на ваш email
  send_emails.py              # Реальная рассылка по final manifest
  run.py                      # CLI-запуск с флагами
  state/
    generation_state.json     # Последний зафиксированный номер (не в git)
  templates/
    Template_word.docx                    # Обычный шаблон (по умолчанию)
    Template_word_greeting_long.docx      # Длинное обращение (> 20 символов)
    Template_word_greeting_short.docx     # Короткое обращение (< 10 символов)
    email_template.txt                    # Текст письма
  data/
    banks.xlsx                # Входные данные
  output/
    docx/                     # Финальные Word-файлы
    pdf/                      # Финальные PDF
    reports/                  # Финальные отчёты
    preview/
      docx/                   # Preview Word-файлы
      pdf/                    # Preview PDF
      reports/                # Preview отчёты и TXT ошибок
  logs/
    app.log                   # Лог приложения
  src/
    main.py
    config.py
    state_manager.py
    name_utils.py
    error_report.py
    manifest_report.py
    email_pipeline.py
    email_report.py
    preflight.py
    output_utils.py
    template_selector.py
    models.py
    excel_reader.py
    document_generator.py
    pdf_converter.py
    email_sender.py
    status_report.py
    validators.py
    logger_setup.py
  .env.example
  config.yaml
  requirements.txt
  README.md
```

## 5. Плейсхолдеры в Word-шаблоне

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

```text
Исх № {{OUT_NUMBER}} от {{DATE}}.

Председателю правления
{{BANK_LEGAL_NAME}}
{{MR_MS}} {{CHAIR_SHORT_DATIVE}}

{{GREETING_WORD}} {{GREETING_NAME}},
```

### Три шаблона для выравнивания обращения

Для визуального выравнивания обращения используются 3 шаблона. Это практичный костыль под текущий Word/PDF-макет: разные буквы имеют разную ширину, поэтому выравнивание не математически точное.

Бот считает длину **финального** `greeting_name_final` (после `build_greeting_name`, с пробелами внутри ФИО, без «Уважаемый/Уважаемая» и запятой):

| Длина `greeting_name_final` | Шаблон |
|---|---|
| больше 20 символов | `Template_word_greeting_long.docx` (на 3 пробела меньше перед обращением) |
| меньше 10 символов | `Template_word_greeting_short.docx` (на 3 пробела больше перед обращением) |
| иначе | `Template_word.docx` |

Примеры длины: `Каммуна Наринбаевна` (18) → обычный; `Шухрат Атабаев` (15) → обычный; очень длинные имена → long; короткие вроде `Волкан Гюл` (10) → обычный; `Волкан` одно слово не пройдёт preflight.

Создать long/short копии из основного шаблона:

```bash
python scripts/create_greeting_templates.py
```

В manifest колонка `template_used` показывает, какой шаблон был применён для строки.

## 6. Как подготовить Excel

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
| `greeting_name` | Имя для обращения (минимум 2 слова или восстановимо) | Каммуна Наринбаевна |
| `pdf_filename` | Имя PDF-файла | Предложение_по_картам_Алокабанк.pdf |
| `email_bank_name` | Название банка в теме/теле email | АК "Алокабанк" |

## 7. Настройка `.env`

```bash
copy .env.example .env
```

```env
SMTP_HOST=
SMTP_PORT=587
SMTP_USER=
SMTP_PASSWORD=
FROM_EMAIL=
FROM_NAME=RenSer Technologies

# Used only on first state/generation_state.json creation.
# Preview generations do not spend this number.
START_OUT_NUMBER=315
```

## 8. Настройка `config.yaml`

```yaml
app:
  stop_on_error: false
  delay_between_emails_seconds: 15   # задержка в send_emails.py между письмами
  explicit_date: null
  date_format: "%d.%m.%Y"
```

Параметр `dry_run` в `config.yaml` больше не влияет на генерацию PDF. Рассылка управляется отдельными скриптами (`email_preview.py`, `send_emails.py`).

## 9. Установка и запуск

### Требования

- Windows 10/11
- Python 3.10+
- **Microsoft Word** (для конвертации PDF)
- Доступ к SMTP-серверу (для реальной отправки)

### Установка

```bash
cd project
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
```

### Типичный рабочий процесс

1. Правите шаблон, Excel, текст.
2. Запускаете `python preview.py` — сколько угодно раз, номера не тратятся.
3. Проверяете `output/preview/pdf/`.
4. Когда всё готово — `python final_generate.py` (с подтверждением `YES`).
5. При необходимости — `python confirm_counter.py` для проверки счётчика.
6. `python email_preview.py` — смотрите, кому уйдут письма.
7. `python email_test_send.py` — одно тестовое письмо на ваш email.
8. `python send_emails.py` — реальная рассылка по manifest (подтверждение `SEND`).

## 10. Рассылка email

Рассылка — **отдельный этап** поверх уже готовых PDF. `send_emails.py` **не читает Excel напрямую**, а работает по manifest успешных PDF из `output/reports/`.

### Порядок запуска

```bash
python preview.py
python final_generate.py
python email_preview.py
python email_test_send.py
python send_emails.py
```

| Скрипт | Назначение |
|---|---|
| `preview.py` | Проверка PDF, СВХИ не тратятся |
| `final_generate.py` | Финальные PDF, СВХИ фиксируются, создаётся manifest |
| `email_preview.py` | Показывает, кому будут письма; создаёт `ген_N_email_report.xlsx` со статусом `dry_run` |
| `email_test_send.py` | Одно тестовое письмо на введённый email с префиксом `[TEST]` |
| `send_emails.py` | Реальная рассылка; требует ввода `SEND` |

### Выбор manifest

По умолчанию берётся последний final manifest из `output/reports/ген_*_manifest.xlsx`. Preview manifest из `output/preview/reports/` **не используется**.

```bash
python send_emails.py --manifest output/reports/ген_5_manifest.xlsx
python email_preview.py --manifest output/reports/ген_5_manifest.xlsx
python email_test_send.py --manifest output/reports/ген_5_manifest.xlsx
```

### Продолжение после сбоя

Если рассылка прервалась, повторный `send_emails.py` **не отправляет повторно** строки со статусом `sent` в `ген_N_email_report.xlsx`. Повторяются только `failed` и `pending`.

```bash
python send_emails.py --resend-all   # принудительно отправить все заново
```

Сопоставление строк: `generation_id + out_number + recipient_email + pdf_path`.

### Email report

После каждой рассылки: `output/reports/ген_N_email_report.xlsx`

Статусы: `pending`, `dry_run`, `sent`, `failed`, `skipped_pdf_missing`, `skipped_invalid_email`.

### Переменные в email-шаблоне

В `templates/email_template.txt` и `config.yaml → email.subject_template`:

| Переменная | Описание |
|---|---|
| `{{BANK_NAME}}` | Название банка для письма |
| `{{BANK_LEGAL_NAME}}` | Юр. название |
| `{{OUT_NUMBER}}` | Исходящий номер |
| `{{DATE}}` | Дата письма |
| `{{GREETING_WORD}}` | Уважаемый / Уважаемая |
| `{{GREETING_NAME}}` | Имя в обращении |
| `{{CHAIR_FULL_NAME}}` | Полное ФИО председателя |

Если переменная не найдена в данных — остаётся как есть, в лог пишется warning.

### SMTP

Заполните `.env`: `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASSWORD`, `FROM_EMAIL`, `FROM_NAME`.

Если SMTP не настроен, `send_emails.py` остановится с ошибкой:

```text
SMTP не настроен. Заполните .env: SMTP_HOST, FROM_EMAIL, SMTP_USER, SMTP_PASSWORD.
```

## 11. Отчёты

### Preflight validation

В начале каждого запуска бот проверяет все строки Excel до создания DOCX/PDF:

- обязательные поля, email, gender;
- корректное обращение `GREETING_NAME` минимум из 2 слов;
- отсутствие служебных слов в итоговом обращении.

Ошибочные строки пропускаются и попадают в status report / TXT ошибок.

### Excel status report

После каждого запуска: `status_YYYYMMDD_HHMMSS.xlsx` в соответствующей папке `reports/`.

Статусы различают этапы:

| Статус | Значение |
|---|---|
| `validation_failed` | Ошибка preflight, документ не создавался |
| `pdf_generated` | PDF создан, готов к рассылке |
| `pdf_failed` | DOCX создан, PDF не создан |

### Manifest успешных PDF

Реестр только успешно созданных PDF:

- Preview: `output/preview/reports/ген_N_manifest.xlsx`
- Final: `output/reports/ген_N_manifest.xlsx`

Колонки: `generation_id`, `mode`, `excel_row`, `bank_name`, `bank_legal_name`, `email_bank_name`, `chair_full_name`, `greeting_word`, `greeting_name_final`, `template_used`, `recipient_email`, `cc_email`, `bcc_email`, `out_number`, `letter_date`, `docx_path`, `pdf_path`, `email_status`, `created_at`.

Пути в manifest — относительные к корню проекта.

### TXT-файл ошибок

Создаётся только если есть ошибки:

- Preview: `output/preview/reports/ген_N_ошибки.txt`
- Final: `output/reports/ген_N_ошибки.txt`

`N` — номер запуска из `state/generation_state.json` → `last_generation_id`.

## 12. Логи

- Консоль — режим, номера, итоги.
- Файл `logs/app.log` — полный лог с ошибками.

## 13. Частые ошибки

| Проблема | Решение |
|---|---|
| Номер «улетает» при тестах | Используйте `python preview.py`, не `final_generate.py` |
| Номер зафиксировался без PDF | СВХИ фиксируется только после успешного PDF |
| Preview показывает старые файлы | Preview автоматически очищает `output/preview/` перед запуском |
| Одно слово после «Уважаемый» | Исправьте `greeting_name` или `chair_full_name` в Excel |
| Неправильный email | Проверьте `recipient_email` в Excel |
| Microsoft Word не установлен | Установите Word; без него PDF не создаётся |
| Остались `{{...}}` в Word | Проверьте плейсхолдеры в шаблоне |

## Быстрый старт

1. `pip install -r requirements.txt`
2. `copy .env.example .env`
3. Заполните `data/banks.xlsx` (или файл из `config.yaml`).
4. `python preview.py` → проверьте `output/preview/pdf/`.
5. `python final_generate.py` → финальная генерация с фиксацией номеров.
6. `python email_preview.py` → проверка рассылки.
7. `python send_emails.py` → реальная отправка.
