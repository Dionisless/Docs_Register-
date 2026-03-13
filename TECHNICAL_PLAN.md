# Технический план развития LanDocs Registrator

## Текущая архитектура (v1)

```
landocs_register.py  (~878 строк)
    Зависимости: pywin32, openpyxl, tkinter (stdlib)
    Сборка: PyInstaller --onefile --windowed

    Модули:
    - Клавиатура/буфер:  _send_tab, _send_ctrl_c, read_current_field, navigate_tabs
    - Парсинг LanDocs:   extract_landocs_data_in(), extract_landocs_data_out()
    - Excel запись:       write_to_excel_in(), write_to_excel_out()
    - ViewDir:            find_latest_in_viewdir()
    - Утилиты:            sanitize_for_filename, parse_date, fmt_date_*, calc_folder_num
    - UI:                 RegistrationApp(tk.Tk) с ttk.Notebook [Входящие | Исходящие]
```

---

## Новые фичи: обзор

| # | Фича | Новые зависимости | Приоритет |
|---|-------|-------------------|-----------|
| F1 | Импорт из журнала входящих (по номеру + году) | -- | Высокий |
| F2 | Процесс "Таблицы уставок": drag-and-drop, парсинг, запись, раскладка | python-docx | Высокий |
| F3 | Вывод изменений (синий цвет) из таблиц Word | python-docx | Средний |
| F4 | Обновление карт уставок в Visio + экспорт PDF | win32com.client (COM) | Средний |
| F5 | Выкладывание карт в ДЭБ (web-автоматизация) | selenium или pyautogui | Низкий |

### Новые зависимости для requirements.txt
```
python-docx>=1.1.0       # F2, F3: парсинг Word таблиц уставок
# win32com.client уже есть в pywin32  # F4: управление Visio через COM
# selenium — только для F5, добавляется позднее
```

---

## Новая архитектура UI

Переход от плоского Notebook к **двухуровневой иерархии вкладок**:

```
RegistrationApp (tk.Tk)
 +-- root_frame
      +-- top_notebook (ttk.Notebook)  ---- вкладки ПРОЦЕССОВ
      |   |
      |   +-- tab "Регистрация писем"
      |   |    +-- inner_notebook
      |   |         +-- "Входящие"   (текущий функционал + F1)
      |   |         +-- "Исходящие"  (текущий функционал)
      |   |
      |   +-- tab "Таблицы уставок"
      |        +-- inner_notebook  ---- шаги процесса
      |             +-- "Шаг 0: Входные данные"  (импорт из журнала + drag&drop таблиц)
      |             +-- "Шаг 1: Парсинг"         (центральная таблица с данными)
      |             +-- "Шаг 2: Запись в таблицы" (вставка письма о выдаче)
      |             +-- "Шаг 3: Реестры"          (запись в реестр / сводную)
      |             +-- "Шаг 4: Раскладка"        (перемещение файлов по папкам)
      |             +-- "Шаг 5: Изменения"        (вывод синих строк)  [F3]
      |             +-- "Шаг 6: Карты"            (обновление Visio)   [F4]
      |             +-- "Шаг 7: ДЭБ"              (веб-загрузка)       [F5]
      |
      +-- status_bar (Label, общий)
```

### Принцип: каждый шаг -- отдельный Frame, прозрачность данных

Каждый шаг показывает спарсенные данные в **редактируемых** полях. Переход к следующему шагу -- по кнопке. Это позволяет пользователю проверить и исправить данные перед действием.

---

## F1: Импорт из журнала входящих

### Задача
Во вкладке "Входящие" добавить возможность заполнить поля из уже существующей строки журнала Excel (по номеру вх + году).

### UI
В `_build_incoming_tab()` добавить секцию **перед** "Данные из LanDocs":

```
+-- LabelFrame "Импорт из журнала"
|   +-- [Entry "№ вх:" width=20]  [Entry "Год:" width=6, default=текущий]
|   +-- [Button "Найти в журнале"]
```

### Реализация

#### Новая функция: `lookup_incoming_journal(incoming_num: str, year: str) -> dict | None`

```python
def lookup_incoming_journal(incoming_num: str, year: str) -> dict | None:
    """
    Ищет строку в журнале входящих по столбцу B (№ вх) на листе,
    соответствующем году.

    Алгоритм:
    1. Открыть EXCEL_PATH_IN (read_only=True, data_only=True)
    2. Определить лист:
       - Если year совпадает с именем листа -> использовать его
       - Иначе -> последний лист
    3. Пройти столбец B снизу вверх
    4. Сравнить значение ячейки с incoming_num (case-insensitive, strip)
    5. Если найдено -> вернуть dict с полями:
       {
         'date':         ws.cell(row, 1).value,
         'incoming_num': ws.cell(row, 2).value,
         'letter_num':   ws.cell(row, 3).value,
         'subject':      ws.cell(row, 4).value,
         'author':       ws.cell(row, 5).value,
         'signed_by':    ws.cell(row, 6).value,
         'folder_num':   ws.cell(row, 7).value,
         'keywords':     ws.cell(row, 9).value,
         'related':      ws.cell(row, 10).value,
         'hyperlink':    ws.cell(row, 3).hyperlink.target if hyperlink else '',
       }
    6. Если не найдено -> return None
    """
```

#### Метод UI: `_on_import_journal(self)`

```python
def _on_import_journal(self):
    num = self.in_import_num_var.get().strip()
    year = self.in_import_year_var.get().strip()
    if not num:
        messagebox.showwarning("Внимание", "Введите № вх.", parent=self)
        return
    result = lookup_incoming_journal(num, year)
    if result is None:
        messagebox.showinfo("Не найдено",
            f"Запись «{num}» не найдена в журнале за {year} год.", parent=self)
        return
    # Заполняем in_data из найденной строки, конвертируя в формат LanDocs
    self.in_data = {
        'date':         str(result.get('date', '')),
        'incoming_num': str(result.get('incoming_num', '')),
        'letter_num':   str(result.get('letter_num', '')),
        'subject':      str(result.get('subject', '')),
        'signatory':    '',   # в журнале хранится signed_by целиком
        'correspondent': str(result.get('signed_by', '')),
        'related':      str(result.get('related', '')),
        'file_link':    str(result.get('hyperlink', '')),
    }
    self._apply_incoming_data()
    # Заполняем пользовательские поля из журнала
    self.in_author_var.set(str(result.get('author', '-')))
    self.in_keywords_var.set(str(result.get('keywords', '')))
    self._reparse_status.set(f"Импортировано: {num}")
```

### Связь с модулем таблиц уставок

Данные `self.in_data` после импорта из журнала используются на Шаге 0 таблиц уставок как источник "Письма о выставлении" (номер и дата). Для этого в `RegistrationApp` хранится ссылка `self.in_data`, которая доступна из любой вкладки.

---

## F2: Процесс "Таблицы уставок"

### Общая структура данных

```python
@dataclass
class TableEntry:
    """Одна таблица уставок в процессе обработки."""
    file_path: str          # полный путь к файлу Word
    form_type: str          # 'old' | 'new' (определяется автоматически)
    object_name: str        # Объект (ПС 220 кВ Зелёный угол)
    dispatch_name: str      # Диспетчерское наименование
    table_number: str       # Номер таблицы (только для old: "24-123")
    letter_num: str         # № письма о выставлении (из журнала)
    letter_date: str        # Дата письма (из журнала)
    registry_row: int       # Строка в реестре (0 = не найдена)
    archive_candidate: str  # Путь к таблице-кандидату на архивацию
    status: str             # 'pending' | 'parsed' | 'written' | 'registered' | 'moved'
```

Список всех таблиц хранится в `self.ustavki_entries: list[TableEntry]`.

### Центральная таблица (UI-виджет)

Используем `ttk.Treeview` со скроллбарами (вертикальный + горизонтальный):

```python
# Столбцы центральной таблицы
COLUMNS = [
    ('file',           'Файл',              200),
    ('form',           'Форма',              60),
    ('object',         'Объект',            180),
    ('dispatch',       'Дисп. наимен.',     180),
    ('table_num',      '№ таблицы',         100),
    ('letter_num',     '№ письма вых.',     140),
    ('letter_date',    'Дата выст.',        100),
    ('registry_row',   'Строка реестра',     100),
    ('archive',        'Архивная таблица',  200),
    ('status',         'Статус',            100),
]
```

Реализация:
```python
def _build_ustavki_table(self, parent):
    """Создает Treeview с горизонтальным и вертикальным скроллом."""
    container = ttk.Frame(parent)
    container.grid(row=0, column=0, sticky='nsew')
    parent.rowconfigure(0, weight=1)
    parent.columnconfigure(0, weight=1)

    self._tree = ttk.Treeview(container, columns=[c[0] for c in COLUMNS],
                               show='headings', height=15)
    for col_id, col_name, col_width in COLUMNS:
        self._tree.heading(col_id, text=col_name)
        self._tree.column(col_id, width=col_width, minwidth=60)

    vsb = ttk.Scrollbar(container, orient='vertical', command=self._tree.yview)
    hsb = ttk.Scrollbar(container, orient='horizontal', command=self._tree.xview)
    self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    self._tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')
    container.rowconfigure(0, weight=1)
    container.columnconfigure(0, weight=1)

    # Двойной клик по ячейке -> редактирование
    self._tree.bind('<Double-1>', self._on_tree_double_click)
```

Редактирование ячеек -- через `ttk.Entry`, который появляется поверх ячейки:
```python
def _on_tree_double_click(self, event):
    """Создает Entry поверх ячейки Treeview для редактирования."""
    region = self._tree.identify_region(event.x, event.y)
    if region != 'cell':
        return
    col = self._tree.identify_column(event.x)   # '#1', '#2', ...
    row = self._tree.identify_row(event.y)
    # Получить bbox, создать Entry, по Enter/FocusOut -- записать обратно
    ...
```

---

### Шаг 0: Входные данные (drag-and-drop + журнал)

#### UI

```
+-- LabelFrame "Письмо о выставлении"
|   +-- Label "№ вх:"    [in_data.incoming_num, readonly]
|   +-- Label "Дата:"    [in_data.date, readonly]
|   +-- Label "Тема:"    [in_data.subject, readonly]
|   +-- Подсказка: "Сначала импортируйте/запарсите письмо во вкладке Входящие"
|
+-- LabelFrame "Таблицы уставок"
|   +-- [Большая область для drag-and-drop / кнопка "Добавить файлы..."]
|   +-- Список добавленных файлов (Listbox)
|   +-- [Button "Удалить выбранные"]
|
+-- [Button "Далее: Парсинг >>"]
```

#### Drag-and-drop

Tkinter не поддерживает DnD из проводника нативно. Решения:

**Вариант A (рекомендуемый): tkinterdnd2**
```python
# Нужно добавить в зависимости: tkinterdnd2
# Или включить DLL в сборку PyInstaller
from tkinterdnd2 import DND_FILES, TkinterDnD

class RegistrationApp(TkinterDnD.Tk):  # вместо tk.Tk
    ...

# В _build_step0:
drop_area.drop_target_register(DND_FILES)
drop_area.dnd_bind('<<Drop>>', self._on_drop_files)

def _on_drop_files(self, event):
    # event.data содержит пути через пробел (или в фигурных скобках)
    paths = self._parse_dnd_paths(event.data)
    for p in paths:
        if p.lower().endswith(('.doc', '.docx')):
            self._add_table_file(p)
```

**Вариант B (без допзависимостей): кнопка "Добавить файлы"**
```python
def _add_files_dialog(self):
    files = filedialog.askopenfilenames(
        title="Выберите таблицы уставок",
        initialdir=USTAVKI_EXEC_FOLDER,  # [вставить путь]
        filetypes=[("Word файлы", "*.doc *.docx"), ("Все файлы", "*.*")],
    )
    for f in files:
        self._add_table_file(f)
```

**Рекомендация:** Реализовать оба варианта. Если tkinterdnd2 доступен -- DnD работает. Всегда есть кнопка "Добавить файлы" как fallback.

**Проверка наличия tkinterdnd2:**
```python
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
    _BASE_CLASS = TkinterDnD.Tk
except ImportError:
    HAS_DND = False
    _BASE_CLASS = tk.Tk

class RegistrationApp(_BASE_CLASS):
    ...
```

#### Добавление tkinterdnd2 в сборку

```
pip install tkinterdnd2
```

В .spec файле PyInstaller:
```python
# Добавить в datas:
import tkinterdnd2
datas=[(os.path.dirname(tkinterdnd2.__file__), 'tkinterdnd2')]
```

---

### Шаг 1: Парсинг данных из Word-таблиц

#### Определение формы (старая/новая)

```python
def detect_table_form(doc_path: str) -> str:
    """
    Определяет форму таблицы: 'old' или 'new'.

    Эвристика (уточнить на реальных файлах):
    - Новая форма: в документе есть ключевое слово/фраза, характерное для новой формы
      (например конкретный заголовок или структура таблицы)
    - Старая форма: в имени файла или в документе есть номер формата "YY-NNN"

    PLACEHOLDER: точные критерии определяются после анализа
    Приложений 1 и 2 (примеры таблиц).
    """
```

#### Парсинг данных из первой таблицы документа

```python
def parse_ustavki_table(doc_path: str) -> dict:
    """
    Парсит данные из таблицы на первой странице Word-документа.

    Использует python-docx для .docx и win32com для .doc.

    Возвращает:
    {
        'object':        str,  # "ПС 220 кВ Зелёный угол"
        'dispatch_name': str,  # Диспетчерское наименование
        'table_number':  str,  # "24-123" (для старых), "" для новых
        'form_type':     str,  # 'old' | 'new'
    }

    Алгоритм:
    1. Открыть документ
    2. Взять первую таблицу (doc.tables[0])
    3. Пройти по строкам/ячейкам и извлечь:
       - Строку с "Объект" / "Подстанция" -> object
       - Строку с "Диспетчерское наименование" -> dispatch_name
    4. Для определения номера таблицы (old form):
       - regex по имени файла: r'(\d{2})-(\d+)'  (YY-NNN)
       - или из содержимого документа
    """
```

**Работа с .doc (старый формат):**
```python
def _open_doc_as_docx(doc_path: str) -> str:
    """
    Конвертирует .doc в .docx через COM Word.
    Возвращает путь к временному .docx файлу.
    """
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(doc_path))
    tmp_path = doc_path + 'x'  # .doc -> .docx
    doc.SaveAs2(tmp_path, FileFormat=16)  # wdFormatXMLDocument
    doc.Close()
    word.Quit()
    return tmp_path
```

**PLACEHOLDER:** Точная структура парсинга определяется после анализа Приложений 1 и 2. При реализации нужно передать примеры файлов.

#### Метод UI: кнопка "Парсить все"

```python
def _on_parse_all(self):
    """Парсит все добавленные таблицы и заполняет Treeview."""
    letter_num  = self.in_data.get('incoming_num', '')
    letter_date = self.in_data.get('date', '')

    for entry in self.ustavki_entries:
        try:
            parsed = parse_ustavki_table(entry.file_path)
            entry.object_name   = parsed['object']
            entry.dispatch_name = parsed['dispatch_name']
            entry.table_number  = parsed['table_number']
            entry.form_type     = parsed['form_type']
            entry.letter_num    = letter_num
            entry.letter_date   = letter_date
            entry.status        = 'parsed'
        except Exception as exc:
            entry.status = f'ошибка: {exc}'

    self._refresh_tree()
```

---

### Шаг 2: Запись "Таблицы выданы" в каждый Word-документ

#### Поиск надписи "таблицы выданы:" с допуском до 2 ошибок

```python
def _fuzzy_find_tables_issued(text: str) -> int | None:
    """
    Ищет позицию фразы "таблицы выданы:" в тексте
    с допуском до 2 ошибочных символов (расстояние Левенштейна <= 2).
    Регистр не важен.

    Возвращает индекс начала найденной подстроки или None.

    Реализация через скользящее окно:
    """
    target = "таблицы выданы:"
    text_lower = text.lower()
    target_len = len(target)

    for i in range(len(text_lower) - target_len + 1):
        window = text_lower[i:i + target_len]
        dist = _levenshtein(window, target)
        if dist <= 2:
            return i
    return None


def _levenshtein(s1: str, s2: str) -> int:
    """Расстояние Левенштейна -- стандартный алгоритм, O(n*m)."""
    if len(s1) < len(s2):
        return _levenshtein(s2, s1)
    if len(s2) == 0:
        return len(s1)
    prev_row = range(len(s2) + 1)
    for i, c1 in enumerate(s1):
        curr_row = [i + 1]
        for j, c2 in enumerate(s2):
            insertions = prev_row[j + 1] + 1
            deletions = curr_row[j] + 1
            substitutions = prev_row[j] + (c1 != c2)
            curr_row.append(min(insertions, deletions, substitutions))
        prev_row = curr_row
    return prev_row[-1]
```

#### Запись в Word-документ

```python
def write_issued_to_doc(doc_path: str, letter_num: str, letter_date: str,
                        hyperlink_path: str):
    """
    Находит 'таблицы выданы:' в документе и вставляет после неё:
    ' <letter_num> от <дата в формате гггг_мм_дд>'
    letter_num -- гиперссылка на файл письма.

    Алгоритм (python-docx для .docx):
    1. Пройти все параграфы doc.paragraphs
    2. Для каждого параграфа: _fuzzy_find_tables_issued(paragraph.text)
    3. Если нашли:
       a. Найти run, содержащий конец фразы "выданы:"
       b. После этого run добавить новый run с текстом
       c. Добавить гиперссылку (через oxml manipulation)
    4. Сохранить документ

    Для .doc: конвертировать через COM -> обработать .docx -> конвертировать обратно
    """
```

**Добавление гиперссылки в docx (через oxml):**
```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def _add_hyperlink(paragraph, url, text):
    """Добавляет гиперссылку в параграф Word."""
    part = paragraph.part
    r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
                          is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    # Стиль гиперссылки (синий, подчеркнутый)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0563C1')
    rPr.append(u)
    rPr.append(color)
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    paragraph._element.append(hyperlink)
```

---

### Шаг 3: Запись в реестры

#### 3A: Новая форма -> Реестр уставок

```
Путь: [вставить ссылку на файл реестра уставок]
Формат: Excel (.xlsx)
Столбцы для записи: "№вх", "Письмо о выставлении", "Дата выставления"
                     [вставить номера столбцов]
```

```python
def write_to_registry_new(entry: TableEntry, registry_path: str):
    """
    Записывает данные о письме выставления в реестр уставок (новая форма).

    Поиск строки:
    1. Открыть registry_path
    2. Взять последний лист
    3. Идти снизу вверх
    4. Для каждой строки сравнить:
       - Столбец X (Объект) == entry.object_name
       - Столбец Y (Дисп. наименование) == entry.dispatch_name
       - Столбец Z (№ письма) == entry.letter_num_original  (номер нашего исх. письма)
    5. Если все 3 совпали -> записать в ту строку:
       - Столбец A: entry.letter_num  (№ вх)
       - Столбец B: entry.letter_num  (Письмо о выставлении)
       - Столбец C: entry.letter_date (Дата выставления)
    6. Если не найдено -> вернуть None (обработка в UI)
    """
```

**Обработка "не найдено" в UI:**
```python
def _on_registry_not_found(self, entry: TableEntry, candidates: list):
    """
    Показывает диалог:
    - Список кандидатов (совпадение 2 из 3 полей) с радиокнопками
    - Поле для ручного ввода номера строки
    - Кнопка "Пропустить"
    """
    dialog = tk.Toplevel(self)
    dialog.title(f"Строка не найдена: {entry.dispatch_name}")
    dialog.transient(self)
    dialog.grab_set()
    ...
```

**Поиск кандидатов (2 из 3):**
```python
def find_registry_candidates(ws, entry, obj_col, disp_col, letter_col) -> list:
    """Ищет строки с совпадением минимум 2 из 3 полей."""
    candidates = []
    for row in range(ws.max_row, 1, -1):
        obj_val  = str(ws.cell(row, obj_col).value or '')
        disp_val = str(ws.cell(row, disp_col).value or '')
        let_val  = str(ws.cell(row, letter_col).value or '')

        matches = sum([
            obj_val.strip().lower() == entry.object_name.strip().lower(),
            disp_val.strip().lower() == entry.dispatch_name.strip().lower(),
            let_val.strip().lower() == entry.letter_num.strip().lower(),
        ])
        if matches >= 2:
            candidates.append({
                'row': row,
                'object': obj_val,
                'dispatch': disp_val,
                'letter': let_val,
                'matches': matches,
            })
    return candidates
```

#### 3B: Старая форма -> Сводная по таблицам

```
Путь: [вставить ссылку на сводную]
Формат: Excel (.xlsx)
Листы: годы ("2024", "2025", ...)
Столбец номеров таблиц: [вставить номер]
Столбец письма об исполнении: [вставить номер]
```

```python
def write_to_summary_old(entry: TableEntry, summary_path: str):
    """
    Записывает данные в 'Сводную по таблицам' (старая форма).

    1. Открыть summary_path
    2. Определить лист по году из entry.table_number: "24-123" -> лист "2024"
       (или последний лист)
    3. Найти строку: столбец [N] содержит entry.table_number
    4. Записать в столбец [M]: номер письма об исполнении
    """
```

---

### Шаг 4: Раскладывание таблиц по папкам

#### Конфигурация путей

```python
# PLACEHOLDER: заполнить конкретные пути
USTAVKI_EXEC_FOLDER = r"[вставить путь к папке для исполнения]"

# Структура папок объектов (пример):
# \\server\path\Зеленый угол\
#     Текущие\
#         файл_таблицы.docx
#     Архив\
#         старый_файл_таблицы.docx
```

#### Маппинг "официальное название -> жаргонное/папка"

Автоматическое определение папки объекта через нечёткий поиск:

```python
def find_object_folder(object_name: str, base_folder: str) -> str | None:
    """
    Находит папку объекта по его официальному названию.

    Алгоритм:
    1. Получить список подпапок в base_folder
    2. Из object_name извлечь ключевые слова:
       - Убрать стандартные префиксы: "ПС", "220 кВ", "110 кВ", "станция" и тд
       - Оставить основное название: "Зелёный угол" -> "зеленый угол"
    3. Для каждой подпапки:
       - Нормализовать: нижний регистр, убрать ё->е, пунктуацию
       - Проверить вхождение ключевых слов
    4. Вернуть лучшее совпадение или None
    """
    # Нормализация
    def normalize(s):
        s = s.lower().replace('ё', 'е')
        s = re.sub(r'[^\w\s]', '', s)
        return s

    # Извлечь имя объекта без типовых префиксов
    name = re.sub(r'^(ПС|АТЭЦ|ТЭЦ|ГЭС|СЭС)\s+\d+\s*кВ\s+', '', object_name, flags=re.IGNORECASE)
    name_norm = normalize(name)

    best_match, best_score = None, 0
    for folder_name in os.listdir(base_folder):
        folder_path = os.path.join(base_folder, folder_name)
        if not os.path.isdir(folder_path):
            continue
        folder_norm = normalize(folder_name)
        # Считаем совпадающие слова
        name_words = set(name_norm.split())
        folder_words = set(folder_norm.split())
        common = name_words & folder_words
        score = len(common) / max(len(name_words), 1)
        if score > best_score:
            best_score = score
            best_match = folder_path

    return best_match if best_score > 0.3 else None
```

#### Поиск архивных кандидатов (близость имён файлов)

```python
def find_archive_candidates(new_file: str, current_folder: str, top_n: int = 5) -> list:
    """
    Ищет в current_folder файлы, наиболее похожие на new_file.
    Последние 5 символов каждого имени (до расширения) отрезаются как шум.

    Возвращает список (path, score) отсортированный по убыванию score.

    Алгоритм близости -- SequenceMatcher из difflib (stdlib, без зависимостей):
    """
    from difflib import SequenceMatcher

    new_name = os.path.splitext(os.path.basename(new_file))[0]
    new_name_trimmed = new_name[:-5] if len(new_name) > 5 else new_name
    new_name_trimmed = new_name_trimmed.lower()

    candidates = []
    for fname in os.listdir(current_folder):
        fpath = os.path.join(current_folder, fname)
        if not os.path.isfile(fpath):
            continue
        if fpath == new_file:
            continue
        name = os.path.splitext(fname)[0]
        name_trimmed = name[:-5] if len(name) > 5 else name
        name_trimmed = name_trimmed.lower()

        score = SequenceMatcher(None, new_name_trimmed, name_trimmed).ratio()
        candidates.append((fpath, fname, score))

    candidates.sort(key=lambda x: x[2], reverse=True)
    return candidates[:top_n]
```

#### UI: выбор архивных кандидатов

В Treeview каждой таблицы показываются кандидаты, пользователь выбирает из списка:

```python
def _show_archive_selector(self, entry: TableEntry, candidates: list):
    """
    Диалог со списком кандидатов для архивации.
    Каждый кандидат -- строка с чекбоксом и показателем близости.
    Один выбранный по умолчанию (наибольшая близость).
    """
```

#### Логика перемещения

```python
def move_tables(entry: TableEntry, object_folder: str):
    """
    1. Найти подпапку "Текущие" (или "Текущая", case-insensitive)
    2. Найти подпапку "Архив"
    3. Если entry.archive_candidate:
       shutil.move(entry.archive_candidate, архив_папка)
    4. shutil.copy2(entry.file_path, текущие_папка)
    """
    current_dir = _find_subfolder(object_folder, ['текущие', 'текущая'])
    archive_dir = _find_subfolder(object_folder, ['архив'])

    if not current_dir or not archive_dir:
        raise FileNotFoundError(
            f"Не найдены папки Текущие/Архив в {object_folder}")

    if entry.archive_candidate and os.path.exists(entry.archive_candidate):
        dest = os.path.join(archive_dir, os.path.basename(entry.archive_candidate))
        shutil.move(entry.archive_candidate, dest)

    dest_new = os.path.join(current_dir, os.path.basename(entry.file_path))
    shutil.copy2(entry.file_path, dest_new)
```

---

## F3: Вывод изменений (синий цвет) в таблицах

### Шаг 5 в UI

#### Обнаружение синего текста в Word

```python
from docx.shared import RGBColor

def extract_blue_rows(doc_path: str) -> list[list[str]]:
    """
    Находит строки таблиц Word, содержащие текст синего цвета.

    Определение "синего":
    - Точное RGB: (0, 0, 255) или близкие оттенки
    - Допуск: R < 80, G < 80, B > 150
    - Или theme color = accent1 (стандартный синий Word)

    Алгоритм:
    1. Для каждой таблицы в документе:
    2.   Для каждой строки таблицы:
    3.     Для каждой ячейки:
    4.       Для каждого run в параграфах ячейки:
    5.         Проверить run.font.color.rgb
    6.         Если синий -> помечаем строку
    7.   Собираем помеченные строки как list[str] (текст ячеек)
    """
    from docx import Document
    doc = Document(doc_path)
    blue_rows = []

    for table in doc.tables:
        for row in table.rows:
            has_blue = False
            row_data = []
            for cell in row.cells:
                cell_text = cell.text
                row_data.append(cell_text)
                for para in cell.paragraphs:
                    for run in para.runs:
                        color = run.font.color
                        if color and color.rgb:
                            r, g, b = color.rgb[0], color.rgb[1], color.rgb[2]
                            if r < 80 and g < 80 and b > 150:
                                has_blue = True
                        elif color and color.theme_color:
                            # theme_color может быть ACCENT_1 и т.д.
                            if 'ACCENT' in str(color.theme_color):
                                has_blue = True
            if has_blue:
                blue_rows.append(row_data)

    return blue_rows
```

#### Генерация Word-документа со сводкой изменений

```python
def generate_changes_report(entries: list[TableEntry], output_path: str):
    """
    Создаёт Word-документ со сводкой изменений.

    Структура:
    ---
    Таблица 1
    Объект: ПС 220 кВ Зелёный угол    Дисп. наим.: ...

    | Столбец 1 | Столбец 2 | ... |
    | данные     | данные    | ... |

    Таблица 2
    ...
    ---
    """
    from docx import Document
    doc = Document()

    for i, entry in enumerate(entries, 1):
        doc.add_heading(f'Таблица {i}', level=2)
        doc.add_paragraph(
            f'Объект: {entry.object_name}    '
            f'Дисп. наименование: {entry.dispatch_name}')

        blue_rows = extract_blue_rows(entry.file_path)
        if not blue_rows:
            doc.add_paragraph('Изменений (синий цвет) не обнаружено.')
            continue

        # Создать таблицу Word
        num_cols = max(len(row) for row in blue_rows)
        table = doc.add_table(rows=len(blue_rows), cols=num_cols)
        table.style = 'Table Grid'
        for r, row_data in enumerate(blue_rows):
            for c, cell_text in enumerate(row_data):
                table.cell(r, c).text = cell_text

    doc.save(output_path)
    os.startfile(output_path)  # Открыть в Word
```

---

## F4: Обновление карт уставок (Visio)

### Шаг 6 в UI

```
Путь к картам: [вставить адрес]
```

#### Замена гиперссылок в Visio через COM

```python
def update_visio_map(visio_path: str, old_link: str, new_link: str,
                     new_table_number: str):
    """
    Открывает Visio файл, находит гиперссылку на old_link,
    заменяет её на new_link и обновляет текст гиперссылки.

    Требует: Microsoft Visio установлен на машине.
    """
    import win32com.client
    visio = win32com.client.Dispatch('Visio.Application')
    visio.Visible = False

    doc = visio.Documents.Open(os.path.abspath(visio_path))

    for page in doc.Pages:
        for shape in page.Shapes:
            for i in range(shape.Hyperlinks.Count):
                hl = shape.Hyperlinks.Item(i + 1)  # 1-based
                if old_link.lower() in hl.Address.lower():
                    hl.Address = new_link
                    # Обновить текст (номер таблицы)
                    if new_table_number:
                        hl.Description = new_table_number
                    break

    doc.Save()

    # Сохранить как PDF
    pdf_folder = r"[вставить путь к папке PDF]"
    pdf_name = os.path.splitext(os.path.basename(visio_path))[0] + '.pdf'
    pdf_path = os.path.join(pdf_folder, pdf_name)
    doc.ExportAsFixedFormat(
        1,  # visFixedFormatPDF
        pdf_path,
        0,  # visDocExIntentPrint
        0,  # visPrintAll
    )

    doc.Close()
    visio.Quit()
```

**Определение old_link (архивная таблица):**
```python
# old_link = путь к таблице, которую мы переместили в Архив (entry.archive_candidate)
# new_link = новый путь таблицы в папке Текущие
```

**Парсинг номера таблицы из имени файла (для новых таблиц):**
```python
def extract_table_number_from_filename(filename: str) -> str:
    """Извлекает номер таблицы формата YY-NNN из имени файла."""
    m = re.search(r'(\d{2})-(\d+)', filename)
    return m.group(0) if m else ''
```

---

## F5: Выкладывание в ДЭБ (веб-автоматизация)

### Шаг 7 в UI

> Реализация зависит от конкретного HTML интерфейса ДЭБ. Ниже -- общий каркас.

#### Вариант A: Selenium (если есть браузер + chromedriver)

```python
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def upload_to_deb(object_name: str, visio_path: str, pdf_path: str,
                  deb_url: str):
    """
    Автоматически загружает карту уставок в ДЭБ.

    1. Открыть deb_url (раздел с картами уставок)
    2. Найти объект по тексту object_name -> кликнуть
    3. Дождаться загрузки страницы объекта
    4. Нажать кнопку редактирования [ссылка на кусок html]
    5. Загрузить visio файл в поле [html]
    6. Загрузить pdf файл в поле [html]
    7. Нажать "Сохранить" [html]
    """
```

#### Вариант B: pyautogui (если браузер открыт вручную)

Более простой, но менее надёжный подход. Пользователь открывает страницу ДЭБ, программа кликает по координатам / ищет картинки.

**Рекомендация:** Selenium предпочтительнее, но требует chromedriver. Если на рабочем месте нет доступа к интернету -- chromedriver нужно положить рядом с .exe. Альтернатива -- использовать IE/Edge через COM.

**PLACEHOLDER:** Реализация уточняется после получения HTML-страниц ДЭБ.

---

## Конфигурация: вынос путей

Все пути вынести в секцию конфигурации вверху файла (или в отдельный config.py / .ini):

```python
# ── Конфигурация ──────────────────────────────────────────────────────────────

# Журналы регистрации
EXCEL_PATH_IN  = r"\\Prim-fs-serv\primrdu\СРЗА\Журнал регистрации корреспонденции\Журнал регистрации входящей документации.xlsx"
EXCEL_PATH_OUT = r"\\Prim-fs-serv\primrdu\СРЗА\Журнал регистрации корреспонденции\Журнал регистрации исходящей документации.xlsx"

# Базовая папка для сохранения писем
DEFAULT_SAVE_FOLDER = r"\\Prim-fs-serv\primrdu\СРЗА\Дела СРЗА\19 Переписка"

# Реестр уставок (новая форма)
REGISTRY_PATH = r"[вставить ссылку на файл реестра уставок]"
REGISTRY_OBJ_COL    = None  # [вставить номер столбца "Объект"]
REGISTRY_DISP_COL   = None  # [вставить номер столбца "Дисп. наименование"]
REGISTRY_LETTER_COL = None  # [вставить номер столбца "№ письма"]
REGISTRY_VX_COL     = None  # [вставить номер столбца "№вх"]
REGISTRY_ISSUE_COL  = None  # [вставить номер столбца "Письмо о выставлении"]
REGISTRY_DATE_COL   = None  # [вставить номер столбца "Дата выставления"]

# Сводная по таблицам (старая форма)
SUMMARY_PATH = r"[вставить ссылку на сводную по таблицам]"
SUMMARY_NUM_COL   = None  # [вставить номер столбца "Номер таблицы"]
SUMMARY_ISSUE_COL = None  # [вставить номер столбца "Письмо об исполнении"]

# Папка с таблицами для исполнения
USTAVKI_EXEC_FOLDER = r"[вставить путь]"

# Папка с картами уставок (Visio)
MAPS_FOLDER = r"[вставить адрес]"

# Папка для PDF-экспорта карт
MAPS_PDF_FOLDER = r"[вставить название папки]"

# ДЭБ
DEB_MAPS_URL = r"[вставить ссылку на раздел с картами уставок]"
```

---

## Обновление зависимостей и сборки

### requirements.txt
```
pywin32>=306
openpyxl>=3.1.0
python-docx>=1.1.0
tkinterdnd2>=0.3.0
pyinstaller>=6.0.0
```

### build.bat -- добавить python-docx и tkinterdnd2
```bat
pip install pywin32 openpyxl python-docx tkinterdnd2 pyinstaller
```

### PyInstaller .spec -- добавить данные tkinterdnd2
```python
import tkinterdnd2
datas = [(os.path.dirname(tkinterdnd2.__file__), 'tkinterdnd2')]
```

---

## PLACEHOLDER-ы: что нужно заполнить перед реализацией

Отмечены как `[вставить ...]` в тексте. Полный список:

| # | Что нужно | Где используется |
|---|-----------|-----------------|
| 1 | Ссылка на реестр уставок (.xlsx) | `REGISTRY_PATH` |
| 2 | Номера столбцов реестра: Объект, Дисп.наим., №письма, №вх, Письмо о выст., Дата | `REGISTRY_*_COL` |
| 3 | Ссылка на сводную по таблицам (.xlsx) | `SUMMARY_PATH` |
| 4 | Номер столбца "Номер таблицы" в сводной | `SUMMARY_NUM_COL` |
| 5 | Номер столбца "Письмо об исполнении" в сводной | `SUMMARY_ISSUE_COL` |
| 6 | Путь к папке для исполнения (таблицы) | `USTAVKI_EXEC_FOLDER` |
| 7 | Путь к картам уставок (Visio) | `MAPS_FOLDER` |
| 8 | Путь к папке PDF-экспорта карт | `MAPS_PDF_FOLDER` |
| 9 | URL раздела ДЭБ с картами | `DEB_MAPS_URL` |
| 10 | HTML-страницы ДЭБ для контекста | Для F5 (Selenium-скрипт) |
| 11 | Приложение 1 (пример старой таблицы Word) | Для точного парсинга в F2 |
| 12 | Приложение 2 (пример новой таблицы Word) | Для точного парсинга в F2 |

---

## Порядок реализации (рекомендуемый)

```
Этап 1 — Базовая инфраструктура (не требует PLACEHOLDER-ов)
    [x] F1: Импорт из журнала входящих
    [ ] Рефакторинг UI: двухуровневый Notebook
    [ ] Центральная таблица (Treeview) для таблиц уставок
    [ ] DnD / добавление файлов

Этап 2 — Парсинг и запись (нужны Приложения 1, 2)
    [ ] Шаг 1: Парсинг Word-таблиц (python-docx)
    [ ] Шаг 2: Запись "таблицы выданы" в документы

Этап 3 — Реестры (нужны ссылки и номера столбцов)
    [ ] Шаг 3A: Запись в реестр (новая форма)
    [ ] Шаг 3B: Запись в сводную (старая форма)

Этап 4 — Файловые операции (нужны пути)
    [ ] Шаг 4: Раскладывание таблиц по папкам

Этап 5 — Анализ изменений (нужны приложения для тестов)
    [ ] Шаг 5: Извлечение синих строк, генерация отчёта Word

Этап 6 — Visio (нужен путь к картам)
    [ ] Шаг 6: Обновление гиперссылок в Visio + PDF экспорт

Этап 7 — ДЭБ (нужны HTML-страницы и URL)
    [ ] Шаг 7: Автоматизация веб-загрузки
```
