# Помощник Опера — MVP v0.1

Первый MVP: загрузка исходного Excel → выбор документа «Проработка» → генерация заполненного Excel по шаблону → скачивание.

## Запуск

```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
uvicorn backend.app.main:app --reload
```

Откройте `http://127.0.0.1:8000`.

## Сопоставление полей «Проработки»

- `номер ктк` → `номер ктк`
- `REF` → `Букинг`
- `статус` → `Клиент`
- `фактический объем, м3` → `объем, м3`
- `вес, кг` → `вес, кг`
- `кол-во мест` → `кол-во мест`
