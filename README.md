# Elexam
![Python version](https://img.shields.io/badge/python-3.12|3.13-blue)  
<img src="https://github.com/Weendx/elexam/blob/main/attachments/preview.jpg?raw=true" height="200" alt="elexam screenshot">

Программа для обработки таблицы с желающими сдать экзамен. Кроме этого, ещё имеет несколько дополнительных функций. Разработана на python 3.13, поэтому кроссплатформенность тоже имеется.

## Скачать

- [Windows](https://github.com/Weendx/elexam/releases/latest/download/elexam.exe)
- [Linux](https://github.com/Weendx/elexam/releases/latest/download/elexam)

## Установка

Установка не требуется, достаточно просто запустить исполняемый файл

## Параметры запуска

- `--tui` - Запуск TUI интерфейса
- `--settings` - Посмотреть расположение файла с настройками
- `--version` - Вывести текущую версию программы

## Сборка

```
git clone https://notabug.org/Weendx/elexam.git
cd elexam
python -m venv venv

# Для linux:
source venv/bin/activate
# Для windows:
# venv\Scripts\activate

pip install -r requirements.txt
```

Теперь можно запустить программу, используя python:
```
python3 src/tui.py
# или python3 src/app.py --tui
```

Если нужно собрать программу в исполняемый файл, то нужно к тому, что выше, добавить:
```
pip install pyinstaller
pyinstaller --onefile -n elexam src/tui.py
```
После окончания сборки исполяемый файл появится в папке `dist`.

Чтобы была возможность запуска программы из любой точки системы, добавьте директорию с ней в переменную окружения `PATH`.
